#!/usr/bin/env python3
"""
ospf_from_xlsx.py

Read the "protocol_ospf" worksheet of router_interfaces.xlsx, match the
routers against the Ansible inventory, and generate the Ansible playbook
push_ospf.yaml.

    python3 ospf_from_xlsx.py -x router_interfaces_1.xlsx -i inventory.ini
    ansible-playbook -i inventory.ini push_ospf.yaml --check --diff
    ansible-playbook -i inventory.ini push_ospf.yaml

Worksheet columns (header row required):
    hostname, interface, protocol, process_id, router_id, area, cost,
    network_type, passive, subnet

Rules
  * Based on the inventory: only hosts present in the inventory are written to
    the playbook. A worksheet host that is NOT in the inventory is an error
    (nothing is generated) unless --ignore-missing-hosts is given, so no
    listed interface is ever silently dropped.
  * Every OSPF interface row is configured with "ip ospf <pid> area <area>"
    (plus "ip ospf cost" when a cost is given). A coverage check at the end
    proves rows read == rows configured.
  * network_type blank / null  ->  NO "ip ospf network" line is generated.
  * passive = yes  ->  "passive-interface <ifname>" under the OSPF process.
  * The "subnet" column is not used for configuration; it is only cross-checked
    against the Interfaces sheet and reported as a warning if inconsistent.

Requires: openpyxl only. PyYAML is optional (needed only to read a YAML-format
inventory without ansible-inventory, and to self-check the generated file).
"""

import argparse
import ipaddress
import json
import re
import subprocess
import sys
from collections import OrderedDict
from pathlib import Path
from textwrap import indent

import openpyxl

try:  # optional: only used to read YAML inventories and to validate the output
    import yaml
except ImportError:  # the script works without it for INI inventories
    yaml = None

BLANK_TOKENS = {"", "none", "null", "nan", "n/a", "na", "-"}
TRUE_TOKENS = {"yes", "y", "true", "t", "1", "passive"}
FALSE_TOKENS = {"no", "n", "false", "f", "0"}
VALID_NETWORK_TYPES = {
    "broadcast",
    "non-broadcast",
    "point-to-point",
    "point-to-multipoint",
    "point-to-multipoint non-broadcast",
}
REQUIRED_COLUMNS = ["hostname", "interface", "process_id", "area"]
SHEET_CANDIDATES = ["protocol_ospf", "protocol_opsf", "ospf", "protocol"]

PLAYBOOK_TASKS = """\
  tasks:
    - name: Configure OSPF process (router-id, passive interfaces)
      cisco.ios.ios_config:
        parents: "router ospf {{ item.process_id }}"
        lines: "{{ item.router_lines }}"
      loop: "{{ ospf_config[inventory_hostname].processes }}"
      loop_control:
        label: "{{ inventory_hostname }} router ospf {{ item.process_id }}"
      when: item.router_lines | length > 0

    - name: Enable OSPF on every listed interface
      cisco.ios.ios_config:
        parents: "interface {{ item.1.name }}"
        lines: "{{ item.1.lines }}"
      loop: "{{ ospf_config[inventory_hostname].processes | subelements('interfaces') }}"
      loop_control:
        label: "{{ inventory_hostname }} {{ item.1.name }}"

    - name: Save configuration when anything changed
      cisco.ios.ios_config:
        save_when: modified

    - name: Show OSPF neighbours
      cisco.ios.ios_command:
        commands: show ip ospf neighbor
      register: ospf_neighbors
      changed_when: false

    - name: Print OSPF neighbours
      ansible.builtin.debug:
        var: ospf_neighbors.stdout_lines
"""


# --------------------------------------------------------------------------
# helpers
# --------------------------------------------------------------------------
def clean(value):
    """Stripped string, or None for blank / null-like cells."""
    if value is None:
        return None
    text = str(value).strip()
    return None if text.lower() in BLANK_TOKENS else text


def to_int_or_str(value):
    """1 / 1.0 / '1' -> 1 ; dotted-quad areas like '0.0.0.0' stay strings."""
    text = clean(value)
    if text is None:
        return None
    try:
        number = float(text)
        if number.is_integer():
            return int(number)
    except ValueError:
        pass
    return text


def to_bool(value, where):
    text = clean(value)
    if text is None:
        return False
    low = text.lower()
    if low in TRUE_TOKENS:
        return True
    if low in FALSE_TOKENS:
        return False
    raise ValueError(f"{where}: cannot interpret passive value {value!r} (use yes/no)")


def _scalar(value):
    """JSON scalars are valid YAML: quoted strings, bare ints."""
    return json.dumps(value) if isinstance(value, str) else str(value)


def to_yaml(obj, level=0):
    """Minimal block-style YAML writer for nested dict / list / str / int."""
    pad = "  " * level
    out = []
    if isinstance(obj, dict):
        for key, val in obj.items():
            if isinstance(val, (dict, list)) and val:
                out.append(f"{pad}{json.dumps(str(key))}:")
                out.append(to_yaml(val, level + 1))
            elif isinstance(val, list):
                out.append(f"{pad}{json.dumps(str(key))}: []")
            elif isinstance(val, dict):
                out.append(f"{pad}{json.dumps(str(key))}: {{}}")
            else:
                out.append(f"{pad}{json.dumps(str(key))}: {_scalar(val)}")
    else:  # list
        for item in obj:
            if isinstance(item, dict) and item:
                lines = to_yaml(item, level + 1).splitlines()
                lines[0] = pad + "- " + lines[0][len(pad) + 2:]
                out.extend(lines)
            elif isinstance(item, list):
                raise ValueError("nested lists are not supported")
            else:
                out.append(f"{pad}- {_scalar(item)}")
    return "\n".join(out)


def find_sheet(wb, requested):
    names = {ws.title.lower(): ws.title for ws in wb.worksheets}
    if requested:
        if requested.lower() in names:
            return wb[names[requested.lower()]]
        sys.exit(f"ERROR: sheet {requested!r} not found. Sheets: {wb.sheetnames}")
    for candidate in SHEET_CANDIDATES:
        if candidate in names:
            return wb[names[candidate]]
    sys.exit(f"ERROR: no OSPF sheet found (looked for {SHEET_CANDIDATES}). "
             f"Sheets: {wb.sheetnames}. Use --sheet NAME.")


def read_table(ws):
    """Yield (excel_row_number, {header: value}) for every non-empty data row."""
    rows = ws.iter_rows(values_only=True)
    header = next(rows, None)
    if header is None:
        sys.exit(f"ERROR: sheet {ws.title!r} is empty")
    header = [str(h).strip().lower() if h is not None else "" for h in header]
    missing = [c for c in REQUIRED_COLUMNS if c not in header]
    if missing:
        sys.exit(f"ERROR: sheet {ws.title!r} is missing column(s): {missing}")
    for number, row in enumerate(rows, start=2):
        if all(cell is None or str(cell).strip() == "" for cell in row):
            continue
        yield number, dict(zip(header, row))


# --------------------------------------------------------------------------
# inventory
# --------------------------------------------------------------------------
def inventory_hosts(path):
    """Host names in an Ansible inventory.

    Uses `ansible-inventory --list` if available (handles INI, YAML, groups,
    dynamic inventories); otherwise falls back to a simple INI/YAML parser.
    """
    path = Path(path)
    if not path.exists():
        sys.exit(f"ERROR: inventory {path} not found")

    try:
        out = subprocess.run(["ansible-inventory", "-i", str(path), "--list"],
                             capture_output=True, text=True, check=True).stdout
        hostvars = json.loads(out).get("_meta", {}).get("hostvars", {})
        if hostvars:
            return list(hostvars)
    except (FileNotFoundError, subprocess.CalledProcessError, ValueError):
        pass

    text = path.read_text()
    if path.suffix in {".yml", ".yaml"}:
        if yaml is None:
            sys.exit("ERROR: reading a YAML inventory needs PyYAML (pip install pyyaml) "
                     "or ansible-inventory on the PATH")
        data = yaml.safe_load(text) or {}
        hosts = []

        def walk(node):
            if isinstance(node, dict):
                for key, value in node.items():
                    if key == "hosts" and isinstance(value, dict):
                        hosts.extend(value)
                    else:
                        walk(value)

        walk(data)
        return list(OrderedDict.fromkeys(hosts))

    hosts, in_hosts = [], True
    for line in text.splitlines():
        line = line.split("#", 1)[0].split(";", 1)[0].strip()
        if not line:
            continue
        if line.startswith("["):
            in_hosts = not re.search(r":(vars|children)\]", line)
            continue
        if in_hosts:
            hosts.append(line.split()[0])
    return list(OrderedDict.fromkeys(hosts))


def match_host(name, inventory):
    """Exact match first, then case-insensitive. Returns the inventory name or None."""
    if name in inventory:
        return name
    return {h.lower(): h for h in inventory}.get(name.lower())


# --------------------------------------------------------------------------
# Interfaces sheet (sanity check only)
# --------------------------------------------------------------------------
def load_interface_ips(wb):
    """{(host_lower, interface_lower): 'ip/mask' or None} from the wide Interfaces sheet."""
    ips = {}
    sheet = next((ws for ws in wb.worksheets if ws.title.lower() == "interfaces"), None)
    if sheet is None:
        return ips
    rows = list(sheet.iter_rows(values_only=True))
    if not rows:
        return ips
    header = [str(h).strip().lower() if h else "" for h in rows[0]]
    for row in rows[1:]:
        record = dict(zip(header, row))
        host = clean(record.get("hostname"))
        if not host:
            continue
        index = 0
        while f"interface_{index}" in record:
            ifname = clean(record.get(f"interface_{index}"))
            ip = clean(record.get(f"ip_address_{index}"))
            mask = clean(record.get(f"subnet_mask_{index}"))
            if ifname:
                ips[(host.lower(), ifname.lower())] = f"{ip}/{mask}" if ip and mask else None
            index += 1
    return ips


# --------------------------------------------------------------------------
# main
# --------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description="Generate push_ospf.yaml from the protocol_ospf worksheet")
    ap.add_argument("-x", "--xlsx", default="router_interfaces.xlsx", help="workbook path")
    ap.add_argument("-i", "--inventory", default="inventory.ini", help="Ansible inventory")
    ap.add_argument("-s", "--sheet", help="OSPF sheet name (default: auto-detect protocol_ospf)")
    ap.add_argument("-o", "--output", default="push_ospf.yaml", help="playbook to write")
    ap.add_argument("--ignore-missing-hosts", action="store_true",
                    help="skip (with a warning) worksheet hosts that are not in the inventory")
    args = ap.parse_args()

    if not Path(args.xlsx).exists():
        sys.exit(f"ERROR: workbook {args.xlsx} not found")
    wb = openpyxl.load_workbook(args.xlsx, data_only=True)
    ws = find_sheet(wb, args.sheet)
    inventory = inventory_hosts(args.inventory)
    iface_ips = load_interface_ips(wb)
    print(f"Workbook : {args.xlsx}   sheet: {ws.title!r}")
    print(f"Inventory: {args.inventory}   hosts: {', '.join(inventory) or '(none)'}")

    errors, warnings = [], []
    # host -> pid -> {"router_id", "passive": [], "interfaces": OrderedDict(name -> entry)}
    data = OrderedDict()
    rows_read = rows_configured = 0

    for number, rec in read_table(ws):
        where = f"{ws.title}!row {number}"
        proto = clean(rec.get("protocol"))
        if proto and proto.lower() != "ospf":
            warnings.append(f"{where}: protocol {proto!r} is not ospf - skipped")
            continue

        host_raw, ifname = clean(rec.get("hostname")), clean(rec.get("interface"))
        if not host_raw or not ifname:
            errors.append(f"{where}: hostname and interface are both required")
            continue
        rows_read += 1

        host = match_host(host_raw, inventory)
        if host is None:
            (warnings if args.ignore_missing_hosts else errors).append(
                f"{where}: host {host_raw!r} is not in the inventory")
            continue

        pid, area = to_int_or_str(rec.get("process_id")), to_int_or_str(rec.get("area"))
        if pid is None or area is None:
            errors.append(f"{where}: process_id and area are required ({host} {ifname})")
            continue

        proc = data.setdefault(host, OrderedDict()).setdefault(
            pid, {"router_id": None, "passive": [], "interfaces": OrderedDict()})

        router_id = clean(rec.get("router_id"))
        if router_id:
            try:
                ipaddress.IPv4Address(router_id)
            except ValueError:
                errors.append(f"{where}: router_id {router_id!r} is not an IPv4 address")
            if proc["router_id"] and proc["router_id"] != router_id:
                errors.append(f"{where}: {host} process {pid} has conflicting router_id "
                              f"({proc['router_id']} vs {router_id})")
            proc["router_id"] = proc["router_id"] or router_id

        if ifname.lower() in (i.lower() for i in proc["interfaces"]):
            errors.append(f"{where}: {host} {ifname} listed twice for process {pid}")
            continue

        lines = [f"ip ospf {pid} area {area}"]
        cost = to_int_or_str(rec.get("cost"))
        if cost is not None:
            lines.append(f"ip ospf cost {cost}")

        # network_type blank/null -> NO "ip ospf network" line at all
        net_type = clean(rec.get("network_type"))
        if net_type is not None:
            net_type = " ".join(net_type.lower().split())
            if net_type not in VALID_NETWORK_TYPES:
                errors.append(f"{where}: network_type {net_type!r} invalid; "
                              f"use one of {sorted(VALID_NETWORK_TYPES)}")
            lines.append(f"ip ospf network {net_type}")

        try:
            if to_bool(rec.get("passive"), where):
                proc["passive"].append(ifname)
        except ValueError as exc:
            errors.append(str(exc))

        proc["interfaces"][ifname] = {"name": ifname, "lines": lines}
        rows_configured += 1

        # ---- sanity checks against the Interfaces sheet (warnings only) ----
        key = (host.lower(), ifname.lower())
        if iface_ips:
            if key not in iface_ips:
                warnings.append(f"{where}: {host} {ifname} is not on the Interfaces sheet")
            elif iface_ips[key] is None and not ifname.lower().startswith("loopback"):
                warnings.append(f"{where}: {host} {ifname} has no IP address on the Interfaces "
                                "sheet - OSPF will not run on an interface without an IP")
        subnet = clean(rec.get("subnet"))
        if subnet and iface_ips.get(key):
            try:
                if (ipaddress.ip_interface(iface_ips[key]).ip
                        not in ipaddress.ip_network(subnet, strict=False)):
                    warnings.append(f"{where}: subnet {subnet} does not contain {host} {ifname} "
                                    f"address {iface_ips[key]} (subnet column is informational only)")
            except ValueError:
                warnings.append(f"{where}: could not parse subnet {subnet!r}")

    for host in inventory:
        if host not in data:
            warnings.append(f"inventory host {host} has no OSPF rows in {ws.title!r} - not in playbook")
    for host, procs in data.items():
        for pid, proc in procs.items():
            if not proc["router_id"]:
                warnings.append(f"{host} process {pid}: no router_id given; IOS will pick one itself")

    if rows_read != rows_configured and not args.ignore_missing_hosts:
        errors.append(f"coverage check failed: {rows_read} rows read, {rows_configured} configured")

    for w in warnings:
        print(f"WARNING: {w}")
    if errors:
        for e in errors:
            print(f"ERROR  : {e}", file=sys.stderr)
        sys.exit(f"\nAborted - {len(errors)} error(s); playbook not written.")
    if not data:
        sys.exit("ERROR: no OSPF configuration to write")

    # ---- build the per-host data embedded in the playbook ----
    ospf_config = OrderedDict()
    for host, procs in data.items():
        plist = []
        for pid, proc in procs.items():
            router_lines = []
            if proc["router_id"]:
                router_lines.append(f"router-id {proc['router_id']}")
            router_lines += [f"passive-interface {i}" for i in proc["passive"]]
            plist.append({"process_id": pid,
                          "router_lines": router_lines,
                          "interfaces": list(proc["interfaces"].values())})
        ospf_config[host] = {"processes": plist}

    # Host order follows the inventory so the play pattern is predictable.
    play_hosts = [h for h in inventory if h in ospf_config]
    vars_yaml = to_yaml({"ospf_config": dict(ospf_config)}) + "\n"

    header = f"""\
---
# push_ospf.yaml  -  GENERATED by ospf_from_xlsx.py - do not edit by hand.
# Source : {Path(args.xlsx).name}  [sheet: {ws.title}]
# Hosts  : {', '.join(play_hosts)}
#
#   ansible-playbook -i {args.inventory} push_ospf.yaml --check --diff   # dry run
#   ansible-playbook -i {args.inventory} push_ospf.yaml
#
# Requires: ansible-galaxy collection install cisco.ios ansible.netcommon
# Inventory must set ansible_connection=ansible.netcommon.network_cli,
# ansible_network_os=cisco.ios.ios and credentials for the routers.
#
# Interfaces whose network_type is blank in the worksheet get NO
# "ip ospf network" line. Changing a router-id on a running process needs
# "clear ip ospf process" (or a reload) to take effect.

- name: Push OSPF configuration from protocol_ospf worksheet
  hosts: {','.join(play_hosts)}
  gather_facts: false
  connection: ansible.netcommon.network_cli

  vars:
    ansible_network_os: cisco.ios.ios
"""
    with open(args.output, "w") as fh:
        fh.write(header)
        fh.write(indent(vars_yaml, "    "))
        fh.write("\n")
        fh.write(PLAYBOOK_TASKS)

    # self-check the generated file when PyYAML is available
    if yaml is not None:
        yaml.safe_load(Path(args.output).read_text())

    for host in play_hosts:
        n = sum(len(p["interfaces"]) for p in ospf_config[host]["processes"])
        print(f"  {host}: {n} interface(s)")
    print(f"\nWrote {args.output}")
    print(f"Coverage: {rows_configured}/{rows_read} worksheet interface rows configured.")


if __name__ == "__main__":
    main()
