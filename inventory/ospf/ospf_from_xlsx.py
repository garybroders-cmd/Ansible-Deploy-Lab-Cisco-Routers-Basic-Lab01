#!/usr/bin/env python3
"""
ospf_from_xlsx.py <workbook.xlsm> <hostname> [--sheet NAME]

Reads the OSPF data for ONE host from the workbook and prints it as JSON:

    {"ospf": {"processes": [{"process_id": 1, "router_id": "10.255.0.1",
                             "passive": [], "interfaces": [
                               {"name": "GigabitEthernet0/0", "area": 0,
                                "cost": 100, "network_type": "point-to-point"}, ...]}]}}

Used by ospf_playbook.yaml in the same way bgp_from_xlsx.py is used by bgp_playbook.yaml.

Sources
  * protocol_ospf sheet, range A2:J60 (header on row 1):
      hostname, interface, protocol, process_id, router_id, area, cost,
      network_type, passive, subnet
    A blank hostname cell takes the hostname from the row above.
  * Interfaces sheet (hostname, interface_N / ip_address_N / subnet_mask_N / description_N):
      - any interface with an IP address that is NOT listed on protocol_ospf is added
        with the host's own process / area / cost (management interfaces are skipped);
      - a host with no protocol_ospf rows at all is built entirely from this sheet
        (router-id = its Loopback0 address).

Rules
  * Loopback interfaces NEVER get an "ip ospf network" line.
  * A physical /30 or /31 interface with a blank network_type gets point-to-point.
  * Hostnames match the inventory name ignoring case, punctuation and domain suffix
    (ivos3 = iVOS-3 = ivos3.lab.local).
  * Warnings go to stderr; stdout is JSON only. Exit code 2 if the host is not found.
"""
import argparse
import ipaddress
import json
import re
import sys
import warnings
from collections import OrderedDict

import openpyxl

warnings.filterwarnings("ignore")

FIRST_ROW, LAST_ROW = 2, 60
SHEET_CANDIDATES = ["protocol_ospf", "protocol_opsf", "ospf", "protocol"]
REQUIRED = ["hostname", "interface", "process_id", "area"]
TRUE_TOKENS = {"yes", "y", "true", "t", "1", "passive"}
FALSE_TOKENS = {"", "no", "n", "false", "f", "0"}
VALID_NETWORK_TYPES = {"broadcast", "non-broadcast", "point-to-point",
                       "point-to-multipoint", "point-to-multipoint non-broadcast"}
IFACE_PREFIXES = OrderedDict([
    ("loopback", "Loopback"), ("loop", "Loopback"), ("lo", "Loopback"),
    ("gigabitethernet", "GigabitEthernet"), ("gig", "GigabitEthernet"), ("gi", "GigabitEthernet"),
    ("fastethernet", "FastEthernet"), ("fa", "FastEthernet"),
    ("tengigabitethernet", "TenGigabitEthernet"), ("te", "TenGigabitEthernet"),
    ("ethernet", "Ethernet"), ("eth", "Ethernet"),
    ("tunnel", "Tunnel"), ("tu", "Tunnel"), ("vlan", "Vlan"), ("vl", "Vlan"),
    ("port-channel", "Port-channel"), ("po", "Port-channel"),
])


def warn(msg):
    print(f"WARNING: {msg}", file=sys.stderr)


def die(msg, code=1):
    print(f"ERROR: {msg}", file=sys.stderr)
    sys.exit(code)


def clean(value):
    if value is None:
        return None
    text = str(value).strip()
    return text or None


def to_num(value):
    text = clean(value)
    if text is None:
        return None
    try:
        number = float(text)
        return int(number) if number == int(number) else text
    except ValueError:
        return text


def norm_iface(name):
    if name is None:
        return None
    text = re.sub(r"\s+", "", name)
    match = re.match(r"^([A-Za-z-]+?)(\d.*)$", text)
    if not match:
        return text
    prefix, rest = match.groups()
    return IFACE_PREFIXES.get(prefix.lower(), prefix) + rest


def _norm(name):
    return re.sub(r"[^a-z0-9]", "", name.lower())


def same_host(sheet_name, inventory_name):
    a, b = sheet_name.strip(), inventory_name.strip()
    return (a.lower() == b.lower()
            or _norm(a) == _norm(b)
            or _norm(a.split(".")[0]) == _norm(b.split(".")[0]))


def find_sheet(wb, requested):
    names = {n.lower(): n for n in wb.sheetnames}
    for cand in ([requested] if requested else []) + SHEET_CANDIDATES:
        if cand and cand.lower() in names:
            return wb[names[cand.lower()]]
    die(f"OSPF sheet not found (looked for {requested or SHEET_CANDIDATES}); "
        f"sheets in workbook: {wb.sheetnames}")


def read_rows(ws, max_col):
    """Header from row 1, data rows FIRST_ROW..LAST_ROW; hostname filled down."""
    header_row = next(ws.iter_rows(min_row=1, max_row=1, max_col=max_col, values_only=True), ())
    header = [str(h).strip().lower() if h is not None else "" for h in header_row]
    out, last_host = [], None
    for number, row in enumerate(ws.iter_rows(min_row=FIRST_ROW, max_row=LAST_ROW,
                                               max_col=max_col, values_only=True), start=FIRST_ROW):
        if all(c is None or str(c).strip() == "" for c in row):
            continue
        rec = dict(zip(header, row))
        rec["_row"] = number
        if clean(rec.get("hostname")):
            last_host = clean(rec["hostname"])
        rec["hostname"] = last_host
        out.append(rec)
    return header, out


def load_interfaces(wb, host):
    """[(ifname, ip, mask, description)] for the host from the Interfaces sheet."""
    names = {n.lower(): n for n in wb.sheetnames}
    if "interfaces" not in names:
        return []
    ws = wb[names["interfaces"]]
    _, rows = read_rows(ws, ws.max_column)
    found = []
    for rec in rows:
        if not rec.get("hostname") or not same_host(rec["hostname"], host):
            continue
        index = 0
        while f"interface_{index}" in rec:
            ifname = norm_iface(clean(rec.get(f"interface_{index}")))
            ip = clean(rec.get(f"ip_address_{index}"))
            mask = clean(rec.get(f"subnet_mask_{index}"))
            if ifname and ip and mask:
                found.append((ifname, ip, mask, clean(rec.get(f"description_{index}")) or ""))
            index += 1
    return found


def prefixlen(ip, mask):
    try:
        return ipaddress.ip_interface(f"{ip}/{mask}").network.prefixlen
    except ValueError:
        return None


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawTextHelpFormatter)
    ap.add_argument("workbook")
    ap.add_argument("hostname")
    ap.add_argument("--sheet", help="OSPF sheet name (default: protocol_ospf)")
    args = ap.parse_args()
    host = args.hostname

    try:
        wb = openpyxl.load_workbook(args.workbook, data_only=True)
    except Exception as exc:
        die(f"cannot open workbook {args.workbook}: {exc}")

    ws = find_sheet(wb, args.sheet)
    header, rows = read_rows(ws, 10)
    missing = [c for c in REQUIRED if c not in header]
    if missing:
        die(f"sheet {ws.title!r} is missing column(s): {missing}")

    mine = [r for r in rows if r["hostname"] and same_host(r["hostname"], host)]
    iface_sheet = load_interfaces(wb, host)
    iface_ip = {n.lower(): (ip, mask) for n, ip, mask, _ in iface_sheet}

    if not mine and not iface_sheet:
        known = sorted({r["hostname"] for r in rows if r["hostname"]})
        die(f"host {host!r} not found on {ws.title!r} (A{FIRST_ROW}:J{LAST_ROW}) or the "
            f"Interfaces sheet. Hostnames on {ws.title!r}: {known}", code=2)

    # Defaults for interfaces not listed on the OSPF sheet: this host's first row,
    # else the first row of any host, else process 1 / area 0 / cost 100.
    basis = mine[0] if mine else (rows[0] if rows else None)
    if basis:
        default = (to_num(basis.get("process_id")), to_num(basis.get("area")), to_num(basis.get("cost")))
    else:
        default = (1, 0, 100)
    if default[0] is None or default[1] is None:
        default = (1, 0, 100)

    procs = OrderedDict()   # pid -> {"router_id", "passive", "interfaces": OrderedDict}

    def proc_for(pid):
        return procs.setdefault(pid, {"process_id": pid, "router_id": None,
                                      "passive": [], "interfaces": OrderedDict()})

    errors = []
    for rec in mine:
        where = f"{ws.title}!row {rec['_row']}"
        ifname = norm_iface(clean(rec.get("interface")))
        if not ifname:
            errors.append(f"{where}: interface is blank for {host}")
            continue
        pid, area = to_num(rec.get("process_id")), to_num(rec.get("area"))
        if pid is None or area is None:
            errors.append(f"{where}: process_id and area are required ({host} {ifname})")
            continue
        proc = proc_for(pid)

        rid = clean(rec.get("router_id"))
        if rid:
            try:
                ipaddress.IPv4Address(rid)
                if proc["router_id"] and proc["router_id"] != rid:
                    errors.append(f"{where}: {host} process {pid} has conflicting router_id "
                                  f"({proc['router_id']} vs {rid})")
                proc["router_id"] = proc["router_id"] or rid
            except ValueError:
                errors.append(f"{where}: router_id {rid!r} is not an IPv4 address")

        is_loop = ifname.lower().startswith("loopback")
        net = clean(rec.get("network_type"))
        net = " ".join(net.lower().split()) if net else None
        if net and net not in VALID_NETWORK_TYPES:
            errors.append(f"{where}: network_type {net!r} invalid; use {sorted(VALID_NETWORK_TYPES)}")
            net = None
        if is_loop:
            if net:
                warn(f"{where}: network_type {net!r} ignored on {ifname} (loopbacks take none)")
            net = None
        elif net is None and ifname.lower() in iface_ip:
            plen = prefixlen(*iface_ip[ifname.lower()])
            if plen is not None and plen >= 30:
                net = "point-to-point"
                warn(f"{where}: {ifname} has no network_type; point-to-point applied (/30 or /31)")

        passive = (clean(rec.get("passive")) or "").lower()
        if passive in TRUE_TOKENS:
            proc["passive"].append(ifname)
        elif passive not in FALSE_TOKENS:
            errors.append(f"{where}: cannot interpret passive value {passive!r} (use yes/no)")

        proc["interfaces"][ifname.lower()] = {"name": ifname, "area": area,
                                              "cost": to_num(rec.get("cost")),
                                              "network_type": net}
    if errors:
        die("\n  ".join(["problems in the workbook:"] + errors))

    # Add IP-addressed interfaces that the OSPF sheet does not list.
    d_pid, d_area, d_cost = default
    if not procs and iface_sheet:
        warn(f"{host}: no rows on {ws.title!r}; building OSPF from the Interfaces sheet "
             f"(process {d_pid}, area {d_area}, cost {d_cost})")
    for ifname, ip, mask, desc in iface_sheet:
        pid = next(iter(procs), d_pid)
        proc = proc_for(pid)
        if ifname.lower() in proc["interfaces"]:
            continue
        if "manage" in desc.lower():
            warn(f"{host} {ifname}: management interface - not added to OSPF")
            continue
        plen = prefixlen(ip, mask)
        net = ("point-to-point" if not ifname.lower().startswith("loopback")
               and plen is not None and plen >= 30 else None)
        proc["interfaces"][ifname.lower()] = {"name": ifname, "area": d_area,
                                              "cost": d_cost, "network_type": net}
        warn(f"{host} {ifname} ({ip}/{plen}) is not on {ws.title!r}; added from the Interfaces sheet")

    for pid, proc in procs.items():
        if not proc["router_id"]:
            lo = iface_ip.get("loopback0")
            if lo:
                proc["router_id"] = lo[0]
                warn(f"{host} process {pid}: no router_id on the sheet; using Loopback0 {lo[0]}")
            else:
                warn(f"{host} process {pid}: no router_id; IOS will choose one itself")

    result = {"ospf": {"processes": [
        {"process_id": p["process_id"], "router_id": p["router_id"], "passive": p["passive"],
         "interfaces": list(p["interfaces"].values())} for p in procs.values()]}}
    print(json.dumps(result, indent=2))


if __name__ == "__main__":
    main()
