#!/usr/bin/env python3
"""
bootstrap_from_xlsx.py

Create one bootstrap config per router from router_interfaces.xlsx, laid out
exactly like the example bootstrap file. Output file name:

    <hostname>_bootstrap.cfg         e.g.  iVOS1_bootstrap.cfg

Files are saved in the "bootstraps" directory (created if needed).

Usage
    python3 bootstrap_from_xlsx.py -x router_interfaces.xlsx
    python3 bootstrap_from_xlsx.py -x router_interfaces.xlsx --host iVOS1
    python3 bootstrap_from_xlsx.py -x router_interfaces.xlsx -o other_dir/   # override the directory

Where each value comes from
    hostname              Interfaces sheet, "hostname" column
    management interface  Interfaces sheet: the interface whose description is
                          "Management" (change with --mgmt-description, or force
                          one interface for every router with --mgmt-interface)
    management IP / mask  Interfaces sheet, same row as that interface
    domain name           --domain          (default lab.local)
    username / privilege  --username / --privilege   (default ansible / 15)
    user password         --password  or env BOOTSTRAP_PASSWORD
    enable secret         --enable-secret or env BOOTSTRAP_ENABLE_SECRET
The last four are not in the workbook, so they are options. The defaults match
the example file; CHANGE THE PASSWORDS for anything but a lab.

Only openpyxl is required.
"""

import argparse
import ipaddress
import os
import re
import sys
from pathlib import Path

import openpyxl

TEMPLATE = """\
! Boot strap file


enable
configure terminal
hostname {hostname}
ip domain-name {domain}
!
username {username} privilege {privilege} secret {password}
!
interface {mgmt_if}
 description {mgmt_desc}
 ip address {mgmt_ip} {mgmt_mask}
 no shutdown
!
crypto key generate rsa modulus {rsa_bits}
ip ssh version 2
!
line vty 0 4
 login local
 transport input ssh
!
enable secret {enable_secret}
end
write memory
"""

BLANK = {"", "none", "null", "nan", "n/a", "na", "-"}


def clean(value):
    if value is None:
        return None
    text = str(value).strip()
    return None if text.lower() in BLANK else text


def find_interfaces_sheet(wb, requested):
    names = {ws.title.lower(): ws.title for ws in wb.worksheets}
    if requested:
        if requested.lower() not in names:
            sys.exit(f"ERROR: sheet {requested!r} not found. Sheets: {wb.sheetnames}")
        return wb[names[requested.lower()]]
    return wb[names["interfaces"]] if "interfaces" in names else wb.worksheets[0]


def read_interfaces(ws):
    """Return {hostname: [ {name, ip, mask, description}, ... ]} (file order kept).

    Supports the wide layout used in the workbook (interface_0, ip_address_0,
    subnet_mask_0, description_0, interface_1, ...) and a long layout with one
    row per interface (hostname, interface, ip_address, subnet_mask, description).
    """
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        sys.exit(f"ERROR: sheet {ws.title!r} is empty")
    header = [str(h).strip().lower() if h is not None else "" for h in rows[0]]
    if "hostname" not in header:
        sys.exit(f"ERROR: sheet {ws.title!r} has no 'hostname' column")

    routers = {}
    for number, row in enumerate(rows[1:], start=2):
        rec = dict(zip(header, row))
        host = clean(rec.get("hostname"))
        if not host:
            continue  # blank row
        interfaces = routers.setdefault(host, [])

        if "interface" in header:  # long layout
            name = clean(rec.get("interface"))
            if name:
                interfaces.append({"name": name, "ip": clean(rec.get("ip_address")),
                                   "mask": clean(rec.get("subnet_mask")),
                                   "description": clean(rec.get("description")), "row": number})
        else:  # wide layout
            index = 0
            while f"interface_{index}" in rec:
                name = clean(rec.get(f"interface_{index}"))
                if name:
                    interfaces.append({"name": name,
                                       "ip": clean(rec.get(f"ip_address_{index}")),
                                       "mask": clean(rec.get(f"subnet_mask_{index}")),
                                       "description": clean(rec.get(f"description_{index}")),
                                       "row": number})
                index += 1
    return routers


def pick_management(host, interfaces, args):
    """Return (interface_dict, error_message)."""
    if args.mgmt_interface:
        found = [i for i in interfaces if i["name"].lower() == args.mgmt_interface.lower()]
        what = f"interface {args.mgmt_interface!r}"
    else:
        want = args.mgmt_description.lower()
        found = [i for i in interfaces if (i["description"] or "").lower() == want]
        what = f"an interface with description {args.mgmt_description!r}"
    if not found:
        return None, f"{host}: {what} not found on the Interfaces sheet"
    if len(found) > 1:
        names = ", ".join(i["name"] for i in found)
        return None, f"{host}: more than one match for {what} ({names}); use --mgmt-interface"
    return found[0], None


def main():
    ap = argparse.ArgumentParser(description="Create <hostname>_bootstrap.cfg for each router in ./bootstraps")
    ap.add_argument("-x", "--xlsx", default="router_interfaces.xlsx", help="workbook path")
    ap.add_argument("-s", "--sheet", help="interfaces sheet (default: 'Interfaces', else first sheet)")
    ap.add_argument("-o", "--out-dir", default="bootstraps",
                    help="directory for the .cfg files (default: bootstraps)")
    ap.add_argument("--host", action="append", metavar="NAME",
                    help="only this router (repeatable); default: every router in the sheet")
    ap.add_argument("--mgmt-description", default="Management",
                    help="description that marks the management interface (default: Management)")
    ap.add_argument("--mgmt-interface", help="use this interface name for every router instead")
    ap.add_argument("--domain", default="lab.local")
    ap.add_argument("--username", default="ansible")
    ap.add_argument("--privilege", type=int, default=15)
    ap.add_argument("--password", default=os.environ.get("BOOTSTRAP_PASSWORD", "StrongPass123"))
    ap.add_argument("--enable-secret", default=os.environ.get("BOOTSTRAP_ENABLE_SECRET", "StrongPass123"))
    ap.add_argument("--rsa-bits", type=int, default=2048)
    ap.add_argument("--name-format", default="{hostname}_bootstrap.cfg",
                    help="output file name pattern (default: {hostname}_bootstrap.cfg)")
    args = ap.parse_args()

    if not Path(args.xlsx).exists():
        sys.exit(f"ERROR: workbook {args.xlsx} not found")
    wb = openpyxl.load_workbook(args.xlsx, data_only=True)
    ws = find_interfaces_sheet(wb, args.sheet)
    routers = read_interfaces(ws)
    if not routers:
        sys.exit(f"ERROR: no routers found on sheet {ws.title!r}")

    if args.host:
        wanted = {h.lower() for h in args.host}
        unknown = wanted - {h.lower() for h in routers}
        if unknown:
            sys.exit(f"ERROR: not on sheet {ws.title!r}: {', '.join(sorted(unknown))}")
        routers = {h: i for h, i in routers.items() if h.lower() in wanted}

    print(f"Workbook: {args.xlsx}   sheet: {ws.title!r}   routers: {', '.join(routers)}")

    errors, warnings, rendered = [], [], {}
    for host, interfaces in routers.items():
        if not re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9_-]*", host):
            errors.append(f"{host}: hostname contains characters not valid in a device hostname")
            continue
        mgmt, err = pick_management(host, interfaces, args)
        if err:
            errors.append(err)
            continue
        if not mgmt["ip"] or not mgmt["mask"]:
            errors.append(f"{host}: management interface {mgmt['name']} has no IP address / mask "
                          f"(Interfaces row {mgmt['row']})")
            continue
        try:
            ipaddress.IPv4Address(mgmt["ip"])
            ipaddress.IPv4Network(f"0.0.0.0/{mgmt['mask']}")  # validates a dotted mask
        except ValueError:
            errors.append(f"{host}: invalid management address {mgmt['ip']} / {mgmt['mask']}")
            continue

        rendered[host] = TEMPLATE.format(
            hostname=host, domain=args.domain, username=args.username,
            privilege=args.privilege, password=args.password,
            mgmt_if=mgmt["name"], mgmt_desc=mgmt["description"] or args.mgmt_description,
            mgmt_ip=mgmt["ip"], mgmt_mask=mgmt["mask"],
            rsa_bits=args.rsa_bits, enable_secret=args.enable_secret)

    if errors:
        for e in errors:
            print(f"ERROR  : {e}", file=sys.stderr)
        sys.exit(f"\nAborted - {len(errors)} error(s); no files written.")

    if args.password == "StrongPass123" or args.enable_secret == "StrongPass123":
        warnings.append("using the example password/secret 'StrongPass123' - set --password / "
                        "--enable-secret (or BOOTSTRAP_PASSWORD / BOOTSTRAP_ENABLE_SECRET)")

    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    for host, text in rendered.items():
        target = out_dir / args.name_format.format(hostname=host)
        target.write_text(text)
        target.chmod(0o600)  # the file contains passwords
        print(f"Wrote {target}")
    for w in warnings:
        print(f"WARNING: {w}")
    print(f"\n{len(rendered)} bootstrap file(s) written.")


if __name__ == "__main__":
    main()
