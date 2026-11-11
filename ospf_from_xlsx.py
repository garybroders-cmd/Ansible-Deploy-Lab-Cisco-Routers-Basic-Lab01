import argparse
import ipaddress
import os
import re
import sys
from collections import OrderedDict

from openpyxl import load_workbook

ABBREV = [
    (r"^(gigabitethernet|gig|gi)(?=\d)", "GigabitEthernet"),
    (r"^(fastethernet|fa)(?=\d)", "FastEthernet"),
    (r"^(tengigabitethernet|te)(?=\d)", "TenGigabitEthernet"),
    (r"^(loopback|lo)(?=\d)", "Loopback"),
    (r"^(ethernet|eth)(?=\d)", "Ethernet"),
]


def norm_if(name):
    """Expand short names (Gi0/1 -> GigabitEthernet0/1) so sheets compare equal."""
    name = str(name).strip().replace(" ", "")
    for pattern, full in ABBREV:
        m = re.match(pattern, name, re.IGNORECASE)
        if m:
            return full + name[m.end():]
    return name


def clean(v):
    if v is None:
        return None
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    v = str(v).strip()
    return v or None


def read_sheet(wb, sheet):
    if sheet not in wb.sheetnames:
        sys.exit(f"ERROR: sheet '{sheet}' not found. Sheets available: {wb.sheetnames}")
    rows = list(wb[sheet].iter_rows(values_only=True))
    if not rows:
        sys.exit(f"ERROR: sheet '{sheet}' is empty")
    header = [str(h).strip().lower() if h is not None else "" for h in rows[0]]
    out = []
    for row in rows[1:]:
        if not any(v is not None and str(v).strip() for v in row):
            continue
        out.append({k: clean(v) for k, v in zip(header, row) if k})
    return header, out


def interface_ips(wb, sheet):
    """host -> {interface: ip or None} from the wide-format Interfaces sheet."""
    header, rows = read_sheet(wb, sheet)
    if "hostname" not in header:
        sys.exit(f"ERROR: sheet '{sheet}' has no 'hostname' column. Header row: {header}")
    result = {}
    for r in rows:
        host = r.get("hostname")
        if not host:
            continue
        ifs = result.setdefault(host, {})
        for key, val in r.items():
            m = re.match(r"^interface_(\d+)$", key)
            if m and val:
                ifs[norm_if(val)] = r.get(f"ip_address_{m.group(1)}")
    return result


SUBNET_COLS = ("subnet", "subnets", "network", "networks")
MAX_SUBNETS = 10


def _is_mask(bits):
    inv = bits ^ 0xFFFFFFFF
    return inv & (inv + 1) == 0          # contiguous 1s then 0s


def _is_wildcard(bits):
    return bits & (bits + 1) == 0        # contiguous 0s then 1s


def parse_subnet(text):
    """Return (IPv4Network, normalised_flag). Accepts CIDR, addr+mask, addr+wildcard, bare host."""
    t = text.strip()
    if "/" in t:
        spec = t
    else:
        parts = t.split()
        if len(parts) == 1:
            spec = parts[0] + "/32"
        elif len(parts) == 2:
            addr, second = parts
            bits = int(ipaddress.IPv4Address(second))
            if second == "0.0.0.0":
                prefix = 32                                   # host wildcard
            elif _is_mask(bits):
                prefix = bin(bits).count("1")
            elif _is_wildcard(bits):
                prefix = 32 - bin(bits).count("1")
            else:
                raise ValueError(f"'{second}' is not a valid mask or wildcard")
            spec = f"{addr}/{prefix}"
        else:
            raise ValueError("expected CIDR, 'address mask' or 'address wildcard'")
    try:
        return ipaddress.IPv4Network(spec, strict=True), False
    except ValueError as exc:
        if "host bits" not in str(exc):
            raise
        return ipaddress.IPv4Network(spec, strict=False), True


def split_subnets(cell):
    return [x.strip() for x in re.split(r"[,;\r\n]+", cell) if x.strip()]


def truthy(v):
    return v is not None and str(v).strip().lower() in ("yes", "y", "true", "1")


def build(host, pid, rows, known_ifs, warnings):
    router_id = next((r["router_id"] for r in rows if r.get("router_id")), None)
    passive, iface_lines, networks, seen = [], [], [], {}

    for r in rows:
        ifname = norm_if(r["interface"]) if r.get("interface") else None
        area = r.get("area")
        label = f"{host} {ifname or '(subnet row)'}"

        if area is None:
            warnings.append(f"{label}: area is blank, row skipped")
            continue

    # ---- interface-level OSPF ----
    if ifname:
        use_if = True

        if host in known_ifs:
            if ifname in known_ifs[host]:
                if not known_ifs[host][ifname]:
                    warnings.append(
                        f"{label}: IP is blank on the Interfaces sheet "
                        "(interface will be shut down), interface OSPF skipped"
                    )
                    use_if = False
            elif not ifname.lower().startswith("loopback"):
                warnings.append(
                    f"{label}: not found on the Interfaces sheet, "
                    "OSPF written anyway"
                )
        else:
            warnings.append(
                f"{host}: not on the Interfaces sheet at all"
            )

        if use_if:
            block = [f"interface {ifname}"]

            network_type = str(
                r.get("network_type") or ""
            ).strip()

            if network_type:
                block.append(
                    f" ip ospf network {network_type}"
                )
            cost = r.get("cost")
            if cost not in (None, "", "nan"):
                block.append(
                    f" ip ospf cost {cost}"
                )

            block.append(
                f" ip ospf {pid} area {area}"
            )

            block.append("!")

            iface_lines.extend(block)

            if truthy(r.get("passive")):
                passive.append(ifname)

        # ---- subnet -> router ospf network statement ----
        cell = next((r[c] for c in SUBNET_COLS if r.get(c)), None)
        items = split_subnets(cell) if cell else []
        if len(items) > MAX_SUBNETS:
            warnings.append(
                f"{label}: {len(items)} subnets in one cell, only the first {MAX_SUBNETS} are used "
                f"(ignored: {', '.join(items[MAX_SUBNETS:])})")
            items = items[:MAX_SUBNETS]
        for item in items:
            try:
                net, normalised = parse_subnet(item)
            except ValueError as exc:
                warnings.append(f"{label}: subnet '{item}' not understood ({exc}), skipped")
                continue
            if normalised:
                warnings.append(f"{label}: subnet '{item}' had host bits set, used {net}")
            if net in seen:
                if seen[net] != str(area):
                    warnings.append(f"{host}: {net} listed in areas {seen[net]} and {area}, first one kept")
                continue
            seen[net] = str(area)
            networks.append((net, area))

    if not iface_lines and not networks:
        return None

    lines = [f"router ospf {pid}"]
    if router_id:
        lines.append(f" router-id {router_id}")
    for p in passive:
        lines.append(f" passive-interface {p}")
    for net, area in networks:
        lines.append(f" ! {net}")
        lines.append(f" network {net.network_address} {net.hostmask} area {area}")
    lines.append("!")
    return "\n".join(lines + iface_lines) + "\n"


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("workbook", nargs="?", default="router_interfaces.xlsx")
    ap.add_argument("--sheet", default="Protocol")
    ap.add_argument("--interfaces-sheet", default="Interfaces")
    ap.add_argument("--out", default="ospf_configs")
    ap.add_argument("--host", help="only build this device")
    args = ap.parse_args()

    if not os.path.isfile(args.workbook):
        sys.exit(f"ERROR: workbook not found: {args.workbook}")
    wb = load_workbook(args.workbook, data_only=True)

    header, rows = read_sheet(wb, args.sheet)
    for col in ("hostname", "interface", "protocol"):
        if col not in header:
            sys.exit(f"ERROR: sheet '{args.sheet}' has no '{col}' column. Header row: {header}")

    if not any(c in header for c in SUBNET_COLS):
        print(f"NOTE: no 'subnet' column found on sheet '{args.sheet}' (header: {header}), "
              "so no network statements will be generated.", file=sys.stderr)

    known_ifs = interface_ips(wb, args.interfaces_sheet) if args.interfaces_sheet in wb.sheetnames else {}

    groups = OrderedDict()
    for r in rows:
        if (r.get("protocol") or "").lower() != "ospf":
            continue
        host = r.get("hostname")
        if not host or (args.host and host != args.host):
            continue
        groups.setdefault((host, r.get("process_id") or "1"), []).append(r)

    if not groups:
        sys.exit("No OSPF rows found" + (f" for host '{args.host}'" if args.host else "") + ".")

    os.makedirs(args.out, exist_ok=True)
    warnings, per_host = [], OrderedDict()
    for (host, pid), host_rows in groups.items():
        text = build(host, pid, host_rows, known_ifs, warnings)
        if text:
            per_host.setdefault(host, []).append(text)

    for host, blocks in per_host.items():
        path = os.path.join(args.out, f"{host}.cfg")
        with open(path, "w") as f:
            f.write("\n".join(blocks))
        print(f"===== {host}  ->  {path} =====")
        print("\n".join(blocks))

    if warnings:
        print("WARNINGS:", file=sys.stderr)
        for w in warnings:
            print(f"  - {w}", file=sys.stderr)


if __name__ == "__main__":
    main()
