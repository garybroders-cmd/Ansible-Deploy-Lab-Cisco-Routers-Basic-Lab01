#!/usr/bin/env python3
"""Read BGP data for one host from the workbook, validate it, and print JSON
(Ansible vars) or write a host_vars file.

Works with .xlsx and .xlsm (macro-enabled) workbooks. Macros are ignored; the
script only reads cell values.

Usage:
  bgp_from_xlsx.py router_interfaces.xlsm iVOS1               # JSON to stdout
  bgp_from_xlsx.py router_interfaces.xlsm iVOS1 --yaml host_vars/iVOS1.yml
  bgp_from_xlsx.py router_interfaces.xlsm --check-all         # validate every router, no output

Validation errors stop the script (exit 1) so Ansible never pushes bad config.
Warnings go to stderr and do not stop it.
"""
import argparse, ipaddress, json, os, re, sys, warnings, zipfile
import openpyxl
from openpyxl.utils.exceptions import InvalidFileException

warnings.filterwarnings("ignore", message="Data Validation extension")
warnings.filterwarnings("ignore", message="Unknown extension")

SUPPORTED_EXT = (".xlsx", ".xlsm", ".xltx", ".xltm")
ROLES = {"reflector", "client", "peer"}
YESNO = {"yes", "no", "y", "n", "true", "false", "1", "0"}
IFACE_RE = re.compile(r"^[A-Za-z][A-Za-z\-]*\d+(/\d+)*(\.\d+)?$")  # Loopback0, Gi0/1, Gi0/0.10


def yes(v):
    return str(v).strip().lower() in ("yes", "y", "true", "1")


def clean(v):
    if v is None or str(v).strip() == "":
        return None
    if isinstance(v, float) and v.is_integer():
        return int(v)
    return v if isinstance(v, int) else str(v).strip()


def read_sheet(wb, name):
    if name not in wb.sheetnames:
        return None
    ws = wb[name]
    # strip header whitespace and ignore columns with a blank header
    hdr = [str(c.value).strip() if c.value is not None else None for c in ws[1]]
    out = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        if all(v is None or str(v).strip() == "" for v in r):
            continue
        out.append({k: clean(v) for k, v in zip(hdr, r) if k})
    return out


def is_ip(v):
    try:
        ipaddress.IPv4Address(str(v))
        return True
    except ValueError:
        return False


def is_asn(v):
    try:
        return 1 <= int(v) <= 4294967295
    except (TypeError, ValueError):
        return False


def validate(host, b, groups, neighbors, errors, warns):
    e = lambda m: errors.append(f"{host}: {m}")
    w = lambda m: warns.append(f"{host}: {m}")

    # --- protocol_bgp row ---
    if not is_asn(b.get("asn")):
        e(f"asn '{b.get('asn')}' is missing or not a valid AS number")
    if not is_ip(b.get("router_id")):
        e(f"router_id '{b.get('router_id')}' is missing or not a valid IPv4 address")
    role = (b.get("role") or "").lower()
    if role not in ROLES:
        e(f"role '{b.get('role')}' must be one of {sorted(ROLES)}")
    for col in ("log_neighbor_changes", "default_ipv4_unicast"):
        if b.get(col) is not None and str(b[col]).lower() not in YESNO:
            e(f"{col} '{b[col]}' must be yes or no")
    if role == "reflector":
        if not b.get("cluster_id"):
            e("role is reflector but cluster_id is blank")
        elif not is_ip(b["cluster_id"]):
            e(f"cluster_id '{b['cluster_id']}' is not a valid IPv4 address")
    elif b.get("cluster_id"):
        w(f"cluster_id is set but role is '{role}', so it will be ignored")

    # --- peer-groups ---
    seen_pg = set()
    for g in groups:
        name = g.get("peer_group")
        if not name:
            e("a bgp_peer_groups row has no peer_group name")
            continue
        if name in seen_pg:
            e(f"peer-group {name} is defined more than once")
        seen_pg.add(name)
        if not is_asn(g.get("remote_as")):
            e(f"peer-group {name}: remote_as '{g.get('remote_as')}' is missing or invalid")
        us = g.get("update_source")
        if us and (is_ip(us) or not IFACE_RE.match(str(us))):
            e(f"peer-group {name}: update_source '{us}' must be an interface name such as Loopback0, not an IP")
        rrc = g.get("route_reflector_client")
        if rrc is not None and str(rrc).lower() not in YESNO:
            e(f"peer-group {name}: route_reflector_client '{rrc}' must be yes or no")
        if yes(rrc) and role != "reflector":
            e(f"peer-group {name}: route_reflector_client=yes but role is '{role}' (needs reflector)")
    if role == "reflector" and not any(yes(g.get("route_reflector_client")) for g in groups) \
            and not neighbors:
        w("role is reflector but no peer-group has route_reflector_client = yes")

    # --- neighbors ---
    seen_ip = set()
    rid = b.get("router_id")
    for n in neighbors:
        ip = n.get("neighbor_ip")
        if not is_ip(ip):
            e(f"neighbor_ip '{ip}' is missing or not a valid IPv4 address")
            continue
        if ip in seen_ip:
            e(f"neighbor {ip} is listed more than once")
        seen_ip.add(ip)
        if ip == rid:
            e(f"neighbor {ip} is this router's own router_id (row under the wrong hostname?)")
        pgn = n.get("peer_group")
        if pgn and pgn not in seen_pg:
            e(f"neighbor {ip}: peer_group '{pgn}' is not defined in bgp_peer_groups for this host")
        if not pgn and not is_asn(n.get("remote_as")):
            e(f"neighbor {ip}: needs a valid remote_as (or a peer_group)")
        if pgn and n.get("remote_as"):
            w(f"neighbor {ip}: remote_as ignored because it inherits from peer-group {pgn}")
        us = n.get("update_source")
        if us and (is_ip(us) or not IFACE_RE.match(str(us))):
            e(f"neighbor {ip}: update_source '{us}' must be an interface name such as Loopback0, not an IP")
        act = n.get("activate")
        if act is not None and str(act).lower() not in YESNO:
            e(f"neighbor {ip}: activate '{act}' must be yes or no")
        # typo catcher: loopbacks normally share a prefix (10.255.0.x)
        if is_ip(rid) and ipaddress.ip_address(ip) not in ipaddress.ip_network(f"{rid}/24", strict=False):
            w(f"neighbor {ip} is outside {ipaddress.ip_network(f'{rid}/24', strict=False)} (router_id {rid}); typo?")
        # iBGP/eBGP sanity
        ras = n.get("remote_as") if not pgn else next(
            (g.get("remote_as") for g in groups if g.get("peer_group") == pgn), None)
        if is_asn(ras) and is_asn(b.get("asn")) and int(ras) != int(b["asn"]) and role in ("reflector",):
            w(f"neighbor {ip}: remote_as {ras} differs from local {b['asn']} (eBGP) on a reflector")


def open_workbook(path):
    """Open .xlsx/.xlsm/.xltx/.xltm for reading, with clear errors."""
    if not os.path.isfile(path):
        sys.exit(f"Workbook not found: {path}")
    ext = os.path.splitext(path)[1].lower()
    if ext not in SUPPORTED_EXT:
        sys.exit(f"Unsupported file type '{ext}'. Supported: {', '.join(SUPPORTED_EXT)}")
    try:
        # keep_vba is not needed: we only read values, and macros are never executed.
        # data_only=True returns cached formula results (the sheet must have been
        # saved in Excel at least once after any formulas were entered).
        return openpyxl.load_workbook(path, data_only=True, keep_vba=False)
    except (InvalidFileException, zipfile.BadZipFile) as exc:
        sys.exit(f"Could not open {path}: {exc}. Is it corrupt, password-protected, or still open "
                 "as a temporary '~$' lock file?")
    except PermissionError:
        sys.exit(f"Permission denied reading {path}")


def load(path):
    wb = open_workbook(path)
    for s in ("protocol_bgp",):
        if s not in wb.sheetnames:
            sys.exit(f"Sheet '{s}' not found in {path}. Sheets present: {wb.sheetnames}. "
                     "Use the updated workbook that contains the BGP sheets.")
    return (read_sheet(wb, "protocol_bgp"),
            read_sheet(wb, "bgp_peer_groups") or [],
            read_sheet(wb, "bgp_neighbors") or [])


def build(host, base_rows, group_rows, neigh_rows, errors, warns):
    mine = [r for r in base_rows if r.get("hostname") == host]
    if not mine:
        errors.append(f"{host}: no protocol_bgp row with this hostname (check spelling/case and blank cells)")
        return None
    if len(mine) > 1:
        errors.append(f"{host}: appears {len(mine)} times in protocol_bgp (must be one row per router)")
    b = mine[0]
    strip = lambda rows: [{k: v for k, v in r.items() if k != "hostname"}
                          for r in rows if r.get("hostname") == host]
    groups, neighbors = strip(group_rows), strip(neigh_rows)
    validate(host, b, groups, neighbors, errors, warns)
    role = (b.get("role") or "").lower()
    for g in groups:
        g["route_reflector_client"] = yes(g.get("route_reflector_client"))
    for n in neighbors:
        n["activate"] = yes(n.get("activate"))
    return {
        "asn": b.get("asn"), "router_id": b.get("router_id"),
        "cluster_id": b.get("cluster_id") if role == "reflector" else None,
        "role": role,
        "log_neighbor_changes": yes(b.get("log_neighbor_changes")),
        "default_ipv4_unicast": yes(b.get("default_ipv4_unicast")),
        "peer_groups": groups, "neighbors": neighbors,
    }


def report(errors, warns, strict=False):
    if strict:
        errors = errors + [f"(strict) {m}" for m in warns]
        warns = []
    for m in warns:
        print(f"WARNING: {m}", file=sys.stderr)
    if errors:
        for m in errors:
            print(f"ERROR: {m}", file=sys.stderr)
        sys.exit(f"{len(errors)} validation error(s) found; nothing was generated.")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("workbook", help="path to the .xlsx or .xlsm workbook")
    ap.add_argument("host", nargs="?")
    ap.add_argument("--yaml", help="write {bgp: ...} to this host_vars file")
    ap.add_argument("--strict", action="store_true",
                    help="treat warnings (e.g. neighbor outside the router_id /24) as errors")
    ap.add_argument("--check-all", action="store_true", help="validate every router in protocol_bgp")
    a = ap.parse_args()

    base, groups, neighbors = load(a.workbook)
    errors, warns = [], []

    # catch rows whose hostname matches no router
    known = {r.get("hostname") for r in base}
    for label, rows in (("bgp_peer_groups", groups), ("bgp_neighbors", neighbors)):
        for r in rows:
            if r.get("hostname") not in known:
                errors.append(f"{label}: hostname '{r.get('hostname')}' has no row in protocol_bgp")

    if a.check_all:
        for h in sorted(x for x in known if x):
            build(h, base, groups, neighbors, errors, warns)
        report(errors, warns, a.strict)
        print(f"OK: {len(known)} router(s) validated", file=sys.stderr)
        sys.exit(0)

    if not a.host:
        ap.error("host is required unless --check-all is used")
    data = build(a.host, base, groups, neighbors, errors, warns)
    report(errors, warns, a.strict)
    if a.yaml:
        import yaml
        with open(a.yaml, "w") as f:
            yaml.safe_dump({"bgp": data}, f, sort_keys=False)
    else:
        print(json.dumps({"bgp": data}, indent=2))
