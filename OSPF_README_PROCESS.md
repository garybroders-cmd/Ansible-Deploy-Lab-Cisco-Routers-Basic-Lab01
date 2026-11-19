This process turns a spreadsheet into a router OSPF config, with Ansible running the steps. </br>


Fill in the workbook. router_interfaces.xlsm is the single source of truth. The Interfaces sheet lists each router’s interfaces, IPs, masks and descriptions. The protocol_ospf sheet has one row per interface that should run OSPF: hostname, interface, process ID, router ID, area, cost, network type and passive.
Run the playbook per host. ospf_playbook.yaml calls ospf_from_xlsx.py once for each router in the inventory, passing the workbook and the hostname. The hostname has to match the sheet, though case, punctuation and domain suffix are ignored.
Extract and validate the data. The script reads that host’s rows and applies the rules:</br>
A row needs a hostname, process_id and area, and an interface that exists on the Interfaces sheet for that host. Otherwise, it’s skipped with a warning.</br>
Loopbacks never get ip ospf network unless specified. </br>
A physical /30 or /31 link with no network type gets point-to-point.
If a process has no router ID, the script falls back to the Loopback0 address.</br>
Bad router IDs, network types or passive values stop the run with an error.</br>
Output JSON. The script prints the host’s OSPF data as JSON on stdout, and warnings go to stderr. Ansible reads the JSON as a variable.
Render with Jinja. ospf.j2 turns that data into IOS config: a router ospf block with the router ID and passive interfaces, then ip ospf <process> area <area> under each interface, plus cost and network type where set.
Push to the device. Ansible applies the rendered config to the router. After that, you’d check show ip ospf neighbor and show ip route ospf.

This mirrors your BGP process, which uses bgp_from_xlsx.py and bgp_playbook.yaml the same way. Fixing a router now means editing a spreadsheet cell and re-running the playbook, not editing config on the device.</br>
