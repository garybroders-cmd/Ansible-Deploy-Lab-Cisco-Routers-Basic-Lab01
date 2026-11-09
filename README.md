# Ansible-Deploy-Lab-Cisco-Routers-Basic-Lab01
Ansible to configure a basic router lab. Stage 1 

## Docker Lab Environment (Raspberry Pi 5)

This lab runs on a Raspberry Pi 5 using Docker containers for automation, management, and web application testing.

### Containers

- Flocio-az
- Portainer
- Ansible
- Alpine SSH 1
- Alpine SSH 2
- Alpine SSH 3
- NGINX
- Web Server 1
- Web Server 2

### Purpose

This environment provides:

- Ansible automation and orchestration <-- This is working of its own management interface assigned using mcxvlan (??)
</br>
- SSH-managed Linux hosts for testing
- Web server deployment and configuration management
- NGINX reverse proxy and load balancing
- Container management through Portainer
- Hands-on infrastructure-as-code and DevOps lab experience

### Lab Topology
Test the playbook and connectivity.

ansible-playbook -i inventory.ini test.yaml 

Check the different modules 
sudo python3 ospf_fromxls.py  <-- this assigns the OSPF network to the router and interface according to the spreadsheet.</br>
sudo python3 read_xlsx.py router_interfaces.xlsx <-- reads the interfaces from the Excel spreadsheet workbook called interfaces. </br>


This updates the script on the server.</br>

wc -l ospf_from_xlsx.py</br>
md5sum ospf_frm_xlsx.py</br>
sudo python -m py_compile ospf_from_xlsx.py && echo OK</br>

to configure routers after bootstrap is installed.</br>

sudo ./ansible_bgp_reflector_config <-- this should be executable so that ou can run different modules or delete modules—</br>

sudo chmod 777 ansible_bgp_reflector_config</br>


