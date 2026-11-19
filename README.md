<img width="1465" height="808" alt="image" src="https://github.com/user-attachments/assets/3028b681-ec1a-4bc3-9302-7829ad91fe5c" />




# Ansible-Deploy-Lab-Cisco-Routers-Basic-Lab01

Status: work in progress. The inventory files currently provide the basics: connecting to each router and configuring its WAN and LAN interfaces. More protocols and features will be added as the project develops.</>

I set up a Raspberry Pi 5 as an Ansible server to automate router configuration. Device details live in an Excel spreadsheet, and a Python 3 script reads that data and runs the right playbooks against each router, so there’s no more configuring them one by one.

## Docker Lab Environment (Raspberry Pi 5)

This lab runs on a Raspberry Pi 5 using Docker containers for automation, management, and web application testing.

### Purpose

This environment provides:

- Ansible automation and orchestration</br>


### The Files Purpose
./create_bootstraps_configs - This will look at the router_interafces.xlsm interfaces and write a simple bootstrap configuration for routers to be accessed by the ansible rapsberry pi.  Configs are stored in a directory called bootstraps.</br>

<img width="889" height="514" alt="image" src="https://github.com/user-attachments/assets/a3c41ca8-7518-4642-a3e1-19da943a6a42" />


My Directory Structure running in a python virtual environment. 


Copy these configurations to each router, etc.</br>

./lab1 - Run the Ansible playbook against the devices listed in inventory.ini, then call the playbook for each protocol (for example, bgp_playbook.yaml). I deliberately run them one at a time so I can troubleshoot faults and make changes as I go. Each protocol has its own folder holding its playbook and templates.</br>
more ansi  
The bgp.j2 file is a Jinja2 template that generates each router’s BGP configuration. Rather than writing a separate config for every router, I write one template with placeholders, and Ansible fills in each device’s values (AS number, neighbors, networks) when the playbook runs.</br>

Amended the /ospf Python files so I can see the test LAN on the two relectore routers and advertised via BGP.

<img width="984" height="859" alt="image" src="https://github.com/user-attachments/assets/e59581e9-f074-450a-bc2c-a9ccaac6ee81" />











