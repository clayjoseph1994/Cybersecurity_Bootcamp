# ELK_Stack_Project
This is a project I worked on in the cybersecurity bootcamp where several resources were used to deploy an ELK Stack (Elasticsearch, Filebeat [instead of logstash in this instance, so I guess it's more of an EFK Stack], and Kibana). This Stack's intended function is to be a monitoring server to aggregate logs from two web servers that have DVWA (Damn Vulnerable Web Application) installed on them. I utilized Docker and Ansible to install the DVWA and ELK Stack container images on the VMs. I will provide more details for this project within this repository.

## Automated ELK Stack Deployment

The files in this repository were used to configure the network depicted below.

![alt text](https://github.com/clayjoseph1994/ELK_Stack_Project/blob/master/Images/ELK_Stack_TopologyDetails.png) 

[NOTE: the network diagram has differing IP addresses from the ones Azure assigned me by default]

These files have been tested and used to generate a live ELK deployment on Azure. They can be used to either recreate the entire deployment pictured above. Alternatively, select portions of the yml file(s) may be used to install only certain pieces of it, such as Filebeat.

This document contains the following details:
Description of the Topology,
Access Policies,
ELK Configuration,
Beats in Use,
Machines Being Monitored, and 
How to Use the Ansible Build


### Description of the Topology

The main purpose of this network is to expose a load-balanced and monitored instance of DVWA, the D*mn Vulnerable Web Application.

Load balancing ensures that the application will be highly available, in addition to restricting inbound access to the network.
The load balancer ensures that resources required to process incoming traffic will be shared by both vulnerable web servers. Access controls will ensure that only authorized users will be able to connect.

Integrating an ELK server allows users to easily monitor the vulnerable VMs for changes to the file system.

-Filebeat monitors the log files or locations specified, collects log events, and forwards them for indexing

-Metricbeat records system and application metrics. It allows for viewing cpu,memory,disk, and network metrics, in addition to also allowing monitoring for apache, docker, and other applications installed.

The configuration details of each machine may be found below.
_Note: Use the [Markdown Table Generator](http://www.tablesgenerator.com/markdown_tables) to add/remove values from the table_.

| Name           | Function   | IP Address | Operating System |   |
|----------------|------------|------------|------------------|---|
| JumpBox        | Gateway    | 10.1.0.4   | Linux            |   |
| DVWA 1 (Web-1) | Web Server | 10.1.0.5   | Linux            |   |
| DVWA 2 (Web-2) | Web Server | 10.1.0.6   | Linux            |   |
| ELK-Server     | Monitoring | 10.2.0.4   | Linux            |   |

### Access Policies

The machines on the internal network are not exposed to the public Internet. 

Only the JumpBox machine can accept connections from the Internet. Access to this machine is only allowed from the following IP addresses:
*My public IP*

Machines within the network can only be accessed by each other.
The DVWA machines (Web-1 and Web-2) can send traffic to the ELK-Server

A summary of the access policies in place can be found in the table below:

| Name       | Public Access | Allowed IPs  |
|------------|---------------|--------------|
| JumpBox    | Yes           | My Public IP |
| ELK-Server | No            | 10.1.0.1-254 |
| Web-1      | No            | 10.1.0.1-254 |
| Web-2      | No            | 10.1.0.1.254 |

In addition to the above, Azure has provisioned a load balancer in front of all machines except for the jump box. The load balancer's targets are organized into the following availability zones:

Availability Zone 1: Web-1 + Web-2
Availability Zone 2: ELK

## ELK Server Configuration
The ELK VM exposes an Elastic Stack instance. Docker is used to download and manage an ELK container. 

Ansible was used to automate configuration of the ELK machine. No configuration was performed manually, which is advantageous because this allows for easy replication and portability. 

The Ansible playbook implements the following tasks: Installing Docker, Installing pip3, Installing the Docker Python Module, Setting the system memory necessary for this machine to work properly, and Downloading and launching a Docker ELK container

The Ansible playbook can be found in the YAML file elk.yml in this repo, and is also included here for convenience
 
    "---" (ignore quotes, I am ust including to circumvent the markup language turning this into a line)
    #- name: Config elk VM with Docker
     hosts: elkservers
     remote_user: puzzlegeek
     become: true
     tasks:
  
       #Use apt module
      - name: Install docker.io
        apt:
          update_cache: yes
          name: docker.io
          state: present
        
        #Use apt module
      - name: Install pip3
        apt:
          force_apt_get: yes
          name: python3-pip
          state: present

        #Use pip module
      - name: Install Docker python module
        pip:
          name: docker
          state: present

        #Use sysctl module
      - name: Use more memory
        sysctl:
          name: vm.max_map_count
          value: "262144"
          state: present
          reload: yes

        #Use docker_container module
      - name: download and launch a docker elk container
        docker_container:
          name: elk
          image: sebp/elk:761
          state: started
          restart_policy: always
          published_ports:
            - 5601:5601
            - 9200:9200
            - 5044:5044

The following screenshot displays the result of running `docker ps` after successfully configuring the ELK instance.
![alt text](https://github.com/clayjoseph1994/ELK_Stack_Project/blob/master/Images/DockerPS.PNG)

### Target Machines & Beats
This ELK server is configured to monitor the following machines:
Web-1 and Web-2 (Both running DVWA)

We have installed the following Beats on these machines:
-Filebeat: Filebeat detects changes to the filesystem. Specifically, we use it to collect Apache logs.
-Metricbeat: Metricbeat detects changes in system metrics, such as CPU usage. We use it to detect SSH login attempts, failed sudo escalations, and CPU/RAM statistics.
-Packetbeat: Packetbeat collects packets that pass through the NIC, similar to Wireshark. We use it to generate a trace of all activity that takes place on the network, in case later forensic analysis should be warranted.

The playbook below installs Metricbeat on the target hosts. The playbook for installing Filebeat is not included, but looks essentially identical — simply replace metricbeat with filebeat, and it will work as expected.

### Using the Playbook
In order to use the playbook, you will need to have an Ansible control node already configured. Assuming you have such a control node provisioned, we will need to perform the following steps:
-Copy the playbooks to the Ansible Control Node
-Run each playbook on the appropriate targets 

SSH into the control node (in my case it was the Docker Anisble container on the JumpBox) and follow the steps below:

The easiest way to copy the playbooks is to use Git:

    $ cd /etc/ansible
    $ mkdir files
    #Clone Repository + IaC Files
    $ git clone https://github.com/clayjoseph1994/ELK_Stack_Project.git
    #Move Playbooks and hosts file Into `/etc/ansible`
    $ cp ELK_Stack_Project/playbooks/* .
    $ cp ELK_Stack_Project/files/* ./files

Next, you must create a hosts file to specify which VMs to run each playbook on. Run the commands below or copy the hosts file I have included in the files directory:

    $ cd /etc/ansible
    $ cat > hosts <<EOF
    [webservers]
    10.1.0.5
    10.2.0.6

    [elk]
    10.2.0.4
    EOF

After this, the commands below run the playbook:

    $ cd /etc/ansible
    $ ansible-playbook elk.yml elk
    $ ansible-playbook filebeat-playbook.yml webservers
    $ ansible-playbook metricbeat-playbook.yml webservers

To verify success, wait five minutes to give ELK time to start up.

Then, navigate to the public IP of your ELK Server machine using port 5601. An example would be entering 0.0.0.0:5601 in your web browser. This should take you to the Kibana home screen.

To ensure filebeat is working properly, from your ELK Server home page (Kibana) click on Add Log Data then choose System Logs. Then you simply have to scroll to the page bottom and click Check Data. If all is well, it will display "Data successfully received from this module"

If this is not the case, you likely need to add a filebeat configuration file and add your ELK Server IP address in this config.

From back on your Ansible container on the JumpBox:
          
    #curl https://gist.githubusercontent.com/slape/5cc350109583af6cbe577bbcc0710c93/raw/eca603b72586fbe148c11f9c87bf96a63cb25760/Filebeat > /etc/ansible/files/filebeat-config.yml

The filebeat page lists the details of editing this correctly, but all I had to change was the line under Elasticsearch output. Because the private IP of my ELK Server is 10.2.0.4, I added this to the config file in this section like shown below:

#-------------------------- Elasticsearch output -------------------------------
output.elasticsearch:
  #Boolean flag to enable or disable the output module.
  #enabled: true

  #Array of hosts to connect to.
  #Scheme and port can be left out and will be set to the default (http and 9200)
  #In case you specify and additional path, the scheme is required: http://localhost:9200/path
  #IPv6 addresses should always be defined as: https://[2001:db8::1]:9200
  hosts: ["10.2.0.4:9200"]
  username: "elastic"
  password: "changeme" # TODO: Change this to the password you set

Just make sure you also add the ":9200" after your ELK Server private IP address.

To check if metricbeat is working properly, reload your ELK Server home page (Kibana) and click on Add Metric Data. Click on Docker Metrics, then scroll to the bottom and click on Check data. If the response is "Data successfulyy received from this module" everything is working properly, and you are done!

If this is not the case, check your metricbeat config file. If it is missing, run the following to install it:
curl https://gist.githubusercontent.com/slape/58541585cc1886d2e26cd8be557ce04c/raw/0ce2c7e744c54513616966affb5e9d96f5e12f73/metricbeat

Make sure the parts of the file specified below match this (or whatever IPs you have for your machines, but this is what mine is configured for and it should work with the files I have provided)
:
#============================== Kibana =====================================

#Starting with Beats version 6.0.0, the dashboards are loaded via the Kibana API.
#This requires a Kibana endpoint configuration.
setup.kibana:
  host: "10.2.0.4:5601"

  #Kibana Host
  #Scheme and port can be left out and will be set to the default (http and 5601)
  #In case you specify and additional path, the scheme is required: http://localhost:5601/path
  #IPv6 addresses should always be defined as: https://[2001:db8::1]:5601
  #host: "localhost:5601"

  #Kibana Space ID
  #ID of the Kibana Space into which the dashboards should be loaded. By default,
  #the Default Space will be used.
  #space.id:

#============================= Elastic Cloud ==================================

#These settings simplify using Metricbeat with the Elastic Cloud (https://cloud.elastic.co/).

#The cloud.id setting overwrites the `output.elasticsearch.hosts` and
#`setup.kibana.host` options.
#You can find the `cloud.id` in the Elastic Cloud web UI.
#cloud.id:

#The cloud.auth setting overwrites the `output.elasticsearch.username` and
#`output.elasticsearch.password` settings. The format is `<user>:<pass>`.
#cloud.auth:

#================================ Outputs =====================================

#Configure what output to use when sending the data collected by the beat.

#-------------------------- Elasticsearch output ------------------------------
output.elasticsearch:
  #Array of hosts to connect to.
  hosts: ["10.2.0.4:9200"]
  username: "elastic"
  password: "changeme"


    Congratulations! You have an ELK Stack Server monitoring system logs on two web servers with filebeat and metrics with metricbeat. If you have any issues/questions let me know, and I will try my best to help!
