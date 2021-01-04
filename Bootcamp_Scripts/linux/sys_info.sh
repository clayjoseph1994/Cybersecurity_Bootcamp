#!/bin/bash
#This script will create research directory in user's home directory
#It will then create a file named sys_info.txt in this new directory
#The commands listed will output the following:
#A title and today's date.
#The uname info for the machine.
#The machine's IP address narrowed to a single line for my particular machine
#The Hostname.

#The DNS info.
#The Memory info.
#The CPU info.
#The Disk usage.
#The currently logged on users.
#Find files with permission 777
#Fine top 10 processes 

mkdir ~/research 2> /dev/null

#send all output to research directory in a file named sys_info.txt

echo "A Quick System Audit Script" > ~/research/sys_info.txt
date >> ~/research/sys_info.txt
echo "" >> ~/research/sys_info.txt
echo "Machine Type Info:" >> ~/research/sys_info.txt
echo $MACHTYPE >> ~/research/sys_info.txt
echo -e "Uname info: $(uname -a) \n" >> ~/research/sys_info.txt
echo -e "IP Info: $(ip addr | head -9 | tail -1) \n" >> ~/research/sys_info.txt
echo -e "Hostname: $(hostname -s) \n" >> ~/research/sys_info.txt
echo "DNS Servers: " >> ~/research/sys_info.txt
cat /etc/resolv.conf >> ~/research/sys_info.txt

echo -e "\nMemory Info:" >> ~/research/sys_info.txt
free >> ~/research/sys_info.txt
echo -e "\nCPU Info:" >> ~/research/sys_info.txt
lscpu | grep CPU >> ~/research/sys_info.txt
echo -e "\nDisk Usage:" >> ~/research/sys_info.txt
df -H | head -2 >> ~/research/sys_info.txt
echo -e "\nWho is logged in: \n $(who -a) \n" >> ~/research/sys_info.txt
echo -e "\nSUID Files:" >> ~/research/sys_info.txt
find / -type f -perm 777 >> ~/research/sys_info.txt
echo -e "\nTop 10 Processes" >> ~/research/sys_info.txt
ps aux -m | awk {'print $1, $2, $3, $4, $11'} | head >> ~/research/sys_info.txt