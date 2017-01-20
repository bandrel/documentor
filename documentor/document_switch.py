#!/usr/bin/env python3
import io
from ciscoconfparse import CiscoConfParse
import argparse
from netmiko import ConnectHandler
import re
from openpyxl import Workbook
import threading
import queue
import json


class WorkerThread(threading.Thread):
    def __init__(self, _queue):
        threading.Thread.__init__(self)
        self.queue = _queue

    def run(self):
        while True:
            current_host = self.queue.get()
            connect_to_switch(username, password, current_host, mode)
            self.queue.task_done()


# noinspection PyUnboundLocalVariable
def connect_to_switch(_username, _password, _host, _mode):
    global host_details
    parsed_output = []
    if _mode == "ssh":
        device = dict(device_type='cisco_ios',
                      ip=_host,
                      username=_username,
                      password=_password,
                      port=port,
                      secret='',
                      verbose=False)
    elif _mode == "telnet":
        device = dict(device_type='cisco_ios_telnet',
                      ip=_host,
                      username=_username,
                      password=_password,
                      port=port,
                      secret='',
                      verbose=False)
    print('Connecting to %s with %s' % (_host, _mode))
    # Create instance of SSHClient object
    # noinspection PyUnboundLocalVariable
    net_connect = ConnectHandler(**device)
    # Automatically add untrusted hosts (make sure okay for security policy in your environment)
    print('%s connection established to %s' % (_mode, _host))
    # Use invoke_shell to establish an 'interactive session'
    hostname = re.sub(r'[#>].*', '', net_connect.find_prompt())
    net_connect.send_command('term len 0')
    output = net_connect.send_command('show run')
    net_connect.send_command('terminal no length')
    if verbose_mode:
        print(output)
    fake_file = io.StringIO(output)
    cisco_parser = CiscoConfParse(fake_file, factory=True)
    interface_dna = cisco_parser.find_objects_dna('IOSIntfLine')
    for sw_int in interface_dna:
        if sw_int.port_type != 'Vlan':
            int_name = sw_int.name
            description = str(sw_int.description)
            admin_shutdown = str(sw_int.is_shutdown)
            if sw_int.is_switchport:
                _mode = 'Switchport'
                switch_access_vlan = str(sw_int.access_vlan)
                switch_trunk_native = str(sw_int.native_vlan)
                switch_trunk_allowed_vlans = sw_int.trunk_vlans_allowed.text
                if sw_int.has_manual_switch_access:
                    switchport_mode = 'Access'
                elif sw_int.has_manual_switch_trunk:
                    switchport_mode = 'Trunk'
                else:
                    switchport_mode = 'Dynamic'
            else:
                _mode = 'Routed'
                switch_access_vlan = 'N/A'
                switchport_mode = 'N/A'
                switch_trunk_native = 'N/A'
                switch_trunk_allowed_vlans = 'N/A'

        parsed_output.append([int_name,
                              description,
                              admin_shutdown,
                              _mode,
                              switchport_mode,
                              switch_access_vlan,
                              switch_trunk_native,
                              switch_trunk_allowed_vlans
                              ])

    lock.acquire()
    host_details[hostname] = parsed_output
    lock.release()
    return


# Declaration of global variables
hosts = []
with open('config.json') as config_file:
    config = json.load(config_file)
username = config['username']
password = config['password']
threads = config['threads']
mode = config['mode'].lower()
port = config['port']
host_details = {}

arg_parser = argparse.ArgumentParser(
        description='script to log into a switch and document the switchport configurations')
arg_parser.add_argument('--hostnames', '-H', type=str, help='Comma separated list of switches to connect to')
arg_parser.add_argument('--inputfile', '-i', type=str, help='File containing line separated switches to connect to')
arg_parser.add_argument('--outputfile', '-o', type=str, help='File containing line separated switches to connect to')
arg_parser.add_argument('--verbose', '-v', type=bool, help='enable verbose output')
args = arg_parser.parse_args()

if args.verbose:
    verbose_mode = True
else:
    verbose_mode = False

if args.hostnames:
    for host in args.hostnames.split(','):
        hosts.append(host)
if args.inputfile:
    with open(args.inputfile, 'r') as f:
        for line in f:
            hosts.append(line.rstrip())
if args.outputfile:
    outputfile = args.outputfile
else:
    outputfile = 'output.xlsx'
lock = threading.Lock()

queue = queue.Queue()
for host in hosts:
    queue.put(host)

for i in range(threads):
    worker = WorkerThread(queue)
    worker.setDaemon(True)
    worker.start()

queue.join()
header = ['Interface Name',
          'Description',
          'Shutdown',
          'Switchport/Routed',
          'Switchport Mode',
          'Access VLAN ID',
          'Trunk Native VLANs',
          'Trunk Allowed VLANs'
          ]
# setup excel workbook for output
wb = Workbook()
for switch in host_details.keys():
    active_sheet = wb.create_sheet(switch)
    active_sheet.append(header)
    for interface in host_details[switch]:
        active_sheet.append(interface)

wb.remove_sheet(wb['Sheet'])
wb.save(outputfile)
