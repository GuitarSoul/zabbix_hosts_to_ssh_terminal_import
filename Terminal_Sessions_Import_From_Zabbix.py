
def zabbix_get_devices():
    import sys, yaml
    import logging
    import re
    import pandas as pd
    from pyzabbix import ZabbixAPI

    # All specific info is written in yaml file
    with open("config.yml", 'r', encoding="utf-8") as ymlfile:
        config = yaml.load(ymlfile, Loader=yaml.FullLoader)
        
    # For ssh sessions to remember your username as default
    username_tacacs = input('Tacacs Username For Default SSH Session : ')
    if config['debug'] == 'True':
        stream = logging.StreamHandler(sys.stdout)
        stream.setLevel(logging.DEBUG)
        log = logging.getLogger('pyzabbix')
        log.addHandler(stream)
        log.setLevel(logging.DEBUG)

    zapi = ZabbixAPI(config['zabbix.url'])
    username = config['zabbix.username']
    password = config['zabbix.password']
    #zapi.session.verify = False
    zapi.login(username, password)

    hosts = zapi.host.get(monitored_hosts='1',output=['hostid','groups','name','inventory'],selectInventory=['serialno_a','serialno_b','os','model'],selectGroups=['name'])

    #host_list = []
    list_dict =[]
    host_dict = {}
    for host in hosts:
        search_ip = zapi.hostinterface.get(hostids=host['hostid'], output=['ip'])
        if not host['inventory']:
            host['inventory'] = {'serialno_a': '', 'serialno_b': '', 'os': '','model':''}

        host_dict = {'Site_Address': host['groups'][0]['name'],
                     'Hostname': host['name'],
                     'IP': search_ip[0]['ip'],
                     'Serial_Number_A': host['inventory']['serialno_a'],
                     'Serial_Number_B': host['inventory']['serialno_b'],
                     'OS_Version': host['inventory']['os'],
                     'Model': host['inventory']['model']
                     }

        list_dict.append(host_dict)

    dataframe = pd.DataFrame(list_dict)
    writer = pd.ExcelWriter('Hosts.xlsx')
    dataframe.to_excel(writer, index=False)
    writer.save()
    return list_dict, username_tacacs

def create_sessions():
    # This function uses a text file with data about all needed devices from Zabbix
    # and creates files with sessions for Putty, SuperPutty, SecureCRT, XShell

    from jinja2 import Environment, FileSystemLoader
    import re

    devices_list, username_tacacs = zabbix_get_devices()

    # Create a txt file with hosts for Putty. Then, this use this file to run the createPuttySessions.ps1
    #curr_dir = os.path.dirname(os.path.abspath(__file__))
    env = Environment(loader=FileSystemLoader('Templates'))
    template = env.get_template('Putty_Sessions_Import_Template.txt')
    with open('Putty_Sessions\\puttyhosts.txt', 'w') as f:
        f.write(template.render(devices_list=devices_list))

    # Create a xml file with hosts for SuperPutty
    #curr_dir = os.path.dirname(os.path.abspath(__file__))
    env = Environment(loader=FileSystemLoader('Templates'))
    template = env.get_template('Super_Putty_Sessions_Import_Template.txt')
    with open('Super_Putty_Sessions\\Sessions_SuperPutty_Import.xml', 'w') as f:
        f.write(template.render(devices_list=devices_list, username=username_tacacs))

    # Create a csv file with hosts for SecureCRT. Then, this use this file to run the Import_sessions.py
    #curr_dir = os.path.dirname(os.path.abspath(__file__))
    env = Environment(loader=FileSystemLoader('Templates'))
    template = env.get_template('SecureCRT_Sessions_Import_Template.txt')
    with open('SecureCRT_Sessions\\Sessions_SecureCRT_Import.csv', 'w') as f:
        f.write(template.render(devices_list=devices_list, username=username_tacacs))

    # Create a csv file with hosts for XShell
    #curr_dir = os.path.dirname(os.path.abspath(__file__))
    env = Environment(loader=FileSystemLoader('Templates'))
    template = env.get_template('XShell_Sessions_Import_Template.txt')
    with open('XShell_Sessions\\Sessions_XShell_Import.csv', 'w') as f:
        f.write(template.render(devices_list=devices_list, username=username_tacacs))

create_sessions()