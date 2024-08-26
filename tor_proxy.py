import os
import time
import stem.process
import re
import requests
import json
from datetime import datetime
from stem import Signal
from stem.control import Controller

class TorProxy:
    def __init__(self, socks_port=9055, control_port=9051, password=None):
        self.socks_port = socks_port
        self.control_port = control_port
        self.password = password
        self.tor_path = r"Tor\tor\tor.exe"
        self.tor_process = None
        self.proxies = {
            'http': f'socks5://127.0.0.1:{self.socks_port}',
            'https': f'socks5://127.0.0.1:{self.socks_port}'
        }

    def start(self):
        self.tor_process = stem.process.launch_tor_with_config(
            config={
                'SocksPort': str(self.socks_port),
                'ControlPort': str(self.control_port),
                'CookieAuthentication': '1',
                'MaxCircuitDirtiness': '5',
                'GeoIPFile': 'https://raw.githubusercontent.com/torproject/tor/main/src/config/geoip',
            },
            init_msg_handler=lambda line: print(line) if re.search('Bootstrapped', line) else False,
            tor_cmd=self.tor_path
        )
        print("TOR proxy started.")

    def stop(self):
        if self.tor_process:
            self.tor_process.kill()
            print("TOR proxy stopped.")

    def renew_connection(self):
        with Controller.from_port(port=self.control_port) as controller:
            if self.password:
                controller.authenticate(password=self.password)
            else:
                controller.authenticate()
            controller.signal(Signal.NEWNYM)
            time.sleep(5)
            ip_address = self.get_ip()
            print(f"New Tor connection established {ip_address}.")

    def get_ip(self):
        try:
            
            response = requests.get("https://ipinfo.io/json", proxies=self.proxies)
            result = json.loads(response.content)
            return f'TOR IP [{datetime.now().strftime("%d-%m-%Y %H:%M:%S")}]: {result["ip"]} {result["country"]}'
        
        #OR check on http://check.torproject.org
        
        except Exception as e:
            return f"Failed to retrieve IP information: {e}"

            