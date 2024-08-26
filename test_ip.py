import os
import stem.process
import re
import requests
import json
from datetime import datetime
from stem import Signal
from stem.control import Controller

class TorProxy:
    def __init__(self, socks_port=9058, control_port=9052, password=None):
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
            print("New Tor connection established.")

    def get_ip(self):
        try:
            response = requests.get("https://ipinfo.io/json", proxies=self.proxies)
            result = json.loads(response.content)
            return f'TOR IP [{datetime.now().strftime("%d-%m-%Y %H:%M:%S")}]: {result["ip"]} {result["country"]}'
        except Exception as e:
            return f"Failed to retrieve IP information: {e}"

if __name__ == "__main__":
    tor_proxy = TorProxy()
    try:
        tor_proxy.start()
        print(tor_proxy.get_ip())
        tor_proxy.renew_connection()
        print(tor_proxy.get_ip())
    finally:
        tor_proxy.stop()
