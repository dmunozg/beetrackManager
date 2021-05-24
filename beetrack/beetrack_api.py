import requests, json, os


class BeetrackAPI:
    def __init__(self, api_key, base_url):
        self.base_url = base_url
        self.api_key = api_key
        self.headers = {
            "Content-Type": "application/json",
            "X-AUTH-TOKEN": self.api_key,
        }

    def create_route(self, payload):
        url = self.base_url + "/routes"
        r = requests.post(url, json=payload, headers=self.headers).json()
        return r

    def get_route(self, id):
        url = self.base_url + "/routes/" + str(id)
        r = requests.get(url, headers=self.headers).json()
        return r

    def update_route(self, id, payload):  # Verificar si se puede borrar
        url = self.base_url + "/routes/" + str(id)
        r = requests.put(url, json=payload, headers=self.headers).json()
        return r

    def update_route_dispatch(self, route_id, data):
        url = self.base_url + "/routes/" + str(route_id)
        print({"LastMile UD Function": {"URL": url, "Payload": data}})
        r = requests.put(url, json=data, headers=self.headers).json()
        return r

    def create_truck(self, payload):
        url = self.base_url + "/trucks"
        r = requests.post(url, json=payload, headers=self.headers).json()
        return r

    def get_trucks(self):
        url = self.base_url + "/trucks"
        r = requests.get(url, headers=self.headers).json()
        return r

    def update_dispatch(self, id, payload):
        url = self.base_url + "/dispatches/" + str(id)
        r = requests.put(url, json=payload, headers=self.headers).json()
        return r

    def filter_dispatch(self, tag, route_id):
        url = self.base_url + "/dispatches?cf[{}]={}&rd=5".format(tag, route_id)
        r = requests.get(url, headers=self.headers).json()
        return r

    def get_dispatch(self, guide):
        url = self.base_url + "/dispatches/" + str(guide)
        r = requests.get(url, headers=self.headers).json()
        return r

    def create_dispatch(self, payload):
        url = self.base_url + "/dispatches"
        r = requests.post(url, json=payload, headers=self.headers).json()
        return r
