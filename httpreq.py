import requests

r = requests.get('http://192.168.0.111')
print(r.status_code)