import requests
proxies = {
    'http': '172.17.16.79:3128',
    'https': '172.17.16.79:3128',
}
resp = requests.get('http://grant-miller.com', proxies=proxies)
print('resp.text=', resp.text)