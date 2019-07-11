import urllib.request

proxyHandler = urllib.request.ProxyHandler({
    'http': 'http://172.17.16.79:3128',
    'https': 'https://172.17.16.79:3128',
})
newOpener = urllib.request.build_opener(
    proxyHandler,
    urllib.request.ProxyBasicAuthHandler()
)
urllib.request.install_opener(newOpener)

resp = urllib.request.urlopen('http://www.grant-miller.com')
print('resp.read()=', resp.read())


