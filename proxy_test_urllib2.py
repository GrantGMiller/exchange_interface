import urllib.request

# set up authentication info
authinfo = urllib.request.HTTPBasicAuthHandler()
authinfo.add_password(None, 'http://172.17.16.79:3128', 'proxy_user', 'Extron10')
proxy_support = urllib.request.ProxyHandler(
    {"http": "http://172.17.16.79:3128"}
)

# build a new opener that adds authentication and caching FTP handlers
opener = urllib.request.build_opener(
    proxy_support,
    authinfo,
    #urllib.request.CacheFTPHandler()
)

# install it
urllib.request.install_opener(opener)

f = urllib.request.urlopen('http://www.python.org/')
