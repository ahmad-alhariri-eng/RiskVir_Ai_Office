import urllib.request
import re

try:
    html = urllib.request.urlopen('https://office-addin-two.vercel.app/').read().decode('utf-8')
    m = re.search(r'src=["\']([^"\']*index[^"\']*\.js)["\']', html)
    if m:
        js_url = m.group(1)
        if not js_url.startswith('http'):
            js_url = 'https://office-addin-two.vercel.app' + js_url
        print('Script:', js_url)
        js = urllib.request.urlopen(js_url).read().decode('utf-8')
        print('Alwaysdata URL present:', 'alwaysdata' in js)
    else:
        print('No script found')
except Exception as e:
    print('Error:', e)
