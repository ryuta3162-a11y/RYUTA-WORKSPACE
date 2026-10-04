import csv, json, sys, urllib.request, collections
path = sys.argv[1]
rows = list(csv.DictReader(open(path, encoding="cp932", errors="replace")))
p = collections.OrderedDict()
for r in rows:
    d = p.setdefault(r['uid'], {'name': r['name']})
    d[r['q_id']] = r['answer']
out = [[u, d['name'], d.get('7', ''), d.get('8', ''), d.get('11', ''), d.get('12', '')] for u, d in p.items() if d.get('12') or d.get('7')]
url = 'https://script.google.com/macros/s/AKfycbzNYW-n-jmwyU-haP8Dl5oGI_kx8JXk3CE7A52FhSOJGughsp90qeDLuN-6sCr8Cu5T/exec'
body = json.dumps({'api': 'importLeaveSurvey', 'token': 'kyodo-ws-refresh-7f3c91', 'rows': out}).encode('utf-8')
req = urllib.request.Request(url, data=body, headers={'Content-Type': 'application/json'})
print(len(out), urllib.request.urlopen(req, timeout=300).read().decode('utf-8')[:500])
