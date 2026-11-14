import re, io, base64, pathlib, os, time
from PIL import Image

root = pathlib.Path(__file__).resolve().parent.parent
src, out = root / 'src', root / 'dist'
out.mkdir(exist_ok=True)

def env(k):
    f = root / '.env'
    if f.exists():
        for line in f.read_text(encoding='utf-8').splitlines():
            if line.startswith(k + '='):
                return line.split('=', 1)[1].strip()
    return os.environ.get(k, '')

for n in ('Index.html', 'GalianInput.html', 'Dashboard.html'):
    if not (src / n).exists():
        raise SystemExit(f'File src/{n} belum ada. Salin dari Apps Script lama.')

rd = lambda n: (src / n).read_text(encoding='utf-8')
url = env('URL_EXEC')
if not url:
    raise SystemExit('Isi URL_EXEC di file .env (salin dari .env.example).')

dash = rd('Dashboard.html')
shim = rd('shim.js').replace('__URL_EXEC__', url)
head = ('<meta name="viewport" content="width=device-width,initial-scale=1">'
        '<meta name="theme-color" content="#07183a"><link rel="manifest" href="manifest.json">'
        '<link rel="apple-touch-icon" href="icon-192.png"><meta name="apple-mobile-web-app-capable" content="yes">'
        '<title>Monitoring Lapangan</title><script>' + shim + '</script>')
sw_reg = "<script>if('serviceWorker' in navigator)navigator.serviceWorker.register('sw.js')</script>"
html = rd('Index.html').replace('<head>', '<head>' + head, 1)
html = html.replace('</body>', rd('GalianInput.html') + dash + sw_reg + '</body>', 1)
(out / 'index.html').write_text(html, encoding='utf-8')

m = re.search(r'data:image/png;base64,([A-Za-z0-9+/=]+)', dash)
if not m:
    raise SystemExit('Logo base64 tidak ditemukan di Dashboard.html')
logo = Image.open(io.BytesIO(base64.b64decode(m.group(1)))).convert('RGBA')
for s in (192, 512):
    bg = Image.new('RGBA', (s, s), '#07183a')
    l = logo.copy(); l.thumbnail((int(s * .7),) * 2)
    bg.paste(l, ((s - l.width) // 2, (s - l.height) // 2), l)
    bg.convert('RGB').save(out / f'icon-{s}.png')

(out / 'manifest.json').write_text(rd('manifest.json'), encoding='utf-8')
(out / 'sw.js').write_text(rd('sw.js').replace('__V__', str(int(time.time()))), encoding='utf-8')
print('Build selesai ->', out)
import re, io, base64, pathlib, os, time
from PIL import Image

root = pathlib.Path(__file__).resolve().parent.parent
src, out = root / 'src', root / 'dist'
out.mkdir(exist_ok=True)

def env(k):
    f = root / '.env'
    if f.exists():
        for line in f.read_text(encoding='utf-8').splitlines():
            if line.startswith(k + '='):
                return line.split('=', 1)[1].strip()
    return os.environ.get(k, '')

for n in ('Index.html', 'GalianInput.html', 'Dashboard.html', 'Mobile.html'):
    if not (src / n).exists():
        raise SystemExit(f'File src/{n} belum ada. Salin dari Apps Script lama.')

rd = lambda n: (src / n).read_text(encoding='utf-8')
url = env('URL_EXEC')
if not url:
    raise SystemExit('Isi URL_EXEC di file .env (salin dari .env.example).')

dash = rd('Dashboard.html')
shim = rd('shim.js').replace('__URL_EXEC__', url)
head = ('<meta name="viewport" content="width=device-width,initial-scale=1">'
        '<meta name="theme-color" content="#07183a"><link rel="manifest" href="manifest.json">'
        '<link rel="apple-touch-icon" href="icon-192.png"><meta name="apple-mobile-web-app-capable" content="yes">'
        '<title>Monitoring Lapangan</title><script>' + shim + '</script>')
sw_reg = "<script>if('serviceWorker' in navigator)navigator.serviceWorker.register('sw.js')</script>"
html = rd('Index.html').replace('<head>', '<head>' + head, 1)
html = html.replace('</body>', rd('GalianInput.html') + dash + rd('Mobile.html') + sw_reg + '</body>', 1)
(out / 'index.html').write_text(html, encoding='utf-8')

m = re.search(r'data:image/png;base64,([A-Za-z0-9+/=]+)', dash)
if not m:
    raise SystemExit('Logo base64 tidak ditemukan di Dashboard.html')
logo = Image.open(io.BytesIO(base64.b64decode(m.group(1)))).convert('RGBA')
for s in (192, 512):
    bg = Image.new('RGBA', (s, s), '#07183a')
    l = logo.copy(); l.thumbnail((int(s * .7),) * 2)
    bg.paste(l, ((s - l.width) // 2, (s - l.height) // 2), l)
    bg.convert('RGB').save(out / f'icon-{s}.png')

(out / 'manifest.json').write_text(rd('manifest.json'), encoding='utf-8')
(out / 'sw.js').write_text(rd('sw.js').replace('__V__', str(int(time.time()))), encoding='utf-8')
print('Build selesai ->', out)