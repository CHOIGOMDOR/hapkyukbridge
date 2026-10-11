"""index.html 을 앱 스크립트용 apps-script/Index.html 로 만들어요.
- 이미지는 한 번씩만 파일 안에 넣고, 폰트는 드라이브에서 불러와서 제외해요."""
import base64, os, re
os.chdir(os.path.join(os.path.dirname(__file__), '..'))
s = open('index.html', encoding='utf-8').read()
s = s.replace('@font-face{font-family:"Memoment";src:url("assets/fonts/MemomentKkukkukk.woff2") format("woff2");font-display:swap}', '')
# 정적 <img> 의 src 는 제거 (자바스크립트가 채워요)
s = re.sub(r'(<img [^>]*?) src="assets/img/[^"]+"', r'\1', s)
data = {}
for n in ['cheer-friends', 'cheer-pompom', 'lying-trio']:
    p = 'assets/img/%s.png' % n
    data[p] = 'data:image/png;base64,' + base64.b64encode(open(p, 'rb').read()).decode()
lit = '{' + ','.join('"%s":"%s"' % (k, v) for k, v in data.items()) + '}'
s = s.replace('var IMG_DATA = {};', 'var IMG_DATA = ' + lit + ';', 1)
open('apps-script/Index.html', 'w', encoding='utf-8').write(s)
print(len(s) // 1024, 'KB')
