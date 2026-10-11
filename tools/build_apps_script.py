"""index.html 을 앱 스크립트용 apps-script/Index.html 로 만들어요. 이미지와 폰트를 파일 안에 넣어요."""
import base64, os, re
os.chdir(os.path.join(os.path.dirname(__file__), '..'))
s = open('index.html', encoding='utf-8').read()
# 폰트: 실제로 쓰는 글자만 담은 작은 버전(약 560KB)을 파일 안에 넣어요 → 드라이브 설정 없이 바로 적용돼요
fdata = base64.b64encode(open('assets/fonts/MemomentKkukkukk.subset.woff2', 'rb').read()).decode()
s = s.replace('src:url("assets/fonts/MemomentKkukkukk.woff2") format("woff2")', 'src:url(data:font/woff2;base64,' + fdata + ') format("woff2")')
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
