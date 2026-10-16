# -*- coding: utf-8 -*-
"""치즈볼 공부 노트의 디자인을 다른 웹페이지에 그대로 쓰기 위한 키트(design-kit/)를 만들어요."""
import re, shutil, os, sys
root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
os.chdir(root)
src = open('index.html', encoding='utf-8').read()
kit = 'design-kit'
os.makedirs(kit + '/assets/img', exist_ok=True)
os.makedirs(kit + '/assets/fonts', exist_ok=True)

css = re.search(r'<style>\n(.*?)</style>', src, re.S).group(1)
open(kit + '/theme.css', 'w', encoding='utf-8').write(
    '/* 치즈볼 공부 노트 디자인 전체 CSS (색·배경·패턴·애니메이션·카드·탭·칩·토스트…). 클래스 설명은 DESIGN.md 참고 */\n' + css)

def grab(start, end):
    a = src.index(start); b = src.index(end, a)
    return src[a:b]

js = []
js.append('/* 디자인에 관련된 JS만 모았어요. 앱 데이터(시트, 퀴즈 등) 코드는 reference.html 에 있어요 */')
js.append(grab('var THEMES = [', 'function applyTheme(){') + grab('function applyTheme(){', '\n}\n') + '\n}\n')
js.append(grab('function kindOf(u){', '/* ============ 데이터 ============ */'))
js.append(grab('function showTab(b){', '\n[].forEach.call(document.querySelectorAll("nav button"),function(b){') )
js.append(grab('function tick(){', '\n}\n') + '\n}\n')
js.append(grab('var toastT;function toast', '\nfunction lget'))
open(kit + '/theme.js', 'w', encoding='utf-8').write('\n'.join(js))

shutil.copy('index.html', kit + '/reference.html')
for f in os.listdir('assets/img'):
    shutil.copy('assets/img/' + f, kit + '/assets/img/' + f)
shutil.copy('assets/fonts/MemomentKkukkukk.subset.woff2', kit + '/assets/fonts/')
shutil.copy('assets/fonts/MemomentKkukkukk.woff2', kit + '/assets/fonts/')
print('ok')
