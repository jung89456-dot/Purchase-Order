"""Claude 아티팩트(호스팅 페이지)용 단일 HTML 을 만든다.

    python3 scripts/build_artifact.py 출력경로.html
    python3 scripts/build_artifact.py --standalone 발주서_자동변환기.html

--standalone: 라이브러리까지 전부 한 파일에 넣은 완전한 HTML. 받아서 더블클릭하면 바로 열린다.

- 아티팩트는 <!doctype>/<html>/<head>/<body> 를 스스로 감싸므로 본문만 남긴다.
- CSS 와 앱 코드는 페이지 안에 넣고, 라이브러리(vendor/)는 페이지와 함께 올리는 파일로 둔다.
- 처음 여는 사람에게는 샘플 데이터를 채워서 보여 준다(PO_AUTO_DEMO).
- vendor/ 사본은 출력 폴더에 함께 만든다. 아티팩트는 제어 문자·U+FFFD 가 그대로 든 파일을 받지 않으므로,
  SheetJS 코드표 문자열의 ASCII 밖 문자와 제어 문자를 같은 뜻의 \\uXXXX 표기로 바꾼다.
"""
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
VENDOR = ('vendor/xlsx.full.min.js', 'vendor/exceljs.min.js')


def main(out):
    html = (ROOT / 'index.html').read_text(encoding='utf-8')
    title = re.search(r'<title>.*?</title>', html, re.S).group(0)
    body = re.search(r'<body>(.*)</body>', html, re.S).group(1)
    css = (ROOT / 'css' / 'style.css').read_text(encoding='utf-8')

    def script(m):
        src = m.group(1)
        if src in VENDOR:
            return m.group(0)
        code = (ROOT / src).read_text(encoding='utf-8')
        return f'<script>\n{code}\n</script>'

    body = re.sub(r'<script src="([^"]+)"></script>', script, body)
    body = body.replace('<script>\n(function () {\n  \'use strict\';\n\n  var APP_VERSION',
                        '<script>window.PO_AUTO_DEMO = true;</script>\n<script>\n(function () {\n  \'use strict\';\n\n  var APP_VERSION', 1)
    assert 'PO_AUTO_DEMO = true' in body
    page = f'{title}\n<style>\n{css}\n</style>\n{body.strip()}\n'
    Path(out).parent.mkdir(parents=True, exist_ok=True)
    Path(out).write_text(page, encoding="utf-8")
    print(out, len(page.encode('utf-8')), 'bytes')
    for rel in VENDOR:
        dst = Path(out).parent / rel
        dst.parent.mkdir(parents=True, exist_ok=True)
        dst.write_text(ascii_js((ROOT / rel).read_text(encoding='utf-8')), encoding='ascii')
        print(dst)


def ascii_js(src):
    """문자열 안의 ASCII 밖 문자·제어 문자를 \\uXXXX 로 (역슬래시로 이스케이프된 경우는 없어야 함)"""
    out = []
    for i, ch in enumerate(src):
        o = ord(ch)
        if 0x20 <= o < 0x80 or ch in '\t\n\r':
            out.append(ch)
            continue
        k, j = 0, i - 1
        while j >= 0 and src[j] == '\\':
            k, j = k + 1, j - 1
        assert k % 2 == 0, f'escaped special char at {i}'
        if o > 0xFFFF:
            o -= 0x10000
            out.append('\\u%04x\\u%04x' % (0xD800 + (o >> 10), 0xDC00 + (o & 0x3FF)))
        else:
            out.append('\\u%04x' % o)
    return ''.join(out)


def standalone(out):
    """라이브러리까지 모두 넣은 한 파일짜리 완전한 HTML (내 컴퓨터에서 더블클릭용)"""
    html = (ROOT / 'index.html').read_text(encoding='utf-8')
    css = (ROOT / 'css' / 'style.css').read_text(encoding='utf-8')
    html = html.replace('<link rel="stylesheet" href="css/style.css">', f'<style>\n{css}\n</style>')

    def script(m):
        code = (ROOT / m.group(1)).read_text(encoding='utf-8')
        assert '</script' not in code.lower(), m.group(1)
        return f'<script>\n{code}\n</script>'

    html = re.sub(r'<script src="([^"]+)"></script>', script, html)
    html = html.replace('<script>\n(function () {\n  \'use strict\';\n\n  var APP_VERSION',
                        '<script>window.PO_AUTO_DEMO = true;</script>\n<script>\n(function () {\n  \'use strict\';\n\n  var APP_VERSION', 1)
    assert 'PO_AUTO_DEMO = true' in html and 'src=' not in re.sub(r'<script>.*?</script>', '', html, flags=re.S)
    Path(out).parent.mkdir(parents=True, exist_ok=True)
    Path(out).write_text(html, encoding='utf-8')
    print(out, len(html.encode('utf-8')), 'bytes')


if __name__ == '__main__':
    if sys.argv[1] == '--standalone':
        standalone(sys.argv[2])
    else:
        main(sys.argv[1])
