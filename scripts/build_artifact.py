"""Claude 아티팩트(호스팅 페이지)용 단일 HTML 을 만든다.

    python3 scripts/build_artifact.py 출력경로.html

- 아티팩트는 <!doctype>/<html>/<head>/<body> 를 스스로 감싸므로 본문만 남긴다.
- CSS 와 앱 코드는 페이지 안에 넣고, 라이브러리(vendor/)는 페이지와 함께 올리는 파일로 둔다.
- 처음 여는 사람에게는 샘플 데이터를 채워서 보여 준다(PO_AUTO_DEMO).
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


if __name__ == '__main__':
    main(sys.argv[1])
