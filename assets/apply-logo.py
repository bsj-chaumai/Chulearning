#!/usr/bin/env python3
# ============================================================================
# apply-logo.py — ロゴ反映スクリプト（Python 標準ライブラリのみ・追加インストール不要）
# ----------------------------------------------------------------------------
# ロゴの唯一の正本は  assets/bravesoft-logo.png  です。
# 公式ロゴをここに上書き保存してから、このスクリプトを 1 回実行すると、
# 「中にロゴを焼き込んでいる」成果物テンプレに反映されます。
#
#   実行:  python3 assets/apply-logo.py
#
# 反映先:
#   - 02_Report_Template/report_template.html … base64 埋め込みを差し替え（高さ固定・縦横比は自動）
#   - 02_Report_Template/report_template.pptx … スライド内のロゴ画像を差し替え（最も多用されている画像＝ロゴと判定。図は温存）
#
# 自動参照で “差し替えるだけ” で反映される（このスクリプト不要）:
#   - index.html              … assets/bravesoft-logo.png を相対参照
#   - 各 Markdown 報告書        … ../assets/ もしくは ../../assets/ を相対参照
#
# 注意（PPTX）: 新ロゴの縦横比が従来と大きく異なる場合、PPTX 表紙/最終ページで
#   伸びて見えることがあります。その場合は PPTX を再ビルド（AI に「PPTXのロゴ反映して」）
#   してください。HTML/Markdown/index は縦横比を自動で保ちます。
# ============================================================================
import base64, hashlib, os, re, sys, zipfile, shutil, collections

KIT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LOGO = os.path.join(KIT, "assets", "bravesoft-logo.png")
HTML = os.path.join(KIT, "02_Report_Template", "report_template.html")
PPTX = os.path.join(KIT, "02_Report_Template", "report_template.pptx")

def fail(msg): print("✗", msg); sys.exit(1)

if not os.path.exists(LOGO):
    fail(f"ロゴが見つかりません: {LOGO}\n  公式ロゴをこのパスに保存してから再実行してください。")
logo_bytes = open(LOGO, "rb").read()
if logo_bytes[:8] != b"\x89PNG\r\n\x1a\n":
    fail("assets/bravesoft-logo.png は PNG ではありません（PNG で保存してください）。")
print(f"● ロゴ読込: assets/bravesoft-logo.png ({len(logo_bytes):,} bytes)")

# --- 1) HTML: base64 を差し替え（ロゴだけが data:image/png;base64 を使っている） ---
if os.path.exists(HTML):
    html = open(HTML, encoding="utf-8").read()
    uri = "data:image/png;base64," + base64.b64encode(logo_bytes).decode()
    new, n = re.subn(r'data:image/png;base64,[A-Za-z0-9+/=]+', uri, html)
    if n:
        open(HTML, "w", encoding="utf-8").write(new)
        print(f"✓ report_template.html: ロゴ {n} 箇所を差し替え")
    else:
        print("… report_template.html: 埋め込みロゴが見つかりませんでした（スキップ）")
else:
    print("… report_template.html が無いのでスキップ")

# --- 2) PPTX: 最も多用されている画像（=ロゴ）を新ロゴに差し替え。図(ユニーク画像)は温存 ---
if os.path.exists(PPTX):
    with zipfile.ZipFile(PPTX) as z:
        names = z.namelist()
        media = [n for n in names if n.startswith("ppt/media/")]
        digests = {m: hashlib.md5(z.read(m)).hexdigest() for m in media}
    freq = collections.Counter(digests.values())
    logo_hash = freq.most_common(1)[0][0] if freq else None          # 最頻＝ロゴ
    targets = [m for m, d in digests.items() if d == logo_hash]
    if not targets:
        print("… PPTX: ロゴ画像を特定できませんでした（スキップ）")
    else:
        tmp = PPTX + ".tmp"
        with zipfile.ZipFile(PPTX) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = logo_bytes if item.filename in targets else zin.read(item.filename)
                zout.writestr(item, data)
        shutil.move(tmp, PPTX)
        print(f"✓ report_template.pptx: ロゴ画像 {len(targets)} 箇所を差し替え（図は温存）")
else:
    print("… report_template.pptx が無いのでスキップ")

print("\n完了。index.html / Markdown は assets/bravesoft-logo.png を自動参照します（差し替えれば即反映）。")
print("※ PPTX で縦横比が崩れる場合は AI に「PPTXのロゴ反映して」と依頼してください（正しい比率で再ビルドします）。")
