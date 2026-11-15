# tools/test_export.js で作った2つのディレクトリの xlsx を、作成日時（docProps）以外の中身で比較する
# 使い方: python3 tools/compare_xlsx.py <変更前> <変更後>   … 差が無ければ「differences 0」
import sys, zipfile, os
a, b = sys.argv[1], sys.argv[2]; bad = 0; n = 0
for f in sorted(os.listdir(a)):
    if not f.endswith('.xlsx'): continue
    n += 1
    za, zb = zipfile.ZipFile(os.path.join(a, f)), zipfile.ZipFile(os.path.join(b, f))
    na = sorted(x for x in za.namelist() if not x.startswith('docProps/'))
    nb = sorted(x for x in zb.namelist() if not x.startswith('docProps/'))
    if na != nb: print(f, 'parts differ'); bad += 1; continue
    for x in na:
        if za.read(x) != zb.read(x): print(f, x, 'differs'); bad += 1
print('compared', n, 'files; differences', bad)
print('api same:', open(a + '/_api.txt').read() == open(b + '/_api.txt').read())
sys.exit(1 if bad else 0)
