#!/usr/bin/env python3
"""pptx 안의 끊어진 r:id 참조(관계 파일에 없는 hlinkClick 등)를 제거한다.
사용: fix_dangling.py in.pptx out.pptx   (끊어진 참조 목록 출력)
"""
import sys, re, zipfile, posixpath

def run(src, dst):
    zi = zipfile.ZipFile(src); names = set(zi.namelist()); report = []
    zo = zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED)
    for it in zi.infolist():
        n = it.filename; d = zi.read(n)
        if n.endswith(".xml") and n.startswith("ppt/") and "/_rels/" not in n:
            relp = posixpath.join(posixpath.dirname(n), "_rels", posixpath.basename(n) + ".rels")
            ids = set(re.findall(r'Id="([^"]+)"', zi.read(relp).decode())) if relp in names else set()
            s = d.decode("utf-8")
            def kill(m):
                rid = re.search(r'r:id="([^"]+)"', m.group(0))
                if rid and rid.group(1) not in ids:
                    report.append((n, rid.group(1))); return ""
                return m.group(0)
            s2 = re.sub(r'<a:hlinkClick\b[^>]*?/>', kill, s)
            s2 = re.sub(r'<a:hlinkClick\b[^>]*?[^/]>.*?</a:hlinkClick>', kill, s2, flags=re.S)
            for rid in re.findall(r'r:(?:id|embed|link)="([^"]+)"', s2):
                if rid not in ids: report.append((n, rid, "남은 끊어진 참조"))
            d = s2.encode("utf-8")
        zo.writestr(it, d)
    zo.close()
    return report

def sanitize_inplace(path):
    """저장된 pptx에서 끊어진 참조를 제거(제자리). PowerPoint 복구 메시지 방지용."""
    import os, tempfile, shutil
    fd, tmp = tempfile.mkstemp(suffix=".pptx"); os.close(fd)
    try:
        rep = run(path, tmp)
        shutil.move(tmp, path)
    finally:
        if os.path.exists(tmp): os.remove(tmp)
    return rep


if __name__ == "__main__":
    r = run(sys.argv[1], sys.argv[2])
    print("제거/남은 참조:", len(r)); [print(" ", x) for x in r]
