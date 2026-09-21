#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
매치메이커 프로 → 탁구 리그전 평균 득점 크롤러
  2026-2학기 「2-1 ~ 2-4 탁구수행평가」 대회에서 학생별 경기 득점을 긁어
  야구수행평가 앱의 「리그 득점 일괄 입력」에 그대로 붙여넣을 형식으로 뽑는다.

  사용법:  python3 crawl.py            # 결과를 화면과 out/ 에 저장
           python3 crawl.py --quiet    # 파일만
"""
import json, os, sys, urllib.request, datetime

DB   = "https://matchmaker-pro-90a1b-default-rtdb.asia-southeast1.firebasedatabase.app"
OUT  = os.path.join(os.path.dirname(os.path.abspath(__file__)), "out")
# 매치메이커 대회 이름 → 야구앱 학급 이름
CLASS_OF = {"2-1 탁구수행평가": "1반", "2-2 탁구수행평가": "2반",
            "2-3 탁구수행평가": "3반", "2-4 탁구수행평가": "4반"}


def fetch(path):
    with urllib.request.urlopen(DB + path, timeout=60) as r:
        return json.load(r)


def is_bye(side):
    return "bye" in str((side or {}).get("id", ""))


def collect(t):
    """대회 하나 → {이름: [득점, 득점, …]}  (11점 단세트, 세트 점수가 곧 득점)"""
    out, total, done = {}, 0, 0
    for g in (t.get("generatedGroups") or []):
        for rd in (g.get("rounds") or []):
            for m in (rd.get("matches") or []):
                h, a = m.get("home"), m.get("away")
                if is_bye(h) or is_bye(a):
                    continue
                total += 1
                sc = m.get("setScores") or []
                if m.get("status") != "completed" or not sc:
                    continue
                done += 1
                for side, key in ((h, "h"), (a, "a")):
                    nm = (side or {}).get("name")
                    if not nm:
                        continue
                    pts = sum(int((s or {}).get(key) or 0) for s in sc)
                    out.setdefault(nm, []).append(pts)
    return out, done, total


def score(avg):
    """평가계획서 척도 — 9.0↑20 / 7.0~16 / 5.0~12 / 3.0~8 / 미만 4"""
    return 20 if avg >= 9 else 16 if avg >= 7 else 12 if avg >= 5 else 8 if avg >= 3 else 4


def main():
    quiet = "--quiet" in sys.argv
    os.makedirs(OUT, exist_ok=True)
    ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    tours = fetch("/tournaments.json") or {}

    report, pastes = [], {}
    for tid, t in tours.items():
        if not t:
            continue
        name = ((t.get("config") or {}).get("name") or "").strip()
        cn = CLASS_OF.get(name)
        if not cn:
            continue
        data, done, total = collect(t)
        lines = []
        for nm in sorted(data):
            v = data[nm]
            lines.append("%s:%s" % (nm, ",".join(str(x) for x in v)))
        pastes[cn] = "\n".join(lines)
        avgs = {nm: sum(v) / len(v) for nm, v in data.items()}
        report.append({
            "class": cn, "name": name, "id": tid,
            "done": done, "total": total,
            "students": len(data),
            "detail": sorted(
                ({"이름": nm, "경기": len(data[nm]), "총득점": sum(data[nm]),
                  "평균": round(avgs[nm], 2), "점수": score(avgs[nm])}
                 for nm in data), key=lambda r: -r["평균"]),
        })

    report.sort(key=lambda r: r["class"])
    json.dump({"crawledAt": ts, "report": report}, open(os.path.join(OUT, "latest.json"), "w"),
              ensure_ascii=False, indent=1)
    for cn, txt in pastes.items():
        open(os.path.join(OUT, "붙여넣기_%s.txt" % cn), "w").write(txt)

    lines = ["📊 탁구 리그전 크롤링 — %s" % ts, ""]
    for r in report:
        pct = (100 * r["done"] // r["total"]) if r["total"] else 0
        lines.append("【%s】 %s — 경기 %d/%d (%d%%) · 기록된 학생 %d명"
                     % (r["class"], r["name"], r["done"], r["total"], pct, r["students"]))
        if r["done"] == 0:
            lines.append("   아직 입력된 경기가 없습니다")
        else:
            for x in r["detail"][:5]:
                lines.append("   %-4s %d경기 평균 %.2f → %d점"
                             % (x["이름"], x["경기"], x["평균"], x["점수"]))
            if len(r["detail"]) > 5:
                lines.append("   …외 %d명" % (len(r["detail"]) - 5))
        lines.append("")
    txt = "\n".join(lines)
    open(os.path.join(OUT, "요약.txt"), "w").write(txt)
    if not quiet:
        print(txt)
    return 0


if __name__ == "__main__":
    sys.exit(main())
