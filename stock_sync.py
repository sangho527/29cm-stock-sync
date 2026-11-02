# -*- coding: utf-8 -*-
"""
셀메이트 재고 → 29커넥트(29CM) 옵션 재고 입력 자동화

흐름 (엑셀 한 행마다)
 1. 29CM_상품관리 엑셀 B열 상품명 읽기 (색칠된 행은 건너뜀)
 2. 셀메이트 재고현황에서 검색 → '등급' 오른쪽 '재고' 열 수량을 옵션별로 합산
 3. 29커넥트 상품관리에서 같은 이름 검색 → 판매상태 '판매중'인 상품번호 클릭
 4. 새 창에서 '옵션/재고 설정' 이동 → 옵션명에 맞춰 변환된 재고 입력
 5. 일시정지 → 직접 확인 후 [수정완료]는 본인이 누르고, [계속하기] 클릭
    → 엑셀 B열 초록색 표시 후 다음 행으로

실행 전: '크롬_자동화모드.bat'으로 크롬을 열고 셀메이트/29커넥트에 로그인,
         각각 재고현황 목록 탭, 상품관리 목록 탭을 열어둔 상태여야 합니다.
사용법:  python stock_sync.py            (2행부터)
         python stock_sync.py --start 16 (16행부터)
"""
import os
import re
import sys
import time
import datetime
import tkinter as tk
from tkinter import scrolledtext

import xlwings as xw
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

# ============================== 설정 ==============================
CDP_URL = "http://127.0.0.1:9222"

EXCEL_BOOK = "29CM_상품관리"       # 열려 있는 엑셀 파일 이름(일부)
EXCEL_SHEET = None                 # 시트 이름. None이면 현재 보고 있는 시트
NAME_COL = "B"
DONE_COLOR = (112, 173, 71)        # 완료 표시 초록

SELLMATE_URL_KEY = "sellmate"                 # 셀메이트 탭 주소에 들어있는 글자
CM_LIST_URL_KEY = "partner-item.29cm.co.kr/list"
WAREHOUSE_COUNT = 9                           # 기준 필터에 선택돼 있어야 할 창고 수

LOG_FILE = "재고동기화_로그.txt"
DIAG_DIR = "진단"


def convert_stock(n: int) -> int:
    """셀메이트 재고 → 29CM 입력값 (여기 숫자만 바꾸면 규칙이 바뀝니다)"""
    if n is None or n <= 0:
        return 0
    if n >= 100:
        return 100           # 100개 이상 → 100
    if n >= 50:
        return 50            # 50~99 → 50   (예: 70 → 50)
    if n >= 20:
        return n // 10 * 10  # 20~49 → 일의 자리 버림 (예: 37 → 30)
    if n >= 10:
        return 7             # 10~19 → 7
    return int(n * 0.7)      # 1~9 → 70% 버림 (9→6, 5→3, 1→0)


# 한글/영문 색상명 통일 (옵션명 매칭용)
COLOR_MAP = {
    "블랙": "black", "검정": "black", "화이트": "white", "흰색": "white",
    "아이보리": "ivory", "크림": "cream", "베이지": "beige", "브라운": "brown",
    "카멜": "camel", "그레이": "gray", "grey": "gray", "회색": "gray",
    "차콜": "charcoal", "멜란지": "melange", "네이비": "navy", "블루": "blue",
    "소라": "skyblue", "스카이블루": "skyblue", "하늘": "skyblue", "민트": "mint",
    "그린": "green", "카키": "khaki", "올리브": "olive", "옐로우": "yellow",
    "옐로": "yellow", "레몬": "lemon", "오렌지": "orange", "레드": "red",
    "버건디": "burgundy", "와인": "wine", "핑크": "pink", "퍼플": "purple",
    "라벤더": "lavender", "오트밀": "oatmeal", "모카": "mocha", "실버": "silver",
    "골드": "gold", "데님": "denim", "연청": "lightdenim", "중청": "middenim",
    "진청": "darkdenim", "프리": "free", "f": "free", "fr": "free",
}
# ===================================================================


def log(msg):
    line = f"[{datetime.datetime.now():%H:%M:%S}] {msg}"
    print(line)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def tokens(text: str) -> set:
    t = (text or "").lower()
    t = re.sub(r"[^0-9a-z가-힣]+", " ", t)
    out = set()
    for w in t.split():
        w = COLOR_MAP.get(w, w)
        out.add(w)
    return out


def squash(s):
    return re.sub(r"\s+", "", s or "").lower()


def to_int(s):
    s = re.sub(r"[^\d\-]", "", s or "")
    try:
        return int(s)
    except ValueError:
        return 0


# --------------------------- 일시정지 창 ---------------------------
class Pauser:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("29CM 재고 동기화")
        self.root.attributes("-topmost", True)
        self.root.geometry("640x460")
        self.title = tk.Label(self.root, font=("맑은 고딕", 12, "bold"),
                              wraplength=600, justify="left")
        self.title.pack(anchor="w", padx=10, pady=(10, 4))
        self.text = scrolledtext.ScrolledText(self.root, font=("맑은 고딕", 10))
        self.text.pack(fill="both", expand=True, padx=10)
        self.btns = tk.Frame(self.root)
        self.btns.pack(pady=10)
        self.choice = None
        self.root.protocol("WM_DELETE_WINDOW", lambda: self._set("stop"))
        self.root.withdraw()

    def _set(self, v):
        self.choice = v

    def ask(self, title, body, buttons, pump_page=None):
        """buttons: [(라벨, 값)]. 창을 띄우고 버튼이 눌릴 때까지 대기."""
        for w in self.btns.winfo_children():
            w.destroy()
        for label, val in buttons:
            tk.Button(self.btns, text=label, width=18, font=("맑은 고딕", 10),
                      command=lambda v=val: self._set(v)).pack(side="left", padx=5)
        self.title.config(text=title)
        self.text.delete("1.0", "end")
        self.text.insert("1.0", body)
        self.choice = None
        self.root.deiconify()
        self.root.lift()
        while self.choice is None:
            self.root.update()
            if pump_page is not None:
                try:
                    pump_page.wait_for_timeout(100)   # 브라우저 연결 유지
                except Exception:
                    time.sleep(0.1)
            else:
                time.sleep(0.1)
        self.root.withdraw()
        self.root.update()
        return self.choice


# ----------------------------- 공통 ------------------------------
READ_GRID_JS = r"""
(must) => {
  const norm = s => (s || '').replace(/\s+/g, ' ').trim();
  for (const t of document.querySelectorAll('table')) {
    const headRow = [...t.querySelectorAll('tr')].find(tr => {
      const txt = [...tr.children].map(c => norm(c.innerText));
      return must.every(m => txt.includes(m));
    });
    if (!headRow) continue;
    const headers = [...headRow.children].map(c => norm(c.innerText));
    let body = [...t.querySelectorAll('tr')].filter(tr => tr !== headRow);
    if (!body.some(r => r.children.length === headers.length)) {
      // jqGrid 처럼 헤더/본문 표가 분리된 경우
      const grid = t.closest('.ui-jqgrid, [class*="grid"], [class*="Grid"]');
      if (grid) body = [...grid.querySelectorAll('tr')].filter(tr => tr !== headRow && !t.contains(tr));
    }
    const rows = body.map(tr => [...tr.children].map(c => norm(c.innerText)))
                     .filter(r => r.length >= headers.length - 1 && r.some(x => x));
    return {headers, rows};
  }
  return null;
}
"""


def read_grid(page, must):
    for fr in page.frames:
        try:
            res = fr.evaluate(READ_GRID_JS, must)
            if res:
                return res
        except Exception:
            pass
    return None


def find_in_frames(page, selector):
    for fr in page.frames:
        try:
            loc = fr.locator(selector)
            if loc.count() > 0:
                return loc.first
        except Exception:
            pass
    return None


def settle(page, ms=1200):
    try:
        page.wait_for_load_state("networkidle", timeout=10000)
    except PWTimeout:
        pass
    page.wait_for_timeout(ms)


def dump(page, tag):
    os.makedirs(DIAG_DIR, exist_ok=True)
    stamp = f"{datetime.datetime.now():%H%M%S}_{tag}"
    try:
        page.screenshot(path=os.path.join(DIAG_DIR, f"{stamp}.png"), full_page=True)
    except Exception:
        pass
    for i, fr in enumerate(page.frames):
        try:
            with open(os.path.join(DIAG_DIR, f"{stamp}_frame{i}.html"), "w", encoding="utf-8") as f:
                f.write(fr.content())
        except Exception:
            pass
    log(f"  진단 파일 저장: {DIAG_DIR}\\{stamp}*")


def find_tab(ctx, key):
    for p in ctx.pages:
        if key in p.url:
            return p
    return None


# ---------------------------- 엑셀 -------------------------------
def open_sheet():
    for app in xw.apps:
        for b in app.books:
            if EXCEL_BOOK in b.name:
                return b.sheets[EXCEL_SHEET] if EXCEL_SHEET else b.sheets.active
    sys.exit(f"열려 있는 엑셀에서 '{EXCEL_BOOK}' 파일을 찾지 못했어요.")


def is_colored(rng):
    c = rng.color
    return c is not None and tuple(c) != (255, 255, 255)


# --------------------------- 셀메이트 ----------------------------
SM_SEARCH_INPUT = ("xpath=//th[contains(normalize-space(.),'통합검색')]"
                   "/following-sibling::td[1]//input[@type='text' or not(@type)]")


def sellmate_stock(sm, name, pauser):
    """셀메이트에서 검색 → {옵션문자열: 재고합계}"""
    sm.bring_to_front()

    # 기준 필터 9개 창고 확인
    while True:
        btn = find_in_frames(sm, "xpath=//*[self::button or self::span or self::div]"
                                 "[contains(normalize-space(.),'selected')]")
        txt = btn.inner_text().strip() if btn else ""
        if f"{WAREHOUSE_COUNT} selected" in txt:
            break
        r = pauser.ask("셀메이트 '기준' 필터 확인 필요",
                       f"기준 필터가 '{txt or '확인 불가'}' 상태예요.\n"
                       f"{WAREHOUSE_COUNT}개 창고를 선택한 뒤 [계속하기]를 눌러주세요.\n"
                       f"(이미 맞게 돼 있으면 [그대로 진행])",
                       [("계속하기", "retry"), ("그대로 진행", "go"), ("중지", "stop")], sm)
        if r == "stop":
            raise KeyboardInterrupt
        if r == "go":
            break

    inp = find_in_frames(sm, SM_SEARCH_INPUT)
    if inp is None:
        dump(sm, "셀메이트_검색창없음")
        raise RuntimeError("셀메이트 통합검색 입력칸을 찾지 못했어요.")
    inp.fill(name)
    inp.press("Enter")
    settle(sm)

    grid = read_grid(sm, ["등급", "재고"])
    if not grid:
        dump(sm, "셀메이트_표없음")
        raise RuntimeError("셀메이트 결과 표('등급'·'재고' 열)를 찾지 못했어요.")

    h = grid["headers"]
    gi = h.index("등급")
    si = next((i for i in range(gi + 1, len(h)) if h[i] == "재고"), gi + 1)
    opt_idx = [i for i, x in enumerate(h)
               if re.search(r"옵션|색상|컬러|사이즈", x) and "코드" not in x]
    name_idx = next((i for i, x in enumerate(h) if "상품명" in x), None)

    rows = [r for r in grid["rows"] if len(r) > si]
    if name_idx is not None:
        exact = [r for r in rows if len(r) > name_idx and squash(r[name_idx]) == squash(name)]
        if exact:
            rows = exact

    stock = {}
    for r in rows:
        if opt_idx:
            key = " ".join(r[i] for i in opt_idx if i < len(r)).strip()
        else:
            key = r[name_idx] if name_idx is not None else " ".join(r)
        stock[key] = stock.get(key, 0) + to_int(r[si])
    return stock


# --------------------------- 29커넥트 ----------------------------
CM_NAME_INPUT = "xpath=(//*[normalize-space(text())='상품명']/following::input[1])[1]"


def cm_search(cm, name):
    cm.bring_to_front()
    inp = cm.locator(CM_NAME_INPUT)
    inp.fill(name)
    cm.get_by_role("button", name="검색하기").click()
    settle(cm)
    grid = read_grid(cm, ["상품번호", "상품명", "판매상태"])
    if not grid:
        return []
    h = grid["headers"]
    ni, mi, ti = h.index("상품번호"), h.index("상품명"), h.index("판매상태")
    out = []
    for r in grid["rows"]:
        if len(r) > max(ni, mi, ti):
            out.append({"no": r[ni], "name": r[mi], "status": r[ti]})
    return out


OPTION_ROWS_JS = r"""
() => {
  const norm = s => (s || '').replace(/\s+/g, ' ').trim();
  const val = c => {
    const i = c.querySelector('input:not([type=checkbox]), select');
    return i ? norm(i.value) : norm(c.innerText);
  };
  const res = []; let idx = 0;
  for (const t of document.querySelectorAll('table')) {
    const hr = [...t.querySelectorAll('tr')].find(tr =>
      [...tr.children].some(c => /재고/.test(c.innerText || '')));
    if (!hr) continue;
    const hs = [...hr.children].map(c => norm(c.innerText));
    const si = hs.findIndex(x => /재고/.test(x));
    const oi = hs.map((x, i) => i).filter(i => i !== si && /옵션|컬러|색상|사이즈|size|color/i.test(hs[i]));
    for (const tr of t.querySelectorAll('tr')) {
      if (tr === hr) continue;
      const cells = [...tr.children];
      if (cells.length !== hs.length) continue;
      const inp = cells[si] && cells[si].querySelector('input');
      if (!inp) continue;
      let opt = oi.map(i => val(cells[i])).filter(Boolean).join(' ');
      if (!opt) opt = cells.filter((c, i) => i !== si).map(c => norm(c.innerText)).filter(Boolean).join(' ');
      inp.setAttribute('data-sync-idx', String(idx));
      res.push({idx, option: opt, current: inp.value});
      idx++;
    }
  }
  return res;
}
"""


def fill_options(detail, sm_stock):
    """29CM 옵션 행마다 셀메이트 재고 매칭 후 입력. 결과 리포트 문자열 반환."""
    nav = detail.get_by_text("옵션/재고 설정", exact=True)
    nav.last.click()
    detail.wait_for_timeout(1500)

    rows = detail.evaluate(OPTION_ROWS_JS)
    if not rows:
        dump(detail, "29CM_옵션표없음")
        return None

    sm_tok = {k: tokens(k) for k in sm_stock}
    lines, used = [], set()
    for r in rows:
        ot = tokens(r["option"])
        cands = [k for k, t in sm_tok.items() if ot and ot <= t]
        loc = detail.locator(f'[data-sync-idx="{r["idx"]}"]')
        if not cands:
            loc.evaluate("e => e.style.outline = '3px solid red'")
            lines.append(f"❌ {r['option']:<30} 셀메이트에서 못 찾음 (입력 안 함, 현재 {r['current']})")
            continue
        key = min(cands, key=lambda k: len(sm_tok[k]))
        used.add(key)
        n = sm_stock[key]
        v = convert_stock(n)
        loc.fill(str(v))
        loc.evaluate("e => e.style.outline = '3px solid #2f80ed'")
        lines.append(f"✅ {r['option']:<30} 셀메이트 {n:>4} → 입력 {v:>3}   (기존 {r['current']})")

    extra = [k for k in sm_stock if k not in used]
    if extra:
        lines.append("\n[셀메이트에는 있는데 29CM 옵션과 매칭 안 된 항목]")
        lines += [f"  · {k}: {sm_stock[k]}" for k in extra]
    return "\n".join(lines)


# ----------------------------- 메인 ------------------------------
def main():
    start = 2
    if "--start" in sys.argv:
        start = int(sys.argv[sys.argv.index("--start") + 1])

    sh = open_sheet()
    last = sh.range(f"{NAME_COL}{sh.cells.last_cell.row}").end("up").row
    pauser = Pauser()

    with sync_playwright() as p:
        try:
            browser = p.chromium.connect_over_cdp(CDP_URL)
        except Exception:
            sys.exit("크롬에 연결하지 못했어요. '크롬_자동화모드.bat'으로 크롬을 먼저 열어주세요.")
        ctx = browser.contexts[0]
        sm = find_tab(ctx, SELLMATE_URL_KEY)
        cm = find_tab(ctx, CM_LIST_URL_KEY)
        if not sm or not cm:
            sys.exit("셀메이트 재고현황 탭과 29커넥트 상품관리 탭을 둘 다 열어두세요.")

        log(f"===== 시작: {start}~{last}행 =====")
        try:
            for row in range(start, last + 1):
                cell = sh.range(f"{NAME_COL}{row}")
                name = str(cell.value or "").strip()
                if not name or is_colored(cell):
                    continue
                log(f"{row}행: {name}")

                try:
                    sm_stock = sellmate_stock(sm, name, pauser)
                except RuntimeError as e:
                    r = pauser.ask(f"{row}행 '{name}' - 셀메이트 오류", str(e),
                                   [("다음 상품", "next"), ("중지", "stop")], sm)
                    if r == "stop":
                        break
                    continue
                if not sm_stock:
                    log("  셀메이트 검색 결과 없음 → 건너뜀")
                    continue
                log(f"  셀메이트: {sm_stock}")

                items = [i for i in cm_search(cm, name) if i["status"] == "판매중"]
                if not items:
                    log("  29CM 판매중 상품 없음 → 건너뜀")
                    continue

                done_any, stop = False, False
                for it in items:
                    pages_before = len(ctx.pages)
                    link = cm.get_by_text(it["no"], exact=True).first
                    try:
                        with ctx.expect_page(timeout=8000) as pi:
                            link.click()
                        detail = pi.value
                    except PWTimeout:
                        detail = cm if len(ctx.pages) == pages_before else ctx.pages[-1]
                    detail.wait_for_load_state("domcontentloaded")
                    detail.bring_to_front()
                    detail.get_by_text("옵션/재고 설정", exact=True).last.wait_for(timeout=20000)

                    report = fill_options(detail, sm_stock)
                    if report is None:
                        report = ("옵션/재고 표를 찾지 못해 입력하지 못했어요.\n"
                                  "직접 입력하거나 [건너뛰기]를 눌러주세요.\n"
                                  f"(진단 파일이 '{DIAG_DIR}' 폴더에 저장됐어요)")
                    log("  " + report.replace("\n", "\n  "))

                    r = pauser.ask(
                        f"{row}행 · {it['name']} ({it['no']})",
                        report + "\n\n파란 테두리 = 입력함, 빨간 테두리 = 매칭 실패\n"
                                 "확인 후 브라우저에서 [수정완료]를 직접 누르고 [계속하기]를 눌러주세요.",
                        [("계속하기 (완료 표시)", "ok"), ("건너뛰기", "skip"), ("중지", "stop")],
                        detail)
                    if detail is not cm:
                        try:
                            detail.close()
                        except Exception:
                            pass
                    if r == "ok":
                        done_any = True
                    elif r == "stop":
                        stop = True
                        break

                if done_any:
                    cell.color = DONE_COLOR
                    log("  → 완료 (초록 표시)")
                if stop:
                    break
        except KeyboardInterrupt:
            pass
        log("===== 종료 =====")
        pauser.ask("작업 종료", f"끝났어요. 자세한 기록은 '{LOG_FILE}'에 있어요.",
                   [("닫기", "x")])


if __name__ == "__main__":
    main()
