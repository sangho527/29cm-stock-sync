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
ENDED_COLOR = (255, 255, 0)        # 29CM 판매종료 표시 노랑

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
        return (n - 10) // 10 * 10   # 50~99 → 10 빼고 일의 자리 버림 (예: 70 → 60, 55 → 40)
    if n >= 12:
        return n // 10 * 10          # 12~49 → 일의 자리 버림 (예: 37 → 30, 15 → 10)
    if n >= 6:
        return int(n * 0.7)          # 6~11 → 70% 후 버림 (예: 10 → 7, 6 → 4)
    return n                         # 1~5 → 그대로


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
        self.btns = tk.Frame(self.root)
        self.btns.pack(side="bottom", pady=10)          # 버튼을 항상 맨 아래에 보이게
        self.text = scrolledtext.ScrolledText(self.root, font=("맑은 고딕", 10))
        self.text.pack(fill="both", expand=True, padx=10)
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
    // 합쳐진 칸(rowspan/colspan)을 아래 행에도 채워 넣어 열 위치가 밀리지 않게 함
    const carry = []; const rows = [];
    for (const tr of body) {
      const out = []; let col = 0;
      const fillCarry = () => { while (carry[col] && carry[col].rem > 0) { out[col] = carry[col].text; carry[col].rem--; col++; } };
      for (const c of tr.children) {
        fillCarry();
        const t = norm(c.innerText), rs = c.rowSpan || 1, cs = c.colSpan || 1;
        for (let k = 0; k < cs; k++) { out[col] = t; if (rs > 1) carry[col] = {text: t, rem: rs - 1}; col++; }
      }
      while (col < carry.length) { if (carry[col] && carry[col].rem > 0) { out[col] = carry[col].text; carry[col].rem--; } col++; }
      if (out.length >= headers.length - 1 && out.some(x => x)) rows.push(out.map(x => x || ''));
    }
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


# '통합검색' 드롭다운 바로 뒤에 있는 텍스트 입력칸에 표시를 붙임
MARK_SM_INPUT_JS = r"""
() => {
  const sel = [...document.querySelectorAll('select')].find(s =>
    [...s.options].some(o => o.text.replace(/\s+/g, '') === '통합검색'));
  if (!sel) return false;
  const inputs = [...document.querySelectorAll('input')].filter(i =>
    (!i.type || ['text', 'search'].includes(i.type)) && i.offsetParent !== null &&
    (sel.compareDocumentPosition(i) & Node.DOCUMENT_POSITION_FOLLOWING));
  if (!inputs.length) return false;
  document.querySelectorAll('[data-sync-search]').forEach(e => e.removeAttribute('data-sync-search'));
  inputs[0].setAttribute('data-sync-search', '1');
  return true;
}
"""

SELECTED_JS = r"""
() => {
  for (const el of document.querySelectorAll('button, span, a, div')) {
    const t = (el.innerText || '').replace(/\s+/g, ' ').trim();
    if (/^\d+ selected$/.test(t)) return t;      // 정확히 'N selected' 인 요소만
  }
  return '';
}
"""


def selected_label(page):
    for fr in page.frames:
        try:
            t = fr.evaluate(SELECTED_JS)
            if t:
                return t
        except Exception:
            pass
    return ""


# 셀메이트 결과: 각 행에서 '[수정]'이 붙은 칸 = 옵션, 그 칸 기준으로 '등급' 오른쪽 '재고' 칸을 찾음
SM_ROWS_JS = r"""
() => {
  const norm = s => (s || '').replace(/\s+/g, ' ').trim();
  const head = [...document.querySelectorAll('tr')].find(tr => {
    const t = [...tr.children].map(c => norm(c.innerText));
    return t.includes('상품옵션') && t.includes('재고');
  });
  if (!head) return null;
  const hs = [...head.children].map(c => norm(c.innerText));
  const oh = hs.indexOf('상품옵션'), gi = hs.indexOf('등급');
  let si = gi >= 0 ? hs.findIndex((x, i) => i > gi && x === '재고') : -1;
  if (si < 0) si = hs.indexOf('재고');
  const off = si - oh;
  const out = []; let cur = '', lines = [];
  for (const tr of document.querySelectorAll('tr')) {
    if (tr === head) continue;
    const cells = [...tr.children];
    const pc = cells.find(c => /상품코드|등록일자/.test(c.innerText || ''));
    if (pc) {                                             // 상품정보 칸(첫 옵션 행에만 있음)
      cur = norm(pc.innerText);
      lines = [...pc.querySelectorAll('*')].map(e => norm(e.innerText)).concat(pc.innerText.split('\n').map(norm)).filter(Boolean);
    }
    const oc = cells.findIndex(c => /\[수정\]/.test(c.innerText || '') && norm(c.innerText).length < 60);
    if (oc < 0 || !cells[oc + off]) continue;
    const option = norm(cells[oc].innerText.replace(/\[(수정|내역)\]|\(품절\)/g, ' '));
    out.push({product: cur, lines, option, stock: norm(cells[oc + off].innerText)});
  }
  return out;
}
"""


SM_CACHE = {}   # 이번 실행 동안 셀메이트 검색 결과 기억 (같은 상품의 다른 컬러 행은 다시 검색하지 않음)


def find_sm_input(sm):
    for _ in range(20):                  # 새로 고침 중이면 최대 10초 기다림
        for fr in sm.frames:
            try:
                if fr.evaluate(MARK_SM_INPUT_JS):
                    return fr.locator('[data-sync-search="1"]').first
            except Exception:
                pass
        inp = find_in_frames(sm, SM_SEARCH_INPUT)
        if inp is not None:
            return inp
        sm.wait_for_timeout(500)
    dump(sm, "셀메이트_검색창없음")
    raise RuntimeError("셀메이트 통합검색 입력칸을 찾지 못했어요.")


def sellmate_stock(sm, name, pauser):
    """셀메이트에서 검색 → {옵션문자열: 재고합계}"""
    sm.bring_to_front()

    # 기준 필터 9개 창고 확인
    while True:
        txt = selected_label(sm)
        if txt == f"{WAREHOUSE_COUNT} selected":
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

    # 전체 이름으로 먼저 검색하고, 결과가 없으면 끝 단어(컬러)를 하나씩 빼고 다시 검색
    # 예) 'Louvre Handmade Wool Coat CAMEL' → 'Louvre Handmade Wool Coat'
    words = name.split()
    terms = [" ".join(words[:len(words) - k]) for k in range(0, 3) if len(words) - k >= 2]
    saw_table = False
    for term in terms:
        if term in SM_CACHE:
            if SM_CACHE[term]:
                log(f"  셀메이트 '{term}' 결과 재사용 (이미 검색함)")
                return SM_CACHE[term]
            continue
        inp = find_sm_input(sm)          # 검색하면 페이지가 새로 고쳐지므로 매번 다시 찾음
        inp.fill(term, timeout=10000)
        inp.press("Enter")
        settle(sm)
        res = None
        for fr in sm.frames:
            try:
                res = fr.evaluate(SM_ROWS_JS)
                if res is not None:
                    break
            except Exception:
                pass
        if res is not None:
            saw_table = True
        if not res:
            if res is not None:
                SM_CACHE[term] = {}
            continue
        # 상품명이 정확히 같은 상품만 (비슷한 이름의 다른 상품 재고가 섞이지 않게)
        rows = [r for r in res if any(squash(l) == squash(term) for l in r["lines"])]
        if not rows:
            rows = [r for r in res if squash(term) in squash(r["product"])]
        if not rows:
            SM_CACHE[term] = {}
            continue
        if term != name:
            log(f"  셀메이트는 '{term}'(으)로 검색됨 → 옵션의 컬러로 구분")
        stock = {}
        for r in rows:
            key = r["option"] or "(옵션없음)"
            stock[key] = stock.get(key, 0) + to_int(r["stock"])
        SM_CACHE[term] = stock
        return stock

    if not saw_table:
        dump(sm, "셀메이트_표없음")
    return {}


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


# 29CM 상세 화면의 '옵션/재고 설정' 섹션만 대상으로 함 (다른 섹션은 절대 건드리지 않음)
FIND_SECTION_JS = r"""
() => {
  const sec = [...document.querySelectorAll('section')].find(s => {
    const h = s.querySelector(':scope > button h2');
    return h && h.innerText.replace(/\s+/g, ' ').trim() === '옵션/재고 설정';
  });
  if (!sec) return 'none';
  document.querySelectorAll('[data-sync-section]').forEach(e => e.removeAttribute('data-sync-section'));
  sec.setAttribute('data-sync-section', '1');
  return sec.children.length > 1 ? 'open' : 'closed';   // 접혀 있으면 내용 영역이 없음
}
"""

OPTION_ROWS_JS = r"""
() => {
  const sec = document.querySelector('[data-sync-section="1"]');
  if (!sec) return [];
  const norm = s => (s || '').replace(/\s+/g, ' ').trim();
  const val = c => {
    const i = c.querySelector('input:not([type=checkbox]):not([type=radio]), select');
    if (i) return norm(i.value);
    // 화면에 보이는 글자만, 칩(BROWN / S)마다 띄어서 읽기 (숨은 '색상'·'사이즈' 라벨 제외)
    const parts = []; const w = document.createTreeWalker(c, NodeFilter.SHOW_TEXT);
    while (w.nextNode()) {
      const t = norm(w.currentNode.textContent), el = w.currentNode.parentElement;
      if (!t || !el) continue;
      const st = getComputedStyle(el);
      if (st.display === 'none' || st.visibility === 'hidden' || el.getClientRects().length === 0) continue;
      const r = el.getBoundingClientRect();
      if (r.width <= 1 || r.height <= 1) continue;
      parts.push(t);
    }
    return parts.length ? parts.join(' ') : norm(c.innerText);
  };
  document.querySelectorAll('[data-sync-idx]').forEach(e => e.removeAttribute('data-sync-idx'));
  const res = []; let idx = 0; let hs = null, si = -1, oi = [];
  for (const tr of sec.querySelectorAll('tr')) {          // 섹션 안의 행만
    const cells = [...tr.children];
    if (cells.some(c => c.tagName === 'TH')) {             // 헤더 행
      const h = cells.map(c => norm(c.innerText));
      const k = h.findIndex(x => /재고/.test(x));
      if (k >= 0) {
        hs = h; si = k;
        oi = h.map((x, i) => i).filter(i => h[i] === '옵션명');
        if (!oi.length) oi = h.map((x, i) => i).filter(i => i !== si &&
          /옵션명|컬러|색상|사이즈|size|color/i.test(h[i]) && !/번호|코드|금액|가격|추가|상태/.test(h[i]));
      }
      continue;
    }
    if (!hs || cells.length !== hs.length) continue;
    const inp = cells[si].querySelector('input:not([type=checkbox]):not([type=radio])');
    if (!inp || inp.disabled || inp.readOnly) continue;
    let opt = oi.map(i => val(cells[i])).filter(Boolean).join(' ');
    if (!opt) opt = cells.filter((c, i) => i !== si).map(c => norm(c.innerText)).filter(Boolean).join(' ');
    inp.setAttribute('data-sync-idx', String(idx));
    res.push({idx, option: opt, current: inp.value});
    idx++;
  }
  return res;
}
"""

# 옵션/재고 섹션 밖의 모든 입력값 스냅샷 (입력 전후 비교로 다른 정보가 바뀌지 않았는지 검사)
SNAPSHOT_OUTSIDE_JS = r"""
() => {
  const sec = document.querySelector('[data-sync-section="1"]');
  return [...document.querySelectorAll('input, textarea, select')]
    .filter(e => !(sec && sec.contains(e)))
    .map(e => (e.type === 'checkbox' || e.type === 'radio') ? String(e.checked) : (e.value || ''));
}
"""


def open_option_section(detail):
    """옵션/재고 설정 섹션이 접혀 있으면 펼침. 이미 펼쳐져 있으면 아무것도 누르지 않음."""
    state = detail.evaluate(FIND_SECTION_JS)
    if state == "none":
        return False
    sec = detail.locator('[data-sync-section="1"]')
    if state == "closed":
        sec.locator(":scope > button").click()
        for _ in range(30):
            detail.wait_for_timeout(300)
            if detail.evaluate(FIND_SECTION_JS) == "open":
                break
    sec.scroll_into_view_if_needed()
    detail.wait_for_timeout(800)
    return detail.evaluate(FIND_SECTION_JS) == "open"


def fill_options(detail, sm_stock, product_name=""):
    """29CM 옵션 행마다 셀메이트 재고 매칭 후 입력. 결과 리포트 문자열 반환."""
    if not open_option_section(detail):
        dump(detail, "29CM_옵션섹션못염")
        return None

    rows = detail.evaluate(OPTION_ROWS_JS)
    if not rows:
        dump(detail, "29CM_옵션표없음")
        return None

    before = detail.evaluate(SNAPSHOT_OUTSIDE_JS)

    name_tok = tokens(product_name)
    sm_tok = {k: tokens(k) for k in sm_stock}
    lines, used = [], set()
    for r in rows:
        ot = tokens(r["option"])
        loc = detail.locator(f'[data-sync-section="1"] [data-sync-idx="{r["idx"]}"]')
        # 셀메이트 옵션 단어가 29CM 옵션에 모두 들어 있고, 29CM 옵션 단어는 (셀메이트 옵션 + 상품명)으로 모두 설명돼야 함
        # 셀메이트 옵션 단어는 (29CM 옵션명 + 상품명)에, 29CM 옵션명 단어는 (셀메이트 옵션 + 상품명)에 모두 있어야 짝
        # 예) 셀메이트 'CAMEL M' ↔ 29CM 'CAMEL M' 또는 'M'(상품명에 CAMEL) / 셀메이트 'M'(상품명에 BROWN) ↔ 29CM 'BROWN M'
        cands = [k for k, t in sm_tok.items()
                 if ot and t and t <= (ot | name_tok) and ot <= (t | name_tok)]
        if cands:
            score = lambda k: (len(sm_tok[k] & ot), len(sm_tok[k]))
            best = max(score(k) for k in cands)
            cands = [k for k in cands if score(k) == best]
        if len(cands) != 1:
            loc.evaluate("e => e.style.outline = '3px solid red'")
            why = "셀메이트에서 못 찾음" if not cands else f"후보 여러 개({', '.join(cands)})"
            lines.append(f"❌ {r['option']:<24} {why} → 입력 안 함 (현재 {r['current']})")
            continue
        key = cands[0]
        used.add(key)
        n = sm_stock[key]
        v = convert_stock(n)
        loc.fill(str(v))
        loc.evaluate("e => e.style.outline = '3px solid #2f80ed'")
        lines.append(f"✅ {r['option']:<24} ← 셀메이트 '{key}' {n}개 → 입력 {v}   (기존 {r['current']})")

    extra = [k for k in sm_stock if k not in used]
    if extra:
        lines.append("\n[참고] 이번에 쓰지 않은 셀메이트 옵션 (다른 컬러 옵션이면 정상)")
        lines += [f"  · {k}: {sm_stock[k]}" for k in extra]

    after = detail.evaluate(SNAPSHOT_OUTSIDE_JS)
    if before != after:
        lines.insert(0, "⚠️⚠️ 경고: 옵션/재고 외 다른 칸의 값이 바뀐 것 같아요. "
                        "[수정완료]를 누르지 말고 창을 닫거나 [건너뛰기] 하세요. ⚠️⚠️\n")
    else:
        lines.insert(0, "🔒 옵션/재고 설정 외 다른 정보(가격·판매·배송·이미지 등)는 변경되지 않았어요.\n")
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
                try:
                    log(f"{row}행: {name}")

                    # 29CM 먼저 검색 (판매중 상품이 있을 때만 셀메이트 조회)
                    found = cm_search(cm, name)
                    items = [i for i in found if i["status"] == "판매중"]
                    if not items:
                        if found and all(i["status"] == "판매종료" for i in found):
                            info = "\n".join(f"  · {i['no']}  {i['name']}  [{i['status']}]" for i in found)
                            log("  29CM 검색 결과가 판매종료뿐 → 일시정지")
                            r = pauser.ask(
                                f"{row}행 '{name}' - 판매종료 상품",
                                f"29CM 검색 결과가 판매종료 상품뿐이에요.\n\n{info}\n\n"
                                "[계속하기]를 누르면 엑셀 B열에 노란색으로 표시하고 다음 상품으로 넘어가요.",
                                [("계속하기 (노란색 표시)", "ok"), ("표시 없이 넘어가기", "skip"),
                                 ("중지", "stop")], cm)
                            if r == "stop":
                                break
                            if r == "ok":
                                cell.color = ENDED_COLOR
                                log("  → 판매종료 (노란색 표시)")
                        else:
                            log("  29CM 판매중 상품 없음 → 건너뜀")
                        continue

                    try:
                        sm_stock = sellmate_stock(sm, name, pauser)
                    except RuntimeError as e:
                        r = pauser.ask(f"{row}행 '{name}' - 셀메이트 오류", str(e),
                                       [("다음 상품", "next"), ("중지", "stop")], sm)
                        if r == "stop":
                            break
                        continue
                    if not sm_stock:
                        log("  셀메이트 검색 결과 없음")
                        r = pauser.ask(f"{row}행 '{name}' - 셀메이트 검색 결과 없음",
                                       "셀메이트에서 이 상품을 찾지 못했어요 (끝의 컬러를 빼고 다시 검색해도 없음).\n"
                                       "29CM 재고는 바꾸지 않고 다음 상품으로 넘어가요.",
                                       [("다음 상품", "next"), ("중지", "stop")], sm)
                        if r == "stop":
                            break
                        continue
                    log(f"  셀메이트: {sm_stock}")

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
                        detail.locator("section > button h2", has_text="옵션/재고 설정").first.wait_for(timeout=20000)

                        report = fill_options(detail, sm_stock, f"{name} {it['name']}")   # 엑셀 상품명 + 29CM 상품명의 컬러로 구분
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
                    raise
                except Exception as e:
                    log(f"  예기치 못한 오류: {type(e).__name__}: {e}")
                    for pg in (sm, cm):
                        try:
                            dump(pg, "예기치못한오류")
                        except Exception:
                            pass
                    r = pauser.ask(f"{row}행 '{name}' - 오류",
                                   f"처리 중 오류가 났어요.\n{type(e).__name__}: {str(e)[:500]}\n\n"
                                   "29CM 상세 창이 열려 있다면 [수정완료]를 누르지 말고 닫아주세요.\n"
                                   "[다음 상품]을 누르면 이 행은 그대로 두고 넘어가요.",
                                   [("다음 상품", "next"), ("중지", "stop")])
                    if r == "stop":
                        break

        except KeyboardInterrupt:
            pass
        log("===== 종료 =====")
        pauser.ask("작업 종료", f"끝났어요. 자세한 기록은 '{LOG_FILE}'에 있어요.",
                   [("닫기", "x")])


if __name__ == "__main__":
    main()
