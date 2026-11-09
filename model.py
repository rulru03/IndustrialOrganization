import json, io
import numpy as np
import openpyxl
book = None
EXPECTED = {'basic': 380, 'small': 165, 'medium': 83, 'large': 33}

def emit(x):
    return json.dumps(x, ensure_ascii=False, allow_nan=False)

def initialize(raw, level):
    global book, codes, names, L, M, V, sheets, valid
    wb = openpyxl.load_workbook(io.BytesIO(bytes(raw)), read_only=True, data_only=True)
    book = {s: list(wb[s].values) for s in wb.sheetnames}
    wb.close()
    sheets = list(book)
    base = book['국산투입계수(Ad)']
    # 부문 코드: 9로 시작하는 코드(합계·소계 등)를 제외 (대분류는 A, C01 같은 문자 코드 포함)
    codes = [str(c) for c in base[4][2:] if c is not None and not str(c).startswith('9')]
    n = EXPECTED[level]
    if len(codes) != n or len(set(codes)) != len(codes):
        raise ValueError('선택한 분류는 %d개 부문이어야 하는데 %d개가 읽혔습니다. 파일을 확인하세요.' % (n, len(codes)))
    nm = {str(r[0]): str(r[1]) for r in base[6:] if r and r[0] is not None}
    names = [nm[c] for c in codes]
    def matrix(s):
        rows = book[s]
        ri = {str(r[0]): i for i, r in enumerate(rows) if r}
        ci = {str(c): i for i, c in enumerate(rows[4])}
        a = np.array([[rows[ri[c]][ci[k]] for k in codes] for c in codes], dtype=float)
        if not np.isfinite(a).all(): raise ValueError(s + ': 비수치 또는 결측 데이터')
        return a
    L = matrix('생산유발계수'); M = matrix('수입유발계수'); V = matrix('부가가치유발계수')
    ad = matrix('국산투입계수(Ad)'); am = matrix('수입투입계수(Am)')
    rows = book['총투입계수(A)']; ci = {str(c): i for i, c in enumerate(rows[4])}
    vr = next(r for r in rows if r and str(r[0]) == '9690')
    v = np.array([vr[ci[c]] for c in codes], dtype=float)
    I = np.eye(n)
    residuals = {'생산': float(np.max(np.abs((I - ad) @ L - I))),
                 '수입': float(np.max(np.abs(am @ L - M))),
                 '부가가치': float(np.max(np.abs(v[:, None] * L - V)))}
    valid = max(residuals.values()) < 1e-8
    if not valid: raise ValueError('계수 검증 실패: ' + str(residuals))
    return emit({'sectors': [{'code': c, 'name': n_} for c, n_ in zip(codes, names)],
                 'sheets': sheets, 'residuals': residuals})

def explore(sheet, rowcode):
    rows = book[sheet]
    choices = [{'code': str(r[0]), 'name': str(r[1])} for r in rows[6:] if r and len(r) > 1 and r[0] is not None and r[1] is not None]
    chosen = next((r for r in rows[6:] if r and str(r[0]) == rowcode), None)
    if chosen is None: chosen = next(r for r in rows[6:] if r and r[0] is not None and r[1] is not None)
    cells = []
    for i in range(2, len(rows[4])):
        if rows[4][i] is None: continue
        val = chosen[i] if i < len(chosen) else None
        if isinstance(val, (float, int)): val = float(val)
        elif val is not None: val = str(val)
        cells.append({'code': str(rows[4][i]), 'name': str(rows[5][i] or '').replace('\n', ' '), 'value': val})
    u = str(rows[3][0] or '')
    unit = u.replace('단위 :', '').strip() if '단위' in u else '단위 없음 (최종수요 1단위당 비율인 계수)'
    return emit({'rows': choices, 'selected': str(chosen[0]), 'name': str(chosen[1]), 'unit': unit, 'cells': cells})

def simulate(payload):
    if not valid: raise ValueError('검증된 데이터가 없습니다.')
    shocks = json.loads(payload); f = np.zeros(len(codes))
    for s in shocks:
        a = float(s['amount'])
        if not np.isfinite(a) or a < 0: raise ValueError('수요 증가액은 유한한 0 이상의 수여야 합니다.')
        f[codes.index(s['code'])] += a
    if not f.sum() > 0: raise ValueError('한 산업 이상의 양의 수요 증가액을 입력하세요.')
    x = L @ f; m = M @ f; v = V @ f
    result = [{'code': c, 'name': n_, 'production': float(a), 'imports': float(b), 'valueAdded': float(d)}
              for c, n_, a, b, d in zip(codes, names, x, m, v)]
    result.sort(key=lambda r: r['production'], reverse=True)
    return emit({'demand': float(f.sum()), 'production': float(x.sum()), 'imports': float(m.sum()),
                 'valueAdded': float(v.sum()), 'multiplier': float(x.sum() / f.sum()), 'rows': result})

def linkage():
    if not valid: raise ValueError('검증된 데이터가 없습니다.')
    col = L.sum(axis=0)   # 열합계: 산업 j 최종수요 1단위가 전 산업에 유발하는 생산
    row = L.sum(axis=1)   # 행합계: 모든 산업 최종수요 1단위씩에 대해 산업 i가 받는 생산
    infl = col / col.mean(); sens = row / row.mean()
    vcol = V.sum(axis=0); mcol = M.sum(axis=0)
    def kind(i, s):
        if i > 1 and s > 1: return '전·후방 연쇄 모두 큼'
        if i > 1: return '후방 연쇄 큼(수요 견인형)'
        if s > 1: return '전방 연쇄 큼(공급 기반형)'
        return '연쇄 작음'
    rows = [{'code': c, 'name': n_, 'influence': float(i), 'sensitivity': float(s),
             'productionMult': float(a), 'valueAddedMult': float(b), 'importMult': float(d), 'type': kind(i, s)}
            for c, n_, i, s, a, b, d in zip(codes, names, infl, sens, col, vcol, mcol)]
    return emit({'rows': rows})
