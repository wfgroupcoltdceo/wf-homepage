/* =====================================================================
 * WF GROUP 수익 시뮬레이터 — 대리점주(점주) 계산 엔진
 * 기준: "20260718 통합 수익 시뮬레이터 (본사+점주)_11_5.xlsx" 가정·대리점주 시트
 * 모든 수식은 엑셀과 1:1 대응하며, 엑셀 재계산 값과 자동 대조 검증됨.
 * ===================================================================== */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.WFSim = factory();
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  // ---- 종합소득세 누진세율표 (2023년 귀속 이후 현행) ----
  var TAX_BRACKETS = [
    { lower: 0,          rate: 0.06, ded: 0 },
    { lower: 14000001,   rate: 0.15, ded: 1260000 },
    { lower: 50000001,   rate: 0.24, ded: 5760000 },
    { lower: 88000001,   rate: 0.35, ded: 15440000 },
    { lower: 150000001,  rate: 0.38, ded: 19940000 },
    { lower: 300000001,  rate: 0.40, ded: 25940000 },
    { lower: 500000001,  rate: 0.42, ded: 35940000 },
    { lower: 1000000001, rate: 0.45, ded: 65940000 }
  ];

  function bracketOf(x) {
    var b = TAX_BRACKETS[0];
    for (var i = 0; i < TAX_BRACKETS.length; i++) {
      if (x >= TAX_BRACKETS[i].lower) b = TAX_BRACKETS[i];
    }
    return b;
  }
  function incomeTax(x) { // 산출세액
    var b = bracketOf(x);
    return x * b.rate - b.ded;
  }

  // ---- 모델 기본값 (가정 시트 초록/파랑 섹션 — 점주 공개분) ----
  var DEFAULTS = {
    m30: {
      n: 30, joinMonth: 8, term: 24,
      mp: 1800000, ach: 1.0,
      shr: 0.4, lossRate: 0, lossShare: 0.4, cm: 0.4,
      st: 0, fr: 0, pl: 3,
      resNew: 125000, resRenew: 100000, setupFee: 1000000, setupOn: 1,
      server: 250000, taxAgent: 0, inst: 3, dep: 5000000,
      nvShare: 0.2, nvLag: 1, cp1: 0.7, cp1m1: 0.5, cp2Lag: 2, cardLag: 1,
      otherIncome: 40000000, vatType: 1
    },
    m50: {
      n: 50, joinMonth: 8, term: 24,
      mp: 2500000, ach: 1.0,
      shr: 0.4, lossRate: 0, lossShare: 0.4, cm: 0.4,
      st: 0, fr: 0, pl: 3,
      resNew: 115000, resRenew: 100000, setupFee: 1500000, setupOn: 1,
      server: 250000, taxAgent: 0, inst: 3, dep: 7500000,
      nvShare: 0.2, nvLag: 1, cp1: 0.7, cp1m1: 0.5, cp2Lag: 2, cardLag: 1,
      otherIncome: 40000000, vatType: 1
    }
  };

  var VAT = {
    simpleRate: 0.3,          // 간이과세 업종 부가가치율
    simpleExempt: 48000000,   // 납부의무 면제 기준
    simpleKeep: 104000000,    // 간이과세 유지 한도 (연)
    licenseTax: 40500         // 통신판매업 등록면허세 (일반과세 시, 사업자당 연)
  };

  function calMonth(joinMonth, m) { // 개월차 → 달력 월
    return ((joinMonth + m - 2) % 12 + 12) % 12 + 1;
  }
  function calYearOffset(joinMonth, m) { // 개월차 → 시작 연도 대비 몇 해 뒤인지
    return Math.floor((joinMonth - 1 + m - 1) / 12);
  }

  // ---- 본사 설정(관리자 패널) 기본값 — %는 정수, 금액은 만원 단위 ----
  var DEFAULT_CFG = {
    term: 24,
    mp30: 180, mp50: 250, ach: 100,
    shr: 40, lossRate: 0, lossShare: 40, cm: 40,
    st: 0, fr: 0, pl: 3, setupOn: 1,
    resNew30: 12.5, resNew50: 11.5, resRenew30: 10, resRenew50: 10,
    setup30: 100, setup50: 150, server: 25, taxAgent: 0,
    nvShare: 20, nvLag: 1, cp1: 70, cp1m1: 50, cp2Lag: 2, cardLag: 1
  };

  /* ---------------------------------------------------------------
   * 사이트 기본값 오버라이드 (assets/js/wf-defaults.js)
   * 각 시뮬레이터의 "기본값으로 저장"이 만들어 주는 파일.
   * 이 파일이 먼저 로드되어 window.WF_DEFAULTS 를 정의해 두면
   * 아래에서 출고 기본값 위에 덮어써진다 → 모든 방문자에게 적용.
   * 우선순위: 코드 출고값 < wf-defaults.js < 브라우저 localStorage
   * --------------------------------------------------------------- */
  var GLOBAL = (typeof self !== "undefined") ? self : null;
  var SITE = (GLOBAL && GLOBAL.WF_DEFAULTS && typeof GLOBAL.WF_DEFAULTS === "object") ? GLOBAL.WF_DEFAULTS : null;

  function overlay(target, patch, whitelist) {
    if (!patch || typeof patch !== "object") return target;
    for (var k in patch) {
      if (!Object.prototype.hasOwnProperty.call(patch, k)) continue;
      if (whitelist && !Object.prototype.hasOwnProperty.call(whitelist, k)) continue;
      if (typeof patch[k] === "number" && isFinite(patch[k])) target[k] = patch[k];
    }
    return target;
  }

  if (SITE) {
    overlay(DEFAULT_CFG, SITE.cfg, DEFAULT_CFG);
    /* 점주 시뮬레이터 전용 기본값 (본사 설정 패널에 없는 항목)
       sim 그룹은 시뮬레이터 화면의 입력 단위 그대로 쓴다:
         join     0 = 자동(다음 달), 1~12 = 해당 월 고정
         inst     할부 개월
         otherMan 사업 외 연간 소득 (만원)
         vatType  1 = 간이과세, 2 = 일반과세
         setupOn  1 = 세팅비 포함 */
    var simWL = { inst: 1, vatType: 1, term: 1 };
    overlay(DEFAULTS.m30, SITE.sim, simWL);
    overlay(DEFAULTS.m50, SITE.sim, simWL);
    if (SITE.sim) {
      var om = SITE.sim.otherMan;
      if (typeof om === "number" && isFinite(om) && om >= 0) {
        DEFAULTS.m30.otherIncome = DEFAULTS.m50.otherIncome = om * 10000;
      }
      var jm = SITE.sim.join;
      if (typeof jm === "number" && jm >= 1 && jm <= 12) {
        DEFAULTS.m30.joinMonth = DEFAULTS.m50.joinMonth = jm;
      }
    }
  }

  /* 공개 파일에 담아도 되는 그룹만 허용한다.
     본사 내부 변수(hqi: 비상주 원가·본사 수수료율·소개보상·부가세율·을 배분율)는
     암호화된 본사 전용 페이지 안에만 존재해야 하므로 절대 내보내지 않는다. */
  var PUBLIC_GROUPS = { cfg: 1, sim: 1, rev: 1 };

  /**
   * 현재 화면 값을 wf-defaults.js 파일로 내려받는다.
   * patch = { cfg:{...}, sim:{...}, rev:{...} } — 넘긴 그룹만 갱신하고
   * 다른 페이지가 저장해 둔 그룹은 그대로 보존한다.
   * hqi 등 비공개 그룹은 넘겨도 조용히 무시된다.
   */
  function exportDefaults(patch) {
    var base = {};
    if (SITE) { try { base = JSON.parse(JSON.stringify(SITE)); } catch (e) { base = {}; } }
    for (var gg in base) {
      if (!Object.prototype.hasOwnProperty.call(PUBLIC_GROUPS, gg)) delete base[gg];
    }
    for (var g in patch) {
      if (!Object.prototype.hasOwnProperty.call(patch, g)) continue;
      if (!Object.prototype.hasOwnProperty.call(PUBLIC_GROUPS, g)) continue;
      base[g] = Object.assign(base[g] || {}, patch[g]);
    }
    var d = new Date();
    function p2(x) { return (x < 10 ? "0" : "") + x; }
    var stamp = d.getFullYear() + "-" + p2(d.getMonth() + 1) + "-" + p2(d.getDate()) +
      " " + p2(d.getHours()) + ":" + p2(d.getMinutes());
    var txt =
      "/* =====================================================================\n" +
      " * WF GROUP 시뮬레이터 — 사이트 기본값\n" +
      " * 시뮬레이터 화면의 [기본값으로 저장] 버튼이 자동 생성한 파일입니다.\n" +
      " * 생성 시각: " + stamp + "\n" +
      " *\n" +
      " * 사용법: 이 파일을  assets/js/wf-defaults.js  자리에 덮어쓰고\n" +
      " *         GitHub에 푸시하면 모든 방문자의 기본값이 바뀝니다.\n" +
      " *         (직접 손으로 고쳐도 됩니다)\n" +
      " *\n" +
      " * ※ 본사 내부 변수(비상주 원가·본사 수수료율·소개보상·부가세율·을 배분율)는\n" +
      " *   공개 파일에 노출되면 안 되므로 여기에 저장되지 않습니다.\n" +
      " *   해당 값은 본사 전용 페이지의 브라우저에만 저장됩니다.\n" +
      " * ===================================================================== */\n" +
      "window.WF_DEFAULTS = " + JSON.stringify(base, null, 2) + ";\n";
    try {
      var blob = new Blob([txt], { type: "application/javascript;charset=utf-8" });
      var url = URL.createObjectURL(blob);
      var a = document.createElement("a");
      a.href = url; a.download = "wf-defaults.js";
      document.body.appendChild(a); a.click();
      setTimeout(function () { document.body.removeChild(a); URL.revokeObjectURL(url); }, 0);
      return true;
    } catch (e) { return false; }
  }

  /** 사이트 기본값 파일이 실제로 적용 중인지 (UI 안내용) */
  function siteDefaultsActive() { return !!SITE; }

  /** 본사 설정(cfg, DEFAULT_CFG 단위)을 엔진 파라미터에 적용 */
  function applyCfg(p, cfg, model) {
    var m30 = (model === "m30");
    p.mp = (m30 ? cfg.mp30 : cfg.mp50) * 10000;
    p.ach = cfg.ach / 100;
    p.shr = cfg.shr / 100;
    p.lossRate = cfg.lossRate / 100;
    p.lossShare = cfg.lossShare / 100;
    p.cm = cfg.cm / 100;
    p.st = cfg.st; p.fr = cfg.fr; p.pl = cfg.pl;
    p.setupOn = cfg.setupOn;
    p.term = cfg.term;
    p.resNew = (m30 ? cfg.resNew30 : cfg.resNew50) * 10000;
    p.resRenew = (m30 ? cfg.resRenew30 : cfg.resRenew50) * 10000;
    p.setupFee = (m30 ? cfg.setup30 : cfg.setup50) * 10000;
    p.server = cfg.server * 10000;
    p.taxAgent = cfg.taxAgent * 10000;
    p.nvShare = cfg.nvShare / 100;
    p.nvLag = cfg.nvLag;
    p.cp1 = cfg.cp1 / 100;
    p.cp1m1 = cfg.cp1m1 / 100;
    p.cp2Lag = cfg.cp2Lag;
    p.cardLag = cfg.cardLag;
    return p;
  }

  /** 계약 개월 수 보정 — 1~60개월 (기본 24) */
  function termOf(p) {
    var t = Math.round(p && p.term);
    if (!isFinite(t) || t < 1) t = 24;
    if (t > 60) t = 60;
    return t;
  }
  /** 연장(재계약) 재투자가 발생하는 개월차 목록 — 13, 25, 37 … (계약 기간 안쪽만) */
  function renewMonthsOf(term) {
    var out = [];
    for (var m = 13; m <= term; m += 12) out.push(m);
    return out;
  }

  /**
   * 점주 시뮬레이션 — 엑셀 '대리점주 수익 시뮬레이터' 시트와 동일.
   * p.term(계약 개월 수)만큼 계산하며, term=24면 엑셀과 완전히 일치한다.
   */
  function simulate(p) {
    var TERM = termOf(p);
    var YN = Math.ceil(TERM / 12);                              // 연차 수 (마지막 해는 짧을 수 있음)
    var app = p.mp * p.ach;                                     // 가정 B12
    var inv1 = p.resNew * p.n + p.setupFee * p.setupOn + p.server + p.taxAgent; // B6
    var inv2 = p.resRenew * p.n + p.server + p.taxAgent;        // B7
    var RENEW = renewMonthsOf(TERM);
    function isRenew(m) { return RENEW.indexOf(m) >= 0; }
    var Tsched = function (m) { return (m > p.st) ? app * (1 + 1 / p.cm) : 0; }; // 정산예정액(발생월 기준)
    var cpShare = 1 - p.nvShare;
    var net13 = app * p.shr - app * p.lossRate * p.lossShare;   // 정상 월 순수입 (B13)

    // 수수료·손실 송금액 — 매출발생 x개월차분
    function Cf(x) {
      if (x <= p.st) return 0;
      return app * ((x <= p.st + p.fr ? 0 : 1 - p.shr) + p.lossRate * p.lossShare);
    }
    var lumpEnd = p.pl * app * ((1 - p.shr) + p.lossRate * p.lossShare); // 계약 종료 시 잔여분

    // 등록면허세 (일반과세): 가입 시점 + 매년 1월, 사업자당 40,500원
    var licUnit = (p.vatType === 2) ? VAT.licenseTax * p.n : 0;
    var licY = []; for (var yz = 0; yz < YN; yz++) licY.push(0);

    var rows = [];
    var cumD = 0, cumH = 0, cumS = 0, prevL = 0, prevThr = 0;
    var recovery = null;
    for (var m = 1; m <= TERM; m++) {
      var yi = Math.floor((m - 1) / 12);   // 0-based 연차
      var C = (m > p.st) ? app : 0;
      var D = (m > p.st) ? app * (1 + 1 / p.cm) : 0;
      var E = Tsched(m - p.nvLag) * p.nvShare;
      var F = p.cp1 * cpShare * (Tsched(m - 1) * p.cp1m1 + Tsched(m - 2) * (1 - p.cp1m1));
      var G = (1 - p.cp1) * cpShare * Tsched(m - p.cp2Lag);
      var H = E + F + G;
      var Q = (m - p.cardLag > p.st) ? app / p.cm : 0;
      var R = 0;
      if (p.inst > 1) {
        if (m >= 1 + p.cardLag && m < 1 + p.cardLag + p.inst) R += inv1 / p.inst;
        for (var ri = 0; ri < RENEW.length; ri++) {
          var rr = RENEW[ri];
          if (m >= rr + p.cardLag && m < rr + p.cardLag + p.inst) R += inv2 / p.inst;
        }
      }
      var I = Q + R;
      var J = Cf(m - p.pl) + (m === TERM ? lumpEnd : 0);
      var LIC = (licUnit > 0 && (m === 1 || calMonth(p.joinMonth, m) === 1)) ? licUnit : 0;
      licY[yi] += LIC;
      var K = H - I - J - LIC;
      var P = (p.inst <= 1) ? ((m === 1 ? inv1 : 0) + (isRenew(m) ? inv2 : 0)) : 0;
      var flow = (m === 1) ? (K - P + p.dep) : (prevL + K - P);
      var S = Math.max(0, -flow);
      var L = Math.max(0, flow);
      cumD += D; cumH += H; cumS += S;
      var Mrecv = cumD - cumH; // 미수 정산금 잔액
      var thr = p.dep + cumS;
      if (recovery === null && m >= 2 && L >= thr && prevL < prevThr) recovery = m;

      // 송금 대상 매출월 라벨
      var k = m - p.pl, O;
      if (m === TERM && p.pl > 0) {
        O = (k > p.st ? calMonth(p.joinMonth, k) + "월분 + " : "") + "잔여 " + p.pl + "개월분 일시청구(종료 시)";
      } else {
        O = (k <= p.st) ? "-" : calMonth(p.joinMonth, k) + "월분";
      }

      // 발생 기준 월 순수입 (세팅기간 0 · 면제구간 100% 수취 · 이후 배분율)
      var netM = (m <= p.st) ? 0
               : (m <= p.st + p.fr) ? app * (1 - p.lossRate * p.lossShare)
               : net13;
      var invM = (m === 1) ? inv1 : (isRenew(m) ? inv2 : 0);

      rows.push({
        m: m, y: yi + 1, cal: calMonth(p.joinMonth, m), C: C, D: D, E: E, F: F, G: G, H: H,
        I: I, Q: Q, R: R, J: J, K: K, P: P, L: L, S: S, Mrecv: Mrecv, O: O, lic: LIC,
        netM: netM, invM: invM
      });
      prevL = L; prevThr = thr;
    }

    // ---- 종합소득세 (엑셀 6번) ----
    var baseTax = incomeTax(p.otherIncome);
    function addedTax(profit) {
      var tb = p.otherIncome + Math.max(0, profit);
      var add = incomeTax(tb) - baseTax;
      return { income: add, local: add * 0.1, total: add * 1.1, marginal: bracketOf(tb).rate, base: tb };
    }

    /* ---- 연차별 요약 (발생 기준) ----
       마지막 해가 12개월이 안 될 수 있으므로 모든 값을 실제 해당 개월에서 집계한다.
       term=24면 기존 1년차/2년차 계산식과 완전히 동일한 값이 나온다. */
    var years = [];
    for (var y = 0; y < YN; y++) {
      var from = y * 12 + 1, to = Math.min(TERM, (y + 1) * 12);
      var gross = 0, invY = 0, revMon = 0, feeMon = 0;
      for (var t = from - 1; t < to; t++) {
        gross += rows[t].netM;
        invY += rows[t].invM;
        if (rows[t].m > p.st) revMon++;                 // 매출이 발생하는 달
        if (rows[t].m > p.st + p.fr) feeMon++;          // 수수료를 내는 달
      }
      var pre = gross - invY;
      var tax = addedTax(pre);
      // 부가세
      var rev = app * revMon;
      var out = rev * 0.1 / 1.1;
      var fee = app * feeMon * (1 - p.shr);
      var inn = -(fee + invY) * 0.1 / 1.1;
      var general = out + inn;
      var simple = (rev < VAT.simpleExempt) ? 0 : rev * VAT.simpleRate * 0.1;
      var vatY = (p.vatType === 1 ? simple : general) + licY[y];
      years.push({
        y: y + 1, from: from, to: to, months: to - from + 1,
        gross: gross, inv: invY, pre: pre, tax: tax,
        rev: rev, out: out, fee: fee, "in": inn, general: general, simple: simple,
        lic: licY[y], vat: vatY, post: pre - tax.total - vatY,
        roi: invY > 0 ? pre / invY : null,
        roiPost: invY > 0 ? (pre - tax.total - vatY) / invY : null
      });
    }
    var simpleKeepOk = (years.length > 1 ? years[1].rev : years[0].rev) <= VAT.simpleKeep;

    // ---- 누적 수익금 (발생 기준) — 각 연차 마지막 달에 그 해 세금(종소세+부가세+면허세) 반영 ----
    var cumPre = 0, bePre = null, bePost = null;
    for (var t2 = 0; t2 < TERM; t2++) {
      var rw = rows[t2];
      cumPre += rw.netM - rw.invM;
      var taxCum = 0;
      for (var yy = 0; yy < years.length; yy++) {
        if (rw.m >= years[yy].to) taxCum += years[yy].tax.total + years[yy].vat;
      }
      rw.cumPre = cumPre;
      rw.cumPost = cumPre - taxCum;
      if (bePre === null && rw.cumPre >= 0) bePre = rw.m;
      if (bePost === null && rw.cumPost >= 0) bePost = rw.m;
    }

    // 회수 판정 문자열 (엑셀 B87)
    var recoveryLabel;
    if (recovery !== null) recoveryLabel = recovery;
    else recoveryLabel = (rows[0].L >= p.dep + rows[0].S) ? "음수 구간 없음" : (TERM + "개월 내 미회수");

    var Y1 = years[0], Y2 = years[1] || null;
    var last = rows[TERM - 1];
    var mid = rows[Math.min(11, TERM - 1)];   // 12개월차 (계약이 더 짧으면 마지막 달)

    return {
      params: p, term: TERM, yearCount: YN, years: years,
      app: app, inv1: inv1, inv2: inv2, net13: net13, rows: rows,
      renewMonths: RENEW,
      // --- 하위 호환 필드 (기존 화면 코드가 쓰던 이름) ---
      y1Gross: Y1.gross, y1Pre: Y1.pre, y2Gross: Y2 ? Y2.gross : 0, y2Pre: Y2 ? Y2.pre : 0,
      tax1: Y1.tax, tax2: Y2 ? Y2.tax : addedTax(0),
      vat: {
        y1rev: Y1.rev, y1out: Y1.out, y1fee: Y1.fee, y1in: Y1["in"],
        y1general: Y1.general, y1simple: Y1.simple, licY1: Y1.lic, licY2: Y2 ? Y2.lic : 0,
        vat1: Y1.vat, y2general: Y2 ? Y2.general : 0, y2simple: Y2 ? Y2.simple : 0,
        vat2: Y2 ? Y2.vat : 0, simpleKeepOk: simpleKeepOk
      },
      y1Post: Y1.post, y2Post: Y2 ? Y2.post : 0,
      roi1: Y1.roi, roi2: Y2 ? Y2.roi : null,
      roiPost1: Y1.roiPost, roiPost2: Y2 ? Y2.roiPost : null,
      recovery: recoveryLabel,
      breakevenPre: bePre, breakevenPost: bePost,
      cumPreEnd: last.cumPre, cumPostEnd: last.cumPost,
      cumPost12: mid.cumPost, cumPost24: last.cumPost,
      cumPre12: mid.cumPre, cumPre24: last.cumPre,
      bal12: mid.L, recv12: mid.Mrecv,
      totalInject: cumS,
      minBal: Math.min.apply(null, rows.map(function (r) { return r.L; })),
      // 정상월 정산 참고치 (B90~93)
      navMonthly: app * (1 + 1 / p.cm) * p.nvShare,
      cpMonthly: app * (1 + 1 / p.cm) * cpShare,
      cardMonthly: app / p.cm
    };
  }

  /**
   * 권장 통장 예치금 — 현재 조건에서 '추가 현금 투입'이 0이 되는 최소 금액.
   * 예치금은 1개월차에만 유입되므로, 미보정 누적현금(K-P의 누계)의 최대 결손폭이 필요액이다.
   * unit(기본 50만 원) 단위로 올림.
   */
  function recommendDeposit(p, unit) {
    unit = unit || 500000;
    var BIG = 1e11; // 하한(0) 클리핑이 발생하지 않을 만큼 큰 예치금
    var q = {};
    for (var k in p) { if (Object.prototype.hasOwnProperty.call(p, k)) q[k] = p[k]; }
    q.dep = BIG;
    var r = simulate(q);
    var need = 0;
    for (var i = 0; i < r.rows.length; i++) {
      var run = r.rows[i].L - BIG; // = Σ(K-P) 누계 (미보정)
      if (-run > need) need = -run;
    }
    if (need <= 0) return 0;
    return Math.ceil(need / unit) * unit;
  }

  return {
    DEFAULTS: DEFAULTS, DEFAULT_CFG: DEFAULT_CFG, applyCfg: applyCfg, VAT: VAT,
    TAX_BRACKETS: TAX_BRACKETS, simulate: simulate, calMonth: calMonth,
    calYearOffset: calYearOffset, incomeTax: incomeTax, recommendDeposit: recommendDeposit,
    exportDefaults: exportDefaults, siteDefaultsActive: siteDefaultsActive,
    SITE_DEFAULTS: SITE
  };
});
