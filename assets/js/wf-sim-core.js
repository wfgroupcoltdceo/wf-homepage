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
      n: 30, joinMonth: 8,
      mp: 1800000, ach: 1.0,
      shr: 0.4, lossRate: 0.04, lossShare: 0.4, cm: 0.4,
      st: 1, fr: 1, pl: 3,
      resNew: 125000, resRenew: 100000, setupFee: 1000000, setupOn: 1,
      server: 250000, taxAgent: 0, inst: 3, dep: 5000000,
      nvShare: 0.2, nvLag: 1, cp1: 0.7, cp1m1: 0.5, cp2Lag: 2, cardLag: 1,
      otherIncome: 40000000, vatType: 1
    },
    m50: {
      n: 50, joinMonth: 8,
      mp: 2500000, ach: 1.0,
      shr: 0.4, lossRate: 0.04, lossShare: 0.4, cm: 0.4,
      st: 1, fr: 1, pl: 3,
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

  // ---- 본사 설정(관리자 패널) 기본값 — %는 정수, 금액은 만원 단위 ----
  var DEFAULT_CFG = {
    mp30: 180, mp50: 250, ach: 100,
    shr: 40, lossRate: 4, lossShare: 40, cm: 40,
    st: 1, fr: 1, pl: 3, setupOn: 1,
    resNew30: 12.5, resNew50: 11.5, resRenew30: 10, resRenew50: 10,
    setup30: 100, setup50: 150, server: 25, taxAgent: 0,
    nvShare: 20, nvLag: 1, cp1: 70, cp1m1: 50, cp2Lag: 2, cardLag: 1
  };

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

  /** 점주 시뮬레이션 — 엑셀 '대리점주 수익 시뮬레이터' 시트와 동일 */
  function simulate(p) {
    var app = p.mp * p.ach;                                     // 가정 B12
    var inv1 = p.resNew * p.n + p.setupFee * p.setupOn + p.server + p.taxAgent; // B6
    var inv2 = p.resRenew * p.n + p.server + p.taxAgent;        // B7
    var T = function (m) { return (m > p.st) ? app * (1 + 1 / p.cm) : 0; }; // 정산예정액(발생월 기준)
    var cpShare = 1 - p.nvShare;
    var net13 = app * p.shr - app * p.lossRate * p.lossShare;   // 정상 월 순수입 (B13)

    // 수수료·손실 송금액 — 매출발생 x개월차분
    function Cf(x) {
      if (x <= p.st) return 0;
      return app * ((x <= p.st + p.fr ? 0 : 1 - p.shr) + p.lossRate * p.lossShare);
    }
    var lump24 = p.pl * app * ((1 - p.shr) + p.lossRate * p.lossShare); // 종료 시 잔여분

    // 등록면허세 (일반과세): 가입 시점 + 매년 1월, 사업자당 40,500원
    var licUnit = (p.vatType === 2) ? VAT.licenseTax * p.n : 0;
    var licY1 = 0, licY2 = 0;

    var rows = [];
    var cumD = 0, cumH = 0, cumS = 0, prevL = 0, prevThr = 0;
    var recovery = null;
    for (var m = 1; m <= 24; m++) {
      var C = (m > p.st) ? app : 0;
      var D = (m > p.st) ? app * (1 + 1 / p.cm) : 0;
      var E = T(m - p.nvLag) * p.nvShare;
      var F = p.cp1 * cpShare * (T(m - 1) * p.cp1m1 + T(m - 2) * (1 - p.cp1m1));
      var G = (1 - p.cp1) * cpShare * T(m - p.cp2Lag);
      var H = E + F + G;
      var Q = (m - p.cardLag > p.st) ? app / p.cm : 0;
      var R = 0;
      if (p.inst > 1) {
        if (m >= 1 + p.cardLag && m < 1 + p.cardLag + p.inst) R += inv1 / p.inst;
        if (m >= 13 + p.cardLag && m < 13 + p.cardLag + p.inst) R += inv2 / p.inst;
      }
      var I = Q + R;
      var J = Cf(m - p.pl) + (m === 24 ? lump24 : 0);
      var LIC = (licUnit > 0 && (m === 1 || calMonth(p.joinMonth, m) === 1)) ? licUnit : 0;
      if (m <= 12) licY1 += LIC; else licY2 += LIC;
      var K = H - I - J - LIC;
      var P = (p.inst <= 1) ? ((m === 1 ? inv1 : 0) + (m === 13 ? inv2 : 0)) : 0;
      var flow = (m === 1) ? (K - P + p.dep) : (prevL + K - P);
      var S = Math.max(0, -flow);
      var L = Math.max(0, flow);
      cumD += D; cumH += H; cumS += S;
      var Mrecv = cumD - cumH; // 미수 정산금 잔액
      var thr = p.dep + cumS;
      if (recovery === null && m >= 2 && L >= thr && prevL < prevThr) recovery = m;

      // 송금 대상 매출월 라벨
      var k = m - p.pl, O;
      if (m === 24 && p.pl > 0) {
        O = (k > p.st ? calMonth(p.joinMonth, k) + "월분 + " : "") + "잔여 " + p.pl + "개월분 일시청구(종료 시)";
      } else {
        O = (k <= p.st) ? "-" : calMonth(p.joinMonth, k) + "월분";
      }

      // 발생 기준 월 순수입 (세팅기간 0 · 면제구간 100% 수취 · 이후 배분율)
      var netM = (m <= p.st) ? 0
               : (m <= p.st + p.fr) ? app * (1 - p.lossRate * p.lossShare)
               : net13;
      var invM = (m === 1) ? inv1 : (m === 13 ? inv2 : 0);

      rows.push({
        m: m, cal: calMonth(p.joinMonth, m), C: C, D: D, E: E, F: F, G: G, H: H,
        I: I, Q: Q, R: R, J: J, K: K, P: P, L: L, S: S, Mrecv: Mrecv, O: O, lic: LIC,
        netM: netM, invM: invM
      });
      prevL = L; prevThr = thr;
    }

    // ---- 연차 요약 (발생 기준) ----
    var y1Gross = app * p.fr * (1 - p.lossRate * p.lossShare) + net13 * (12 - p.st - p.fr); // B71
    var y1Pre = y1Gross - inv1;   // B73 세전
    var y2Gross = net13 * 12;     // B79
    var y2Pre = y2Gross - inv2;   // B81

    // ---- 종합소득세 (엑셀 6번) ----
    var baseTax = incomeTax(p.otherIncome);
    function addedTax(profit) {
      var tb = p.otherIncome + Math.max(0, profit);
      var add = incomeTax(tb) - baseTax;
      return { income: add, local: add * 0.1, total: add * 1.1, marginal: bracketOf(tb).rate, base: tb };
    }
    var tax1 = addedTax(y1Pre), tax2 = addedTax(y2Pre);

    // ---- 부가세 (엑셀 7번) ----
    var y1rev = app * (12 - p.st);                              // B130
    var y1out = y1rev * 0.1 / 1.1;                              // B131
    var y1fee = (y1rev - app * p.fr) * (1 - p.shr);             // B132
    var y1in = -(y1fee + inv1) * 0.1 / 1.1;                     // B133
    var y1general = y1out + y1in;                               // B134
    var y1simple = (y1rev < VAT.simpleExempt) ? 0 : y1rev * VAT.simpleRate * 0.1; // B135
    var vat1 = (p.vatType === 1 ? y1simple : y1general) + licY1;  // B137 — 면허세는 실제 납부 횟수 반영
    var y2rev = app * 12;
    var y2out = y2rev * 0.1 / 1.1;
    var y2in = -(y2rev * (1 - p.shr) + inv2) * 0.1 / 1.1;
    var y2general = y2out + y2in;
    var y2simple = (y2rev < VAT.simpleExempt) ? 0 : y2rev * VAT.simpleRate * 0.1;
    var vat2 = (p.vatType === 1 ? y2simple : y2general) + licY2;  // B141
    var simpleKeepOk = y2rev <= VAT.simpleKeep;

    var y1Post = y1Pre - tax1.total - vat1;  // B76
    var y2Post = y2Pre - tax2.total - vat2;  // B84

    // ---- 누적 수익금 (발생 기준) ----
    // 12개월차에 1년차 세금(종소세+부가세+면허세), 24개월차에 2년차 세금을 반영
    var cumPre = 0, bePre = null, bePost = null;
    for (var t = 0; t < 24; t++) {
      var rw = rows[t];
      cumPre += rw.netM - rw.invM;
      var taxCum = (rw.m >= 12 ? tax1.total + vat1 : 0) + (rw.m >= 24 ? tax2.total + vat2 : 0);
      rw.cumPre = cumPre;
      rw.cumPost = cumPre - taxCum;
      if (bePre === null && rw.cumPre >= 0) bePre = rw.m;
      if (bePost === null && rw.cumPost >= 0) bePost = rw.m;
    }

    // 회수 판정 문자열 (엑셀 B87)
    var recoveryLabel;
    if (recovery !== null) recoveryLabel = recovery;
    else recoveryLabel = (rows[0].L >= p.dep + rows[0].S) ? "음수 구간 없음" : "24개월 내 미회수";

    return {
      params: p, app: app, inv1: inv1, inv2: inv2, net13: net13, rows: rows,
      y1Gross: y1Gross, y1Pre: y1Pre, y2Gross: y2Gross, y2Pre: y2Pre,
      tax1: tax1, tax2: tax2,
      vat: {
        y1rev: y1rev, y1out: y1out, y1fee: y1fee, y1in: y1in,
        y1general: y1general, y1simple: y1simple, licY1: licY1, licY2: licY2, vat1: vat1,
        y2general: y2general, y2simple: y2simple, vat2: vat2, simpleKeepOk: simpleKeepOk
      },
      y1Post: y1Post, y2Post: y2Post,
      roi1: y1Pre / inv1, roi2: y2Pre / inv2,
      recovery: recoveryLabel,
      breakevenPre: bePre, breakevenPost: bePost,
      cumPost12: rows[11].cumPost, cumPost24: rows[23].cumPost,
      bal12: rows[11].L, recv12: rows[11].Mrecv,
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

  return { DEFAULTS: DEFAULTS, DEFAULT_CFG: DEFAULT_CFG, applyCfg: applyCfg, VAT: VAT, TAX_BRACKETS: TAX_BRACKETS, simulate: simulate, calMonth: calMonth, incomeTax: incomeTax, recommendDeposit: recommendDeposit };
});
