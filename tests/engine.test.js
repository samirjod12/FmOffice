const { Workbook, shiftFormula, formatValue, E } = require('../app/js/engine.js');
let pass = 0, fail = 0;
function eq(name, got, want) {
  const g = got && got.code ? got.code : got;
  const ok = typeof want === 'number' && typeof g === 'number' ? Math.abs(g - want) < 1e-9 : g === want;
  if (ok) pass++; else { fail++; console.log('FAIL', name, '=> got', g, 'want', want); }
}
const wb = new Workbook();
const set = (a, t) => { const m = /^([A-Z]+)(\d+)$/.exec(a); const c = [...m[1]].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1; wb.setInput(0, +m[2] - 1, c, t); };
const val = (a) => { const m = /^([A-Z]+)(\d+)$/.exec(a); const c = [...m[1]].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1; return wb.getValue(0, +m[2] - 1, c); };
const calc = f => { set('Z99', f); return val('Z99'); };
[['A1', '10'], ['A2', '20'], ['A3', '30'], ['A4', 'hello'], ['B1', '1'], ['B2', '2'], ['B3', '3'], ['C1', 'apple'], ['C2', 'banana'], ['C3', 'apple']].forEach(([a, t]) => set(a, t));
eq('sum', calc('=SUM(A1:A3)'), 60);
eq('sum text ignored', calc('=SUM(A1:A4)'), 60);
eq('avg', calc('=AVERAGE(A1:A3)'), 20);
eq('prec', calc('=2+3*4'), 14);
eq('pow unary', calc('=-2^2'), 4);
eq('paren', calc('=(2+3)*4'), 20);
eq('pct', calc('=50%'), 0.5);
eq('concat', calc('="a"&"b"&1'), 'ab1');
eq('if', calc('=IF(A1>5,"big","small")'), 'big');
eq('if lazy err', calc('=IF(TRUE,1,1/0)'), 1);
eq('div0', calc('=1/0'), '#DIV/0!');
eq('iferror', calc('=IFERROR(1/0,"x")'), 'x');
eq('vlookup exact', calc('=VLOOKUP(2,B1:C3,2,FALSE)'), 'banana');
eq('vlookup approx', calc('=VLOOKUP(2.5,B1:C3,2)'), 'banana');
eq('index/match', calc('=INDEX(A1:A3,MATCH(20,A1:A3,0))'), 20);
eq('countif', calc('=COUNTIF(C1:C3,"apple")'), 2);
eq('countif op', calc('=COUNTIF(A1:A3,">15")'), 2);
eq('sumif', calc('=SUMIF(C1:C3,"apple",A1:A3)'), 40);
eq('sumifs', calc('=SUMIFS(A1:A3,C1:C3,"apple",B1:B3,">1")'), 30);
eq('sumproduct', calc('=SUMPRODUCT(A1:A3,B1:B3)'), 140);
eq('sumproduct arr', calc('=SUMPRODUCT((A1:A3>10)*B1:B3)'), 5);
eq('round', calc('=ROUND(2.345,2)'), 2.35);
eq('round neg', calc('=ROUND(-2.5,0)'), -3);
eq('left', calc('=LEFT("Hello",2)'), 'He');
eq('proper', calc('=PROPER("hello wORLD")'), 'Hello World');
eq('text fmt', calc('=TEXT(1234.5,"#,##0.00")'), '1,234.50');
eq('text date', calc('=TEXT(DATE(2024,3,5),"yyyy-mm-dd")'), '2024-03-05');
eq('text month', calc('=TEXT(DATE(2024,3,5),"mmmm d, yyyy")'), 'March 5, 2024');
eq('cross sheet', (() => { wb.addSheet('Data'); wb.setInput(1, 0, 0, '99'); return calc("=Data!A1+1"); })(), 100);
eq('cycle', (() => { set('D1', '=D2'); set('D2', '=D1'); return val('D1'); })(), '#CIRC!');
eq('unknown fn', calc('=FOO(1)'), '#NAME?');
eq('bad parse', calc('=1+'), '#NAME?');
eq('compare str', calc('="a"="A"'), true);
eq('whole col', calc('=SUM(B:B)'), 6);
eq('and/or', calc('=AND(1,OR(0,1))'), true);
eq('date diff', calc('=DATEDIF(DATE(2020,1,15),DATE(2024,3,1),"Y")'), 4);
eq('weekday', calc('=WEEKDAY(DATE(2024,3,5))'), 3);
eq('pmt', Math.round(calc('=PMT(0.05/12,360,200000)') * 100) / 100, -1073.64);
// shift
eq('shift', shiftFormula('=A1+$B$2+C$3', 2, 1), '=B3+$B$2+D$3');
eq('shift range', shiftFormula('=SUM(A1:B2)', 1, 0), '=SUM(A2:B3)');
eq('shift off', shiftFormula('=A1', -1, 0), '=#REF!');
// formatting
eq('fmt pct', formatValue(0.256, '0.0%'), '25.6%');
eq('fmt cur', formatValue(-1234.5, '$#,##0.00'), '-$1,234.50');
eq('fmt gen', formatValue(0.1 + 0.2), '0.3');
eq('fmt big', formatValue(1234567890123), '1.23457E+12');
eq('fmt time', formatValue(0.75, 'h:mm AM/PM'), '6:00 PM');
eq('fmt sci', formatValue(12345, '0.00E+00'), '1.23E+04');
// insert / delete rows
const w2 = new Workbook();
w2.setInput(0, 0, 0, '1'); w2.setInput(0, 1, 0, '2'); w2.setInput(0, 2, 0, '3'); w2.setInput(0, 3, 0, '=SUM(A1:A3)');
w2.insertRows(0, 1, 1);
eq('insert row ref', w2.getInput(0, 4, 0), '=SUM(A1:A4)');
eq('insert row val', w2.getValue(0, 4, 0), 6);
w2.deleteRows(0, 1, 1);
eq('delete row ref', w2.getInput(0, 3, 0), '=SUM(A1:A3)');
w2.deleteRows(0, 0, 1);
eq('delete first', w2.getInput(0, 2, 0), '=SUM(A1:A2)');
eq('delete first val', w2.getValue(0, 2, 0), 5);
// json round trip
const w3 = Workbook.fromJSON(JSON.parse(JSON.stringify(w2.toJSON())));
eq('json', w3.getValue(0, 2, 0), 5);
// literal parse
set('E1', '12%'); eq('pct literal', val('E1'), 0.12); eq('pct display', wb.display(0, 0, 4), '12%');
set('E2', '2024-03-05'); eq('date literal display', wb.display(0, 1, 4), '2024-03-05');
set('E3', '1,234'); eq('comma literal', val('E3'), 1234);
set('E4', "'123"); eq('apostrophe', val('E4'), '123');

// ---- dynamic arrays, names, LET, new functions ----
const w4 = new Workbook();
const S4 = (a, t) => { const m = /^([A-Z]+)(\d+)$/.exec(a); const c = [...m[1]].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1; w4.setInput(0, +m[2] - 1, c, t); };
const V4 = a => { const m = /^([A-Z]+)(\d+)$/.exec(a); const c = [...m[1]].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1; const v = w4.getValue(0, +m[2] - 1, c); return v && v.code ? v.code : v; };
[['A1', '5'], ['A2', '3'], ['A3', '9'], ['A4', '1'], ['B1', 'x'], ['B2', 'y'], ['B3', 'x'], ['B4', 'z']].forEach(([a, t]) => S4(a, t));
S4('D1', '=SORT(A1:A4)'); eq('spill sort first', V4('D1'), 1); eq('spill sort last', V4('D4'), 9); eq('spill display', w4.display(0, 3, 3), '9');
S4('F1', '=SUM(D1:D4)'); eq('formula over spill', V4('F1'), 18);
S4('G1', '=FILTER(A1:A4,B1:B4="x")'); eq('filter 1', V4('G1'), 5); eq('filter 2', V4('G2'), 9); eq('filter spill blocked? no', V4('G3'), null);
S4('H1', '=SEQUENCE(3,2)'); eq('seq', V4('I3'), 6);
S4('J1', '=UNIQUE(B1:B4)'); eq('unique', V4('J3'), 'z');
S4('D8', 'blocker'); S4('D7', '=SEQUENCE(3)'); eq('spill blocked', V4('D7'), '#SPILL!');
w4.names.RATE_ = 'Sheet1!$A$1'; w4.names.RNG = 'Sheet1!$A$1:$A$4'; w4.names.K = '2';
S4('L1', '=SUM(RNG)*K'); eq('named range', V4('L1'), 36);
S4('L2', '=LET(x,A1,y,A2,x*y)'); eq('let', V4('L2'), 15);
S4('L3', '=OFFSET(A1,1,0)'); eq('offset', V4('L3'), 3);
S4('L4', '=INDIRECT("A3")'); eq('indirect', V4('L4'), 9);
S4('L5', '=TEXTBEFORE("a-b-c","-",2)'); eq('textbefore', V4('L5'), 'a-b');
S4('L6', '=PERCENTILE(A1:A4,0.5)'); eq('percentile', V4('L6'), 4);
S4('L7', '=ROUND(IRR({-100,60,60}),4)'); // array literal not supported -> NAME
S4('M1', '-100'); S4('M2', '60'); S4('M3', '60'); S4('L8', '=ROUND(IRR(M1:M3),4)'); eq('irr', V4('L8'), 0.1307);
S4('L9', '=ROUND(NPER(0.01,-100,1000),2)'); eq('nper', V4('L9'), 10.59);
S4('L10', '=NETWORKDAYS(DATE(2024,3,4),DATE(2024,3,10))'); eq('networkdays', V4('L10'), 5);
S4('L11', '=SLOPE(A1:A4,M1:M4)'); eq('slope defined', typeof V4('L11'), 'number');
S4('L12', '=HYPERLINK("https://x.y","Click")'); eq('hyperlink', V4('L12'), 'Click');
S4('L13', '=VSTACK(A1:A2,A3:A4)'); eq('vstack', V4('L16'), 1);
w4.insertRows(0, 0, 1); eq('name adjusts on insert', w4.names.RNG, 'Sheet1!$A$2:$A$5');
const j4 = JSON.parse(JSON.stringify(w4.toJSON())); eq('names persist', Workbook.fromJSON(j4).names.K, '2');
console.log(`${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
