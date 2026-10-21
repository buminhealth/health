// Test server: serves the app HTML and routes /script.google.com/macros/s/x/exec to the real Code.gs
// running on the gasmock stand-ins.
const http = require('http');
const fs = require('fs');
const vm = require('vm');
const url = require('url');
const { load } = require('./gasmock');

function start(codePath, htmlPath, port) {
  const g = load(codePath);
  const D = s => vm.runInContext(`new Date(${JSON.stringify(s)})`, g.ctx);
  // 초기 설정 (setupSystem 대신 직접)
  g.props.TOKEN_SECRET = 'test-secret';
  g.props.DATA_ROOT_ID = g.ctx.DriveApp.createFolder('data').getId();
  g.props.ADMIN_PW_SALT = 'salt';
  g.props.ADMIN_PW_HASH = g.api.sha256Hex_('salt' + 'admin-pass-1');
  const ss = g.ctx.SpreadsheetApp.getActiveSpreadsheet();
  const auth = ss.insertSheet('부서인증');
  auth.appendRow(['부서', 'PIN', '담당자', '사용(Y/N)', '부서장 이메일', '비고']);
  auth.appendRow(['61병동', '1234', '', 'Y', 'head61@bumin.co.kr', '']);
  auth.appendRow(['71병동', '5678', '', 'Y', '', '']);
  auth.appendRow(['9병동', '', '', 'Y', '', 'PIN 미입력']);
  const cfgSheet = ss.insertSheet('관리자설정');
  cfgSheet.appendRow(['키', '값', '수정일시']);
  cfgSheet.appendRow(['앱주소', 'https://buminhealth.github.io/health/app.html', '']);
  cfgSheet.appendRow(['보건관리자이메일', 'admin@bumin.co.kr', '']);
  const emp = ss.insertSheet('직원목록');
  emp.appendRow(['사번', '이름', '부서', '입사일', '성별']);
  emp.appendRow(['126051302', '김가나', '61병동', D('2026-06-30T15:00:00Z'), '여']);   // KST 2026-07-01
  emp.appendRow(['126051303', '이다라', '71병동', D('2026-07-01T15:00:00Z'), '여']);
  emp.appendRow(['126051304', '박마바', '61병동', D('2026-07-02T15:00:00Z'), '남']);
  const chem = ss.insertSheet('chemicals');
  chem.appendRow(['제품명', '제조사', '부서', '성분명', 'CAS', '함유량']);
  chem.appendRow(['테스트소독제', 'A사', '61병동', '에탄올', '64-17-5', '70']);

  const server = http.createServer((req, res) => {
    const u = url.parse(req.url, true);
    if (u.pathname === '/script.google.com/macros/s/x/exec') {
      if (req.method === 'GET') {
        const out = g.api.doGet({ parameter: u.query });
        res.writeHead(200, { 'Content-Type': 'application/javascript; charset=utf-8' });
        return res.end(out.text);
      }
      let body = '';
      req.setEncoding('utf8');
      req.on('data', c => body += c);
      req.on('end', () => {
        const out = g.api.doPost({ postData: { contents: body } });
        res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
        res.end(out.text);
      });
      return;
    }
    if (u.pathname === '/' || u.pathname === '/index.html') {
      res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
      return res.end(fs.readFileSync(htmlPath));
    }
    res.writeHead(404); res.end();
  });
  return new Promise(r => server.listen(port, () => r({ server, g })));
}
module.exports = { start };
