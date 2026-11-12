// 네이버 로그인 버튼 선택 테스트(실제 크롬 필요). 실행: npm run test:login
// 로그인 화면이 가로형·세로형 버튼을 따로 그려도 보이는 '로그인' 버튼을 누르는지 확인한다.
const fs = require("node:fs");
const vm = require("node:vm");
const assert = require("node:assert/strict");
const path = require("node:path");

const source = fs.readFileSync(path.join(__dirname, "naver-booking-sync.js"), "utf8");
const scope = { require, __dirname, process, console, setTimeout };
vm.runInNewContext(
  source.slice(0, source.lastIndexOf("main().catch")) +
    "\nglobalThis.testApi = { chromium, tryNaverCredentialLogin, maskSecrets };",
  scope
);

(async () => {
  const { chromium, tryNaverCredentialLogin, maskSecrets } = scope.testApi;
  assert.equal(maskSecrets("id myid pw secret", { naverId: "myid", naverPassword: "secret" }), "id [아이디] pw [비밀번호]");
  assert.equal(maskSecrets("그대로", { naverId: "", naverPassword: "" }), "그대로");

  const browser = await chromium.launch({ channel: "chrome", headless: true });
  try {
    const page = await browser.newPage();
    for (const visibleId of ["loginBtn_column", "loginBtn_row"]) {
      await page.setContent(`<input id="id"><input id="pw" type="password">
        <button type="button" id="loginBtn_column" style="display:${visibleId === "loginBtn_column" ? "block" : "none"}">로그인</button>
        <button type="button" id="loginBtn_row" style="display:${visibleId === "loginBtn_row" ? "block" : "none"}">로그인</button>
        <button type="button">패스키 로그인</button>
        <script>document.querySelectorAll('[id^=loginBtn]').forEach(button => button.onclick = () => {
          window.clickedLogin = button.id;
          document.querySelectorAll('input').forEach(input => input.remove());
        });</script>`);
      await tryNaverCredentialLogin(page, { naverId: "test", naverPassword: "test", timeoutMs: 2000 });
      assert.equal(await page.evaluate(() => window.clickedLogin), visibleId);
    }
    console.log("PASS: 두 가지 네이버 로그인 화면에서 보이는 로그인 버튼을 누릅니다.");
  } finally {
    await browser.close();
  }
})().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
