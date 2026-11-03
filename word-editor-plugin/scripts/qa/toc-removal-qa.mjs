/** Regression: cached TOC rows and live heading counts can disagree.
 * Remove via the ribbon, then inspect the exported DOCX and block list.
 * Run after build: node scripts/qa/toc-removal-qa.mjs
 */
import { openApp, createReport } from './harness.mjs';
import { buildDocx, p, stylesXml } from './docx-fixtures.mjs';

const report = createReport('TOC removal preserves adjacent user content', { strict: true });
const app = await openApp({ name: 'toc-removal', port: Number(process.env.QA_PORT ?? 9512) });
const cases = [
  { name: 'stale cached result', headings: 4, rows: 2 },
  { name: 'no remaining sources', headings: 0, rows: 2 },
  { name: 'single row', headings: 1, rows: 1 },
  { name: 'more than one block-list page', headings: 250, rows: 250 },
];
const styles = stylesXml().replace('</w:styles>', '<w:style w:type="paragraph" w:styleId="TOC1"><w:name w:val="toc 1"/><w:basedOn w:val="Normal"/></w:style></w:styles>');

function fixture({ headings, rows }) {
  const begin = '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>';
  const end = '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
  const cached = Array.from({ length: rows }, (_, index) => '<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr>' + (index === 0 ? begin : '') + `<w:r><w:t>CACHED_ROW_${index}</w:t><w:tab/><w:t>1</w:t></w:r>` + (index === rows - 1 ? end : '') + '</w:p>').join('');
  return buildDocx({ styles, body:
    Array.from({ length: headings }, (_, index) => p(`Heading${index}`, '<w:pStyle w:val="Heading1"/>')).join('') +
    cached + p('USER_MUST_SURVIVE', '<w:pStyle w:val="TOC1"/>') + p('BODY_MUST_SURVIVE'),
  });
}

try {
  await app.cdp.send('Emulation.setDeviceMetricsOverride', { width: 1600, height: 1000, deviceScaleFactor: 1, mobile: false });
  for (const test of cases) {
    const buf = fixture(test);
    const url = 'data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,' + buf.toString('base64');
    await app.js(`window.__qaHost.replies['fs.pickUserFile']=()=>Promise.resolve({success:true,error:null,data:{token:${JSON.stringify(test.name)},url:${JSON.stringify(url)},name:'toc-safety.docx',size:${buf.length},access:'readwrite'}})`);
    await app.tab('קובץ');
    await app.click('פתח קובץ', { after: 500 });
    await app.clickSel('.open-browse', 0, { after: 7000 });
    const before = JSON.parse(await app.js('(async()=>{const d=window.__qa.sd().activeEditor.doc;return JSON.stringify({toc:await d.toc.list({}),blocks:await d.blocks.list({limit:1})})})()'));
    if (before.toc.total !== 1 || before.blocks.total !== test.headings + test.rows + 2) {
      report.fail(test.name, `Invalid fixture: ${JSON.stringify(before)}`);
      continue;
    }
    await app.tab('הפניות');
    await app.click('הסר', { after: 2000 });
    const xml = (await app.docx())['word/document.xml'];
    const after = JSON.parse(await app.js('(async()=>{const d=window.__qa.sd().activeEditor.doc;return JSON.stringify({toc:await d.toc.list({}),blocks:await d.blocks.list({limit:1})})})()'));
    const status = await app.status();
    const messages = await app.messages();
    const safe = xml.includes('USER_MUST_SURVIVE') && xml.includes('BODY_MUST_SURVIVE') &&
      !xml.includes('CACHED_ROW_') && after.toc.total === 0 && after.blocks.total === test.headings + 2 &&
      !status?.error && !messages.some((message) => message.method === 'ui.showError');
    report[safe ? 'pass' : 'fail'](test.name, `${test.rows} cached rows removed, ${test.headings} headings preserved; blocks ${before.blocks.total}→${after.blocks.total}; status=${JSON.stringify(status)}`);
  }
} finally {
  app.close();
}
process.exit(report.print() > 0 ? 1 : 0);
