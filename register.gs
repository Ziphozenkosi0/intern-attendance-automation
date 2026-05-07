const DESTINATION_FOLDER_ID = '1MmwmcTTdq-fgpkQeFPqCuy2rXdjg3tQQ';
const SEND_HOUR             = 7;
const TIMEZONE              = 'Africa/Johannesburg';
const RECIPIENT_EMAIL       = 'brenton@digifycx.com';

const RANGE_START = new Date('2026-05-04');
const RANGE_END   = new Date('2026-10-29');

const INTERNS = [
  ['Sibane Ngubane',               'Day',   'IT'],
  ['Asonwabe Ntshiyantshiya',       'Day',   'IT'],
  ['Amahle Ngwenya',               'Day',   'IT'],
  ['Busiswa Majozi',               'Night', 'IT'],
  ['Tristan Govender',             'Mid',   'IT'],
  ['Rorisang Mokati',              'Mid',   'IT'],
  ['Thulile Gqada',                'Night', 'IT'],
  ['Xolelwa Mandisa Mazibuko',     'Day',   'Marketing / Designing'],
  ['Sabelo Tshazi',                'Day',   'Marketing / Designing'],
  ['Sfundo Xulu',                  'Day',   'Marketing / Designing'],
  ['Nomfundo Mqadi',               'Day',   'Marketing / Designing'],
  ['Sbongokuhle Nene',             'Day',   'Marketing / Designing'],
  ['Hamida Moosa',                 'Night', 'Marketing / Designing'],
  ['Motlalepule Lerato Ivy Khauta','Day',   'Marketing / Designing'],
  ['Lindeka Bele',                 'Night', 'Marketing / Designing'],
  ['Samkelisiwe Gqada',            'Night', 'Marketing / Designing'],
  ['Thoko Sindiswa Xulu',          'Night', 'Marketing / Designing'],
  ['Tokoza Bangani',               'Night', 'Marketing / Designing'],
  ['Alrique Usher',                'Day',   'Zendesk'],
  ['Olwethu Zamokuhle Mahlaba',    'Night', 'Zendesk'],
  ['Nomfundo Mtiyane',             'Day',   'Zendesk'],
  ['Sbahle Ngidi',                 'Mid',   'Zendesk'],
  ['Sanelisiwe Mbele',             'Mid',   'Zendesk'],
  ['Siphiwe Sibisi',               'Night', 'Zendesk'],
  ['Ncebekazi Makhanya',           'Mid',   'Zendesk'],
  ['Nothando Shangase',            'Mid',   'Zendesk'],
  ['Attiyah Bibi Abdul Sathar',    'Night', 'Zendesk'],
  ['Nomonde Bhengu',               'Night', 'Zendesk'],
  ['Nosipho Mbalenhle Nkwanyana',  'Night', 'Zendesk'],
  ['Gugu Xulu',                    'Day',   'Zendesk'],
  ['Alison Rajpal',                'Day',   'Zendesk'],
];

const DAYS = ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday'];
const SKILL_COLORS = {
  'IT':                    { bg: '#D5E8D4', font: '#1A3A1A' },
  'Marketing / Designing': { bg: '#FFE6CC', font: '#3A1A00' },
  'Zendesk':               { bg: '#DAE8FC', font: '#001A3A' },
};

function createWeeklyRegister() {
  const today = new Date();
  if (today.getDay() !== 1) { console.log('Not Monday — skipping.'); return; }
  if (today < RANGE_START || today > RANGE_END) { console.log('Out of range — skipping.'); return; }

  const mondayStr  = formatDate(today);
  const weekEnd    = new Date(today);
  weekEnd.setDate(today.getDate() + 6);
  const fileName   = `Capaciti Register — WC ${zeroPad(today.getDate())}${zeroPad(today.getMonth()+1)} (${mondayStr})`;

  try {
    const folder   = DriveApp.getFolderById(DESTINATION_FOLDER_ID);
    const existing = folder.getFilesByName(fileName);
    if (existing.hasNext()) { console.log('Already exists: ' + fileName); return; }

    const ss = SpreadsheetApp.create(fileName);
    const ws = ss.getActiveSheet();
    ws.setName(`WC ${zeroPad(today.getDate())}${zeroPad(today.getMonth()+1)}`);

    const file = DriveApp.getFileById(ss.getId());
    folder.addFile(file);
    DriveApp.getRootFolder().removeFile(file);

    ws.setColumnWidth(1, 220);
    ws.setColumnWidth(2, 65);
    ws.setColumnWidth(3, 150);
    for (let d = 0; d < 7; d++) {
      ws.setColumnWidth(4 + d * 2, 90);
      ws.setColumnWidth(5 + d * 2, 90);
    }

    // Title
    ws.getRange(1,1,1,17).merge()
      .setValue(`  CAPACITI INTERN ATTENDANCE REGISTER  ·  WEEK COMMENCING ${formatDisplayDate(today).toUpperCase()}  `)
      .setBackground('#1A3A5C').setFontColor('#FFFFFF').setFontFamily('Calibri')
      .setFontSize(13).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
    ws.setRowHeight(1, 36);

    // Day headers row 2
    ws.getRange(2,1,1,3).setValues([['Name','Shift','Skill']])
      .setBackground('#1F4E79').setFontColor('#FFFFFF').setFontFamily('Calibri')
      .setFontSize(10).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true,true,true,true,true,true,'#B0C4DE',SpreadsheetApp.BorderStyle.SOLID);

    for (let d = 0; d < 7; d++) {
      const colStart = 4 + d * 2;
      const dayDate  = new Date(today); dayDate.setDate(today.getDate() + d);
      const isWknd   = d >= 5;
      const bg       = isWknd ? '#595959' : '#1F4E79';
      ws.getRange(2, colStart, 1, 2).merge()
        .setValue(`${DAYS[d]}\n${zeroPad(dayDate.getDate())}/${zeroPad(dayDate.getMonth()+1)}`)
        .setBackground(bg).setFontColor('#FFFFFF').setFontFamily('Calibri')
        .setFontSize(10).setFontWeight('bold').setHorizontalAlignment('center')
        .setVerticalAlignment('middle').setWrap(true)
        .setBorder(true,true,true,true,true,true,'#1F4E79',SpreadsheetApp.BorderStyle.MEDIUM);
    }
    ws.setRowHeight(2, 30);

    // Sub-headers row 3
    ws.getRange(3,1,1,3).setBackground('#1F4E79')
      .setBorder(true,true,true,true,true,true,'#B0C4DE',SpreadsheetApp.BorderStyle.SOLID);
    for (let d = 0; d < 7; d++) {
      const colStart = 4 + d * 2;
      const bg       = d >= 5 ? '#595959' : '#2E75B6';
      ['Attended ','On Time '].forEach((lbl, offset) => {
        ws.getRange(3, colStart + offset).setValue(lbl)
          .setBackground(bg).setFontColor('#FFFFFF').setFontFamily('Calibri')
          .setFontSize(9).setFontWeight('bold').setHorizontalAlignment('center')
          .setVerticalAlignment('middle').setWrap(true)
          .setBorder(true,true,true,true,true,true,'#B0C4DE',SpreadsheetApp.BorderStyle.SOLID);
      });
    }
    ws.setRowHeight(3, 20);

    // Intern rows
    let dataRow   = 4;
    let lastSkill = null;
    INTERNS.forEach(([name, shift, skill], i) => {
      if (skill !== lastSkill) {
        lastSkill = skill;
        const sc  = SKILL_COLORS[skill] || { bg:'#EEEEEE', font:'#000000' };
        ws.getRange(dataRow,1,1,17).merge()
          .setValue(`  ── ${skill.toUpperCase()} ──`)
          .setBackground(sc.bg).setFontColor(sc.font).setFontFamily('Calibri')
          .setFontSize(10).setFontWeight('bold').setHorizontalAlignment('left')
          .setVerticalAlignment('middle')
          .setBorder(true,true,true,true,true,true,'#1F4E79',SpreadsheetApp.BorderStyle.MEDIUM);
        ws.setRowHeight(dataRow, 16);
        dataRow++;
      }
      const bg   = i % 2 === 0 ? '#EBF3FB' : '#FFFFFF';
      const cbBg = i % 2 === 0 ? '#D6EAF8' : '#EAF4FB';

      ws.getRange(dataRow,1).setValue(name).setBackground(bg).setFontFamily('Calibri')
        .setFontSize(10).setHorizontalAlignment('left').setVerticalAlignment('middle')
        .setBorder(true,true,true,true,true,true,'#B0C4DE',SpreadsheetApp.BorderStyle.SOLID);
      ws.getRange(dataRow,2).setValue(shift).setBackground(bg).setFontFamily('Calibri')
        .setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle')
        .setBorder(true,true,true,true,true,true,'#B0C4DE',SpreadsheetApp.BorderStyle.SOLID);
      ws.getRange(dataRow,3).setValue(skill).setBackground(bg).setFontFamily('Calibri')
        .setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle')
        .setBorder(true,true,true,true,true,true,'#B0C4DE',SpreadsheetApp.BorderStyle.SOLID);

      for (let d = 0; d < 7; d++) {
        const colStart = 4 + d * 2;
        const cellBg   = d >= 5 ? '#F2F2F2' : cbBg;
        ws.getRange(dataRow, colStart, 1, 2).insertCheckboxes()
          .setBackground(cellBg).setHorizontalAlignment('center').setVerticalAlignment('middle')
          .setBorder(true,true,true,true,true,true,'#B0C4DE',SpreadsheetApp.BorderStyle.SOLID);
      }
      ws.setRowHeight(dataRow, 20);
      dataRow++;
    });

    // Totals row
    ws.getRange(dataRow,1,1,3).merge()
      .setValue('TOTAL ATTENDED').setBackground('#1F4E79').setFontColor('#FFFFFF')
      .setFontFamily('Calibri').setFontSize(10).setFontWeight('bold')
      .setHorizontalAlignment('left').setVerticalAlignment('middle')
      .setBorder(true,true,true,true,true,true,'#B0C4DE',SpreadsheetApp.BorderStyle.SOLID);
    for (let d = 0; d < 7; d++) {
      const colStart = 4 + d * 2;
      const attCol   = columnLetter(colStart);
      const otCol    = columnLetter(colStart + 1);
      ws.getRange(dataRow, colStart)
        .setFormula(`=COUNTIF(${attCol}4:${attCol}${dataRow-1},TRUE)`)
        .setBackground('#2E75B6').setFontColor('#FFFFFF').setFontWeight('bold')
        .setHorizontalAlignment('center')
        .setBorder(true,true,true,true,true,true,'#B0C4DE',SpreadsheetApp.BorderStyle.SOLID);
      ws.getRange(dataRow, colStart+1)
        .setFormula(`=COUNTIF(${otCol}4:${otCol}${dataRow-1},TRUE)`)
        .setBackground('#2E75B6').setFontColor('#FFFFFF').setFontWeight('bold')
        .setHorizontalAlignment('center')
        .setBorder(true,true,true,true,true,true,'#B0C4DE',SpreadsheetApp.BorderStyle.SOLID);
    }
    ws.setRowHeight(dataRow, 20);
    ws.setFrozenRows(3);
    //ws.setFrozenColumns(3);

    // Email
    const ccEmail = 'mohammed@cxexperts.co.za';
    const subject = ` Capaciti Intern Register — WC ${formatDisplayDate(today)}`;
    const body    = `Dear Brenton,\n\n`
                  + `Please find this week's Capaciti Intern Attendance Register below.\n\n`
                  + ` Register  : ${fileName}\n`
                  + ` Week      : ${formatDisplayDate(today)} – ${formatDisplayDate(weekEnd)}\n`
                  + ` Open Here : ${ss.getUrl()}\n\n`
                  + `Kindly ensure it is completed by end of the week.\n\n`
                  + `— Capaciti Register Automation`;
    ss.addEditor(RECIPIENT_EMAIL);
GmailApp.sendEmail(RECIPIENT_EMAIL, subject, body, { cc: ccEmail });
    console.log('✅ Done! ' + fileName);

  } catch(e) {
    console.error('❌ ' + e.message);
    GmailApp.sendEmail(Session.getActiveUser().getEmail(),
      '⚠️ Capaciti Register Failed', `Error: ${e.message}`);
  }
}

function setupWeeklyTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'createWeeklyRegister')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('createWeeklyRegister')
    .timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(SEND_HOUR).inTimezone(TIMEZONE).create();
  console.log('✅ Trigger set — every Monday at 7:00 AM SAST');
}

function testCreateNow() {
  console.log('🧪 Testing...');
  createWeeklyRegister();
}

function formatDate(d) {
  return `${d.getFullYear()}-${zeroPad(d.getMonth()+1)}-${zeroPad(d.getDate())}`;
}
function formatDisplayDate(d) {
  const m = ['January','February','March','April','May','June','July','August','September','October','November','December'];
  return `${d.getDate()} ${m[d.getMonth()]} ${d.getFullYear()}`;
}
function zeroPad(n) { return String(n).padStart(2,'0'); }
function columnLetter(col) {
  let l = '';
  while (col > 0) { const mod = (col-1)%26; l = String.fromCharCode(65+mod)+l; col = Math.floor((col-1)/26); }
  return l;
}


function testEmailOnly() {
  const ccEmail = Session.getActiveUser().getEmail();
  console.log('Sending to: ' + ccEmail);
  GmailApp.sendEmail(
    RECIPIENT_EMAIL,
    '📋 Test Email — Capaciti Register',
    'This is a test email from the Capaciti Register Automation script.',
    { cc: ccEmail }
  );
  console.log('✅ Email sent!');
}
