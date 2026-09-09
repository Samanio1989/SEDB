const CONFIG = {
  SPREADSHEET_ID: '1T-tiELJyShN0j1ugKnZa_omCo0ONYJI9_aYYy4T-dC4',
  SHEET_NAME: 'การตอบแบบฟอร์ม 1',
  APP_TITLE: 'ฐานข้อมูลการศึกษาและวิเคราะห์ชุมชน ด้วยกระบวนการวิศวกรสังคม',
  TIMEZONE: 'Asia/Bangkok'
};

/**
 * Web API สำหรับ GitHub Pages
 * - JSON: เปิด /exec โดยตรง
 * - JSONP: /exec?callback=ชื่อฟังก์ชัน
 */
function doGet(e) {
  try {
    const data = getAppData();
    const callback = e && e.parameter ? clean_(e.parameter.callback) : '';

    if (callback && /^[A-Za-z_$][0-9A-Za-z_$\.]*$/.test(callback)) {
      return ContentService
        .createTextOutput(callback + '(' + JSON.stringify(data) + ');')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }

    return ContentService
      .createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    const errorPayload = {
      error: true,
      message: err && err.message ? err.message : String(err),
      records: [],
      dashboard: emptyDashboard_(),
      filters: emptyFilters_()
    };
    const callback = e && e.parameter ? clean_(e.parameter.callback) : '';
    if (callback && /^[A-Za-z_$][0-9A-Za-z_$\.]*$/.test(callback)) {
      return ContentService
        .createTextOutput(callback + '(' + JSON.stringify(errorPayload) + ');')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService
      .createTextOutput(JSON.stringify(errorPayload))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getAppData() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sh = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sh) throw new Error('ไม่พบชีต: ' + CONFIG.SHEET_NAME);

  const values = sh.getDataRange().getDisplayValues();
  if (!values.length) return emptyPayload_();

  const headers = values[0].map(normalizeHeader_);
  validateHeaders_(headers);

  const rows = values
    .slice(1)
    .filter(row => row.some(v => clean_(v) !== ''));

  const records = rows.map((row, index) => {
    const record = {
      id: index + 1,
      timestamp: get_(row, headers, 'ประทับเวลา'),
      collector: get_(row, headers, '1.1 ชื่อ-สกุล ผู้เก็บข้อมูล'),
      phone: get_(row, headers, '1.2 เบอร์โทรศัพท์ติดต่อ'),
      informant: get_(row, headers, '1.3 ผู้ให้ข้อมูลหลัก'),
      areaName: get_(row, headers, '2.1 ชื่อพื้นที่ดำเนินการ'),
      targetGroup: get_(row, headers, '2.2 กลุ่มเป้าหมาย'),
      location: get_(row, headers, '2.3 ที่ตั้งพื้นที่'),

      // ใช้จังหวัดจากคอลัมน์ 2.6 โดยตรง ไม่วิเคราะห์จากข้อความที่ตั้งพื้นที่
      province: get_(row, headers, '2.6 จังหวัดของพื้นที่'),

      visitDate: get_(row, headers, '2.4 วันที่ลงพื้นที่'),
      issue: get_(row, headers, '2.5 ประเด็นปัญหาหลักของพื้นที่'),
      driveLink: get_(row, headers, '3.1 Link Google Drive เอกสารประกอบ'),
      general: get_(row, headers, '4.1.1 ลักษณะทั่วไปของพื้นที่'),
      problemAnalysis: get_(row, headers, '4.1.2 ปัญหาหลักของพื้นที่ และวิเคราะห์สาเหตุของปัญหา'),
      needs: get_(row, headers, '4.1.3 ความต้องการของพื้นที่'),
      lifeClock: get_(row, headers, 'ข้อมูลนาฬิกาชีวิต'),
      pastPresent: get_(row, headers, '6.1 ลักษณะการดำเนินงานของพื้นที่ในอดีตและปัจจุบัน'),
      turningPoint: get_(row, headers, '6.2 เหตุการณ์สำคัญที่ทำให้เกิดการเปลี่ยนแปลงในการดำเนินงาน'),
      currentProblem: get_(row, headers, '6.3 ปัญหาที่เกิดขึ้นในปัจจุบัน เริ่มตั้งแต่เมื่อใด และลักษณะของปัญหา'),
      currentSituation: get_(row, headers, '6.4 ปัจจุบันสถานการณ์ของพื้นที่'),
      agencies: get_(row, headers, '6.5 หน่วยงานที่เข้ามาดำเนินการในพื้นที่')
    };

    record.targetGroup = clean_(record.targetGroup) || 'ไม่ระบุกลุ่มเป้าหมาย';
    record.province = normalizeProvince_(record.province);
    record.issueTags = splitTags_(record.issue);
    record.driveLink = normalizeDriveLink_(record.driveLink);
    return record;
  });

  const targetCounts = countBy_(records.map(r => r.targetGroup));
  const provinceCounts = countBy_(records.map(r => r.province));
  const issueCounts = countBy_(records.flatMap(r => r.issueTags));

  return {
    appTitle: CONFIG.APP_TITLE,
    source: {
      spreadsheetId: CONFIG.SPREADSHEET_ID,
      sheetName: CONFIG.SHEET_NAME,
      provinceColumn: '2.6 จังหวัดของพื้นที่'
    },
    updatedAt: Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'dd/MM/yyyy HH:mm'),
    records: records,
    dashboard: {
      totalAreas: records.length,
      totalProvinces: Object.keys(provinceCounts).length,
      totalTargetGroups: Object.keys(targetCounts).length,
      totalIssues: Object.keys(issueCounts).length,
      targetGroups: toCountArray_(targetCounts),
      provinces: toCountArray_(provinceCounts),
      issues: toCountArray_(issueCounts)
    },
    filters: {
      provinces: Object.keys(provinceCounts).sort(localeSort_),
      targetGroups: Object.keys(targetCounts).sort(localeSort_),
      issues: Object.keys(issueCounts).sort(localeSort_)
    }
  };
}

function validateHeaders_(headers) {
  const required = [
    '2.1 ชื่อพื้นที่ดำเนินการ',
    '2.2 กลุ่มเป้าหมาย',
    '2.3 ที่ตั้งพื้นที่',
    '2.6 จังหวัดของพื้นที่',
    '2.4 วันที่ลงพื้นที่',
    '2.5 ประเด็นปัญหาหลักของพื้นที่',
    '3.1 Link Google Drive เอกสารประกอบ'
  ];

  const missing = required.filter(name => !headers.includes(normalizeHeader_(name)));
  if (missing.length) {
    throw new Error('ไม่พบคอลัมน์ที่จำเป็น: ' + missing.join(', '));
  }
}

function normalizeProvince_(value) {
  const province = clean_(value)
    .replace(/^จังหวัด\s*/i, '')
    .replace(/^จ\.\s*/i, '')
    .trim();
  return province || 'ไม่ระบุจังหวัด';
}

function normalizeDriveLink_(value) {
  const link = clean_(value);
  if (!link) return '';
  return /^https?:\/\//i.test(link) ? link : '';
}

function emptyPayload_() {
  return {
    appTitle: CONFIG.APP_TITLE,
    source: {
      spreadsheetId: CONFIG.SPREADSHEET_ID,
      sheetName: CONFIG.SHEET_NAME,
      provinceColumn: '2.6 จังหวัดของพื้นที่'
    },
    updatedAt: '',
    records: [],
    dashboard: emptyDashboard_(),
    filters: emptyFilters_()
  };
}

function emptyDashboard_() {
  return {
    totalAreas: 0,
    totalProvinces: 0,
    totalTargetGroups: 0,
    totalIssues: 0,
    targetGroups: [],
    provinces: [],
    issues: []
  };
}

function emptyFilters_() {
  return { provinces: [], targetGroups: [], issues: [] };
}

function normalizeHeader_(value) {
  return String(value || '')
    .replace(/\u00a0/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function get_(row, headers, wanted) {
  const target = normalizeHeader_(wanted);
  const index = headers.indexOf(target);
  return index >= 0 ? clean_(row[index]) : '';
}

function clean_(value) {
  return String(value == null ? '' : value)
    .replace(/\u00a0/g, ' ')
    .replace(/[ \t]+/g, ' ')
    .trim();
}

function splitTags_(text) {
  return [...new Set(
    String(text || '')
      .split(/[,，;；\n]+/)
      .map(clean_)
      .filter(Boolean)
  )];
}

function countBy_(items) {
  return items.reduce((acc, item) => {
    const key = clean_(item);
    if (key) acc[key] = (acc[key] || 0) + 1;
    return acc;
  }, {});
}

function toCountArray_(obj) {
  return Object.entries(obj)
    .map(([name, count]) => ({ name, count }))
    .sort((a, b) => b.count - a.count || localeSort_(a.name, b.name));
}

function localeSort_(a, b) {
  return String(a).localeCompare(String(b), 'th');
}
