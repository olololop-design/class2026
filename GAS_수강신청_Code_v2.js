// ============================================================
// 수강신청 앱 v2 - Google Apps Script 백엔드 (비밀번호 인증 추가)
// ============================================================
const SHEET_ID = 'YOUR_GOOGLE_SHEET_ID'; // ← 여기 수정
const MAX_PER_CLASS = 27;
const TOTAL_COURSES = 10;

const COURSES = [
  'AI인공지능', 'K-POP댄스', '게임개발자', '레진공예가', '마술사',
  '바리스타', '쇼콜라티에', '파티시에', '퍼스널컬러전문가', '퍼퓸디자이너'
];

// ============================================================
// doGet
// ============================================================
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// 메인 POST 핸들러
// ============================================================
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    let result;

    if      (action === 'checkStudent')  result = handleCheckStudent(data.grade, data.classNum, data.studentNum);
    else if (action === 'setPassword')   result = handleSetPassword(data.grade, data.classNum, data.studentNum, data.password);
    else if (action === 'login')         result = handleLogin(data.grade, data.classNum, data.studentNum, data.password);
    else if (action === 'getStatus')     result = handleGetStatus(data.grade, data.classNum, data.studentNum);
    else if (action === 'register')      result = handleRegister(data.grade, data.classNum, data.studentNum, data.period, data.courseIndex);
    else if (action === 'getCounts')     result = handleGetCounts();
    else result = { success: false, message: '알 수 없는 액션입니다.' };

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================================
// [1단계] 학생 존재 확인 + 비밀번호 설정 여부 반환
// ============================================================
function handleCheckStudent(grade, classNum, studentNum) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const studentSheet = ss.getSheetByName('학생명단');
  if (!studentSheet) return { success: false, message: '학생명단 시트가 없습니다.' };

  const data = studentSheet.getDataRange().getValues();
  // 헤더: [학년, 반, 번호, 이름, 비밀번호(해시)]
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (String(row[0]) === String(grade) &&
        String(row[1]) === String(classNum) &&
        String(row[2]) === String(studentNum)) {
      const hasPassword = row[4] !== '' && row[4] !== undefined && row[4] !== null;
      return {
        success: true,
        studentName: row[3],
        hasPassword: hasPassword   // true면 로그인, false면 비번 설정
      };
    }
  }
  return { success: false, message: '학번 정보를 찾을 수 없습니다. 반과 번호를 확인해주세요.' };
}

// ============================================================
// [2단계-A] 최초 비밀번호 설정
// ============================================================
function handleSetPassword(grade, classNum, studentNum, password) {
  if (!password || password.length < 4) {
    return { success: false, message: '비밀번호는 4자리 이상이어야 해요.' };
  }

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const studentSheet = ss.getSheetByName('학생명단');
  const data = studentSheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (String(row[0]) === String(grade) &&
        String(row[1]) === String(classNum) &&
        String(row[2]) === String(studentNum)) {

      // 이미 비밀번호가 있으면 설정 불가 (선생님만 초기화 가능)
      if (row[4] !== '' && row[4] !== undefined && row[4] !== null) {
        return { success: false, message: '이미 비밀번호가 설정되어 있어요. 선생님께 초기화를 요청하세요.' };
      }

      // 단순 해시 (GAS에서 SHA-256 불가 → Utilities.computeDigest 사용)
      const hashed = hashPassword(password);
      studentSheet.getRange(i + 1, 5).setValue(hashed);
      SpreadsheetApp.flush();

      return { success: true, studentName: row[3], message: '비밀번호가 설정되었어요! 🎉' };
    }
  }
  return { success: false, message: '학생 정보를 찾을 수 없습니다.' };
}

// ============================================================
// [2단계-B] 로그인 (비밀번호 확인)
// ============================================================
function handleLogin(grade, classNum, studentNum, password) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const studentSheet = ss.getSheetByName('학생명단');
  const data = studentSheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (String(row[0]) === String(grade) &&
        String(row[1]) === String(classNum) &&
        String(row[2]) === String(studentNum)) {

      const storedHash = row[4];
      const inputHash = hashPassword(password);

      if (storedHash === inputHash) {
        return { success: true, studentName: row[3] };
      } else {
        return { success: false, message: '비밀번호가 틀렸어요. 다시 확인해주세요.' };
      }
    }
  }
  return { success: false, message: '학생 정보를 찾을 수 없습니다.' };
}

// ============================================================
// 비밀번호 해시 (SHA-256)
// ============================================================
function hashPassword(password) {
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    password,
    Utilities.Charset.UTF_8
  );
  return digest.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

// ============================================================
// 수강신청 현황 조회 (인증 후 호출)
// ============================================================
function handleGetStatus(grade, classNum, studentNum) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const regSheet = ss.getSheetByName('수강신청');
  if (!regSheet) return { success: true, period1: null, period2: null, counts: getCounts(ss) };

  const data = regSheet.getDataRange().getValues();
  const studentKey = `${grade}-${classNum}-${studentNum}`;
  let period1 = null, period2 = null;

  for (let i = 1; i < data.length; i++) {
    if (`${data[i][0]}-${data[i][1]}-${data[i][2]}` === studentKey) {
      period1 = data[i][3] !== '' ? data[i][3] : null;
      period2 = data[i][4] !== '' ? data[i][4] : null;
      break;
    }
  }

  return { success: true, period1, period2, counts: getCounts(ss) };
}

// ============================================================
// 수강신청 / 변경
// ============================================================
function handleRegister(grade, classNum, studentNum, period, courseIndex) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const regSheet = getOrCreateRegSheet(ss);
    const data = regSheet.getDataRange().getValues();
    const studentKey = `${grade}-${classNum}-${studentNum}`;

    let studentRowIndex = -1;
    let existingP1 = null, existingP2 = null;

    for (let i = 1; i < data.length; i++) {
      if (`${data[i][0]}-${data[i][1]}-${data[i][2]}` === studentKey) {
        studentRowIndex = i + 1;
        existingP1 = data[i][3] !== '' ? Number(data[i][3]) : null;
        existingP2 = data[i][4] !== '' ? Number(data[i][4]) : null;
        break;
      }
    }

    // 같은 강좌 중복 방지
    if (period === 1 && existingP2 === courseIndex)
      return { success: false, message: '2교시에 이미 선택한 강좌는 1교시에 선택할 수 없어요.' };
    if (period === 2 && existingP1 === courseIndex)
      return { success: false, message: '1교시에 이미 선택한 강좌는 2교시에 선택할 수 없어요.' };

    // 정원 확인
    const counts = getCounts(ss);
    const cntArr = period === 1 ? counts.p1 : counts.p2;
    const existingForPeriod = period === 1 ? existingP1 : existingP2;
    if (existingForPeriod !== courseIndex && (cntArr[courseIndex] || 0) >= MAX_PER_CLASS)
      return { success: false, message: `${COURSES[courseIndex]} 강좌가 마감되었어요.` };

    // 저장
    if (studentRowIndex === -1) {
      const newRow = [grade, classNum, studentNum, '', ''];
      if (period === 1) newRow[3] = courseIndex; else newRow[4] = courseIndex;
      regSheet.appendRow(newRow);
    } else {
      regSheet.getRange(studentRowIndex, period === 1 ? 4 : 5).setValue(courseIndex);
    }

    return { success: true, message: `${period}교시 ${COURSES[courseIndex]} 신청 완료!`, counts: getCounts(ss) };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// 전체 현황 조회
// ============================================================
function handleGetCounts() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  return { success: true, counts: getCounts(ss) };
}

function getCounts(ss) {
  const regSheet = ss.getSheetByName('수강신청');
  const p1 = Array(TOTAL_COURSES).fill(0);
  const p2 = Array(TOTAL_COURSES).fill(0);
  if (!regSheet) return { p1, p2 };
  const data = regSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const v1 = data[i][3], v2 = data[i][4];
    if (v1 !== '' && v1 !== null && !isNaN(v1)) p1[Number(v1)]++;
    if (v2 !== '' && v2 !== null && !isNaN(v2)) p2[Number(v2)]++;
  }
  return { p1, p2 };
}

function getOrCreateRegSheet(ss) {
  let sheet = ss.getSheetByName('수강신청');
  if (!sheet) {
    sheet = ss.insertSheet('수강신청');
    sheet.appendRow(['학년', '반', '번호', '1교시', '2교시']);
  }
  return sheet;
}

// ============================================================
// 초기화 함수 (GAS 편집기에서 수동 1회 실행)
// 학생명단 시트: [학년, 반, 번호, 이름, 비밀번호해시]
// ============================================================
function initializeSheets() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  let studentSheet = ss.getSheetByName('학생명단');
  if (!studentSheet) studentSheet = ss.insertSheet('학생명단');
  studentSheet.clearContents();
  studentSheet.appendRow(['학년', '반', '번호', '이름', '비밀번호(해시)']);

  // 샘플 데이터 (실제 명단으로 교체하세요)
  const sampleData = [];
  for (let c = 1; c <= 9; c++) {
    for (let n = 1; n <= 27; n++) {
      sampleData.push([3, c, n, `3-${c}-${n}번학생`, '']); // 비번 빈칸 = 미설정
    }
  }
  studentSheet.getRange(2, 1, sampleData.length, 5).setValues(sampleData);

  let regSheet = ss.getSheetByName('수강신청');
  if (!regSheet) regSheet = ss.insertSheet('수강신청');
  regSheet.clearContents();
  regSheet.appendRow(['학년', '반', '번호', '1교시', '2교시']);

  SpreadsheetApp.flush();
  Logger.log('초기화 완료! 학생명단 이름을 실제 학생으로 교체하세요.');
}

// ============================================================
// 특정 학생 비밀번호 초기화 (선생님이 GAS 편집기에서 실행)
// 실행 전 아래 반/번호를 수정하세요
// ============================================================
function resetStudentPassword() {
  const TARGET_CLASS = 1;   // ← 반 수정
  const TARGET_NUM   = 1;   // ← 번호 수정

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('학생명단');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]) === String(TARGET_CLASS) &&
        String(data[i][2]) === String(TARGET_NUM)) {
      sheet.getRange(i + 1, 5).setValue('');
      SpreadsheetApp.flush();
      Logger.log(`3학년 ${TARGET_CLASS}반 ${TARGET_NUM}번 비밀번호 초기화 완료`);
      return;
    }
  }
  Logger.log('학생을 찾을 수 없습니다.');
}
