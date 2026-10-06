/** 고2 아침 공부 챌린지 - Google Apps Script 서버 코드
 * 아래 한 값만 교사가 설정합니다. 학생의 이름·학번은 저장하지 않습니다.
 */
const SPREADSHEET_ID = '여기에_구글시트_ID_입력';
const SHEET_NAME = '아침공부인증';
const TZ = 'Asia/Seoul';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('고2 아침 공부 챌린지')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getAppData(nickname) {
  nickname = cleanNickname_(nickname);

  const rows = readRows_();
  const mine = rows
    .filter(r => r.nickname === nickname)
    .sort((a, b) => b.createdAt - a.createdAt);

  const feed = rows
    .sort((a, b) => b.createdAt - a.createdAt)
    .slice(0, 20)
    .map(feedItem_);

  return {
    days: mine.length,
    records: mine.slice(0, 20).map(r => ({
      date: r.date,
      promise: r.promise
    })),
    feed: feed
  };
}

function submitChallenge(form) {
  const nickname = cleanNickname_(form.nickname);
  const promise = String(form.promise || '').trim().slice(0, 80);

  if (!promise) {
    throw new Error('오늘의 다짐을 입력해 주세요.');
  }

  if (
    !form.photoData ||
    !/^data:image\/(jpeg|jpg|png|webp);base64,/.test(form.photoData)
  ) {
    throw new Error('인증 사진을 촬영하거나 등록해 주세요.');
  }

  const rows = readRows_();
  const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');

  // 같은 닉네임은 하루에 한 번만 인증
  if (rows.some(r => r.nickname === nickname && r.date === today)) {
    throw new Error('오늘은 이미 인증을 완료했어요.');
  }

  // 사진 DataURL은 시트 셀 제한을 고려해 화면에서 자동 압축됨
  if (form.photoData.length > 48000) {
    throw new Error('사진 용량이 큽니다. 다시 촬영하거나 작은 사진을 등록해 주세요.');
  }

  const sheet = getSheet_();
  const cumulativeDays =
    rows.filter(r => r.nickname === nickname).length + 1;

  sheet.appendRow([
    String(Date.now()),
    new Date(),
    nickname,
    promise,
    form.photoData,
    cumulativeDays
  ]);

  const updatedRows = readRows_();

  const mine = updatedRows
    .filter(r => r.nickname === nickname)
    .sort((a, b) => b.createdAt - a.createdAt);

  const all = updatedRows
    .sort((a, b) => b.createdAt - a.createdAt)
    .slice(0, 20);

  return {
    days: mine.length,
    records: mine.slice(0, 20).map(r => ({
      date: r.date,
      promise: r.promise
    })),
    feed: all.map(feedItem_)
  };
}

function getSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow([
      'ID',
      '인증 시간',
      '닉네임',
      '오늘의 결심',
      '사진 DataURL',
      '누적 공부일수'
    ]);
    sheet.setFrozenRows(1);
  }

  return sheet;
}

function readRows_() {
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();

  if (values.length < 2) return [];

  return values
    .slice(1)
    .filter(row => row[2])
    .map(row => ({
      id: String(row[0]),
      createdAt: new Date(row[1]),
      date: Utilities.formatDate(new Date(row[1]), TZ, 'yyyy-MM-dd'),
      nickname: String(row[2]),
      promise: String(row[3]),
      photo: String(row[4] || ''),
      days: Number(row[5] || 0)
    }));
}

function feedItem_(row) {
  return {
    nickname: row.nickname,
    time: Utilities.formatDate(row.createdAt, TZ, 'M/d HH:mm'),
    promise: row.promise,
    photo: row.photo
  };
}

function cleanNickname_(value) {
  const nickname = String(value || '')
    .trim()
    .replace(/[<>]/g, '');

  if (nickname.length < 2 || nickname.length > 12) {
    throw new Error('닉네임은 2~12글자로 입력해 주세요.');
  }

  return nickname;
}