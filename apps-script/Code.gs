/**
 * 열매똑똑 해커톤 본선 심사 - Google Apps Script 웹앱
 *
 * 사용법:
 * 1. https://script.google.com 접속
 * 2. "새 프로젝트" 생성
 * 3. 아래 코드 전체를 붙여넣기
 * 4. 상단 메뉴: 배포 → 새 배포 → 유형: 웹앱
 *    - 다음 사용자로 실행: 나
 *    - 액세스 권한: 모든 사용자
 * 5. 첫 배포 직후 권한 승인 화면이 뜨면 "고급 → 안전하지 않은 페이지로 이동"으로 승인
 * 6. 배포 후 표시되는 "웹앱 URL"을 복사
 * 7. 심사표 앱(HTML)의 관리자 모드에서 해당 URL을 입력
 *
 * ⚠️ v4 업데이트: 서명 이미지가 Drive 폴더에 저장되고 시트에는 링크로 첨부됩니다.
 *    - 서명 폴더 ID: 1ykR-pm4h8ONKWhUiT7n47hh-jW3MB78X
 *    - 시트 컬럼이 16개 → 18개로 변경됨 (서명링크, 서명시각 추가)
 *    - 기존 시트가 있다면 "심사결과" 탭을 삭제하면 새 헤더로 자동 재생성됨
 *
 * ⚠️ 권한 주의:
 *    Drive API 권한이 필요합니다. 첫 실행 시 권한 동의 화면이 뜨면 승인하세요.
 *    동일 Google 계정으로 시트와 Drive 폴더 모두에 접근 가능해야 합니다.
 */

const SHEET_ID = '1tlePa82FRUOeAvnn7WlTSOJJLTQ2HOjh8Arh3EmnEIM';
const SHEET_NAME = '심사결과';
const SIGNATURE_FOLDER_ID = '1ykR-pm4h8ONKWhUiT7n47hh-jW3MB78X';

const HEADERS = [
  '저장시각', '심사위원번호', '심사위원성함', '심사위원소속',
  '팀ID', '앱명', '제출기관', '제출자', '부문',
  '문제정의명확성(25)', '실행완성도(25)', '개선효과설득력(25)', '확산가능성(25)',
  '감점(-10~0)', '총점(100)', '심사의견',
  '서명링크', '서명시각'
];

/**
 * POST: 심사 점수 + 서명 저장
 */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(HEADERS);
      sheet.setFrozenRows(1);
      const header = sheet.getRange(1, 1, 1, HEADERS.length);
      header.setBackground('#141416');
      header.setFontColor('#ffffff');
      header.setFontWeight('bold');
      sheet.autoResizeColumns(1, HEADERS.length);
    }

    // 같은 심사위원 + 같은 팀 데이터가 있으면 업데이트
    const lastRow = sheet.getLastRow();
    let updateRow = -1;

    if (lastRow > 1) {
      const range = sheet.getRange(2, 2, lastRow - 1, 4);
      const values = range.getValues();
      for (let i = 0; i < values.length; i++) {
        if (String(values[i][0]) === String(data.judgeId) &&
            String(values[i][3]) === String(data.teamId)) {
          updateRow = i + 2;
          break;
        }
      }
    }

    // 서명 이미지를 Drive에 업로드
    let signatureUrl = '';
    let signedAtFormatted = '';
    if (data.signature) {
      try {
        signatureUrl = uploadSignature(data);
      } catch (sigErr) {
        signatureUrl = '업로드 실패: ' + sigErr.toString();
      }
    }
    if (data.signedAt) {
      try {
        const d = new Date(data.signedAt);
        signedAtFormatted = Utilities.formatDate(d, 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
      } catch (e) {
        signedAtFormatted = String(data.signedAt);
      }
    }

    const now = new Date();
    const timestamp = Utilities.formatDate(now, 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');

    const row = [
      timestamp,
      data.judgeId,
      data.judgeName,
      data.judgeAffiliation || '',
      data.teamId,
      data.teamName,
      data.teamOrg || '',
      data.teamSubmitter || '',
      data.teamSector || '',
      data.scores.clarity,
      data.scores.execution,
      data.scores.impact,
      data.scores.spread,
      data.scores.deduction,
      data.total,
      data.comment || '',
      signatureUrl,
      signedAtFormatted
    ];

    if (updateRow > 0) {
      sheet.getRange(updateRow, 1, 1, row.length).setValues([row]);
    } else {
      sheet.appendRow(row);
    }

    return ContentService
      .createTextOutput(JSON.stringify({
        ok: true,
        updated: updateRow > 0,
        signatureUrl: signatureUrl
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * 서명 base64 PNG를 Drive에 업로드
 * 파일명: signature_J{judgeId}_{teamId}_{timestamp}.png
 * 같은 심사위원+팀의 기존 파일이 있으면 휴지통으로 이동 후 새로 업로드
 */
function uploadSignature(data) {
  const folder = DriveApp.getFolderById(SIGNATURE_FOLDER_ID);

  // 기존 파일 정리 (같은 judgeId + teamId)
  const prefix = `signature_J${data.judgeId}_${data.teamId}_`;
  const existing = folder.searchFiles(`title contains '${prefix}'`);
  while (existing.hasNext()) {
    const f = existing.next();
    if (f.getName().indexOf(prefix) === 0) {
      f.setTrashed(true);
    }
  }

  // base64 디코딩
  let base64 = data.signature;
  if (base64.indexOf('data:image/') === 0) {
    base64 = base64.split(',')[1];
  }
  const bytes = Utilities.base64Decode(base64);

  // 파일명: signature_J1_A1_20260427_143205.png
  const now = new Date();
  const stamp = Utilities.formatDate(now, 'Asia/Seoul', 'yyyyMMdd_HHmmss');
  const filename = `${prefix}${stamp}.png`;

  const blob = Utilities.newBlob(bytes, 'image/png', filename);
  const file = folder.createFile(blob);

  // 공유 설정 (링크 가진 사람 누구나 볼 수 있음)
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) {
    // 도메인 정책으로 공유 제한된 경우 무시
  }

  // 메타데이터를 description에 기록
  const meta = `심사위원: ${data.judgeName} (J${data.judgeId})\n` +
    `작품: ${data.teamId} - ${data.teamName}\n` +
    `제출자: ${data.teamSubmitter || ''}\n` +
    `서명 시각: ${data.signedAt || ''}`;
  file.setDescription(meta);

  return file.getUrl();
}

/**
 * GET: 시트 데이터 조회 (관리자 화면 새로고침용)
 */
function doGet(e) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet || sheet.getLastRow() < 2) {
      return ContentService
        .createTextOutput(JSON.stringify({ ok: true, data: [] }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const lastRow = sheet.getLastRow();
    const values = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();

    const data = values.map(row => ({
      timestamp: row[0] instanceof Date
        ? Utilities.formatDate(row[0], 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss')
        : row[0],
      judgeId: row[1],
      judgeName: row[2],
      judgeAffiliation: row[3],
      teamId: row[4],
      teamName: row[5],
      teamOrg: row[6],
      teamSubmitter: row[7],
      teamSector: row[8],
      clarity: row[9],
      execution: row[10],
      impact: row[11],
      spread: row[12],
      deduction: row[13],
      total: row[14],
      comment: row[15],
      signatureUrl: row[16] || '',
      signedAt: row[17] || ''
    }));

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, data: data }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * 권한 부여용 수동 테스트 함수
 * Apps Script 에디터에서 한 번 실행 → 권한 동의 후 사용 가능
 */
function testFolderAccess() {
  const folder = DriveApp.getFolderById(SIGNATURE_FOLDER_ID);
  Logger.log('폴더명: ' + folder.getName());
  Logger.log('폴더 URL: ' + folder.getUrl());
  const ss = SpreadsheetApp.openById(SHEET_ID);
  Logger.log('시트명: ' + ss.getName());
  Logger.log('✓ 모든 권한 정상');
}
