/** 웹 앱 배포: 실행 사용자 ‘나’, 액세스 ‘모든 사용자’으로 설정하세요. */
const FOLDER_ID = '1_VhjNfTPKWy-v_WsY-fZ6jT7VAaqN4Cg';
const MAX_BYTES = 10 * 1024 * 1024;

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Upload')
    .setTitle('집회신고 접수증 사진 보관')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// HTML form을 google.script.run으로 전달하면 파일이 Blob으로 수신됩니다.
function uploadReceipt(form) {
  if (!form || !['경기본부', '인천본부', '서울본부', '대세충본부'].includes(form.headquarters)) {
    throw new Error('본부를 선택해 주세요.');
  }
  const metadata = {headquarters: form.headquarters};
  for (const [key, label, limit] of [['company', '건설사명', 60], ['officerContact', '정보관 이름 및 연락처', 120], ['reporter', '신고인 이름', 60]]) {
    const value = typeof form[key] === 'string' ? form[key].trim() : '';
    if (!value || value.length > limit) throw new Error(label + '을(를) ' + limit + '자 이내로 입력해 주세요.');
    metadata[key] = value;
  }
  const blob = form.receipt;
  if (!blob || typeof blob.getBytes !== 'function') throw new Error('사진을 선택해 주세요.');
  const bytes = blob.getBytes();
  if (!bytes.length || bytes.length > MAX_BYTES) throw new Error('사진은 10MB 이하로 선택해 주세요.');
  const u = i => (bytes[i] + 256) % 256;
  let type, extension;
  if (u(0) === 255 && u(1) === 216 && u(2) === 255) {
    type = 'image/jpeg'; extension = 'jpg';
  } else if ([137,80,78,71,13,10,26,10].every((v,i) => u(i) === v)) {
    type = 'image/png'; extension = 'png';
  } else if (bytes.length >= 12 && String.fromCharCode(...bytes.slice(0,4).map(x=>(x+256)%256)) === 'RIFF' && String.fromCharCode(...bytes.slice(8,12).map(x=>(x+256)%256)) === 'WEBP') {
    type = 'image/webp'; extension = 'webp';
  } else {
    throw new Error('JPG, PNG, WEBP 사진만 저장할 수 있습니다. HEIC 사진은 JPG로 변환해 주세요.');
  }
  const clean = value => String(value || '').trim().replace(/[\\/:*?"<>|\r\n\t]/g, '_').slice(0,60);
  const company = clean(form.company) || '건설사미입력';
  const timestamp = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyyMMdd_HHmmss');
  const name = `${form.headquarters}_${company}_${timestamp}_${Utilities.getUuid().slice(0,8)}.${extension}`;
  const file = DriveApp.getFolderById(FOLDER_ID).createFile(blob.setContentType(type).setName(name));
  try {
    // 사진과 입력 정보를 하나로 보관: Drive 파일 상세정보의 ‘설명’에서 확인.
    file.setDescription(JSON.stringify({...metadata, submittedAt: new Date().toISOString()}, null, 2));
  } catch (error) {
    file.setTrashed(true); // 정보가 기록되지 않은 사진을 성공으로 남기지 않음.
    throw new Error('입력 정보 저장에 실패했습니다. 다시 시도해 주세요.');
  }
  // 공유 권한을 변경하지 않습니다. 지정 폴더의 접근 정책에 따라 보관합니다.
  return {name: file.getName()};
}
