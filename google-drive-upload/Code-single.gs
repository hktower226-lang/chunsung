/** 웹 앱 배포: 실행 사용자 ‘나’, 액세스 ‘모든 사용자’으로 설정하세요. */
const FOLDER_ID = '1_VhjNfTPKWy-v_WsY-fZ6jT7VAaqN4Cg';
const MAX_BYTES = 10 * 1024 * 1024;

function doGet() {
  return HtmlService.createHtmlOutput(UPLOAD_PAGE)
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

// 화면과 저장 코드를 함께 배포하여 HTML 파일 혼동을 방지합니다.
const UPLOAD_PAGE = `<!doctype html>
<html lang="ko"><head><meta charset="utf-8"><base target="_top"><style>
*{box-sizing:border-box}body{font-family:system-ui,sans-serif;background:#f3f6fb;color:#243b59;margin:0;padding:24px 16px}main{max-width:480px;margin:24px auto;background:white;border:1px solid #e0e7f1;border-radius:16px;padding:24px}h1{font-size:22px;margin:0 0 12px}p{font-size:14px;line-height:1.8;color:#66778f}label{display:block;margin:20px 0 8px;font-size:14px;font-weight:650}input,select,button{font:inherit;width:100%;padding:12px;border-radius:8px;border:1px solid #ccd7e7;min-height:46px}input[type=file]{font-size:14px}button{margin-top:24px;background:#315fc6;border:0;color:#fff;font-weight:700;cursor:pointer}button:disabled{opacity:.55;cursor:wait}#status{white-space:pre-line;margin:18px 0 0;overflow-wrap:anywhere}a{display:block;margin-top:12px;color:#315fc6}input:focus,select:focus,button:focus-visible{outline:2px solid #527fe6;outline-offset:2px}
</style></head><body><main><h1>집회신고 접수증 등록 · 새 화면</h1><p>경기본부(수도권통합본부) Drive 폴더에 저장합니다. 사진과 입력 정보를 함께 보관하며 대시보드 신고 내역은 자동 등록되지 않습니다.</p>
<form id="uploadForm"><label for="headquarters">본부</label><select id="headquarters" name="headquarters" required><option value="">본부를 선택하세요</option><option>경기본부</option><option>인천본부</option><option>서울본부</option><option>대세충본부</option></select><label for="company">건설사명</label><input id="company" name="company" maxlength="60" placeholder="예: 디엘이앤씨(주)" required><label for="officerContact">정보관 이름 및 연락처</label><input id="officerContact" name="officerContact" maxlength="120" placeholder="예: 박재웅 / 010-3583-0914" required><label for="reporter">신고인 이름</label><input id="reporter" name="reporter" maxlength="60" placeholder="예: 남형욱" required><label for="receipt">접수증 사진</label><input id="receipt" name="receipt" type="file" accept="image/jpeg,image/png,image/webp" required><p>JPG · PNG · WEBP / 사진 한 장당 최대 10MB<br>저장 후 다른 사진을 선택해 계속 올릴 수 있습니다.</p><button id="submit" type="submit">Google Drive에 저장</button></form><p id="status" role="status" aria-live="polite"></p></main>
<script>
const form=document.getElementById('uploadForm'),button=document.getElementById('submit'),status=document.getElementById('status');
form.addEventListener('submit',event=>{
 event.preventDefault();if(button.disabled||!form.reportValidity())return;const file=form.receipt.files[0];
 if(!file||!file.size||file.size>10*1024*1024){status.textContent='10MB 이하의 사진을 선택해 주세요.';return}
 button.disabled=true;status.textContent='저장 중입니다. 창을 닫지 마세요.';
 google.script.run.withSuccessHandler(result=>{button.disabled=false;status.textContent='접수가 완료되었습니다.\\n'+result.name+'\\n사진과 입력 정보가 담당자 Drive에 저장되었습니다.';form.reset();}).withFailureHandler(error=>{button.disabled=false;status.textContent='저장 실패: '+(error.message||'다시 시도해 주세요.');}).uploadReceipt(form);
});
</script></body></html>
`;
