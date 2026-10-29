# 수도권 본부 본사집회신고 공유 대시보드

`index.html`이 `data.json`을 읽는 정적 웹사이트입니다. 첨부 엑셀의 9개 열과 4개 본부를 반영했습니다. 엑셀의 실제 신고 내역은 비어 있으므로 초기 목록도 비어 있습니다. 기존 Streamlit 앱과 엑셀은 보존했습니다.

## 내용 수정

GitHub에서 **data.json → 연필(Edit) → 수정 → Commit changes**로 저장합니다. JSON은 큰따옴표를 쓰고 마지막 항목 뒤에는 쉼표를 넣지 않습니다. `reports` 배열에 다음 형태의 객체를 추가하세요. 아래 값은 형식 설명용 예시이며 실제 신고가 아닙니다.

```json
{
  "no": 1,
  "headquarters": "경기본부",
  "receiptNumber": "접수번호 입력",
  "company": "건설사명 입력",
  "location": "개최장소 및 행진진로 입력",
  "period": "집회 시작일 ~ 종료일 입력",
  "organizer": "개최기관 입력",
  "reportedAt": "신고일자 입력",
  "reporter": "신고인 입력"
}
```

`headquarters`는 최상위 `headquarters` 목록 중 하나여야 합니다. 접수번호는 앞자리 0을 보존하도록 문자열로 입력하세요. 날짜와 기간은 입력한 문자열 그대로 표시합니다. `updatedAt`은 자료 기준일(예: `2026-10-29`)이며 수정 시 직접 갱신하세요. `title`, `description`, 본부 목록도 JSON에서 바꿀 수 있습니다.

검색, 본부 필터, 본부별 건수, CSV 내려받기, 인쇄를 지원합니다. 웹 화면에서 직접 편집하거나 GitHub에 저장하는 기능은 없습니다. GitHub의 JSON 파일을 수정하면 배포가 실행되고 웹페이지에 반영됩니다.

## GitHub Pages 배포

저장소 **Settings → Pages → Build and deployment → Source → GitHub Actions**를 선택합니다. `main`에 HTML 또는 JSON 변경을 올리면 `.github/workflows/pages.yml`이 자동 배포합니다. 최초 설정 후 Actions에서 워크플로를 수동 실행할 수도 있습니다. 실제 사이트 주소는 배포 성공 후 Pages 설정 또는 Actions 결과에서 확인하세요. Pages 활성화 및 사용 가능 여부는 저장소 권한·요금제에 따라 달라집니다.

배포 파일에는 `index.html`, `data.json`만 포함됩니다. Pages 게시 시 신고인·장소 등 JSON에 입력한 정보도 사이트 방문자가 열람할 수 있으므로 공유 범위에 맞는 자료만 입력하세요.

## 로컬 확인

```sh
cd /workspace/chunsung
python3 -m http.server 8080 --bind 127.0.0.1
```

브라우저로 서버에서 확인하세요. HTML을 파일로 직접 열면 브라우저의 보안 정책으로 JSON을 읽지 못할 수 있습니다. 별도 빌드나 패키지 설치는 필요 없습니다.
