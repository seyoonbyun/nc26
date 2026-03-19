# Accelerate 2026 — 구글 시트 연동 설정 가이드

## 전체 동작 흐름
```
1. 국외 유저 → pricing.html에서 플랜 선택 → 모달 폼 작성 → 제출
2. Apps Script → Google Sheets에 데이터 저장 (Status: Pending)
3. 어드민 → 구글 시트에서 등록 목록 확인 + 결제 사이트에서 수기 확인
4. 어드민 → 결제 확인된 유저의 Status 셀에 "결제완료" 입력
5. Apps Script → onEdit 트리거가 감지 → 해당 유저에게 확인 이메일 자동 발송
6. 유저 → 티켓번호 + 플랜정보 + 행사안내가 포함된 컨펌레터 이메일 수신
   (이메일은 유저가 선택한 언어로 발송: en/ja/zh/ko)
```

## 설정 순서

### 1단계: Google Sheets 생성
- [Google Sheets](https://sheets.google.com)에서 새 스프레드시트 생성

### 2단계: Apps Script 코드 붙여넣기
1. 스프레드시트에서 **확장 프로그램 > Apps Script**
2. 기본 Code.gs 내용을 이 프로젝트의 `Code.gs` 내용으로 전체 교체
3. 저장 (Ctrl+S)

### 3단계: 웹 앱 배포 (폼 데이터 수신용)
1. **배포 > 새 배포**
2. 유형: **웹 앱**
3. 설정:
   - 실행 사용자: **나**
   - 액세스 권한: **모든 사용자**
4. **배포** → URL 복사

### 4단계: onEdit 트리거 설정 (결제완료 감지용)
> ⚠️ 중요: 이 설정을 해야 "결제완료" 입력 시 이메일이 자동 발송됩니다.

1. Apps Script 편집기 왼쪽 메뉴에서 **⏰ 트리거** 클릭
2. 우측 하단 **+ 트리거 추가** 클릭
3. 설정:
   - 실행할 함수: **onSheetEdit**
   - 이벤트 소스: **스프레드시트에서**
   - 이벤트 유형: **수정 시**
4. **저장** → Google 계정 권한 승인

### 5단계: pricing.html에 URL 연결
브라우저에서 pricing.html을 열고, 콘솔(F12)에서 실행:
```js
localStorage.setItem('accelerate_script_url', '여기에_배포URL');
```

## 어드민 사용법

### 구글 시트 컬럼 구조
| A | B | C | D | E | F | G | H | I | J |
|---|---|---|---|---|---|---|---|---|---|
| Timestamp | Name | Nationality | Email | Phone | Plan | Price | Language | Memo | Status |

### 결제 처리 방법
1. 시트에서 새 등록 확인 (Status: Pending)
2. 외부 결제 사이트에서 해당 유저의 결제 내역 수기 확인
3. J열(Status)에 **`결제완료`** 입력 후 Enter
4. 자동으로 이메일 발송 → Status가 **`결제완료 ✓ 메일발송`** 으로 변경됨
5. 발송 실패 시 **`결제완료 ✗ 발송실패: 에러메시지`** 표시

## 참고사항
- "결제완료" 정확히 입력해야 트리거가 동작합니다 (공백, 대소문자 주의)
- Gmail 일일 발송 한도: 무료 계정 100건/일, Workspace 1,500건/일
- Apps Script 코드 수정 시 **새 배포**를 해야 웹 앱(doPost)에 반영됩니다
- onEdit 트리거는 코드 수정만으로 자동 반영됩니다 (재배포 불필요)
