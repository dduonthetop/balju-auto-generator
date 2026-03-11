# 배포 가이드 (GitHub Pages + API 서버)

## 구조
- 프론트: `automatic.html` -> GitHub Pages
- 백엔드: `automatic_server.py` -> Render/Railway 같은 서버

## 1) GitHub 저장소 생성 및 업로드
```powershell
cd C:\Users\ksuja\Desktop\balju
git init
git add .
git commit -m "Add automatic excel generator web + server"
git branch -M main
git remote add origin https://github.com/dduonthetop/balju-auto-generator.git
git push -u origin main
```

## 2) GitHub Pages 켜기
1. GitHub 저장소 -> `Settings` -> `Pages`
2. `Build and deployment`:
   - Source: `Deploy from a branch`
   - Branch: `main` / `/ (root)`
3. 저장 후 URL 확인:
   - 메인 페이지: `https://dduonthetop.github.io/balju-auto-generator/`
   - 직접 페이지: `https://dduonthetop.github.io/balju-auto-generator/automatic.html`

## 3) API 서버 배포 (Render 예시)
1. Render에서 `New +` -> `Web Service`
2. GitHub 저장소 연결
3. 설정:
   - Build Command: `pip install -r requirements.txt`
   - Start Command: `python automatic_server.py`
   - Environment:
     - `PORT` (Render가 자동 주입)
     - `HOST=0.0.0.0`
4. 배포 후 서버 URL 확인:
   - 예: `https://balju-api.onrender.com`

## 4) 프론트에서 API 주소 입력
- GitHub Pages URL 접속
- 화면의 `API 서버 주소`에 서버 URL 입력
  - 예: `https://balju-api.onrender.com`
- 기준파일/참고파일 업로드 후 생성

## 5) 현재 상태
- GitHub Pages 메인 URL 응답 확인 완료: `https://dduonthetop.github.io/balju-auto-generator/`
- Render는 저장소 연결과 첫 배포 승인에 계정 로그인이 필요하므로 브라우저에서 1회 진행해야 한다.

## 주의
- GitHub Pages는 정적 사이트만 가능해서 `apply_mapping.py` 직접 실행 불가.
- 반드시 API 서버가 따로 떠 있어야 업로드/생성/다운로드가 동작함.
