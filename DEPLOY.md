# 배포 가이드 (GitHub Pages 기준)

## 구조
- 기본 사용: `index.html` / `automatic.html` -> GitHub Pages
- 선택 서버: `automatic_server.py` -> Render/Railway 같은 서버

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

## 3) 사용 방법
- GitHub Pages URL만 열면 된다.
- 별도 API 서버 주소 입력은 필요 없다.
- 기준 파일/참고 파일 업로드 후 바로 생성 가능하다.

## 4) API 서버 배포 (선택 사항, Render 예시)
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

## 5) 현재 상태
- GitHub Pages 메인 URL 응답 확인 완료: `https://dduonthetop.github.io/balju-auto-generator/`
- Render는 선택 사항이며, Python 3.13 호환 수정까지 반영 완료했다.

## 주의
- 기본 배포는 GitHub Pages만 사용한다.
- `automatic_server.py`는 필요할 때만 별도 서버로 띄운다.
