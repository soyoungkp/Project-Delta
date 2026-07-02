# 배포 가이드 — Streamlit Community Cloud (포트폴리오용)

목표: 이 BQ 비교 도구를 인터넷 공개 URL(`https://이름.streamlit.app`)로 올려 포트폴리오에 링크한다.

## 현재 상태 (확인 완료)
- ✅ `app.py` 는 이미 `st.file_uploader` 로 파일을 받는다 → **클라우드 배포에 코드 수정 불필요.**
- ✅ `requirements.txt` 추가됨 (streamlit / pandas / numpy / openpyxl).
- ✅ `.gitignore` 에 `.claude/` 추가 (로컬 설정 미커밋), 회사 xlsx(`sample/*.xlsx`)는 계속 제외.
- ✅ `python -m py_compile app.py sample/*.py` 통과.
- 브랜치: `master`, 원격(remote) 아직 없음.

## 1. GitHub 올리기

개인 포트폴리오라면 커밋 author를 개인 계정으로 먼저 설정 권장:
```bash
git config user.name "본인이름"
git config user.email "개인이메일@example.com"
```

커밋 & push (GitHub에서 Public repo 생성 후):
```bash
git add requirements.txt .gitignore DEPLOY.md
git commit -m "Add requirements.txt and prepare Streamlit Cloud deployment"

git remote add origin https://github.com/<내아이디>/<repo이름>.git
git push -u origin master
```
> 포트폴리오는 **Public** 저장소가 코드까지 보여줄 수 있어 유리. 회사 실제 BQ xlsx는 .gitignore로 이미 제외되므로 올라가지 않는다(확인: `git status` 에 `sample/*.xlsx` 안 보임).

## 2. Streamlit Community Cloud 배포
1. https://share.streamlit.io → **GitHub 계정으로 로그인**.
2. **Create app → Deploy a public app from GitHub**.
3. 입력: Repository=`<내아이디>/<repo이름>`, Branch=`master`, Main file path=`app.py`.
4. App URL 서브도메인 지정(예: `bq-compare`) → **Deploy**.
5. 1~3분 후 `https://bq-compare.streamlit.app` 형태 주소 완성 → 포트폴리오에 링크.

## 3. 무료 티어 참고
- 미접속 시 잠자기 → 방문하면 자동 기동(수십 초). 포트폴리오엔 무방.
- 코드 push 시 자동 재배포.
- 메모리 ~1GB. BQ 엑셀 비교엔 충분.

## (선택) 포트폴리오 데모 데이터
방문자가 파일 없이도 도구를 체험하게 하려면, 익명화한 데모 BQ 2개를 만들어
`sample/demo_Previous_BQ.xlsx`, `sample/demo_Revised_BQ.xlsx` 로 저장하고
`.gitignore` 에 예외(`!sample/demo_*.xlsx`)를 추가한 뒤, `app.py` 업로더 위에
"샘플로 체험하기" 버튼을 두면 된다. (원하면 실제 파일을 구조 그대로 익명화해 만들어 줄 수 있음.)
