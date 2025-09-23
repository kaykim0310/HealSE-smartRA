# 🚀 GitHub 업로드 가이드

## 📦 완성된 파일

**KOSHA-KRAS-System.zip** 파일에 모든 것이 포함되어 있습니다!

## 🔗 GitHub에 업로드하는 방법

### 방법 1: 웹 인터페이스 사용 (가장 쉬움)

1. **GitHub 접속**: https://github.com
2. **새 저장소 생성**: "New repository" 클릭
3. **저장소 이름**: `kosha-kras-system`
4. **설명**: `KOSHA KRAS 위험성평가 자동 분석 시스템`
5. **Public** 선택 (또는 Private)
6. **Create repository** 클릭

7. **파일 업로드**:
   - "uploading an existing file" 링크 클릭
   - ZIP 파일을 압축 해제한 후 모든 파일을 드래그 앤 드롭
   - 또는 "choose your files" 클릭하여 파일 선택

8. **커밋 메시지**: `Initial commit: KOSHA KRAS Risk Assessment System`
9. **Commit new files** 클릭

### 방법 2: Git 명령어 사용

```bash
# ZIP 파일 압축 해제
unzip KOSHA-KRAS-System.zip
cd kosha-kras-system

# Git 초기화
git init
git add .
git commit -m "Initial commit: KOSHA KRAS Risk Assessment System"

# GitHub 저장소와 연결 (저장소 URL을 본인 것으로 변경)
git remote add origin https://github.com/YOUR-USERNAME/kosha-kras-system.git
git branch -M main
git push -u origin main
```

## 📋 포함된 파일들

- ✅ **app.py**: 메인 애플리케이션
- ✅ **start.bat**: Windows 실행 스크립트
- ✅ **start.sh**: Linux/Mac 실행 스크립트
- ✅ **requirements.txt**: Python 의존성
- ✅ **Dockerfile**: Docker 지원
- ✅ **docker-compose.yml**: Docker Compose 설정
- ✅ **README.md**: 완전한 프로젝트 문서
- ✅ **.gitignore**: Git 무시 파일
- ✅ **uploads/**, **outputs/**: 필요한 디렉토리

## 🎯 추천 저장소 설정

### 저장소 이름
- `kosha-kras-system`
- `workplace-risk-assessment`
- `kosha-safety-analyzer`

### 태그 (Topics)
- `kosha`
- `risk-assessment`
- `workplace-safety`
- `python`
- `flask`
- `excel`
- `ai-analysis`

### 라이선스
- MIT License (권장)

## 🌟 완료 후 할 일

1. **README 확인**: 저장소에서 README가 제대로 표시되는지 확인
2. **Issues 활성화**: 사용자 피드백을 받기 위해
3. **Releases 생성**: v1.0.0 태그로 첫 번째 릴리스 생성
4. **별표 요청**: 사용자들에게 ⭐ 표시 요청

## 🎉 성공!

GitHub 저장소가 생성되면 다음과 같은 URL을 얻게 됩니다:
`https://github.com/YOUR-USERNAME/kosha-kras-system`

이제 누구나 이 링크로 접속해서 프로젝트를 다운로드하고 사용할 수 있습니다! 🚀
