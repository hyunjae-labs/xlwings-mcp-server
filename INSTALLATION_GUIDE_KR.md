# xlwings MCP Server 설치 가이드

## 사전 준비
1. **Python 설치** → [python.org](https://www.python.org/downloads/) 에서 다운로드
   - **검증된 버전: Python 3.13.5** (이 가이드 작성 및 테스트 완료 버전)
   - 다른 버전이 이미 설치되어 있다면 재설치 불필요 - 기존 버전으로 진행
   - 설치할 때 **"Add Python to PATH"** 체크 필수
2. **VSCode + Continue 확장** 설치되어 있어야 함

---

## 설치 개요

**xlwings-mcp-server**는 Python 패키지(라이브러리)입니다.
- Excel을 제어하는 코드가 들어있음
- `pip install`로 설치하면 Python 내부에 저장됨
- 설치 후 시스템 전역에서 `python -m xlwings_mcp`로 실행 가능
- 특정 폴더나 위치에 제약 없음

---

## 방법 1: 표준 설치 (권장)

### 1단계: 명령 프롬프트 열기
- Windows 키 + R → `cmd` 입력 → Enter

### 2단계: 패키지 설치
```cmd
pip install xlwings-mcp-server
```
- Enter 누르고 설치 완료 대기 (1-2분)

### 3단계: Python 경로 확인
```cmd
where python
```
출력 예시들:
- `C:\Python313\python.exe`
- `C:\Users\[사용자명]\AppData\Local\Programs\Python\Python313\python.exe`
- `C:\Users\[사용자명]\miniconda3\python.exe`

**주의: 실제 출력된 경로를 복사하세요**

### 4단계: Continue 설정 파일 만들기
1. 파일 탐색기에서 주소창에 입력: `%USERPROFILE%\.continue`
2. `mcpServers` 폴더 만들기 (없으면)
3. 그 안에 `xlwings-mcp-server.yaml` 텍스트 파일 생성
4. 아래 내용 붙여넣기:

```yaml
name: xlwings-mcp-server
version: 0.1.7
description: Excel MCP Server
schema: v1
mcpServers:
  - name: xlwings-mcp-server
    command: "여기에_3단계_경로_붙여넣기"
    args: ["-m", "xlwings_mcp", "stdio"]
```

**실제 예시** (3단계 결과가 `C:\Python313\python.exe`인 경우):
```yaml
    command: "C:/Python313/python.exe"
```

**설명**: 
- `command:` = Python 실행 파일 경로 (슬래시 `/` 사용)
- `args:` = Python에게 xlwings_mcp 모듈을 실행하라는 명령

### 5단계: VSCode 재시작

---

## 작동 원리

`pip install xlwings-mcp-server` 실행 시:
1. PyPI 저장소에서 패키지 다운로드
2. Python의 `site-packages` 디렉토리에 설치
3. 설치 완료 후 시스템 전역에서 `python -m xlwings_mcp` 명령으로 실행 가능

---

## 방법 2: 가상환경 설치 (선택사항)

가상환경을 만들어서 다른 Python 프로그램과 충돌 방지

### PowerShell에서 실행:
```powershell
# 홈 폴더에 전용 폴더 생성
cd $env:USERPROFILE
mkdir xlwings-mcp-server
cd xlwings-mcp-server

# 가상환경 생성 및 패키지 설치
python -m venv .venv
.\.venv\Scripts\activate
pip install xlwings-mcp-server

# Python 경로 확인
where python
# 결과: C:\Users\[사용자명]\xlwings-mcp-server\.venv\Scripts\python.exe
```

이후 Continue 설정은 방법 1의 4-5단계와 동일 (경로만 다름)

---

## 수동 설치 (대체 방법)

### 1. Python 설치
1. [Python 3.12](https://www.python.org/downloads/) 다운로드
2. 설치 시 **"Add Python to PATH"** 체크 ✅
3. 설치 완료

### 2. 명령 프롬프트에서 실행
```cmd
cd %USERPROFILE%
mkdir xlwings-mcp-server
cd xlwings-mcp-server
python -m venv .venv
.venv\Scripts\activate
pip install xlwings-mcp-server
where python
```
마지막 명령어 결과 복사 (예: `C:\Users\[사용자명]\xlwings-mcp-server\.venv\Scripts\python.exe`)

### 3. Continue 설정 파일 생성
1. 폴더 열기: `%USERPROFILE%\.continue\mcpServers`
2. `xlwings-mcp-server.yaml` 파일 생성
3. 내용 입력 (경로 수정 필요):

```yaml
name: xlwings-mcp-server
version: 0.1.7
description: Excel MCP Server
schema: v1
mcpServers:
  - name: xlwings-mcp-server
    command: "C:/Users/[사용자명]/xlwings-mcp-server/.venv/Scripts/python.exe"
    args: ["-m", "xlwings_mcp", "stdio"]
```
⚠️ 경로의 `\`를 `/`로 변경

### 4. VSCode 재시작

---

## 문제 해결

**Continue에서 서버가 안 보임**
- YAML 파일 경로 확인: `%USERPROFILE%\.continue\mcpServers\xlwings-mcp-server.yaml`
- 경로 슬래시 방향 확인 (`/` 사용)
- VSCode 완전 재시작

**Python 오류**
- PowerShell 관리자 권한 실행 확인
- 컴퓨터 재시작 후 재시도