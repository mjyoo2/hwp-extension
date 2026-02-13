# HWP/HWPX Editor for VSCode

[![GitHub](https://img.shields.io/badge/GitHub-mjyoo2%2Fhwp--extension-blue?logo=github)](https://github.com/mjyoo2/hwp-extension)

VSCode에서 한글(HWP/HWPX) 문서를 열고 편집할 수 있는 확장 프로그램입니다.

## Features

### 문서 보기 및 편집
- **HWPX 파일**: 읽기 및 편집 지원 (XML 기반 최신 포맷)
- **HWP 파일**: 읽기 지원 (바이너리 레거시 포맷, 읽기 전용)

### 지원 기능
- 텍스트 내용 보기 및 편집
- 단락 추가/수정/삭제
- 테이블 보기 및 편집
- 문서 메타데이터 확인
- 문서 구조 탐색

### MCP (Model Context Protocol) 서버
AI 도구(Claude 등)와 연동하여 문서를 자동으로 편집할 수 있는 MCP 서버를 포함합니다.

## Installation

1. VSCode에서 Extensions (Ctrl+Shift+X) 열기
2. "HWPX Editor" 검색
3. Install 클릭

## Usage

### 기본 사용법
1. HWPX 파일을 VSCode에서 열기
2. 자동으로 HWPX Editor가 활성화됨
3. 문서 내용 확인 및 편집

### MCP 서버 사용 (AI 연동)

Command Palette (Ctrl+Shift+P)에서:
- `HWPX: Show MCP Server Configuration` - MCP 설정 정보 확인
- `HWPX: Copy MCP Server Path` - MCP 서버 경로 복사

#### Claude Code에서 사용

`.vscode/mcp.json` 파일 생성:

```json
{
  "mcpServers": {
    "hwpx": {
      "command": "node",
      "args": ["${extensionPath}/out/mcp-server.js"]
    }
  }
}
```

## Supported File Formats

| 포맷 | 확장자 | 읽기 | 쓰기 |
|------|--------|------|------|
| HWPX | .hwpx | O | O |
| HWP | .hwp | O | X (읽기 전용) |

> **Note**: HWP 파일은 읽기 전용으로 지원됩니다. 편집이 필요한 경우 한컴오피스에서 HWPX로 변환 후 사용해주세요.

## Requirements

- VSCode 1.107.0 이상

## Known Issues

- HWP 파일은 읽기 전용 (편집하려면 HWPX로 변환 필요)
- 일부 복잡한 서식은 표시되지 않을 수 있습니다
- 암호화된 HWP 파일은 지원하지 않습니다

## Release Notes

### 0.2.0
- HWP 파일 읽기 지원 추가 (읽기 전용)
- OLE Compound File 파싱 구현

### 0.1.0
- 최초 릴리스
- HWPX 파일 읽기/쓰기 지원
- MCP 서버 포함

## License

MIT

## Contributing

GitHub: https://github.com/mjyoo2/hwp-extension

버그 리포트 및 기능 요청은 GitHub Issues를 이용해주세요.
