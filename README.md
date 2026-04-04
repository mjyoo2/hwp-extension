# HWP/HWPX Editor for VSCode

[![GitHub](https://img.shields.io/badge/GitHub-mjyoo2%2Fhwp--extension-blue?logo=github)](https://github.com/mjyoo2/hwp-extension)

VSCode에서 한글(HWP/HWPX) 문서를 열고 편집할 수 있는 확장 프로그램입니다.

## Features

### 문서 보기 및 편집
- **HWPX 파일**: 읽기 및 편집 지원 (XML 기반 최신 포맷)
- **HWP 파일**: 읽기 지원 (바이너리 레거시 포맷)

### 포맷 변환
- **HWPX to HWP**: HWPX 문서를 HWP 바이너리로 변환 (MCP 또는 Extension)
- **HWP to HWPX**: HWP 문서를 HWPX로 변환하여 편집 가능하게 전환

### 지원 기능
- 텍스트 내용 보기 및 편집
- 단락 추가/수정/삭제
- 테이블 보기 및 편집 (보더, 배경색, 병합 셀 지원)
- 이미지, 수학식, 각주/미주 보기
- 헤더/푸터 보기
- 문서 메타데이터 확인
- 문서 구조 탐색

### MCP (Model Context Protocol) 서버
AI 도구(Claude 등)와 연동하여 문서를 자동으로 편집할 수 있는 MCP 서버를 포함합니다.

주요 MCP 도구:
- `open_document` - HWP/HWPX 문서 열기
- `save_document` - 문서 저장 (.hwp 또는 .hwpx 확장자로 포맷 변환 가능)
- `get_document_text`, `get_paragraphs`, `get_tables` - 문서 내용 조회
- `search_text`, `replace_text` - 텍스트 검색/치환
- `insert_table`, `insert_image` - 요소 삽입
- 외 50+ 문서 조작 도구

## Installation

1. VSCode에서 Extensions (Ctrl+Shift+X) 열기
2. "HWPX Editor" 검색
3. Install 클릭

## Usage

### 기본 사용법
1. HWP 또는 HWPX 파일을 VSCode에서 열기
2. 자동으로 HWP/HWPX Editor가 활성화됨
3. 문서 내용 확인 및 편집

### 포맷 변환
- **HWP를 HWPX로**: HWP 파일을 열면 저장 시 자동으로 HWPX로 변환
- **MCP로 변환**: `save_document`에서 `output_path`의 확장자로 변환 포맷 지정

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

| 포맷 | 확장자 | 읽기 | 쓰기 | 변환 |
|------|--------|------|------|------|
| HWPX | .hwpx | O | O | HWPX to HWP |
| HWP | .hwp | O | O (변환) | HWP to HWPX |

## Requirements

- VSCode 1.107.0 이상

## Known Issues

- 일부 복잡한 서식은 표시되지 않을 수 있습니다
- 암호화된 HWP 파일은 지원하지 않습니다
- 수학식(equation) 요소는 변환 시 인라인으로 처리되어 element 구조가 달라질 수 있습니다

## Release Notes

### 0.4.0
- HWPX to HWP, HWP to HWPX 양방향 변환 기능 추가
- 공유 파서/라이터 모듈 (shared/) 도입
- 테이블 보더/배경색/헤더/스타일 변환 정확도 대폭 개선
- Merged cell (colSpan/rowSpan) 변환 지원
- MCP 서버 빌드 안정화

### 0.3.0
- 테이블 배경색 및 테두리선 렌더링 오류 수정
- HWP 바이너리 파서 강화

### 0.2.0
- HWP 파일 읽기 지원 추가 (읽기 전용)
- OLE Compound File 파싱 구현

### 0.1.0
- 최초 릴리스
- HWPX 파일 읽기/쓰기 지원
- MCP 서버 포함

## Disclaimer

본 프로젝트는 공개된 한글 문서 파일 형식 사양(HWPML, HWP 5.0)을 기반으로 독자적으로 구현한 비공식 도구입니다. 한글과컴퓨터(Hancom)와는 아무런 관련이 없으며, 상업적 목적이 아닌 개인/연구 용도로 개발되었습니다. "HWP", "HWPX"는 한글과컴퓨터의 파일 형식명입니다.

## License

MIT

## Contributing

GitHub: https://github.com/mjyoo2/hwp-extension

버그 리포트 및 기능 요청은 GitHub Issues를 이용해주세요.
