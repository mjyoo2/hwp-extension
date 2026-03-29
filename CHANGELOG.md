# Changelog

All notable changes to "HWPX Editor" extension will be documented in this file.

## [0.3.0] - 2026-03-29

### Fixed
- 테이블 배경색 및 테두리선 렌더링 오류 수정
- 이미지 레이아웃 오류 수정

### Improved
- HWP 바이너리 파서 대폭 강화 (standalone parser 개선)
- MCP 서버에 HWP 바이너리 파일 읽기 지원 추가
- MCP 서버 문서 조작 도구 확장 (HwpxDocument, HwpxParser)
- Webview UI 렌더링 개선 (테이블, 이미지)
- 미사용 코드 및 파일 정리

## [0.2.0] - 2025-02-23

### Added
 HWP 바이너리 파일 읽기 지원 (읽기 전용)
 OLE Compound File 자체 파서 구현
 HWP 전용 Custom Editor (HWP Viewer)
 HWP → HWPX 변환 커맨드 추가

### Improved
 HWPX 파서 대폭 강화 (더 많은 HWPML 요소 지원)
 Webview UI 개선 및 확장
 MCP 서버 도구 및 문서 조작 기능 확장

## [0.1.0] - 2025-01-12

### Added
- HWPX 파일 읽기/쓰기 지원
- 텍스트 편집 기능
- 테이블 보기 및 편집
- 문서 메타데이터 확인
- MCP (Model Context Protocol) 서버 포함
- AI 도구 연동 지원 (Claude 등)
