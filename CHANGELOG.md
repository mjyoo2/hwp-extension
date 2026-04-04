# Changelog

All notable changes to "HWPX Editor" extension will be documented in this file.

## [0.4.0] - 2026-04-04

### Added
- HWPX to HWP 변환 기능 (MCP save_document에서 .hwp 확장자로 저장)
- HWP to HWPX 변환 기능 (MCP save_document에서 .hwpx 확장자로 저장)
- 공유 파서/라이터 모듈 (shared/) 도입, Extension과 MCP 서버 간 코드 재사용
- HWP 바이너리 Writer (OLE Compound File 생성)
- 라운드트립 변환 테스트 스위트 (16개 파일, 14개 완벽 통과)

### Fixed
- 테이블 셀 보더 변환 시 완전 손실 문제 해결 (99.99% 보존)
- 테이블 셀 배경색 변환 시 손실 문제 해결
- 헤더/푸터 변환 시 손실 문제 해결
- lineSpacing, fontColor, alignment 등 스타일 변환 오류 수정
- Merged cell (colSpan/rowSpan) 처리 대폭 개선
- 셀 내 중첩 테이블/이미지 변환 지원
- .hwp 파일이 실제 HWPX(ZIP)인 경우 매직 바이트 감지로 자동 처리
- Webview에서 style="none" 보더 렌더링 오류 수정
- MCP 서버 TypeScript 빌드 에러 수정

### Improved
- HWP 바이너리 파서 대폭 강화 (보더, 헤더/푸터, 페이지 설정 파싱 개선)
- Underline을 객체 형식으로 보존 (type, shape, color)
- HWPX 파서 셀 width 계산 개선 (NaN 방지)
- MCP save_document에 OLE 시그니처 검증 및 포맷 정보 추가

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
