/**
 * HWPX 인코더 관련 상수
 *
 * HWPX 파일 생성에 필요한 모든 설정값을 한곳에 모아 관리합니다.
 * 초중급 개발자도 쉽게 수정할 수 있도록 명확한 주석을 포함했습니다.
 */

// ==================== 파일 형식 관련 ====================

/** HWPX 파일의 MIME 타입 (ZIP 파일 내 mimetype 파일에 저장됨) */
export const HWPX_MIME_TYPE = "application/vnd.hancom.hwpx";

/** HWPX 파일의 버전 */
export const HWPX_VERSION = "1.2";

// ==================== XML 네임스페이스 ====================

/** HWPX XML 문서에서 사용되는 네임스페이스 */
export const NAMESPACES = {
  /** Hancom 문서 네임스페이스 */
  HANCOM: "http://www.hancom.co.kr/hwp/xml",
  /** Hancom 공통 네임스페이스 */
  HANCOM_COMMON: "http://www.hancom.co.kr/hwp/xml/common",
  /** Hancom 버전 네임스페이스 */
  HANCOM_VERSION: "http://www.hancom.co.kr/hwp/xml/version",
  /** Hancom 속성 네임스페이스 */
  HANCOM_PROP: "http://www.hancom.co.kr/hwp/xml/property",
} as const;

/** XML 헤더에 포함될 네임스페이스 선언 문자열 */
export const NAMESPACE_DECLARATIONS = {
  HEAD: `xmlns:hh="${NAMESPACES.HANCOM}" xmlns:hc="${NAMESPACES.HANCOM_COMMON}" xmlns:hv="${NAMESPACES.HANCOM_VERSION}" xmlns:hp="${NAMESPACES.HANCOM_PROP}"`,
  SECTION: `xmlns:hs="${NAMESPACES.HANCOM}" xmlns:hp="${NAMESPACES.HANCOM_PROP}"`,
} as const;

// ==================== 단위 변환 ====================

/** 1 포인트 (pt) 를 HWPUNIT 으로 변환한 값 (HWPX 내부 단위) */
export const HWPUNIT_PER_PT = 1000;

/** 1 인치 (inch) 를 포인트 (pt) 로 변환한 값 */
export const PT_PER_INCH = 72;

/** 1 인치 (inch) 를 픽셀로 변환한 값 (표준 96 DPI 기준) */
export const PIXELS_PER_INCH = 96;

/** 픽셀을 포인트로 변환하는 계수 */
export const PT_PER_PIXEL = PT_PER_INCH / PIXELS_PER_INCH; // 0.75

// ==================== 페이지 크기 (A4 기준) ====================

/** A4 용지 너비 (포인트) */
export const A4_WIDTH_PT = 794;

/** A4 용지 높이 (포인트) */
export const A4_HEIGHT_PT = 1123;

/** A4 용지 상단 여백 (포인트) */
export const A4_TOP_MARGIN_PT = 72;

/** A4 용지 하단 여백 (포인트) */
export const A4_BOTTOM_MARGIN_PT = 72;

/** A4 용지 왼쪽 여백 (포인트) */
export const A4_LEFT_MARGIN_PT = 83;

/** A4 용지 오른쪽 여백 (포인트) */
export const A4_RIGHT_MARGIN_PT = 83;

/** A4 용지 상단 머리글 영역 (포인트) */
export const A4_HEADER_MARGIN_PT = 40;

/** A4 용지 하단 바닥글 영역 (포인트) */
export const A4_FOOTER_MARGIN_PT = 40;

// ==================== 폰트 관련 ====================

/** 기본 한글 폰트 */
export const DEFAULT_KOREAN_FONT = "Batang";

/** 기본 영문 폰트 */
export const DEFAULT_LATIN_FONT = "Arial";

/** 기본 폰트 크기 (포인트) */
export const DEFAULT_FONT_SIZE_PT = 10;

// ==================== 줄 간격 ====================

/** 기본 줄 간격 (폰트 크기의 %) */
export const DEFAULT_LINE_SPACING_PERCENT = 160;

/** 최소 줄 간격 (HWPUNIT) */
export const MIN_LINE_SPACING_HWPUNIT = 1000;

// ==================== 스타일 ID ====================

/** 바탕글 스타일 ID */
export const STYLE_ID_BATANGGUL = "0";

/** 개요 1 스타일 ID */
export const STYLE_ID_GYEYOE1 = "1";

/** 개요 2 스타일 ID */
export const STYLE_ID_GYEYOE2 = "2";

// ==================== 이미지 관련 ====================

/** 이미지 파일이 저장될 디렉토리 경로 (ZIP 내) */
export const IMAGE_DIR_PATH = "image/";

/** 이미지 파일명 접미사 */
export const IMAGE_FILE_EXTENSION = ".png";

// ==================== 테이블 관련 ====================

/** 테이블 테두리 기본 두께 (HWPUNIT) */
export const TABLE_BORDER_WIDTH_HWPUNIT = 500; // 0.5pt

/** 테이블 셀 기본 패딩 (HWPUNIT) */
export const TABLE_CELL_PADDING_HWPUNIT = 200; // 0.2pt

// ==================== 색상 코드 ====================

/** 검정색 (hex) */
export const COLOR_BLACK = "#000000";

/** 흰색 (hex) */
export const COLOR_WHITE = "#FFFFFF";

/** 회색 (워터마크용) */
export const COLOR_GRAY = "#CCCCCC";

/** 코드 블록 배경색 */
export const COLOR_CODE_BLOCK_BG = "#f4f4f4";

// ==================== ZIP 파일 구조 ====================

/** HWPX 파일 내 파일 목록 */
export const HWPX_FILE_STRUCTURE = [
  { name: "mimetype", mimeType: "" },
  { name: ".hwpversions/version.xml", mimeType: "application/xml" },
  { name: "header.xml", mimeType: "application/xml" },
  { name: "section0.xml", mimeType: "application/xml" },
  { name: "content.hpf", mimeType: "application/octet-stream" },
] as const;

/** ZIP 파일 생성 시 사용할 압축 레벨 (0=압축안함, 9=최대압축) */
export const ZIP_COMPRESSION_LEVEL = 9;
