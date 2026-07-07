import type { Tool } from "./types";

// 이미지 위치: public/tools/kocarc-automation/{logo.png, card.png, guide/*.png}
// TODO 확인 필요: downloadUrl 릴리스 자산명, changelog 날짜, featured 여부
export const kocarcAutomation: Tool = {
  slug: "kocarc-automation",
  name: "KOCARC 자동등록",
  tagline:
    "엑셀에 환자 정보를 채워 넣으면 KOCARC eCRF(ecrf.kr/kocarc)에 심정지 환자를 한 명씩 자동으로 로그인·입력·저장까지 일괄 등록합니다.",
  category: "desktop",
  status: "active",
  icon: "ClipboardPlus",
  logo: "/tools/kocarc-automation/logo.png",
  screenshot: "/tools/kocarc-automation/card.png",
  featured: false, // TODO: 대표 도구로 노출할지 결정
  hasGuide: false,
  downloadUrl:
    // TODO: 릴리스에 올린 실제 exe 자산명으로 교체 (파일명에 공백/한글이 있으면 URL 인코딩 필요)
    "https://github.com/yujeong0411/kocarc_automation/releases/latest/download/KOCARC_%EC%9E%90%EB%8F%99%EB%93%B1%EB%A1%9D.exe",
  quickStart: [
    "실행파일을 다운로드해서 더블클릭하세요.",
    "PC 보호창이 뜨면 추가정보를 누르고 실행을 누르세요.",
    "[빈 양식 만들기]로 엑셀 양식을 만들고, 회색·빨강 안내를 보며 환자 정보를 채우세요.",
    "채운 엑셀 파일을 선택하고 eCRF 아이디·비밀번호를 입력하세요.",
    "[시작]을 누르면 실시간 로그를 보며 자동 등록됩니다.",
  ],
  features: [
    {
      icon: "FileSpreadsheet",
      title: "검증 내장 엑셀 양식",
      desc: "드롭다운·입력 규칙이 들어간 빈 양식을 버튼 한 번으로 생성",
    },
    {
      icon: "Palette",
      title: "색으로 보는 입력 안내",
      desc: "회색은 해당 없는 칸, 빨강은 규칙에 어긋난 값 — 채우기 전 미리 확인",
    },
    {
      icon: "ShieldCheck",
      title: "사이트 저장검증 재사용",
      desc: "eCRF 자체 검사(checkInput)를 통과한 환자만 실제로 저장",
    },
    {
      icon: "History",
      title: "중복 방지 · 이어하기",
      desc: "진행 상황을 기록해 중단·재시작해도 이미 등록한 환자는 건너뜀",
    },
    {
      icon: "ScrollText",
      title: "실시간 로그",
      desc: "등록 진행 상황을 창에서 실시간으로 확인하고 언제든 중지",
    },
    {
      icon: "RotateCcw",
      title: "처음부터 새로 시작",
      desc: "체크박스 하나로 이전 진행 기록을 지우고 1번 환자부터 다시",
    },
  ],
  requirements: [
    "Windows 10 이상",
    "Google Chrome 설치",
    "KOCARC eCRF(ecrf.kr/kocarc) 계정 (아이디·비밀번호)",
    "인터넷 연결 — 실제 연구 DB에 생성·저장됩니다",
  ],
  blogPosts: [],
  changelog: [
    {
      version: "1.0.0",
      date: "2026.07.07", // TODO: 실제 릴리스 날짜로 교체
      note: "최초 정식 릴리스. 검증 내장 엑셀 양식 생성, 색 기반 입력 안내(회색·빨강), 사이트 저장검증(checkInput) 재사용, 중복 생성 방지·이어하기, 실시간 로그, 처음부터 새로 시작 옵션",
    },
  ],
};
