# 와이프의 업무 시간 단축 및 효율을 위해서 만든 유틸 모음 (feelsUtil)

불철주야 고생하시는 내무부장관님께 바치는 도구이옵니다.

## 현재 버전

v1.0.2

## 주요 기능 및 수정사항

*   **v1.0.2 (2025-07-06) 변경 사항:**
    *   `PasswordWindow` 실행 시 비밀번호 입력 필드에 자동으로 포커스가 가도록 수정.
    *   경기번호 계산기 "강수" 열 정렬 로직 개선:
        *   숫자 값은 선택된 정렬 방향(오름차순/내림차순)에 따라 올바르게 정렬되도록 수정.
        *   "선"이 포함된 문자열은 일반 문자열 정렬을 따르며, 숫자 값보다 뒤에 정렬되도록 유지.

*   **v1.0.1 변경 사항:**
    *   **경기 번호 계산기:**
        *   일반 경기 번호 계산 로직.
        *   '자유품새' 종목에 대한 특수 계산 로직 개선:
            *   참가 인원에 따른 예선, 본선, 결선 규칙 적용.
            *   조를 나눌 때 남는 인원이 첫 조부터 순서대로 들어가도록 수정.
            *   조의 수가 항상 짝수가 되도록 조정.
        *   '강수' 필드에서 '강' 문자 제거 및 '준결승', '결승' 값 조정.
        *   '경기수' 필드에 '1~8' 대신 숫자 '8'이 표시되도록 수정.
    *   **사용자 인터페이스 (GUI) 개선:**
        *   '+' 버튼 폰트 크기 확대.
        *   각 입력 행에 '-' 버튼 추가 및 해당 행 삭제 기능 구현.
        *   행 삭제 시 번호 자동 재정렬 기능 추가.
        *   전체 창 및 계산 결과 필드에 스크롤바 추가.
        *   입력 필드 영역 마우스 휠 스크롤 기능 추가.
        *   모든 GUI 창 하단에 저작권 푸터 추가 (Copyright (c) FEELJAE-WON. All rights reserved.).
        *   GUI 타이틀에 버전 및 빌드 날짜 표시.
    *   **파일 관리:**
        *   결과 파일 다운로드 이름이 항상 "경기번호계산_{현재시간}.xlsx" 형식으로 고정.
        *   `.exe` 파일 아이콘 적용.
        *   빌드 시 `tesseract.exe` 포함 제외.

## 다운로드

최신 실행 파일 (`feelsUtil.exe`)은 프로젝트의 [GitHub Releases](https://github.com/Feeljae-Won/challenge_utils/releases) 페이지에서 다운로드할 수 있습니다.