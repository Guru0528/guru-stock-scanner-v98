GURU V9.8 CLOSE TRADER FINAL

[목적]
- 종가매매 기능만 유지한 최종 정리 버전
- 급등주 검색/급등검증 UI 없음
- V9.1 원본 종가매매 검색 및 점수/검증 흐름 유지
- PC + 휴대폰 브라우저 사용 지원

[PC 실행]
1. ZIP을 완전히 압축 해제합니다.
2. KIS API 키를 기존 방식(.env)으로 설정합니다.
3. V9_8_PC_RUN.bat 실행
4. PC 브라우저: http://localhost:8501

[휴대폰 - 같은 Wi-Fi]
1. PC에서 V9_8_PC_RUN.bat을 실행한 상태로 둡니다.
2. 검은 창에 표시되는 Network URL을 휴대폰 브라우저에 입력합니다.
3. PC와 휴대폰은 같은 Wi-Fi/공유기 네트워크여야 합니다.
4. Windows 방화벽이 Python/8501 포트를 차단하면 허용해야 합니다.

[휴대폰 - 외부에서도 사용 / 클라우드]
- Streamlit Community Cloud 등에 이 폴더를 배포하면 PC가 꺼져 있어도 휴대폰에서 접속할 수 있습니다.
- 배포 Entry point: app.py (또는 streamlit_app.py)
- Secrets에 다음 값을 등록합니다.
  KIS_APP_KEY="본인 키"
  KIS_APP_SECRET="본인 시크릿"
  KIS_MODE="REAL"
- API 키/시크릿은 GitHub 저장소나 ZIP에 직접 적지 마십시오.

[종가매매 사용 순서]
- 15:10~15:20: 장중 후보검색 = 예비 후보 확인
- 16:00 이후: 장마감 확정검색 = 최종 판단 기준
- TOP: 장마감 결과 우선 표시
- 애프터: 선택 종목의 이후 가격 확인 보조

[중요]
- 검색기는 매매 판단 보조 도구이며 수익을 보장하지 않습니다.
- 장중 후보보다 장마감 확정검색 결과를 우선 사용하십시오.
