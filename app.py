import os,sys,subprocess,json,time
from pathlib import Path
from datetime import datetime
import pandas as pd
import streamlit as st

ROOT=Path(__file__).resolve().parent
RESULTS=ROOT/"results"; DATA=ROOT/"data"
RESULTS.mkdir(exist_ok=True);DATA.mkdir(exist_ok=True)

st.set_page_config(page_title="GURU V10 CLOUD PRO",page_icon="📈",layout="wide",initial_sidebar_state="collapsed")
st.markdown("""<style>
.block-container{padding:.7rem .8rem 3rem;max-width:1080px}
h1{font-size:1.55rem!important}.stButton>button{width:100%;min-height:3.1rem;font-weight:800;border-radius:13px}
[data-testid="stMetric"]{background:#f7f9fc;border:1px solid #e6ebf2;padding:9px;border-radius:12px}
div[data-testid="stExpander"]{border-radius:12px}
</style>""",unsafe_allow_html=True)

def secret_env():
    env=os.environ.copy()
    for k in ("KIS_APP_KEY","KIS_APP_SECRET","KIS_MODE"):
        try:
            if k in st.secrets and st.secrets[k]: env[k]=str(st.secrets[k])
        except Exception: pass
    return env

def run_mode(mode):
    env=secret_env();env["SCANNER_MODE"]=mode
    progress_file=DATA/"scan_progress.json"
    try: progress_file.unlink(missing_ok=True)
    except Exception: pass
    before={p.resolve() for p in RESULTS.glob("*.xlsx")}
    status_box=st.empty();bar=st.progress(0);metrics=st.empty();detail=st.empty();reasons_box=st.empty()
    proc=subprocess.Popen([sys.executable,str(ROOT/"scanner_engine.py")],cwd=str(ROOT),env=env,
                          stdout=subprocess.PIPE,stderr=subprocess.STDOUT,text=True,encoding="utf-8",errors="replace",bufsize=1)
    started=time.time();last={}
    while proc.poll() is None:
        try:
            if progress_file.exists(): last=json.loads(progress_file.read_text(encoding="utf-8"))
        except Exception: pass
        done=int(last.get("phase_done",last.get("done",0)) or 0)
        total=int(last.get("phase_total",last.get("total",0)) or 0)
        universe=int(last.get("universe_total",0) or 0)
        phase=str(last.get("phase","검색") or "검색")
        pctv=int(last.get("overall_pct",0) or 0)
        if not pctv and total: pctv=min(99,int(done*100/total))
        pctv=max(0,min(99,pctv))
        bar.progress(pctv)
        if total:
            if phase=="1차 고속필터":
                msg=f"1차 고속필터 · 전체 {universe or total:,}종목 중 {done:,}종목 확인 · 전체진행 {pctv}%"
            elif phase=="2차 정밀분석":
                msg=f"2차 정밀분석 · 1차 검색 {universe:,}종목 완료 · 정밀분석 {done:,}/{total:,}종목 · 전체진행 {pctv}%"
            else:
                msg=f"{phase} · {done:,}/{total:,}종목 · 전체진행 {pctv}%"
            status_box.info(msg)
        else:
            status_box.info("검색 준비 중 · API 인증/종목목록 확인")
        with metrics.container():
            a,b,c,d,e=st.columns(5)
            a.metric("API 성공",f"{int(last.get('api_ok',0) or 0):,}")
            b.metric("API 실패",f"{int(last.get('api_fail',0) or 0):,}")
            c.metric("제외",f"{int(last.get('excluded',0) or 0):,}")
            d.metric("지표계산",f"{int(last.get('indicator_ok',0) or 0):,}")
            e.metric("현재 후보",f"{int(last.get('pass_count',0) or 0):,}")
        elapsed=int(time.time()-started);eta=int(last.get("eta",0) or 0)
        detail.caption(f"경과 {elapsed//60:02d}분 {elapsed%60:02d}초 · 현재 단계 예상 남은시간 {eta//60:02d}분 {eta%60:02d}초 · 숫자는 단계별로 고정 표시됩니다.")
        reasons=last.get("reasons") or {}
        if reasons: reasons_box.caption("현재 주요 분류/탈락: "+" · ".join(f"{k} {v:,}" for k,v in reasons.items()))
        time.sleep(.5)
    out=""
    if proc.stdout:
        try: out=proc.stdout.read()
        except Exception: pass
    rc=proc.wait()
    if rc!=0:
        status_box.error("검색 중 오류가 발생했습니다.");bar.progress(0)
        return False,None,0,f"검색 오류 (코드 {rc})",out,out
     # 검색 완료 후 가장 최근 생성/갱신된 결과파일 찾기
    files = sorted(
        RESULTS.glob("*.xlsx"),
        key=lambda p: p.stat().st_mtime,
        reverse=True
    )

    if mode == "AFTER":
        files = [p for p in files if "AFTER" in p.name.upper()]
    else:
        files = [p for p in files if "AFTER" not in p.name.upper()]

    result = files[0] if files else None

    if not result:
        status_box.error("검색은 종료됐지만 새 결과파일이 없습니다.")
        return False,None,0,"검색은 종료됐지만 새 결과파일이 없습니다.",out,""

    try:
        sh="애프터확인" if mode=="AFTER" else "TOP10_한눈에보기"
        d=pd.read_excel(result,sheet_name=sh)
        bar.progress(100)
        status_box.success(f"검색 완료 · TOP {len(d)}개 · 결과 저장 완료")
        return True,result,len(d),f"{len(d)}개",out,""
    except Exception as ex:
        status_box.error("결과파일 검증 중 오류가 발생했습니다.")
        return False,result,0,f"결과 검증 실패: {ex}",out,""

def latest(pattern):
    fs=sorted(RESULTS.glob(pattern),key=lambda p:p.stat().st_mtime,reverse=True)
    return fs[0] if fs else None  
def topfile():
    fs = [
        p for p in RESULTS.glob("*.xlsx")
        if "AFTER" not in p.name.upper()
    ]
    fs = sorted(fs, key=lambda p: p.stat().st_mtime, reverse=True)
    return fs[0] if fs else None

def afterfile():
    fs = [
        p for p in RESULTS.glob("*.xlsx")
        if "AFTER" in p.name.upper()
    ]
    fs = sorted(fs, key=lambda p: p.stat().st_mtime, reverse=True)
    return fs[0] if fs else None

def won(v):
    try:
        if pd.isna(v):
            return "-"
        return f"{float(v):,.0f}원"
    except:
        return "-"

def pct(v):
    try:
        if pd.isna(v):
            return "표본없음"
        return f"{float(v)*100:.1f}%"
    except:
        return "표본없음"

def readtop():
    p = topfile()
    if not p:
        return p, pd.DataFrame()
    try:
        return p, pd.read_excel(
            p,
            sheet_name="TOP10_한눈에보기",
            dtype={"종목코드": str}
        )
    except:
        return p, pd.DataFrame()
st.title("GURU V10 CLOUD PRO")
st.caption("종가매매 · 익일~5일 스윙 · 2중 백테스트 검증 · PC OFF 모바일")

tabs=st.tabs(["🏆 TOP","🔎 검색","🌙 애프터","🧮 매매계산","📊 결과","⚙️ 상태"])

with tabs[0]:
    p,df=readtop()
    if p:st.caption("최신 결과 "+datetime.fromtimestamp(p.stat().st_mtime).strftime("%m/%d %H:%M"))
    if df.empty:st.info("아직 결과가 없습니다. 검색 탭에서 장마감 확정검색을 실행하세요.")
    else:
        st.success(f"TOP 후보 {len(df)}개")
        with st.expander("V10 검증등급 읽는 법"):
            st.write("강화검증: 개별 종목 과거표본 30건 이상. 시장표본보완: 개별 표본은 적지만 같은 전략군의 시장 전체 표본이 100건 이상. 검증부족: 둘 다 부족.")
            st.write("목표 선도달률은 높을수록, 손절률은 낮을수록 유리합니다. 통계는 미래 수익을 보장하지 않습니다.")
        for _,r in df.iterrows():
            rank=int(r.get("순위",0)); name=str(r.get("종목명","")); grade=str(r.get("등급",""))
            val=str(r.get("검증등급","검증부족"))
            with st.expander(f"{rank}위 {name} · {grade} · {r.get('최종추천점수','')}점 · {val}",expanded=rank<=3):
                a,b,c=st.columns(3);a.metric("종가",won(r.get("종가")));b.metric("손절",won(r.get("손절가")));c.metric("익일목표",won(r.get("익일목표가")))
                st.write(f"**매수구간** {won(r.get('매수구간하단'))} ~ {won(r.get('매수구간상단'))}")
                st.write(f"**3일 / 5일 목표** {won(r.get('3일목표가'))} / {won(r.get('5일목표가'))}")
                st.write(f"**전략** {r.get('선정전략','')}　|　**보유판정** {r.get('추천보유유형','')}")
                st.write(f"개별표본 **{r.get('백테스트표본',0)}건** · 익일 {pct(r.get('익일+2%선도달률'))} · 3일 {pct(r.get('3일+3%선도달률'))} · 5일 {pct(r.get('5일+5%선도달률'))} · 손절 {pct(r.get('손절선도달률'))}")
                st.write(f"시장전략표본 **{r.get('시장전략표본',0)}건** · 익일 {pct(r.get('시장익일+2%'))} · 3일 {pct(r.get('시장3일+3%'))} · 5일 {pct(r.get('시장5일+5%'))} · 손절 {pct(r.get('시장손절률'))}")

with tabs[1]:
    st.subheader("검색 실행")
    st.info("클라우드 배포 후에는 PC가 꺼져 있어도 이 화면에서 검색할 수 있습니다.")
    if st.button("① 15:10~15:20 장중 예비검색",type="primary"):
        ok,p,n,msg,out,err=run_mode("INTRADAY")
        if ok:st.success(f"장중 검색 완료 · TOP {n}개");st.rerun()
        else:
            st.error(msg)
            with st.expander("오류 상세"):st.code((err or out)[-5000:])
    if st.button("② 16:00 이후 장마감 확정검색",type="primary"):
        ok,p,n,msg,out,err=run_mode("CLOSE")
        if ok:st.success(f"장마감 확정검색 완료 · TOP {n}개");st.rerun()
        else:
            st.error(msg)
            with st.expander("오류 상세"):st.code((err or out)[-5000:])
    if st.button("③ 애프터마켓 TOP 재확인"):
        with st.spinner("애프터 현재가 확인 중..."):ok,p,n,msg,out,err=run_mode("AFTER")
        if ok:st.success("애프터 확인 완료");st.rerun()
        else:st.error(msg)

with tabs[2]:
    p=afterfile()
    if st.button("🔄 애프터 새로 확인",type="primary"):
        with st.spinner("확인 중..."):ok,_,_,msg,_,_=run_mode("AFTER")
        if ok:st.rerun()
        else:st.error(msg)
    if not p:st.info("애프터 결과가 없습니다.")
    else:
        try:d=pd.read_excel(p,sheet_name="애프터확인",dtype={"종목코드":str})
        except:d=pd.DataFrame()
        for _,r in d.iterrows():
            st.markdown(f"### {r.get('순위','')}위 {r.get('종목명','')} · {r.get('애프터판정','')}")
            a,b,c=st.columns(3);a.metric("정규장",won(r.get("정규장종가")));b.metric("현재",won(r.get("애프터현재가")));c.metric("등락",pct(r.get("종가대비등락률")))
            st.divider()

with tabs[3]:
    capital=st.number_input("투입금액",100000,step=100000,value=3000000)
    entry=st.number_input("예상 매수가",1,step=10,value=10000);stop=st.number_input("손절가",1,step=10,value=9700)
    risk=st.slider("허용손실률",.2,3.0,1.0,.1)/100
    q1=int(capital//entry);q2=int((capital*risk)//max(entry-stop,1));qty=max(0,min(q1,q2))
    a,b,c=st.columns(3);a.metric("권장수량",f"{qty:,}주");b.metric("매수금액",won(qty*entry));c.metric("손절손실",won(qty*max(entry-stop,0)))

with tabs[4]:
    fs=sorted(RESULTS.glob("*.xlsx"),key=lambda p:p.stat().st_mtime,reverse=True)[:30]
    if fs:
        st.dataframe(pd.DataFrame([{"시간":datetime.fromtimestamp(f.stat().st_mtime).strftime("%m/%d %H:%M"),"파일":f.name} for f in fs]),use_container_width=True,hide_index=True)
    else:st.info("결과 없음")

with tabs[5]:
    local=(ROOT/".env").exists()
    cloud=False
    try:cloud=bool(st.secrets.get("KIS_APP_KEY",""))
    except:pass
    st.write("API 인증:", "✅ 준비됨" if (local or cloud) else "❌ 설정 필요")
    st.write("실행환경:", "클라우드 Secrets" if cloud else "PC .env" if local else "미설정")
    st.write("설정파일: Guru_Stock_Scanner_V10_CLOUD_PRO.xlsx")
    st.warning("APP KEY/SECRET은 GitHub에 올리지 말고 Streamlit Secrets에만 저장하세요.")
    st.caption("자동주문 기능은 포함하지 않습니다.")
