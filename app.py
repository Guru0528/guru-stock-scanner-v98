import os,sys,subprocess,json,time,math
from pathlib import Path
from datetime import datetime
import pandas as pd
import streamlit as st

ROOT=Path(__file__).resolve().parent; RESULTS=ROOT/"results"; DATA=ROOT/"data"
RESULTS.mkdir(exist_ok=True); DATA.mkdir(exist_ok=True)
sys.path.insert(0,str(ROOT))
import scanner_engine as eng

st.set_page_config(page_title="GURU V9.8 CLOSE TRADER FINAL",page_icon="📈",layout="wide",initial_sidebar_state="collapsed")
st.markdown("""<style>
.block-container{padding:0.8rem .8rem 3rem;max-width:1050px}
h1{font-size:1.55rem!important}.stButton>button{width:100%;min-height:3.1rem;font-weight:800;border-radius:13px}
[data-testid="stMetric"]{background:#f7f9fc;border:1px solid #e6ebf2;padding:10px;border-radius:13px}
div[data-testid="stExpander"]{border-radius:13px}
.small{font-size:.84rem;color:#687386}.buy{font-weight:800}
@media (max-width: 700px){
 .block-container{padding:.45rem .45rem 2rem;max-width:100%}
 h1{font-size:1.3rem!important}
 [data-testid="stMetric"]{padding:7px}
 .stButton>button{min-height:2.8rem}
}
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
    pf=DATA/"scan_progress.json"
    try: pf.unlink(missing_ok=True)
    except Exception: pass
    before={p.resolve() for p in RESULTS.glob("*.xlsx")}
    status=st.empty();bar=st.progress(0);metrics=st.empty();detail=st.empty()
    log_path=ROOT/"logs"/"scanner_live.log"
    log_path.parent.mkdir(exist_ok=True)
    log_fp=open(log_path,"w",encoding="utf-8",buffering=1)
    proc=subprocess.Popen([sys.executable,str(ROOT/"scanner_engine.py")],cwd=str(ROOT),env=env,
                          stdout=log_fp,stderr=subprocess.STDOUT,text=True)
    started=time.time();last={}
    while proc.poll() is None:
        try:
            if pf.exists(): last=json.loads(pf.read_text(encoding="utf-8"))
        except Exception: pass
        done=int(last.get("done",0) or 0);total=int(last.get("total",0) or 0);pctv=min(100,int(done*100/total)) if total else 0
        bar.progress(pctv);status.info(f"V9.1 원본검색 · {done:,}/{total:,}종목 ({pctv}%)" if total else "검색 준비 중 · API 인증/종목목록 확인")
        with metrics.container():
            a,b,c,d,e=st.columns(5)
            a.metric("API 성공",f"{int(last.get('api_ok',0) or 0):,}");b.metric("API 실패",f"{int(last.get('api_fail',0) or 0):,}")
            c.metric("제외",f"{int(last.get('excluded',0) or 0):,}");d.metric("지표계산",f"{int(last.get('indicator_ok',0) or 0):,}")
            e.metric("현재 후보",f"{int(last.get('pass_count',0) or 0):,}")
        elapsed=int(time.time()-started);eta=int(last.get("eta",0) or 0)
        detail.caption(f"경과 {elapsed//60:02d}분 {elapsed%60:02d}초 · 예상 남은시간 {eta//60:02d}분 {eta%60:02d}초")
        time.sleep(.5)
    rc=proc.wait(); log_fp.close()
    try: out=log_path.read_text(encoding="utf-8",errors="replace")
    except Exception: out=""
    after={q.resolve() for q in RESULTS.glob("*.xlsx")};created=sorted(after-before,key=lambda x:Path(x).stat().st_mtime,reverse=True)
    result=Path(created[0]) if created else None
    if rc!=0:return False,None,0,f"검색 오류 (코드 {rc})",out,out
    if not result:return False,None,0,"검색은 종료됐지만 새 결과파일이 없습니다.",out,""
    try:
        sheet="애프터확인" if mode=="AFTER" else "TOP10_한눈에보기";chk=pd.read_excel(result,sheet_name=sheet)
        bar.progress(100);status.success(f"검색 완료 · {len(chk)}개 · Excel 저장 완료")
        return True,result,len(chk),f"결과파일 생성 확인 · {len(chk)}개 후보",out,""
    except Exception as ex:return False,result,0,f"결과파일 검증 실패: {ex}",out,""

def latest(pattern):
    fs=sorted(RESULTS.glob(pattern),key=lambda p:p.stat().st_mtime,reverse=True);return fs[0] if fs else None
def close_file(): return latest("Guru_Stock_Scanner_V9_1_CLOSE_*.xlsx")
def intra_file(): return latest("Guru_Stock_Scanner_V9_1_INTRADAY_*.xlsx")
def after_file(): return latest("Guru_Stock_Scanner_V9_1_AFTER_*.xlsx")
def read_top(p):
    if not p:return pd.DataFrame()
    try:
        df=pd.read_excel(p,sheet_name="TOP10_한눈에보기",dtype={"종목코드":str})
        required=["순위","종목명","종목코드","종가","매수구간하단","매수구간상단","손절가",
                  "익일목표가","기술점수","종가매매점수","검증점수","최종판정"]
        missing=[c for c in required if c not in df.columns]
        if missing:
            st.error("현재 결과파일은 V9.8 형식이 아닙니다. 장마감 확정검색을 새로 실행하세요. 누락: "+", ".join(missing))
            return pd.DataFrame()
        # 동일 행 가격 데이터의 기본 무결성 검사
        bad=[]
        for i,r in df.iterrows():
            try:
                close=float(r["종가"]); lo=float(r["매수구간하단"]); hi=float(r["매수구간상단"])
                stop=float(r["손절가"]); target=float(r["익일목표가"])
                if not (lo <= hi and stop < close < target and lo < close*1.10 and hi > close*0.90):
                    bad.append(i)
            except Exception:
                bad.append(i)
        if bad:
            st.error(f"결과 데이터 무결성 오류 {len(bad)}건을 발견했습니다. 잘못된 행은 표시하지 않습니다.")
            df=df.drop(index=bad).reset_index(drop=True)
        return df.copy(deep=True)
    except Exception as ex:
        st.error(f"결과파일 읽기 오류: {ex}")
        return pd.DataFrame()
def won(v):
    try:return f"{float(v):,.0f}원"
    except:return "-"
def pct(v):
    try:
        x=float(v)
        if math.isnan(x): return "산출불가"
        return f"{x*100:.1f}%"
    except:return "산출불가"
def watch_path():return DATA/"watchlist.json"
def get_watch():
    try:return json.loads(watch_path().read_text(encoding="utf-8"))
    except:return []
def save_watch(x):watch_path().write_text(json.dumps(x,ensure_ascii=False,indent=2),encoding="utf-8")

st.title("GURU V9.8 CLOSE TRADER FINAL")
st.caption("종가매매 전용 · V9.1 원본검색/점수체계 유지 · 급등주 기능 제거 · PC/휴대폰 지원")

tabs=st.tabs(["🏆 TOP","🔎 검색","🌙 애프터","⭐ 관심","🧮 매매계산","📊 기록","⚙️ 설정"])

with tabs[0]:
    p=close_file() or intra_file();df=read_top(p)
    if p:st.caption("최신 결과 · "+datetime.fromtimestamp(p.stat().st_mtime).strftime("%m/%d %H:%M"))
    if df.empty:st.info("검색 결과가 없습니다. '검색'에서 먼저 실행하세요.")
    else:
        st.caption("사용 결과파일: "+p.name)
        for rec in df.to_dict(orient="records"):
            r=dict(rec)
            rank=int(r.get("순위",0)); name=str(r.get("종목명","")); grade=str(r.get("등급",""))
            decision=str(r.get("최종판정",""))
            cscore=r.get("종가매매점수",r.get("V9추천점수",""))
            title=f"{rank}위  {name}  · {decision} · 종가매매 {cscore}점"
            with st.expander(title,expanded=rank<=3):
                a,b,c=st.columns(3);a.metric("종가",won(r.get("종가")));b.metric("손절",won(r.get("손절가")));c.metric("익일목표",won(r.get("익일목표가")))
                st.write(f"**매수구간** {won(r.get('매수구간하단'))} ~ {won(r.get('매수구간상단'))}")
                st.write(f"**3일 / 5일 목표** {won(r.get('3일목표가'))} / {won(r.get('5일목표가'))}")
                st.write(f"**최종판정** {decision}　·　**기술점수** {r.get('기술점수','-')}　·　**검증점수** {r.get('검증점수','-')}")
                st.write(f"**보유** {r.get('추천보유유형','')}　·　**전략** {r.get('선정전략','')}")
                st.caption(f"백테스트 {int(r.get('백테스트표본',0) or 0)}건 · 신뢰도 {r.get('백테스트신뢰도','')} · 익일 +2% {pct(r.get('익일+2%선도달률'))} · 손절선 도달 {pct(r.get('손절선도달률'))}")
                if st.button("⭐ 관심종목 추가",key=f"w{rank}_{r.get('종목코드')}"):
                    w=get_watch(); item={"code":str(r.get("종목코드")).zfill(6),"name":name,"added":datetime.now().isoformat(timespec="minutes")}
                    if not any(x["code"]==item["code"] for x in w):w.append(item);save_watch(w)
                    st.success("추가했습니다.")

with tabs[1]:
    st.subheader("종가매매 검색")
    st.info("15:10~15:20 장중 검색은 종가매매 예비후보 선정용입니다. 최종 매수 판단은 16:00 이후 장마감 확정검색 결과를 기준으로 하세요.")
    if st.button("① 15:10~15:20 장중 예비검색",type="primary"):
        with st.spinner("코스피/코스닥 전 종목을 검색하고 결과파일을 검증 중입니다..."):
            ok,result,rows,msg,out,err=run_mode("INTRADAY")
        if ok:
            st.success(f"장중 예비검색 완료 · 후보 {rows}개")
            st.caption(result.name)
            st.rerun()
        else:
            st.error(msg)
            with st.expander("오류 상세"): st.code((err or out)[-5000:])
    if st.button("② 16:00 이후 장마감 확정검색",type="primary"):
        with st.spinner("코스피/코스닥 전 종목 확정검색 + 결과파일 검증 중입니다..."):
            ok,result,rows,msg,out,err=run_mode("CLOSE")
        if ok:
            st.success(f"장마감 확정검색 완료 · TOP 후보 {rows}개")
            st.caption(result.name)
            st.rerun()
        else:
            st.error(msg)
            with st.expander("오류 상세"): st.code((err or out)[-5000:])
    if st.button("③ 애프터마켓 TOP 재확인"):
        with st.spinner("TOP 종목 애프터 현재가 확인 + 결과 검증 중입니다..."):
            ok,result,rows,msg,out,err=run_mode("AFTER")
        if ok:
            st.success(f"애프터 확인 완료 · {rows}개")
            st.rerun()
        else:
            st.error(msg)
            with st.expander("오류 상세"): st.code((err or out)[-5000:])

with tabs[2]:
    p=after_file()
    if st.button("🔄 애프터 현재가 새로 확인",type="primary"):
        with st.spinner("확인 중..."):
            ok,result,rows,msg,out,err=run_mode("AFTER")
        if ok: st.rerun()
        else: st.error(msg)
    if not p:st.info("애프터 결과가 없습니다.")
    else:
        try:ad=pd.read_excel(p,sheet_name="애프터확인",dtype={"종목코드":str})
        except:ad=pd.DataFrame()
        for _,r in ad.iterrows():
            judge=str(r.get("애프터판정",""));icon={"매수관찰":"🟢","추격금지":"🔴","급락주의":"🔴","상승주의":"🟡","눌림관찰":"🔵"}.get(judge,"⚪")
            st.markdown(f"### {icon} {r.get('순위','')}위 {r.get('종목명','')} · {judge}")
            a,b,c=st.columns(3);a.metric("정규장",won(r.get("정규장종가")));b.metric("현재",won(r.get("애프터현재가")));c.metric("등락",pct(r.get("종가대비등락률")))
            st.divider()

with tabs[3]:
    w=get_watch()
    if not w:st.info("TOP 화면에서 관심종목을 추가하세요.")
    for i,x in enumerate(w):
        c1,c2=st.columns([4,1]);c1.write(f"**{x['name']}**　{x['code']}　추가 {x['added']}")
        if c2.button("삭제",key=f"del{i}"):w.pop(i);save_watch(w);st.rerun()

with tabs[4]:
    st.subheader("1회 매매 수량 계산")
    capital=st.number_input("투입금액(원)",min_value=100000,step=100000,value=3000000)
    entry=st.number_input("예상 매수가",min_value=1,step=10,value=10000)
    stop=st.number_input("손절가",min_value=1,step=10,value=9700)
    riskpct=st.slider("계좌 허용손실률",0.2,3.0,1.0,0.1)/100
    qty_cap=int(capital//entry); risk_per=max(entry-stop,1); qty_risk=int((capital*riskpct)//risk_per)
    qty=max(0,min(qty_cap,qty_risk));loss=qty*risk_per
    a,b,c=st.columns(3);a.metric("권장수량",f"{qty:,}주");b.metric("매수금액",won(qty*entry));c.metric("손절시 손실",won(loss))
    st.caption("수수료·세금·슬리피지는 별도입니다. 계산기는 주문을 실행하지 않습니다.")

with tabs[5]:
    fs=sorted(RESULTS.glob("*.xlsx"),key=lambda p:p.stat().st_mtime,reverse=True)[:30]
    if not fs:st.info("저장된 결과가 없습니다.")
    else:
        hist=pd.DataFrame([{"시간":datetime.fromtimestamp(f.stat().st_mtime).strftime("%m/%d %H:%M"),"파일":f.name,
                            "유형":"애프터" if "_AFTER_" in f.name else "장마감" if "_CLOSE_" in f.name else "장중"} for f in fs])
        st.dataframe(hist,use_container_width=True,hide_index=True)

with tabs[6]:
    st.write("API 환경파일:", "✅ 연결됨" if (ROOT/".env").exists() else "❌ .env 없음")
    st.write("설정파일: `Guru_Stock_Scanner_V9_1_FIX.xlsx`")
    st.markdown("**권장 자동화 시각**  \n15:15 장중 예비검색 → 16:05 장마감 확정검색 → 17:00/19:00 애프터 재확인")
    st.warning("API Key/Secret은 화면이나 공개 GitHub에 올리지 마세요. 서버의 환경변수/Secret에 보관하세요.")
    st.info("V9는 분석·검색 보조용이며 자동 주문은 포함하지 않았습니다.")
