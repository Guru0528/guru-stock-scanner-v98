import os,io,time,json,zipfile,traceback,re,math
from pathlib import Path
from datetime import datetime,timedelta
from concurrent.futures import ThreadPoolExecutor,as_completed
from threading import Lock
import requests,pandas as pd,numpy as np
from dotenv import load_dotenv
from openpyxl import load_workbook

ROOT=Path(__file__).resolve().parent
DATA=ROOT/"data";RESULTS=ROOT/"results";LOGS=ROOT/"logs"
for p in (DATA,RESULTS,LOGS): p.mkdir(exist_ok=True)
load_dotenv(ROOT/".env")
KEY=os.getenv("KIS_APP_KEY","").strip(); SECRET=os.getenv("KIS_APP_SECRET","").strip()
MODE=os.getenv("KIS_MODE","REAL").upper()
BASE="https://openapi.koreainvestment.com:9443" if MODE=="REAL" else "https://openapivts.koreainvestment.com:29443"
book=load_workbook(ROOT/"Guru_Stock_Scanner_V9_1_FIX.xlsx",data_only=True)
CFG={r[0]:r[1] for r in book["조건설정"].iter_rows(min_row=2,values_only=True) if r[0]}
STAMP=datetime.now().strftime("%Y%m%d_%H%M%S");TOK=DATA/"token.json";ERR=LOGS/f"errors_{STAMP}.txt"
PROGRESS=DATA/"scan_progress.json"

# V10.1 - KIS API 안정화용 전역 호출 간격 제어
_API_LOCK=Lock()
_LAST_API_CALL=0.0
def api_wait():
    global _LAST_API_CALL
    gap=float(CFG.get("REQUEST_SLEEP",0.12) or 0.12)
    with _API_LOCK:
        now=time.monotonic()
        wait=gap-(now-_LAST_API_CALL)
        if wait>0: time.sleep(wait)
        _LAST_API_CALL=time.monotonic()

def write_progress(**kw):
    base={"status":"running","time":datetime.now().isoformat(timespec="seconds")}
    base.update(kw)
    tmp=PROGRESS.with_suffix(".tmp")
    try:
        tmp.write_text(json.dumps(base,ensure_ascii=False),encoding="utf-8")
        tmp.replace(PROGRESS)
    except Exception:
        pass

def elog(s):
    with ERR.open("a",encoding="utf-8") as f:f.write(str(s)+"\n")

def get_token():
    if TOK.exists():
        try:
            d=json.loads(TOK.read_text(encoding="utf-8"))
            if d.get("access_token") and d.get("expires_at",0)>time.time()+300:return d["access_token"]
        except: pass
    if not KEY or not SECRET: raise RuntimeError(".env에 KIS_APP_KEY / KIS_APP_SECRET가 필요합니다.")
    r=requests.post(BASE+"/oauth2/tokenP",json={"grant_type":"client_credentials","appkey":KEY,"appsecret":SECRET},timeout=30)
    r.raise_for_status();d=r.json()
    if "access_token" not in d: raise RuntimeError("토큰 발급 실패: "+str(d))
    d["expires_at"]=time.time()+23*3600;TOK.write_text(json.dumps(d),encoding="utf-8");return d["access_token"]

def clean_name(raw,code):
    raw=re.sub(r"^KR7\d{9,12}","",raw).strip()
    raw=re.sub(r"^[A-Z0-9]{12}","",raw).strip()
    return raw or code

def masters():
    out=[]
    for market,fn in [("KOSPI","kospi_code.mst.zip"),("KOSDAQ","kosdaq_code.mst.zip")]:
        cache=DATA/f"{market}_master.csv"
        try:
            r=requests.get("https://new.real.download.dws.co.kr/common/master/"+fn,timeout=30);r.raise_for_status()
            with zipfile.ZipFile(io.BytesIO(r.content)) as z:lines=z.read(z.namelist()[0]).decode("cp949","ignore").splitlines()
            rows=[]
            for line in lines:
                code=line[:9].strip()
                if len(code)==6 and code.isdigit():
                    chunk=line[9:90].strip()
                    name=clean_name(chunk,code)
                    name=re.sub(r"^KR7"+re.escape(code)+r"\d{3}","",name).strip()
                    name=re.split(r"\s{2,}|\d{8,}",name)[0].strip() or code
                    rows.append((code,name,market))
            if rows:
                pd.DataFrame(rows,columns=["code","name","market"]).to_csv(cache,index=False,encoding="utf-8-sig");out+=rows
        except Exception as e:
            elog(f"MASTER {market}: {e}")
            if cache.exists():out+=list(pd.read_csv(cache,dtype=str).itertuples(index=False,name=None))
    return pd.DataFrame(out,columns=["code","name","market"]).drop_duplicates("code").reset_index(drop=True)

def security_type(name):
    n=str(name).upper().replace(" ","")
    if any(x in n for x in ["KODEX","TIGER","RISE","SOL","ACE","HANARO","KOSEF","ARIRANG","PLUS","TIMEFOLIO","FOCUS","TREX","KBSTAR","ETF","ETN"]):
        return "ETF_ETN"
    if "스팩" in str(name) or "SPAC" in n:return "SPAC"
    # Korean preferred shares: 우/우B/1우/2우B etc. Avoid false positive for normal words by checking suffix.
    if re.search(r"(우|우B|우C|[1-9]우|[1-9]우B)$",str(name).strip(),re.I):return "PREFERRED"
    return "COMMON"

def excluded_type(t):
    if t=="ETF_ETN" and CFG["EXCLUDE_ETF_ETN"]=="Y":return True
    if t=="SPAC" and CFG["EXCLUDE_SPAC"]=="Y":return True
    if t=="PREFERRED" and CFG["EXCLUDE_PREFERRED"]=="Y":return True
    return False

def fetch(code,tk,s,e):
    h={"authorization":"Bearer "+tk,"appkey":KEY,"appsecret":SECRET,"tr_id":"FHKST03010100","custtype":"P"}
    p={"FID_COND_MRKT_DIV_CODE":"J","FID_INPUT_ISCD":code,"FID_INPUT_DATE_1":s,"FID_INPUT_DATE_2":e,
       "FID_PERIOD_DIV_CODE":"D","FID_ORG_ADJ_PRC":"1"}
    last=""
    for a in range(3):
        api_wait()
        try:
            r=requests.get(BASE+"/uapi/domestic-stock/v1/quotations/inquire-daily-itemchartprice",headers=h,params=p,timeout=20)
            if r.status_code==200:
                j=r.json()
                if str(j.get("rt_cd"))=="0": return j.get("output2",[]),""
                last=f'{j.get("msg_cd","")} {j.get("msg1","")}'
                if any(x in last.lower() for x in ["초당","rate","limit","too many"]):
                    time.sleep(1.0*(a+1)); continue
            else:
                last=f"HTTP {r.status_code}: {r.text[:150]}"
                if r.status_code in (429,500,502,503,504):
                    time.sleep(1.0*(a+1)); continue
        except Exception as ex:
            last=repr(ex); time.sleep(.8*(a+1))
    return [],last

def history(code,tk):
    end=datetime.now().date();start=end-timedelta(days=int(CFG["LOOKBACK_CALENDAR_DAYS"]))
    rows=[];cur=end
    while cur>=start:
        s=max(start,cur-timedelta(days=90))
        rr,msg=fetch(code,tk,s.strftime("%Y%m%d"),cur.strftime("%Y%m%d"))
        if not rr and msg:return [],msg
        rows.extend(rr);cur=s-timedelta(days=1)
    return rows,""

def frame(rows):
    a=[]
    for x in rows:
        try:a.append({"date":str(x["stck_bsop_date"]),"open":float(x["stck_oprc"]),"high":float(x["stck_hgpr"]),
                      "low":float(x["stck_lwpr"]),"close":float(x["stck_clpr"]),"volume":float(x["acml_vol"]),
                      "turnover":float(x.get("acml_tr_pbmn") or 0)})
        except:pass
    return pd.DataFrame(a).drop_duplicates("date").sort_values("date").reset_index(drop=True) if a else pd.DataFrame()

def enrich(d,n=14):
    d=d.copy()
    d["ma5"]=d.close.rolling(5).mean();d["ma20"]=d.close.rolling(20).mean();d["ma60"]=d.close.rolling(60).mean()
    d["v20"]=d.volume.rolling(20).mean()
    delta=d.close.diff();up=delta.clip(lower=0);dn=-delta.clip(upper=0)
    au=up.ewm(alpha=1/n,adjust=False,min_periods=n).mean();ad=dn.ewm(alpha=1/n,adjust=False,min_periods=n).mean()
    d["rsi14"]=100-100/(1+au/ad.replace(0,np.nan))
    tr=pd.concat([d.high-d.low,(d.high-d.close.shift()).abs(),(d.low-d.close.shift()).abs()],axis=1).max(axis=1)
    d["atr14"]=tr.ewm(alpha=1/n,adjust=False,min_periods=n).mean()
    u=d.high.diff();dd=-d.low.diff();plus=np.where((u>dd)&(u>0),u,0.);minus=np.where((dd>u)&(dd>0),dd,0.)
    d["plus_di"]=100*pd.Series(plus,index=d.index).ewm(alpha=1/n,adjust=False,min_periods=n).mean()/d.atr14.replace(0,np.nan)
    d["minus_di"]=100*pd.Series(minus,index=d.index).ewm(alpha=1/n,adjust=False,min_periods=n).mean()/d.atr14.replace(0,np.nan)
    dx=100*(d.plus_di-d.minus_di).abs()/(d.plus_di+d.minus_di).replace(0,np.nan)
    d["adx14"]=dx.ewm(alpha=1/n,adjust=False,min_periods=n).mean()
    d["high20_prev"]=d.high.shift(1).rolling(20).max()
    rng=(d.high-d.low).replace(0,np.nan)
    d["close_location"]=(d.close-d.low)/rng
    d["upper_wick"]=(d.high-np.maximum(d.open,d.close))/rng
    d["body_ratio"]=(d.close-d.open).abs()/rng
    d["bull"]=d.close>=d.open
    return d

def technical_score(x,p):
    close=float(x.close);ma20=float(x.ma20);ma60=float(x.ma60);high20=float(x.high20_prev)
    dist=close/ma20-1;vr=float(x.volume/x.v20) if x.v20 else 0
    highdist=close/high20-1 if high20 else 999;atrpct=float(x.atr14/close) if close else 999
    cl=float(x.close_location);uw=float(x.upper_wick);rsi=float(x.rsi14);adx=float(x.adx14)
    turn=float(x.turnover or close*x.volume);sc=0
    if cl>=.85:sc+=15
    elif cl>=float(CFG["MIN_CLOSE_LOCATION"]):sc+=10
    if bool(x.bull) and uw<=.20:sc+=10
    elif uw<=float(CFG["MAX_UPPER_WICK_RATIO"]):sc+=5
    if abs(dist)<=float(CFG["TOUCH_TOLERANCE"]) and ma20>float(p.ma20):sc+=15
    elif ma20>float(p.ma20):sc+=8
    if close>ma60:sc+=10
    if vr>=1.5:sc+=10
    elif vr>=1.2:sc+=8
    elif vr>=float(CFG["MIN_VOLUME_RATIO"]):sc+=4
    if turn>=10_000_000_000:sc+=5
    elif turn>=float(CFG["MIN_TURNOVER"]):sc+=3
    if 45<=rsi<=68:sc+=10
    elif float(CFG["RSI_MIN"])<=rsi<=float(CFG["RSI_MAX"]):sc+=5
    if adx>=25 and float(x.plus_di)>float(x.minus_di):sc+=10
    elif adx>=float(CFG["ADX_MIN"]) and float(x.plus_di)>float(x.minus_di):sc+=6
    if -.08<=highdist<=0:sc+=10
    elif 0<highdist<=.01:sc+=6
    if .015<=atrpct<=.045:sc+=5
    elif atrpct<=float(CFG["MAX_ATR_PCT"]):sc+=2
    return sc,dist,vr,turn,highdist,atrpct,cl,uw

def strategy(x,p,dist,vr,hd):
    tags=[]
    if abs(dist)<=float(CFG["TOUCH_TOLERANCE"]) and float(x.ma20)>float(p.ma20) and float(x.close)>float(x.ma60):tags.append("A_20일선눌림")
    if -.035<=hd<=.01 and float(x.close_location)>=.75 and float(x.plus_di)>float(x.minus_di):tags.append("B_전고점돌파직전")
    if vr>=1.5 and float(x.close)>float(x.ma20)>float(x.ma60) and float(x.adx14)>=22 and float(x.close_location)>=.75:tags.append("C_거래량추세가속")
    return "+".join(tags) if tags else "기본"

def grade(sc):
    if sc>=float(CFG["GRADE_A_PLUS"]):return "A+"
    if sc>=float(CFG["GRADE_A"]):return "A"
    if sc>=float(CFG["GRADE_B"]):return "B"
    return "제외"

def grade_rank(g): return {"A+":1,"A":2,"B":3,"제외":9}.get(g,9)

def bt_conf(n):
    if n>=50:return "높음"
    if n>=30:return "양호"
    if n>=10:return "보통"
    return "낮음"

def first_touch_outcomes(q, i, entry, atr):
    stop=max(entry*(1-float(CFG["STOP_MAX_PCT"])),entry-atr*float(CFG["STOP_ATR"]))
    targets=[entry*(1+float(CFG["NEXTDAY_TARGET_PCT"])),
             entry*(1+float(CFG["DAY3_TARGET_PCT"])),
             entry*(1+float(CFG["DAY5_TARGET_PCT"]))]
    hit1=hit3=hit5=False; stop_first=False; alive=True
    for day in range(1,6):
        r=q.iloc[i+day];lo=float(r.low);hi=float(r.high)
        # Conservative OHLC rule: if stop and a target are both touched on same day, stop is assumed first.
        if alive and lo<=stop:
            stop_first=True;alive=False
            break
        if day==1 and hi>=targets[0]:hit1=True
        if day<=3 and hi>=targets[1]:hit3=True
        if day<=5 and hi>=targets[2]:hit5=True
    f=q.iloc[i+1:i+6]
    mfe=float(f.high.max()/entry-1);mae=float(f.low.min()/entry-1)
    return hit1,hit3,hit5,stop_first,mfe,mae

def historical_stats(q):
    sig=[];minsc=float(CFG["BACKTEST_MIN_SCORE"]);maxn=int(CFG["BACKTEST_SIGNALS_PER_STOCK"])
    for i in range(1,len(q)-5):
        x=q.iloc[i];p=q.iloc[i-1]
        sc,dist,vr,turn,hd,ap,cl,uw=technical_score(x,p)
        if sc<minsc:continue
        if float(x.close)<float(CFG["MIN_PRICE"]) or turn<float(CFG["MIN_TURNOVER"]):continue
        if cl<float(CFG["MIN_CLOSE_LOCATION"]) or uw>float(CFG["MAX_UPPER_WICK_RATIO"]):continue
        if ap>float(CFG["MAX_ATR_PCT"]):continue
        if strategy(x,p,dist,vr,hd)=="기본":continue
        sig.append(first_touch_outcomes(q,i,float(x.close),float(x.atr14)))
    sig=sig[-maxn:]
    if not sig:return {"bt_n":0,"bt_next2":None,"bt_3d3":None,"bt_5d5":None,"bt_stop_first":None,"bt_mfe5":None,"bt_mae5":None}
    a=np.array(sig,dtype=float)
    return {"bt_n":len(a),"bt_next2":a[:,0].mean(),"bt_3d3":a[:,1].mean(),"bt_5d5":a[:,2].mean(),
            "bt_stop_first":a[:,3].mean(),"bt_mfe5":a[:,4].mean(),"bt_mae5":a[:,5].mean()}

def weighted_bt_score(bt):
    n=int(bt.get("bt_n") or 0)
    if n<int(CFG["MIN_BT_SAMPLE_FOR_BONUS"]):return None
    # 0..100: favor 3-day result, then next-day and 5-day, penalize stop-first.
    raw=100*(0.25*float(bt["bt_next2"])+0.40*float(bt["bt_3d3"])+0.20*float(bt["bt_5d5"])+0.15*(1-float(bt["bt_stop_first"])))
    # shrink small samples toward neutral 50.
    reliability=min(1.0,n/50.0)
    return 50+(raw-50)*reliability

def holding_type(bt):
    n=int(bt.get("bt_n") or 0)
    if n<10:return "검증부족"
    a=float(bt["bt_next2"]);b=float(bt["bt_3d3"]);c=float(bt["bt_5d5"])
    if a>=.50 and a>=b-.05:return "익일형"
    if b>=.55 and b>=c:return "2~3일형"
    if c>=.50:return "3~5일형"
    return "보수관찰"

def recommendation(final_score,conf):
    if final_score>=90 and conf in ("양호","높음"):return "★★★★★ 적극관찰"
    if final_score>=85:return "★★★★☆ 우선관찰"
    if final_score>=78:return "★★★☆☆ 관찰"
    return "★★☆☆☆ 보조후보"

def analyze(item,tk):
    typ=security_type(item.name)
    b={"code":item.code,"name":item.name,"market":item.market,"security_type":typ,"api_ok":False,"bars":0,
       "indicator_ok":False,"final_pass":False,"fail_reason":"","api_message":""}
    if excluded_type(typ):
        b["fail_reason"]="EXCLUDED_"+typ;return b
    rows,msg=history(item.code,tk)
    if not rows:b.update(fail_reason="API_ERROR",api_message=msg);return b
    b["api_ok"]=True;d=frame(rows);b["bars"]=len(d)
    if len(d)<int(CFG["MIN_BARS"]):b["fail_reason"]="BARS_SHORT";return b
    q=enrich(d).dropna().reset_index(drop=True)
    if len(q)<80:b["fail_reason"]="INDICATOR_NA";return b
    b["indicator_ok"]=True;x=q.iloc[-1];p=q.iloc[-2]
    sc,dist,vr,turn,hd,ap,cl,uw=technical_score(x,p);strat=strategy(x,p,dist,vr,hd);gr=grade(sc)
    fails=[]
    if float(x.close)<float(CFG["MIN_PRICE"]):fails.append("PRICE")
    if turn<float(CFG["MIN_TURNOVER"]):fails.append("TURNOVER")
    if vr<float(CFG["MIN_VOLUME_RATIO"]):fails.append("VOLUME")
    if not(float(CFG["RSI_MIN"])<=float(x.rsi14)<=float(CFG["RSI_MAX"])):fails.append("RSI")
    if float(x.adx14)<float(CFG["ADX_MIN"]):fails.append("ADX")
    if ap>float(CFG["MAX_ATR_PCT"]):fails.append("ATR_HIGH")
    if cl<float(CFG["MIN_CLOSE_LOCATION"]):fails.append("CLOSE_WEAK")
    if uw>float(CFG["MAX_UPPER_WICK_RATIO"]):fails.append("UPPER_WICK")
    if float(x.close)<=float(x.ma60):fails.append("BELOW_MA60")
    if float(x.plus_di)<=float(x.minus_di):fails.append("DI")
    if sc<float(CFG["MIN_SCORE"]):fails.append("SCORE")
    if strat=="기본":fails.append("NO_CORE_PATTERN")
    bt=historical_stats(q);conf=bt_conf(int(bt["bt_n"]));bt_score=weighted_bt_score(bt)
    final_score=round(sc if bt_score is None else 0.80*sc+0.20*bt_score,1)
    hold=holding_type(bt);rec=recommendation(final_score,conf)
    entry=float(x.close)
    stop=max(entry*(1-float(CFG["STOP_MAX_PCT"])),min(float(x.ma20)-float(x.atr14)*.5,entry-float(x.atr14)*float(CFG["STOP_ATR"])))
    b.update(date=x.date,close=entry,open=float(x.open),high=float(x.high),low=float(x.low),
             ma5=float(x.ma5),ma20=float(x.ma20),ma60=float(x.ma60),ma20_distance=dist,
             volume_ratio=vr,turnover=turn,rsi14=float(x.rsi14),adx14=float(x.adx14),
             plus_di=float(x.plus_di),minus_di=float(x.minus_di),atr14=float(x.atr14),atr_pct=ap,
             high20=float(x.high20_prev),high20_distance=hd,close_location=cl,upper_wick_ratio=uw,
             strategy=strat,technical_score=sc,grade=gr,grade_rank=grade_rank(gr),
             entry_low=round(entry-max(float(x.atr14)*.20,entry*.005),0),
             entry_high=round(entry+min(float(x.atr14)*.15,entry*.005),0),
             stop=round(stop,0),nextday_target=round(entry*(1+float(CFG["NEXTDAY_TARGET_PCT"])),0),
             day3_target=round(entry*(1+float(CFG["DAY3_TARGET_PCT"])),0),
             day5_target=round(entry*(1+float(CFG["DAY5_TARGET_PCT"])),0),
             nextday_gap_limit=round(entry*(1+float(CFG["MAX_GAP_UP_PCT"])),0),
             bt_confidence=conf,bt_score=round(bt_score,1) if bt_score is not None else None,
             final_recommend_score=final_score,holding_type=hold,recommendation=rec)
    b.update(bt);b["fail_reason"]=" / ".join(fails);b["final_pass"]=not fails
    return b



def fast_prefilter(row,tk):
    """V10.2: 짧은 최근 구간으로 명백한 탈락만 먼저 제거."""
    try:
        # 상품 제외는 기존 analyze()와 동일하게 최종 단계에서도 다시 확인됨.
        e=datetime.now().strftime("%Y%m%d")
        s=(datetime.now()-timedelta(days=150)).strftime("%Y%m%d")
        arr,err=fetch(str(row.code),tk,s,e)
        if err or not arr:
            return row,"API_ERROR"
        d=frame(arr)
        if d is None or len(d)<35:
            return row,"DATA_SHORT"
        d=enrich(d)
        r=d.iloc[-1]
        close=float(r["close"]); vol=float(r["volume"])
        turn=float(close*vol)
        # 1차 필터는 의도적으로 느슨하게: 명백한 저유동/약추세만 제외
        if float(r.get("volume_ratio",1) or 1)<0.35: return row,"FAST_VOLUME"
        if float(r.get("ma20",close) or close)>0 and close<float(r["ma20"])*0.88: return row,"FAST_TREND"
        return row,"PASS"
    except Exception:
        return row,"API_ERROR"

def close_scan():
    started=time.time()
    write_progress(status="starting",message="API 인증 및 종목목록 준비 중",done=0,total=0)
    tk=get_token();u=masters();limit=int(CFG["DIAGNOSTIC_LIMIT"] or 0)
    if limit>0:u=u.head(limit)
    total=len(u)
    print(f"[V10.2 FAST] 검색대상 {total:,}종목",flush=True)

    # 1차 고속필터
    pf_start=time.time(); pf_pass=[]; pf_retry=[]; pf_fail=0; pf_ex=0
    write_progress(status="running",phase="1차 고속필터",message="최근 데이터로 후보 압축 중",
                   done=0,total=total,api_ok=0,api_fail=0,excluded=0,indicator_ok=0,pass_count=0,elapsed=0,eta=0,reasons={})
    # API 안정성을 위해 1차는 낮은 동시성 사용
    with ThreadPoolExecutor(max_workers=2) as pex:
        pfs=[pex.submit(fast_prefilter,r,tk) for r in u.itertuples(index=False)]
        for i,f in enumerate(as_completed(pfs),1):
            r,reason=f.result()
            if reason=="PASS": pf_pass.append(r)
            elif reason=="API_ERROR":
                pf_fail+=1
                pf_retry.append(r)
            else: pf_ex+=1
            if i%10==0 or i==total:
                elapsed=int(time.time()-pf_start); eta=int((elapsed/max(i,1))*(total-i))
                write_progress(status="running",phase="1차 고속필터",message="최근 데이터로 후보 압축 중",
                    done=i,total=total,universe_total=total,phase_done=i,phase_total=total,overall_pct=int(i*50/max(total,1)),
                    api_ok=i-pf_fail,api_fail=pf_fail,excluded=pf_ex,indicator_ok=0,
                    pass_count=len(pf_pass),elapsed=elapsed,eta=eta,
                    reasons={"1차통과":len(pf_pass),"1차제외":pf_ex,"API오류":pf_fail})

    # API 오류 종목은 누락시키지 않고 기존 정밀분석에서 재시도
    # 정상 통과 종목을 우선 정밀분석한다.
    # 1차 API 오류 종목도 2차에서 반드시 재시도하여 누락을 방지한다.
    phase2_rows=pf_pass+pf_retry
    u2=pd.DataFrame([r._asdict() for r in phase2_rows]) if phase2_rows else u.head(0).copy()
    if u2.empty:
        u2=u.copy()
    universe_total=total
    total=len(u2)
    phase2_total=total
    phase2_start=time.time()
    write_progress(status="running",phase="2차 정밀분석",message="1차 통과 종목 2년 분석/백테스트 중",
                   done=0,total=total,universe_total=universe_total,phase_done=0,phase_total=phase2_total,overall_pct=50,
                   api_ok=0,api_fail=0,excluded=pf_ex,indicator_ok=0,pass_count=0,elapsed=0,eta=0,
                   reasons={"1차통과":len(pf_pass),"1차재시도":len(pf_retry),"1차제외":pf_ex})
    rows=[];ok=fail=ind=pas=excluded=0
    with ThreadPoolExecutor(max_workers=max(1,min(int(CFG.get("MAX_WORKERS",4) or 4),6))) as ex:
        fs=[ex.submit(analyze,r,tk) for r in u2.itertuples(index=False)]
        for i,f in enumerate(as_completed(fs),1):
            try:
                z=f.result();rows.append(z);excluded+=int(str(z.get("fail_reason","")).startswith("EXCLUDED_"))
                ok+=int(z["api_ok"]);fail+=int((not z["api_ok"]) and not str(z.get("fail_reason","")).startswith("EXCLUDED_"))
                ind+=int(z["indicator_ok"]);pas+=int(z["final_pass"])
            except Exception:
                fail+=1;elog(traceback.format_exc())
            if i%10==0 or i==total:
                elapsed=int(time.time()-phase2_start)
                eta=int((elapsed/max(i,1))*(total-i)) if i else 0
                reason_counts={}
                for rr in rows:
                    fr=str(rr.get("fail_reason","") or "")
                    if not fr and rr.get("final_pass"): key="PASS"
                    elif fr.startswith("EXCLUDED_"): key="상품제외"
                    elif fr=="API_ERROR": key="API오류"
                    elif fr: key=fr.split(" / ")[0]
                    else: key="기타"
                    reason_counts[key]=reason_counts.get(key,0)+1
                top_reasons=dict(sorted(reason_counts.items(),key=lambda x:x[1],reverse=True)[:8])
                print(f"{i:,}/{total:,} | API성공 {ok:,} 실패 {fail:,} | 제외 {excluded:,} | 지표 {ind:,} | PASS {pas:,} | ETA {eta//60:02d}:{eta%60:02d}",flush=True)
                overall_pct=50+int(i*50/max(phase2_total,1))
                write_progress(status="running",phase="2차 정밀분석",message="1차 통과/재시도 종목 2년 분석·백테스트 중",
                    done=i,total=phase2_total,universe_total=universe_total,phase_done=i,phase_total=phase2_total,overall_pct=min(99,overall_pct),
                    api_ok=ok,api_fail=fail,excluded=pf_ex+excluded,indicator_ok=ind,pass_count=pas,elapsed=elapsed,eta=eta,reasons=top_reasons)
    allx=pd.DataFrame(rows); cand=allx[allx.indicator_ok==True].copy() if not allx.empty else pd.DataFrame()
    if not cand.empty:cand=cand.sort_values(["final_pass","grade_rank","final_recommend_score","technical_score","turnover"],ascending=[False,True,False,False,False])
    # V10: 시장 전체 과거표본을 이용한 2차 검증.
    # 개별 종목의 표본이 적어도 같은 전략군의 시장 전체 과거 사례를 함께 보여준다.
    if not allx.empty and "strategy" in allx.columns:
        valid_bt=allx[(allx.get("bt_n",0).fillna(0)>0) & allx["strategy"].notna()].copy()
        pooled={}
        for strat,g in valid_bt.groupby("strategy"):
            n=float(g["bt_n"].fillna(0).sum())
            if n<=0: continue
            def wavg(col):
                x=g[[col,"bt_n"]].dropna()
                return float((x[col]*x["bt_n"]).sum()/x["bt_n"].sum()) if len(x) and x["bt_n"].sum()>0 else None
            pooled[strat]={"시장표본":int(n),"시장익일2":wavg("bt_next2"),"시장3일3":wavg("bt_3d3"),
                           "시장5일5":wavg("bt_5d5"),"시장손절":wavg("bt_stop_first")}
        def poolval(r,key):
            return pooled.get(r.get("strategy"),{}).get(key)
        allx["market_bt_n"]=allx.apply(lambda r:poolval(r,"시장표본"),axis=1)
        allx["market_bt_next2"]=allx.apply(lambda r:poolval(r,"시장익일2"),axis=1)
        allx["market_bt_3d3"]=allx.apply(lambda r:poolval(r,"시장3일3"),axis=1)
        allx["market_bt_5d5"]=allx.apply(lambda r:poolval(r,"시장5일5"),axis=1)
        allx["market_bt_stop"]=allx.apply(lambda r:poolval(r,"시장손절"),axis=1)
        allx["validation_label"]=allx.apply(
            lambda r:"강화검증" if (r.get("bt_n") or 0)>=30 else
                     ("시장표본보완" if (r.get("market_bt_n") or 0)>=100 else "검증부족"),axis=1)
    top=allx[allx.final_pass==True].copy() if not allx.empty else pd.DataFrame()
    if not top.empty:
        top=top.sort_values(["grade_rank","final_recommend_score","technical_score","turnover"],ascending=[True,False,False,False]).head(int(CFG["TOP_N"])).reset_index(drop=True)
        top.insert(0,"rank",range(1,len(top)+1))
    topmap={"rank":"순위","name":"종목명","code":"종목코드","market":"시장","grade":"등급","final_recommend_score":"최종추천점수",
    "recommendation":"추천도","holding_type":"추천보유유형","strategy":"선정전략","close":"종가","entry_low":"매수구간하단",
    "entry_high":"매수구간상단","stop":"손절가","nextday_target":"익일목표가","day3_target":"3일목표가","day5_target":"5일목표가",
    "nextday_gap_limit":"익일추격금지가","bt_confidence":"백테스트신뢰도","bt_n":"백테스트표본","bt_next2":"익일+2%선도달률",
    "bt_3d3":"3일+3%선도달률","bt_5d5":"5일+5%선도달률","bt_stop_first":"손절선도달률",
    "validation_label":"검증등급","market_bt_n":"시장전략표본","market_bt_next2":"시장익일+2%","market_bt_3d3":"시장3일+3%",
    "market_bt_5d5":"시장5일+5%","market_bt_stop":"시장손절률"}
    detailmap={"name":"종목명","code":"종목코드","market":"시장","date":"기준일","technical_score":"기술점수","final_recommend_score":"최종추천점수",
    "grade":"등급","strategy":"선정전략","close":"종가","ma5":"5일선","ma20":"20일선","ma60":"60일선","ma20_distance":"20일선이격률",
    "volume_ratio":"거래량배수","turnover":"거래대금(원)","rsi14":"RSI14","adx14":"ADX14","plus_di":"+DI","minus_di":"-DI","atr14":"ATR14",
    "atr_pct":"ATR비율","high20":"최근20일고점","high20_distance":"전고점거리","close_location":"종가강도","upper_wick_ratio":"윗꼬리비율",
    "bt_confidence":"백테스트신뢰도","bt_n":"백테스트표본","bt_next2":"익일+2%선도달률","bt_3d3":"3일+3%선도달률","bt_5d5":"5일+5%선도달률",
    "bt_stop_first":"손절선도달률","validation_label":"검증등급","market_bt_n":"시장전략표본",
    "market_bt_next2":"시장익일+2%","market_bt_3d3":"시장3일+3%","market_bt_5d5":"시장5일+5%","market_bt_stop":"시장손절률",
    "bt_mfe5":"5일최대상승폭","bt_mae5":"5일최대하락폭","holding_type":"추천보유유형","recommendation":"추천도"}
    topkr=top[[c for c in topmap if c in top.columns]].rename(columns=topmap) if not top.empty else pd.DataFrame(columns=list(topmap.values()))
    det=cand[[c for c in detailmap if c in cand.columns]].rename(columns=detailmap) if not cand.empty else pd.DataFrame(columns=list(detailmap.values()))
    summary=pd.DataFrame({"항목":["검색대상","상품제외","API성공","API실패","지표계산성공","최종후보","TOP출력"],"수량":[len(u),excluded,ok,fail,ind,pas,len(top)]})
    reasons=allx.fail_reason.fillna("").replace("","PASS").value_counts().reset_index() if not allx.empty else pd.DataFrame(columns=["탈락사유","건수"])
    if len(reasons.columns)==2:reasons.columns=["탈락사유","건수"]
    mode_tag=os.getenv("SCANNER_MODE","CLOSE").upper(); path=RESULTS/f"Guru_Stock_Scanner_V10_{mode_tag}_{STAMP}.xlsx"
    with pd.ExcelWriter(path,engine="openpyxl") as w:
        topkr.to_excel(w,index=False,sheet_name="TOP10_한눈에보기");det.to_excel(w,index=False,sheet_name="상세분석")
        summary.to_excel(w,index=False,sheet_name="실행요약");reasons.to_excel(w,index=False,sheet_name="탈락사유")
        pd.DataFrame([CFG]).to_excel(w,index=False,sheet_name="적용조건")
    from openpyxl import load_workbook as LW
    from openpyxl.styles import Font,PatternFill,Alignment
    from openpyxl.utils import get_column_letter
    b=LW(path)
    for sn in ["TOP10_한눈에보기","상세분석"]:
        ws=b[sn];ws.freeze_panes="A2";ws.auto_filter.ref=ws.dimensions;ws.sheet_view.showGridLines=False
        for cell in ws[1]:
            cell.fill=PatternFill("solid",fgColor="17365D");cell.font=Font(color="FFFFFF",bold=True);cell.alignment=Alignment(horizontal="center",wrap_text=True)
        for i in range(1,ws.max_column+1):ws.column_dimensions[get_column_letter(i)].width=16
    ws=b["TOP10_한눈에보기"];ws.column_dimensions["B"].width=18;ws.column_dimensions["G"].width=22;ws.column_dimensions["I"].width=30
    hdr={c.value:c.column for c in ws[1]}
    for h in ["익일+2%선도달률","3일+3%선도달률","5일+5%선도달률","손절선도달률","시장익일+2%","시장3일+3%","시장5일+5%","시장손절률"]:
        if h in hdr:
            for r in range(2,ws.max_row+1):ws.cell(r,hdr[h]).number_format="0.0%"
    if "등급" in hdr:
        for r in range(2,ws.max_row+1):
            v=ws.cell(r,hdr["등급"]).value
            ws.cell(r,hdr["등급"]).fill=PatternFill("solid",fgColor="E2F0D9" if v=="A+" else "D9EAF7" if v=="A" else "FFF2CC")
    if "백테스트신뢰도" in hdr:
        for r in range(2,ws.max_row+1):
            if ws.cell(r,hdr["백테스트신뢰도"]).value=="낮음":ws.cell(r,hdr["백테스트신뢰도"]).fill=PatternFill("solid",fgColor="F4CCCC")
    if "손절선도달률" in hdr:
        for r in range(2,ws.max_row+1):
            v=ws.cell(r,hdr["손절선도달률"]).value
            if isinstance(v,(int,float)) and v>=.5:ws.cell(r,hdr["손절선도달률"]).fill=PatternFill("solid",fgColor="F4CCCC")
    b.active=0;b.save(path)
    write_progress(status="done",message="검색 완료",done=total,total=total,api_ok=ok,api_fail=fail,excluded=excluded,indicator_ok=ind,pass_count=pas,top_count=len(top),elapsed=int(time.time()-started),result=str(path.name))
    print("완료:",path,flush=True)



def current_price(code, tk):
    """KIS 국내주식 현재가 REST. 애프터마켓 운영시간에는 API가 제공하는 최신 체결/현재가를 사용."""
    h={"authorization":"Bearer "+tk,"appkey":KEY,"appsecret":SECRET,"tr_id":"FHKST01010100","custtype":"P"}
    p={"FID_COND_MRKT_DIV_CODE":"J","FID_INPUT_ISCD":str(code).zfill(6)}
    try:
        api_wait()
        r=requests.get(BASE+"/uapi/domestic-stock/v1/quotations/inquire-price",headers=h,params=p,timeout=20)
        if r.status_code!=200:return None,f"HTTP {r.status_code}"
        j=r.json()
        if str(j.get("rt_cd"))!="0":return None,f'{j.get("msg_cd","")} {j.get("msg1","")}'
        o=j.get("output") or {}
        price=o.get("stck_prpr")
        if price in (None,"","0"):return None,"현재가 없음"
        return float(price),""
    except Exception as e:return None,repr(e)

def latest_close_result():
    files=sorted(RESULTS.glob("Guru_Stock_Scanner_V10_*CLOSE*.xlsx"),key=lambda p:p.stat().st_mtime,reverse=True)
    if not files:
        # compatibility: close_scan filename may not contain CLOSE in older runs
        files=sorted(RESULTS.glob("Guru_Stock_Scanner_V10_MULTI_*.xlsx"),key=lambda p:p.stat().st_mtime,reverse=True)
    return files[0] if files else None

def after_scan():
    src=latest_close_result()
    if src is None:
        print("애프터마켓 확인용 장마감 결과가 없습니다. 먼저 ②_장마감_확정검색.bat을 실행하세요.")
        return
    try:
        top=pd.read_excel(src,sheet_name="TOP10_한눈에보기",dtype={"종목코드":str})
    except Exception as e:
        print("TOP10 읽기 실패:",e);return
    if top.empty:
        print("장마감 TOP10 후보가 없습니다.");return
    top=top.head(int(CFG.get("AFTER_TOP_N",10))).copy()
    tk=get_token();rows=[]
    for _,r in top.iterrows():
        code=str(r["종목코드"]).replace(".0","").zfill(6);close=float(r["종가"])
        cur,msg=current_price(code,tk)
        chg=(cur/close-1) if cur else None
        if cur is None:judge="API확인필요"
        elif chg>=float(CFG.get("AFTER_CHASE_PCT",0.04)):judge="추격금지"
        elif chg<=float(CFG.get("AFTER_DROP_WARN_PCT",-0.03)):judge="급락주의"
        elif float(CFG.get("AFTER_BUY_LOW_PCT",-0.02))<=chg<=float(CFG.get("AFTER_BUY_HIGH_PCT",0.02)):judge="매수관찰"
        elif chg>float(CFG.get("AFTER_BUY_HIGH_PCT",0.02)):judge="상승주의"
        else:judge="눌림관찰"
        rows.append({"순위":r.get("순위"),"종목명":r.get("종목명"),"종목코드":code,"등급":r.get("등급"),
                     "최종추천점수":r.get("최종추천점수"),"정규장종가":close,"애프터현재가":cur,
                     "종가대비등락률":chg,"애프터판정":judge,"추천보유유형":r.get("추천보유유형"),
                     "선정전략":r.get("선정전략"),"API메시지":msg})
    out=pd.DataFrame(rows)
    path=RESULTS/f"Guru_Stock_Scanner_V10_AFTER_{STAMP}.xlsx"
    with pd.ExcelWriter(path,engine="openpyxl") as w:
        out.to_excel(w,index=False,sheet_name="애프터확인")
        top.to_excel(w,index=False,sheet_name="장마감TOP10")
        pd.DataFrame([CFG]).to_excel(w,index=False,sheet_name="적용조건")
    from openpyxl import load_workbook as LW
    from openpyxl.styles import Font,PatternFill,Alignment
    from openpyxl.utils import get_column_letter
    b=LW(path);ws=b["애프터확인"];ws.freeze_panes="A2";ws.auto_filter.ref=ws.dimensions;ws.sheet_view.showGridLines=False
    for c in ws[1]:
        c.fill=PatternFill("solid",fgColor="17365D");c.font=Font(color="FFFFFF",bold=True);c.alignment=Alignment(horizontal="center",wrap_text=True)
    for i in range(1,ws.max_column+1):ws.column_dimensions[get_column_letter(i)].width=17
    ws.column_dimensions["B"].width=18;ws.column_dimensions["I"].width=16;ws.column_dimensions["K"].width=30
    hdr={c.value:c.column for c in ws[1]}
    if "종가대비등락률" in hdr:
        for rr in range(2,ws.max_row+1):ws.cell(rr,hdr["종가대비등락률"]).number_format="0.00%"
    if "애프터판정" in hdr:
        colors={"매수관찰":"E2F0D9","추격금지":"F4CCCC","급락주의":"F4CCCC","상승주의":"FFF2CC","눌림관찰":"D9EAF7"}
        for rr in range(2,ws.max_row+1):
            v=ws.cell(rr,hdr["애프터판정"]).value
            if v in colors:ws.cell(rr,hdr["애프터판정"]).fill=PatternFill("solid",fgColor=colors[v])
    b.active=0;b.save(path)
    print("애프터 확인 완료:",path)

def intraday_scan():
    # Uses the same technical engine against the latest day data available from KIS.
    # It is intentionally labeled preliminary because the regular-session candle is not finalized.
    print("장중 종가매수 예비검색: 당일 진행 중 데이터 기준이므로 16:00 확정검색으로 재검증하세요.")
    close_scan()

def main():
    mode=os.getenv("SCANNER_MODE",str(CFG.get("MODE","CLOSE"))).upper()
    if mode=="AFTER":after_scan()
    elif mode=="INTRADAY":intraday_scan()
    else:close_scan()




if __name__=="__main__":
    try:
        main()
    except Exception as e:
        write_progress(status="error",message=str(e),traceback=traceback.format_exc())
        raise
