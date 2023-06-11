import pandas as pd;
import requests;
from time import sleep; 
from datetime import datetime as dt,timedelta;
import xlwings as xw
import numpy as np

s= requests.session();opt_chain={'time':[''],'coi':[''],'poi':[''],'sum':[''],'tc':[''],'tp':[''],'bcoi':[''],'bpoi':[''],'bsum':[''],'btc':[''],'btp':[''],'cvol':[''],'pvol':[''],'vsum':[''],'cpr':[''],'ppr':['']};
nextchain={'time':[''],'coi':[''],'poi':[''],'sum':[''],'tc':[''],'tp':[''],'bcoi':[''],'bpoi':[''],'bsum':[''],'btc':[''],'btp':[''],'fcoi':[''],'fpoi':[''],'fsum':[''],'ftc':[''],'ftp':[''],'cpr':[''],'ppr':['']};
emp=emp2=pd.DataFrame();empty=pd.read_csv('empty.csv')

def opt_nse():
    
    df=pd.DataFrame(opt_chain);df2=pd.DataFrame(nextchain);  file=xw.Book('book.xlsx');curr=pd.read_csv('optchain.csv');
    file.sheets('sheet1').range('a1').value=empty;empty['ftp']=empty['ftc']='';file.sheets('sheet2').range('a1').value=empty
    cum=pd.read_csv('optnext.csv') ;global emp,emp2;nfurl='https://www.nseindia.com/api/option-chain-indices?symbol=NIFTY' # a='https://www.nseindia.com/api/quote-equity?symbol=SBIN'
    s= requests.session();    URL = 'https://www.nseindia.com/';close=0#print('url done1') 
    s.headers.update({'Accept': '/','Accept-Encoding':'gzip, deflate, br', 'Accept-Language': 'en-US,en;q=0.9', 'Cache-Control': 'max-age=0', 'Connection': 'keep-alive',
    'Sec-Fetch-Dest': 'empty', 'Sec-Fetch-Mode': 'cors', 'Sec-Fetch-Site': 'none', 'Sec-Fetch-User': '?1', 'Upgrade-Insecure-Requests': '1',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.116 Safari/537.36'
    });r1 = s.get(URL);print('url done')  # nf = s.get("https://www1.nseindia.com/live_market/dynaContent/live_watch/option_chain/optionKeys.jsp?symbolCode=-10006&symbol=NIFTY&symbol=NIFTY&instrument=-&date=-&segmentLink=17&symbolCount=2&segmentLink=17") # bnf = s.get("https://www.nseindia.com/api/option-chain-indices?symbol=BANKNIFTY").json();#print(nf)
    nf = s.get(nfurl).json();fs = s.get("https://www.nseindia.com/api/option-chain-indices?symbol=FINNIFTY").json();print('done')
    nf = nf['records']['data'];nf = pd.DataFrame(pd.json_normalize(nf));a=nf[['expiryDate']];a['exp']=a['expiryDate'];a.drop_duplicates(subset=['expiryDate'],inplace=True);
    a['expiryDate'] = pd.to_datetime(a['expiryDate']);a.sort_values(by='expiryDate', ascending=True,inplace=True);
    exp=a.iloc[0][1];next=a.iloc[1][1];
    fs = fs['records']['data'];fs = pd.DataFrame(pd.json_normalize(fs));a=fs[['expiryDate']];a['exp']=a['expiryDate'];a.drop_duplicates(subset=['expiryDate'],inplace=True);
    a['expiryDate'] = pd.to_datetime(a['expiryDate']);a.sort_values(by='expiryDate', ascending=True,inplace=True);
    fexp=a.iloc[0][1];
    print('Datetime is:',exp,next,fexp)
    
    for i in range(500):
        print(s.get(URL).status_code);nf = s.get(nfurl).json();#print('done');#r1 = s.get(URL);#print('responce',r1)#print(nf.status_code)print('responce');
        bnf = s.get("https://www.nseindia.com/api/option-chain-indices?symbol=BANKNIFTY").json();
        fs = s.get("https://www.nseindia.com/api/option-chain-indices?symbol=FINNIFTY").json()
        fs = fs['records']['data']; fs = pd.DataFrame(pd.json_normalize(fs)); fs=fs[fs['expiryDate']==fexp] ;#fs.to_csv('fs.csv');break
        fs=fs[['strikePrice','PE.openInterest','PE.changeinOpenInterest','CE.openInterest','CE.changeinOpenInterest','PE.change','CE.change']]
        cvol = nf['filtered']['CE']['totVol']/10**5;pvol=nf['filtered']['PE']['totVol']/10**5;#print(cvol)
        nf = nf['records']['data'];bnf= bnf['records']['data'];bnf= pd.DataFrame(pd.json_normalize(bnf));bnf2=bnf[bnf['expiryDate']==next] 
        bnf=bnf[bnf['expiryDate']==exp] ;  nf = pd.DataFrame(pd.json_normalize(nf));#nf.to_csv('nf.csv')
        nf2=nf[nf['expiryDate']==next] ; nf=nf[nf['expiryDate']==exp] 
        bnf=bnf[['strikePrice','PE.openInterest','PE.changeinOpenInterest','CE.openInterest','CE.changeinOpenInterest']]
        bpoi=bnf['PE.changeinOpenInterest'].sum()/10**3;bcoi=bnf['CE.changeinOpenInterest'].sum()/10**3
        btp=bnf['PE.openInterest'].sum()/10**4;btc=bnf['CE.openInterest'].sum()/10**4# df=pd.DataFrame(opt_chain);        # r = r['allSec']['data']
        nf=nf[['strikePrice','PE.openInterest','PE.changeinOpenInterest','CE.openInterest','CE.changeinOpenInterest','PE.change','CE.change']];#print(nf)
        poi=nf['PE.changeinOpenInterest'].sum()/10**3;coi=nf['CE.changeinOpenInterest'].sum()/10**3
        tp=nf['PE.openInterest'].sum()/10**4;tc=nf['CE.openInterest'].sum()/10**4
        now2=dt.now();now=now2.strftime("%H:%M");b=int(now.split(':')[0])*60+int(now.split(':')[1]);
        fpoi=fs['PE.changeinOpenInterest'].sum()/10**3;fcoi=fs['CE.changeinOpenInterest'].sum()/10**3
        ftp=fs['PE.openInterest'].sum()/10**4;ftc=fs['CE.openInterest'].sum()/10**4# if coi<0:coi=("{:04d}".format(int(coi)));   # else:coi=("{:03d}".format(int(poi)))
        bnf2=bnf2[['strikePrice','PE.openInterest','PE.changeinOpenInterest','CE.openInterest','CE.changeinOpenInterest']]
        bpoi2=bnf2['PE.changeinOpenInterest'].sum()/10**3;bcoi2=bnf2['CE.changeinOpenInterest'].sum()/10**3
        btp2=bnf2['PE.openInterest'].sum()/10**4;btc2=bnf2['CE.openInterest'].sum()/10**4# df=pd.DataFrame(opt_chain);    # r = r['allSec']['data']
        nf2=nf2[['strikePrice','PE.openInterest','PE.changeinOpenInterest','CE.openInterest','CE.changeinOpenInterest','PE.change','CE.change']]
        poi2=nf2['PE.changeinOpenInterest'].sum()/10**3;coi2=nf2['CE.changeinOpenInterest'].sum()/10**3
        tp2=nf2['PE.openInterest'].sum()/10**4;tc2=nf2['CE.openInterest'].sum()/10**4
        coi2=coi+coi2;poi2=poi+poi2;bcoi2=bcoi+bcoi2;bpoi2=bpoi+bpoi2;
        df.iloc[0,0]=now; df.iloc[0,1]=int(coi);df.iloc[0,2]=int(poi);df.iloc[0,3]=round(poi-coi);
        df.iloc[0,4]=round(tc); df.iloc[0,5]=round(tp); df.iloc[0,6]=int(bcoi);df.iloc[0,7]=((int(bpoi)));df.iloc[0,8]=round(bpoi-bcoi);
        df.iloc[0,9]=round(btc);df.iloc[0,10]=round(btp);#all values in lacs
        df.iloc[0,11]=round(cvol);df.iloc[0,12]=round(pvol);df.iloc[0,13]=round(pvol-cvol);df.iloc[0,14]=round(nf['CE.change'].sum())
        df.iloc[0,15]=round(nf['PE.change'].sum())
        df.to_csv('optchain.csv',index=False,mode='a',header=False);
        curr=curr.append(df);
        file.sheets('sheet1').range('a1').value=curr
        df2.iloc[0,0]=now; df2.iloc[0,1]=int(coi2);df2.iloc[0,2]=int(poi2);df2.iloc[0,3]=round(poi2-coi2);
        df2.iloc[0,4]=round(tc2); df2.iloc[0,5]=round(tp2)
        df2.iloc[0,6]=int(bcoi2);df2.iloc[0,7]=int(bpoi2);df2.iloc[0,8]=round(bpoi2-bcoi2);
        df2.iloc[0,9]=round(btc2); df2.iloc[0,10]=round(btp2);#all values in lacs
        df2.iloc[0,11]=int(fcoi);df2.iloc[0,12]=int(fpoi);df2.iloc[0,13]=round(fpoi-fcoi);
        df2.iloc[0,14]=round(ftc); df2.iloc[0,15]=round(ftp);df2.iloc[0,16]=round(nf2['CE.change'].sum())
        df2.iloc[0,17]=round(nf2['PE.change'].sum())
        df2.to_csv('optnext.csv',index=False,mode='a',header=False)
        
        cum=cum.append(df2); file.sheets('sheet2').range('a1').value=cum;print(now2.strftime("%H:%M:%S")) ;
        if ((b>= 927)|(b<= 555))  :
            print("time's up sleep till 7pm", b);#break;
            if b<=931:
                for j in range(50):
                    sleep(300);n=dt.now().strftime("%H:%M");b=int(n.split(':')[0])*60+int(n.split(':')[1]);#break;
                    if b>=1140:close=1;break
        elif (close==0):close=0
        else: print("time's up ",b);break;
        sleep(10)
opt_nse()