import pandas as pd, numpy as np
X='metricas.xlsx'
B=['TAPO','ZIBATA','CHAPALITA','AVCONS','SANJACINTO','XOCHIMILCO']
h1=pd.read_excel(X,sheet_name='Hoja1',header=None)
h2=pd.read_excel(X,sheet_name='Hoja2',header=None)
def block(df,r0,r1):
    t=df.iloc[r0:r1,0:7].copy(); t.columns=['k']+B; t=t.set_index('k')
    return t.replace('-',np.nan).astype(float)
ent_h=block(h1,38,62)   # rows 39-62 excel -> idx 38..61
ent_h.index=range(24)
daily=h1.iloc[65:142,0:7].copy(); daily.columns=['fecha']+B; daily=daily.set_index('fecha'); daily.index=pd.to_datetime(daily.index); daily=daily.astype(float)
mos_h=block(h2,44,68); mos_h.index=range(24)
uni_h=block(h2,72,96); uni_h.index=range(24)
