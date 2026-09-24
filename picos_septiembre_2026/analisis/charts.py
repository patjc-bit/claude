exec(open('load.py').read())
import matplotlib; matplotlib.use('Agg')
import matplotlib.pyplot as plt, matplotlib.ticker, matplotlib.dates as mdates
from matplotlib.colors import LinearSegmentedColormap, TwoSlopeNorm
SURF='#fcfcfb'; INK='#0b0b0b'; INK2='#52514e'; MUTED='#898781'; GRID='#e1e0d9'; AXIS='#c3c2b7'
C1,C2,C3='#2a78d6','#eb6834','#1baf7a'
plt.rcParams.update({'font.family':'Liberation Sans','font.size':9,'axes.edgecolor':AXIS,'axes.labelcolor':INK2,
 'xtick.color':MUTED,'ytick.color':MUTED,'axes.facecolor':SURF,'figure.facecolor':SURF,'axes.titlecolor':INK,
 'axes.spines.top':False,'axes.spines.right':False,'axes.grid':True,'grid.color':GRID,'grid.linewidth':0.6})

# ---------- Fig 1: small multiples, indexed to own max (hours 7-22)
H=list(range(7,23))
fig,axs=plt.subplots(2,3,figsize=(9.2,5.4),sharex=True,sharey=True)
for ax,b in zip(axs.flat,B):
    e=ent_h[b].loc[H]; m=mos_h[b].loc[H]; u=uni_h[b].loc[H]
    ax.axvspan(14,16,color='#f0efec',zorder=0,lw=0)
    ax.plot(H,e/e.max()*100,color=C1,lw=2,label='Entradas',solid_capstyle='round')
    ax.plot(H,m/m.max()*100,color=C2,lw=2,ls=(0,(5,2)),label='Mostrador desatendido')
    ax.plot(H,u/u.max()*100,color=C3,lw=2,ls=(0,(1.2,1.6)),label='Unifila')
    for s,c in ((e,C1),(m,C2),(u,C3)):
        h=s.idxmax(); ax.plot([h],[100],'o',ms=5.5,color=c,mec=SURF,mew=1.5,zorder=5)
    grp='anterior' if b in ['TAPO','ZIBATA','CHAPALITA'] else 'nueva (desde 2-sep)'
    ax.set_title(f"{b}  ·  {grp}",fontsize=9.5,loc='left',fontweight='bold')
    ax.set_ylim(0,108); ax.set_xticks([7,10,13,16,19,22]); ax.set_xticklabels(['7h','10h','13h','16h','19h','22h'])
for ax in axs[:,0]: ax.set_ylabel('% del máximo propio')
h,l=axs[0,0].get_legend_handles_labels()
fig.legend(h,l,loc='upper left',ncol=3,frameon=False,bbox_to_anchor=(0.01,0.985),fontsize=9,labelcolor=INK2)
fig.text(0.99,0.965,'Franja gris = 14:00–16:00 (hora de comida)',ha='right',color=MUTED,fontsize=8)
fig.tight_layout(rect=(0,0,1,0.93))
fig.savefig('fig1_perfil_horario.png',dpi=200); plt.close(fig)

# ---------- Fig 2: diverging heatmap, unifila alerts per entry, indexed to branch's own 9-20h mean
H2=list(range(8,22))
bb=['TAPO','ZIBATA','AVCONS','SANJACINTO','XOCHIMILCO','CHAPALITA']
r=(uni_h/ent_h).loc[H2,bb]
idx=(r/r.loc[9:20].mean())
cmap=LinearSegmentedColormap.from_list('div',['#1c5cab','#86b6ef','#f0efec','#f0a3a2','#c23433'])
fig,ax=plt.subplots(figsize=(9.2,3.3))
im=ax.imshow(idx.T.values,aspect='auto',cmap=cmap,norm=TwoSlopeNorm(1.0,0.5,1.5))
ax.grid(False)
ax.set_xticks(range(len(H2))); ax.set_xticklabels([f'{h}h' for h in H2])
ax.set_yticks(range(len(bb))); ax.set_yticklabels([b+(' *' if b=='CHAPALITA' else '') for b in bb],color=INK)
for i,b in enumerate(bb):
    for j,h in enumerate(H2):
        v=idx.loc[h,b]; ax.text(j,i,f'{v:.2f}',ha='center',va='center',fontsize=7.2,color=INK if 0.72<v<1.3 else '#ffffff')
for s in ax.spines.values(): s.set_visible(False)
cb=fig.colorbar(im,ax=ax,fraction=0.025,pad=0.01); cb.outline.set_visible(False); cb.ax.tick_params(labelsize=7.5,colors=MUTED)
cb.set_label('1.0 = promedio de la sucursal',fontsize=7.5,color=INK2)
ax.set_title('Alertas de unifila por cliente que entra (índice vs. promedio 9–20h de cada sucursal)',loc='left',fontsize=9.5,fontweight='bold')
fig.text(0.01,0.01,'* CHAPALITA: el contador de entradas parece subcontar (ver sección de calidad de datos); leer con cautela.',fontsize=7.5,color=MUTED)
fig.tight_layout(rect=(0,0.04,1,1))
fig.savefig('fig2_unifila_por_cliente.png',dpi=200); plt.close(fig)

# ---------- Fig 3: daily entries old branches + 7d mean, with event markers
d=daily.replace(0,np.nan).copy(); d.loc['2026-09-15']=np.nan; d.loc['2026-07-08','TAPO']=np.nan; d.loc['2026-07-06','CHAPALITA']=np.nan
fig,ax=plt.subplots(figsize=(9.2,3.9))
for b,c in (('ZIBATA',C1),('TAPO',C2),('CHAPALITA',C3)):
    s=d[b]; ax.plot(s.index,s,color=c,lw=0.8,alpha=0.35)
    rm=s.rolling(7,center=True,min_periods=5).mean()
    ax.plot(rm.index,rm,color=c,lw=2.2,label=f'{b} (media móvil 7 días)')
    last=rm.dropna(); ax.text(last.index[-1]+pd.Timedelta(days=0.8),last.iloc[-1],b,color=INK2,va='center',fontsize=8.5,fontweight='bold')
ev=[('2026-07-16','Inician vacaciones SEP'),('2026-08-03','¿Inicio semestre\nuniversitario? (hipótesis)'),('2026-08-31','Regreso a clases SEP'),('2026-09-16','16-sep\n(feriado)')]
for dt,lab in ev:
    x=pd.Timestamp(dt); ax.axvline(x,color=MUTED,lw=0.8,ls=(0,(3,3)))
    ax.text(x+pd.Timedelta(days=0.6),1175,lab,fontsize=7.5,color=INK2,va='top')
ax.set_ylim(0,1200); ax.set_xlim(pd.Timestamp('2026-07-01'),pd.Timestamp('2026-09-21'))
ax.xaxis.set_major_locator(mdates.WeekdayLocator(byweekday=0,interval=2)); MES={7:'jul',8:'ago',9:'sep'}
ax.xaxis.set_major_formatter(matplotlib.ticker.FuncFormatter(lambda x,p: f"{mdates.num2date(x).day}-{MES[mdates.num2date(x).month]}"))
ax.set_ylabel('Entradas por día')
ax.legend(loc='lower left',bbox_to_anchor=(0,1.0),frameon=False,fontsize=8,ncol=3,labelcolor=INK2)
ax.set_title('Entradas diarias — sucursales anteriores (línea tenue = dato diario)',loc='left',fontsize=9.5,fontweight='bold',pad=22)
fig.tight_layout(); fig.savefig('fig3_diario.png',dpi=200); plt.close(fig)
print('ok')
