# Segmentation + features, mirrored exactly in JS (digits.js)
import numpy as np
N=20
def otsu(g):
    h=np.bincount(g.ravel(),minlength=256).astype(np.float64); tot=g.size
    sumall=(np.arange(256)*h).sum(); sb=0; wb=0; best=0; thr=127
    for t in range(256):
        wb+=h[t]
        if wb==0: continue
        wf=tot-wb
        if wf==0: break
        sb+=t*h[t]; mb=sb/wb; mf=(sumall-sb)/wf
        v=wb*wf*(mb-mf)**2
        if v>best: best=v; thr=t
    return thr
from scipy import ndimage
def label(ink):
    lab,n=ndimage.label(ink,structure=np.ones((3,3)))
    comps=[]
    areas=ndimage.sum(np.ones_like(lab),lab,index=np.arange(1,n+1)) if n else []
    for k,sl in enumerate(ndimage.find_objects(lab)):
        comps.append({'id':k+1,'y0':sl[0].start,'y1':sl[0].stop,'x0':sl[1].start,'x1':sl[1].stop,'area':int(areas[k])})
    return lab,comps
def segment(gray):
    """gray: uint8 HxW crop of a word. returns list of components with features (sorted LTR)"""
    H,W=gray.shape
    thr=otsu(gray); ink=gray<=thr
    if ink.mean()>0.5: ink=~ink
    lab,comps=label(ink)
    minA=max(3,0.006*H*H)
    comps=[c for c in comps if c['area']>=minA]
    if not comps: return [],lab
    comps.sort(key=lambda c:c['x0'])
    # merge stacked parts (x-overlap >= 50% of narrower)
    merged=[]
    for c in comps:
        c=dict(c,ids=[c['id']])
        if merged:
            m=merged[-1]; ov=min(m['x1'],c['x1'])-max(m['x0'],c['x0'])
            oy=min(m['y1'],c['y1'])-max(m['y0'],c['y0'])
            if ov>=0.5*min(m['x1']-m['x0'],c['x1']-c['x0']) and oy<=0.2*min(m['y1']-m['y0'],c['y1']-c['y0']):
                m.update(x0=min(m['x0'],c['x0']),x1=max(m['x1'],c['x1']),y0=min(m['y0'],c['y0']),y1=max(m['y1'],c['y1']),area=m['area']+c['area'],ids=m['ids']+c['ids']); continue
        merged.append(c)
    hs=np.array([m['y1']-m['y0'] for m in merged]); mx=hs.max()
    tall=[m for m in merged if (m['y1']-m['y0'])>=0.5*mx]
    R=float(np.median([m['y1']-m['y0'] for m in tall]))
    mid=float(np.median([(m['y0']+m['y1'])/2 for m in tall]))
    for m in merged:
        w=m['x1']-m['x0']; h=m['y1']-m['y0']; S=max(w,h)
        sq=np.zeros((S,S),np.float32)
        sub=np.isin(lab[m['y0']:m['y1'],m['x0']:m['x1']],m['ids']).astype(np.float32)
        oy=(S-h)//2; ox=(S-w)//2; sq[oy:oy+h,ox:ox+w]=sub
        img=np.zeros((N,N),np.float32)
        cs=np.zeros((S+1,S+1),np.float32); cs[1:,1:]=sq.cumsum(0).cumsum(1)
        for i in range(N):
            y0=(i*S)//N; y1=max(y0+1,((i+1)*S)//N)
            for j in range(N):
                x0=(j*S)//N; x1=max(x0+1,((j+1)*S)//N)
                img[i,j]=(cs[y1,x1]-cs[y0,x1]-cs[y1,x0]+cs[y0,x0])/((y1-y0)*(x1-x0))
        extra=[h/R, w/R, ((m['y0']+m['y1'])/2-mid)/R, m['area']/float(w*h)]
        m['feat']=np.concatenate([img.ravel(),np.array(extra,np.float32)])
    return merged,lab
