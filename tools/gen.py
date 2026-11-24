import sys, random, glob, io, numpy as np
from PIL import Image, ImageDraw, ImageFont, ImageFilter
sys.path.insert(0,'/tmp/claude-0/digits'); from seg import segment
CLS=[str(i) for i in range(10)]+['e%d'%i for i in range(10)]+['.','/','-',':','x']
E='٠١٢٣٤٥٦٧٨٩'; F='۰۱۲۳۴۵۶۷۸۹'
def cls_of(ch):
    if ch.isdigit() and ch in '0123456789': return CLS.index(ch)
    if ch in E: return CLS.index('e%d'%E.index(ch))
    if ch in F: return CLS.index('e%d'%F.index(ch))
    if ch in '٫,.٬': return CLS.index('.')
    if ch in '/': return CLS.index('/')
    if ch in '-': return CLS.index('-')
    if ch in ':': return CLS.index(':')
    return CLS.index('x')
LET='بتثجحخدذرزسشصضطظعغفقكلمنيءرسSRVATINvx×%٪#'
def rnd_string(r):
    east=r.random()<0.6; D=(E if r.random()<0.85 else F) if east else '0123456789'
    sep='٫' if east and r.random()<0.6 else (',' if east and r.random()<0.3 else '.')
    k=r.random(); dg=lambda n:''.join(r.choice(D) for _ in range(n))
    if k<0.45: s=dg(r.randint(1,5))+(sep+dg(2) if r.random()<0.7 else '')
    elif k<0.6: s=dg(2)+'/'+dg(2)+'/'+dg(4)
    elif k<0.75: s=dg(r.randint(8,15))
    elif k<0.85: s=dg(r.randint(1,3))+r.choice(['×','x','-',':','%'])+dg(r.randint(1,3))
    else: s=dg(r.randint(2,6))
    if r.random()<0.15: s=r.choice(LET)+s
    if r.random()<0.15: s=s+r.choice(LET)
    return s
from fontTools.ttLib import TTFont
_CM={}
def cmap(fp):
    if fp not in _CM: _CM[fp]=set(TTFont(fp).getBestCmap().keys())
    return _CM[fp]
def fix(s,cm):
    out=''
    for ch in s:
        if ord(ch) in cm: out+=ch; continue
        if ch in F and ord(E[F.index(ch)]) in cm: out+=E[F.index(ch)]; continue
        if ch=='٫' and ord(',') in cm: out+=','; continue
        return None
    return out
def render(r,font_path):
    size=r.randint(20,64); ft=ImageFont.truetype(font_path,size)
    s=fix(rnd_string(r),cmap(font_path))
    if not s: return []
    W=int(size*(len(s)*1.2+2)); H=int(size*2.2)
    bg=r.randint(170,255); fg=r.randint(0,90)
    im=Image.new('L',(W,H),bg); d=ImageDraw.Draw(im)
    x=size*0.5; y=size*0.4; boxes=[]
    for ch in s:
        d.text((x,y),ch,font=ft,fill=fg)
        bb=ft.getbbox(ch); boxes.append((x+bb[0],x+bb[2],ch))
        x+=ft.getlength(ch)+r.uniform(-0.03,0.15)*size
    if r.random()<0.5:
        im=im.rotate(r.uniform(-1.5,1.5),resample=Image.BILINEAR,fillcolor=bg)
    f=r.uniform(max(0.35,14/size),1.0)
    im=im.resize((max(8,int(W*f)),max(8,int(H*f))),Image.BILINEAR if r.random()<0.5 else Image.LANCZOS)
    if r.random()<0.6: im=im.filter(ImageFilter.GaussianBlur(r.uniform(0.2,1.2)))
    a=np.asarray(im).astype(np.float32)
    if r.random()<0.6: a=a+np.random.default_rng(r.randint(0,1<<30)).normal(0,r.uniform(2,14),a.shape)
    a=np.clip(a,0,255).astype(np.uint8); im=Image.fromarray(a)
    if r.random()<0.5:
        b=io.BytesIO(); im.save(b,'JPEG',quality=r.randint(30,90)); im=Image.open(b).convert('L')
    a=np.asarray(im)
    boxes=[(b0*f,b1*f,ch) for b0,b1,ch in boxes]
    # tight crop like a Tesseract word box
    ink=a<(int(bg)+int(fg))/2
    ys,xs=np.where(ink)
    if len(ys)==0: return []
    y0,y1,x0,x1=ys.min(),ys.max()+1,xs.min(),xs.max()+1
    h=y1-y0; pad=max(2,int(round(0.15*h)))
    Y0,X0=max(0,y0-pad),max(0,x0-pad)
    crop=a[Y0:y1+pad,X0:x1+pad]
    comps,_=segment(crop)
    out=[]; owner={}
    for c in comps:
        cx0,cx1=c['x0']+X0,c['x1']+X0
        hits=[k for k,b in enumerate(boxes) if cx0<= (b[0]+b[1])/2 <=cx1]
        if len(hits)!=1:
            # fragment of a char (no char center inside) -> mark the char nearest as fragmented
            mids=[abs((b[0]+b[1])/2-(cx0+cx1)/2) for b in boxes]; owner.setdefault(int(np.argmin(mids)),[]).append(None); continue
        owner.setdefault(hits[0],[]).append((c['feat'],cls_of(boxes[hits[0]][2])))
    for k,v in owner.items():
        if len(v)==1 and v[0] is not None: out.append(v[0])
    return out
def work(args):
    seed,fonts,n=args; r=random.Random(seed); X=[];Y=[]
    for i in range(n):
        for fe,c in render(r,r.choice(fonts)): X.append(fe);Y.append(c)
    return np.array(X,np.float32),np.array(Y,np.int16)
if __name__=='__main__':
    from multiprocessing import Pool
    fonts=sorted(glob.glob('/tmp/claude-0/fonts/*.ttf'))
    test_fonts=[f for f in fonts if 'Naskh' in f or 'Tajawal' in f]
    train_fonts=[f for f in fonts if f not in test_fonts]
    which=sys.argv[1]; n=int(sys.argv[2])
    fl=train_fonts if which=='train' else test_fonts
    with Pool(4) as p: res=p.map(work,[(i*7919+(0 if which=='train' else 99991),fl,n//16) for i in range(16)])
    X=np.concatenate([a for a,b in res]);Y=np.concatenate([b for a,b in res])
    np.savez_compressed(f'/tmp/claude-0/digits/{which}.npz',X=X,Y=Y)
    print(which,X.shape,np.bincount(Y,minlength=len(CLS)))
