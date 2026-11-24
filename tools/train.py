import numpy as np, json
CLS=[str(i) for i in range(10)]+['e%d'%i for i in range(10)]+['.','/','-',':','x']
d=np.load('train.npz');X,Y=d['X'],d['Y'].astype(np.int64)
t=np.load('test.npz');Xt,Yt=t['X'],t['Y'].astype(np.int64)
# clip extras
def prep(X):
    X=X.copy(); X[:,-4:-1]=np.clip(X[:,-4:-1],-3,3); return X
X,Xt=prep(X),prep(Xt)
mu=X.mean(0); sd=X.std(0)+1e-3
X=(X-mu)/sd; Xt=(Xt-mu)/sd
rng=np.random.default_rng(0); C=len(CLS); H1=256
W1=rng.normal(0,np.sqrt(2/X.shape[1]),(X.shape[1],H1)).astype(np.float32); b1=np.zeros(H1,np.float32)
W2=rng.normal(0,np.sqrt(2/H1),(H1,C)).astype(np.float32); b2=np.zeros(C,np.float32)
P=[W1,b1,W2,b2]; M=[np.zeros_like(p) for p in P]; V=[np.zeros_like(p) for p in P]
def fwd(X):
    h=np.maximum(0,X@W1+b1); return h,h@W2+b2
lr=2e-3; step=0; bs=256
for ep in range(14):
    idx=rng.permutation(len(X))
    for k in range(0,len(X),bs):
        j=idx[k:k+bs]; x=X[j]; y=Y[j]
        x=x+rng.normal(0,0.05,x.shape).astype(np.float32)
        h,o=fwd(x); o=o-o.max(1,keepdims=True); p=np.exp(o); p/=p.sum(1,keepdims=True)
        g=p; g[np.arange(len(y)),y]-=1; g/=len(y)
        gW2=h.T@g+1e-5*W2; gb2=g.sum(0); gh=g@W2.T; gh[h<=0]=0
        gW1=x.T@gh+1e-5*W1; gb1=gh.sum(0)
        step+=1
        for i,(p_,g_) in enumerate(zip(P,[gW1,gb1,gW2,gb2])):
            M[i]=0.9*M[i]+0.1*g_; V[i]=0.999*V[i]+0.001*g_*g_
            p_-=lr*(M[i]/(1-0.9**step))/(np.sqrt(V[i]/(1-0.999**step))+1e-8)
    lr*=0.75
    _,o=fwd(Xt); acc=(o.argmax(1)==Yt).mean()
    print(ep,'test acc',round(acc,4),flush=True)
_,o=fwd(Xt); pr=o.argmax(1)
conf=np.zeros((C,C),int)
for a,b in zip(Yt,pr): conf[a,b]+=1
for i in range(C):
    errs=[(CLS[j],conf[i,j]) for j in range(C) if j!=i and conf[i,j]>3]
    print(CLS[i], round(conf[i,i]/max(1,conf[i].sum()),3), errs)
json.dump({'N':20,'cls':CLS,'mu':mu.round(5).tolist(),'sd':sd.round(5).tolist(),'W1':W1.round(4).tolist(),'b1':b1.round(4).tolist(),'W2':W2.round(4).tolist(),'b2':b2.round(4).tolist()},open('model.json','w'))
