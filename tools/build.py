import sys,base64
t=open(sys.argv[1],encoding='utf-8').read(); d=open(sys.argv[2],encoding='utf-8').read().replace('</','<\\/')
logo='data:image/png;base64,'+base64.b64encode(open(sys.argv[4],'rb').read()).decode()
t=t.replace('/*__LOGO__*/',logo)
open(sys.argv[3],'w',encoding='utf-8').write(t.replace('/*__DATA__*/null',d))
