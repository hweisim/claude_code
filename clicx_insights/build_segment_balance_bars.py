rows=[("-","Borrower",74137),("-","Have Balance",74045),("-","No Balance",534085),("-","No Deposit",216022),
      ("LA","Have Balance",25996),("LA","No Balance",54083),("LA","No Deposit",84300)]
T=sum(r[2] for r in rows)
stops=[(0,(0xF0,0xF1,0xF2)),(90e3,(0xD5,0xE4,0xE1)),(270e3,(0xA3,0xD3,0xCC)),(545e3,(0x5B,0xA4,0xA0))]
def col(v):
    for (a,ca),(b,cb) in zip(stops,stops[1:]):
        if v<=b:
            t=(v-a)/(b-a); return '#%02X%02X%02X'%tuple(round(x+(y-x)*t) for x,y in zip(ca,cb))
    return '#5BA4A0'
X0=416; PX=98.5/200e3; ys=[52,131,209,288,366,444,523]; bh=58
h=['<!doctype html><meta charset="utf-8"><style>body{margin:0;background:#fff}#c{position:relative;width:890px;height:575px;font-family:Arial,"Liberation Sans",Helvetica,sans-serif;color:#2b2b2b}'
   '.t{position:absolute;font-size:26px;line-height:30px;white-space:nowrap}.v{position:absolute;font-size:24px;line-height:28px;white-space:nowrap}.p{color:#4E79A7}</style><div id="c">']
for gx in [X0+98.5*i for i in range(5)]:
    h.append(f'<div style="position:absolute;left:{gx:.1f}px;top:14px;width:1px;height:546px;background:#E6E6E6"></div>')
h.append('<div style="position:absolute;left:14px;top:325px;width:836px;height:2px;background:#D9D9D9"></div>')
h.append('<div style="position:absolute;left:416px;top:14px;width:1px;height:546px;background:#BDBDBD"></div>')
h.append('<div class="t" style="left:22px;top:36px">-</div>')
h.append('<div class="t" style="left:34px;top:350px;font-size:24px;line-height:34px;color:#2E3B4E">Loan<br>Applicants<br>(Not<br>booked)</div>')
for (g,lab,v),y in zip(rows,ys):
    w=v*PX
    h.append(f'<div class="t" style="left:210px;top:{y-15}px">{lab}</div>')
    h.append(f'<div style="position:absolute;left:{X0}px;top:{y-bh/2}px;width:{w:.1f}px;height:{bh}px;background:{col(v)}"></div>')
    h.append(f'<div class="v" style="left:{X0+w+9:.1f}px;top:{y-15}px">{round(v/1000)}K<span class="p"> ({round(100*v/T)}%)</span></div>')
h.append('</div>')
open('chart.html','w').write(''.join(h)); print(T,[ (r[1],round(r[2]/1000),round(100*r[2]/T)) for r in rows])
