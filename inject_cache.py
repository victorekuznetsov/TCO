# -*- coding: utf-8 -*-
"""Встраивает кэшированные значения формул в xlsx (формулы и оформление сохраняются),
чтобы просмотрщики без пересчёта (мобильные) показывали числа, а не 0."""
import formulas, warnings, zipfile, shutil, re, os, sys
from xml.sax.saxutils import escape
warnings.filterwarnings("ignore")
ERR={"#DIV/0!","#REF!","#VALUE!","#NAME?","#N/A","#NUM!","#NULL!","None"}

def compute(path):
    xl=formulas.ExcelModel().loads(path).finish(); sol=xl.calculate(); vmap={}
    for k,v in sol.items():
        ku=k.upper()
        if "'!" not in ku: continue
        sheet=ku.split("]")[-1].split("'!")[0]; cell=ku.split("!")[-1].replace("$","")
        if not re.match(r"^[A-Z]+\d+$",cell): continue
        try: val=v.value[0,0]
        except:
            try: val=v.value
            except: continue
        vmap[(sheet,cell)]=val
    return vmap

def bad(x):
    s=str(x); return s in ERR or s.startswith("=")

def inject(path,out):
    vmap=compute(path)
    tmp=out+".work"
    if os.path.exists(tmp): shutil.rmtree(tmp)
    with zipfile.ZipFile(path) as z: z.extractall(tmp)
    wbxml=open(f"{tmp}/xl/workbook.xml",encoding="utf-8").read()
    name2rid={m.group(1).upper():m.group(2) for m in re.finditer(r'<sheet[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"',wbxml)}
    rels=open(f"{tmp}/xl/_rels/workbook.xml.rels",encoding="utf-8").read()
    rid2path={}
    for rel in re.finditer(r'<Relationship\b[^>]*/?>',rels):
        s=rel.group(0); idm=re.search(r'Id="([^"]+)"',s); tm=re.search(r'Target="([^"]+)"',s)
        if idm and tm: rid2path[idm.group(1)]=tm.group(1)
    def norm(tgt):
        t=tgt.lstrip("/")
        if t.startswith("xl/"): return f"{tmp}/{t}"
        return f"{tmp}/xl/{t}"
    total=0
    for sname,rid in name2rid.items():
        tgt=rid2path.get(rid)
        if not tgt: continue
        p=norm(tgt)
        if not os.path.exists(p): continue
        xml=open(p,encoding="utf-8").read(); cnt=[0]
        def repl(mm):
            ref,attrs,fpart=mm.group(1),mm.group(2),mm.group(3)
            val=vmap.get((sname,ref))
            if val is None or bad(val): return mm.group(0)
            if isinstance(val,bool): t=' t="b"'; vs="1" if val else "0"
            elif isinstance(val,(int,float)): t=""; vs=repr(float(val))
            else: t=' t="str"'; vs=escape(str(val))
            a=re.sub(r'\s+t="[^"]*"','',attrs); cnt[0]+=1
            return f'<c r="{ref}"{a}{t}>{fpart}<v>{vs}</v></c>'
        xml=re.sub(r'<c r="([A-Z]+\d+)"([^>]*)>(<f[^>]*>.*?</f>)(?:<v[^>]*/>|<v>.*?</v>)?</c>',repl,xml,flags=re.DOTALL)
        open(p,"w",encoding="utf-8").write(xml); total+=cnt[0]
    if os.path.exists(out): os.remove(out)
    with zipfile.ZipFile(out,"w",zipfile.ZIP_DEFLATED) as z:
        for root,_,files in os.walk(tmp):
            for fn in files:
                fp=os.path.join(root,fn); z.write(fp,os.path.relpath(fp,tmp))
    shutil.rmtree(tmp); print("Встроено кэш-значений:",total)

if __name__=="__main__":
    inject(sys.argv[1],sys.argv[2])
