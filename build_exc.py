# -*- coding: utf-8 -*-
"""
Расчётная модель TCO горных экскаваторов по методике аннуитета (руб/м3).
Воспроизводит логику корпоративной модели сравнения экскаваторов
(пост-налоговый дисконтированный аннуитет, разложенный на операционный и
инвестиционный) в виде чистого переиспользуемого input-driven файла.

Формируемый файл: Модель_TCO_экскаваторов_аннуитет.xlsx
"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference, Series
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import SeriesLabel
from openpyxl.chart.data_source import StrRef
from openpyxl.worksheet.datavalidation import DataValidation

S_MET="Методика"; S_PAR="Параметры"; S_IN="Ввод данных"
S_PROD="Производительность"; S_ANN="Расчёт аннуитета"; S_CMP="Сравнение"; S_SENS="Чувствительность"

EXC=["D","E","F","G","H","I"]          # 6 экскаваторов
NEXC=6

# ---- палитра ----
DARK="1F3864"; MID="2E5496"; INP="FFF2CC"; RES="E2EFDA"; HEAD="D6DCE5"; GREY="F2F2F2"; BEST="C6EFCE"; ACC="FCE4D6"
Fd=PatternFill("solid",fgColor=DARK); Fm=PatternFill("solid",fgColor=MID); Fi=PatternFill("solid",fgColor=INP)
Fr=PatternFill("solid",fgColor=RES); Fh=PatternFill("solid",fgColor=HEAD); Fg=PatternFill("solid",fgColor=GREY)
Fb=PatternFill("solid",fgColor=BEST); Fa=PatternFill("solid",fgColor=ACC)
FT=Font(size=15,bold=True,color="FFFFFF"); FSub=Font(size=10,italic=True,color="FFFFFF")
FSec=Font(size=11,bold=True,color="FFFFFF"); FHd=Font(size=9,bold=True,color=DARK)
FB=Font(size=10,bold=True); FN=Font(size=10); FRz=Font(size=10,bold=True,color="1F3864"); FU=Font(size=8,italic=True,color="808080")
thin=Side(style="thin",color="BFBFBF"); BD=Border(left=thin,right=thin,top=thin,bottom=thin)
L=Alignment("left",vertical="center",wrap_text=True); C=Alignment("center",vertical="center",wrap_text=True); R=Alignment("right",vertical="center")
M="#,##0"; M1="#,##0.0"; M2="#,##0.00"; P1="0.0%"; P2="0.00%"; NUM="#,##0.000"

wb=Workbook()

def title(ws,t,s,last="K"):
    ws.merge_cells(f"B1:{last}1"); ws["B1"]=t; ws["B1"].font=FT; ws["B1"].alignment=L
    ws.merge_cells(f"B2:{last}2"); ws["B2"]=s; ws["B2"].font=FSub; ws["B2"].alignment=L
    for r in (1,2):
        for col in range(2,ord(last)-64+1): ws.cell(row=r,column=col).fill=Fd

def section(ws,row,text,c0="B",c1="K"):
    ws.merge_cells(f"{c0}{row}:{c1}{row}"); cc=ws[f"{c0}{row}"]; cc.value=text; cc.font=FSec; cc.alignment=L
    for col in range(ord(c0)-64,ord(c1)-64+1): ws.cell(row=row,column=col).fill=Fm

# финансовые генераторы формул
def PVAF(r,n): return f"IF({r}=0,{n},(1-(1+{r})^-{n})/{r})"

# ================================================================= #
# ЛИСТ: ПАРАМЕТРЫ
# ================================================================= #
wp=wb.active; wp.title=S_PAR; wp.sheet_view.showGridLines=False
title(wp,"ГЛОБАЛЬНЫЕ ПАРАМЕТРЫ МОДЕЛИ","Единые допущения для всех сравниваемых экскаваторов. Жёлтые ячейки — ввод.",last="F")
wp.column_dimensions["A"].width=2; wp.column_dimensions["B"].width=46
wp.column_dimensions["C"].width=13; wp.column_dimensions["D"].width=12; wp.column_dimensions["E"].width=50

def phdr(row):
    for col,t in (("B","Параметр"),("C","Значение"),("D","Ед."),("E","Комментарий")):
        c=wp[f"{col}{row}"]; c.value=t; c.font=FHd; c.fill=Fh; c.border=BD; c.alignment=C

prow={}
def padd(row,key,label,val,unit,fmt,comment):
    prow[key]=row
    wp[f"B{row}"]=label; wp[f"B{row}"].font=FN; wp[f"B{row}"].border=BD; wp[f"B{row}"].alignment=L
    c=wp[f"C{row}"]; c.value=val; c.fill=Fi; c.border=BD; c.alignment=C; c.font=FB
    if fmt: c.number_format=fmt
    wp[f"D{row}"]=unit; wp[f"D{row}"].font=FU; wp[f"D{row}"].alignment=C; wp[f"D{row}"].border=BD
    wp[f"E{row}"]=comment; wp[f"E{row}"].font=FU; wp[f"E{row}"].alignment=L; wp[f"E{row}"].border=BD

def pcalc(row,key,label,formula,unit,fmt,comment):
    prow[key]=row
    wp[f"B{row}"]=label; wp[f"B{row}"].font=FN; wp[f"B{row}"].border=BD; wp[f"B{row}"].alignment=L
    c=wp[f"C{row}"]; c.value="="+formula; c.fill=Fr; c.border=BD; c.alignment=C; c.font=FRz
    if fmt: c.number_format=fmt
    wp[f"D{row}"]=unit; wp[f"D{row}"].font=FU; wp[f"D{row}"].alignment=C; wp[f"D{row}"].border=BD
    wp[f"E{row}"]=comment; wp[f"E{row}"].font=FU; wp[f"E{row}"].alignment=L; wp[f"E{row}"].border=BD

def pc(key): return f"'{S_PAR}'!C{prow[key]}"

section(wp,4,"Финансовые допущения","B","E"); phdr(5)
padd(6,"disc","Ставка дисконтирования (WACC)",0.12,"%/год",P1,"Стоимость капитала для приведения к текущей стоимости")
padd(7,"tax","Налог на прибыль",0.25,"%",P1,"Даёт налоговый щит на затраты и амортизацию")
padd(8,"horizon","Срок эксплуатации (горизонт)",10,"лет",M,"Период сравнения жизненного цикла")
padd(9,"dep_years","Срок амортизации (налоговый учёт)",5.17,"лет",M2,"Линейная амортизация; 62 мес ≈ 5,17 года")
padd(10,"fuel_p","Цена дизельного топлива",54.83,"руб/кг","#,##0.00","Средняя цена на площадке эксплуатации")

section(wp,12,"Курсы валют","B","E"); phdr(13)
padd(14,"cny","Курс юаня (CNY)",11.2068,"руб",NUM,"Для оборудования из Китая")
padd(15,"usd","Курс доллара (USD)",76.1258,"руб",NUM,"")
padd(16,"eur","Курс евро (EUR)",86.8976,"руб",NUM,"")

section(wp,18,"Параметры производительности (карьер)","B","E"); phdr(19)
padd(20,"kfv","Календарный фонд времени",8760,"час/год",M,"24 ч × 365 дн")
padd(21,"dens","Плотность породы",2.787,"т/м³",NUM,"В целике")
padd(22,"loosen","Коэффициент разрыхления",1.5,"коэф.",M2,"")
padd(23,"fill","Коэффициент наполнения ковша",0.9,"коэф.",M2,"")
padd(24,"loss","Коэффициент потерь при экскавации",0.9,"коэф.",M2,"")
padd(25,"cap","Грузоподъёмность самосвала",136,"т",M,"Эталонный самосвал для расчёта")
padd(26,"pos","Время постановки самосвала",32,"сек",M,"Время оборота а/с под погрузку")
padd(27,"shift","Длительность смены",12,"час",M,"Для пересчёта простоев")

section(wp,29,"Расчётные финансовые коэффициенты","B","E"); phdr(30)
pcalc(31,"S","Коэф. аннуитета Σдиск.факторов (r, N)",PVAF(pc('disc'),pc('horizon')),"коэф.",NUM,"Сумма дисконт-факторов за срок службы")
_d=pc('dep_years'); _r=pc('disc')
_dsf=f"(1/{_d})*((1-(1+{_r})^-INT({_d}))/{_r}+({_d}-INT({_d}))*(1+{_r})^-(INT({_d})+1))"
pcalc(32,"dsf","Коэф. приведённой амортизации",_dsf,"коэф.",NUM,"Приведённая амортизация на 1 руб. цены (налог. щит)")

# легенда
lg=34
wp[f"B{lg}"]="Легенда:"; wp[f"B{lg}"].font=FB
wp[f"C{lg}"].fill=Fi; wp[f"C{lg}"].border=BD; wp[f"D{lg}"]="— ввод данных"; wp[f"D{lg}"].font=FN; wp.merge_cells(f"D{lg}:E{lg}")
wp[f"C{lg+1}"].fill=Fr; wp[f"C{lg+1}"].border=BD; wp[f"D{lg+1}"]="— расчётные показатели"; wp[f"D{lg+1}"].font=FN; wp.merge_cells(f"D{lg+1}:E{lg+1}")

print("Параметры:",prow)

# ================================================================= #
# ЛИСТ: ВВОД ДАННЫХ (6 экскаваторов)
# ================================================================= #
wi=wb.create_sheet(S_IN); wi.sheet_view.showGridLines=False
title(wi,"ВВОД ИСХОДНЫХ ДАННЫХ ПО ЭКСКАВАТОРАМ","Сравнение до 6 машин одного класса. Пример: экскаваторы класса 2000 (~12 м³).")
wi.column_dimensions["A"].width=2; wi.column_dimensions["B"].width=42; wi.column_dimensions["C"].width=12
for cc in EXC: wi.column_dimensions[cc].width=15
wi.column_dimensions["K"].width=34

HR=4
wi[f"B{HR}"]="Показатель"; wi[f"B{HR}"].font=FHd; wi[f"B{HR}"].fill=Fh; wi[f"B{HR}"].border=BD; wi[f"B{HR}"].alignment=C
wi[f"C{HR}"]="Ед."; wi[f"C{HR}"].font=FHd; wi[f"C{HR}"].fill=Fh; wi[f"C{HR}"].border=BD; wi[f"C{HR}"].alignment=C
for i,cc in enumerate(EXC):
    c=wi[f"{cc}{HR}"]; c.value=f"Экскаватор {i+1}"; c.font=FHd; c.fill=Fh; c.border=BD; c.alignment=C
wi[f"K{HR}"]="Пояснение"; wi[f"K{HR}"].font=FHd; wi[f"K{HR}"].fill=Fh; wi[f"K{HR}"].border=BD; wi[f"K{HR}"].alignment=C

names=["Komatsu PC2000-8\nMining Solutions","Komatsu PC2000-11R\nИСТК","Zoomlion ZE2000G\nОрикс",
       "Sany SY2000\nАСТ","Shantui SE2000LCW","TZCO TZ2000\nДзикай"]
supplier=["Mining Solutions","ИСТК","Орикс","АСТ","Строймпорттехника","Дзикай"]
engines=["Komatsu SAA12V140E","Komatsu SAA12V140E-7","Cummins QST30","Cummins QSK38","Weichai 12M33","Weichai 12M33"]
price_val=[3800,3401.64,20129,19621.62,16742.05,15400]
price_cur=["USD","USD","CNY","CNY","CNY","CNY"]
install=[0,0,0,0,0,0]
bucket=[10,10,10,10,9.6,12]
cycle=[37.3,37.3,37.3,37.3,37.3,37.3]
ktg=[0.840,0.808,0.738,0.738,0.738,0.738]
downtime=[266.1]*6
fuel_kgh=[88.46,85.81,90.23,92.04,102,120]
maint=[42690,42690,34556,65685,34556,51705]
bucket_m3=[0.87]*6
personnel=[10015]*6

LAY=[
 ("SEC","A. ИДЕНТИФИКАЦИЯ"),
 ("R","name","Наименование / модель","—",names,"Название варианта",True,None),
 ("R","supplier","Поставщик","—",supplier,"",True,None),
 ("R","engine","Двигатель","—",engines,"",True,None),
 ("SEC","B. СТОИМОСТЬ"),
 ("R","price_val","Цена за единицу","тыс.вал.",price_val,"В валюте контракта",False,M),
 ("R","price_cur","Валюта цены","—",price_cur,"CNY / USD / EUR / RUB",True,None),
 ("R","install","Монтаж, доставка, ПНР","тыс.руб",install,"Дополнительно к цене",False,M),
 ("SEC","C. ПРОИЗВОДИТЕЛЬНОСТЬ"),
 ("R","bucket","Геометрический объём ковша","м³",bucket,"Участвует в расчёте производительности",False,M1),
 ("R","cycle","Время цикла экскавации","сек",cycle,"Черпание–поворот–разгрузка–поворот",False,M1),
 ("SEC","D. РЕЖИМ ЭКСПЛУАТАЦИИ"),
 ("R","ktg","КТГ (коэф. техготовности), средний","коэф.",ktg,"Средний за срок службы (доля готового времени)",False,P1),
 ("R","downtime","Ежесменные простои","мин/см",downtime,"ОТМ, ожидание а/с, прочие организационные",False,M1),
 ("SEC","E. ТОПЛИВО"),
 ("R","fuel_kgh","Удельный расход ДТ","кг/час",fuel_kgh,"Ключевой драйвер OPEX",False,M1),
 ("SEC","F. ГОДОВЫЕ ЗАТРАТЫ (средние)"),
 ("R","maint","ТОиР и сервис (ТО, ТР, ППР, ремонты)","тыс.руб/год",maint,"Среднегодовые затраты на обслуживание и ремонт",False,M),
 ("R","bucket_m3","Расходники на ковш (зубья, коронки)","руб/м³",bucket_m3,"Удельные на кубометр вынутой породы",False,M2),
 ("R","personnel","Расходы на персонал (экипаж)","тыс.руб/год",personnel,"ФОТ с отчислениями на 1 машину",False,M),
]
irow={}
r=HR+2
for it in LAY:
    if it[0]=="SEC":
        section(wi,r,it[1],"B","K"); r+=1; continue
    _,key,label,unit,vals,comment,istext,fmt=it
    irow[key]=r
    wi[f"B{r}"]=label; wi[f"B{r}"].font=FN; wi[f"B{r}"].border=BD; wi[f"B{r}"].alignment=L
    wi[f"C{r}"]=unit; wi[f"C{r}"].font=FU; wi[f"C{r}"].border=BD; wi[f"C{r}"].alignment=C
    for i,cc in enumerate(EXC):
        c=wi[f"{cc}{r}"]; c.value=vals[i]; c.fill=Fi; c.border=BD; c.font=FN
        c.alignment=L if istext else C
        if fmt: c.number_format=fmt
    wi[f"K{r}"]=comment; wi[f"K{r}"].font=FU; wi[f"K{r}"].border=BD; wi[f"K{r}"].alignment=L
    r+=1
# валидация валют
dv=DataValidation(type="list",formula1='"CNY,USD,EUR,RUB"',allow_blank=False); wi.add_data_validation(dv)
dv.add(f"D{irow['price_cur']}:I{irow['price_cur']}")

def inp(key,col): return f"'{S_IN}'!{col}{irow[key]}"
print("Ввод:",irow)

# ================================================================= #
# ЛИСТ: ПРОИЗВОДИТЕЛЬНОСТЬ
# ================================================================= #
wpr=wb.create_sheet(S_PROD); wpr.sheet_view.showGridLines=False
title(wpr,"РАСЧЁТ ПРОИЗВОДИТЕЛЬНОСТИ","Часовая и годовая производительность (тыс.м³) по каждому экскаватору.")
wpr.column_dimensions["A"].width=2; wpr.column_dimensions["B"].width=42; wpr.column_dimensions["C"].width=12
for cc in EXC: wpr.column_dimensions[cc].width=15

PHR=4
wpr[f"B{PHR}"]="Показатель"; wpr[f"B{PHR}"].font=FHd; wpr[f"B{PHR}"].fill=Fh; wpr[f"B{PHR}"].border=BD; wpr[f"B{PHR}"].alignment=C
wpr[f"C{PHR}"]="Ед."; wpr[f"C{PHR}"].font=FHd; wpr[f"C{PHR}"].fill=Fh; wpr[f"C{PHR}"].border=BD; wpr[f"C{PHR}"].alignment=C
for i,cc in enumerate(EXC):
    c=wpr[f"{cc}{PHR}"]; c.value=f"='{S_IN}'!{cc}{irow['name']}"; c.font=FHd; c.fill=Fh; c.border=BD; c.alignment=C

prd={}; pr=PHR+1
def padd2(key,label,unit,fn,fmt=M1,res=False,bold=False):
    global pr
    prd[key]=pr
    b=wpr[f"B{pr}"]; b.value=label; b.border=BD; b.alignment=L; b.font=(FRz if res else (FB if bold else FN))
    wpr[f"C{pr}"]=unit; wpr[f"C{pr}"].font=FU; wpr[f"C{pr}"].border=BD; wpr[f"C{pr}"].alignment=C
    for cc in EXC:
        c=wpr[f"{cc}{pr}"]; c.value="="+fn(cc); c.number_format=fmt; c.border=BD; c.alignment=R
        c.font=(FRz if res else FN)
        if res: c.fill=Fr
    pr+=1
def pref(key,col): return f"{col}{prd[key]}"

def cur_rate(col):
    cur=inp('price_cur',col)
    return f'IF({cur}="CNY",{pc("cny")},IF({cur}="USD",{pc("usd")},IF({cur}="EUR",{pc("eur")},1)))'

padd2("price_rub","Стоимость в рублях","тыс.руб",
      lambda c:f"{inp('price_val',c)}*{cur_rate(c)}+{inp('install',c)}",M,bold=True)
padd2("rock","Порода в ковше","т",
      lambda c:f"{inp('bucket',c)}*{pc('dens')}*({pc('fill')}/{pc('loosen')})*{pc('loss')}",M2)
padd2("buckets","Ковшей на самосвал (циклов)","шт",
      lambda c:f"ROUNDUP({pc('cap')}/{pref('rock',c)},0)",M)
padd2("trucks","Самосвалов за час работы","а/с/ч",
      lambda c:f"3600/({pref('buckets',c)}*{inp('cycle',c)}+{pc('pos')})",M2)
padd2("th","Часовая производительность","т/час",
      lambda c:f"{pc('cap')}*{pref('trucks',c)}",M)
padd2("m3h","Часовая производительность","м³/час",
      lambda c:f"{pref('th',c)}/{pc('dens')}",M1)
padd2("tech","Время техготовности","час/год",
      lambda c:f"{pc('kfv')}*{inp('ktg',c)}",M)
padd2("eff","Эффективное (рабочее) время","час/год",
      lambda c:f"{pref('tech',c)}*(1-{inp('downtime',c)}/(60*{pc('shift')}))",M)
padd2("kio","КИО (коэф. использования)","коэф.",
      lambda c:f"{pref('eff',c)}/{pc('kfv')}",P1)
padd2("am3","Годовая производительность","тыс.м³/год",
      lambda c:f"{pref('m3h',c)}*{pref('eff',c)}/1000",M,res=True)

def prodref(key,col): return f"'{S_PROD}'!{col}{prd[key]}"
print("Производительность:",prd)

# ================================================================= #
# ЛИСТ: РАСЧЁТ АННУИТЕТА
# ================================================================= #
wa=wb.create_sheet(S_ANN); wa.sheet_view.showGridLines=False
title(wa,"РАСЧЁТ АННУИТЕТА (РУБ/М³)","Пост-налоговый приведённый аннуитет: операционный + инвестиционный. Разбивка по статьям.")
wa.column_dimensions["A"].width=2; wa.column_dimensions["B"].width=44; wa.column_dimensions["C"].width=12
for cc in EXC: wa.column_dimensions[cc].width=15

AHR=4
wa[f"B{AHR}"]="Статья"; wa[f"B{AHR}"].font=FHd; wa[f"B{AHR}"].fill=Fh; wa[f"B{AHR}"].border=BD; wa[f"B{AHR}"].alignment=C
wa[f"C{AHR}"]="Ед."; wa[f"C{AHR}"].font=FHd; wa[f"C{AHR}"].fill=Fh; wa[f"C{AHR}"].border=BD; wa[f"C{AHR}"].alignment=C
for i,cc in enumerate(EXC):
    c=wa[f"{cc}{AHR}"]; c.value=f"='{S_IN}'!{cc}{irow['name']}"; c.font=FHd; c.fill=Fh; c.border=BD; c.alignment=C

ar={}; a=AHR+1
def asec(t):
    global a; section(wa,a,t,"B","I"); a+=1
def aadd(key,label,unit,fn,fmt=M,res=False,bold=False):
    global a
    ar[key]=a
    b=wa[f"B{a}"]; b.value=label; b.border=BD; b.alignment=L; b.font=(FRz if res else (FB if bold else FN))
    wa[f"C{a}"]=unit; wa[f"C{a}"].font=FU; wa[f"C{a}"].border=BD; wa[f"C{a}"].alignment=C
    for cc in EXC:
        c=wa[f"{cc}{a}"]; c.value="="+fn(cc); c.number_format=fmt; c.border=BD; c.alignment=R
        c.font=(FRz if res else FN)
        if res: c.fill=Fr
    a+=1
def aref(key,col): return f"{col}{ar[key]}"

# --- годовые денежные затраты ---
asec("ГОДОВЫЕ ДЕНЕЖНЫЕ ЗАТРАТЫ (средние, тыс.руб/год)")
aadd("fuel","Дизтопливо",  "тыс.руб/год",
     lambda c:f"{prodref('eff',c)}*{inp('fuel_kgh',c)}*{pc('fuel_p')}/1000")
aadd("maint","ТОиР и сервис","тыс.руб/год", lambda c:f"{inp('maint',c)}")
aadd("bucket","Расходники на ковш","тыс.руб/год", lambda c:f"{inp('bucket_m3',c)}*{prodref('am3',c)}")
aadd("pers","Расходы на персонал","тыс.руб/год", lambda c:f"{inp('personnel',c)}")
aadd("cash","ИТОГО денежные затраты (OPEX)","тыс.руб/год",
     lambda c:f"{aref('fuel',c)}+{aref('maint',c)}+{aref('bucket',c)}+{aref('pers',c)}",bold=True)

# --- аннуитет по статьям (тыс.руб/год) ---
asec("ГОДОВОЙ АННУИТЕТ ПО СТАТЬЯМ (тыс.руб/год)")
aadd("a_inv","Инвестиционный аннуитет (возврат капитала)","тыс.руб/год",
     lambda c:f"{prodref('price_rub',c)}/{pc('S')}")
aadd("a_fuel","— дизтопливо","тыс.руб/год", lambda c:f"{aref('fuel',c)}")
aadd("a_maint","— ТОиР и сервис","тыс.руб/год", lambda c:f"{aref('maint',c)}")
aadd("a_bucket","— расходники на ковш","тыс.руб/год", lambda c:f"{aref('bucket',c)}")
aadd("a_pers","— персонал","тыс.руб/год", lambda c:f"{aref('pers',c)}")
aadd("a_tax","Налог на прибыль (щит: OPEX+амортизация)","тыс.руб/год",
     lambda c:f"-{pc('tax')}*({aref('cash',c)}+{prodref('price_rub',c)}*{pc('dsf')}/{pc('S')})")
aadd("a_op","Операционный аннуитет (после налога)","тыс.руб/год",
     lambda c:f"{aref('cash',c)}+{aref('a_tax',c)}",bold=True)
aadd("a_tot","ОБЩИЙ АННУИТЕТ","тыс.руб/год",
     lambda c:f"{aref('a_inv',c)}+{aref('a_op',c)}",bold=True)

# --- удельный аннуитет руб/м³ ---
asec("УДЕЛЬНЫЙ АННУИТЕТ (руб/м³) — ГЛАВНЫЙ КРИТЕРИЙ")
def perm3(key): return lambda c:f"{aref(key,c)}/{prodref('am3',c)}"
aadd("u_inv","Инвестиционный","руб/м³",perm3("a_inv"),M2)
aadd("u_fuel","Дизтопливо","руб/м³",perm3("a_fuel"),M2)
aadd("u_maint","ТОиР и сервис","руб/м³",perm3("a_maint"),M2)
aadd("u_bucket","Расходники на ковш","руб/м³",perm3("a_bucket"),M2)
aadd("u_pers","Персонал","руб/м³",perm3("a_pers"),M2)
aadd("u_tax","Налог на прибыль (щит)","руб/м³",perm3("a_tax"),M2)
aadd("u_op","Операционный аннуитет","руб/м³",perm3("a_op"),M2,bold=True)
aadd("u_tot","ОБЩИЙ АННУИТЕТ","руб/м³",perm3("a_tot"),M2,res=True)

asec("ДОПОЛНИТЕЛЬНО")
aadd("sebest","Себестоимость (без дисконта)","руб/м³",
     lambda c:f"{aref('cash',c)}/{prodref('am3',c)}",M2)
aadd("cap_share","Доля инвестиций в аннуитете","%",
     lambda c:f"{aref('u_inv',c)}/{aref('u_tot',c)}",P1)

def annref(key,col): return f"'{S_ANN}'!{col}{ar[key]}"
print("Аннуитет:",ar)

# ================================================================= #
# ЛИСТ: СРАВНЕНИЕ
# ================================================================= #
wm=wb.create_sheet(S_CMP); wm.sheet_view.showGridLines=False
title(wm,"СРАВНИТЕЛЬНЫЙ АНАЛИЗ ЭКСКАВАТОРОВ","Ранжирование по удельному аннуитету (руб/м³). Зелёным — лучший вариант.")
wm.column_dimensions["A"].width=2; wm.column_dimensions["B"].width=40; wm.column_dimensions["C"].width=11
for cc in EXC: wm.column_dimensions[cc].width=15

MHR=4
wm[f"B{MHR}"]="Показатель"; wm[f"B{MHR}"].font=FHd; wm[f"B{MHR}"].fill=Fh; wm[f"B{MHR}"].border=BD; wm[f"B{MHR}"].alignment=C
wm[f"C{MHR}"]="Ед."; wm[f"C{MHR}"].font=FHd; wm[f"C{MHR}"].fill=Fh; wm[f"C{MHR}"].border=BD; wm[f"C{MHR}"].alignment=C
for i,cc in enumerate(EXC):
    c=wm[f"{cc}{MHR}"]; c.value=f"='{S_IN}'!{cc}{irow['name']}"; c.font=FHd; c.fill=Fh; c.border=BD; c.alignment=C

metrics=[
 ("Стоимость (в рублях)","price_rub",prodref,M,"тыс.руб"),
 ("Годовая производительность","am3",prodref,M,"тыс.м³"),
 ("КИО","kio",prodref,P1,"коэф."),
 ("Себестоимость (без дисконта)","sebest",annref,M2,"руб/м³"),
 ("Инвестиционный аннуитет","u_inv",annref,M2,"руб/м³"),
 ("Операционный аннуитет","u_op",annref,M2,"руб/м³"),
 ("ОБЩИЙ АННУИТЕТ","u_tot",annref,M2,"руб/м³"),
]
mrow={}; m=MHR+1
for label,key,reff,fmt,unit in metrics:
    mrow[key]=m
    b=wm[f"B{m}"]; b.value=label; b.border=BD; b.alignment=L; b.font=(FB if key=="u_tot" else FN)
    wm[f"C{m}"]=unit; wm[f"C{m}"].font=FU; wm[f"C{m}"].border=BD; wm[f"C{m}"].alignment=C
    for cc in EXC:
        c=wm[f"{cc}{m}"]; c.value="="+reff(key,cc); c.number_format=fmt; c.border=BD; c.alignment=R
        c.font=(FB if key=="u_tot" else FN)
        if key=="u_tot": c.fill=Fr
    m+=1

# рейтинг
section(wm,m,"РАНЖИРОВАНИЕ (критерий: минимум аннуитета руб/м³)","B","I"); m+=1
tot_rng=f"D{mrow['u_tot']}:I{mrow['u_tot']}"
wm[f"B{m}"]="Место в рейтинге"; wm[f"B{m}"].font=FB; wm[f"B{m}"].border=BD
for cc in EXC:
    c=wm[f"{cc}{m}"]; c.value=f"=RANK({cc}{mrow['u_tot']},{tot_rng},1)"; c.font=FB; c.border=BD; c.alignment=C
m+=1
wm[f"B{m}"]="Отклонение от лучшего"; wm[f"B{m}"].font=FN; wm[f"B{m}"].border=BD
for cc in EXC:
    c=wm[f"{cc}{m}"]; c.value=f"={cc}{mrow['u_tot']}/MIN({tot_rng})-1"; c.number_format=P1; c.border=BD; c.alignment=C; c.font=FN
m+=2

# рекомендация
names_rng=f"D{MHR}:I{MHR}"
wm[f"B{m}"]="РЕКОМЕНДУЕМЫЙ ВАРИАНТ:"; wm[f"B{m}"].font=Font(bold=True,size=12,color=DARK)
wm.merge_cells(f"C{m}:I{m}")
rec=wm[f"C{m}"]
rec.value=f'=INDEX({names_rng},MATCH(MIN({tot_rng}),{tot_rng},0))&"  ("&TEXT(MIN({tot_rng}),"0.00")&" руб/м³)"'
rec.font=Font(bold=True,size=12,color="006100"); rec.alignment=C
for cc in EXC[:0]: pass
for col in ["C","D","E","F","G","H","I"]: wm[f"{col}{m}"].fill=Fb; wm[f"{col}{m}"].border=BD
m+=2

# структура аннуитета для диаграммы
section(wm,m,"СТРУКТУРА УДЕЛЬНОГО АННУИТЕТА (руб/м³)","B","I"); m+=1
wm[f"B{m}"]="Статья"; wm[f"B{m}"].font=FHd; wm[f"B{m}"].fill=Fh; wm[f"B{m}"].border=BD
for cc in EXC:
    c=wm[f"{cc}{m}"]; c.value=f"='{S_IN}'!{cc}{irow['name']}"; c.font=FHd; c.fill=Fh; c.border=BD; c.alignment=C
stop=m+1
struct=[("Инвестиционный","u_inv"),("Дизтопливо","u_fuel"),("ТОиР и сервис","u_maint"),
        ("Расходники на ковш","u_bucket"),("Персонал","u_pers"),("Налог (щит)","u_tax")]
m=stop
for label,key in struct:
    b=wm[f"B{m}"]; b.value=label; b.font=FN; b.border=BD; b.alignment=L
    for cc in EXC:
        c=wm[f"{cc}{m}"]; c.value="="+annref(key,cc); c.number_format=M2; c.border=BD; c.alignment=R; c.font=FN
    m+=1
sbot=m-1

catref=Reference(wm,min_col=4,min_row=stop-1,max_col=9,max_row=stop-1)
chart1=BarChart(); chart1.type="col"; chart1.grouping="stacked"; chart1.overlap=100
chart1.title="Структура аннуитета по вариантам, руб/м³"; chart1.height=9; chart1.width=20
chart1.y_axis.title="руб/м³"
for i in range(len(struct)):
    row_i=stop+i
    vals=Reference(wm,min_col=4,min_row=row_i,max_col=9,max_row=row_i)
    ser=Series(vals,title_from_data=False)
    ser.tx=SeriesLabel(strRef=StrRef(f"'{S_CMP}'!$B${row_i}"))
    chart1.series.append(ser)
chart1.set_categories(catref)
wm.add_chart(chart1,"K4")

chart2=BarChart(); chart2.type="col"; chart2.title="Общий аннуитет, руб/м³"; chart2.height=8; chart2.width=11
chart2.y_axis.title="руб/м³"
d2=Reference(wm,min_col=4,min_row=mrow['u_tot'],max_col=9,max_row=mrow['u_tot'])
c2=Reference(wm,min_col=4,min_row=MHR,max_col=9,max_row=MHR)
chart2.add_data(d2,from_rows=True); chart2.set_categories(c2); chart2.legend=None
chart2.dataLabels=DataLabelList(); chart2.dataLabels.showVal=True; chart2.dataLabels.numFmt="0.0"
wm.add_chart(chart2,"K23")

chart3=BarChart(); chart3.type="col"; chart3.title="Стоимость экскаватора, тыс.руб"; chart3.height=8; chart3.width=11
d3=Reference(wm,min_col=4,min_row=mrow['price_rub'],max_col=9,max_row=mrow['price_rub'])
chart3.add_data(d3,from_rows=True); chart3.set_categories(c2); chart3.legend=None
wm.add_chart(chart3,"T23")
print("Сравнение:",mrow)

# ================================================================= #
# ЛИСТ: ЧУВСТВИТЕЛЬНОСТЬ
# ================================================================= #
ws=wb.create_sheet(S_SENS); ws.sheet_view.showGridLines=False
title(ws,"АНАЛИЗ ЧУВСТВИТЕЛЬНОСТИ","Торнадо ключевых драйверов и влияние производительности (КТГ) на аннуитет.")
ws.column_dimensions["A"].width=2; ws.column_dimensions["B"].width=34
for cc in "CDEFGH": ws.column_dimensions[cc].width=15

ws["B4"]="Анализируемый экскаватор (1–6):"; ws["B4"].font=FB
sel=ws["C4"]; sel.value=1; sel.fill=Fi; sel.border=BD; sel.alignment=C; sel.font=FB
SEL="$C$4"
dv2=DataValidation(type="whole",operator="between",formula1="1",formula2="6"); ws.add_data_validation(dv2); dv2.add(sel)
ws["D4"]=f"=INDEX('{S_IN}'!D{irow['name']}:I{irow['name']},{SEL})"; ws["D4"].font=Font(bold=True,color=ACC if False else "C55A11"); ws["D4"].alignment=L
ws.merge_cells("D4:H4")

def selann(key): return f"INDEX('{S_ANN}'!D{ar[key]}:I{ar[key]},{SEL})"
def selprod(key): return f"INDEX('{S_PROD}'!D{prd[key]}:I{prd[key]},{SEL})"

section(ws,6,"БАЗОВЫЕ ВЕЛИЧИНЫ ВЫБРАННОГО ВАРИАНТА","B","H")
base=[
 ("b_tot","Общий аннуитет","руб/м³",selann("u_tot"),M2),
 ("b_inv","— инвестиционный","руб/м³",selann("u_inv"),M2),
 ("b_fuel","— дизтопливо","руб/м³",selann("u_fuel"),M2),
 ("b_maint","— ТОиР и сервис","руб/м³",selann("u_maint"),M2),
 ("b_pers","— персонал","руб/м³",selann("u_pers"),M2),
 ("b_am3","Годовая производительность","тыс.м³",selprod("am3"),M),
 ("b_fix","Постоянные статьи (инв+ТОиР+перс+щит)","руб/м³",
   f"{selann('u_inv')}+{selann('u_maint')}+{selann('u_pers')}+{selann('u_tax')}",M2),
 ("b_var","Переменные статьи (топливо+ковш)","руб/м³",
   f"{selann('u_fuel')}+{selann('u_bucket')}",M2),
]
brow={}; b=7
for key,label,unit,fml,fmt in base:
    brow[key]=b
    ws[f"B{b}"]=label; ws[f"B{b}"].font=FN; ws[f"B{b}"].border=BD; ws[f"B{b}"].alignment=L
    ws[f"C{b}"]=unit; ws[f"C{b}"].font=FU; ws[f"C{b}"].border=BD; ws[f"C{b}"].alignment=C
    c=ws[f"D{b}"]; c.value="="+fml; c.number_format=fmt; c.border=BD; c.alignment=R; c.font=FN
    b+=1
def br(key): return f"$D${brow[key]}"
TAX=pc('tax'); DSF=pc('dsf')

# торнадо
tr=b+1; section(ws,tr,"ТОРНАДО: ВЛИЯНИЕ ДРАЙВЕРОВ НА ОБЩИЙ АННУИТЕТ (руб/м³)","B","H"); tr+=1
for i,h in enumerate(["Драйвер","Диапазон","При снижении","При росте","Размах"]):
    c=ws.cell(row=tr,column=2+i,value=h); c.font=FHd; c.fill=Fh; c.border=BD; c.alignment=C
tr+=1
# (label, pct, net-contribution formula using base rows)
drivers=[
 ("Расход / цена ДТ",0.20, f"{br('b_fuel')}*(1-{TAX})"),
 ("Затраты на ТОиР",0.25, f"{br('b_maint')}*(1-{TAX})"),
 ("Цена приобретения",0.15, f"{br('b_inv')}*(1-{TAX}*{DSF})"),
 ("Расходы на персонал",0.15, f"{br('b_pers')}*(1-{TAX})"),
]
ttop=tr
for label,pct,net in drivers:
    ws.cell(row=tr,column=2,value=label).font=FN; ws.cell(row=tr,column=2).border=BD; ws.cell(row=tr,column=2).alignment=L
    ws.cell(row=tr,column=3,value=f"±{int(pct*100)}%").alignment=C; ws.cell(row=tr,column=3).border=BD; ws.cell(row=tr,column=3).font=FN
    ws.cell(row=tr,column=4,value=f"={br('b_tot')}-{pct}*({net})")
    ws.cell(row=tr,column=5,value=f"={br('b_tot')}+{pct}*({net})")
    ws.cell(row=tr,column=6,value=f"=E{tr}-D{tr}")
    for col in (4,5,6):
        cc=ws.cell(row=tr,column=col); cc.number_format=M2; cc.border=BD; cc.alignment=R; cc.font=FN
    tr+=1
tbot=tr-1
# вспом. отклонения для диаграммы
ws.cell(row=ttop-1,column=9,value="низкое").font=FU; ws.cell(row=ttop-1,column=10,value="высокое").font=FU
for i in range(ttop,tbot+1):
    ws.cell(row=i,column=9,value=f"=D{i}-{br('b_tot')}").number_format=M2
    ws.cell(row=i,column=10,value=f"=E{i}-{br('b_tot')}").number_format=M2
tch=BarChart(); tch.type="bar"; tch.grouping="stacked"; tch.overlap=100
tch.title="Торнадо: отклонение аннуитета от базового, руб/м³"; tch.height=6.5; tch.width=15
dl=Reference(ws,min_col=9,min_row=ttop,max_col=9,max_row=tbot)
dh=Reference(ws,min_col=10,min_row=ttop,max_col=10,max_row=tbot)
cats=Reference(ws,min_col=2,min_row=ttop,max_col=2,max_row=tbot)
tch.add_data(dl); tch.add_data(dh); tch.set_categories(cats)
tch.series[0].graphicalProperties.solidFill="70AD47"; tch.series[1].graphicalProperties.solidFill="C55A11"
tch.legend=None
ws.add_chart(tch,"J8")

# чувствительность к КТГ / производительности
ur=tbot+3; section(ws,ur,"ВЛИЯНИЕ ПРОИЗВОДИТЕЛЬНОСТИ (КТГ / УТИЛИЗАЦИИ)","B","H"); ur+=1
ws.cell(row=ur,column=2,value="Изменение производительности").font=FHd
ws.cell(row=ur,column=2).fill=Fh; ws.cell(row=ur,column=2).border=BD; ws.cell(row=ur,column=2).alignment=C
levels=[-0.15,-0.10,0.0,0.10,0.15]
for j,d in enumerate(levels):
    c=ws.cell(row=ur,column=3+j,value=f"{'+' if d>0 else ''}{int(d*100)}%")
    c.font=FHd; c.fill=Fh; c.border=BD; c.alignment=C
ur+=1
ws.cell(row=ur,column=2,value="Общий аннуитет, руб/м³").font=FB
ws.cell(row=ur,column=2).border=BD; ws.cell(row=ur,column=2).alignment=L
for j,d in enumerate(levels):
    # постоянные статьи делятся на больший объём, переменные ~const на м³
    c=ws.cell(row=ur,column=3+j,value=f"={br('b_fix')}/(1+{d})+{br('b_var')}")
    c.number_format=M2; c.border=BD; c.alignment=R; c.font=(FB)
    if abs(d)<1e-9: c.fill=Fr
ur+=2
ws.cell(row=ur,column=2,value="Рост производительности (КТГ, утилизация) распределяет постоянные затраты "
        "на больший объём — удельный аннуитет снижается. Это ключевой рычаг эффективности.").font=FU
ws.merge_cells(f"B{ur}:H{ur}"); ws[f"B{ur}"].alignment=L
print("Чувствительность готова")

# ================================================================= #
# ЛИСТ: МЕТОДИКА (первым)
# ================================================================= #
wme=wb.create_sheet(S_MET,0); wme.sheet_view.showGridLines=False
title(wme,"МЕТОДИКА: АННУИТЕТ ЗАТРАТ (РУБ/М³) НА ВЛАДЕНИЕ ЭКСКАВАТОРАМИ","Пост-налоговый дисконтированный аннуитет жизненного цикла для сравнения машин одного класса.")
wme.column_dimensions["A"].width=2; wme.column_dimensions["B"].width=4; wme.column_dimensions["C"].width=120
def mh(row,t):
    c=wme.cell(row=row,column=3,value=t); c.font=FSec; c.fill=Fm; c.alignment=L; wme.cell(row=row,column=2).fill=Fm
def mt(row,t,h=None):
    c=wme.cell(row=row,column=3,value=t); c.font=FN; c.alignment=Alignment("left","top",wrap_text=True)
    if h: wme.row_dimensions[row].height=h
blocks=[
 (4,"h","1. НАЗНАЧЕНИЕ"),
 (5,"t","Экономическое сравнение нескольких экскаваторов одного класса по удельному аннуитету затрат (руб/м³ вынутой "
       "горной массы) за полный срок эксплуатации. Метрика позволяет корректно сопоставлять машины с разной ценой, "
       "производительностью и структурой затрат. Лучшим считается вариант с минимальным аннуитетом руб/м³.",60),
 (6,"h","2. ПРОИЗВОДИТЕЛЬНОСТЬ"),
 (7,"t","Порода в ковше = объём ковша × плотность × (наполнение/разрыхление) × коэф. потерь. Ковшей на самосвал = "
       "ОКРУГЛ.ВВЕРХ(грузоподъёмность / порода в ковше). Самосвалов в час = 3600 / (ковшей × время цикла + время "
       "постановки а/с). Часовая производительность (т/ч) = грузоподъёмность × самосвалов в час; м³/ч = т/ч / плотность.",60),
 (8,"t","Годовой объём (тыс.м³) = м³/ч × эффективное время. Эффективное время = КФВ × КТГ × (1 − простои/смена). "
       "КТГ — коэф. технической готовности (надёжность), КИО — коэф. использования = эффективное время / КФВ.",45),
 (9,"h","3. ЗАТРАТЫ ЖИЗНЕННОГО ЦИКЛА"),
 (10,"t","• Инвестиции: цена машины (пересчёт валюты в рубли по курсу) + монтаж/ПНР.",None),
 (11,"t","• Дизтопливо = эффективное время × удельный расход (кг/ч) × цена (руб/кг).",None),
 (12,"t","• ТОиР и сервис: плановое ТО, ремонты, ППР (среднегодовые).",None),
 (13,"t","• Расходники на ковш (зубья, коронки) — удельные, руб/м³.",None),
 (14,"t","• Расходы на персонал (экипаж, ФОТ с отчислениями).",None),
 (15,"t","• Амортизация (налоговый учёт) и налог на прибыль — формируют налоговый щит.",None),
 (16,"h","4. АННУИТЕТ (ЯДРО МЕТОДИКИ)"),
 (17,"t","Все потоки приводятся к текущей стоимости по ставке дисконтирования, затем преобразуются в эквивалентный "
        "годовой платёж (аннуитет): Аннуитет = NPV затрат / Σ(дисконт-факторов). Модель пост-налоговая: годовой поток "
        "= денежные затраты × (1 − ставка налога) − налоговый щит амортизации.",55),
 (18,"t","Инвестиционный аннуитет = Стоимость / Σ(дисконт-факторов) — возврат капитала. Операционный аннуитет = "
        "затраты после налога. Общий аннуитет = инвестиционный + операционный. Удельный аннуитет (руб/м³) = "
        "общий аннуитет / годовой объём производства.",45),
 (19,"t","Налоговый щит амортизации приведён через коэффициент амортизации (линейная за срок налогового учёта), "
        "поэтому машины с более высокой ценой получают больший вычет в первые годы.",45),
 (20,"h","5. ПОРЯДОК РАБОТЫ"),
 (21,"t","Шаг 1. «Параметры» — ставка, налог, курсы валют, цена ДТ, параметры карьера (плотность, самосвал, простои).",None),
 (22,"t","Шаг 2. «Ввод данных» — по каждому экскаватору: цена, ковш, цикл, КТГ, расход ДТ, годовые затраты.",None),
 (23,"t","Шаг 3. «Производительность» и «Расчёт аннуитета» — авторасчёт (просмотр).",None),
 (24,"t","Шаг 4. «Сравнение» — рейтинг, рекомендация, диаграммы. «Чувствительность» — торнадо и влияние КТГ.",None),
 (25,"h","6. ДОПУЩЕНИЯ"),
 (26,"t","• Машины сравниваются в одном классе и одной горнотехнической задаче; разница в производительности учтена "
        "через удельный аннуитет руб/м³.",30),
 (27,"t","• Годовые затраты приняты усреднёнными по сроку службы (при необходимости отразите капремонты через "
        "среднегодовые затраты ТОиР). Ставка и цены — согласованы (номинальные/реальные).",30),
 (28,"t","• Методика воспроизводит корпоративный подход «аннуитета руб/м³»; закрытая форма даёт результат, "
        "эквивалентный погодовому дисконтированному расчёту при равномерных затратах.",30),
]
for it in blocks:
    row,kind=it[0],it[1]
    (mh if kind=="h" else mt)(row, it[2], *( (it[3],) if kind=="t" and len(it)>3 and it[3] else () ))
# легенда
lg=30
wme.cell(row=lg,column=3,value="Обозначения:").font=FB
wme.cell(row=lg+1,column=2).fill=Fi; wme.cell(row=lg+1,column=2).border=BD
wme.cell(row=lg+1,column=3,value="жёлтые ячейки — ввод данных").font=FN
wme.cell(row=lg+2,column=2).fill=Fr; wme.cell(row=lg+2,column=2).border=BD
wme.cell(row=lg+2,column=3,value="зелёные ячейки — итоговые показатели / лучший вариант").font=FN

# порядок листов и сохранение
order=[S_MET,S_PAR,S_IN,S_PROD,S_ANN,S_CMP,S_SENS]
wb._sheets.sort(key=lambda s:order.index(s.title))
wb.active=0
wb.calculation.fullCalcOnLoad=True
OUT="Модель_TCO_экскаваторов_аннуитет.xlsx"
wb.save(OUT)
print("СОХРАНЕНО:",OUT)




