from flask import Flask, request, jsonify
from flask_cors import CORS
import openpyxl, io, base64, math
from openpyxl.styles import Alignment, Font
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

app = Flask(__name__)
CORS(app)

TEMPLATE_B64 = open('template.b64').read().strip()
CAP_MESES = {12:1, 4:3, 2:6}

# ─── helpers ────────────────────────────────────────────
def set_cell_bg(cell, color_hex):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color_hex)
    tcPr.append(shd)

def fmt_usd(v):
    return f'${abs(v):,.0f}' if v >= 0 else f'–${abs(v):,.0f}'

def fmt_pct(v, d=4):
    return f'{v*100:.{d}f}%'

def saldo_f(hip, ta, cu, mes):
    s = hip
    for _ in range(int(mes)):
        s -= (cu - s * ta)
    return max(s, 0)

def vf_f(tb, n, pmt):
    return pmt * (math.pow(1 + tb, n) - 1) / tb

def calc_escenario(precio, hip, fon, ta, tb, cu, diff, h, a):
    n = int(h * 12)
    s = saldo_f(hip, ta, cu, n)
    pv = precio * math.pow(1 + a / 100, h)
    inN = pv - s
    gB = inN - fon
    vf = vf_f(tb, n, diff) if n > 0 and diff != 0 else 0
    rT = gB - vf
    vpn = rT / math.pow(1 + tb, n)
    return {'s': s, 'pv': pv, 'inN': inN, 'gB': gB, 'vf': vf, 'rT': rT, 'vpn': vpn}

def parse_datos(d):
    precio = float(d['precio'])
    pct_ini = float(d['pctIni']) / 100
    gastos = float(d['gastos'])
    tasa = float(d['tasa']) / 100
    cap = int(d['cap'])
    plazo = int(d['plazo'])
    cok = float(d['cok']) / 100
    alq = float(d['alq'])
    mnt = float(d['mnt'])
    imp = float(d['imp'])
    mng = float(d['mng'])
    h1, h2, h3 = float(d['h1']), float(d['h2']), float(d['h3'])
    a1, a2, a3, a4 = float(d['a1']), float(d['a2']), float(d['a3']), float(d['a4'])

    if cap == 12: ta = tasa / 12
    elif cap == 4: ta = math.pow(1 + tasa / 4, 1/3) - 1
    else: ta = math.pow(1 + tasa / 2, 1/6) - 1
    tb = math.pow(1 + cok, 1/12) - 1

    ini = precio * pct_ini
    hip = precio - ini
    n_meses = plazo * 12
    cu_f = math.pow(1 + ta, n_meses)
    cu = hip * ta * cu_f / (cu_f - 1)
    fon = ini + gastos
    cok_m = fon * tb
    gi = mnt + imp + mng
    tbuy = cu + gi + cok_m
    diff = tbuy - alq
    hors = [h1, h2, h3]
    aprs = [a1, a2, a3, a4]
    return dict(precio=precio, pct_ini=pct_ini, gastos=gastos, tasa=tasa, cap=cap,
                plazo=plazo, cok=cok, alq=alq, mnt=mnt, imp=imp, mng=mng,
                ta=ta, tb=tb, ini=ini, hip=hip, n_meses=n_meses, cu=cu,
                fon=fon, cok_m=cok_m, gi=gi, tbuy=tbuy, diff=diff,
                hors=hors, aprs=aprs,
                nombre=d.get('nombre',''), ciudad=d.get('ciudad',''), n_equipo=d.get('n',1))

# ─── EXCEL ──────────────────────────────────────────────
@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok'})

@app.route('/generar-excel', methods=['POST'])
def generar_excel():
    try:
        d = request.json
        orig = base64.b64decode(TEMPLATE_B64)
        wb = openpyxl.load_workbook(io.BytesIO(orig))
        ws1 = wb['Datos']
        ws2 = wb['Matriz de Sensibilidad']
        ws1.comments = {}
        ws2.comments = {}
        p = parse_datos(d)
        cm = CAP_MESES.get(int(d.get('cap', 12)), 1)
        ws1['C7'] = float(d['tasa']) / 100
        ws1['E7'] = 12
        ws1['E8'] = cm
        ws1['E12'] = cm
        ws1['E13'] = 1
        ws1['H7'] = float(d['cok']) / 100
        ws1['E18'] = float(d['precio'])
        ws1['E19'] = float(d['pctIni']) / 100
        ws1['E22'] = int(d['plazo'])
        ws1['E28'] = float(d['gastos'])
        ws1['E34'] = float(d['alq'])
        ws1['E38'] = float(d['mnt'])
        ws1['E39'] = float(d['imp'])
        ws1['E40'] = float(d['mng'])
        ws1['E46'] = float(d['h1'])
        ws1['E48'] = float(d['h2'])
        ws1['E50'] = float(d['h3'])
        ws2['D3'] = float(d['a1']) / 100
        ws2['E3'] = float(d['a2']) / 100
        ws2['F3'] = float(d['a3']) / 100
        ws2['G3'] = float(d['a4']) / 100
        wrap = Alignment(wrap_text=True, vertical='top')
        for row in ws1.iter_rows():
            for cell in row:
                if cell.column == 11 and cell.value:
                    cell.alignment = wrap
        font_base = Font(name='Aptos Narrow', size=11)
        ws2['F3'].font = font_base
        ws2['G3'].font = font_base
        for col, w in [('A',10.664),('E',13.109),('F',12.219),('J',10.664),('K',10.887),('L',10.664),('N',18.777),('O',10.664)]:
            ws1.column_dimensions[col].width = w
        for col, w in [('A',10.664),('B',16.664),('C',2.887),('D',12.219),('H',10.664)]:
            ws2.column_dimensions[col].width = w
        for row, h in [(24,48.0),(44,30.0),(69,45.0),(71,50.25)]:
            ws1.row_dimensions[row].height = h
        ws2.row_dimensions[6].height = 22.5
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return jsonify({'status': 'ok', 'file': base64.b64encode(buf.read()).decode('utf-8')})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

# ─── WORD ───────────────────────────────────────────────
@app.route('/generar-word', methods=['POST'])
def generar_word():
    try:
        d = request.json
        p = parse_datos(d)
        precio,pct_ini,gastos,tasa,cap,plazo,cok = p['precio'],p['pct_ini'],p['gastos'],p['tasa'],p['cap'],p['plazo'],p['cok']
        alq,mnt,imp,mng = p['alq'],p['mnt'],p['imp'],p['mng']
        ta,tb,ini,hip,n_meses,cu = p['ta'],p['tb'],p['ini'],p['hip'],p['n_meses'],p['cu']
        fon,cok_m,gi,tbuy,diff = p['fon'],p['cok_m'],p['gi'],p['tbuy'],p['diff']
        hors,aprs = p['hors'],p['aprs']
        nombre,ciudad,n_equipo = p['nombre'],p['ciudad'],p['n_equipo']
        n1 = nombre.split()[0] if nombre else 'El cliente'
        cap_nom = {12:'mensual',4:'trimestral',2:'semestral'}[cap]
        if cap==12: ta_formula = f'TEM_A = TNA {tasa*100:.1f}%÷12 = {fmt_pct(ta)}'
        elif cap==4: ta_formula = f'TET={tasa*100:.1f}%/4={fmt_pct(tasa/4)}; TEM_A=(1+TET)^(1/3)–1={fmt_pct(ta)}'
        else: ta_formula = f'TES={tasa*100:.1f}%/2={fmt_pct(tasa/2)}; TEM_A=(1+TES)^(1/6)–1={fmt_pct(ta)}'

        saldos = {h: saldo_f(hip, ta, cu, int(h*12)) for h in hors}
        matriz = {h: {a: calc_escenario(precio,hip,fon,ta,tb,cu,diff,h,a)['vpn'] for a in aprs} for h in hors}
        todos_vpn = [(h,a,matriz[h][a]) for h in hors for a in aprs]
        todos_neg = all(v<0 for _,_,v in todos_vpn)
        mejor = max(todos_vpn, key=lambda x: x[2])
        eq = {}
        for h in hors:
            n = int(h*12)
            s = saldos[h]
            vf = vf_f(tb,n,diff) if n>0 and diff!=0 else 0
            p_req = s + fon + vf
            t_req = math.pow(p_req/precio,1/h)-1 if precio>0 else 0
            eq[h] = {'p_req':p_req,'t_req':t_req}

        doc = Document()
        for section in doc.sections:
            section.top_margin = Inches(1)
            section.bottom_margin = Inches(1)
            section.left_margin = Inches(1.18)
            section.right_margin = Inches(1.18)
        doc.styles['Normal'].font.name = 'Arial'
        doc.styles['Normal'].font.size = Pt(11)

        def heading(text, size=13, color='1F3864'):
            pp = doc.add_paragraph()
            pp.paragraph_format.space_before = Pt(10)
            pp.paragraph_format.space_after = Pt(4)
            r = pp.add_run(text)
            r.bold = True; r.font.size = Pt(size)
            r.font.color.rgb = RGBColor.from_string(color)
            return pp

        def note_box(text):
            t = doc.add_table(rows=1, cols=1)
            t.style = 'Table Grid'
            c = t.cell(0,0)
            set_cell_bg(c, 'DEEAF1')
            pp = c.paragraphs[0]
            pp.paragraph_format.space_before = Pt(3)
            pp.paragraph_format.space_after = Pt(3)
            r = pp.add_run(text)
            r.font.size = Pt(9); r.italic = True
            doc.add_paragraph()

        def item(text, bold_prefix=None):
            pp = doc.add_paragraph(style='List Number')
            pp.paragraph_format.space_after = Pt(4)
            if bold_prefix and text.startswith(bold_prefix):
                idx = text.index('—')+1 if '—' in text else len(bold_prefix)
                r1 = pp.add_run(text[:idx]); r1.bold=True; r1.font.size=Pt(11)
                r2 = pp.add_run(text[idx:]); r2.font.size=Pt(11)
            else:
                r = pp.add_run(text); r.font.size=Pt(11)

        def hdr_row(tbl, headers, bg='2E75B6'):
            for j,h in enumerate(headers):
                c = tbl.rows[0].cells[j]
                set_cell_bg(c,'2E75B6')
                r = c.paragraphs[0].add_run(h)
                r.bold=True; r.font.color.rgb=RGBColor(255,255,255); r.font.size=Pt(10)

        # CARATULA
        for txt,sz,bold in [
            ('ANÁLISIS DE DECISIÓN COMPRAR VS. RENTAR',16,True),
            (f'Caso {nombre}: {ciudad}',13,True),
            ('Metodología con dos tasas diferenciadas',11,False),
            ('','11',False),
            (f'Equipo {n_equipo}',11,True),
            ('Curso: Gestión Financiera (FIN83010)',10,False),
            ('Docente: Mag. Juan Antonio Lillo Paredes',10,False),
            ('Escuela de Postgrado — Universidad San Ignacio de Loyola',10,False),
        ]:
            pp = doc.add_paragraph()
            pp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if txt:
                r = pp.add_run(txt); r.bold=bold
                r.font.size = Pt(int(sz)) if isinstance(sz,str) else Pt(sz)
                if txt == 'ANÁLISIS DE DECISIÓN COMPRAR VS. RENTAR':
                    r.font.color.rgb = RGBColor.from_string('1F3864')
        doc.add_page_break()

        note_box(f'Nota metodológica: Este caso usa dos tasas diferenciadas. TASA A hipotecaria (TEM {fmt_pct(ta)}) para cuota y saldo. TASA B de oportunidad (TEM {fmt_pct(tb)}) para costo de capital, acumulación de sobrecostos y descuento del VPN. Sin gastos de venta en este análisis.')

        heading('RESUMEN EJECUTIVO')
        items_re = [
            f'El presente análisis evalúa la decisión financiera de {nombre} entre comprar un inmueble de {fmt_usd(precio)} en {ciudad} o continuar rentando por {fmt_usd(alq)} mensuales, aplicando los conceptos del valor del dinero en el tiempo del caso Harvard W14404 con la extensión metodológica de dos tasas diferenciadas.',
            f'La metodología emplea dos tasas: TASA A hipotecaria (TEM={fmt_pct(ta)}) derivada de TNA {tasa*100:.1f}% capitalizable {cap_nom}; y TASA B de oportunidad (TEM={fmt_pct(tb)}) derivada del COK {cok*100:.0f}% anual.',
            f'Los resultados muestran VPN entre {fmt_usd(min(todos_vpn,key=lambda x:x[2])[2])} y {fmt_usd(mejor[2])}. El escenario {"menos desfavorable" if todos_neg else "más favorable"} es {int(mejor[0])} años con {mejor[1]:.1f}% de apreciación (VPN = {fmt_usd(mejor[2])}).',
            f'Recomendación: {"Continuar rentando. Para VPN=0 en " + str(int(hors[0])) + " años se requiere " + f"{eq[hors[0]]["t_req"]*100:.1f}% de apreciación anual sostenida." if todos_neg else "La compra puede ser viable con horizontes largos y apreciación alta."}',
        ]
        for t in items_re: item(t)
        doc.add_page_break()

        heading('2.1  FUNDAMENTOS FINANCIEROS  (Miembro 1)')
        note_box('Pasos 1 a 4 — Tasas, cuota, costo de oportunidad y diferencia mensual')
        items_21 = [
            f'El caso {nombre} parte de: inmueble {fmt_usd(precio)}, renta {fmt_usd(alq)}/mes, enganche {pct_ini*100:.0f}% ({fmt_usd(ini)}), gastos {fmt_usd(gastos)}, hipoteca {tasa*100:.1f}% capitalizable {cap_nom} a {plazo} años, COK {cok*100:.0f}% anual.',
            f'PASO 1 — Tasas: {ta_formula}. COK: {fmt_pct(tb)} mensual. Valores completos: TASA A={ta:.10f}; TASA B={tb:.10f}.',
            f'PASO 2 — Cuota: hipoteca {fmt_usd(hip)} a {n_meses} meses con TASA A. =PAGO({fmt_pct(ta)},{n_meses},–{fmt_usd(hip)}) = {fmt_usd(cu)}/mes.',
            f'PASO 3 — Costo de oportunidad: {fmt_usd(fon)} inmovilizados × TASA B {fmt_pct(tb)} = {fmt_usd(cok_m)}/mes.',
            'PASO 4 — Diferencia mensual:',
        ]
        for t in items_21: item(t)

        tbl4 = doc.add_table(rows=9, cols=3)
        tbl4.style = 'Table Grid'
        hdr_row(tbl4, ['Concepto','Monto mensual','Tasa aplicada'])
        filas4 = [
            (f'Cuota hipotecaria (=PAGO, TASA A)', fmt_usd(cu), f'TASA A {fmt_pct(ta)}'),
            ('Mantenimiento', fmt_usd(mnt), 'Dato'),
            ('Impuesto predial', fmt_usd(imp), 'Dato'),
            ('Mantenimiento general', fmt_usd(mng), 'Dato'),
            (f'Costo de oportunidad (TASA B)', fmt_usd(cok_m), f'TASA B {fmt_pct(tb)}'),
            ('Total costo mensual comprar', fmt_usd(tbuy), ''),
            ('Menos: renta alternativa', f'–{fmt_usd(alq)}', 'Dato'),
            ('DIFERENCIA MENSUAL (sobrecosto)', fmt_usd(diff), 'Flujo clave'),
        ]
        for i,(c1,c2,c3) in enumerate(filas4):
            rr = tbl4.rows[i+1]
            bg = 'FFE699' if 'DIFERENCIA' in c1 or 'Total' in c1 else ('DEEAF1' if i%2==0 else 'FFFFFF')
            for j,txt in enumerate([c1,c2,c3]):
                cc = rr.cells[j]; set_cell_bg(cc,bg)
                r = cc.paragraphs[0].add_run(txt)
                r.bold = 'DIFERENCIA' in c1 or 'Total' in c1; r.font.size=Pt(10)
        doc.add_paragraph()
        items_21b = [
            f'La diferencia mensual de {fmt_usd(diff)} es el flujo incremental fundamental. Acumulada con TASA B durante el horizonte debe ser superada por la ganancia neta de venta para que comprar sea conveniente.',
            f'Fondos inmovilizados: {fmt_usd(fon)} ({pct_ini*100:.0f}% del precio). Sin gastos de venta al final, los ingresos netos de venta son el precio futuro menos el saldo hipotecario pendiente.',
            f'La barrera acumulada (VF sobrecostos) crece exponencialmente. A mayor TASA B y/o horizonte más largo, más alta la barrera que la apreciación debe superar.',
        ]
        for t in items_21b: item(t)
        doc.add_page_break()

        heading('2.2  ANÁLISIS TEMPORAL Y ESCENARIOS  (Miembro 2)')
        note_box('Pasos 5 y 6 — Saldos hipotecarios, VPN por horizonte y escenario')

        items_22a = [
            f'Se analizan {len(hors)} horizontes ({", ".join(str(int(h))+" años" for h in hors)}) y {len(aprs)} escenarios de apreciación ({", ".join(str(a)+"%" for a in aprs)}), generando {len(hors)*len(aprs)} combinaciones. Saldos con TASA A, VPN con TASA B.',
            'PASO 5 — Saldos hipotecarios pendientes por horizonte (=VA con TASA A):',
        ]
        for t in items_22a: item(t)

        tbl5 = doc.add_table(rows=len(hors)+1, cols=4)
        tbl5.style = 'Table Grid'
        hdr_row(tbl5, ['Horizonte','Meses pagados','Meses restantes','Saldo pendiente (TASA A)'])
        for i,h in enumerate(hors):
            rr = tbl5.rows[i+1]
            s = saldos[h]; mp=int(h*12); mr=n_meses-mp
            for j,txt in enumerate([f'{int(h)} años',str(mp),str(mr),fmt_usd(s)]):
                cc=rr.cells[j]; set_cell_bg(cc,'DEEAF1' if i%2==0 else 'FFFFFF')
                r=cc.paragraphs[0].add_run(txt); r.bold=j==3; r.font.size=Pt(10)
        doc.add_paragraph()

        h_ej,a_ej = hors[0],aprs[2]
        r_ej = calc_escenario(precio,hip,fon,ta,tb,cu,diff,h_ej,a_ej)
        n_ej = int(h_ej*12)
        item(f'PASO 6 — Secuencia VPN (ejemplo: {int(h_ej)} años, {a_ej:.1f}% apreciación):')

        tbl6 = doc.add_table(rows=7, cols=3)
        tbl6.style = 'Table Grid'
        hdr_row(tbl6, ['#','Operación','Cálculo y resultado'])
        filas6 = [
            ('a','Precio de venta futuro',f'{fmt_usd(precio)}×(1+{a_ej:.1f}%)^{int(h_ej)} = {fmt_usd(r_ej["pv"])}'),
            ('b','Ingreso neto (precio – saldo TASA A)',f'{fmt_usd(r_ej["pv"])} – {fmt_usd(r_ej["s"])} = {fmt_usd(r_ej["inN"])}'),
            ('c','Ganancia bruta',f'{fmt_usd(r_ej["inN"])} – {fmt_usd(fon)} = {fmt_usd(r_ej["gB"])}'),
            ('d','VF sobrecostos (TASA B)',f'=VF({fmt_pct(tb)},{n_ej},–{fmt_usd(diff)}) = {fmt_usd(r_ej["vf"])}'),
            ('e','Resultado neto en t=N',f'{fmt_usd(r_ej["gB"])} – {fmt_usd(r_ej["vf"])} = {fmt_usd(r_ej["rT"])}'),
            ('f','VPN hoy (TASA B)',f'=VA({fmt_pct(tb)},{n_ej},0,{fmt_usd(r_ej["rT"])}) = {fmt_usd(r_ej["vpn"])}'),
        ]
        for i,(a,b,c) in enumerate(filas6):
            rr=tbl6.rows[i+1]; bg='FFE699' if i==5 else ('DEEAF1' if i%2==0 else 'FFFFFF')
            for j,txt in enumerate([a,b,c]):
                cc=rr.cells[j]; set_cell_bg(cc,bg)
                r=cc.paragraphs[0].add_run(txt); r.bold=i==5; r.font.size=Pt(9)
        doc.add_paragraph()

        item('Tabla de sensibilidad completa — VPN en el período 0:')
        tbl_mx = doc.add_table(rows=len(hors)+1, cols=len(aprs)+1)
        tbl_mx.style = 'Table Grid'
        c0 = tbl_mx.rows[0].cells[0]; set_cell_bg(c0,'2E75B6')
        r0 = c0.paragraphs[0].add_run('Horizonte \\ Apreciación'); r0.bold=True; r0.font.color.rgb=RGBColor(255,255,255); r0.font.size=Pt(9)
        for j,a in enumerate(aprs):
            cc=tbl_mx.rows[0].cells[j+1]; set_cell_bg(cc,'2E75B6')
            r=cc.paragraphs[0].add_run(f'{a:.1f}%'); r.bold=True; r.font.color.rgb=RGBColor(255,255,255); r.font.size=Pt(10)
        for i,h in enumerate(hors):
            rr=tbl_mx.rows[i+1]
            cc=rr.cells[0]; set_cell_bg(cc,'D9E1F2')
            r=cc.paragraphs[0].add_run(f'{int(h)} años'); r.bold=True; r.font.size=Pt(10)
            for j,a in enumerate(aprs):
                vpn=matriz[h][a]; cc=rr.cells[j+1]
                bg='E2EFDA' if vpn>0 else ('FFE7E7' if vpn<-50000 else 'FFF2CC')
                set_cell_bg(cc,bg)
                r=cc.paragraphs[0].add_run(fmt_usd(vpn)); r.bold=True; r.font.size=Pt(10)
                r.font.color.rgb=RGBColor(0x27,0x50,0x0a) if vpn>0 else RGBColor(0x7b,0x1c,0x14)
        doc.add_paragraph()

        item('Punto de equilibrio — apreciación mínima para VPN = 0:')
        tbl_eq = doc.add_table(rows=len(hors)+1, cols=3)
        tbl_eq.style = 'Table Grid'
        hdr_row(tbl_eq, ['Horizonte','Precio de venta requerido','Apreciación de equilibrio'])
        for i,h in enumerate(hors):
            rr=tbl_eq.rows[i+1]; e=eq[h]
            for j,txt in enumerate([f'{int(h)} años',fmt_usd(e['p_req']),f'{e["t_req"]*100:.2f}% anual']):
                cc=rr.cells[j]; set_cell_bg(cc,'FFE699' if j==2 else ('DEEAF1' if i%2==0 else 'FFFFFF'))
                r=cc.paragraphs[0].add_run(txt); r.bold=j==2; r.font.size=Pt(10)
        doc.add_paragraph()
        doc.add_page_break()

        heading('2.3  EVALUACIÓN Y RECOMENDACIÓN  (Miembro 3)')
        note_box('Sensibilidad al COK, factores cualitativos y recomendación final')

        item(f'La evaluación integral {"confirma que ningún escenario genera VPN positivo" if todos_neg else "muestra escenarios con VPN positivo"} para {n1}. La tasa de oportunidad del {cok*100:.0f}% anual es el parámetro más crítico del modelo.')

        item(f'Sensibilidad al COK — escenario referencia ({int(hors[0])} años, {aprs[2]:.1f}% apreciación):')
        cok_pct = cok*100
        cok_vars = [c for c in [cok_pct-4,cok_pct-2,cok_pct,cok_pct+2,cok_pct+4] if c>0]
        h_ref,a_ref = hors[0],aprs[2]
        tbl_cok = doc.add_table(rows=len(cok_vars)+1, cols=4)
        tbl_cok.style = 'Table Grid'
        hdr_row(tbl_cok, ['COK anual','TEM_B mensual','VF sobrecostos','VPN'])
        for i,ck in enumerate(cok_vars):
            tb2=math.pow(1+ck/100,1/12)-1; cop2=fon*tb2; diff2=cu+gi+cop2-alq
            n_r=int(h_ref*12); s2=saldos[h_ref]
            vf2=vf_f(tb2,n_r,diff2) if n_r>0 and diff2!=0 else 0
            pv2=precio*math.pow(1+a_ref/100,h_ref); gB2=pv2-s2-fon
            rT2=gB2-vf2; vpn2=rT2/math.pow(1+tb2,n_r)
            is_base=abs(ck-cok_pct)<0.01
            rr=tbl_cok.rows[i+1]; bg='FFE699' if is_base else ('DEEAF1' if i%2==0 else 'FFFFFF')
            for j,txt in enumerate([f'{ck:.0f}%{"(base)" if is_base else ""}',fmt_pct(tb2),fmt_usd(vf2),fmt_usd(vpn2)]):
                cc=rr.cells[j]; set_cell_bg(cc,bg)
                r=cc.paragraphs[0].add_run(txt); r.bold=is_base; r.font.size=Pt(10)
        doc.add_paragraph()

        items_23b = [
            f'Factores cualitativos favorables a la compra: estabilidad habitacional, cobertura contra inflación, ahorro forzoso mediante amortización de capital, potencial de apreciación del mercado de {ciudad}.',
            f'Factores cualitativos favorables a rentar: diferencia mensual de {fmt_usd(diff)} puede comprometer la liquidez, mayor flexibilidad de movilidad, capital de {fmt_usd(fon)} concentrado en un solo activo.',
            f'Para que comprar sea conveniente en el horizonte de {int(hors[0])} años se requiere: precio futuro de {fmt_usd(eq[hors[0]]["p_req"])}, equivalente a apreciación sostenida del {eq[hors[0]]["t_req"]*100:.2f}% anual.',
            f'Recomendación final: {"Se recomienda a " + n1 + " continuar rentando. Los " + str(len(todos_vpn)) + " escenarios analizados generan VPN negativo. La compra requiere apreciación del " + f"{eq[hors[0]]["t_req"]*100:.1f}%" + "% anual sostenida en " + str(int(hors[0])) + " años, horizonte de " + str(int(hors[-1])) + "+ años, o COK sustancialmente menor al " + f"{cok*100:.0f}%" + " actual." if todos_neg else "Evaluar compra bajo horizontes de " + str(int(hors[-1])) + "+ años con apreciación del " + f"{aprs[-1]:.1f}%" + "% anual."}',
        ]
        for t in items_23b: item(t)
        doc.add_page_break()

        heading('REFERENCIAS BIBLIOGRÁFICAS')
        refs = [
            'Cleary, S. y Foerster, S. (2014). Teaching Note W14404: Time Value of Money: The Buy Versus Rent Decision. Ivey Business School, Western University.',
            f'Documento del curso. (2026). Cálculos Detallados — Caso {nombre}. FIC-Lab, Escuela de Postgrado USIL.',
            'Documento del curso. (2026). Guía Simplificada: Caso Rebecca Young — Comprar vs. Rentar. FIC-Lab, USIL.',
            'Brealey, R.A. (2020). Fundamentos de finanzas corporativas (13a ed.). McGraw-Hill.',
        ]
        for ref in refs:
            pp = doc.add_paragraph()
            pp.paragraph_format.left_indent = Inches(0.5)
            pp.paragraph_format.first_line_indent = Inches(-0.5)
            pp.paragraph_format.space_after = Pt(6)
            r = pp.add_run(ref); r.font.size=Pt(10)

        buf = io.BytesIO()
        doc.save(buf); buf.seek(0)
        return jsonify({'status':'ok','file':base64.b64encode(buf.read()).decode('utf-8')})
    except Exception as e:
        import traceback
        return jsonify({'status':'error','message':str(e),'trace':traceback.format_exc()}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
