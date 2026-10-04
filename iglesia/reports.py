"""Reportes de miembros (Excel/PDF, según permisos) y reporte completo (solo administrador)."""
import datetime
from io import BytesIO
from xml.sax.saxutils import escape

from flask import Blueprint, render_template, send_file, request

from .db import audit, get_db
from .security import NAME_SQL, admin_required, pastoral_required, excel_safe, login_required, scope, where

bp = Blueprint('reports', __name__)

XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'


def _sheet(ws, headers, rows):
    from openpyxl.styles import Font
    ws.append(headers)
    for c in ws[1]:
        c.font = Font(bold=True)
    ws.freeze_panes = 'A2'
    for r in rows:
        ws.append([excel_safe(v) for v in (r.values() if isinstance(r, dict) else r)])
    for col in ws.columns:
        width = max(len(str(c.value)) if c.value is not None else 0 for c in col[:200])
        ws.column_dimensions[col[0].column_letter].width = min(max(width + 2, 10), 60)


def filtered_members():
    db=get_db();sc,sp=scope('m')
    rows=db.execute('SELECT m.*,mi.Name AS MinistryName FROM Members m LEFT JOIN Ministries mi ON mi.Id=m.MinistryId'+where(sc)+' ORDER BY m.FullName',sp).fetchall()
    result=[];today=datetime.date.today()
    for r in rows:
        age=None
        try:
            birth=datetime.date.fromisoformat(r['BirthDate']);age=today.year-birth.year-((today.month,today.day)<(birth.month,birth.day))
        except (ValueError,TypeError):pass
        if request.args.get('ministry',type=int) and r['MinistryId']!=request.args.get('ministry',type=int):continue
        if request.args.get('status') and r['Status']!=request.args['status']:continue
        if request.args.get('baptized')=='1' and not r['BaptismDate']:continue
        if request.args.get('baptized')=='0' and r['BaptismDate']:continue
        if request.args.get('from') and (not r['JoinDate'] or r['JoinDate']<request.args['from']):continue
        if request.args.get('to') and (not r['JoinDate'] or r['JoinDate']>request.args['to']):continue
        low=request.args.get('min_age',type=int);high=request.args.get('max_age',type=int)
        if low is not None and (age is None or age<low):continue
        if high is not None and (age is None or age>high):continue
        result.append([r['FullName'],r['Cedula'],r['Phone'],r['MinistryName'],r['Status'],r['JoinDate'],r['BaptismDate'],age,r['ChurchRole']])
    return result

HEADERS=['Nombre','Cédula','Teléfono','Ministerio','Estado','Ingreso','Bautizo','Edad','Cargo']

@bp.route('/reports')
@login_required
def reports():
    return render_template('reports.html', title='REPORTES',ministries=get_db().execute("SELECT Id,Name FROM Ministries WHERE Name <> '__COMUNIDADES_ORACION__' ORDER BY Name").fetchall())

@bp.route('/reports/members.xlsx')
@login_required
def report_xlsx():
    from openpyxl import Workbook
    wb=Workbook();wb.active.title='Miembros';_sheet(wb.active,HEADERS,filtered_members())
    bio=BytesIO();wb.save(bio);bio.seek(0)
    audit('Exportar','Reporte',details='Miembros Excel filtrado');get_db().commit()
    return send_file(bio,as_attachment=True,download_name='reporte_miembros.xlsx',mimetype=XLSX_MIME)

@bp.route('/reports/members.pdf')
@login_required
def report_pdf():
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4,landscape
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.platypus import Paragraph,SimpleDocTemplate,Spacer,Table,TableStyle
    styles=getSampleStyleSheet();cell=styles['BodyText'];cell.fontSize=7;cell.leading=9
    def p(v):return Paragraph(escape(str(v if v is not None else '')),cell)
    data=[[p(v) for v in r] for r in [HEADERS]+filtered_members()]
    table=Table(data,repeatRows=1,colWidths=[145,85,75,95,65,65,65,35,65])
    table.setStyle(TableStyle([('GRID',(0,0),(-1,-1),.25,colors.grey),('BACKGROUND',(0,0),(-1,0),colors.lightgrey),('VALIGN',(0,0),(-1,-1),'TOP')]))
    bio=BytesIO();SimpleDocTemplate(bio,pagesize=landscape(A4),leftMargin=20,rightMargin=20).build([Paragraph('Sistema Iglesia Web · Reporte de miembros',styles['Heading2']),Spacer(1,12),table]);bio.seek(0)
    audit('Exportar','Reporte',details='Miembros PDF filtrado');get_db().commit()
    return send_file(bio,as_attachment=True,download_name='reporte_miembros.pdf',mimetype='application/pdf')


@bp.route('/reports/complete.xlsx')
@pastoral_required  # incluye todos los miembros y el historial: antes lo podía descargar cualquier usuario
def complete_xlsx():
    from openpyxl import Workbook
    db = get_db()
    wb = Workbook()
    wb.remove(wb.active)
    sets = [
        ('Miembros', f'SELECT {NAME_SQL},Cedula,Phone,Status,JoinDate,BaptismDate,ChurchRole FROM Members WHERE DeletedAt IS NULL ORDER BY 1 COLLATE NOCASE',
         ['Nombre', 'Cédula', 'Teléfono', 'Estado', 'Ingreso', 'Bautizo', 'Cargo']),
        ('Ministerios', "SELECT Name,Leader,Assistant,Notes FROM Ministries WHERE Name <> '__COMUNIDADES_ORACION__' ORDER BY Name",
         ['Ministerio', 'Encargado', 'Asistente', 'Notas']),
        ('Actividad', 'SELECT CreatedAt,Action,Entity,Details FROM AuditLog ORDER BY Id DESC',
         ['Fecha', 'Acción', 'Módulo', 'Detalle']),
        ('Conversiones', "SELECT c.FullName,c.ConversionDate,coalesce(mi.Name,''),coalesce(u.FullName,''),coalesce(c.Notes,'') FROM Conversions c LEFT JOIN Ministries mi ON mi.Id=c.MinistryId LEFT JOIN Users u ON u.Id=c.RegisteredBy ORDER BY c.ConversionDate DESC",
         ['Nombre','Fecha','Ministerio','Registrado por','Observación']),
    ]
    for title, sql, heads in sets:
        _sheet(wb.create_sheet(title), heads, db.execute(sql).fetchall())
    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    audit('Exportar', 'Reporte', details='Reporte completo Excel')
    db.commit()
    return send_file(bio, as_attachment=True, download_name='reporte_completo.xlsx', mimetype=XLSX_MIME)

def _baptism_rows():
    db=get_db(); sc,sp=scope('m')
    sql='''SELECT m.MemberCode,m.FullName,m.Phone,coalesce(mi.Name,'') MinistryName,m.BaptismDate,m.MinistryId
           FROM Members m LEFT JOIN Ministries mi ON mi.Id=m.MinistryId'''+where(sc,"m.BaptismDate IS NOT NULL AND m.BaptismDate<>''")
    rows=db.execute(sql+' ORDER BY m.BaptismDate DESC,m.FullName',sp).fetchall(); out=[]
    year=request.args.get('year',type=int); month=request.args.get('month',type=int); ministry=request.args.get('ministry',type=int); term=(request.args.get('q') or '').strip().lower()
    for r in rows:
        if year and not str(r['BaptismDate']).startswith(str(year)+'-'):continue
        if month and str(r['BaptismDate'])[5:7] != f'{month:02d}':continue
        if ministry and r['MinistryId']!=ministry:continue
        if term and term not in ((r['FullName'] or '')+' '+(r['MemberCode'] or '')).lower():continue
        out.append(r)
    return out

@bp.route('/reports/baptisms')
@login_required
def baptisms():
    rows=_baptism_rows(); db=get_db(); year=datetime.date.today().year
    sc,sp=scope('m'); all_rows=db.execute("SELECT m.BaptismDate FROM Members m"+where(sc,"m.BaptismDate IS NOT NULL AND m.BaptismDate<>''"),sp).fetchall()
    totals={'year':sum(1 for r in all_rows if str(r['BaptismDate']).startswith(str(year)+'-')),'all':len(all_rows)}
    years=sorted({int(str(r['BaptismDate'])[:4]) for r in all_rows if str(r['BaptismDate'])[:4].isdigit()}|{year},reverse=True)
    ministries=db.execute("SELECT Id,Name FROM Ministries WHERE Name <> '__COMUNIDADES_ORACION__' ORDER BY Name").fetchall()
    return render_template('baptisms_report.html',title='BAUTIZADOS',rows=rows,totals=totals,years=years,ministries=ministries)

@bp.route('/reports/baptisms.xlsx')
@login_required
def baptisms_xlsx():
    from openpyxl import Workbook
    rows=[[r['MemberCode'],r['FullName'],r['Phone'],r['MinistryName'],r['BaptismDate']] for r in _baptism_rows()]
    wb=Workbook(); wb.active.title='Bautizados'; _sheet(wb.active,['Código MEM','Nombre','Teléfono','Ministerio','Fecha de Bautismo'],rows)
    bio=BytesIO(); wb.save(bio); bio.seek(0); audit('Exportar','Reporte',details='Bautizados Excel'); get_db().commit()
    return send_file(bio,as_attachment=True,download_name='reporte_bautizados.xlsx',mimetype=XLSX_MIME)

@bp.route('/reports/baptisms.pdf')
@login_required
def baptisms_pdf():
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4,landscape
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.platypus import Paragraph,SimpleDocTemplate,Spacer,Table,TableStyle
    styles=getSampleStyleSheet(); cell=styles['BodyText']; cell.fontSize=8; cell.leading=10
    def pp(v):return Paragraph(escape(str(v or '')),cell)
    heads=['Código MEM','Nombre','Teléfono','Ministerio','Fecha de Bautismo']; rows=[[r['MemberCode'],r['FullName'],r['Phone'],r['MinistryName'],r['BaptismDate']] for r in _baptism_rows()]
    data=[[pp(v) for v in x] for x in [heads]+rows]; table=Table(data,repeatRows=1,colWidths=[75,180,100,150,100]); table.setStyle(TableStyle([('GRID',(0,0),(-1,-1),.25,colors.grey),('BACKGROUND',(0,0),(-1,0),colors.lightgrey),('VALIGN',(0,0),(-1,-1),'TOP')]))
    bio=BytesIO(); SimpleDocTemplate(bio,pagesize=landscape(A4),leftMargin=25,rightMargin=25).build([Paragraph('Sistema Iglesia Web · Bautizados',styles['Heading2']),Spacer(1,12),table]); bio.seek(0); audit('Exportar','Reporte',details='Bautizados PDF'); get_db().commit()
    return send_file(bio,as_attachment=True,download_name='reporte_bautizados.pdf',mimetype='application/pdf')
