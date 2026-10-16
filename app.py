import os, io, sqlite3, hashlib, json, uuid
from datetime import date, datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
import pandas as pd
import streamlit as st

APP_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(APP_DIR, 'celcius_sales_cost.db')
EXPORT_DIR = os.path.join(APP_DIR, 'exports')
os.makedirs(EXPORT_DIR, exist_ok=True)

st.set_page_config(page_title='Celcius Sales & Cost System', page_icon='📊', layout='wide')

# ---------- Database ----------
def db():
    con = sqlite3.connect(DB_PATH, check_same_thread=False)
    con.row_factory = sqlite3.Row
    return con

def init_db():
    with db() as c:
        c.executescript('''
        CREATE TABLE IF NOT EXISTS users(id INTEGER PRIMARY KEY, username TEXT UNIQUE, password_hash TEXT, role TEXT, active INTEGER DEFAULT 1, created_at TEXT);
        CREATE TABLE IF NOT EXISTS clients(id INTEGER PRIMARY KEY, name TEXT UNIQUE, created_at TEXT);
        CREATE TABLE IF NOT EXISTS projects(id INTEGER PRIMARY KEY, client_id INTEGER, name TEXT, product TEXT DEFAULT '', active INTEGER DEFAULT 1, created_at TEXT, UNIQUE(client_id,name), FOREIGN KEY(client_id) REFERENCES clients(id));
        CREATE TABLE IF NOT EXISTS categories(id INTEGER PRIMARY KEY, name TEXT UNIQUE, kind TEXT DEFAULT 'Both', active INTEGER DEFAULT 1);
        CREATE TABLE IF NOT EXISTS commercials(id INTEGER PRIMARY KEY, client_id INTEGER, project_id INTEGER, category TEXT, service TEXT DEFAULT '', source TEXT DEFAULT '', destination TEXT DEFAULT '', unit TEXT DEFAULT 'shipment', rate REAL, effective_from TEXT, effective_to TEXT DEFAULT '', status TEXT DEFAULT 'Pending', notes TEXT DEFAULT '', created_by TEXT, created_at TEXT, approved_by TEXT, approved_at TEXT, FOREIGN KEY(client_id) REFERENCES clients(id), FOREIGN KEY(project_id) REFERENCES projects(id));
        CREATE TABLE IF NOT EXISTS sales(id INTEGER PRIMARY KEY, entry_ref TEXT UNIQUE, client_id INTEGER, project_id INTEGER, service_date TEXT, category TEXT, service TEXT DEFAULT '', source TEXT DEFAULT '', destination TEXT DEFAULT '', shipment_ref TEXT DEFAULT '', quantity REAL DEFAULT 1, unit TEXT DEFAULT 'shipment', rate REAL DEFAULT 0, amount REAL DEFAULT 0, imported_amount REAL, remarks TEXT DEFAULT '', status TEXT DEFAULT 'Pending', created_by TEXT, created_at TEXT, updated_at TEXT, source_file TEXT DEFAULT '', approved_by TEXT, approved_at TEXT, billed_at TEXT, billed_annexure_ref TEXT, FOREIGN KEY(client_id) REFERENCES clients(id), FOREIGN KEY(project_id) REFERENCES projects(id));
        CREATE TABLE IF NOT EXISTS costs(id INTEGER PRIMARY KEY, entry_ref TEXT UNIQUE, client_id INTEGER, project_id INTEGER, sales_id INTEGER, cost_date TEXT, category TEXT, vendor TEXT DEFAULT '', reference_no TEXT DEFAULT '', amount REAL, remarks TEXT DEFAULT '', status TEXT DEFAULT 'Pending', created_by TEXT, created_at TEXT, source_file TEXT DEFAULT '', approved_by TEXT, approved_at TEXT, FOREIGN KEY(client_id) REFERENCES clients(id), FOREIGN KEY(project_id) REFERENCES projects(id), FOREIGN KEY(sales_id) REFERENCES sales(id));
        CREATE TABLE IF NOT EXISTS budgets(id INTEGER PRIMARY KEY, client_id INTEGER, project_id INTEGER, category TEXT, month TEXT, limit_amount REAL, status TEXT DEFAULT 'Approved', created_by TEXT, created_at TEXT, UNIQUE(client_id,project_id,category,month));
        CREATE TABLE IF NOT EXISTS annexures(id INTEGER PRIMARY KEY, annexure_ref TEXT, client_id INTEGER, project_id INTEGER, month TEXT, category TEXT DEFAULT 'All categories', version INTEGER, status TEXT, excel_path TEXT, pdf_path TEXT, total REAL, generated_by TEXT, generated_at TEXT, UNIQUE(client_id,project_id,month,version));
        CREATE TABLE IF NOT EXISTS audit_log(id INTEGER PRIMARY KEY, username TEXT, action TEXT, table_name TEXT, record_id INTEGER, details TEXT, created_at TEXT);
        ''')
        cost_columns = {row['name'] for row in c.execute('PRAGMA table_info(costs)').fetchall()}
        if 'sales_id' not in cost_columns:
            c.execute('ALTER TABLE costs ADD COLUMN sales_id INTEGER')
        migrations = {
            'sales': {'approved_by':'TEXT','approved_at':'TEXT','billed_at':'TEXT','billed_annexure_ref':'TEXT'},
            'costs': {'approved_by':'TEXT','approved_at':'TEXT'},
            'commercials': {'approved_by':'TEXT','approved_at':'TEXT'},
            'annexures': {'category':"TEXT DEFAULT 'All categories'"},
        }
        for table, cols in migrations.items():
            existing = {row['name'] for row in c.execute(f'PRAGMA table_info({table})').fetchall()}
            for col, typ in cols.items():
                if col not in existing:
                    c.execute(f'ALTER TABLE {table} ADD COLUMN {col} {typ}')
        if not c.execute('SELECT id FROM users LIMIT 1').fetchone():
            c.execute('INSERT INTO users(username,password_hash,role,created_at) VALUES(?,?,?,?)', ('admin', hashlib.sha256('admin123'.encode()).hexdigest(), 'Admin', datetime.now().isoformat(timespec='seconds')))
        for cat in ['FTL','PTL','Multimodal Transport','Manpower','Last Mile Delivery','Freight','Rent','Packaging','Ad hoc','Warehouse Utility','Consumable','Other']:
            c.execute('INSERT OR IGNORE INTO categories(name,kind,active) VALUES(?,?,1)', (cat,'Both'))

def audit(action, table, rid=None, details=''):
    with db() as c:
        c.execute('INSERT INTO audit_log(username,action,table_name,record_id,details,created_at) VALUES(?,?,?,?,?,?)', (st.session_state.get('username','system'),action,table,rid,str(details),datetime.now().isoformat(timespec='seconds')))

def qdf(sql, params=()):
    with db() as c: return pd.read_sql_query(sql,c,params=params)

def options(sql, params=()):
    df=qdf(sql,params)
    return {f"{r['label']}": int(r['id']) for _,r in df.iterrows()} if not df.empty else {}

def money(x):
    try: return f"₹{float(x):,.0f}"
    except: return '₹0'

def to_num(v, default=0):
    if pd.isna(v) or str(v).strip()=='': return default
    try: return float(str(v).replace(',','').replace('₹','').strip())
    except: return default

def clean_str(v):
    if pd.isna(v): return ''
    return str(v).strip()

def make_ref(prefix): return f'{prefix}-{uuid.uuid4().hex[:16].upper()}'

def parse_import_date(value):
    if pd.isna(value) or str(value).strip() == '':
        return pd.NaT
    # Preserve true Excel dates; interpret ambiguous text dates in day-first format.
    if isinstance(value, (datetime, date, pd.Timestamp)):
        return pd.to_datetime(value, errors='coerce')
    return pd.to_datetime(str(value).strip(), errors='coerce', dayfirst=True)

def get_rate(client_id, project_id, category, service_date, source='', destination='', service='', unit='shipment'):
    df=qdf('''SELECT rate FROM commercials WHERE client_id=? AND project_id=? AND category=?
              AND lower(trim(COALESCE(source,'')))=lower(trim(?))
              AND lower(trim(COALESCE(destination,'')))=lower(trim(?))
              AND lower(trim(COALESCE(service,'')))=lower(trim(?))
              AND lower(trim(COALESCE(unit,'shipment')))=lower(trim(?))
              AND status='Approved' AND effective_from<=?
              AND (effective_to='' OR effective_to IS NULL OR effective_to>=?)
              ORDER BY effective_from DESC,id DESC LIMIT 1''',
              (client_id,project_id,category,source,destination,service,unit,str(service_date),str(service_date)))
    return float(df.iloc[0]['rate']) if not df.empty else None


def get_project_label(project_id):
    df=qdf('SELECT p.name FROM projects p WHERE p.id=?',(project_id,))
    return df.iloc[0]['name'] if not df.empty else ''

init_db()

# ---------- Login ----------
if 'logged_in' not in st.session_state: st.session_state.logged_in=False
if not st.session_state.logged_in:
    st.title('Celcius Sales & Cost Monitoring System')
    st.caption('Local-first sales, commercial, cost and annexure management')
    with st.form('login'):
        u=st.text_input('Username', value='admin')
        p=st.text_input('Password', type='password')
        go=st.form_submit_button('Sign in', type='primary')
    if go:
        row=qdf('SELECT * FROM users WHERE username=? AND active=1',(u,))
        if not row.empty and row.iloc[0]['password_hash']==hashlib.sha256(p.encode()).hexdigest():
            st.session_state.logged_in=True; st.session_state.username=u; st.session_state.role=row.iloc[0]['role']; st.rerun()
        else: st.error('Invalid username or password.')
    st.info('First login: username `admin`, password `admin123`. Change the password from User Management after signing in.')
    st.stop()

role=st.session_state.get('role','Sales'); username=st.session_state.get('username','admin')
CAN_APPROVE = role in ['Admin','Finance']
CAN_EDIT_BUDGETS = role in ['Admin','Finance']
def approval_fields(status):
    return (username, datetime.now().isoformat(timespec='seconds')) if status == 'Approved' else (None, None)
st.title('📊 Celcius Sales & Cost Monitoring System')
st.caption(f'Signed in as **{username}** · Role: {role}')
if st.sidebar.button('Log out'):
    st.session_state.logged_in=False; st.rerun()

# Load selectors
client_opts=options("SELECT id, name AS label FROM clients ORDER BY name")
project_opts=options("SELECT p.id, c.name || ' / ' || p.name AS label FROM projects p JOIN clients c ON c.id=p.client_id WHERE p.active=1 ORDER BY c.name,p.name")
category_list=qdf('SELECT name FROM categories WHERE active=1 ORDER BY name')['name'].tolist()
if not category_list: category_list=['FTL','PTL','Multimodal Transport','Manpower','Last Mile Delivery','Freight','Rent','Packaging','Ad hoc','Other']

pages=['Dashboard','Client & Project Master','Commercial Rates','Daily Sales Entry','Sales Register','Cost Entry Register','P&L Reports','Monthly Cost Limits','Generate Monthly Annexure','Excel Import Center','User Management & Audit']
if role=='Management': pages=['Dashboard','Sales Register','P&L Reports','Monthly Cost Limits','Generate Monthly Annexure','User Management & Audit']
page=st.sidebar.radio('Navigation',pages)

# ---------- Dashboard ----------
if page=='Dashboard':
    st.subheader('Management Dashboard')
    today=date.today(); month_default=today.strftime('%Y-%m')
    month=st.selectbox('Reporting month', sorted(set([month_default]+qdf('SELECT DISTINCT substr(service_date,1,7) AS m FROM sales')['m'].dropna().tolist()+qdf('SELECT DISTINCT substr(cost_date,1,7) AS m FROM costs')['m'].dropna().tolist()), reverse=True))
    s=qdf("SELECT COALESCE(SUM(amount),0) v FROM sales WHERE substr(service_date,1,7)=? AND status='Approved'",(month,)).iloc[0]['v']
    cst=qdf("SELECT COALESCE(SUM(amount),0) v FROM costs WHERE substr(cost_date,1,7)=? AND status='Approved'",(month,)).iloc[0]['v']
    pending_s=int(qdf("SELECT COUNT(*) n FROM sales WHERE substr(service_date,1,7)=? AND status='Pending'",(month,)).iloc[0]['n'])
    pending_c=int(qdf("SELECT COUNT(*) n FROM costs WHERE substr(cost_date,1,7)=? AND status='Pending'",(month,)).iloc[0]['n'])
    k=st.columns(5)
    k[0].metric('Approved Sales',money(s)); k[1].metric('Approved Costs',money(cst)); k[2].metric('Contribution',money(s-cst)); k[3].metric('Contribution %',f'{(s-cst)/s*100:.1f}%' if s else '—'); k[4].metric('Pending Reviews',pending_s+pending_c)
    left,right=st.columns(2)
    with left:
        st.markdown('#### Project-wise sales and contribution')
        df=qdf('''SELECT c.name Client,p.name Project,
                         COALESCE((SELECT SUM(s.amount) FROM sales s WHERE s.project_id=p.id AND substr(s.service_date,1,7)=? AND s.status='Approved'),0) Sales,
                         COALESCE((SELECT SUM(x.amount) FROM costs x WHERE x.project_id=p.id AND substr(x.cost_date,1,7)=? AND x.status='Approved'),0) Costs
                  FROM projects p JOIN clients c ON c.id=p.client_id ORDER BY Sales DESC''',(month,month))
        if not df.empty:
            df['Contribution']=df['Sales']-df['Costs']
            for col in ['Sales','Costs','Contribution']: df[col]=df[col].map(money)
            st.dataframe(df,use_container_width=True,hide_index=True)
        else: st.info('No project data yet. Add a client and project to begin.')
    with right:
        st.markdown('#### Category-wise approved costs')
        costs=qdf("SELECT category AS Category, SUM(amount) AS Cost FROM costs WHERE substr(cost_date,1,7)=? AND status='Approved' GROUP BY category ORDER BY Cost DESC",(month,))
        if not costs.empty: st.bar_chart(costs.set_index('Category'))
        else: st.info('No approved costs for this month.')
    st.markdown('#### Monthly budget status')
    b=qdf('''SELECT b.category Category,b.limit_amount Limit_Amount,COALESCE(x.actual,0) Actual, b.limit_amount-COALESCE(x.actual,0) Remaining FROM budgets b LEFT JOIN (SELECT project_id,category,substr(cost_date,1,7) month,SUM(amount) actual FROM costs WHERE status='Approved' GROUP BY project_id,category,substr(cost_date,1,7)) x ON x.project_id=b.project_id AND x.category=b.category AND x.month=b.month WHERE b.month=? ORDER BY b.category''',(month,))
    if not b.empty:
        b['Over Budget']=b['Actual']>b['Limit_Amount']; st.dataframe(b,use_container_width=True,hide_index=True)
    else: st.caption('No monthly cost limits set for this month yet.')

# ---------- Masters ----------
elif page=='Client & Project Master':
    st.subheader('Client & Project Master')
    a,b=st.columns(2)
    with a:
        st.markdown('#### Add client')
        with st.form('clientform'):
            cname=st.text_input('Client name'); addc=st.form_submit_button('Save client')
        if addc:
            if not cname.strip(): st.error('Client name is required.')
            else:
                try:
                    with db() as c: c.execute('INSERT INTO clients(name,created_at) VALUES(?,?)',(cname.strip(),datetime.now().isoformat(timespec='seconds')))
                    audit('CREATE','clients',details=cname); st.success('Client saved.'); st.rerun()
                except sqlite3.IntegrityError: st.warning('This client already exists.')
        st.dataframe(qdf('SELECT id AS ID,name AS Client,created_at AS Created FROM clients ORDER BY name'),use_container_width=True,hide_index=True)
    with b:
        st.markdown('#### Add project')
        if not client_opts: st.info('Add a client first.')
        else:
            with st.form('projectform'):
                clabel=st.selectbox('Client',list(client_opts)); pname=st.text_input('Project name'); product=st.text_input('Product/service description'); addp=st.form_submit_button('Save project')
            if addp:
                if not pname.strip(): st.error('Project name is required.')
                else:
                    try:
                        with db() as c: c.execute('INSERT INTO projects(client_id,name,product,created_at) VALUES(?,?,?,?)',(client_opts[clabel],pname.strip(),product.strip(),datetime.now().isoformat(timespec='seconds')))
                        audit('CREATE','projects',details=f'{clabel}/{pname}'); st.success('Project saved.'); st.rerun()
                    except sqlite3.IntegrityError: st.warning('This project already exists for that client.')
        st.dataframe(qdf('SELECT p.id ID,c.name Client,p.name Project,p.product Product,p.active Active FROM projects p JOIN clients c ON c.id=p.client_id ORDER BY c.name,p.name'),use_container_width=True,hide_index=True)
    st.markdown('#### Manage categories')
    with st.form('categoryform'):
        newcat=st.text_input('New category'); kind=st.selectbox('Applies to',['Both','Sales','Cost']); addcat=st.form_submit_button('Add category')
    if addcat and newcat.strip():
        with db() as c: c.execute('INSERT OR IGNORE INTO categories(name,kind,active) VALUES(?,?,1)',(newcat.strip(),kind))
        audit('CREATE','categories',details=newcat); st.success('Category added (or already existed).'); st.rerun()
    st.dataframe(qdf('SELECT name Category,kind Applies_To,active Active FROM categories ORDER BY name'),use_container_width=True,hide_index=True)

# ---------- Commercial Rates ----------
elif page=='Commercial Rates':
    st.subheader('Commercial Rate Registration')
    tab1,tab2=st.tabs(['Manual registration','Rate register'])
    with tab1:
        if not client_opts or not project_opts: st.warning('Create client and project first.')
        else:
            with st.form('commercialform'):
                x1,x2,x3=st.columns(3)
                cl=x1.selectbox('Client',list(client_opts),key='rate_client'); rate_project_rows=qdf('SELECT id,name FROM projects WHERE client_id=? AND active=1 ORDER BY name',(client_opts[cl],)); rate_project_opts={str(r['name']):int(r['id']) for _,r in rate_project_rows.iterrows()}; pr=x2.selectbox('Project',list(rate_project_opts) if rate_project_opts else ['No active projects'],key='rate_project'); cat=x3.selectbox('Category',category_list)
                x1,x2,x3=st.columns(3)
                service=x1.text_input('Service / rate description'); source=x2.text_input('Source'); destination=x3.text_input('Destination')
                x1,x2,x3,x4=st.columns(4)
                unit=x1.selectbox('Rate unit',['shipment','kg','box','crate','trip','day','hour','km','piece','other']); rate=x2.number_input('Rate (₹)',min_value=0.0,step=100.0); ef=x3.date_input('Effective from',value=date.today()); et=x4.date_input('Effective to (optional)',value=None)
                status=st.selectbox('Approval status',['Pending','Approved','Rejected'] if CAN_APPROVE else ['Pending','Rejected']); notes=st.text_input('Notes'); save=st.form_submit_button('Save commercial rate')
            if save:
                if not rate_project_opts: st.error('Select a client with an active project.')
                elif status=='Approved' and not CAN_APPROVE: st.error('Only Admin or Finance can approve commercial rates.')
                elif rate<=0: st.error('Rate must be greater than zero.')
                else:
                    apby,apat=approval_fields(status)
                    with db() as c:
                        c.execute('INSERT INTO commercials(client_id,project_id,category,service,source,destination,unit,rate,effective_from,effective_to,status,notes,created_by,created_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)',(client_opts[cl],rate_project_opts[pr],cat,service,source,destination,unit,rate,ef.isoformat(),et.isoformat() if et else '',status,notes,username,datetime.now().isoformat(timespec='seconds')))
                        if status=='Approved': c.execute('UPDATE commercials SET approved_by=?,approved_at=? WHERE id=last_insert_rowid()', (apby,apat))
                    audit('CREATE','commercials',details=f'{cl}/{pr}/{cat}/{rate}'); st.success('Rate saved.'); st.rerun()
    with tab2:
        df=qdf('''SELECT r.id ID,c.name Client,p.name Project,r.category Category,r.service Service,r.source Source,r.destination Destination,r.unit Unit,r.rate Rate,r.effective_from Effective_From,r.effective_to Effective_To,r.status Status,r.approved_by Approved_By,r.approved_at Approved_At,r.notes Notes FROM commercials r JOIN clients c ON c.id=r.client_id JOIN projects p ON p.id=r.project_id ORDER BY r.id DESC''')
        st.dataframe(df,use_container_width=True,hide_index=True)
        if role in ['Admin','Finance'] and not df.empty:
            rid=st.number_input('Rate record ID to review',min_value=1,step=1,key='rateid'); newstatus=st.selectbox('Set status',['Approved','Pending','Rejected'],key='ratestatus')
            if st.button('Update rate approval'):
                if newstatus=='Approved' and not CAN_APPROVE: st.error('Only Admin or Finance can approve commercial rates.')
                else:
                    apby,apat=approval_fields(newstatus)
                    with db() as c: c.execute('UPDATE commercials SET status=?,approved_by=?,approved_at=? WHERE id=?',(newstatus,apby,apat,int(rid)))
                audit('UPDATE','commercials',int(rid),newstatus); st.success('Rate status updated.'); st.rerun()

# ---------- Sales Entry ----------
elif page=='Daily Sales Entry':
    st.subheader('Daily Sales Entry')
    if not client_opts or not project_opts: st.warning('Create client and project first.')
    else:
        with st.form('salesform'):
            a,b,c=st.columns(3)
            cl=a.selectbox('Client',list(client_opts),key='sales_client'); sale_project_rows=qdf('SELECT id,name FROM projects WHERE client_id=? AND active=1 ORDER BY name',(client_opts[cl],)); sale_project_opts={str(r['name']):int(r['id']) for _,r in sale_project_rows.iterrows()}; pr=b.selectbox('Project',list(sale_project_opts) if sale_project_opts else ['No active projects'],key='sales_project'); sd=c.date_input('Service date',value=date.today())
            a,b,c=st.columns(3)
            cat=a.selectbox('Category',category_list); service=b.text_input('Product / service'); shipment=c.text_input('Shipment / docket reference')
            a,b,c=st.columns(3)
            src=a.text_input('Source'); dst=b.text_input('Destination'); unit=c.selectbox('Unit',['shipment','kg','box','crate','trip','day','hour','km','piece','other'])
            a,b,c=st.columns(3)
            qty=a.number_input('Quantity',min_value=0.0,value=1.0,step=1.0); use_rate=b.checkbox('Use approved commercial rate',value=True); manual_rate=b.number_input('Manual rate (₹, if not using commercial)',min_value=0.0,step=100.0,disabled=use_rate); use_override=c.checkbox('Use amount override instead of quantity × rate',value=False); imported_amount=c.number_input('Amount override (₹)',min_value=0.0,step=100.0,value=0.0,disabled=not use_override)
            remarks=st.text_area('Remarks'); save=st.form_submit_button('Save sales entry',type='primary')
        if save:
            client_id=client_opts[cl]
            project_id=sale_project_opts.get(pr)
            if project_id is None:
                st.error('Select a client with an active project.')
            else:
                found=get_rate(client_id,project_id,cat,sd.isoformat(),src,dst,service,unit) if use_rate else None
                if use_rate and found is None:
                    st.error('No approved commercial rate matches client, project, category, source, destination, service, unit and service date. Approve an exact-match rate or uncheck the option and enter a manual rate.')
                elif not use_rate and manual_rate <= 0:
                    st.error('Manual rate must be greater than zero.')
                elif use_override and imported_amount <= 0:
                    st.error('Amount override must be greater than zero.')
                elif qty <= 0:
                    st.error('Quantity must be greater than zero.')
                else:
                    rt=found if use_rate else manual_rate
                    calc=round(qty*rt,2)
                    amt=imported_amount if use_override else calc
                    now=datetime.now().isoformat(timespec='seconds')
                    with db() as c:
                        c.execute('INSERT INTO sales(entry_ref,client_id,project_id,service_date,category,service,source,destination,shipment_ref,quantity,unit,rate,amount,imported_amount,remarks,status,created_by,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)',(make_ref('SAL'),client_id,project_id,sd.isoformat(),cat,service,src,dst,shipment,qty,unit,rt,amt,imported_amount if use_override else None,remarks,'Pending',username,now,now))
                    audit('CREATE','sales',details=f'{cl}/{pr}/{sd}/{amt}')
                    st.success(f'Sales entry saved for {money(amt)}. Entry is pending approval.')
                    st.rerun()
    st.caption('Sales are created as Pending. An authorised reviewer must approve them before they appear in approved dashboard totals or final approved annexures.')

# ---------- Sales Register ----------
elif page=='Sales Register':
    st.subheader('Sales Register & Approval')
    filters=qdf('SELECT DISTINCT substr(service_date,1,7) m FROM sales ORDER BY m DESC')['m'].dropna().tolist()
    month=st.selectbox('Month filter',['All months']+filters)
    sql='''SELECT s.id ID,s.entry_ref Entry_Ref,s.billed_at Billed_At,s.billed_annexure_ref Billed_Annexure_Ref,c.name Client,p.name Project,s.service_date Service_Date,s.category Category,s.service Service,s.source Source,s.destination Destination,s.shipment_ref Shipment_Ref,s.quantity Quantity,s.unit Unit,s.rate Rate,s.amount Amount,s.imported_amount Imported_Amount,s.status Status,s.approved_by Approved_By,s.approved_at Approved_At,s.remarks Remarks,s.created_by Created_By FROM sales s JOIN clients c ON c.id=s.client_id JOIN projects p ON p.id=s.project_id'''
    if month!='All months': sql+=' WHERE substr(s.service_date,1,7)=?'
    sql+=' ORDER BY s.service_date DESC,s.id DESC'
    df=qdf(sql,(month,) if month!='All months' else ())
    st.dataframe(df,use_container_width=True,hide_index=True)
    if not df.empty and CAN_APPROVE:
        with st.expander('Review / correct a sales record'):
            rid=st.number_input('Sales record ID',min_value=1,step=1,key='salesrid'); act=st.selectbox('Action',['Approve','Reject','Return to Pending','Cancel']); reason=st.text_input('Review reason / correction note')
            if st.button('Apply sales review'):
                rec=qdf('SELECT billed_at,status FROM sales WHERE id=?',(int(rid),))
                if rec.empty: st.error('Sales record ID not found.')
                elif rec.iloc[0]['billed_at']: st.error('This sales record is billed and locked; create a new version/adjustment instead.')
                else:
                    status={'Approve':'Approved','Reject':'Rejected','Return to Pending':'Pending','Cancel':'Cancelled'}[act]
                    if status=='Approved' and not CAN_APPROVE: st.error('Only Admin or Finance can approve records.')
                    else:
                        apby,apat=approval_fields(status)
                        with db() as c: c.execute("UPDATE sales SET status=?,approved_by=?,approved_at=?,remarks=CASE WHEN ?<>'' THEN COALESCE(remarks,'') || ' | Review: ' || ? ELSE remarks END,updated_at=? WHERE id=? AND (billed_at IS NULL OR billed_at='')",(status,apby,apat,reason,reason,datetime.now().isoformat(timespec='seconds'),int(rid)))
                        audit('REVIEW','sales',int(rid),f'{status}: {reason}'); st.success('Sales record updated.'); st.rerun()
    if not df.empty:
        st.download_button('Download register as CSV',df.to_csv(index=False).encode('utf-8-sig'),'sales_register.csv','text/csv')

# ---------- Cost Entry ----------
elif page=='Cost Entry Register':
    st.subheader('Cost Entry Register')
    tabs=st.tabs(['Manual cost entry','Cost register'])
    with tabs[0]:
        if not client_opts or not project_opts:
            st.warning('Create client and project first.')
        else:
            cl=st.selectbox('Client',list(client_opts),key='costcl')
            client_id=client_opts[cl]
            project_rows=qdf('SELECT id,name FROM projects WHERE client_id=? AND active=1 ORDER BY name',(client_id,))
            project_choices={str(r['name']):int(r['id']) for _,r in project_rows.iterrows()}
            if not project_choices:
                st.warning('This client has no active projects. Add a project first.')
            else:
                pr=st.selectbox('Project',list(project_choices),key='costpr')
                project_id=project_choices[pr]
                sale_rows=qdf("SELECT id,entry_ref,service_date,category,amount,status FROM sales WHERE client_id=? AND project_id=? AND status='Approved' ORDER BY service_date DESC,id DESC",(client_id,project_id))
                sale_options={'Unlinked project cost (not tied to one sale)':None}
                for _,sale in sale_rows.iterrows():
                    sale_options[f"{sale['entry_ref']} | {sale['service_date']} | {sale['category']} | {money(sale['amount'])}"]=int(sale['id'])
                selected_sale_label=st.selectbox('Link cost to Sale Entry (optional)',list(sale_options),key='cost_sales_link',help='Link to an approved sale to include this cost in that sale category P&L. Leave unlinked for general project costs.')
                selected_sale_id=sale_options[selected_sale_label]
                with st.form('costform'):
                    a,b,c=st.columns(3)
                    cd=a.date_input('Cost date',value=date.today())
                    cat=b.selectbox('Cost category',category_list,key='costcat')
                    vendor=c.text_input('Vendor / transporter')
                    a,b=st.columns(2)
                    ref=a.text_input('Invoice / reference number')
                    amount=b.number_input('Cost amount (₹)',min_value=0.0,step=100.0)
                    remarks=st.text_area('Cost remarks')
                    save=st.form_submit_button('Save cost entry',type='primary')
                if save:
                    if amount<=0:
                        st.error('Cost amount must be greater than zero.')
                    else:
                        with db() as c:
                            c.execute("INSERT INTO costs(entry_ref,client_id,project_id,sales_id,cost_date,category,vendor,reference_no,amount,remarks,status,created_by,created_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)",
                                      (make_ref('CST'),client_id,project_id,selected_sale_id,cd.isoformat(),cat,vendor,ref,amount,remarks,'Pending',username,datetime.now().isoformat(timespec='seconds')))
                        audit('CREATE','costs',details=f'{cl}/{pr}/{cat}/{amount}; linked_sale_id={selected_sale_id}')
                        st.success('Cost saved as Pending.' + (f' Linked to {selected_sale_label}.' if selected_sale_id else ''))
                        st.rerun()
                if not sale_rows.empty:
                    st.caption('Approved sales available for linking to this project')
                    sale_view=sale_rows[['entry_ref','service_date','category','amount']].copy()
                    sale_view['amount']=sale_view['amount'].map(money)
                    st.dataframe(sale_view,use_container_width=True,hide_index=True)
    with tabs[1]:
        df=qdf("SELECT x.id ID,x.entry_ref Entry_Ref,c.name Client,p.name Project,x.sales_id Linked_Sale_ID,s.entry_ref Linked_Sale_Ref,s.category Linked_Sale_Category,x.cost_date Cost_Date,x.category Cost_Category,x.vendor Vendor,x.reference_no Reference,x.amount Amount,x.status Status,x.approved_by Approved_By,x.approved_at Approved_At,x.remarks Remarks,x.created_by Created_By FROM costs x JOIN clients c ON c.id=x.client_id JOIN projects p ON p.id=x.project_id LEFT JOIN sales s ON s.id=x.sales_id ORDER BY x.cost_date DESC,x.id DESC")
        st.dataframe(df,use_container_width=True,hide_index=True)
        if not df.empty:
            st.download_button('Download cost register CSV',df.to_csv(index=False).encode('utf-8-sig'),'cost_register.csv','text/csv')
            if role in ['Admin','Finance','Operations/Cost']:
                with st.expander('Review cost record'):
                    rid=st.number_input('Cost record ID',min_value=1,step=1,key='costrid'); act=st.selectbox('Action',['Approve','Reject','Return to Pending'],key='costact'); reason=st.text_input('Review reason',key='costreason')
                    if st.button('Apply cost review'):
                        status={'Approve':'Approved','Reject':'Rejected','Return to Pending':'Pending'}[act]
                        if status=='Approved' and not CAN_APPROVE: st.error('Only Admin or Finance can approve costs.')
                        else:
                            apby,apat=approval_fields(status)
                            with db() as c: c.execute("UPDATE costs SET status=?,approved_by=?,approved_at=?,remarks=CASE WHEN ?<>'' THEN COALESCE(remarks,'') || ' | Review: ' || ? ELSE remarks END WHERE id=?",(status,apby,apat,reason,reason,int(rid)))
                            audit('REVIEW','costs',int(rid),f'{status}: {reason}'); st.success('Cost record updated.'); st.rerun()

# ---------- P&L Reports ----------
elif page=='P&L Reports':
    st.subheader('Profit & Loss Reports')
    months=qdf("SELECT DISTINCT month FROM (SELECT substr(service_date,1,7) month FROM sales UNION SELECT substr(cost_date,1,7) month FROM costs) WHERE month IS NOT NULL ORDER BY month DESC")['month'].dropna().tolist()
    selected_month=st.selectbox('Reporting period',['All time']+months,key='pnl_month')
    period_sales=" AND substr(s.service_date,1,7)=?" if selected_month!='All time' else ''
    period_costs=" AND substr(co.cost_date,1,7)=?" if selected_month!='All time' else ''
    params_sales=(selected_month,) if selected_month!='All time' else ()
    params_costs=(selected_month,) if selected_month!='All time' else ()
    sales_df=qdf("SELECT s.id,s.client_id,s.project_id,s.service_date,s.category,s.amount,c.name AS Client,p.name AS Project FROM sales s JOIN clients c ON c.id=s.client_id JOIN projects p ON p.id=s.project_id WHERE s.status='Approved'"+period_sales,params_sales)
    costs_df=qdf("SELECT co.id,co.client_id,co.project_id,co.sales_id,co.cost_date,co.category AS Cost_Category,co.amount,c.name AS Client,p.name AS Project,s.category AS Linked_Sale_Category FROM costs co JOIN clients c ON c.id=co.client_id JOIN projects p ON p.id=co.project_id LEFT JOIN sales s ON s.id=co.sales_id WHERE co.status='Approved'"+period_costs,params_costs)
    total_sales=float(sales_df['amount'].sum()) if not sales_df.empty else 0.0
    total_costs=float(costs_df['amount'].sum()) if not costs_df.empty else 0.0
    total_profit=total_sales-total_costs
    k=st.columns(4)
    k[0].metric('Approved Sales',money(total_sales)); k[1].metric('Approved Costs',money(total_costs)); k[2].metric('Gross Profit',money(total_profit)); k[3].metric('Margin',f'{total_profit/total_sales*100:.1f}%' if total_sales else '—')
    t1,t2,t3,t4=st.tabs(['Project-wise P&L','Client-wise P&L','Category-wise P&L','Cost Category Breakdown'])
    def show_pnl(frame):
        if frame.empty:
            st.info('No approved entries found for this reporting period.')
            return
        frame['Gross Profit']=frame['Sales']-frame['Costs']
        frame['Margin %']=frame.apply(lambda r:r['Gross Profit']/r['Sales']*100 if r['Sales'] else 0,axis=1)
        for col in ['Sales','Costs','Gross Profit']:
            frame[col]=frame[col].map(money)
        frame['Margin %']=frame['Margin %'].map(lambda v:f'{v:.1f}%')
        st.dataframe(frame,use_container_width=True,hide_index=True)
        st.download_button('Download P&L CSV',frame.to_csv(index=False).encode('utf-8-sig'),'pnl_report.csv','text/csv',key='download_'+str(abs(hash(tuple(frame.columns)))))
    with t1:
        sgroup=sales_df.groupby(['Client','Project'],dropna=False)['amount'].sum().rename('Sales') if not sales_df.empty else pd.Series(dtype=float,name='Sales')
        cgroup=costs_df.groupby(['Client','Project'],dropna=False)['amount'].sum().rename('Costs') if not costs_df.empty else pd.Series(dtype=float,name='Costs')
        pnl=pd.concat([sgroup,cgroup],axis=1).fillna(0).reset_index().rename(columns={'Client':'Client','Project':'Project'})
        show_pnl(pnl)
    with t2:
        sgroup=sales_df.groupby('Client',dropna=False)['amount'].sum().rename('Sales') if not sales_df.empty else pd.Series(dtype=float,name='Sales')
        cgroup=costs_df.groupby('Client',dropna=False)['amount'].sum().rename('Costs') if not costs_df.empty else pd.Series(dtype=float,name='Costs')
        pnl=pd.concat([sgroup,cgroup],axis=1).fillna(0).reset_index()
        show_pnl(pnl)
    with t3:
        sales_cat=sales_df.groupby('category',dropna=False)['amount'].sum().rename('Sales') if not sales_df.empty else pd.Series(dtype=float,name='Sales')
        linked=costs_df[costs_df['sales_id'].notna()] if not costs_df.empty else costs_df
        linked_cost=linked.groupby('Linked_Sale_Category',dropna=False)['amount'].sum() if not linked.empty else pd.Series(dtype=float)
        if not linked_cost.empty: linked_cost.index=linked_cost.index.map(lambda x:str(x) if pd.notna(x) else 'Uncategorised')
        unlinked=costs_df[costs_df['sales_id'].isna()] if not costs_df.empty else costs_df
        unlinked_cost=unlinked.groupby('Cost_Category',dropna=False)['amount'].sum() if not unlinked.empty else pd.Series(dtype=float)
        if not unlinked_cost.empty: unlinked_cost.index=unlinked_cost.index.map(lambda x:'Unallocated: '+(str(x) if pd.notna(x) else 'Uncategorised'))
        all_cost=pd.concat([linked_cost,unlinked_cost]).groupby(level=0).sum().rename('Costs') if (not linked_cost.empty or not unlinked_cost.empty) else pd.Series(dtype=float,name='Costs')
        cat_pnl=pd.concat([sales_cat,all_cost],axis=1).fillna(0).reset_index()
        if 'category' in cat_pnl.columns: cat_pnl=cat_pnl.rename(columns={'category':'Category'})
        elif 'index' in cat_pnl.columns: cat_pnl=cat_pnl.rename(columns={'index':'Category'})
        show_pnl(cat_pnl)
        st.caption('Linked costs follow the selected sale category. Unlinked costs stay separately labelled as Unallocated.')
    with t4:
        if costs_df.empty:
            st.info('No approved costs found for this reporting period.')
        else:
            bycat=costs_df.groupby('Cost_Category',dropna=False)['amount'].sum().reset_index(name='Approved Cost')
            bycat['Approved Cost']=bycat['Approved Cost'].map(money)
            st.dataframe(bycat.rename(columns={'Cost_Category':'Cost Category'}),use_container_width=True,hide_index=True)

# ---------- Monthly Cost Limits ----------
elif page=='Monthly Cost Limits':
    st.subheader('Monthly Cost Limits by Category')
    if not CAN_EDIT_BUDGETS: st.info('Only Admin or Finance can create or edit monthly cost limits.')
    elif not client_opts or not project_opts: st.warning('Create client and project first.')
    else:
        with st.form('budgetform'):
            a,b,c=st.columns(3); cl=a.selectbox('Client',list(client_opts),key='budgetcl'); pr=b.selectbox('Project',list(project_opts),key='budgetpr'); cat=c.selectbox('Category',category_list,key='budgetcat')
            a,b=st.columns(2); month=a.text_input('Month (YYYY-MM)',value=date.today().strftime('%Y-%m')); limit=b.number_input('Approved monthly limit (₹)',min_value=0.0,step=1000.0)
            save=st.form_submit_button('Save / update monthly limit')
        if save:
            if not CAN_EDIT_BUDGETS: st.error('Only Admin or Finance can edit budgets.')
            else:
              try:
                datetime.strptime(month,'%Y-%m')
                with db() as c:
                    c.execute('''INSERT INTO budgets(client_id,project_id,category,month,limit_amount,status,created_by,created_at) VALUES(?,?,?,?,?,?,?,?) ON CONFLICT(client_id,project_id,category,month) DO UPDATE SET limit_amount=excluded.limit_amount,status=excluded.status,created_by=excluded.created_by,created_at=excluded.created_at''',(client_opts[cl],project_opts[pr],cat,month,limit,'Approved',username,datetime.now().isoformat(timespec='seconds')))
                audit('UPSERT','budgets',details=f'{cl}/{pr}/{cat}/{month}/{limit}'); st.success('Monthly limit saved.'); st.rerun()
              except ValueError: st.error('Month must be YYYY-MM, for example 2026-10.')
    df=qdf('''SELECT b.id ID,c.name Client,p.name Project,b.month Month,b.category Category,b.limit_amount Approved_Limit,COALESCE(SUM(CASE WHEN x.status='Approved' THEN x.amount ELSE 0 END),0) Actual_Cost,b.limit_amount-COALESCE(SUM(CASE WHEN x.status='Approved' THEN x.amount ELSE 0 END),0) Remaining FROM budgets b JOIN clients c ON c.id=b.client_id JOIN projects p ON p.id=b.project_id LEFT JOIN costs x ON x.project_id=b.project_id AND x.category=b.category AND substr(x.cost_date,1,7)=b.month GROUP BY b.id ORDER BY b.month DESC,c.name,p.name,b.category''')
    if not df.empty:
        df['Utilisation %']=df.apply(lambda r: round(r['Actual_Cost']/r['Approved_Limit']*100,1) if r['Approved_Limit'] else 0,axis=1); df['Status']=df.apply(lambda r:'OVER BUDGET' if r['Actual_Cost']>r['Approved_Limit'] else ('NEAR LIMIT' if r['Utilisation %']>=80 else 'Within limit'),axis=1)
        st.dataframe(df,use_container_width=True,hide_index=True)
        st.download_button('Download budget report CSV',df.to_csv(index=False).encode('utf-8-sig'),'monthly_budget_report.csv','text/csv')
    else: st.info('No budgets entered yet.')

# ---------- Annexure ----------
elif page=='Generate Monthly Annexure':
    st.subheader('Generate Monthly Project Annexure')
    st.write('Select the client, project and billing month. Only Approved sales entries are included in the final annexure. Pending or incomplete entries are flagged before generation.')
    if not client_opts or not project_opts: st.warning('Create client and project first.')
    else:
        with st.form('annexureform'):
            a,b,c=st.columns(3); cl=a.selectbox('Client',list(client_opts),key='anncl'); pr=b.selectbox('Project',list(project_opts),key='annpr'); month=c.text_input('Billing month (YYYY-MM)',value=date.today().strftime('%Y-%m'))
            cat_filter=st.selectbox('Category filter',['All categories']+category_list); prepare=st.form_submit_button('Preview annexure')
        if prepare:
            try:
                datetime.strptime(month,'%Y-%m')
                cid=client_opts[cl]; pid=project_opts[pr]
                sql='''SELECT s.*,c.name client_name,p.name project_name FROM sales s JOIN clients c ON c.id=s.client_id JOIN projects p ON p.id=s.project_id WHERE s.client_id=? AND s.project_id=? AND substr(s.service_date,1,7)=?'''
                params=[cid,pid,month]
                if cat_filter!='All categories': sql+=' AND s.category=?'; params.append(cat_filter)
                sql+=' ORDER BY s.service_date,s.id'
                data=qdf(sql,tuple(params))
                st.session_state['ann_data']=data; st.session_state['ann_meta']={'client':cl,'project':pr,'cid':cid,'pid':pid,'month':month,'category':cat_filter}
            except ValueError: st.error('Month must be YYYY-MM.')
        data=st.session_state.get('ann_data')
        meta=st.session_state.get('ann_meta')
        if data is not None and meta:
            if data.empty: st.warning('No sales entries found for this project and month.')
            else:
                unresolved=data[(data['status']!='Approved') | data['service_date'].isna() | (data['category'].fillna('').astype(str).str.strip()=='') | data['amount'].isna() | (pd.to_numeric(data['amount'],errors='coerce').fillna(0)<=0) | (data['quantity'].isna()) | (pd.to_numeric(data['quantity'],errors='coerce').fillna(0)<=0) | (data['billed_at'].notna() & (data['billed_at'].astype(str).str.strip()!=''))]
                st.markdown('#### Review status')
                a,b,c=st.columns(3); a.metric('Records found',len(data)); b.metric('Approved records',int((data.status=='Approved').sum())); c.metric('Unresolved records',len(unresolved))
                if not unresolved.empty:
                    st.warning('Some records are pending, rejected, cancelled, or incomplete. They will not be included in a final approved annexure.')
                    st.dataframe(unresolved[['id','entry_ref','service_date','category','amount','status','remarks']],use_container_width=True,hide_index=True)
                view=data[data.status=='Approved'].copy()
                if not view.empty:
                    view['Line Amount']=view['amount'].apply(lambda x: round(float(x or 0),2))
                    out=view.rename(columns={'entry_ref':'Entry Reference','service_date':'Service Date','source':'Source','destination':'Destination','category':'Category','service':'Product / Service','shipment_ref':'Shipment / Docket Reference','quantity':'Quantity','unit':'Unit','rate':'Approved Rate','status':'Status'})
                    cols=['Entry Reference','Service Date','Source','Destination','Category','Product / Service','Shipment / Docket Reference','Quantity','Unit','Approved Rate','Line Amount','Status']
                    st.markdown('#### Approved annexure preview')
                    st.dataframe(out[cols],use_container_width=True,hide_index=True)
                    subtotals=out.groupby('Category',dropna=False)['Line Amount'].sum().reset_index(name='Subtotal')
                    total=float(out['Line Amount'].sum())
                    st.markdown('#### Category subtotals')
                    st.dataframe(subtotals,use_container_width=True,hide_index=True)
                    st.metric('Annexure grand total',money(total))
                    if unresolved.empty and CAN_APPROVE:
                        approve_final=st.checkbox('I confirm the displayed records are reviewed and ready for client submission')
                        if st.button('Generate approved Excel + PDF annexure',type='primary',disabled=not approve_final):
                            old=qdf('SELECT MAX(version) v FROM annexures WHERE client_id=? AND project_id=? AND month=?',(meta['cid'],meta['pid'],meta['month']))
                            version=int(old.iloc[0]['v'] or 0)+1
                            safe=f"{meta['client']}_{meta['project']}_{meta['month']}_v{version}".replace('/','-').replace(' ','_')
                            excel_path=os.path.join(EXPORT_DIR,safe+'.xlsx'); pdf_path=os.path.join(EXPORT_DIR,safe+'.pdf')
                            with pd.ExcelWriter(excel_path,engine='openpyxl') as writer:
                                pd.DataFrame([['CELCIUS - MONTHLY PROJECT ANNEXURE'],['Client',meta['client']],['Project',meta['project']],['Billing Month',meta['month']],['Annexure Reference',safe],['Prepared By',username],['Generated At',datetime.now().strftime('%d-%b-%Y %H:%M')]]).to_excel(writer,index=False,header=False,sheet_name='Annexure')
                                out[cols].to_excel(writer,index=False,sheet_name='Annexure',startrow=7)
                                subtotals.to_excel(writer,index=False,sheet_name='Category Subtotals')
                                ws=writer.sheets['Annexure']; ws.freeze_panes='A9'; ws.auto_filter.ref=f'A8:L{8+len(out)}'
                                for column_cells in ws.columns:
                                    letter=column_cells[0].column_letter
                                    maxlen=max(len(str(cell.value or '')) for cell in column_cells)
                                    ws.column_dimensions[letter].width=min(max(maxlen+2,12),32)
                                for row in ws.iter_rows():
                                    for cell in row:
                                        if cell.row==1: cell.font=__import__('openpyxl').styles.Font(bold=True,size=15)
                            pdf_ok=False
                            try:
                                from reportlab.lib import colors
                                from reportlab.lib.pagesizes import landscape, A4
                                from reportlab.lib.styles import getSampleStyleSheet
                                from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
                                doc=SimpleDocTemplate(pdf_path,pagesize=landscape(A4),rightMargin=24,leftMargin=24,topMargin=24,bottomMargin=24)
                                styles=getSampleStyleSheet(); story=[Paragraph('CELCIUS - MONTHLY PROJECT ANNEXURE',styles['Title']),Paragraph(f"Client: {meta['client']} &nbsp;&nbsp; Project: {meta['project']} &nbsp;&nbsp; Billing Month: {meta['month']}",styles['Normal']),Paragraph(f'Annexure Ref: {safe} &nbsp;&nbsp; Prepared By: {username} &nbsp;&nbsp; Status: Approved',styles['Normal']),Spacer(1,12)]
                                pdf_cols=['Service Date','Source','Destination','Category','Shipment Ref','Qty','Unit','Rate','Amount']
                                pdfdata=[pdf_cols]
                                for _,r in out.iterrows(): pdfdata.append([str(r['Service Date']),str(r['Source']),str(r['Destination']),str(r['Category']),str(r['Shipment / Docket Reference']),str(r['Quantity']),str(r['Unit']),f"{float(r['Approved Rate']):,.2f}",f"{float(r['Line Amount']):,.2f}"])
                                pdfdata.append(['','','','','','','','Grand Total',f'{total:,.2f}'])
                                tbl=Table(pdfdata,repeatRows=1)
                                tbl.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#17365D')),('TEXTCOLOR',(0,0),(-1,0),colors.white),('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),('GRID',(0,0),(-1,-1),0.4,colors.grey),('FONTSIZE',(0,0),(-1,-1),7),('ALIGN',(5,1),(-1,-1),'RIGHT'),('BACKGROUND',(0,-1),(-1,-1),colors.HexColor('#D9EAF7'))]))
                                story.extend([tbl,Spacer(1,12),Paragraph('Category Subtotals',styles['Heading2'])])
                                subdata=[['Category','Subtotal']]+[[str(r['Category']),f"{float(r['Subtotal']):,.2f}"] for _,r in subtotals.iterrows()]
                                story.append(Table(subdata,style=TableStyle([('BACKGROUND',(0,0),(-1,0),colors.lightgrey),('GRID',(0,0),(-1,-1),0.4,colors.grey)])))
                                doc.build(story); pdf_ok=True
                            except Exception as e: st.warning(f'Excel created. PDF generation failed: {e}. Install requirements and try again.')
                            with db() as c:
                                c.execute('INSERT INTO annexures(annexure_ref,client_id,project_id,month,category,version,status,excel_path,pdf_path,total,generated_by,generated_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?)',(safe,meta['cid'],meta['pid'],meta['month'],meta.get('category','All categories'),version,'Approved',excel_path,pdf_path,total,username,datetime.now().isoformat(timespec='seconds')))
                                billed_now=datetime.now().isoformat(timespec='seconds')
                                c.execute('UPDATE sales SET billed_at=?,billed_annexure_ref=? WHERE id IN ('+','.join('?' for _ in view['id'].tolist())+') AND (billed_at IS NULL OR billed_at='')',(billed_now,safe,*[int(x) for x in view['id'].tolist()]))
                            audit('GENERATE','annexures',details=safe); st.success(f'Annexure generated: {safe}')
                            with open(excel_path,'rb') as f: st.download_button('Download Excel annexure',f.read(),os.path.basename(excel_path), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                            if pdf_ok:
                                with open(pdf_path,'rb') as f: st.download_button('Download PDF annexure',f.read(),os.path.basename(pdf_path),'application/pdf')
                    elif not unresolved.empty: st.info('Resolve or approve all unresolved records before creating a final approved annexure.')
                else: st.info('There are no approved entries available for the annexure.')
    st.markdown('#### Previously generated annexures')
    hist=qdf('''SELECT a.annexure_ref Reference,c.name Client,p.name Project,a.month Month,a.category Category,a.version Version,a.status Status,a.total Total,a.generated_by Generated_By,a.generated_at Generated_At FROM annexures a JOIN clients c ON c.id=a.client_id JOIN projects p ON p.id=a.project_id ORDER BY a.generated_at DESC''')
    if not hist.empty: st.dataframe(hist,use_container_width=True,hide_index=True)

# ---------- Import center ----------
elif page=='Excel Import Center':
    st.subheader('Excel / CSV Import Center')
    st.write('Use this to bring in existing sales or cost records. Preview and map columns before saving. This tool does not assume that every source workbook tab is a sales sheet.')
    kind=st.selectbox('Import into',['Sales register','Cost register','Commercial rates'])
    file=st.file_uploader('Upload Excel or CSV',type=['xlsx','xls','csv'])
    if file:
        try:
            if file.name.lower().endswith('.csv'):
                df=pd.read_csv(file); sheet=None
            else:
                xls=pd.ExcelFile(file); sheet=st.selectbox('Worksheet',xls.sheet_names); df=pd.read_excel(file,sheet_name=sheet)
            df=df.dropna(how='all').copy()
            st.write(f'{len(df)} rows · {len(df.columns)} columns'); st.dataframe(df.head(30),use_container_width=True)
            st.download_button('Download source preview as CSV',df.to_csv(index=False).encode('utf-8-sig'),'import_preview.csv','text/csv')
            st.caption('Choose how to map columns. Existing duplicate records are skipped by reference/date/category/amount fingerprint where possible.')
            columns=['(Not mapped)']+list(df.columns)
            if kind=='Sales register':
                with st.form('salesimportmap'):
                    a,b,c=st.columns(3)
                    datecol=a.selectbox('Service date column',columns,index=columns.index(next((x for x in columns if x.lower().strip() in ['date','date ','service date','service_date']), '(Not mapped)')) if next((x for x in columns if x.lower().strip() in ['date','date ','service date','service_date']),None) else 0)
                    fromcol=b.selectbox('Source / From column',columns,index=next((i for i,x in enumerate(columns) if x.lower().strip() in ['from','source','from city']),0))
                    tocol=c.selectbox('Destination / To column',columns,index=next((i for i,x in enumerate(columns) if x.lower().strip() in ['to','destination','to city']),0))
                    a,b,c=st.columns(3); amountcol=a.selectbox('Amount column',columns,index=next((i for i,x in enumerate(columns) if x.lower().strip() in ['amount','amt','sales amount']),0)); ratecol=b.selectbox('Rate column (optional)',columns); qtycol=c.selectbox('Quantity column (optional)',columns)
                    a,b,c=st.columns(3); catcol=a.selectbox('Category column',columns,index=next((i for i,x in enumerate(columns) if x.lower().strip() in ['category','remark','service','fruit type']),0)); servicecol=b.selectbox('Product/service column',columns); shipcol=c.selectbox('Shipment reference column',columns)
                    cl=st.selectbox('Assign imported rows to client',list(client_opts) if client_opts else ['No clients created']); pr=st.selectbox('Assign imported rows to project',list(project_opts) if project_opts else ['No projects created']); import_status=st.selectbox('Initial status',['Pending','Approved'] if CAN_APPROVE else ['Pending'],index=0); commit=st.form_submit_button('Validate and import sales')
                if commit:
                    if import_status=='Approved' and not CAN_APPROVE: st.error('Only Admin or Finance can import records as Approved.')
                    elif not client_opts or not project_opts or datecol=='(Not mapped)' or amountcol=='(Not mapped)': st.error('Create client/project and map at least date and amount columns.')
                    else:
                        added=0; skipped=0; errors=[]
                        with db() as import_con:
                          for ix,r in df.iterrows():
                            try:
                                d=parse_import_date(r[datecol])
                                if pd.isna(d): raise ValueError('Invalid date')
                                amount=to_num(r[amountcol],-1);
                                if amount<=0: raise ValueError('Missing, zero or invalid amount')
                                qty=to_num(r[qtycol],1) if qtycol!='(Not mapped)' else 1; rate=to_num(r[ratecol],0) if ratecol!='(Not mapped)' else (amount/qty if qty else amount)
                                cat=clean_str(r[catcol]) if catcol!='(Not mapped)' else 'Other'; cat=cat or 'Other'
                                src=clean_str(r[fromcol]) if fromcol!='(Not mapped)' else ''; dst=clean_str(r[tocol]) if tocol!='(Not mapped)' else ''
                                service=clean_str(r[servicecol]) if servicecol!='(Not mapped)' else ''; ship=clean_str(r[shipcol]) if shipcol!='(Not mapped)' else ''
                                fingerprint=import_con.execute('SELECT id FROM sales WHERE client_id=? AND project_id=? AND service_date=? AND category=? AND amount=? AND source=? AND destination=?',(client_opts[cl],project_opts[pr],d.date().isoformat(),cat,amount,src,dst)).fetchone()
                                if fingerprint: skipped+=1; continue
                                apby,apat=approval_fields(import_status)
                                import_con.execute('INSERT INTO sales(entry_ref,client_id,project_id,service_date,category,service,source,destination,shipment_ref,quantity,unit,rate,amount,imported_amount,remarks,status,created_by,created_at,updated_at,source_file,approved_by,approved_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)',(make_ref('IMP'),client_opts[cl],project_opts[pr],d.date().isoformat(),cat,service,src,dst,ship,qty,'shipment',rate,amount,amount,'Imported from '+file.name,import_status,username,datetime.now().isoformat(timespec='seconds'),datetime.now().isoformat(timespec='seconds'),file.name,apby,apat))
                                added+=1
                            except Exception as e: errors.append({'Row':int(ix)+2,'Error':str(e)})
                        audit('IMPORT','sales',details=f'{file.name}; added={added}; skipped={skipped}; errors={len(errors)}')
                        st.success(f'Imported {added} sales rows; skipped {skipped} likely duplicates; errors {len(errors)}.')
                        if errors: st.dataframe(pd.DataFrame(errors),use_container_width=True,hide_index=True)
                        st.rerun()
            elif kind=='Cost register':
                with st.form('costimportmap'):
                    a,b,c=st.columns(3); datecol=a.selectbox('Cost date column',columns,index=next((i for i,x in enumerate(columns) if x.lower().strip() in ['date','cost date','cost_date']),0)); amountcol=b.selectbox('Cost amount column',columns,index=next((i for i,x in enumerate(columns) if x.lower().strip() in ['amount','mm cost','adhoc cost','per day cost','cost']),0)); catcol=c.selectbox('Category column',columns,index=next((i for i,x in enumerate(columns) if x.lower().strip() in ['category','remark','cost category','type']),0))
                    a,b,c=st.columns(3); vendorcol=a.selectbox('Vendor column',columns); refcol=b.selectbox('Reference column',columns); remarkscol=c.selectbox('Remarks column',columns)
                    cl=st.selectbox('Assign to client',list(client_opts) if client_opts else ['No clients created'],key='costimportcl'); pr=st.selectbox('Assign to project',list(project_opts) if project_opts else ['No projects created'],key='costimportpr'); import_status=st.selectbox('Initial status',['Pending','Approved'] if CAN_APPROVE else ['Pending'],key='costimportstatus'); commit=st.form_submit_button('Validate and import costs')
                if commit:
                    if import_status=='Approved' and not CAN_APPROVE: st.error('Only Admin or Finance can import records as Approved.')
                    elif not client_opts or not project_opts: st.error('Create client and project first.')
                    else:
                        added=0; skipped=0; errors=[]
                        with db() as import_con:
                          for ix,r in df.iterrows():
                            try:
                                d=parse_import_date(r[datecol]) if datecol!='(Not mapped)' else pd.NaT
                                if pd.isna(d): raise ValueError('Invalid/missing date')
                                amount=to_num(r[amountcol],-1)
                                if amount<=0: raise ValueError('Missing, zero or invalid amount')
                                cat=clean_str(r[catcol]) if catcol!='(Not mapped)' else 'Other'; cat=cat or 'Other'
                                vendor=clean_str(r[vendorcol]) if vendorcol!='(Not mapped)' else ''; ref=clean_str(r[refcol]) if refcol!='(Not mapped)' else ''; remarks=clean_str(r[remarkscol]) if remarkscol!='(Not mapped)' else ''
                                dup=import_con.execute('SELECT id FROM costs WHERE client_id=? AND project_id=? AND cost_date=? AND category=? AND amount=? AND reference_no=?',(client_opts[cl],project_opts[pr],d.date().isoformat(),cat,amount,ref)).fetchone()
                                if dup: skipped+=1; continue
                                apby,apat=approval_fields(import_status)
                                import_con.execute('INSERT INTO costs(entry_ref,client_id,project_id,cost_date,category,vendor,reference_no,amount,remarks,status,created_by,created_at,source_file,approved_by,approved_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)',(make_ref('IMP'),client_opts[cl],project_opts[pr],d.date().isoformat(),cat,vendor,ref,amount,remarks,import_status,username,datetime.now().isoformat(timespec='seconds'),file.name,apby,apat))
                                added+=1
                            except Exception as e: errors.append({'Row':int(ix)+2,'Error':str(e)})
                        audit('IMPORT','costs',details=f'{file.name}; added={added}; skipped={skipped}; errors={len(errors)}'); st.success(f'Imported {added} cost rows; skipped {skipped} duplicates; errors {len(errors)}.')
                        if errors: st.dataframe(pd.DataFrame(errors),use_container_width=True,hide_index=True)
                        st.rerun()
            else:
                st.caption('Map columns from the selected rate-card sheet. If client/project columns are not present, assign all imported rates to one client and project below.')
                with st.form('commercialimportmap'):
                    a,b,c=st.columns(3)
                    clientcol=a.selectbox('Client column (optional)',columns,key='imp_rate_client')
                    projectcol=b.selectbox('Project column (optional)',columns,key='imp_rate_project')
                    ratecol=c.selectbox('Rate column',columns,index=next((i for i,x in enumerate(columns) if x.lower().strip() in ['rate','amount','commercial rate','price']),0),key='imp_rate_rate')
                    a,b,c=st.columns(3)
                    catcol=a.selectbox('Category column (optional)',columns,key='imp_rate_cat')
                    servicecol=b.selectbox('Service / description column (optional)',columns,key='imp_rate_service')
                    unitcol=c.selectbox('Unit column (optional)',columns,key='imp_rate_unit')
                    a,b,c=st.columns(3)
                    sourcecol=a.selectbox('Source / From column (optional)',columns,key='imp_rate_source')
                    destcol=b.selectbox('Destination / To column (optional)',columns,key='imp_rate_dest')
                    datecol=c.selectbox('Effective-from date column (optional)',columns,key='imp_rate_date')
                    a,b=st.columns(2)
                    assign_client=a.selectbox('Default client',list(client_opts) if client_opts else ['No clients created'],key='imp_rate_default_client')
                    assign_project=b.selectbox('Default project',list(project_opts) if project_opts else ['No projects created'],key='imp_rate_default_project')
                    default_cat=st.selectbox('Default category when no category column is mapped',category_list,key='imp_rate_default_cat')
                    initial_status=st.selectbox('Initial approval status',['Pending','Approved'] if CAN_APPROVE else ['Pending'],index=0,key='imp_rate_status')
                    commit=st.form_submit_button('Validate and import commercial rates')
                if commit:
                    if initial_status=='Approved' and not CAN_APPROVE: st.error('Only Admin or Finance can import rates as Approved.')
                    elif not client_opts or not project_opts or ratecol=='(Not mapped)':
                        st.error('Create at least one client and project, and map a rate column.')
                    else:
                        added=0; skipped=0; errors=[]
                        with db() as import_con:
                          for ix,r in df.iterrows():
                            try:
                                rate_value=to_num(r[ratecol],-1)
                                if rate_value<0: raise ValueError('Missing or invalid rate')
                                clname=clean_str(r[clientcol]) if clientcol!='(Not mapped)' else ''
                                prname=clean_str(r[projectcol]) if projectcol!='(Not mapped)' else ''
                                if clname and clientcol!='(Not mapped)':
                                    matches=qdf('SELECT id FROM clients WHERE lower(name)=lower(?)',(clname,))
                                    if matches.empty: raise ValueError(f'Client not found: {clname}; create it first or leave Client column unmapped and use default')
                                    cid=int(matches.iloc[0]['id'])
                                else: cid=client_opts[assign_client]
                                if prname and projectcol!='(Not mapped)':
                                    matches=qdf('SELECT id FROM projects WHERE client_id=? AND lower(name)=lower(?)',(cid,prname))
                                    if matches.empty: raise ValueError(f'Project not found for client: {prname}; create it first or use default project')
                                    pid=int(matches.iloc[0]['id'])
                                else:
                                    pid=project_opts[assign_project]
                                    project_client=qdf('SELECT client_id FROM projects WHERE id=?',(pid,))
                                    if not project_client.empty and int(project_client.iloc[0]['client_id'])!=cid:
                                        if clname: raise ValueError('Mapped client and default project do not belong together')
                                        cid=int(project_client.iloc[0]['client_id'])
                                cat=clean_str(r[catcol]) if catcol!='(Not mapped)' else default_cat
                                cat=cat or default_cat
                                service=clean_str(r[servicecol]) if servicecol!='(Not mapped)' else ''
                                unit_value=clean_str(r[unitcol]) if unitcol!='(Not mapped)' else 'shipment'; unit_value=unit_value or 'shipment'
                                src=clean_str(r[sourcecol]) if sourcecol!='(Not mapped)' else ''
                                dst=clean_str(r[destcol]) if destcol!='(Not mapped)' else ''
                                if datecol!='(Not mapped)':
                                    d=parse_import_date(r[datecol])
                                    if pd.isna(d): raise ValueError('Invalid effective-from date')
                                    effective=d.date().isoformat()
                                else: effective=date.today().isoformat()
                                dup=import_con.execute('SELECT id FROM commercials WHERE client_id=? AND project_id=? AND category=? AND source=? AND destination=? AND rate=? AND effective_from=?',(cid,pid,cat,src,dst,rate_value,effective)).fetchone()
                                if dup: skipped+=1; continue
                                apby,apat=approval_fields(initial_status)
                                import_con.execute('INSERT INTO commercials(client_id,project_id,category,service,source,destination,unit,rate,effective_from,effective_to,status,notes,created_by,created_at,approved_by,approved_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)',(cid,pid,cat,service,src,dst,unit_value,rate_value,effective,'',initial_status,'Imported from '+file.name,username,datetime.now().isoformat(timespec='seconds'),apby,apat))
                                added+=1
                            except Exception as e: errors.append({'Row':int(ix)+2,'Error':str(e)})
                        audit('IMPORT','commercials',details=f'{file.name}; added={added}; skipped={skipped}; errors={len(errors)}')
                        st.success(f'Imported {added} commercial rates; skipped {skipped} duplicates; errors {len(errors)}.')
                        if errors: st.dataframe(pd.DataFrame(errors),use_container_width=True,hide_index=True)
                        st.rerun()
        except Exception as e: st.error(f'Could not read file: {e}')

# ---------- User management and audit ----------
elif page=='User Management & Audit':
    st.subheader('User Management & Audit Trail')
    if role!='Admin': st.warning('Only Admin can manage users. Audit events are visible below.')
    else:
        with st.form('userform'):
            a,b,c=st.columns(3); nu=a.text_input('New username'); npw=b.text_input('Temporary password',type='password'); nr=c.selectbox('Role',['Admin','Sales','Operations/Cost','Finance','Management'])
            add=st.form_submit_button('Create user')
        if add:
            if not nu.strip() or len(npw)<8: st.error('Enter a username and a password of at least 8 characters.')
            else:
                try:
                    with db() as c: c.execute('INSERT INTO users(username,password_hash,role,created_at) VALUES(?,?,?,?)',(nu.strip(),hashlib.sha256(npw.encode()).hexdigest(),nr,datetime.now().isoformat(timespec='seconds')))
                    audit('CREATE','users',details=nu); st.success('User created.'); st.rerun()
                except sqlite3.IntegrityError: st.warning('Username already exists.')
        with st.expander('Change my password'):
            with st.form('pwform'):
                old=st.text_input('Current password',type='password'); new=st.text_input('New password (minimum 8 characters)',type='password'); ch=st.form_submit_button('Change password')
            if ch:
                row=qdf('SELECT password_hash FROM users WHERE username=?',(username,))
                if not row.empty and row.iloc[0]['password_hash']==hashlib.sha256(old.encode()).hexdigest() and len(new)>=8:
                    with db() as c: c.execute('UPDATE users SET password_hash=? WHERE username=?',(hashlib.sha256(new.encode()).hexdigest(),username))
                    audit('UPDATE','users',details='Password changed'); st.success('Password changed.')
                else: st.error('Current password is incorrect or the new password is too short.')
        st.dataframe(qdf('SELECT id ID,username Username,role Role,active Active,created_at Created_At FROM users ORDER BY username'),use_container_width=True,hide_index=True)
    st.markdown('#### Audit log')
    audit_df=qdf('SELECT created_at Timestamp,username User,action Action,table_name Table_Name,record_id Record_ID,details Details FROM audit_log ORDER BY id DESC LIMIT 500')
    st.dataframe(audit_df,use_container_width=True,hide_index=True)
    if not audit_df.empty: st.download_button('Download audit log CSV',audit_df.to_csv(index=False).encode('utf-8-sig'),'audit_log.csv','text/csv')

st.sidebar.divider()
st.sidebar.caption('Local database: celcius_sales_cost.db')
st.sidebar.caption('Back up the database file regularly. For shared concurrent use, deploy with PostgreSQL and access controls.')
