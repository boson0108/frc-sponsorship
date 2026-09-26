import streamlit as st
import os
import re
import smtplib
import time
import imaplib
import json
import hashlib
import string
import email
from email.header import decode_header
from email.utils import parseaddr
from filelock import FileLock
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import openpyxl

try:
    import pandas as pd
except ImportError:
    pd = None

# ==========================================
# 網頁基本設定
# ==========================================
st.set_page_config(page_title="公關暨招募寄信系統", layout="wide")

USERS_FILE = "users.json"
USERS_LOCK = "users.json.lock"
PROJECTS_FILE = "projects.json"
PROJECTS_LOCK = "projects.json.lock"

def decode_str(s):
    if not s: return ""
    try:
        value, charset = decode_header(s)[0]
        if charset: return value.decode(charset)
        elif isinstance(value, bytes): return value.decode('utf-8', errors='ignore')
        return str(value)
    except Exception:
        return str(s)

class SafeFormatter(string.Formatter):
    def get_value(self, key, args, kwargs):
        if isinstance(key, str): return kwargs.get(key, f"{{{key}}}")
        return super().get_value(key, args, kwargs)

def safe_format_template(template, context_dict):
    if not template: return ""
    try:
        formatter = SafeFormatter()
        return formatter.format(template, **context_dict)
    except Exception:
        res = str(template)
        for k, v in context_dict.items(): res = res.replace(f"{{{k}}}", str(v))
        return res

def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()

def safe_replace_file(src, dst, max_retries=5):
    for attempt in range(max_retries):
        try:
            os.replace(src, dst)
            return
        except PermissionError:
            if attempt == max_retries - 1: raise
            time.sleep(0.02)

def load_users():
    with FileLock(USERS_LOCK, timeout=10):
        if not os.path.exists(USERS_FILE):
            default_users = {"admin": {"password": hash_password("admin123"), "role": "admin", "real_name": "系統管理員"}}
            temp_file = f"{USERS_FILE}.tmp"
            with open(temp_file, "w", encoding="utf-8") as f: json.dump(default_users, f, ensure_ascii=False, indent=4)
            safe_replace_file(temp_file, USERS_FILE)
            return default_users
        with open(USERS_FILE, "r", encoding="utf-8") as f:
            try: return json.load(f)
            except Exception: return {}

def save_users(users_data):
    with FileLock(USERS_LOCK, timeout=10):
        temp_file = f"{USERS_FILE}.tmp"
        with open(temp_file, "w", encoding="utf-8") as f: json.dump(users_data, f, ensure_ascii=False, indent=4)
        safe_replace_file(temp_file, USERS_FILE)

def load_projects():
    with FileLock(PROJECTS_LOCK, timeout=10):
        if not os.path.exists(PROJECTS_FILE): return {}
        with open(PROJECTS_FILE, "r", encoding="utf-8") as f:
            try: 
                data = json.load(f)
                for p_name, p_data in data.items():
                    if "mode" not in p_data: p_data["mode"] = "individual"
                    if "replies" not in p_data: p_data["replies"] = []
                return data
            except Exception: return {}

def save_projects(projects_data):
    with FileLock(PROJECTS_LOCK, timeout=10):
        temp_file = f"{PROJECTS_FILE}.tmp"
        with open(temp_file, "w", encoding="utf-8") as f: json.dump(projects_data, f, ensure_ascii=False, indent=4)
        safe_replace_file(temp_file, PROJECTS_FILE)

def extract_email(contact_str):
    if not contact_str or str(contact_str).lower() in ("nan", "none", "null"): return None
    match = re.search(r'[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+', str(contact_str))
    return match.group(0) if match else None

def read_excel_data(uploaded_file):
    if pd is not None:
        try:
            df = pd.read_excel(uploaded_file, sheet_name="全部彙總清單")
            return df.to_dict(orient="records")
        except Exception:
            try: return pd.read_excel(uploaded_file, sheet_name=0).to_dict(orient="records")
            except Exception: pass
    try:
        wb = openpyxl.load_workbook(uploaded_file, data_only=True)
        sheet_name = "全部彙總清單" if "全部彙總清單" in wb.sheetnames else wb.sheetnames[0]
        ws = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))
        if not rows: return []
        headers = [str(h).strip() if h is not None else "" for h in rows[0]]
        data = []
        for row in rows[1:]:
            if any(c is not None for c in row):
                data.append({headers[i]: (row[i] if i < len(row) and row[i] is not None else "") for i in range(len(headers))})
        return data
    except Exception as e:
        st.error(f"❌ Excel 讀取失敗：{e}")
        return []

DEFAULT_TEMPLATE = ""
# ==========================================
# 狀態初始化與登入介面
# ==========================================
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.username, st.session_state.role, st.session_state.real_name = "", "", ""
if "page" not in st.session_state: st.session_state.page = "home"
if "current_project" not in st.session_state: st.session_state.current_project = None
if "gmail_account" not in st.session_state: st.session_state.gmail_account = ""
if "gmail_password" not in st.session_state: st.session_state.gmail_password = ""

if not st.session_state.logged_in:
    st.title("🔐 公關暨招募寄信系統")
    st.markdown("請輸入您的專屬帳號與密碼以登入系統。")
    with st.form("login_form"):
        login_user = st.text_input("帳號").strip()
        login_pwd = st.text_input("密碼", type="password").strip()
        if st.form_submit_button("登入", type="primary"):
            users_db = load_users()
            if login_user in users_db and users_db[login_user]["password"] == hash_password(login_pwd):
                st.session_state.logged_in = True
                st.session_state.username = login_user
                st.session_state.role = users_db[login_user]["role"]
                st.session_state.real_name = users_db[login_user].get("real_name", login_user)
                st.rerun()
            else: st.error("⚠️ 帳號或密碼錯誤！")
    st.stop()

projects_db = load_projects()
unread_count = sum(1 for p in projects_db.values() for r in p.get("replies", []) if not r.get("read", True) and (st.session_state.role == "admin" or r.get("owner") == st.session_state.real_name))

st.sidebar.markdown(f"👤 登入者：**{st.session_state.real_name}**")
if st.sidebar.button("🚪 登出", use_container_width=True):
    st.session_state.logged_in = False
    st.rerun()
st.sidebar.divider()

nav_options = ["🏠 專案與寄信區", f"📥 收件與回信匣 {'🔴' if unread_count > 0 else ''}"]
if st.session_state.role == "admin": nav_options.append("⚙️ 系統後台管理")
app_mode = st.sidebar.radio("📌 系統功能導覽", nav_options)
st.sidebar.divider()

# ==========================================
# 模式 A：系統後台管理
# ==========================================
if "⚙️ 系統後台管理" in app_mode:
    st.title("⚙️ 系統後台管理")
    tab1, tab2 = st.tabs(["👥 帳號管理與重設密碼", "📊 團隊寄件總覽"])
    with tab1:
        col_new, col_reset = st.columns(2)
        users_db = load_users()
        with col_new:
            st.subheader("➕ 建立新帳號")
            with st.form("add_user_form"):
                n_usr = st.text_input("登入帳號").strip()
                n_name = st.text_input("成員姓名").strip()
                n_pwd = st.text_input("預設密碼", type="password").strip()
                n_role = st.selectbox("帳號權限", ["user (一般)", "admin (管理員)"])
                if st.form_submit_button("新增帳號", type="primary", use_container_width=True):
                    if n_usr in users_db: st.error("帳號已存在！")
                    elif not n_usr or not n_pwd: st.error("不得為空！")
                    else:
                        users_db[n_usr] = {"password": hash_password(n_pwd), "role": "admin" if "admin" in n_role else "user", "real_name": n_name or n_usr}
                        save_users(users_db)
                        st.success(f"成功建立：{n_name}")
                        st.rerun()
        with col_reset:
            st.subheader("🔑 重設成員密碼")
            with st.form("reset_pwd_form"):
                target_user = st.selectbox("選擇帳號", list(users_db.keys()))
                new_pwd = st.text_input("新密碼", type="password").strip()
                if st.form_submit_button("重設密碼", type="primary", use_container_width=True):
                    if not new_pwd: st.error("不得為空！")
                    else:
                        users_db[target_user]["password"] = hash_password(new_pwd)
                        save_users(users_db)
                        st.success("✅ 密碼已重設！")
                        st.rerun()
        st.table([{"帳號": u, "姓名": d.get("real_name", u), "權限": d.get("role", "user")} for u, d in users_db.items()])
    with tab2:
        st.subheader("📂 寄件紀錄總覽")
        for p_name, p_data in projects_db.items():
            sent_list = p_data.get("sent_companies", [])
            with st.expander(f"📁 {p_name} (共 {len(sent_list)} 封)"):
                for record in sent_list:
                    if isinstance(record, dict): st.write(f"- {record.get('company', '?')} (負責人: {record.get('sender', '?')})")
                    else: st.write(f"- {record}")

# ==========================================
# 模式 B：收件與回信匣
# ==========================================
elif "📥 收件與回信匣" in app_mode:
    st.title("📥 廠商回信與通知中心")
    st.subheader("1. 郵件伺服器認證")
    col1, col2, col3 = st.columns([3, 3, 2])
    test_email = col1.text_input("團隊 Gmail 信箱", value=st.session_state.gmail_account).strip()
    test_pwd = col2.text_input("應用程式密碼", value=st.session_state.gmail_password, type="password").strip()
    reply_folder = col3.text_input("歸檔資料夾", value="FRC_Replies")
    
    if st.button("🔄 強制掃描近期回信", type="primary", use_container_width=True):
        if not test_email or not test_pwd: st.error("請輸入信箱密碼！")
        else:
            with st.spinner("🚀 掃描中..."):
                try:
                    mail = imaplib.IMAP4_SSL("imap.gmail.com", timeout=20)
                    mail.login(test_email, test_pwd)
                    status, _ = mail.select(reply_folder)
                    if status != 'OK': mail.create(reply_folder)
                    mail.select("INBOX")
                    date_limit = (pd.Timestamp.now() - pd.Timedelta(days=3)).strftime("%d-%b-%Y")
                    status, messages = mail.search(None, f'(SINCE "{date_limit}")')
                    if status == "OK" and messages[0]:
                        msg_nums = messages[0].split()
                        sent_map = {}
                        for p_name, p_data in projects_db.items():
                            for record in p_data.get("sent_companies", []):
                                if isinstance(record, dict) and record.get("email"):
                                    sent_map[str(record["email"]).strip().lower()] = (p_name, record.get("company"), record.get("sender"))
                        new_reply_count = 0
                        if len(msg_nums) > 100: msg_nums = msg_nums[-100:]
                        for num in msg_nums:
                            res, header_data = mail.fetch(num, "(BODY.PEEK[HEADER.FIELDS (FROM SUBJECT)])")
                            for response_part in header_data:
                                if isinstance(response_part, tuple):
                                    header_msg = email.message_from_bytes(response_part[1])
                                    from_header = decode_str(header_msg.get("From"))
                                    _, addr = parseaddr(from_header)
                                    addr_lower = str(addr).strip().lower()
                                    subject_check = decode_str(header_msg.get("Subject"))
                                    if addr_lower in sent_map:
                                        proj_name, company_name, original_sender = sent_map[addr_lower]
                                        if not any(r["subject"] == subject_check and r["email"] == addr_lower for r in projects_db[proj_name].get("replies", [])):
                                            res, full_msg_data = mail.fetch(num, "(RFC822)")
                                            for f_part in full_msg_data:
                                                if isinstance(f_part, tuple):
                                                    msg = email.message_from_bytes(f_part[1])
                                                    body = "無法解析"
                                                    if msg.is_multipart():
                                                        for part in msg.walk():
                                                            if part.get_content_type() == "text/plain":
                                                                try: body = part.get_payload(decode=True).decode('utf-8', errors='ignore'); break
                                                                except: pass
                                                    else:
                                                        try: body = msg.get_payload(decode=True).decode('utf-8', errors='ignore')
                                                        except: pass
                                                    clipped_body = body[:1000] + ("\n...(略)" if len(body)>1000 else "")
                                                    projects_db[proj_name]["replies"].append({"company": company_name, "email": addr, "subject": subject_check, "body": clipped_body, "read": False, "time": time.strftime("%Y-%m-%d %H:%M"), "owner": original_sender})
                                                    new_reply_count += 1
                                                    mail.copy(num, reply_folder)
                                                    mail.store(num, '+FLAGS', '\\Deleted')
                        mail.expunge()
                        if new_reply_count > 0:
                            save_projects(projects_db)
                            st.success(f"攔截 {new_reply_count} 封新回信！")
                            time.sleep(1)
                            st.rerun()
                        else: st.info("沒有新回信。")
                    else: st.info("無信件。")
                    mail.logout()
                except Exception as e: st.error(f"失敗：{e}")

    st.divider()
    st.subheader("2. 回信匣")
    has_any = False
    for p_name, p_data in projects_db.items():
        replies = p_data.get("replies", [])
        visible = [(idx, r) for idx, r in enumerate(replies) if st.session_state.role == "admin" or r.get("owner") == st.session_state.real_name]
        if not visible: continue
        has_any = True
        st.write(f"#### 📁 {p_name}")
        for real_idx, reply in reversed(visible):
            icon = "🔴" if not reply.get("read", True) else "🟢"
            with st.expander(f"{icon} {reply['company']} ({reply['email']}) - {reply['subject']}"):
                st.caption(f"時間：{reply.get('time', '')} | 負責人：{reply.get('owner', '')}")
                st.text(reply['body'])
                if not reply.get("read", True) and st.button("標為已讀", key=f"read_{p_name}_{real_idx}"):
                    projects_db[p_name]["replies"][real_idx]["read"] = True
                    save_projects(projects_db)
                    st.rerun()
    if not has_any: st.info("尚無紀錄。")
# ==========================================
# 模式 C：專案與寄信區
# ==========================================
elif "🏠 專案與寄信區" in app_mode:
    if st.session_state.page == "home":
        st.title("✉️ 公關寄信系統 - 專案大廳")
        col_new, col_list = st.columns([1, 2])
        with col_new:
            st.subheader("➕ 建立新專案")
            with st.form("new_project_form"):
                new_proj_name = st.text_input("專案命名").strip()
                proj_mode = st.radio("模式", ["個別寄送", "一鍵群發"])
                if st.form_submit_button("建立專案", type="primary", use_container_width=True):
                    if not new_proj_name: st.error("不能為空！")
                    elif new_proj_name in projects_db: st.error("已存在！")
                    else:
                        # 建立專案時，給予預設的主旨
                        projects_db[new_proj_name] = {"sent_companies": [], "template": "", "subject": "【錄取通知】明道中學學生會 — 敬致 {企業／贊助單位}", "replies": [], "mode": "bulk" if "群發" in proj_mode else "individual"}
                        save_projects(projects_db)
                        st.rerun()
        with col_list:
            st.subheader("📂 專案列表")
            for proj_name, proj_data in list(projects_db.items()):
                sents = proj_data.get('sent_companies', [])
                count = len(sents) if st.session_state.role == "admin" else sum(1 for r in sents if isinstance(r, dict) and r.get('sender') == st.session_state.real_name)
                with st.container(border=True):
                    c1, c2, c3 = st.columns([6, 2, 2])
                    with c1:
                        st.write(f"**{proj_name}** [{'群發' if proj_data.get('mode')=='bulk' else '個別'}]")
                        st.caption(f"寄出 {count} 封")
                    with c2:
                        if st.button("開啟", key=f"open_{proj_name}", use_container_width=True):
                            st.session_state.current_project = proj_name
                            st.session_state.page = "project"
                            st.rerun()
                    with c3:
                        if st.session_state.role == "admin" and st.button("刪除", key=f"del_{proj_name}", use_container_width=True):
                            del projects_db[proj_name]
                            save_projects(projects_db)
                            st.rerun()

    elif st.session_state.page == "project":
        curr_proj = st.session_state.current_project
        p_data = projects_db[curr_proj]
        if st.sidebar.button("🔙 返回大廳", type="primary", use_container_width=True):
            st.session_state.page = "home"
            st.rerun()
        st.sidebar.divider()
        st.sidebar.header("🔐 寄件帳號設定")
        sender_email = st.sidebar.text_input("Gmail", value=st.session_state.gmail_account).strip()
        sender_password = st.sidebar.text_input("密碼", value=st.session_state.gmail_password, type="password").strip()
        st.session_state.gmail_account, st.session_state.gmail_password = sender_email, sender_password
        st.sidebar.header("📂 系統設定")
        pdf_dir = st.sidebar.text_input("附件資料夾", value="企劃書檔案").strip()
        backup_folder = st.sidebar.text_input("備份標籤", value="FRC_Sponsorship").strip()

        st.title(f"📁 專案：{curr_proj}")
        st.divider()

        st.header("Step 1: 團隊與信件格式")
        c1, c2, c3 = st.columns(3)
        team_name = c1.text_input("團隊名稱", value="").strip()
        contact_person = c2.text_input("聯絡人", value="").strip()
        contact_phone = c3.text_input("電話", value="").strip()
        
        # 🌟 新增：信件主旨自訂欄位
        email_subject = st.text_input("📌 信件主旨 (標題)", value=p_data.get("subject", "【錄取通知】明道中學學生會 — 敬致 {企業／贊助單位}"), placeholder="例如：【錄取通知】明道中學學生會 — 敬致 {企業／贊助單位}")
        
        email_template = st.text_area("✏️ 內容", value=p_data.get("template", ""), height=200, placeholder="信件內容，可使用 {變數}，並在此處直接貼上網址或連結")
        
        # 儲存變更
        if email_template != p_data.get("template") or email_subject != p_data.get("subject"):
            projects_db[curr_proj]["template"] = email_template
            projects_db[curr_proj]["subject"] = email_subject
            save_projects(projects_db)

        st.header("Step 2 & 3: 載入與寄出")
        
        uploaded_file = st.file_uploader("📊 上傳 Excel 名單", type=["xlsx"])
        
        sent_list = p_data.get("sent_companies", [])
        all_sents = [r.get("company") if isinstance(r, dict) else r for r in sent_list]
        
        if uploaded_file:
            records = read_excel_data(uploaded_file)
            if records:
                if p_data.get("mode") == "bulk":
                    unsent = [r for r in records if r.get('企業／贊助單位') and r.get('企業／贊助單位') not in all_sents]
                    if not unsent: st.success("全數寄出！")
                    else:
                        st.info(f"尚有 {len(unsent)} 名未寄送。")
                        first = unsent[0]
                        fmt_dict = dict(first)
                        fmt_dict.update({"team_name": team_name, "contact_person": contact_person, "contact_phone": contact_phone, "pdf_filename": f"{str(first.get('編號', '000')).zfill(3)}_{first.get('企業／贊助單位')}_贊助企劃書.pdf"})
                        prev_text = safe_format_template(email_template, fmt_dict)
                        
                        # 🌟 新增：預覽主旨變化
                        prev_subj = safe_format_template(email_subject, fmt_dict)
                        st.write(f"**✉️ 預覽主旨：** {prev_subj}")
                        
                        st.write("--- 預覽內容 ---")
                        st.text(prev_text)
                        st.write("-------------")
                        
                        if st.button(f"🚀 群發 {len(unsent)} 人", type="primary"):
                            if not sender_email or not sender_password: st.error("帳號未設！")
                            else:
                                pb = st.progress(0)
                                stxt = st.empty()
                                try:
                                    server = smtplib.SMTP("smtp.gmail.com", 587, timeout=15)
                                    server.starttls()
                                    server.login(sender_email, sender_password)
                                    
                                    try:
                                        imap = imaplib.IMAP4_SSL("imap.gmail.com", timeout=15)
                                        imap.login(sender_email, sender_password)
                                        if imap.select(backup_folder)[0] != 'OK': imap.create(backup_folder)
                                    except:
                                        imap = None
                                        
                                    success = 0
                                    for idx, r_data in enumerate(unsent):
                                        comp = r_data.get('企業／贊助單位')
                                        to_mail = extract_email(r_data.get('聯絡資訊', ''))
                                        stxt.text(f"寄送中: {comp}...")
                                        if to_mail:
                                            pdf_name = f"{str(r_data.get('編號', '000')).zfill(3)}_{comp}_贊助企劃書.pdf"
                                            pdf_p = os.path.join(pdf_dir, pdf_name)
                                            f_dict = dict(r_data)
                                            f_dict.update({"team_name": team_name, "contact_person": contact_person, "contact_phone": contact_phone, "pdf_filename": pdf_name})
                                            
                                            raw_t = safe_format_template(email_template, f_dict)
                                            # 🌟 新增：產生自訂主旨
                                            formatted_subj = safe_format_template(email_subject, f_dict)
                                            
                                            msg = MIMEMultipart()
                                            msg['From'] = sender_email
                                            msg['To'] = to_mail
                                            msg['Subject'] = formatted_subj  # 使用自訂主旨
                                            
                                            msg.attach(MIMEText(raw_t, 'plain', 'utf-8'))
                                                
                                            if os.path.exists(pdf_p):
                                                with open(pdf_p, 'rb') as f:
                                                    att = MIMEApplication(f.read(), _subtype="pdf")
                                                    att.add_header('Content-Disposition', 'attachment', filename=pdf_name)
                                                    msg.attach(att)
                                                    
                                            server.send_message(msg)
                                            if imap:
                                                try: imap.append(backup_folder, '\\Seen', imaplib.Time2Internaldate(time.time()), msg.as_bytes())
                                                except: pass
                                            
                                            projects_db[curr_proj]["sent_companies"].append({"company": comp, "sender": st.session_state.real_name, "email": to_mail})
                                            success += 1
                                            time.sleep(1)
                                        pb.progress((idx+1)/len(unsent))
                                    server.quit()
                                    if imap: imap.logout()
                                    save_projects(projects_db)
                                    stxt.text(f"✅ 完成！寄出 {success} 封")
                                    time.sleep(2)
                                    st.rerun()
                                except Exception as e: st.error(f"中斷：{e}")
                else:
                    c_list = list(set([r.get('企業／贊助單位', '') for r in records if r.get('企業／贊助單位')]))
                    sel_c = st.selectbox("選擇發送對象", c_list) if c_list else None
                    if sel_c:
                        r_data = next((r for r in records if r.get('企業／贊助單位') == sel_c), {})
                        to_mail = extract_email(r_data.get('聯絡資訊', ''))
                        if to_mail: st.success(f"目標信箱: `{to_mail}`")
                        
                        pdf_name = f"{str(r_data.get('編號', '000')).zfill(3)}_{sel_c}_贊助企劃書.pdf"
                        pdf_p = os.path.join(pdf_dir, pdf_name)
                        f_dict = dict(r_data)
                        f_dict.update({"team_name": team_name, "contact_person": contact_person, "contact_phone": contact_phone, "pdf_filename": pdf_name})
                        
                        raw_t = safe_format_template(email_template, f_dict)
                        # 🌟 新增：產生自訂主旨
                        formatted_subj = safe_format_template(email_subject, f_dict)
                        
                        st.write(f"**✉️ 預覽主旨：** {formatted_subj}")
                        st.write("--- 預覽內容 ---")
                        st.text(raw_t)
                        st.write("-------------")
                        
                        is_s = sel_c in all_sents
                        if st.button("再次寄送" if is_s else "確定寄送", type="primary"):
                            if not sender_email or not sender_password or not to_mail: st.error("帳號未設或無Email！")
                            else:
                                with st.spinner('寄送中...'):
                                    try:
                                        server = smtplib.SMTP("smtp.gmail.com", 587, timeout=10)
                                        server.starttls()
                                        server.login(sender_email, sender_password)
                                        
                                        msg = MIMEMultipart()
                                        msg['From'] = sender_email
                                        msg['To'] = to_mail
                                        msg['Subject'] = formatted_subj  # 使用自訂主旨
                                        
                                        msg.attach(MIMEText(raw_t, 'plain', 'utf-8'))
                                        
                                        if os.path.exists(pdf_p):
                                            with open(pdf_p, 'rb') as f:
                                                att = MIMEApplication(f.read(), _subtype="pdf")
                                                att.add_header('Content-Disposition', 'attachment', filename=pdf_name)
                                                msg.attach(att)
                                        
                                        server.send_message(msg)
                                        server.quit()
                                        
                                        try:
                                            imap = imaplib.IMAP4_SSL("imap.gmail.com", timeout=10)
                                            imap.login(sender_email, sender_password)
                                            if imap.select(backup_folder)[0] != 'OK': imap.create(backup_folder)
                                            imap.append(backup_folder, '\\Seen', imaplib.Time2Internaldate(time.time()), msg.as_bytes())
                                            imap.logout()
                                        except: pass
                                        
                                        if not is_s:
                                            projects_db[curr_proj]["sent_companies"].append({"company": sel_c, "sender": st.session_state.real_name, "email": to_mail})
                                            save_projects(projects_db)
                                        st.success("✅ 成功！")
                                        time.sleep(1)
                                        st.rerun()
                                    except Exception as e: st.error(f"失敗：{e}")