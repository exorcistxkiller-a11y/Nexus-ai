import streamlit as st
import psutil
import requests
import ollama
import os
import io
import sys
import contextlib
from PIL import Image
import PyPDF2
from pptx import Presentation

# ---------------------------------------------------------
# Page Setup 
# ---------------------------------------------------------
st.set_page_config(page_title="Nexus AI — Refined Interface", page_icon="✨", layout="wide", initial_sidebar_state="expanded")

# ---------------------------------------------------------
# 1:1 CSS Extraction from Provided HTML
# ---------------------------------------------------------
st.markdown("""
<style>
    .stApp { 
        background: radial-gradient(850px 550px at 72% -10%,rgba(143,124,255,.17),transparent 60%),
                    radial-gradient(650px 480px at 10% 105%,rgba(121,230,255,.10),transparent 60%),
                    #080b16 !important;
        color: #eef3ff !important; 
        font-family: Inter, ui-sans-serif, system-ui, sans-serif !important;
    }
    header[data-testid="stHeader"] { display: none !important; }
    
    section[data-testid="stSidebar"] { 
        background-color: rgba(16,21,34,.94) !important; 
        border-right: 1px solid #273148 !important;
        min-width: 268px !important;
    }
    
    .topbar {
        display: flex; align-items: center; justify-content: space-between; 
        padding: 15px 27px; border-bottom: 1px solid #273148; 
        background: rgba(10,14,24,.72); backdrop-filter: blur(12px);
        margin: -90px -40px 20px -40px; 
        z-index: 999;
    }
    .topbar-title { display: flex; align-items: center; gap: 10px; }
    .topbar-title h1 { font-size: 18px; margin: 0; font-weight: 750; color: #eef3ff; line-height: 1; }
    .topbar-state { font-size: 12px; color: #8994ad; padding-left: 2px; }
    .online-dot { width: 7px; height: 7px; border-radius: 50%; background: #57e39b; box-shadow: 0 0 9px rgba(87,227,155,.55); }
    .deploy-btn-group { display: flex; gap: 9px; }
    .top-btn { height: 36px; padding: 0 16px; border-radius: 9px; border: 1px solid #273148; background: #101522; color: #eef3ff; font-size: 13px; font-weight: 650; cursor: pointer; }
    .top-btn.deploy { border: 0; background: linear-gradient(120deg,#8f7cff,#79e6ff); color: #07101b; }

    .brand { display: flex; align-items: center; gap: 10px; font-size: 20px; font-weight: 750; letter-spacing: .2px; margin-bottom: 22px; color:#eef3ff;}
    .logo { width: 31px; height: 31px; border-radius: 9px; background: linear-gradient(135deg,#8f7cff,#79e6ff); display: grid; place-items: center; color: #07101b; box-shadow: 0 0 25px rgba(121,230,255,.12); }
    .logo svg { width: 20px; height: 20px; }

    div[data-testid="stSidebar"] .stButton > button {
        width: 100% !important; height: 43px !important;
        border: 1px solid #273148 !important; border-radius: 11px !important;
        background: #151b2b !important; color: #eef3ff !important;
        justify-content: flex-start !important; padding: 0 13px !important;
        font-weight: 600 !important; font-size: 13.5px !important;
        transition: background 0.2s, border-color 0.2s;
    }
    div[data-testid="stSidebar"] .stButton > button:hover { border-color: #52627f !important; background: #182035 !important; }
    div[data-testid="stSidebar"] .stButton:first-of-type > button {
        background: linear-gradient(120deg,#8f7cff,#79e6ff) !important;
        border-color: transparent !important; color: #07101b !important;
        box-shadow: 0 8px 25px rgba(121,230,255,.10) !important;
    }

    .status { margin-top: 18px; border: 1px solid #273148; border-radius: 14px; background: rgba(21,27,43,.72); padding: 14px 13px; }
    .status-title { display: flex; align-items: center; gap: 8px; color: #8994ad; font-size: 12px; font-weight: 700; text-transform: uppercase; letter-spacing: .08em; margin-bottom: 14px; }
    .metric { margin: 12px 0; }
    .metric-head { display: flex; justify-content: space-between; font-size: 12.5px; margin-bottom: 6px; color: #eef3ff; }
    .metric-head span { color: #8994ad; }
    .track { height: 5px; background: #252e43; border-radius: 20px; overflow: hidden; }
    .history { margin-top: 20px; }
    .history-title { font-size: 12px; color: #8994ad; font-weight: 700; text-transform: uppercase; letter-spacing: .08em; margin: 0 6px 8px; }
    .history a { display: block; color: #cfd7e9; font-size: 13px; padding: 9px 10px; border-radius: 9px; cursor:pointer; }
    .history a:hover { background: #192136; color: white; }

    [data-testid="stChatMessage"] {
        padding: 13px 16px !important; border: 1px solid #273148 !important;
        border-radius: 15px !important; background: rgba(16,21,34,.92) !important;
        font-size: 14.5px !important; line-height: 1.58 !important; color: #e8edf8 !important;
        box-shadow: 0 8px 24px rgba(0,0,0,.08) !important; max-width: 800px; margin-bottom: 20px;
    }
    [data-testid="stChatMessage"]:nth-child(odd) {
        background: linear-gradient(135deg, rgba(143,124,255,.19), rgba(121,230,255,.08)) !important;
        border-color: rgba(143,124,255,.65) !important;
    }

    .stChatInputContainer { 
        max-width: 850px !important; margin: 0 auto !important;
        border: 1px solid #2c3852 !important; background: rgba(16,21,34,.96) !important;
        border-radius: 18px !important; padding: 11px 12px 10px !important;
        box-shadow: 0 16px 45px rgba(0,0,0,.25) !important;
    }
    .stChatInputContainer:focus-within { border-color: rgba(143,124,255,.85) !important; box-shadow: 0 0 0 1px rgba(143,124,255,.18),0 16px 45px rgba(0,0,0,.3) !important; }

    [data-testid="stSelectbox"] div[data-baseweb="select"] {
        height: 37px !important; min-height: 37px !important;
        border: 1px solid #273148 !important; border-radius: 10px !important;
        background: #151b2b !important; color: #eaf1ff !important;
        font-size: 12.5px !important; cursor: pointer;
    }
    [data-testid="stFileUploaderDropzone"] {
        border: 1px dashed #273148 !important; border-radius: 10px !important;
        background: #151b2b !important; padding: 10px !important;
    }
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------
# Telemetry Helpers & Core Tools
# ---------------------------------------------------------
def get_loaded_models():
    try:
        res = requests.get("http://localhost:11434/api/ps", timeout=1.5)
        if res.status_code == 200: return res.json().get("models", [])
    except: pass
    return []

def force_unload_all():
    models = get_loaded_models()
    for m in models:
        try: requests.post("http://localhost:11434/api/generate", json={"model": m.get("name"), "keep_alive": 0}, timeout=2.0)
        except: pass

def extract_text_from_multiple_files(uploaded_files):
    combined_text = ""
    for uploaded_file in uploaded_files:
        file_type = uploaded_file.name.split('.')[-1].lower()
        if file_type == 'txt':
            text = uploaded_file.getvalue().decode("utf-8")
            combined_text += f"\n--- File: {uploaded_file.name} ---\n" + text + "\n"
        elif file_type == 'pdf':
            reader = PyPDF2.PdfReader(uploaded_file)
            pdf_text = ""
            for page in reader.pages:
                t = page.extract_text()
                if t: pdf_text += t + "\n"
            combined_text += f"\n--- File: {uploaded_file.name} ---\n" + pdf_text + "\n"
    return combined_text

def create_presentation(title_text, points):
    prs = Presentation()
    slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(slide_layout)
    slide.shapes.title.text = title_text
    slide.placeholders[1].text = "Generated autonomously by Nexus AI Core"
    
    slide_layout = prs.slide_layouts[1]
    slide = prs.slides.add_slide(slide_layout)
    slide.shapes.title.text = "Key Insights & Overview"
    tf = slide.placeholders[1].text_frame
    tf.text = points[0] if points else "Core concept generated by local model."
    
    for pt in points[1:]:
        p = tf.add_paragraph()
        p.text = pt
        p.level = 0

    file_stream = io.BytesIO()
    prs.save(file_stream)
    file_stream.seek(0)
    return file_stream

# ---------------------------------------------------------
# Session State & Personas Definition
# ---------------------------------------------------------
available_models = ["qwen2.5-coder:7b-instruct-q4_K_M", "llava:7b", "llama3.1", "deepseek-r1:1.5b"]

personas = {
    "💻 Code Wizard": "You are an expert software engineer and programming tutor. Provide clean, efficient code and crisp technical troubleshooting.",
    "🧮 Logic & Math Professor": "You are an advanced mathematical logic professor. Break down problems step-by-step with clear reasoning and verification.",
    "✨ General Assistant": "You are Nexus AI, a helpful, precise, and balanced personal assistant."
}

if "active_model" not in st.session_state:
    st.session_state.active_model = available_models[0]
if "last_file_names" not in st.session_state:
    st.session_state.last_file_names = []
if "messages" not in st.session_state:
    st.session_state.messages = [{"role": "assistant", "content": "Nexus Core initialized. All systems ready."}]

# ---------------------------------------------------------
# Sidebar Layout (Model Selector, Uploader, Personas, Controls)
# ---------------------------------------------------------
with st.sidebar:
    st.markdown("""
    <div class="brand">
      <div class="logo"><svg viewBox="0 0 24 24" fill="none"><path d="M12 2.5 20 7v10l-8 4.5L4 17V7l8-4.5Z" stroke="currentColor" stroke-width="1.6"/><circle cx="12" cy="12" r="2.1" fill="currentColor"/></svg></div>
      <span>Nexus AI</span>
    </div>
    """, unsafe_allow_html=True)
    
    if st.button("＋ New chat", use_container_width=True):
        st.session_state.messages = []
        st.rerun()
    if st.button("⌫ Purge memory", use_container_width=True):
        force_unload_all()
        st.rerun()

    # --- MOVED TO SIDEBAR: Model Changer ---
    st.markdown("<div class='history-title' style='margin-top:15px;'>Active Model Engine</div>", unsafe_allow_html=True)
    current_index = available_models.index(st.session_state.active_model) if st.session_state.active_model in available_models else 0
    new_selected_model = st.selectbox("Model Selector", available_models, index=current_index, label_visibility="collapsed")

    # --- MOVED TO SIDEBAR: File Uploader ---
    st.markdown("<div class='history-title' style='margin-top:15px;'>Document / Vision Upload</div>", unsafe_allow_html=True)
    uploaded_files = st.file_uploader("Upload Files", type=['png', 'jpg', 'jpeg', 'pdf', 'txt'], accept_multiple_files=True, label_visibility="collapsed")

    # Persona Selector
    st.markdown("<div class='history-title' style='margin-top:15px;'>Active Persona</div>", unsafe_allow_html=True)
    selected_persona_name = st.selectbox("Persona Selector", list(personas.keys()), label_visibility="collapsed")
    current_system_prompt = personas[selected_persona_name]

    # Chat History Export Button
    chat_export_markdown = ""
    for m in st.session_state.messages:
        role_label = "User" if m["role"] == "user" else "Nexus AI"
        chat_export_markdown += f"**{role_label}:** {m['content']}\n\n---\n"
    
    st.markdown("<div style='margin-top: 15px;'></div>", unsafe_allow_html=True)
    st.download_button(
        label="📥 Export Chat History (.md)",
        data=chat_export_markdown,
        file_name="nexus_chat_history.md",
        mime="text/markdown",
        use_container_width=True
    )

    mem = psutil.virtual_memory()
    mem_used_gb = (mem.total - mem.available) / (1024**3)
    mem_total_gb = mem.total / (1024**3)
    cpu = psutil.cpu_percent(interval=0.1) 
    
    st.markdown(f"""
    <div class="status">
      <div class="status-title"><span class="dot"></span>System status</div>
      <div class="metric"><div class="metric-head"><span>RAM used</span><b>{mem_used_gb:.1f} / {mem_total_gb:.0f} GB</b></div><div class="track"><i style="width:{mem.percent}%;background:#79e6ff;display:block;height:100%;border-radius:20px;"></i></div></div>
      <div class="metric"><div class="metric-head"><span>CPU temp</span><b>N/A</b></div><div class="track"><i style="width:0%;background:#ffc15c;display:block;height:100%;border-radius:20px;"></i></div></div>
      <div class="metric"><div class="metric-head"><span>CPU usage</span><b>{cpu}%</b></div><div class="track"><i style="width:{cpu}%;background:#8f7cff;display:block;height:100%;border-radius:20px;"></i></div></div>
    </div>
    """, unsafe_allow_html=True)

# ---------------------------------------------------------
# Main Workspace & Topbar Header
# ---------------------------------------------------------
st.markdown("""
<header class="topbar">
  <div class="topbar-title"><span class="online-dot"></span><h1>Nexus AI</h1><span class="topbar-state">local model · online</span></div>
  <div class="deploy-btn-group"><button class="top-btn">Stop</button><button class="top-btn deploy">Deploy</button></div>
</header>
""", unsafe_allow_html=True)

for idx, msg in enumerate(st.session_state.messages):
    avatar_icon = "👤" if msg["role"] == "user" else "✨"
    with st.chat_message(msg["role"], avatar=avatar_icon):
        st.markdown(msg["content"])
        if "ppt_data" in msg:
            st.download_button(
                label="📥 Download PowerPoint Presentation (.pptx)",
                data=msg["ppt_data"],
                file_name="Nexus_Presentation.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                key=f"ppt_dl_{idx}"
            )
        
        # Code Sandbox execution output
        if msg["role"] == "assistant" and "```python" in msg["content"]:
            if st.button(f"▶ Run Code Snippet #{idx}", key=f"run_code_{idx}"):
                try:
                    parts = msg["content"].split("```python")
                    code_block = parts[1].split("```")[0].strip()
                    
                    st.markdown("**🛡️ Sandbox Execution Output:**")
                    stdout_buffer = io.StringIO()
                    with contextlib.redirect_stdout(stdout_buffer):
                        exec(code_block, {})
                    output_result = stdout_buffer.getvalue()
                    if output_result.strip():
                        st.code(output_result, language="text")
                    else:
                        st.success("Code executed successfully with no printed output.")
                except Exception as ex:
                    st.error(f"Execution Error: {str(ex)}")

# ---------------------------------------------------------
# SMART ROUTER & LOCAL RAG
# ---------------------------------------------------------
current_file_names = [f.name for f in uploaded_files] if uploaded_files else []
if current_file_names != st.session_state.last_file_names:
    st.session_state.last_file_names = current_file_names
    if uploaded_files:
        file_exts = [f.name.split('.')[-1].lower() for f in uploaded_files]
        if any(ext in ['png', 'jpg', 'jpeg'] for ext in file_exts):
            st.session_state.active_model = "llava:7b"
            st.session_state.messages = []
            st.rerun()
        elif any(ext in ['pdf', 'txt'] for ext in file_exts):
            st.session_state.active_model = "llama3.1"
            st.session_state.messages = []
            st.rerun()
else:
    if not uploaded_files and st.session_state.last_file_names:
        st.session_state.last_file_names = []
        st.session_state.active_model = "qwen2.5-coder:7b-instruct-q4_K_M"
        st.session_state.messages = []
        st.rerun()

if new_selected_model != st.session_state.active_model:
    st.session_state.active_model = new_selected_model
    st.session_state.messages = []
    st.rerun()

selected_model = st.session_state.active_model

# ---------------------------------------------------------
# Generation Loop & Advanced Tool Handlers
# ---------------------------------------------------------
if prompt := st.chat_input("Message Nexus AI (Type 'ppt: <topic>' for slides)..."):
    user_msg = {"role": "user", "content": prompt}
    
    if uploaded_files:
        file_exts = [f.name.split('.')[-1].lower() for f in uploaded_files]
        if any(ext in ['pdf', 'txt'] for ext in file_exts):
            combined_docs = extract_text_from_multiple_files(uploaded_files)
            user_msg["content"] = f"Here are the indexed documents:\n\n{combined_docs}\n\nBased on these documents, answer the following: {prompt}"
        elif any(ext in ['png', 'jpg', 'jpeg'] for ext in file_exts):
            img_file = [f for f in uploaded_files if f.name.split('.')[-1].lower() in ['png', 'jpg', 'jpeg']][0]
            raw_img = Image.open(io.BytesIO(img_file.getvalue()))
            raw_img.thumbnail((1600, 1600)) 
            byte_arr = io.BytesIO()
            raw_img.save(byte_arr, format='JPEG')
            user_msg["images"] = [byte_arr.getvalue()]

    st.session_state.messages.append(user_msg)
    
    with st.chat_message("user", avatar="👤"):
        st.markdown(prompt)

    with st.chat_message("assistant", avatar="✨"):
        response_box = st.empty()
        
        if prompt.lower().startswith("ppt:"):
            topic = prompt[4:].strip()
            response_box.markdown(f"⚙️ **Generating presentation slides for:** *{topic}*...")
            
            outline_prompt = f"Provide 4 concise bullet points explaining the topic: {topic}"
            res = ollama.chat(model="llama3.1", messages=[{"role": "user", "content": outline_prompt}])
            content_text = res["message"]["content"]
            points = [line.strip("-* ").strip() for line in content_text.split("\n") if line.strip()][:4]
            if not points:
                points = ["Introduction and background", "Core framework and features", "Real-world impact", "Future outlook"]

            ppt_stream = create_presentation(topic, points)
            
            full_text = f"✅ PowerPoint presentation successfully generated for **{topic}**!\n\nClick the button below to download your `.pptx` file."
            response_box.markdown(full_text)
            st.session_state.messages.append({"role": "assistant", "content": full_text, "ppt_data": ppt_stream.getvalue()})
            st.rerun()

        else:
            full_text = ""
            model_options = {}
            if "llava" in selected_model:
                model_options = {"temperature": 0.1, "repeat_penalty": 1.5} 
            elif "deepseek" in selected_model:
                model_options = {"temperature": 0.6} 
            else:
                model_options = {"temperature": 0.3} 

            api_messages = [{"role": "system", "content": current_system_prompt}] + st.session_state.messages

            try:
                stream = ollama.chat(
                    model=selected_model,
                    messages=api_messages,
                    stream=True,
                    keep_alive="1m",
                    options=model_options
                )
                for chunk in stream:
                    text_chunk = chunk["message"]["content"]
                    text_chunk = text_chunk.replace("<think>", "🧠 **Thinking Process:**\n```text\n")
                    text_chunk = text_chunk.replace("</think>", "\n```\n\n---\n**Final Answer:**\n\n")
                    
                    full_text += text_chunk
                    response_box.markdown(full_text + " ▍")
                
                response_box.markdown(full_text)
                st.session_state.messages.append({"role": "assistant", "content": full_text})
                st.rerun()
            except Exception as e:
                st.error("Error connecting to Ollama.")
                st.code(str(e))