import streamlit as st
import fitz
import re
import json
import requests
import io
import os
import subprocess
import tempfile
from docx import Document
from docx.shared import Inches, Pt
from PIL import Image
from crewai import Agent, Task, Crew, LLM
from concurrent.futures import ThreadPoolExecutor, as_completed
from streamlit_agraph import agraph, Node, Edge, Config
from openai import OpenAI

# ==========================================
# 1. 系统初始化与专业 UI 注入
# ==========================================
st.set_page_config(
    page_title="Sci-Mentor 科研辅助系统", 
    layout="wide",
    initial_sidebar_state="expanded" # 默认展开侧边栏
)

# 隐藏默认菜单和右上角工具栏，保留极简学术风
st.markdown("""
    <style>
    [data-testid="stToolbar"] { display: none !important; }
    #MainMenu { display: none !important; }
    footer { display: none !important; }
    [data-testid="stStatusWidget"] { display: none !important; } /* 隐藏右下角悬浮头像 */
    
    .stApp { 
        background-color: #F7F9FC; 
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
    }
    
    h1, h2, h3, h4 { color: #1F3A5F; font-weight: 600; }
    p, li { color: #333333; line-height: 1.6; }
    
    [data-testid="stSidebar"] {
        background-color: #FFFFFF;
        border-right: 1px solid #E5E7EB;
    }
    
    .stTabs [data-baseweb="tab-list"] { 
        gap: 32px; 
        border-bottom: 1px solid #E5E7EB;
        padding-bottom: 8px;
    }
    .stTabs [data-baseweb="tab"] { 
        font-size: 16px; 
        font-weight: 500; 
        color: #6B7280;
    }
    .stTabs [aria-selected="true"] { 
        color: #1F3A5F !important; 
        border-bottom: 2px solid #1F3A5F !important;
    }
    
    /* 修复按钮文字消失问题 */
    .stButton>button {
        border-radius: 4px;
        background-color: #1F3A5F !important;
        color: #FFFFFF !important;
        font-weight: 500;
        border: none;
        padding: 0.4rem 1.5rem;
        width: 100%;
    }
    .stButton>button:hover {
        background-color: #152A46 !important;
        color: #FFFFFF !important;
    }
    .stButton>button p, .stButton>button div {
        color: #FFFFFF !important;
    }
    
    /* 统一 Slider 颜色为学术蓝 */
    .stSlider [data-baseweb="slider"] div[role="slider"] {
        background-color: #1F3A5F !important;
        border-color: #1F3A5F !important;
    }
    .stSlider [data-baseweb="slider"] > div:nth-child(2) {
        background: #1F3A5F !important;
    }
    .stSlider [data-testid="stTickBar"] > div {
        background-color: #D1D5DB !important;
    }
    .stSlider div[data-testid="stSliderThumbValue"] {
        color: #1F3A5F !important;
    }

    .stTextInput>div>div>input, .stTextArea>div>div>textarea {
        border-radius: 4px;
        border: 1px solid #D1D5DB;
    }
    
    .workflow-text {
        font-size: 14px;
        color: #4B5563;
        margin-bottom: 20px;
        padding-bottom: 10px;
        border-bottom: 1px dashed #D1D5DB;
    }
    </style>
""", unsafe_allow_html=True)

st.title("Sci-Mentor 智能科研辅助系统")
st.markdown("""
<div class='workflow-text'>
    <b>科研工作流：</b> &nbsp; 1. 宏观检索与寻宝 &nbsp;→&nbsp; 2. 本地文献微观精读 &nbsp;→&nbsp; 3. 课题方案生成 &nbsp;→&nbsp; 4. 工业级排版翻译
</div>
""", unsafe_allow_html=True)

# ==========================================
# 2. 侧边栏配置 (从 Secrets 读取)
# ==========================================
with st.sidebar:
    st.subheader("系统环境配置")
    
    # 尝试从 st.secrets 读取默认值，如果在本地跑没有 secrets 文件，则默认为空
    try:
        default_ds_key = st.secrets["DEEPSEEK_API_KEY"]
        default_am_token = st.secrets["AMINER_TOKEN"]
    except:
        default_ds_key = ""
        default_am_token = ""

    deepseek_key = st.text_input("DeepSeek API Key", value=default_ds_key, type="password")
    aminer_token = st.text_input("AMiner Token", value=default_am_token, type="password")
    limit_num = st.slider("单次检索文献上限", 1, 10, 3)
    st.caption("注：云端已预设测试凭证，可直接运行。若失效请在左侧填入您的 Key。")

def init_llm(key):
    return LLM(model="deepseek/deepseek-chat", base_url="https://api.deepseek.com", api_key=key, temperature=0.1)

def draw_graph(data_list):
    nodes, edges, seen = [], [], set()
    for idx, raw in enumerate(data_list):
        try:
            d = json.loads(raw.replace("```json", "").replace("```", "").strip())
            p_name = d.get("文献名称", f"Paper_{idx}")
            target = d.get("关键靶点") or d.get("研究靶点") or "Unknown"
            limit = d.get("局限或临床痛点") or d.get("局限或痛点") or "None"
            method = d.get("干预手段", "")

            if p_name not in seen:
                nodes.append(Node(id=p_name, label=p_name[:12], size=22, color="#1F3A5F"))
                seen.add(p_name)
            if target and target not in ["未提及", "Unknown", "无"]:
                if target not in seen:
                    nodes.append(Node(id=target, label=target, size=18, color="#3B82F6"))
                    seen.add(target)
                edges.append(Edge(source=p_name, target=target, label="靶点", color="#9CA3AF"))
            if method and method not in ["未提及", "Unknown", "无"]:
                if method not in seen:
                    nodes.append(Node(id=method, label=method[:12], size=16, color="#8B5CF6"))
                    seen.add(method)
                edges.append(Edge(source=p_name, target=method, label="干预", color="#9CA3AF"))
            if limit and limit not in ["未提及", "None"]:
                l_id = f"lim_{limit[:8]}"
                if l_id not in seen:
                    nodes.append(Node(id=l_id, label=limit[:12]+"...", size=16, color="#EF4444"))
                    seen.add(l_id)
                edges.append(Edge(source=p_name, target=l_id, label="局限", color="#9CA3AF"))
        except:
            continue
    if nodes:
        st.markdown("##### 知识结构关系图谱")
        agraph(nodes=nodes, edges=edges, config=Config(width="100%", height=400, directed=True, physics=True))

# ==========================================
# 3. 主界面 Tabs
# ==========================================
t2, t1, t4, t3 = st.tabs(["1. 宏观检索与寻宝", "2. 本地文献精读", "3. 课题方案生成", "4. 排版翻译"])

# --- Tab 2: 在线挖掘 ---
with t2:
    st.subheader("双源学术数据检索 (AMiner + PubMed)")
    
    c1, c2, c3 = st.columns([3, 2, 2])
    with c1:
        query = st.text_input("目标检索词", "tumor microenvironment")
    with c2:
        journal = st.text_input("期刊名称过滤 (可选)", "")
    with c3:
        logic_opt = st.selectbox("逻辑组合方式", ["AND", "OR", "NOT", "NAND"])
    
    st.markdown("<br>", unsafe_allow_html=True) 
    use_oa = st.checkbox("解析 PubMed OA 全文数据 (耗时较长)", value=False)

    if st.button("执行在线检索与分析", key="btn_online"):
        if not deepseek_key or not aminer_token:
            st.error("凭证缺失，请在左侧侧边栏配置。")
        else:
            llm = init_llm(deepseek_key)
            raw_data = []
            
            with st.spinner("正在查询数据库并交叉比对..."):
                try:
                    # 抓取 AMiner
                    a_url = "https://datacenter.aminer.cn/gateway/open_platform/api/paper/search/pro"
                    a_res = requests.get(a_url, params={"page": 0, "size": max(limit_num*2, 5), "title": query}, headers={"Authorization": aminer_token}).json()
                    for p in a_res.get("data", []):
                        abs_text = re.sub(r'<[^>]+>', ' ', p.get('abstract', '')).strip()
                        if len(abs_text) > 20:
                            raw_data.append({"src": "AMiner", "title": p.get('title', ''), "venue": p.get('venue_name', ''), "doi": p.get('doi', ''), "id": p.get('id', ''), "abs": abs_text})
                    
                    # 抓取 PubMed
                    p_url = f"https://www.ebi.ac.uk/europepmc/webservices/rest/search?query={query}&format=json&pageSize={max(limit_num*2, 5)}&resultType=core"
                    p_res = requests.get(p_url).json()
                    for p in p_res.get("resultList", {}).get("result", []):
                        abs_text = re.sub(r'<[^>]+>', ' ', p.get('abstractText', '')).strip()
                        if len(abs_text) > 50:
                            pmcid = p.get('pmcid', '')
                            tag = "PubMed (Abstract)"
                            if use_oa and pmcid:
                                tag = "PubMed (Full Text)"
                                try:
                                    xml = requests.get(f"https://www.ebi.ac.uk/europepmc/webservices/rest/{pmcid}/fullTextXML", timeout=10).text
                                    abs_text = re.sub(r'\s+', ' ', re.sub(r'<[^>]+>', ' ', xml))[:20000] + "..."
                                except:
                                    pass
                            raw_data.append({"src": tag, "title": p.get('title', ''), "venue": p.get('journalTitle', ''), "doi": p.get('doi', ''), "id": p.get('id', ''), "abs": abs_text})
                    
                    # 过滤去重
                    filtered, titles_seen = [], set()
                    for p in raw_data:
                        t_low = p['title'].lower().strip()
                        if t_low in titles_seen or not t_low: continue
                        has_q = query.lower() in t_low or query.lower() in p['abs'].lower()
                        has_j = journal.lower() in p['venue'].lower() if journal else False
                        
                        keep = False
                        if logic_opt == "AND": keep = has_q and has_j
                        elif logic_opt == "OR": keep = has_q or has_j
                        elif logic_opt == "NOT": keep = has_q and not has_j
                        elif logic_opt == "NAND": keep = not has_q and not has_j
                        
                        if keep:
                            filtered.append(p)
                            titles_seen.add(t_low)
                    
                    final_papers = filtered[:limit_num]
                    if not final_papers:
                        st.warning("未检索到符合条件的学术文献。")
                        st.stop()

                    st.info(f"检索完成，已提取 {len(final_papers)} 篇核心文献。")
                    
                    paper_texts = []
                    for idx, p in enumerate(final_papers):
                        link = f"https://doi.org/{p['doi']}" if p['doi'] else (f"https://www.aminer.cn/pub/{p['id']}" if p['src']=="AMiner" else f"https://europepmc.org/article/MED/{p['id']}")
                        st.markdown(f"**[{idx+1}]** [{p['title']}]({link}) - *{p['venue']}* `| {p['src']}`")
                        paper_texts.append(f"来源:{p['src']}\n标题:{p['title']}\n期刊:{p['venue']}\n链接:{link}\n内容:{p['abs']}")

                    def extract_entities(info):
                        title, txt = info
                        ag = Agent(role='分析助手', goal='提取结构化数据', backstory='专业医学生', llm=llm)
                        task = Task(description=f'输出 JSON: "文献名称": "{title[:15]}", "研究靶点": "...", "局限或痛点": "..."\n文本:{txt}', expected_output='JSON', agent=ag)
                        return str(Crew(agents=[ag], tasks=[task]).kickoff(inputs={})).replace("```json", "").replace("```", "").strip()

                    pairs = [(p.split('\n')[1].replace('标题:',''), p) for p in paper_texts]
                    entities = [None] * len(pairs)
                    with ThreadPoolExecutor(max_workers=5) as ex:
                        futs = {ex.submit(extract_entities, pr): i for i, pr in enumerate(pairs)}
                        for f in as_completed(futs):
                            entities[futs[f]] = f.result()

                    draw_graph(entities)

                    sum_agent = Agent(role='科研导师', goal='提炼研究空白', backstory='资深学者', llm=llm)
                    sum_task = Task(description=f'输出2个研究空白和2个创新方向，结论附链接。末尾加“关联文献导航”。\n数据:{"".join(entities)}', expected_output='报告', agent=sum_agent)
                    
                    with st.spinner("正在归纳总结..."):
                        report = str(Crew(agents=[sum_agent], tasks=[sum_task]).kickoff(inputs={})).replace("```markdown", "").replace("```", "")
                    
                    st.markdown("### 宏观调研报告")
                    st.markdown(report)
                    st.session_state['gap_1'] = report
                    st.download_button("导出报告", data=report, file_name="macro_report.md")

                except Exception as e:
                    st.error(f"检索异常: {e}")

# --- Tab 1: 本地对比 ---
with t1:
    st.subheader("多文献交叉比对与机制提取")
    files = st.file_uploader("上传至少两篇 PDF 文献", accept_multiple_files=True, type=["pdf"])
    
    if st.button("执行交叉比对", key="btn_local"):
        if not deepseek_key or len(files) < 2:
            st.error("需配置 Key 并上传至少两篇文献")
        else:
            llm = init_llm(deepseek_key)
            texts = []
            for f in files:
                doc = fitz.open(stream=f.read(), filetype="pdf")
                text = re.sub(r'\n{3,}', '\n\n', "\n".join([page.get_text() for page in doc]))
                texts.append((f.name, text[:12000]))
            
            st.info(f"已加载 {len(texts)} 篇文献，数据提取中...")

            def extract_pdf(file_info):
                name, txt = file_info
                ag = Agent(role='数据提取', goal='提取7个维度', backstory='审稿人', llm=llm)
                t = Task(description=f'输出 JSON: "文献名称":"{name}","研究类型/模型","样本量/对象","关键靶点/分子","干预手段","核心机制","主要结论","局限或临床痛点","原文证据"\n文本:{txt}', expected_output='JSON', agent=ag)
                return str(Crew(agents=[ag], tasks=[t]).kickoff(inputs={})).replace("```json", "").replace("```", "").strip()

            data_list = [None] * len(texts)
            with ThreadPoolExecutor(max_workers=5) as ex:
                futs = {ex.submit(extract_pdf, info): i for i, info in enumerate(texts)}
                for f in as_completed(futs):
                    data_list[futs[f]] = f.result()

            draw_graph(data_list)

            ed_agent = Agent(role='主编', goal='生成对比表', backstory='编辑', llm=llm)
            ed_task = Task(description=f'输出 Markdown 表格：【文献名称】|【模型】|【样本量】|【靶点】|【干预】|【机制】|【结论】|【痛点】|【证据】。并输出2个Gap和2个方向，附来源链接。\n数据:{"".join(data_list)}', expected_output='表+报告', agent=ed_agent)
            
            with st.spinner("生成比对报告..."):
                res = str(Crew(agents=[ed_agent], tasks=[ed_task]).kickoff(inputs={})).replace("```markdown", "").replace("```", "")
            
            st.markdown("### 文献对比与研究空白")
            st.markdown(res)
            st.session_state['gap_2'] = res
            st.download_button("导出对比报告", data=res, file_name="micro_comparison.md")

# --- Tab 4: 方案生成 ---
with t4:
    st.subheader("课题研究方案智能生成")
    src_list = ["手动输入"]
    if 'gap_1' in st.session_state: src_list.insert(0, "宏观调研报告")
    if 'gap_2' in st.session_state: src_list.insert(0, "本地对比报告")
    if 'gap_1' in st.session_state and 'gap_2' in st.session_state: src_list.insert(0, "融合上述两份报告")

    src = st.radio("数据输入来源", src_list, horizontal=True)
    
    default_txt = ""
    if src == "融合上述两份报告": default_txt = f"【宏观】\n{st.session_state['gap_1']}\n\n【微观】\n{st.session_state['gap_2']}"
    elif src == "宏观调研报告": default_txt = st.session_state['gap_1']
    elif src == "本地对比报告": default_txt = st.session_state['gap_2']
    else: default_txt = "针对肿瘤免疫逃逸，缺乏有效靶向干预。"

    topic = st.text_area("课题背景与痛点确认", default_txt, height=180)
    
    if st.button("编译研究方案", key="btn_prop"):
        if not deepseek_key or not topic.strip():
            st.error("请确认 Key 和课题信息")
        else:
            llm = init_llm(deepseek_key)
            with st.spinner("构建研究框架与实验设计..."):
                a1 = Agent(role='生信专家', goal='规划数据库分析', backstory='多组学专家', llm=llm)
                a2 = Agent(role='实验专家', goal='设计分子与动物实验', backstory='资深研究员', llm=llm)
                a3 = Agent(role='首席PI', goal='统筹完整方案', backstory='评审专家', llm=llm)
                
                t1 = Task(description=f'规划公共数据库(GEO/TCGA)挖掘策略:\n{topic}', expected_output='生信方案', agent=a1)
                t2 = Task(description=f'设计体内外实验与机制验证:\n{topic}', expected_output='实验方案', agent=a2)
                t3 = Task(description=f'''
                综合上述规划，输出正式《研究方案》。必须用 Markdown 分级标题排版，包含：
                一、基本信息与假说
                二、数据库与数据需求 (明确数据库名与数据类型)
                三、实验设计 (细胞模型、动物模型、机制实验)
                四、数据整理与统计分析 (软件与检验方法)
                五、预期成果
                ''', expected_output='完整方案', agent=a3, context=[t1, t2])

                res = str(Crew(agents=[a1, a2, a3], tasks=[t1, t2, t3]).kickoff(inputs={})).replace("```markdown", "").replace("```", "")
                
            st.markdown("### 科研项目计划书")
            st.markdown(res)
            st.download_button("导出研究方案", data=res, file_name="proposal.md", type="primary")

# --- Tab 3: 翻译 ---
with t3:
    st.subheader("学术文献原位排版翻译")
    f_trans = st.file_uploader("上传英文文献 PDF", type=["pdf"])
    
    if st.button("执行翻译", key="btn_trans"):
        if not deepseek_key or not f_trans:
            st.error("请确认 Key 和文件")
        else:
            with tempfile.TemporaryDirectory() as tmpdir:
                in_path = os.path.join(tmpdir, "in.pdf")
                with open(in_path, "wb") as f: f.write(f_trans.read())
                
                st.info("正在处理 PDF 排版与文本翻译...")
                env = os.environ.copy()
                env["OPENAI_BASE_URL"] = "https://api.deepseek.com/v1"
                env["OPENAI_API_KEY"] = deepseek_key
                
                try:
                    subprocess.run(["pdf2zh", in_path, "-s", "openai:deepseek-chat", "-o", tmpdir], capture_output=True, text=True, env=env, timeout=600)
                    
                    out_files = os.listdir(tmpdir)
                    zh_path, dual_path = None, None
                    for fn in out_files:
                        if fn.endswith("-zh.pdf"): zh_path = os.path.join(tmpdir, fn)
                        elif fn.endswith("-dual.pdf"): dual_path = os.path.join(tmpdir, fn)
                    if not zh_path and not dual_path:
                        for fn in out_files:
                            if fn.endswith(".pdf") and fn != "in.pdf": dual_path = os.path.join(tmpdir, fn)

                    if zh_path or dual_path:
                        st.success("翻译处理完成")
                        c1, c2 = st.columns(2)
                        if zh_path: c1.download_button("下载纯中文版", data=open(zh_path, "rb").read(), file_name="Translated_ZH.pdf")
                        if dual_path: c2.download_button("下载双语对照版", data=open(dual_path, "rb").read(), file_name="Translated_Dual.pdf")
                    else:
                        st.error("未找到翻译输出文件")
                except Exception as e:
                    st.error(f"执行超时或失败: {e}")