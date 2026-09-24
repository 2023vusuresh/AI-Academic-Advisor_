
import os, re, time, math, hashlib
from pathlib import Path
import pandas as pd
import streamlit as st
import openpyxl

# ============================================================
# AIRA — AI Academic Advisor | Assignment #1
# Authoritative sources: the two supplied Excel workbooks.
# RAG: Excel extraction -> normalization -> question routing/entity resolution ->
# 500/100 chunks ->
# SentenceTransformer embeddings -> FAISS retrieval -> rerank ->
# structured verification -> guarded LLM explanation.
# ============================================================

st.set_page_config(page_title="AIRA | Academic Advisor", page_icon="🎓", layout="wide")

ROOT = Path(__file__).resolve().parent
DATA = ROOT / "data"
MINOR_FILE = DATA / "Minor_Courses_for_BTech_Students(3).xlsx"
STRUCT_FILE = DATA / "Semester_Spread_Structures_Sept_2026(3).xlsx"

PASS_GRADES = {"A+","A","A-","B+","B","B-","C+","C","C-","D+","D","D-"}
FAIL_GRADES = {"F","FAIL","FAILED"}

# ---------- styling ----------
st.markdown("""
<style>
:root { --ink:#132238; --muted:#667085; --line:#e5eaf0; --accent:#24527a; --soft:#f6f9fc; }
.block-container{max-width:1120px;padding-top:1rem;padding-bottom:3rem}
.hero{padding:30px 32px;border:1px solid var(--line);border-radius:26px;background:linear-gradient(135deg,#f8fbff 0%,#eef5fb 100%);margin-bottom:20px}
.hero h1{margin:0;color:var(--ink);font-size:2.25rem;letter-spacing:-.02em}
.hero p{color:var(--muted);margin:.45rem 0 0;font-size:1.02rem}
.hero-note{margin-top:14px;color:#526173;font-size:.9rem}
.badge{display:inline-block;padding:5px 10px;border-radius:999px;background:#eaf2f8;color:#24527a;font-size:.78rem;font-weight:700}
.answer{padding:24px 26px;border:1px solid var(--line);border-radius:22px;background:#fff;box-shadow:0 10px 30px rgba(20,33,61,.06);margin-top:12px}
.answer h3{margin:.5rem 0 1rem;color:var(--ink)}
.small{font-size:.86rem;color:#667085}
.source-card{padding:12px 14px;border:1px solid var(--line);border-radius:14px;background:#fafbfd;margin:8px 0}
.source-label{font-weight:700;color:#24527a;font-size:.86rem}
.source-detail{color:#5e6b7a;font-size:.84rem;margin-top:3px}
.question-card{padding:16px;border:1px solid var(--line);border-radius:18px;background:#fff;min-height:112px}
.question-card h4{margin:0 0 5px;color:var(--ink)}
.question-card p{margin:0;color:var(--muted);font-size:.9rem}
.stButton>button{border-radius:12px;border:1px solid #d8e0e8}
</style>
""", unsafe_allow_html=True)

# ---------- data cleaning & organization ----------
def norm(x):
    if x is None or (isinstance(x,float) and pd.isna(x)): return ""
    return re.sub(r"\s+"," ",str(x).replace("\xa0"," ").replace("’","'").strip())

def norm_code(x):
    return re.sub(r"\s+","",norm(x).upper())

def clean_semester(x):
    s=norm(x)
    if not s: return ""
    m=re.search(r"([1-8])",s)
    return f"S{m.group(1)}" if m else s

def clean_prerequisite(x):
    raw=norm(x)
    if not raw: return "", "not_recorded"
    u=raw.upper()
    if u in {"NIL","NONE","NO PREREQUISITE","NO PREREQUISITES"}:
        return "Nil", "explicit_none"
    if u in {"TBD","TO BE DECIDED","TO BE CONFIRMED"}:
        return raw, "pending"
    raw=re.sub(r"\s*[,;]\s*", ", ", raw)
    raw=re.sub(r"\s*/\s*", "/", raw)
    return raw, "recorded"

def batch_marker(x):
    m=re.match(r"^(20\d{2})\s*Batch$",norm(x),re.I)
    return m.group(1) if m else ""

@st.cache_data(show_spinner=False)
def load_authoritative_data():
    """Clean and organize the two supplied Excel workbooks before retrieval."""
    if not MINOR_FILE.exists() or not STRUCT_FILE.exists():
        raise FileNotFoundError("The two authoritative Excel source files are missing from the data folder.")

    # ========================================================
    # 1) MINOR WORKBOOK — header-driven parsing
    # ========================================================
    minors=[]; placeholders=[]
    wb=openpyxl.load_workbook(MINOR_FILE,data_only=True,read_only=True)
    for ws in wb.worksheets:
        vals=list(ws.iter_rows(values_only=True))
        header_idx=None
        for i,row in enumerate(vals[:10]):
            hs=[norm(x).lower() for x in row]
            if "course code" in hs and any("course title" in x for x in hs):
                header_idx=i; break
        if header_idx is None: continue
        hs=[norm(x).lower() for x in vals[header_idx]]
        def find_col(*names):
            for name in names:
                for j,h in enumerate(hs):
                    if h==name.lower(): return j
            for j,h in enumerate(hs):
                if any(name.lower() in h for name in names): return j
            return None
        c={"code":find_col("course code"),"title":find_col("course title"),"l":find_col("l"),"t":find_col("t"),"p":find_col("p"),"credit":find_col("credit"),"semester":find_col("semester"),"pre":find_col("pre-rq","pre-req","prerequisite")}
        batch=""
        for rnum,row in enumerate(vals,start=1):
            for j in range(min(2,len(row))):
                b=batch_marker(row[j])
                if b: batch=b
            if rnum<=header_idx+1: continue
            code_raw=norm(row[c["code"]]) if c["code"] is not None else ""
            title_raw=norm(row[c["title"]]) if c["title"] is not None else ""
            credit=row[c["credit"]] if c["credit"] is not None else None
            sem=clean_semester(row[c["semester"]]) if c["semester"] is not None else ""
            pre_raw=norm(row[c["pre"]]) if c["pre"] is not None else ""
            if title_raw and title_raw.upper() in {"TBD","DON'T KNOW"} and not code_raw:
                placeholders.append({"source_workbook":MINOR_FILE.name,"source_sheet":ws.title,"source_row":rnum,"batch":batch,"placeholder_type":title_raw.upper(),"raw_title":title_raw,"raw_credits":credit,"raw_prerequisite":pre_raw})
                continue
            if not code_raw and not title_raw: continue
            if title_raw.lower() in {"finance minor","marketing minor","start -up minor","start-up minor"}: continue
            cu=code_raw.upper()
            if cu in {"DON'T KNOW","DONT KNOW"}: code_status="not_recorded_in_source"; code=""
            elif cu=="NEW": code_status="new_code_pending"; code=""
            elif code_raw: code_status="recorded"; code=norm_code(code_raw)
            else: code_status="missing_in_source"; code=""
            pre,pre_status=clean_prerequisite(pre_raw)
            minors.append({"source_workbook":MINOR_FILE.name,"source_sheet":ws.title,"source_row":rnum,"minor":ws.title,"batch":batch,"course_code_raw":code_raw,"course_code":code,"course_code_status":code_status,"course_title_raw":title_raw,"course_title":title_raw,"lecture_hours":row[c["l"]] if c["l"] is not None else None,"tutorial_hours":row[c["t"]] if c["t"] is not None else None,"practical_hours":row[c["p"]] if c["p"] is not None else None,"credits":credit,"semester":sem,"prerequisite_raw":pre_raw,"prerequisite":pre,"prerequisite_status":pre_status})

    # ========================================================
    # 2) SEMESTER-SPREAD WORKBOOK — block-driven parsing
    # ========================================================
    semesters=[]; structures=[]; requirements=[]
    wb=openpyxl.load_workbook(STRUCT_FILE,data_only=True,read_only=True)
    for ws in [wb[n] for n in wb.sheetnames[:5]]:
        vals=list(ws.iter_rows(values_only=True))
        blocks=[]
        for c0 in range(len(vals[1])):
            h2=norm(vals[1][c0]).lower()
            if h2 in {"course code","code"}:
                label=""
                for cc0 in [c0,c0+1,c0-1]:
                    if 0<=cc0<len(vals[0]):
                        v=norm(vals[0][cc0])
                        if re.fullmatch(r"S[1-8]",v,re.I): label=v.upper(); break
                if label: blocks.append((c0,label))
        batch=re.search(r"(20\d{2})",ws.title)
        batch=batch.group(1) if batch else ""
        basket=category=""
        for rnum,row in enumerate(vals[2:],start=3):
            b=norm(row[0]) if len(row)>0 else ""; cat=norm(row[1]) if len(row)>1 else ""
            if b: basket=b
            if cat: category=cat
            for c0,sem in blocks:
                vals7=list(row[c0:c0+7])+[None]*7
                code_raw=norm(vals7[0]); title_raw=norm(vals7[1]); pre_raw=norm(vals7[2])
                if not code_raw or not title_raw: continue
                cu=code_raw.upper()
                if cu in {"TBA","TBD"}: code_status="pending"; code=""
                elif cu in {"DON'T KNOW","DONT KNOW"}: code_status="not_recorded_in_source"; code=""
                else: code_status="recorded"; code=norm_code(code_raw)
                pre,pre_status=clean_prerequisite(pre_raw)
                semesters.append({"source_workbook":STRUCT_FILE.name,"source_sheet":ws.title,"source_row":rnum,"batch":batch,"basket":basket,"category":category,"semester":sem,"course_code_raw":code_raw,"course_code":code,"course_code_status":code_status,"course_title_raw":title_raw,"course_title":title_raw,"prerequisite_raw":pre_raw,"prerequisite":pre,"prerequisite_status":pre_status,"lecture_hours":vals7[3],"tutorial_hours":vals7[4],"practical_hours":vals7[5],"credits":vals7[6]})

    # ========================================================
    # 3) STRUCTURE SHEETS — exact column layouts, no invented codes
    # ========================================================
    for ws in [wb[n] for n in wb.sheetnames[5:]]:
        vals=list(ws.iter_rows(values_only=True))
        if ws.title in {"Struct_2022","Struct_2023"}:
            groups=[(5,6,"University Core"),(9,10,"Foundation"),(13,14,"Program Core"),(17,18,"Program Honors"),(21,22,"Honors Track 1"),(25,26,"Honors Track 2"),(29,30,"Internship/Capstone")]
            basket_cols=(1,2)
        else:
            groups=[(3,4,"University Core"),(7,8,"Program Core"),(11,12,"Electives")]
            basket_cols=(14,15)
        year=re.search(r"(20\d{2})",ws.title); year=year.group(1) if year else ""
        for rnum,row in enumerate(vals[2:],start=3):
            for title_col,credit_col,category in groups:
                ti,ci=title_col-1,credit_col-1
                title=norm(row[ti]) if ti<len(row) else ""; credit=row[ci] if ci<len(row) else None
                if not title or title.lower() in {"total credits","courses","course"}: continue
                structures.append({"source_workbook":STRUCT_FILE.name,"source_sheet":ws.title,"source_row":rnum,"structure_version":year,"category":category,"course_title_raw":title,"course_title":title,"credits":credit})
            bi,ci=basket_cols[0]-1,basket_cols[1]-1
            basket=norm(row[bi]) if bi<len(row) else ""; credit=row[ci] if ci<len(row) else None
            if basket and basket.lower() not in {"baskets","components"} and credit not in (None,""):
                requirements.append({"source_workbook":STRUCT_FILE.name,"source_sheet":ws.title,"source_row":rnum,"structure_version":year,"basket":basket,"credits":credit})

    minors_df=pd.DataFrame(minors); semesters_df=pd.DataFrame(semesters); structures_df=pd.DataFrame(structures); requirements_df=pd.DataFrame(requirements); placeholders_df=pd.DataFrame(placeholders)
    course_parts=[]
    for _,r in semesters_df[semesters_df.course_code!=""].iterrows():
        course_parts.append({"course_code":r.course_code,"course_title":r.course_title,"credits":r.credits,"prerequisite":r.prerequisite,"prerequisite_status":r.prerequisite_status,"batch":r.batch,"source_type":"semester_excel","source_sheet":r.source_sheet,"source_row":r.source_row})
    for _,r in minors_df[minors_df.course_code!=""].iterrows():
        course_parts.append({"course_code":r.course_code,"course_title":r.course_title,"credits":r.credits,"prerequisite":r.prerequisite,"prerequisite_status":r.prerequisite_status,"batch":r.batch,"source_type":"minor_excel","source_sheet":r.source_sheet,"source_row":r.source_row})
    courses=pd.DataFrame(course_parts)
    return minors_df,semesters_df,structures_df,requirements_df,courses,placeholders_df
# ---------- synthetic student data ----------
@st.cache_data
def synthetic_students():
    students=pd.DataFrame([
        ["SYN001","BTech","2026",5,72,"Finance"],
        ["SYN002","BTech","2026",6,96,"Psychology"],
        ["SYN003","BTech","2025",7,120,"Finance"],
        ["SYN004","BTech","2026",5,72,"Finance"],
        ["SYN005","BTech","2025",7,118,"Marketing"],
        ["SYN006","BTech","2026",6,94,"Economics"],
    ],columns=["student_id","programme","batch","semester","completed_credits","minor"])
    history=pd.DataFrame([
        ["SYN001","UCOR103","Passed","A"],["SYN001","UCOR203","Passed","B"],
        ["SYN001","UCOR205","Failed","F"],["SYN001","UCOR104","Passed","A"],
        ["SYN002","UCOR103","Passed","A"],["SYN002","UCOR203","Passed","B"],
        ["SYN002","UCOR104","Passed","B+"],
        ["SYN003","UCOR103","Passed","A"],["SYN003","UCOR203","Passed","A"],
        ["SYN003","UCOR205","Passed","B"],["SYN003","UCOR310","Passed","A-"],
        ["SYN004","UCOR103","Passed","A"],
        ["SYN005","UCOR103","Passed","B"],["SYN006","UCOR103","Passed","A"],
    ],columns=["student_id","course_code","status","grade"])
    return students,history

# ---------- RAG corpus ----------
def record_to_text(row):
    parts=[]
    for k,v in row.items():
        if norm(v): parts.append(f"{k}: {norm(v)}")
    return " | ".join(parts)

def tokenize_words(s):
    return re.findall(r"\b[\w&'-]+\b", str(s).lower())

def chunk_text(text, meta, target_tokens=500, overlap=100):
    toks=text.split()
    if len(toks)<=target_tokens:
        return [{"text":text,**meta,"chunk_index":0}]
    out=[]; start=0; idx=0
    step=max(1,target_tokens-overlap)
    while start<len(toks):
        end=min(len(toks),start+target_tokens)
        out.append({"text":" ".join(toks[start:end]),**meta,"chunk_index":idx})
        idx+=1
        if end==len(toks): break
        start+=step
    return out

@st.cache_resource(show_spinner="Preparing academic information…")
def build_rag():
    minors, semesters, structures, requirements, courses, placeholders = load_authoritative_data()
    records=[]
    for _,r in minors.iterrows():
        records.append(record_to_text(r.to_dict()))
    for _,r in semesters.iterrows():
        records.append(record_to_text(r.to_dict()))
    for _,r in structures.iterrows():
        records.append(record_to_text(r.to_dict()))

    chunks=[]
    for i,t in enumerate(records):
        chunks.extend(chunk_text(t,{"record_id":i,"source":"authoritative_excel"}))

    texts=[c["text"] for c in chunks]
    # Primary RAG: dense embeddings + FAISS.
    try:
        from sentence_transformers import SentenceTransformer
        import faiss
        model=SentenceTransformer("all-MiniLM-L6-v2")
        emb=model.encode(texts,normalize_embeddings=True,show_progress_bar=False)
        emb=emb.astype("float32")
        index=faiss.IndexFlatIP(emb.shape[1]); index.add(emb)
        return {"mode":"dense_faiss","chunks":chunks,"texts":texts,"model":model,"index":index,
                "minors":minors,"semesters":semesters,"structures":structures,"requirements":requirements,"courses":courses,"placeholders":placeholders}
    except Exception as e:
        # Explicitly labelled fallback so the system never pretends TF-IDF is dense RAG.
        from sklearn.feature_extraction.text import TfidfVectorizer
        from sklearn.metrics.pairwise import cosine_similarity
        vec=TfidfVectorizer(ngram_range=(1,2),sublinear_tf=True)
        X=vec.fit_transform(texts)
        return {"mode":"lexical_fallback","chunks":chunks,"texts":texts,"vectorizer":vec,"X":X,
                "error":str(e),"minors":minors,"semesters":semesters,"structures":structures,"requirements":requirements,"courses":courses,"placeholders":placeholders}

def retrieve(query, rag, k_candidates=40, k_final=12):
    q=norm(query)
    # query processing / expansion
    expanded=q
    synonyms={"pre requisite":"prerequisite","pre-requisite":"prerequisite",
              "enroll":"eligibility","register":"eligibility","subject":"course",
              "subjects":"courses","available":"offered"}
    for a,b in synonyms.items(): expanded=expanded.replace(a,b)
    codes=re.findall(r"\b[A-Za-z]{2,8}\d{3}\b",q.upper())

    if rag["mode"]=="dense_faiss":
        qv=rag["model"].encode([expanded],normalize_embeddings=True).astype("float32")
        scores,ids=rag["index"].search(qv,k_candidates)
        cand=[(int(i),float(s)) for i,s in zip(ids[0],scores[0]) if i>=0]
    else:
        from sklearn.metrics.pairwise import cosine_similarity
        qv=rag["vectorizer"].transform([expanded])
        scores=cosine_similarity(qv,rag["X"]).ravel()
        ids=scores.argsort()[::-1][:k_candidates]
        cand=[(int(i),float(scores[i])) for i in ids]

    qwords=set(tokenize_words(expanded))
    reranked=[]
    for i,sem_score in cand:
        txt=rag["texts"][i]
        words=set(tokenize_words(txt))
        lexical=len(qwords & words)/max(1,len(qwords))
        bonus=0.0
        for code in codes:
            if code in txt.upper(): bonus+=0.20
        title_terms=[w for w in qwords if len(w)>2]
        title_overlap=sum(1 for w in title_terms if w in txt.lower())/max(1,len(title_terms))
        final=0.72*sem_score+0.18*lexical+0.10*title_overlap+bonus
        reranked.append((final,i,sem_score,lexical))
    reranked.sort(reverse=True)
    return [rag["chunks"][i] | {"score":s,"semantic_score":ss,"lexical_score":ls}
            for s,i,ss,ls in reranked[:k_final]]

# ---------- entity resolution / deterministic verification ----------
def _clean_title(s):
    return re.sub(r"[^a-z0-9]+", " ", norm(s).lower()).strip()

def _query_words(q):
    stop={"what","which","when","where","who","how","does","do","is","are","the","a","an","for","of","to","in","on","and","or","can","i","my","me","please","tell","give","show","list","course","courses","about","with","from","this","that","does","it","its"}
    return [w for w in re.findall(r"[a-z0-9]+", q.lower()) if len(w)>1 and w not in stop]

def find_course(query, courses):
    """Resolve course codes OR course names. Names are first-class user input."""
    q=norm(query)
    qcode=norm_code(q)
    exact_code=courses[courses.course_code.eq(qcode)]
    if len(exact_code)>0:
        # The same code can have harmless punctuation/wording differences across
        # academic-year sheets. Collapse near-identical titles; only true title
        # differences become an ambiguity.
        titles=list(exact_code.course_title.astype(str).map(_clean_title))
        base=set(titles[0].split()) if titles else set()
        near=[]
        for t in titles:
            ts=set(t.split()); sim=len(base & ts)/max(1,len(base | ts))
            if sim>=0.75: near.append(t)
        if len(near)==len(titles): return "exact",exact_code.iloc[0]
        return "ambiguous",exact_code.drop_duplicates(subset=["course_code","course_title"])
    codes=re.findall(r"\b[A-Za-z]{2,8}\d{3}\b",q.upper())
    for code in codes:
        hit=courses[courses.course_code.eq(norm_code(code))]
        if len(hit)>0:
            titles=list(hit.course_title.astype(str).map(_clean_title)); base=set(titles[0].split()) if titles else set()
            if all(len(base & set(t.split()))/max(1,len(base | set(t.split())))>=0.75 for t in titles):
                return "exact",hit.iloc[0]
            return "ambiguous",hit.drop_duplicates(subset=["course_code","course_title"])

    # Exact title or title phrase embedded in the user's natural-language question.
    qt=_clean_title(q)
    title_norm=courses.course_title.map(_clean_title)
    exact_title=courses[title_norm.eq(qt)]
    if len(exact_title)==0:
        embedded=courses[title_norm.map(lambda t: bool(t) and (t in qt or qt in t))]
        exact_title=embedded
    exact_title=exact_title.drop_duplicates(subset=["course_code","course_title"])
    if len(exact_title)==1: return "title",exact_title.iloc[0]
    if len(exact_title)>1: return "ambiguous",exact_title

    words=[w for w in _query_words(q) if w not in {"prerequisite","prereq","credit","credits","offered","offering","semester","eligible","eligibility","register","registration","enroll","enrol","take","available","availability","information","details","detail"}]
    if not words: return "none",None
    scored=[]
    for _,r in courses.drop_duplicates(subset=["course_code","course_title"]).iterrows():
        title_words=set(_clean_title(r.course_title).split())
        overlap=sum(w in title_words for w in words)/max(1,len(words))
        phrase=1.0 if all(w in _clean_title(r.course_title).split() for w in words) else 0.0
        score=0.75*overlap+0.25*phrase
        if score>0: scored.append((score,r))
    scored.sort(key=lambda x:x[0],reverse=True)
    if not scored or scored[0][0]<0.55: return "none",None
    top=scored[0][0]
    ties=[r for s,r in scored if s>=max(0.55,top-0.08)]
    # Only ask a follow-up when the leading matches are genuinely close.
    if len(ties)>1 and top < 1.0:
        return "ambiguous",pd.DataFrame(ties[:8])
    return "title",scored[0][1]

def _course_rows(courses, code):
    return courses[courses.course_code.eq(norm_code(code))].drop_duplicates()

def _source_evidence(rows, label=None):
    out=[]
    for _,r in rows.iterrows():
        source=r.get("source_sheet",""); row=r.get("source_row","")
        stype=norm(r.get("source_type", ""))
        raw_workbook=norm(r.get("source_workbook", ""))
        workbook = ("Minor Courses workbook" if "Minor_Courses" in raw_workbook or stype == "minor_excel" else ("Semester & Structure workbook" if "Semester_Spread" in raw_workbook or stype in {"semester_excel","structure_excel"} else "University Excel source"))
        text=[]
        for col in ["course_code","course_title","credits","prerequisite","semester","minor","batch","lecture_hours","tutorial_hours","practical_hours","category","basket"]:
            if col in r.index and norm(r[col]): text.append(f"{col}: {norm(r[col])}")
        out.append({"source":source,"row":row,"workbook":workbook,"text":("; ".join(text) if text else norm(label or "authoritative Excel record"))})
    return out

def _is_list_query(q):
    return any(x in q for x in ["list all","list the","what courses","which courses","courses in","subjects in","show all","give me all"])

def _minor_match(q, minors):
    ql=q.lower()
    for m in sorted(minors.minor.dropna().unique(), key=len, reverse=True):
        if str(m).lower() in ql: return str(m)
    return None

def _semester_number(q):
    m=re.search(r"\b(?:semester|sem)\s*([1-8])\b",q.lower())
    return f"S{m.group(1)}" if m else None

def _structure_match(q, structures):
    ql=q.lower()
    sheets=[s for s in structures.source_sheet.dropna().unique() if str(s).lower() in ql]
    if sheets: return sheets[0]
    for token in ["2022","2023","2024","2025","2026"]:
        if token in ql:
            hit=structures[structures.source_sheet.astype(str).str.contains(token,case=False,na=False)]
            if len(hit): return hit.source_sheet.iloc[0]
    return None

def _course_fact_answer(code, title, courses, semesters, q):
    rows=_course_rows(courses,code)
    if rows.empty: return None
    # Preserve all distinct authoritative values instead of silently choosing one.
    titles=sorted(set(norm(x) for x in rows.course_title if norm(x)))
    credits=sorted(set(norm(x) for x in rows.credits if norm(x)))
    pres_raw=[norm(x) for x in rows.prerequisite if norm(x)]
    pres=[]
    for x in pres_raw:
        if x.upper() in {"NIL","NONE","NA","N/A","-"}: continue
        if x not in pres: pres.append(x)
    if not pres and pres_raw: pres=["Nil/none recorded"]
    evidence=_source_evidence(rows)
    wants_pre="prereq" in q or "pre req" in q or "pre-requisite" in q
    wants_credit=any(x in q for x in ["credit","credits","how many credit"])

    # Do not merge different academic-year/batch values into one misleading answer.
    if wants_pre:
        pre_pairs=[]
        for _,rr in rows.iterrows():
            pv=norm(rr.get("prerequisite","")) or "Not recorded"
            ver=norm(rr.get("batch","")) or norm(rr.get("source_sheet",""))
            item=(ver,pv)
            if item not in pre_pairs: pre_pairs.append(item)
        pre_values=sorted(set(v for _,v in pre_pairs))
        if len(pre_values)>1:
            detail="; ".join(f"{v}: {p}" for v,p in pre_pairs)
            return {"answer":f"The supplied records show different prerequisite entries for **{code} — {title}** across academic versions: {detail}. Please tell me the applicable batch/year so I can give the correct version.","status":"follow_up","evidence":evidence}

    if wants_credit:
        credit_pairs=[]
        for _,rr in rows.iterrows():
            cv=norm(rr.get("credits","")) or "Not recorded"
            ver=norm(rr.get("batch","")) or norm(rr.get("source_sheet",""))
            item=(ver,cv)
            if item not in credit_pairs: credit_pairs.append(item)
        credit_values=sorted(set(v for _,v in credit_pairs))
        if len(credit_values)>1:
            detail="; ".join(f"{v}: {c} credits" for v,c in credit_pairs)
            return {"answer":f"The supplied records show different credit values for **{code} — {title}** across academic versions: {detail}. Please tell me the applicable batch/year.","status":"follow_up","evidence":evidence}
    if wants_pre and wants_credit:
        val="; ".join(pres) if pres else "Nil/none recorded"
        cval=", ".join(credits) if credits else "not recorded"
        return {"answer":f"{code} — {title}: prerequisite **{val}**; credits **{cval}**.","status":"verified","evidence":evidence}
    if wants_pre:
        val="; ".join(pres) if pres else "Nil/none recorded"
        return {"answer":f"{code} — {title}: the prerequisite recorded in the supplied Excel data is **{val}**.","status":"verified","evidence":evidence}
    if wants_credit:
        val=", ".join(credits) if credits else "not recorded"
        return {"answer":f"{code} — {title}: credits recorded in the supplied Excel data: **{val}**.","status":"verified","evidence":evidence}
    if any(x in q for x in ["all details","all information","everything about","tell me about","details about","information about"]):
        items=[]
        if titles: items.append("Course title: " + "; ".join(titles))
        if credits: items.append("Credits: " + "; ".join(credits))
        if pres: items.append("Prerequisite: " + "; ".join(pres))
        else: items.append("Prerequisite: Nil/none recorded")
        rows2=semesters[semesters.course_code.astype(str).str.upper().eq(code.upper())].drop_duplicates()
        if not rows2.empty:
            items.append("Semester offerings: " + ", ".join(sorted(set(norm(x) for x in rows2.semester))))
            cats=sorted(set(norm(x) for x in rows2.category if norm(x)))
            if cats: items.append("Categories: " + "; ".join(cats))
            baskets=sorted(set(norm(x) for x in rows2.basket if norm(x)))
            if baskets: items.append("Baskets: " + "; ".join(baskets))
            evidence += _source_evidence(rows2)
        return {"answer":f"**{code} — {title}**\n\n"+"\n".join(f"- {x}" for x in items),"status":"verified","evidence":evidence}
    if any(x in q for x in ["semester and category","category and semester","semester/category","category/semester"]):
        rows2=semesters[semesters.course_code.astype(str).str.upper().eq(code.upper())].drop_duplicates()
        if rows2.empty:
            return {"answer":f"No semester-spread record was found for **{code} — {title}** in the supplied workbook.","status":"insufficient","evidence":evidence}
        sems=sorted(set(norm(x) for x in rows2.semester if norm(x)))
        cats=sorted(set(norm(x) for x in rows2.category if norm(x)))
        evidence += _source_evidence(rows2)
        return {"answer":f"**{code} — {title}** is listed in semester(s): **{', '.join(sems) if sems else 'not recorded'}**.\n\nCategory: **{'; '.join(cats) if cats else 'not recorded'}**.","status":"verified","evidence":evidence}
    if any(x in q for x in ["offer","available","when is","which semester","semester is","offered"]):
        rows2=semesters[semesters.course_code.str.upper().eq(code.upper())]
        if rows2.empty:
            return {"answer":f"{code} — {title} is present in the course data, but I do not have a semester-offering record for it in the supplied workbook.","status":"insufficient","evidence":evidence}
        offs=rows2[["source_sheet","source_row","semester","course_title","credits"]].drop_duplicates()
        sems=sorted(set(norm(x) for x in offs.semester))
        evidence += _source_evidence(offs)
        return {"answer":f"{code} — {title} is listed for: **{', '.join(sems)}** in the supplied semester-spread workbook. This is a source listing, not a guarantee of future registration availability.","status":"verified","evidence":evidence}
    return None


# ---------- comprehensive database question router ----------
def _wants_list(q):
    return any(x in q.lower() for x in ["list", "show", "which", "what courses", "give me", "all courses", "all the courses", "how many courses"])

def _format_rows(rows, fields, limit=80):
    out=[]
    for _,r in rows.head(limit).iterrows():
        vals=[]
        for f,label in fields:
            if f in r.index and norm(r[f]):
                value=norm(r[f])
                if f=="course_code" and value.upper() in {"DON'T KNOW","DONT KNOW","DONTKNOW"}: value="not recorded"
                vals.append(f"{label}: {value}")
        if vals: out.append(" • ".join(vals))
    return out

def database_query_answer(query, minors, semesters, structures, courses):
    """Answer questions whose facts can be obtained directly from the two Excel workbooks.
    This is deliberately deterministic; the LLM only verbalizes grounded evidence later.
    Returns None when the query needs course/entity resolution or student logic.
    """
    ql=query.lower()
    # ----- minor database -----
    minor=_minor_match(ql,minors)
    if minor:
        rows=minors[minors.minor.astype(str).str.lower().eq(minor.lower())].copy()
        # Optional semester/batch filters.
        sem=_semester_number(ql)
        if "semester" in ql and any(x in ql for x in ["specified semester", "particular semester", "a semester"]) and not sem:
            return {"answer":f"Which semester should I use for the {minor} minor? Please specify a semester such as Semester 3 or Semester 5.","status":"follow_up","evidence":[]}
        if sem and "semester" in rows.columns:
            rows=rows[rows.semester.astype(str).str.upper().str.replace(" ","").eq(sem)]
        batch=re.search(r"\b(?:batch|year)\s*(2022|2023|2024|2025|2026)\b",ql)
        if batch and "batch" in rows.columns:
            b=batch.group(1); rows=rows[rows.batch.astype(str).str.contains(b,case=False,na=False)]
        rows=rows.drop_duplicates(subset=[c for c in ["course_code","course_title"] if c in rows.columns])
        if len(rows)==0:
            return {"answer":f"No matching {minor} minor course records were found for the specified filters in the supplied Excel data.","status":"insufficient","evidence":[]}
        if _wants_list(ql) or "minor" in ql:
            if any(x in ql for x in ["lecture", "tutorial", "practical", "l t p", "hours"]):
                fields=[("course_code","Code"),("course_title","Course"),("lecture_hours","L"),("tutorial_hours","T"),("practical_hours","P"),("credits","Credits"),("semester","Semester"),("prerequisite","Prerequisite")]
            else:
                fields=[("course_code","Code"),("course_title","Course"),("credits","Credits"),("semester","Semester"),("prerequisite","Prerequisite")]
            items=_format_rows(rows,fields)
            return {"answer":f"The supplied **{minor}** minor workbook records {len(rows)} matching course entries:\n\n"+"\n".join(f"- {x}" for x in items),"status":"verified","evidence":_source_evidence(rows)}

    # ----- semester-spread database -----
    sem=_semester_number(ql)
    if sem and (_wants_list(ql) or any(x in ql for x in ["semester", "offered", "offering"])):
        rows=semesters[semesters.semester.astype(str).str.upper().eq(sem)].copy()
        # If a structure/year is named, constrain to that sheet.
        stname=_structure_match(ql,semesters.rename(columns={"source_sheet":"source_sheet"})) if False else None
        years=re.findall(r"20(?:22|23|24|25|26)",ql)
        if years:
            year=years[0]
            yrrows=rows[rows.source_sheet.astype(str).str.contains(year,case=False,na=False)]
            if len(yrrows): rows=yrrows
        rows=rows.drop_duplicates(subset=["course_code","course_title"])
        if len(rows):
            fields=[("course_code","Code"),("course_title","Course"),("credits","Credits"),("category","Category"),("basket","Basket"),("prerequisite","Prerequisite")]
            items=_format_rows(rows,fields)
            return {"answer":f"The supplied semester-spread workbook lists {len(rows)} distinct course records for **{sem}**:\n\n"+"\n".join(f"- {x}" for x in items),"status":"verified","evidence":_source_evidence(rows)}

    # ----- structure database -----
    stname=_structure_match(ql,structures)
    if any(x in ql for x in ["structure", "university core", "program core", "programme core", "foundation", "elective", "graduation credit", "credit requirement", "degree credit"]):
        if not stname:
            # Questions explicitly asking for a cross-structure comparison can be answered without guessing.
            if "different structures" in ql or "across structures" in ql or "all structures" in ql:
                vals=[]
                for sh in sorted(structures.source_sheet.dropna().unique()):
                    rr=structures[structures.source_sheet.eq(sh)]
                    vals.append(f"{sh}: {len(rr)} course entries")
                return {"answer":"The supplied workbook contains multiple academic structures:\n\n"+"\n".join(f"- {x}" for x in vals)+"\n\nPlease specify the structure/year for a course-level or credit-requirement answer.","status":"follow_up","evidence":[]}
            return {"answer":"Please specify the applicable academic structure/year (for example, Struct_2024, Struct_2025, or Struct_2026_DS). The supplied workbook contains multiple structures, so I will not assume one.","status":"follow_up","evidence":[]}
        rows=structures[structures.source_sheet.eq(stname)].copy()
        # category filter
        cat=None
        for c in ["university core","program core","programme core","foundation","elective","honors","specialization"]:
            if c in ql: cat=c; break
        if cat:
            if cat in {"program core","programme core"}: rows=rows[rows.category.astype(str).str.contains("Program Core",case=False,na=False)]
            elif cat=="university core": rows=rows[rows.category.astype(str).str.contains("University Core",case=False,na=False)]
            elif cat=="foundation": rows=rows[rows.category.astype(str).str.contains("Foundation",case=False,na=False)]
            elif cat=="elective": rows=rows[rows.category.astype(str).str.contains("Elective",case=False,na=False)]
            elif cat=="honors": rows=rows[rows.category.astype(str).str.contains("Honors|Hons",case=False,na=False)]
            elif cat=="specialization": rows=rows[rows.category.astype(str).str.contains("Specialization|Sp Track|Applied",case=False,na=False)]
        rows=rows.drop_duplicates(subset=["course_title","category"])
        if len(rows) and (_wants_list(ql) or cat):
            items=_format_rows(rows,[("course_title","Course"),("credits","Credits"),("category","Category")])
            return {"answer":f"For **{stname}**, the supplied structure workbook records {len(rows)} matching entries:\n\n"+"\n".join(f"- {x}" for x in items),"status":"verified","evidence":_source_evidence(rows)}

    # ----- generic field queries -----
    # Count / filtering by credits across the course records.
    m=re.search(r"\b(\d+(?:\.\d+)?)\s*credits?\b",ql)
    if m and _wants_list(ql):
        val=m.group(1)
        pool=pd.concat([courses.assign(source_group="course"),structures.assign(source_group="structure")],ignore_index=True,sort=False)
        rows=pool[pool.credits.astype(str).str.strip().eq(val)]
        rows=rows.drop_duplicates(subset=[c for c in ["course_code","course_title","category"] if c in rows.columns])
        if len(rows):
            items=_format_rows(rows,[("course_code","Code"),("course_title","Course"),("credits","Credits"),("category","Category")])
            return {"answer":f"I found {len(rows)} matching records with **{val} credits** in the supplied academic data:\n\n"+"\n".join(f"- {x}" for x in items),"status":"verified","evidence":_source_evidence(rows)}

    # No-prerequisite / prerequisite existence query.
    if _wants_list(ql) and "prerequisite" in ql and any(x in ql for x in ["no prerequisite","without prerequisite","nil prerequisite","prerequisite is nil"]):
        rows=courses[courses.prerequisite.map(lambda x: not norm(x) or norm(x).upper() in {"NIL","NONE","NA","N/A","-"})].drop_duplicates(subset=["course_code","course_title"])
        items=_format_rows(rows,[("course_code","Code"),("course_title","Course"),("credits","Credits")])
        return {"answer":f"The supplied course records contain {len(rows)} courses with no prerequisite recorded:\n\n"+"\n".join(f"- {x}" for x in items),"status":"verified","evidence":_source_evidence(rows)}

    return None

def prerequisite_tokens(pre):
    p=norm(pre)
    if not p or p.upper() in {"NIL","NONE","NA","N/A","-"}: return []
    return [norm_code(x) for x in re.findall(r"[A-Za-z]{2,8}\d{3}",p)]

def student_eligibility(student_id, course, history):
    req=prerequisite_tokens(course.get("prerequisite",""))
    if not req: return {"status":"eligible_on_prerequisite_check","required":[],"missing":[],"failed":[]}
    h=history[history.student_id.eq(student_id)]
    passed=set(norm_code(x) for x in h.loc[h.status.astype(str).str.lower().eq("passed"),"course_code"])
    failed=set(norm_code(x) for x in h.loc[~h.status.astype(str).str.lower().eq("passed"),"course_code"])
    missing=[x for x in req if x not in passed and x not in failed]
    failed_req=[x for x in req if x in failed]
    status="not_eligible_based_on_record" if failed_req else ("insufficient_student_record" if missing else "eligible_on_prerequisite_check")
    return {"status":status,"required":req,"missing":missing,"failed":failed_req}

def material_conflicts(course_code, semesters):
    """Only flag a true within-version disagreement. Different academic years are versioned records, not automatic conflicts."""
    rows=semesters[semesters.course_code.astype(str).str.upper().eq(course_code.upper())]
    conflicts=[]
    for sheet,grp in rows.groupby("source_sheet"):
        vals=[]
        for x in grp.prerequisite.astype(str):
            v=norm(x)
            if v not in vals: vals.append(v)
        if len(vals)>1:
            conflicts.append({"source_sheet":sheet,"prerequisites":vals,"sources":[f"{r.source_sheet} row {r.source_row}" for _,r in grp.iterrows()]})
    return conflicts

def direct_answer(query, student_id, rag):
    minors, semesters, structures, requirements, courses = rag["minors"],rag["semesters"],rag["structures"],rag["requirements"],rag["courses"]
    q=norm(query); ql=q.lower(); evidence=[]
    # First route questions that can be answered directly from workbook tables.
    routed=database_query_answer(q,minors,semesters,structures,courses)
    if routed is not None:
        return routed
    if not student_id:
        sid_hit=re.search(r"\bSYN\d{3}\b",q.upper())
        if sid_hit: student_id=sid_hit.group(0)

    # 1) Broad source-database questions first.
    minor_name=_minor_match(ql,minors)
    if minor_name and (_is_list_query(ql) or "minor" in ql) and not any(x in ql for x in ["prereq","credit of","when"]):
        rows=minors[minors.minor.astype(str).str.lower().eq(minor_name.lower())].copy()
        rows=rows.drop_duplicates(subset=["course_code","course_title"]) if "course_code" in rows else rows.drop_duplicates()
        if len(rows):
            items=[f"{norm(r.course_code)} — {norm(r.course_title)} ({norm(r.credits)} credits; Semester {norm(r.semester)})" for _,r in rows.iterrows()]
            evidence=_source_evidence(rows)
            return {"answer":f"The supplied **{minor_name}** minor data contains {len(items)} distinct course records:\n\n"+"\n".join(f"- {x}" for x in items),"status":"verified","evidence":evidence}

    semno=_semester_number(ql)
    if semno and _is_list_query(ql):
        rows=semesters[semesters.semester.astype(str).str.upper().eq(semno)].drop_duplicates(subset=["course_code","course_title"])
        if len(rows):
            items=[f"{norm(r.course_code)} — {norm(r.course_title)} ({norm(r.credits)} credits)" for _,r in rows.iterrows()]
            evidence=_source_evidence(rows)
            return {"answer":f"The supplied semester-spread workbook lists {len(items)} distinct courses in **{semno}**:\n\n"+"\n".join(f"- {x}" for x in items),"status":"verified","evidence":evidence}

    struct=_structure_match(ql,structures)
    if struct and any(x in ql for x in ["graduation credit","credit requirement","credits required","degree credits","how many credits"]):
        req_rows=requirements[requirements.source_sheet.eq(struct)].copy()
        if len(req_rows):
            summary=[f"{norm(r.basket)}: {norm(r.credits)} credits" for _,r in req_rows.iterrows()]
            ev=[{"source":r.source_sheet,"row":r.source_row,"text":f"{norm(r.basket)}: {norm(r.credits)} credits","workbook":"Semester & Structure workbook"} for _,r in req_rows.iterrows()]
            return {"answer":f"For **{struct}**, the supplied structure workbook records these credit requirements:\n\n"+"\n".join(f"- {x}" for x in summary),"status":"verified","evidence":ev}

    if struct and (_is_list_query(ql) or "structure" in ql or "program core" in ql or "programme core" in ql or "university core" in ql or "foundation" in ql or "elective" in ql):
        rows=structures[structures.source_sheet.eq(struct)].copy()
        if "program core" in ql or "programme core" in ql: rows=rows[rows.category.astype(str).str.contains("Program Core",case=False,na=False)]
        elif "university core" in ql: rows=rows[rows.category.astype(str).str.contains("University Core",case=False,na=False)]
        elif "foundation" in ql: rows=rows[rows.category.astype(str).str.contains("Foundation",case=False,na=False)]
        elif "elective" in ql: rows=rows[rows.category.astype(str).str.contains("Elective",case=False,na=False)]
        rows=rows.drop_duplicates(subset=["course_title","category"])
        if len(rows):
            items=[f"{norm(r.course_title)} ({norm(r.credits)} credits)" for _,r in rows.iterrows()]
            evidence=_source_evidence(rows)
            return {"answer":f"For **{struct}**, the supplied structure workbook records {len(items)} matching course entries:\n\n"+"\n".join(f"- {x}" for x in items),"status":"verified","evidence":evidence}

    # Broad category questions without a structure/year must be clarified because
    # the two workbooks contain multiple academic structures.
    if any(x in ql for x in ["university core","foundation category","foundation courses","program core","programme core","electives"]):
        if not struct:
            return {"answer":"I need the applicable academic structure/year before I can give a reliable course list for that category because the supplied workbook contains multiple structures. Please provide the year/structure (for example, 2024, 2025, or 2026).","status":"follow_up","evidence":[]}

    if "minor courses" in ql and not minor_name:
        return {"answer":"Please provide the minor name (for example, Finance, Marketing, Economics, Psychology, Design, Law, or Start-up) so I can retrieve the relevant course records.","status":"follow_up","evidence":[]}

    if "offering" in ql and "minor" in ql and not minor_name:
        return {"answer":"Please provide the minor name so I can identify its course-offering records.","status":"follow_up","evidence":[]}

    if "guarantee" in ql and "offering" in ql:
        return {"answer":"No. A semester listing in the supplied workbook is evidence that the course appears in that semester-spread record; it does not establish a guarantee of future registration availability.","status":"verified","evidence":[]}

    if "failed" in ql and "prerequisite" in ql:
        return {"answer":"If a required prerequisite is recorded as failed, the supplied student-history check treats that prerequisite as not satisfied. The supplied sources do not establish an override or exception process.","status":"verified","evidence":[]}

    if any(x in ql for x in ["source records disagree", "records disagree", "conflicting source", "source conflict", "two source records"]):
        return {"answer":"When source records materially disagree, AIRA does not silently choose one. It identifies the conflicting records, preserves their sheet/row evidence, and asks the user to verify the applicable official academic record.","status":"verified","evidence":[]}

    if "credit requirements for different structures" in ql or "different structures" in ql and "credit" in ql:
        wb2=openpyxl.load_workbook(STRUCT_FILE,data_only=True); allsum=[]
        for ws2 in wb2.worksheets[5:]:
            vals2=list(ws2.iter_rows(values_only=True))
            for rr in vals2:
                if len(rr)>=15 and norm(rr[13]) and norm(rr[14]) and norm(rr[13]).lower() not in {"\"baskets\"","baskets"}:
                    allsum.append(f"{ws2.title}: {norm(rr[13])} — {norm(rr[14])} credits")
        return {"answer":"Credit requirements recorded across the supplied structures:\n\n"+"\n".join(f"- {x}" for x in allsum),"status":"verified","evidence":[{"source":"structure sheets","row":"","text":x} for x in allsum]}

    # 2) Course-name/code resolution. Course name is a first-class input.
    kind, course=find_course(q,courses)
    if kind=="ambiguous":
        rows=course.drop_duplicates(subset=["course_code","course_title"])
        choices=", ".join(sorted(f"{norm(r.course_code)} — {norm(r.course_title)}" for _,r in rows.iterrows()))
        return {"answer":f"I found multiple possible matches: **{choices}**. Please tell me which course you mean. I will not guess.","status":"follow_up","evidence":_source_evidence(rows)}
    if kind=="none":
        if any(x in ql for x in ["can i", "can ", "eligible", "register", "enroll", "enrol", "am i allowed"]):
            return {"answer":"To check eligibility, please provide/select the course you mean. If you are asking about your own eligibility, I also need the relevant student profile or synthetic student ID.","status":"follow_up","evidence":[]}
        if "credit" in ql and any(x in ql for x in ["degree","graduat","required","need"]):
            cats=structures.groupby(["source_sheet","category"],dropna=False).size().reset_index()
            return {"answer":"I need the applicable academic structure/batch to determine a graduation-credit requirement because the supplied workbook contains multiple structure years. Please provide the structure/year (for example, 2024, 2025, or 2026).","status":"follow_up","evidence":_source_evidence(cats)}
        if any(x in ql for x in ["can i","eligible","register","enroll","enrol"]):
            return {"answer":"I could not identify the course in the supplied academic data. Please provide the course name or code exactly as it appears in the university records.","status":"follow_up","evidence":[]}
        return {"answer":"I could not identify the requested information in the supplied university Excel sources. Please provide a more specific course name, course code, minor, semester, structure/year, or other detail.","status":"insufficient","evidence":[]}

    code=norm_code(course.course_code); title=norm(course.course_title)
    fact=_course_fact_answer(code,title,courses,semesters,ql)
    conflicts=material_conflicts(code,semesters)
    if conflicts:
        base_ev=_source_evidence(_course_rows(courses,code))
        return {"answer":f"The supplied semester-spread data contains differing prerequisite records for **{code} — {title}**. I will not silently choose between them. The source records should be verified against the current official academic authority.","status":"conflict","evidence":base_ev}
    if fact:
        # Student-specific wording still requires a profile.
        if any(x in ql for x in ["can i","can ","eligible","register","enroll","enrol","am i allowed"]):
            pass
        else: return fact

    # 3) Student eligibility.
    if any(x in ql for x in ["can i","can we","can ","eligible","register","enroll","enrol","am i allowed"]):
        if not student_id:
            return {"answer":f"I found **{code} — {title}**. To check your eligibility, please select a synthetic student profile so I can compare its recorded history with the source prerequisite.","status":"follow_up","evidence":_source_evidence(_course_rows(courses,code))}
        _,history=synthetic_students(); result=student_eligibility(student_id,course.to_dict(),history)
        ev=_source_evidence(_course_rows(courses,code))
        if result["status"]=="eligible_on_prerequisite_check":
            ans=f"For **{student_id}**, the recorded prerequisite evidence for {code} is satisfied. This verifies the prerequisite check only; the supplied sources do not establish every possible registration rule."
        elif result["status"]=="not_eligible_based_on_record":
            ans=f"For **{student_id}**, the recorded history does not satisfy the prerequisite check for {code}: a required prerequisite is recorded as failed. The supplied sources do not establish an override process."
        else:
            ans=f"I do not have enough student-record information to verify eligibility for {code}. Required prerequisite evidence not demonstrated: {', '.join(result['missing'])}."
        return {"answer":ans,"status":"verified" if result["status"]!="insufficient_student_record" else "insufficient","evidence":ev}

    # 4) General course fact response covers combined questions.
    rows=_course_rows(courses,code)
    ev=_source_evidence(rows)
    credits=sorted(set(norm(x) for x in rows.credits if norm(x)))
    pres_raw=[norm(x) for x in rows.prerequisite if norm(x)]
    pres=[]
    for x in pres_raw:
        if x.upper() in {"NIL","NONE","NA","N/A","-"}: continue
        if x not in pres: pres.append(x)
    if not pres and pres_raw: pres=["Nil/none recorded"]
    return {"answer":f"**{code} — {title}**\n\nCredits: {', '.join(credits) if credits else 'not recorded'}\n\nPrerequisite: {', '.join(pres) if pres else 'Nil/none recorded'}.","status":"verified","evidence":ev}

# ---------- LLM ----------
SYSTEM = """You are AIRA, an academic decision-support assistant.
Use ONLY the supplied university evidence and synthetic student data.
Never invent a course, prerequisite, credit, offering, policy, or registration rule.
If evidence is missing, ambiguous, or conflicting, say so and ask the smallest useful follow-up question.
If structured verification disagrees with an LLM inference, structured verification wins.
Do not claim that an offering record guarantees future availability.
Answer concisely. Do not provide unsupported recommendations."""

def call_gemini(user_prompt):
    key=st.secrets.get("GEMINI_API_KEY",os.getenv("GEMINI_API_KEY",""))
    if not key: return None,"GEMINI_API_KEY is not configured."
    try:
        from google import genai
        client=genai.Client(api_key=key)
        models=[st.secrets.get("GEMINI_MODEL","gemini-2.5-flash"),
                "gemini-2.5-flash-lite"]
        last=None
        for model in dict.fromkeys(models):
            try:
                r=client.models.generate_content(model=model,contents=SYSTEM+"\n\n"+user_prompt)
                txt=getattr(r,"text",None)
                if txt: return txt,None
            except Exception as e: last=e
        return None,f"Gemini request failed: {last}"
    except Exception as e:
        return None,f"Gemini SDK error: {e}"

def llm_layer_answer(layer, query, evidence_text, structured_text):
    if layer=="Basic LLM":
        prompt=f"""Answer the following question as a general language-model baseline.
Do not claim a university-specific fact unless it is explicitly supplied in the question.
QUESTION:
{query}"""
    elif layer=="Structured Prompting":
        prompt=f"""Use a strict academic-advisor format.
QUESTION:
{query}
OUTPUT:
1. Answer or uncertainty
2. What information is missing, if any
3. Do not invent university-specific facts."""
    elif layer=="LLM + RAG":
        prompt=f"""Answer using ONLY these retrieved university records.
<RETRIEVED_EVIDENCE>
{evidence_text}
</RETRIEVED_EVIDENCE>
QUESTION:
{query}"""
    else:
        prompt=f"""Answer using the retrieved university evidence and verified synthetic student context.
<UNIVERSITY_EVIDENCE>
{evidence_text}
</UNIVERSITY_EVIDENCE>
<STUDENT_CONTEXT>
{structured_text}
</STUDENT_CONTEXT>
QUESTION:
{query}"""
    return call_gemini(prompt)

# ---------- USER-FACING UI ----------
minors, semesters, structures, requirements, courses, placeholders = load_authoritative_data()
students, history = synthetic_students()
rag=build_rag()

SOURCE_NAMES = {
    "Minor_Courses_for_BTech_Students(3)": "Minor Courses workbook",
    "Semester_Spread_Structures_Sept_2026(3)": "Semester & Structure workbook",
}

def friendly_source(e):
    raw = str(e.get("source", "University Excel source"))
    label = str(e.get("workbook", SOURCE_NAMES.get(raw.replace(".xlsx", ""), raw)))
    row = e.get("row", "")
    if row and str(row) not in {"nan", "None"}:
        return f"{label} • {raw} • row {row}"
    return f"{label} • {raw}"

def status_message(status):
    return {
        "verified": ("✓ Answer supported by the supplied academic records", "success"),
        "follow_up": ("One more detail is needed — I will not guess", "info"),
        "insufficient": ("I could not find enough information in the supplied academic records", "warning"),
        "conflict": ("The supplied records contain a difference that needs verification", "warning"),
    }.get(status, ("Answer based on the supplied academic records", "info"))

def render_sources(evidence):
    if not evidence:
        st.caption("No specific source record was available for this response.")
        return
    seen=set(); shown=0
    for e in evidence:
        key=(str(e.get("source")),str(e.get("row")),str(e.get("text")))
        if key in seen: continue
        seen.add(key); shown += 1
        st.markdown(
            f'<div class="source-card"><div class="source-label">Source {shown}</div>'
            f'<div class="source-detail">{friendly_source(e)}</div>'
            f'<div class="source-detail">{norm(e.get("text",""))[:500]}</div></div>',
            unsafe_allow_html=True,
        )
        if shown >= 8: break

def answer_question(query, student_id=None):
    t0=time.time()
    hits=retrieve(query,rag)
    base=direct_answer(query,student_id,rag)
    ev="\n".join(f"- {h.get('text','')} (source: {h.get('source','authoritative Excel')})" for h in hits)
    if base.get("evidence"):
        ev += "\nVERIFIED SOURCE RECORDS:\n" + "\n".join(
            f"- {e.get('text','')} | sheet={e.get('source','')} row={e.get('row','')}" for e in base["evidence"]
        )
    student_text=""
    if student_id:
        s=students[students.student_id==student_id].iloc[0]
        h=history[history.student_id==student_id]
        student_text=f"Student: {s.to_dict()}\nHistory: {h.to_dict('records')}"
    llm=None
    if st.session_state.get("use_llm", True):
        llm,_=llm_layer_answer("LLM + RAG + Structured Student Data",query,ev,student_text)
    if base["status"] in {"conflict","insufficient","follow_up"}:
        final=base["answer"]
    elif base.get("evidence"):
        final=base["answer"]
    else:
        final=llm if llm else base["answer"]
    return final, base, hits, time.time()-t0

st.markdown('''<div class="hero">
<span class="badge">Vidyashilp University • Academic Guidance</span>
<h1>🎓 AIRA</h1>
<p>Your simple academic advisor. Ask a question in your own words and get an evidence-based answer.</p>
<div class="hero-note"><b>No technical knowledge needed.</b> Ask using a course name, course code, minor, semester, or student profile.</div>
</div>''', unsafe_allow_html=True)

with st.sidebar:
    st.markdown("### 👤 My student profile")
    student_labels=["No profile selected"]+[f"{r.student_id} • {r.minor} • Semester {r.semester}" for _,r in students.iterrows()]
    selected=st.selectbox("Select only if you want an eligibility check",student_labels,label_visibility="collapsed")
    student_id=None if selected=="No profile selected" else selected.split(" • ")[0]
    st.divider()
    st.markdown("### 💡 You can ask")
    st.caption("📘 Courses and prerequisites")
    st.caption("🎯 Eligibility")
    st.caption("📚 Minors")
    st.caption("🏛️ Academic structure")
    st.caption("📅 Semester offerings")
    st.divider()
    st.caption("Academic information comes only from the 2 supplied university Excel workbooks.")

st.markdown("### What would you like to know?")
cols=st.columns(4)
quick=[
    ("📘 Courses","Prerequisites, credits and course details","What is the prerequisite for UCOR205?"),
    ("🎯 Eligibility","Check a synthetic student's prerequisite eligibility","Can SYN001 take UCOR205?"),
    ("📚 Minors","Explore courses in a minor","What courses are in the Finance minor?"),
    ("🏛️ Structure","Understand semesters and academic structures","What courses are in the Program Core of Struct_2025?"),
]
for c,(title,desc,qtext) in zip(cols,quick):
    with c:
        st.markdown(f'<div class="question-card"><h4>{title}</h4><p>{desc}</p></div>',unsafe_allow_html=True)
        if st.button("Try this",key="quick_"+title,use_container_width=True):
            st.session_state["pending_question"]=qtext

pending=st.session_state.pop("pending_question",None)
q=st.chat_input("Ask in your own words — for example, “What are the prerequisites for Communication Skills?”")
query=q or pending

if query:
    final,base,hits,elapsed=answer_question(query,student_id)
    msg,kind=status_message(base["status"])
    if kind=="success": st.success(msg)
    elif kind=="warning": st.warning(msg)
    else: st.info(msg)
    st.markdown('<div class="answer"><h3>Answer</h3></div>',unsafe_allow_html=True)
    st.markdown(final)
    st.markdown("### 📌 Sources")
    source_items = base.get("evidence") if base["status"] in {"verified", "conflict"} and base.get("evidence") else (hits if base["status"] == "verified" else [])
    render_sources(source_items)
    if base["status"]=="follow_up":
        st.info("Please answer the follow-up question above. I will use your response to continue rather than guessing.")
else:
    st.info("Start with a question above. Try a course name if you do not know its code.")
    st.markdown("### Examples")
    examples=[
        "What is the prerequisite for Financial Management?",
        "How many credits does UCOR205 have?",
        "When is UCOR205 offered?",
        "What courses are in the Finance minor?",
        "Which courses are in Semester 5?",
        "What are the credit requirements for Struct_2024?",
        "Can SYN001 take UCOR205?",
        "What should I do if two records disagree?",
    ]
    for i in range(0,len(examples),2):
        c1,c2=st.columns(2)
        for c,ex in zip((c1,c2),examples[i:i+2]):
            with c:
                if st.button("💬 "+ex,key=f"example_{i}_{ex}",use_container_width=True):
                    st.session_state["pending_question"]=ex
                    st.rerun()

with st.expander("ℹ️ How AIRA works — simple view"):
    st.markdown("""
**1. Clean & organize** — the two supplied Excel workbooks are normalized, cleaned, and converted into structured academic records while preserving their original sheet and row references.

**2. Retrieve** — relevant records are found using semantic search and keyword/course matching.

**3. Verify** — factual answers, eligibility checks, ambiguity, missing information, and material conflicts are checked against the structured records.

**4. Explain** — the LLM turns grounded evidence into a natural-language response. It is not allowed to override verified source facts.

**5. Cite** — the response shows the workbook, sheet, and row used as evidence whenever a specific source record is available.
""")

with st.expander("🔬 Faculty / Evaluation Lab"):
    st.markdown("**Basic LLM → Structured Prompting → LLM + RAG → LLM + RAG + Structured Student Data**")
    st.session_state["use_llm"]=st.checkbox("Use Gemini as the explanation layer",value=True,help="The production answer remains grounded in deterministic source verification.")
    cases=[
        "What is the prerequisite for Communication Skills?","What is the prerequisite for UCOR205?","How many credits does Communication Skills carry?","When is UCOR205 offered?","When is Communication Skills offered?","Tell me all information about UCOR203.","Can SYN001 take UCOR205?","Can I take UCOR205?","What courses are in the Finance minor?","What courses are offered in Semester 5?","What courses are in the Program Core of Struct_2025?","What courses are in the University Core of Struct_2026_DS?","What are the graduation credit requirements in Struct_2024?","Can I graduate with 100 credits?","What is the prerequisite and credit for UCOR203?","Which course is called Communication Skills?","What information is available about Learning to Learn?","Is a listed semester offering a guarantee of registration?","What happens if I failed a prerequisite?","Can I register without a student profile?","Tell me about a course that is not in the source data.","Which courses are in the Foundation category of Struct_2022?","Which courses are in the Electives of Struct_2026_DS?","Which courses have 3 credits?","Which courses have no prerequisite?","Show Finance minor courses in a specified semester.","What are the credit requirements across different structures?","What should I do if two source records disagree?","What does the Finance minor record for lecture, tutorial and practical hours?","Which semester and category contain UCOR205?"
    ]
    st.dataframe(pd.DataFrame({"Test case":range(1,len(cases)+1),"Query":cases}),use_container_width=True,hide_index=True)
