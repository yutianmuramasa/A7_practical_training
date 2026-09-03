
import os
import re
import glob
import pickle
import numpy as np

EMBED_MODEL = "BAAI/bge-small-zh-v1.5"
CACHE_DIR = "rag_cache"
DEFAULT_SYSTEM = (
    "你是教学助手。请只依据给定的资料回答问题；资料中没有的信息要明说"
    "\u201c资料里没有提到\u201d，不要编造。回答简明、准确，可以标注资料出处。"
)


def extract_text(path):
    """读取 txt/md/docx/pdf 文本（复用与 app.py 一致的解析方式）。"""
    ext = os.path.splitext(path)[1].lower()
    if ext in (".txt", ".md"):
        with open(path, "r", encoding="utf-8", errors="ignore") as f:
            return f.read()
    if ext == ".docx":
        from docx import Document
        doc = Document(path)
        return "\n".join(p.text for p in doc.paragraphs)
    if ext == ".pdf":
        import fitz
        doc = fitz.open(path)
        return "\n".join(page.get_text() for page in doc)
    return ""


def chunk_text(text, size=500, overlap=80):
    """把长文本切成有重叠的片段，保证检索时上下文完整。"""
    text = re.sub(r"\s+", " ", text).strip()
    if len(text) <= size:
        return [text] if text else []
    chunks, i = [], 0
    while i < len(text):
        chunks.append(text[i:i + size])
        i += size - overlap
    return chunks


def _load_docs(folder, exts=(".txt", ".md", ".docx", ".pdf")):
    files = []
    for ext in exts:
        files += glob.glob(os.path.join(folder, "**", "*" + ext), recursive=True)
    docs = []  # [(片段文本, 来源文件名), ...]
    for f in files:
        for ch in chunk_text(extract_text(f)):
            if ch:
                docs.append((ch, os.path.basename(f)))
    return docs


class Rag:
    """向量索引 + 检索 + 问答。"""

    def __init__(self):
        self.model = None
        self.faiss = None
        self.chunks = []   # [(文本, 来源)]
        self.index = None

    # ---------- 构建 ----------
    def _ensure_model(self):
        if self.model is None:
            from sentence_transformers import SentenceTransformer
            import faiss
            self.model = SentenceTransformer(EMBED_MODEL)
            self.faiss = faiss

    def embed(self, texts):
        self._ensure_model()
        vecs = self.model.encode(list(texts), normalize_embeddings=True)
        return np.asarray(vecs, dtype="float32")

    def build(self, folder):
        self._ensure_model()
        self.chunks = _load_docs(folder)
        if not self.chunks:
            raise RuntimeError("没有读到任何资料，请确认文件夹里有 txt/md/docx/pdf")
        vecs = self.embed([c for c, _ in self.chunks])
        self.index = self.faiss.IndexFlatIP(vecs.shape[1])
        self.index.add(vecs)
        self.save(folder)
        return self

    # ---------- 缓存 ----------
    def _cache_paths(self, folder):
        tag = os.path.basename(os.path.normpath(folder))
        os.makedirs(CACHE_DIR, exist_ok=True)
        base = os.path.join(CACHE_DIR, tag)
        return base + "_chunks.pkl", base + "_index.faiss"

    def save(self, folder):
        self._ensure_model()
        cp, ip = self._cache_paths(folder)
        with open(cp, "wb") as f:
            pickle.dump(self.chunks, f)
        self.faiss.write_index(self.index, ip)

    def load(self, folder):
        cp, ip = self._cache_paths(folder)
        if not (os.path.exists(cp) and os.path.exists(ip)):
            return None
        self._ensure_model()
        with open(cp, "rb") as f:
            self.chunks = pickle.load(f)
        self.index = self.faiss.read_index(ip)
        return self

    # ---------- 检索 ----------
    def retrieve(self, question, top_k=4):
        qv = self.embed([question])
        scores, idxs = self.index.search(qv, top_k)
        out = []
        for j, i in enumerate(idxs[0]):
            if i < 0 or i >= len(self.chunks):
                continue
            chunk, src = self.chunks[i]
            out.append((chunk, float(scores[0][j]), src))
        return out


def build_rag(folder):
    rag = Rag()
    if rag.load(folder) is None:
        rag.build(folder)
    return rag


def _call_deepseek(user_text, system=DEFAULT_SYSTEM, max_tokens=800, temperature=0.3):
    """生成环节调 DeepSeek（key 从环境变量读，不要写进代码）。"""
    import os as _os
    from openai import OpenAI
    key = _os.environ.get("DEEPSEEK_API_KEY")
    if not key:
        return "[缺少环境变量 DEEPSEEK_API_KEY：请先设置再运行]"
    client = OpenAI(api_key=key, base_url="https://api.deepseek.com/v1")
    resp = client.chat.completions.create(
        model="deepseek-chat",
        messages=[
            {"role": "system", "content": system},
            {"role": "user", "content": user_text},
        ],
        temperature=temperature,
        max_tokens=max_tokens,
    )
    return resp.choices[0].message.content


def rag_ask(rag, question, top_k=4):
    """检索相关片段 + DeepSeek 依据资料回答。"""
    hits = rag.retrieve(question, top_k)
    if not hits:
        return "[检索不到相关资料，换一个问法或先确认资料库内容]"
    context = "\n\n".join(f"【{src}】{chunk}" for chunk, _, src in hits)
    user_text = (
        f"以下是资料片段：\n{context}\n\n"
        f"问题：{question}\n"
        f"请依据上面的资料给出简明、准确的回答；资料里没有的信息请直接说明。"
    )
    return _call_deepseek(user_text)
