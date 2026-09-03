
import sys
from rag_utils import build_rag, rag_ask

def main():
    folder = "knowledge_files"
    question = sys.argv[1] if len(sys.argv) > 1 else "什么是TensorFlow.js？请依据资料回答"
    rag = build_rag(folder)
    print(f"[索引] 资料片段数：{len(rag.chunks)}")
    hits = rag.retrieve(question, top_k=2)
    print("[检索到的片段]")
    for chunk, score, src in hits:
        print(f"  - 来源《{src}》 相似度 {score:.3f}：{chunk[:80]}...")
    print("\n[生成回答]")
    print(rag_ask(rag, question))

if __name__ == "__main__":
    main()
