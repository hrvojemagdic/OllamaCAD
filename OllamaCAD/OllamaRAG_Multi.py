# OllamaRAG_Multi.py
# Folder-based RAG over: PDF, images, text files, CSV, Excel (multi-sheet)
# OCR via Ollama vision model, embeddings via Ollama embeddings, FAISS store, QA via Ollama text model

import os, re, io, csv, json, pickle, hashlib
from dataclasses import dataclass
from typing import List, Dict, Any, Optional

import numpy as np
from tqdm import tqdm
from PIL import Image
import ollama

try:
    from pdf2image import convert_from_path
except Exception:
    convert_from_path = None

# NEW: pandas for Excel
try:
    import pandas as pd
except Exception:
    pd = None


@dataclass
class PipelineConfig:
    poppler_path: Optional[str] = None

    ocr_model: str = "qwen3-vl:8b-instruct-q4_K_M"
    qa_model: str = "gemma3:12b-it-q4_K_M"
    embed_model: str = "qwen3-embedding:8b-q4_K_M"

    dpi: int = 220
    max_ocr_img_side: int = 1600

    chunk_chars: int = 1200
    chunk_overlap: int = 200

    top_k: int = 10

    work_dir: str = "rag_store"
    index_name: str = "faiss.index"
    meta_name: str = "meta.pkl"
    manifest_name: str = "manifest.json"


def ensure_dir(p: str) -> None:
    os.makedirs(p, exist_ok=True)


def normalize_ws(text: str) -> str:
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def chunk_text(text: str, chunk_chars: int, overlap: int) -> List[str]:
    text = normalize_ws(text)
    if not text:
        return []
    chunks = []
    start = 0
    while start < len(text):
        end = min(len(text), start + chunk_chars)
        chunks.append(text[start:end])
        if end == len(text):
            break
        start = max(0, end - overlap)
    return chunks


def downscale_image(img: Image.Image, max_side: int) -> Image.Image:
    w, h = img.size
    m = max(w, h)
    if m <= max_side:
        return img
    scale = max_side / m
    return img.resize((int(w * scale), int(h * scale)), Image.BICUBIC)


def pil_to_png_bytes(img: Image.Image) -> bytes:
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


def sha256_file(path: str) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


class OllamaOCR:
    def __init__(self, model: str, max_side: int = 1600):
        self.model = model
        self.max_side = max_side

    def ocr_image(self, img: Image.Image) -> str:
        img = downscale_image(img.convert("RGB"), self.max_side)
        img_bytes = pil_to_png_bytes(img)

        prompt = (
            "You are an OCR engine. Extract ALL readable text from the image.\n"
            "Rules:\n"
            "- Output only the extracted text.\n"
            "- Preserve paragraphs and line breaks when reasonable.\n"
            "- Do not add commentary.\n"
        )

        resp = ollama.chat(
            model=self.model,
            messages=[{
                "role": "user",
                "content": prompt,
                "images": [img_bytes],
            }],
        )
        return normalize_ws(resp["message"]["content"])


class OllamaEmbedder:
    def __init__(self, model: str):
        self.model = model

    def embed(self, texts: List[str]) -> np.ndarray:
        vecs = []
        for t in tqdm(texts, desc="Embedding (Ollama)"):
            r = ollama.embeddings(model=self.model, prompt=t)
            v = np.asarray(r["embedding"], dtype=np.float32)
            v = v / (np.linalg.norm(v) + 1e-12)
            vecs.append(v)
        return np.vstack(vecs)


class FaissStore:
    def __init__(self):
        import faiss
        self.faiss = faiss
        self.index = None
        self.dim = None

    def build(self, vectors: np.ndarray) -> None:
        self.dim = vectors.shape[1]
        self.index = self.faiss.IndexFlatIP(self.dim)
        self.index.add(vectors.astype(np.float32))

    def search(self, query_vec: np.ndarray, top_k: int):
        if query_vec.ndim == 1:
            query_vec = query_vec[None, :]
        D, I = self.index.search(query_vec.astype(np.float32), top_k)
        return D[0], I[0]

    def save(self, path: str) -> None:
        self.faiss.write_index(self.index, path)

    def load(self, path: str) -> None:
        self.index = self.faiss.read_index(path)
        self.dim = self.index.d


class OllamaQA:
    def __init__(self, model: str):
        self.model = model

    def answer(self, question: str, contexts: List[Dict[str, Any]]) -> str:
        ctx_blocks = []
        for c in contexts:
            tag = c.get("tag", "[src]")
            ctx_blocks.append(f"{tag}\n{c['text']}\n")
        context_text = "\n---\n".join(ctx_blocks)

        system = (
            "You answer questions using ONLY the provided context.\n"
            "Rules:\n"
            "- Plain text only.\n"
            "- If the answer isn't in the context, say: I don't know (not in RAG folder).\n"
            "- Cite sources using tags like [file:... p003 c001] or [file:... sheet:... row012 c000].\n"
        )

        user = (
            f"CONTEXT:\n{context_text}\n\n"
            f"QUESTION:\n{question}\n\n"
            "Write the best possible answer with citations."
        )

        resp = ollama.chat(
            model=self.model,
            messages=[
                {"role": "system", "content": system},
                {"role": "user", "content": user},
            ],
        )
        return resp["message"]["content"].strip()


class RAGPipeline:
    def __init__(self, cfg: PipelineConfig):
        self.cfg = cfg
        ensure_dir(cfg.work_dir)
        self.ocr = OllamaOCR(cfg.ocr_model, cfg.max_ocr_img_side)
        self.embedder = OllamaEmbedder(cfg.embed_model)
        self.qa = OllamaQA(cfg.qa_model)
        self.store = FaissStore()
        self.meta: List[Dict[str, Any]] = []

    def _manifest_path(self):
        return os.path.join(self.cfg.work_dir, self.cfg.manifest_name)

    def _load_manifest(self) -> Dict[str, Any]:
        p = self._manifest_path()
        if not os.path.exists(p):
            return {"files": {}}
        with open(p, "r", encoding="utf-8") as f:
            return json.load(f)

    def _save_manifest(self, m: Dict[str, Any]) -> None:
        p = self._manifest_path()
        with open(p, "w", encoding="utf-8") as f:
            json.dump(m, f, indent=2)

    # NEW: helper to make a compact "col=value | col=value" row string
    def _format_df_row(self, cols: List[str], values: List[Any]) -> str:
        parts = []
        for c, v in zip(cols, values):
            if v is None:
                continue
            s = str(v).strip()
            if not s or s.lower() in ("nan", "none"):
                continue
            parts.append(f"{c}={s}")
        return " | ".join(parts).strip()

    def ingest_folder(self, folder: str) -> None:
        folder = os.path.abspath(folder)
        if not os.path.isdir(folder):
            raise FileNotFoundError(folder)

        manifest = self._load_manifest()
        prev = manifest.get("files", {})

        exts_pdf = {".pdf"}
        exts_img = {".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp"}
        exts_txt = {".txt", ".md", ".log", ".json", ".xml", ".html"}
        exts_csv = {".csv"}
        # NEW: Excel extensions
        exts_xlsx = {".xlsx", ".xls", ".xlsm"}  # add ".xlsb" only if you install pyxlsb

        files = []
        for root, _, names in os.walk(folder):
            for n in names:
                p = os.path.join(root, n)
                ext = os.path.splitext(n.lower())[1]
                if ext in (exts_pdf | exts_img | exts_txt | exts_csv | exts_xlsx):
                    files.append(p)

        if not files:
            raise RuntimeError("No supported files found in folder.")

        all_chunks: List[str] = []
        meta: List[Dict[str, Any]] = []

        for path in tqdm(files, desc="Ingest files"):
            ext = os.path.splitext(path.lower())[1]
            rel = os.path.relpath(path, folder).replace("\\", "/")
            file_hash = sha256_file(path)
            prev[rel] = {"sha256": file_hash}

            if ext == ".pdf":
                if convert_from_path is None:
                    raise RuntimeError("pdf2image not installed or failed to import.")
                imgs = convert_from_path(path, dpi=self.cfg.dpi, poppler_path=self.cfg.poppler_path)
                for i, img in enumerate(imgs):
                    page = i + 1
                    text = self.ocr.ocr_image(img)
                    for ci, ch in enumerate(chunk_text(text, self.cfg.chunk_chars, self.cfg.chunk_overlap)):
                        all_chunks.append(ch)
                        meta.append({
                            "source": os.path.abspath(path),
                            "rel": rel,
                            "file_type": "pdf",
                            "page": page,
                            "chunk_id": ci,
                            "tag": f"[file:{rel} p{page:03d} c{ci:03d}]",
                            "text": ch,
                        })

            elif ext in exts_img:
                img = Image.open(path)
                text = self.ocr.ocr_image(img)
                for ci, ch in enumerate(chunk_text(text, self.cfg.chunk_chars, self.cfg.chunk_overlap)):
                    all_chunks.append(ch)
                    meta.append({
                        "source": os.path.abspath(path),
                        "rel": rel,
                        "file_type": "image",
                        "page": None,
                        "chunk_id": ci,
                        "tag": f"[file:{rel} c{ci:03d}]",
                        "text": ch,
                    })

            elif ext in exts_txt:
                with open(path, "r", encoding="utf-8", errors="ignore") as f:
                    text = normalize_ws(f.read())
                for ci, ch in enumerate(chunk_text(text, self.cfg.chunk_chars, self.cfg.chunk_overlap)):
                    all_chunks.append(ch)
                    meta.append({
                        "source": os.path.abspath(path),
                        "rel": rel,
                        "file_type": "text",
                        "page": None,
                        "chunk_id": ci,
                        "tag": f"[file:{rel} c{ci:03d}]",
                        "text": ch,
                    })

            elif ext in exts_csv:
                with open(path, "r", encoding="utf-8", errors="ignore", newline="") as f:
                    reader = csv.reader(f)
                    for r_i, row in enumerate(reader):
                        line = " | ".join([c.strip() for c in row if c is not None])
                        if not line.strip():
                            continue
                        all_chunks.append(line)
                        meta.append({
                            "source": os.path.abspath(path),
                            "rel": rel,
                            "file_type": "csv",
                            "row": r_i,
                            "chunk_id": 0,
                            "tag": f"[file:{rel} row{r_i:03d} c000]",
                            "text": line,
                        })

            # NEW: Excel (multi-sheet)
            elif ext in exts_xlsx:
                if pd is None:
                    raise RuntimeError("pandas not installed. Run: pip install pandas openpyxl")

                # Read ALL sheets
                # dtype=str keeps values as strings; engine auto-detected (openpyxl for xlsx/xlsm)
                sheets = pd.read_excel(path, sheet_name=None, dtype=str)

                for sheet_name, df in sheets.items():
                    if df is None or df.empty:
                        continue

                    # Normalize columns
                    cols = [str(c).strip() if c is not None else "" for c in df.columns.tolist()]
                    cols = [c if c else f"col{j}" for j, c in enumerate(cols)]

                    # Fill NaNs with empty string
                    df = df.fillna("")

                    for r_i, row in enumerate(df.itertuples(index=False, name=None)):
                        line = self._format_df_row(cols, list(row))
                        if not line.strip():
                            continue

                        all_chunks.append(line)
                        meta.append({
                            "source": os.path.abspath(path),
                            "rel": rel,
                            "file_type": "excel",
                            "sheet": sheet_name,
                            "row": r_i,
                            "chunk_id": 0,
                            "tag": f"[file:{rel} sheet:{sheet_name} row{r_i:03d} c000]",
                            "text": line,
                        })

        if not all_chunks:
            raise RuntimeError("No text extracted from supported files.")

        vectors = self.embedder.embed(all_chunks)
        self.store.build(vectors)
        self.meta = meta

        idx_path = os.path.join(self.cfg.work_dir, self.cfg.index_name)
        meta_path = os.path.join(self.cfg.work_dir, self.cfg.meta_name)

        self.store.save(idx_path)
        with open(meta_path, "wb") as f:
            pickle.dump(self.meta, f)

        manifest["files"] = prev
        self._save_manifest(manifest)

        print(f"Saved index: {idx_path}")
        print(f"Saved metadata: {meta_path}")
        print(f"Chunks indexed: {len(self.meta)}")

    def load(self) -> None:
        idx_path = os.path.join(self.cfg.work_dir, self.cfg.index_name)
        meta_path = os.path.join(self.cfg.work_dir, self.cfg.meta_name)
        if not (os.path.exists(idx_path) and os.path.exists(meta_path)):
            raise FileNotFoundError("Index not found. Build index first.")
        self.store.load(idx_path)
        with open(meta_path, "rb") as f:
            self.meta = pickle.load(f)

    def query(self, question: str) -> str:
        qvec = self.embedder.embed([question])[0]
        _, idxs = self.store.search(qvec, self.cfg.top_k)
        contexts = [self.meta[i] for i in idxs if i >= 0]
        if not contexts:
            return "I don't know (not in RAG folder)."
        return self.qa.answer(question, contexts)


def main():
    import argparse
    p = argparse.ArgumentParser()
    p.add_argument("--dir", type=str, help="Folder to ingest (OllamaRAG)")
    p.add_argument("--ask", type=str, help="Question to ask")
    p.add_argument("--store", type=str, default="rag_store", help="Store dir")
    p.add_argument("--poppler", type=str, default=None, help="Poppler bin path if not in PATH")
    p.add_argument("--topk", type=int, default=10, help="Top-K retrieval")

    # model overrides from add-in
    p.add_argument("--ocr", type=str, default=None, help="OCR (vision) model")
    p.add_argument("--qa", type=str, default=None, help="QA (text) model")
    p.add_argument("--embed", type=str, default=None, help="Embedding model")

    args = p.parse_args()

    cfg = PipelineConfig(
        poppler_path=args.poppler,
        work_dir=args.store,
        top_k=args.topk,
        ocr_model=args.ocr or "qwen3-vl:8b-instruct-q4_K_M",
        qa_model=args.qa or "gemma3:12b-it-q4_K_M",
        embed_model=args.embed or "qwen3-embedding:8b-q4_K_M",
    )

    rag = RAGPipeline(cfg)

    if args.dir:
        rag.ingest_folder(args.dir)
        return

    if args.ask:
        rag.load()
        print(rag.query(args.ask))
        return

    print("Use --dir to ingest or --ask to query.")


if __name__ == "__main__":
    main()
