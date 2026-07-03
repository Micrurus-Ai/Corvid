"""Persistent embeddings index for document Q&A over large folders.

Extract-once + incremental (by mtime/size), semantic retrieval, no file-count cap. The index lives in
%APPDATA%\\AxonIntelligence\\doc_index\\<hash-of-folder>\\ as manifest.json (chunk text + per-file
ranges) and emb.npy (a float32 matrix of chunk embeddings). Re-asking a folder only re-embeds files
that changed since last time, so it stays fast and cheap."""
import os
import json
import hashlib

import numpy as np
from openai import OpenAI

_EMBED_MODEL = os.getenv("ASSISTANT_EMBED_MODEL", "text-embedding-3-small")
_MAX_FILES = 20000
_CHUNK = 1200
_CHUNKS_PER_FILE = 40
_BATCH = 64


def _index_dir(folder):
    base = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), "AxonIntelligence", "doc_index")
    h = hashlib.sha1(os.path.abspath(folder).lower().encode("utf-8")).hexdigest()[:16]
    d = os.path.join(base, h)
    os.makedirs(d, exist_ok=True)
    return d


def _load(folder):
    d = _index_dir(folder)
    try:
        with open(os.path.join(d, "manifest.json"), encoding="utf-8") as f:
            man = json.load(f)
        emb = np.load(os.path.join(d, "emb.npy"))
        return man, emb
    except Exception:
        return None, None


def _save(folder, man, emb):
    d = _index_dir(folder)
    try:
        np.save(os.path.join(d, "emb.npy"), emb.astype("float32"))
        with open(os.path.join(d, "manifest.json"), "w", encoding="utf-8") as f:
            json.dump(man, f)
    except Exception:
        pass


def _chunk_text(text, size=_CHUNK):
    out = []
    for i in range(0, len(text), size):
        c = text[i:i + size].strip()
        if c:
            out.append(c)
    return out


def _embed(client, texts):
    rows = []
    for i in range(0, len(texts), _BATCH):
        batch = texts[i:i + _BATCH]
        r = client.embeddings.create(model=_EMBED_MODEL, input=batch)
        rows.extend(d.embedding for d in r.data)
    return np.array(rows, dtype="float32")


def update_index(folder):
    """Build or incrementally refresh the folder's index. Returns (manifest, emb) or (None, None)."""
    if not folder or not os.path.isdir(folder) or not os.getenv("OPENAI_API_KEY"):
        return None, None
    from axon.knowledge import _read_text, _EXT   # lazy import to avoid a cycle

    old_man, old_emb = _load(folder)
    old_files = (old_man or {}).get("files", {})
    old_chunks = (old_man or {}).get("chunks", [])

    new_chunks = []              # [{"path","text"}]
    new_files = {}              # path -> {mtime,size,start,count}
    reused = []                 # (dest_start, count, src_start)
    to_embed = []               # chunk texts needing embedding
    to_embed_pos = []           # their global row indices in new_chunks

    scanned = 0
    for root, _dirs, files in os.walk(folder):
        for fn in files:
            if os.path.splitext(fn)[1].lower() not in _EXT:
                continue
            path = os.path.join(root, fn)
            try:
                st = os.stat(path)
            except Exception:
                continue
            scanned += 1
            if scanned > _MAX_FILES:
                break
            meta = old_files.get(path)
            start = len(new_chunks)
            if meta and abs(meta.get("mtime", 0) - st.st_mtime) < 1 and meta.get("size") == st.st_size:
                for j in range(meta["count"]):
                    new_chunks.append({"path": path, "text": old_chunks[meta["start"] + j]["text"]})
                reused.append((start, meta["count"], meta["start"]))
                new_files[path] = {"mtime": st.st_mtime, "size": st.st_size, "start": start, "count": meta["count"]}
            else:
                txt = _read_text(path)
                chs = _chunk_text(txt)[:_CHUNKS_PER_FILE]
                for c in chs:
                    to_embed.append(c)
                    to_embed_pos.append(len(new_chunks))
                    new_chunks.append({"path": path, "text": c})
                new_files[path] = {"mtime": st.st_mtime, "size": st.st_size, "start": start, "count": len(chs)}
        if scanned > _MAX_FILES:
            break

    if not new_chunks:
        man = {"folder": folder, "model": _EMBED_MODEL, "files": {}, "chunks": []}
        _save(folder, man, np.zeros((0, 1), dtype="float32"))
        return man, np.zeros((0, 1), dtype="float32")

    client = OpenAI()
    new_rows = _embed(client, to_embed) if to_embed else None
    dim = new_rows.shape[1] if new_rows is not None else (old_emb.shape[1] if old_emb is not None and old_emb.size else 1536)
    emb = np.zeros((len(new_chunks), dim), dtype="float32")
    for dest_start, count, src_start in reused:
        emb[dest_start:dest_start + count] = old_emb[src_start:src_start + count]
    if to_embed_pos:
        emb[to_embed_pos] = new_rows

    man = {"folder": folder, "model": _EMBED_MODEL, "files": new_files, "chunks": new_chunks}
    _save(folder, man, emb)
    return man, emb


def query(folder, question, k=10):
    """Return the top-k (path, chunk_text) most relevant to the question, via semantic similarity."""
    man, emb = update_index(folder)
    if man is None or emb is None or len(man["chunks"]) == 0:
        return []
    client = OpenAI()
    q = np.array(client.embeddings.create(model=_EMBED_MODEL, input=[question]).data[0].embedding, dtype="float32")
    sims = emb.dot(q) / (np.linalg.norm(emb, axis=1) * (np.linalg.norm(q) + 1e-8) + 1e-8)
    top = np.argsort(-sims)[:k]
    return [(man["chunks"][int(i)]["path"], man["chunks"][int(i)]["text"]) for i in top]


def index_stats(folder):
    """Build/refresh the index and report counts (for an explicit 'index my documents')."""
    man, emb = update_index(folder)
    if man is None:
        return None
    return {"files": len(man.get("files", {})), "chunks": len(man.get("chunks", []))}
