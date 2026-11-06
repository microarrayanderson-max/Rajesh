import io
import json
import os
import re
import sqlite3
import time
import datetime
import requests
from concurrent.futures import ThreadPoolExecutor
from flask import Flask, request, jsonify, send_file, render_template, redirect, url_for, flash, g, has_request_context
from flask_login import LoginManager, UserMixin, login_user, logout_user, login_required, current_user
from werkzeug.security import generate_password_hash, check_password_hash

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment
except ImportError:
    Workbook = None
try:
    from docx import Document
    from docx.shared import RGBColor, Pt
    from docx.enum.text import WD_COLOR_INDEX
except ImportError:
    Document = None

app = Flask(__name__)
# Set SECRET_KEY in your environment for anything beyond local testing
app.secret_key = os.environ.get("SECRET_KEY", "LOCAL_DEVELOPMENT_SECRET_KEY_12345")

DB_FILE = "local_users.db"

# FIX 1: the Ensembl REST API lives on rest.ensembl.org, NOT ensembl.org
ENSEMBL_BASES = {
    "GRCh38": "https://rest.ensembl.org",          # current Ensembl
    "GRCh37": "https://grch37.rest.ensembl.org",   # older assembly (grch37.ensembl.org)
}
ENSEMBL_BASE = ENSEMBL_BASES["GRCh38"]
CACHE_DB = "ensembl_cache.db"
HEADERS = {"Content-Type": "application/json", "Accept": "application/json"}
TIMEOUT = 90  # DMD and other huge genes can be slow


# ------------------------------------------------------------------
# Database
# ------------------------------------------------------------------
def get_db():
    return sqlite3.connect(DB_FILE)


def init_db():
    conn = get_db()
    try:
        cursor = conn.cursor()
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS users (
                email TEXT PRIMARY KEY,
                password TEXT NOT NULL,
                security_question TEXT NOT NULL,
                security_answer TEXT NOT NULL
            )
        ''')
        cursor.execute("SELECT 1 FROM users WHERE email='admin@lab.com'")
        if not cursor.fetchone():
            cursor.execute(
                "INSERT INTO users VALUES (?, ?, ?, ?)",
                ("admin@lab.com",
                 generate_password_hash("password123"),
                 "What is your favorite gene?",
                 generate_password_hash("hbb"))
            )
        conn.commit()
    finally:
        conn.close()


init_db()

login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = 'login'


class User(UserMixin):
    def __init__(self, email):
        self.id = email


@login_manager.user_loader
def load_user(user_id):
    conn = get_db()
    try:
        row = conn.execute("SELECT email FROM users WHERE email=?", (user_id,)).fetchone()
    finally:
        conn.close()
    return User(row[0]) if row else None


# ------------------------------------------------------------------
# Auth routes
# ------------------------------------------------------------------
@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        email = request.form.get('email', '').strip()
        password = request.form.get('password', '')

        conn = get_db()
        try:
            row = conn.execute("SELECT password FROM users WHERE email=?", (email,)).fetchone()
        finally:
            conn.close()

        if row and check_password_hash(row[0], password):
            login_user(User(email))
            return redirect(url_for('home'))

        flash("Invalid email or password credentials.", "error")
    return render_template('login.html')


@app.route('/logout')
@login_required
def logout():
    logout_user()
    return redirect(url_for('login'))


@app.route('/forgot-password', methods=['GET', 'POST'])
def forgot_password():
    if request.method == 'POST':
        email = request.form.get('email', '').strip()
        answer = request.form.get('security_answer', '').strip().lower()
        new_password = request.form.get('new_password', '')

        if not new_password:
            flash("New password cannot be empty.", "error")
            return render_template('forgot_password.html')

        conn = get_db()
        try:
            row = conn.execute("SELECT security_answer FROM users WHERE email=?", (email,)).fetchone()
            if row and check_password_hash(row[0], answer):
                conn.execute("UPDATE users SET password=? WHERE email=?",
                             (generate_password_hash(new_password), email))
                conn.commit()
                flash("Password updated successfully! Please login.", "success")
                return redirect(url_for('login'))
        finally:
            conn.close()

        flash("Incorrect Email identifier or security answer verification.", "error")
    return render_template('forgot_password.html')


@app.route('/')
@login_required
def home():
    return render_template('index.html', user=current_user.id)


# ------------------------------------------------------------------
# Ensembl helpers
# ------------------------------------------------------------------
def current_assembly():
    """GRCh38 (default) or GRCh37, chosen with the ?assembly= parameter of the request."""
    if has_request_context() and request.args.get("assembly") == "GRCh37":
        return "GRCh37"
    return "GRCh38"


# ---- local cache: every successful answer is saved and used if Ensembl cannot be reached ----
def _cache_conn():
    conn = sqlite3.connect(CACHE_DB)
    conn.execute("CREATE TABLE IF NOT EXISTS cache (k TEXT PRIMARY KEY, v TEXT NOT NULL, ts TEXT NOT NULL)")
    return conn


def cache_put(key, data):
    try:
        text = json.dumps(data)
        if len(text) > 30_000_000:
            return
        conn = _cache_conn()
        try:
            conn.execute("INSERT OR REPLACE INTO cache VALUES (?, ?, ?)",
                         (key, text, datetime.datetime.utcnow().strftime("%Y-%m-%d %H:%M")))
            conn.commit()
        finally:
            conn.close()
    except Exception:
        pass   # the cache is a bonus; never break a request because of it


def cache_get(key):
    try:
        conn = _cache_conn()
        try:
            row = conn.execute("SELECT v, ts FROM cache WHERE k=?", (key,)).fetchone()
        finally:
            conn.close()
        return (json.loads(row[0]), row[1]) if row else (None, None)
    except Exception:
        return None, None


def _note_cache_use(ts):
    if has_request_context():
        old = getattr(g, "cache_ts", None)
        g.cache_ts = ts if old is None else min(old, ts)


@app.after_request
def _cache_header(resp):
    ts = getattr(g, "cache_ts", None) if has_request_context() else None
    if ts:
        resp.headers["X-Cache-Date"] = ts
    return resp


class EnsemblError(Exception):
    """Raised when Ensembl cannot be reached or returns a server-side problem."""


def _ens_fetch(base, path, params):
    """One network request. Returns JSON, None if not found (400/404); raises EnsemblError otherwise."""
    res = None
    for attempt in range(2):
        try:
            res = requests.get(f"{base}{path}", headers=HEADERS, params=params, timeout=TIMEOUT)
        except requests.Timeout:
            raise EnsemblError("Ensembl took too long to respond (timeout).")
        except requests.RequestException as e:
            raise EnsemblError(
                f"Could not reach Ensembl ({e.__class__.__name__}). "
                "Check your internet connection, firewall or proxy.")
        if res.status_code == 429 and attempt == 0:   # rate limited: wait and retry once
            try:
                wait = float(res.headers.get("Retry-After", 2))
            except ValueError:
                wait = 2
            time.sleep(min(wait, 10))
            continue
        break
    if res.status_code in (400, 404):
        return None
    if res.status_code != 200:
        raise EnsemblError(f"Ensembl returned HTTP {res.status_code}.")
    try:
        return res.json()
    except ValueError:
        raise EnsemblError("Ensembl returned an unreadable response.")


def ens_get(path, params=None, assembly=None):
    """GET from Ensembl REST (GRCh38 or GRCh37). If Ensembl cannot be reached, answers
    saved earlier on this computer are used. Raises EnsemblError if neither works."""
    base = ENSEMBL_BASES[assembly or current_assembly()]
    key = f"{base}{path}|{params!r}"
    try:
        data = _ens_fetch(base, path, params)
    except EnsemblError:
        saved, ts = cache_get(key)
        if saved is not None:
            _note_cache_use(ts)
            return saved
        raise
    if data is not None:
        cache_put(key, data)
    return data


def gene_name_for_transcript(tx):
    """FIX: in a transcript lookup, 'Parent' is the gene ID string, not a dict."""
    parent = tx.get("Parent")
    if isinstance(parent, dict):
        return parent.get("display_name", "Unknown")
    if isinstance(parent, str) and parent:
        try:
            gene = ens_get(f"/lookup/id/{parent}")
        except EnsemblError:
            return parent
        if gene:
            return gene.get("display_name", parent)
        return parent
    return "Unknown"


def summarize_transcript(t):
    return {
        "transcript_id": t.get("id"),
        "biotype": t.get("biotype"),
        "length": t.get("end") - t.get("start") + 1,
        "description": t.get("description", "No description provided"),
        "is_canonical": t.get("is_canonical", 0),
        "name": t.get("display_name"),
        "version": t.get("version"),
        "protein_id": (t.get("Translation") or {}).get("id"),
        "protein_length": (t.get("Translation") or {}).get("length"),
        "exon_count": len(t.get("Exon") or []),
        "intron_count": max(len(t.get("Exon") or []) - 1, 0),
    }


@app.route('/api/search-transcripts', methods=['GET'])
@login_required
def search_transcripts():
    try:
        return _search_transcripts()
    except EnsemblError as e:
        return jsonify({"error": str(e)}), 502


def _search_transcripts():
    user_input = request.args.get('q', '').strip()
    search_type = request.args.get('type', 'auto').lower()

    if not user_input:
        return jsonify({"error": "Query parameter 'q' is required"}), 400

    upper = user_input.upper()
    is_enst = search_type == "transcript" or upper.startswith("ENST")
    is_refseq = search_type == "refseq" or upper.startswith(("NM_", "NR_", "NP_", "XM_", "XR_"))

    # Strip version suffix (.5) only for accession-style input, not gene symbols
    clean_input = user_input.split(".")[0] if (is_enst or is_refseq) else user_input

    gene_data = None

    # 1. Ensembl transcript search
    if is_enst:
        tx = ens_get(f"/lookup/id/{clean_input}", {"expand": 1})
        if tx and tx.get("object_type") == "Transcript":
            return jsonify({
                "gene_name": gene_name_for_transcript(tx),
                "transcripts": [summarize_transcript(tx)]
            })

    # 2. RefSeq route
    # FIX: /xrefs/symbol expects a gene symbol; accession lookup is /xrefs/name
    if is_refseq:
        xref_data = ens_get(f"/xrefs/name/homo_sapiens/{clean_input}")
        if isinstance(xref_data, dict):
            xref_data = [xref_data]
        for xr in (xref_data or []):
            xid, xtype = xr.get("id"), xr.get("type")
            if not xid:
                continue
            if xtype == "gene":
                gene_data = ens_get(f"/lookup/id/{xid}", {"expand": 1})
            elif xtype == "transcript":
                tx = ens_get(f"/lookup/id/{xid}")
                if tx and tx.get("Parent"):
                    gene_data = ens_get(f"/lookup/id/{tx['Parent']}", {"expand": 1})
            if gene_data:
                break

    # 3. Gene symbol fallback
    if not gene_data:
        gene_data = ens_get(f"/lookup/symbol/homo_sapiens/{clean_input}", {"expand": 1})

    if not gene_data or "Transcript" not in gene_data:
        return jsonify({"error": f"Could not find matching records for target: '{user_input}'"}), 404

    return jsonify({
        "gene_name": gene_data.get("display_name", user_input),
        "transcripts": [summarize_transcript(t) for t in gene_data.get("Transcript", [])]
    })


# ------------------------------------------------------------------
# Structure table
# ------------------------------------------------------------------
def build_structural_table(tx_id):
    tx_data = ens_get(f"/lookup/id/{tx_id}", {"expand": 1})
    seq_data = ens_get(f"/sequence/id/{tx_id}", {"type": "genomic"})

    if not tx_data or not seq_data or not seq_data.get("seq"):
        return None, None

    full_seq = seq_data["seq"]
    strand = tx_data.get("strand", 1)
    tx_start = tx_data["start"]
    tx_end = tx_data["end"]

    # Exons in transcript (5'->3') order
    exons = sorted(tx_data.get("Exon", []), key=lambda x: x["start"], reverse=(strand == -1))

    # FIX: Ensembl returns the genomic sequence in the TRANSCRIPT's orientation.
    # For minus-strand genes, offsets must be counted from the transcript END,
    # otherwise exon/intron sequences come from the wrong place.
    def slice_seq(g_start, g_end):
        if strand == 1:
            a, b = g_start - tx_start, g_end - tx_start
        else:
            a, b = tx_end - g_end, tx_end - g_start
        return full_seq[a:b + 1]

    table = []
    for i, exon in enumerate(exons):
        table.append({
            "type": f"Exon {i+1}",
            "id": exon["id"],
            "start": exon["start"],
            "end": exon["end"],
            "start_phase": exon.get("start_phase", "-"),
            "end_phase": exon.get("end_phase", "-"),
            "length": exon["end"] - exon["start"] + 1,
            "sequence": slice_seq(exon["start"], exon["end"]).upper()
        })

        if i < len(exons) - 1:
            nxt = exons[i + 1]
            if strand == 1:
                intron_start, intron_end = exon["end"] + 1, nxt["start"] - 1
            else:
                intron_start, intron_end = nxt["end"] + 1, exon["start"] - 1

            if intron_end < intron_start:  # adjacent/overlapping exons, no intron
                continue

            table.append({
                "type": f"Intron {i+1}-{i+2}",
                "id": f"Intron_{i+1}_{i+2}",
                "start": intron_start,
                "end": intron_end,
                "start_phase": "-",
                "end_phase": "-",
                "length": intron_end - intron_start + 1,
                "sequence": slice_seq(intron_start, intron_end).lower()
            })

    return table, gene_name_for_transcript(tx_data)



def get_refseq_info(tx_id):
    """RefSeq accessions linked to an Ensembl transcript (mRNA NM_/XM_, ncRNA NR_, protein NP_)
    plus MANE status. Never raises: returns whatever could be found."""
    info = {"mrna": [], "ncrna": [], "protein": [], "mane": []}
    try:
        tx = ens_get(f"/lookup/id/{tx_id}", {"mane": 1}) or {}
        for m in (tx.get("MANE") or []):
            info["mane"].append({"type": m.get("type"), "refseq": m.get("refseq_match")})
        xrefs = ens_get(f"/xrefs/id/{tx_id}", {"all_levels": 1}) or []
    except EnsemblError:
        return info
    for x in xrefs:
        db = x.get("dbname", "")
        acc = x.get("display_id") or x.get("primary_id")
        if not acc or not db.startswith("RefSeq"):
            continue
        if "mRNA" in db:
            key = "mrna"
        elif "ncRNA" in db:
            key = "ncrna"
        elif "peptide" in db:
            key = "protein"
        else:
            continue
        if acc not in info[key]:
            info[key].append(acc)
    return info


def refseq_summary(info):
    """One-line text version for exports."""
    parts = []
    ids = info["mrna"] + info["ncrna"]
    parts.append("RefSeq transcript: " + (", ".join(ids) if ids else "none linked"))
    parts.append("RefSeq protein: " + (", ".join(info["protein"]) if info["protein"] else "none linked"))
    if info["mane"]:
        parts.append("MANE: " + ", ".join(f"{m['type']} ({m['refseq']})" for m in info["mane"]))
    return " | ".join(parts)



def refseq_ids_for(tx_id, assembly=None):
    """RefSeq mRNA/ncRNA accessions (NM_/NR_/XM_/XR_) linked to one Ensembl transcript."""
    xrefs = ens_get(f"/xrefs/id/{tx_id}", assembly=assembly) or []
    out = []
    for x in xrefs:
        db = x.get("dbname", "")
        if db.startswith("RefSeq") and ("mRNA" in db or "ncRNA" in db):
            acc = x.get("display_id") or x.get("primary_id")
            if acc and acc not in out:
                out.append(acc)
    return out


@app.route('/api/refseq-batch', methods=['GET'])
@login_required
def refseq_batch():
    """RefSeq IDs for several Ensembl transcripts (used to fill the transcript list)."""
    raw = request.args.get('ids', '')
    ids = [i.strip() for i in raw.split(',') if re.fullmatch(r'ENST\d+', i.strip())][:20]

    asm = current_assembly()

    def worker(tx_id):
        try:
            return tx_id, refseq_ids_for(tx_id, asm)
        except EnsemblError:
            return tx_id, None   # None = lookup failed (shown as n/a)

    with ThreadPoolExecutor(max_workers=4) as pool:
        results = dict(pool.map(worker, ids))
    return jsonify(results)


@app.route('/api/export-selected', methods=['GET'])
@login_required
def export_selected():
    tx_id = request.args.get('transcript_id', '').strip()
    export_format = request.args.get('format', 'json').lower()

    if not tx_id:
        return jsonify({"error": "Parameter 'transcript_id' is required"}), 400

    try:
        table_data, gene_name = build_structural_table(tx_id)
    except EnsemblError as e:
        return jsonify({"error": str(e)}), 502
    if not table_data:
        return jsonify({"error": "Failed to parse the transcript structure"}), 500

    refseq = get_refseq_info(tx_id)

    # ---- Excel ----
    if export_format == 'excel':
        if not Workbook:
            return jsonify({"error": "openpyxl is not installed (pip install openpyxl)"}), 501
        wb = Workbook()
        ws = wb.active
        ws.title = "Structure"
        ws.append(["Element", "Accession ID", "Start", "End", "Start Phase", "End Phase", "Length", "Sequence"])
        for item in table_data:
            ws.append([item["type"], item["id"], item["start"], item["end"],
                       item["start_phase"], item["end_phase"], item["length"], item["sequence"]])

        info_ws = wb.create_sheet("Transcript Info")
        info_ws.append(["Gene", gene_name])
        info_ws.append(["Ensembl transcript", tx_id])
        info_ws.append(["RefSeq mRNA", ", ".join(refseq["mrna"]) or "none linked"])
        info_ws.append(["RefSeq ncRNA", ", ".join(refseq["ncrna"]) or "none linked"])
        info_ws.append(["RefSeq protein", ", ".join(refseq["protein"]) or "none linked"])
        info_ws.append(["MANE", ", ".join(f"{m['type']} ({m['refseq']})" for m in refseq["mane"]) or "-"])

        stream = io.BytesIO()
        wb.save(stream)
        stream.seek(0)
        return send_file(
            stream,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name=f"{tx_id}_structure.xlsx"
        )

    # ---- Word ----
    if export_format == 'word':
        if not Document:
            return jsonify({"error": "python-docx is not installed (pip install python-docx)"}), 501
        doc = Document()
        doc.add_heading(f"Transcript Map: {tx_id} ({gene_name})", level=1)
        doc.add_paragraph(refseq_summary(refseq))

        table = doc.add_table(rows=1, cols=6)
        table.style = "Table Grid"
        # FIX: table.rows is a collection; need the first row -> rows[0]
        hdr_cells = table.rows[0].cells
        for idx, text in enumerate(["Element", "Start", "End", "S-Phase", "Len", "Sequence Segments"]):
            hdr_cells[idx].text = text

        for item in table_data:
            cells = table.add_row().cells
            cells[0].text = str(item["type"])
            cells[1].text = str(item["start"])
            cells[2].text = str(item["end"])
            cells[3].text = str(item["start_phase"])
            cells[4].text = str(item["length"])

            seq = item["sequence"]
            run = cells[5].paragraphs[0].add_run(seq[:50] + ("..." if len(seq) > 50 else ""))
            run.font.name = 'Courier New'
            if item["type"].startswith("Exon"):
                run.font.bold = True
                run.font.color.rgb = RGBColor(0, 51, 102)
            else:
                run.font.color.rgb = RGBColor(128, 128, 128)

        stream = io.BytesIO()
        doc.save(stream)
        stream.seek(0)
        return send_file(
            stream,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            as_attachment=True,
            download_name=f"{tx_id}_report.docx"
        )

    return jsonify({"table_data": table_data, "gene_name": gene_name, "refseq": refseq})


# ------------------------------------------------------------------
# cDNA + protein view
# ------------------------------------------------------------------
_BASES = "TCAG"
_AAS = "FFLLSSSSYY**CC*WLLLLPPPPHHQQRRRRIIIMTTTTNNKKSSRRVVVVAAAADDEEGGGG"
CODON_TABLE = {a + b + c: _AAS[i]
               for i, (a, b, c) in enumerate((x, y, z) for x in _BASES for y in _BASES for z in _BASES)}


def translate_cds(cds):
    """Translate a CDS in frame 1 (standard code). '*' = stop. Unknown codons = 'X'."""
    return "".join(CODON_TABLE.get(cds[i:i + 3], "X") for i in range(0, len(cds) - len(cds) % 3, 3))




# ------------------------------------------------------------------
# Variants on the cDNA (consequence calculated for this transcript)
# ------------------------------------------------------------------
SO_MAP = {
    "missense": ["missense_variant"],
    "synonymous": ["synonymous_variant"],
    "stop_gained": ["stop_gained"],
    "frameshift": ["frameshift_variant"],
    "inframe": ["inframe_insertion", "inframe_deletion"],
    "start_stop_lost": ["start_lost", "stop_lost"],
    "protein_altering": ["protein_altering_variant"],
    "utr": ["5_prime_UTR_variant", "3_prime_UTR_variant"],
}
SEVERITY = ["stop_gained", "frameshift", "start_lost", "stop_lost", "inframe_deletion",
            "inframe_insertion", "protein_altering", "missense", "synonymous",
            "5utr", "3utr", "noncoding"]
_COMP = str.maketrans("ACGT", "TGCA")


def _rank(cls):
    return SEVERITY.index(cls) if cls in SEVERITY else len(SEVERITY)


def compute_variant_marks(records, segs, strand, cdna, cds, cds0, cds1):
    """Turn Ensembl variation records (genomic) into per-cDNA-position marks.
    segs = [(exon_start, exon_end, cdna_offset_before_exon)] in transcript order."""

    def g2c(g):
        for st, en, cum in segs:
            if st <= g <= en:
                return cum + (g - st if strand == 1 else en - g)
        return None

    has_cds = cds is not None and cds0 is not None
    coding = lambda i: has_cds and cds0 <= i <= cds1

    def utr_cls(i):
        if not has_cds:
            return "noncoding"
        return "5utr" if i < cds0 else "3utr"

    by_pos, used = {}, 0
    for v in records:
        alleles = v.get("alleles") or []
        if len(alleles) < 2 or v.get("start") is None or v.get("end") is None:
            continue
        gs, ge = v["start"], v["end"]
        vid = v.get("id")
        is_snv = gs == ge and all(len(a) == 1 and a in "ACGT" for a in alleles)

        if is_snv:
            idx = g2c(gs)
            if idx is None:
                continue
            tx_alleles = [a.translate(_COMP) if strand == -1 else a for a in alleles]
            cls = utr_cls(idx)
            if coding(idx):
                rel = idx - cds0
                ci, p = rel // 3, rel % 3
                codon = cds[ci * 3: ci * 3 + 3]
                best = None
                if len(codon) == 3:
                    ref_aa = CODON_TABLE.get(codon, "X")
                    for alt in tx_alleles:
                        if alt == cdna[idx]:
                            continue
                        new_aa = CODON_TABLE.get(codon[:p] + alt + codon[p + 1:], "X")
                        if ci == 0 and new_aa != "M":
                            c = "start_lost"
                        elif ref_aa != "*" and new_aa == "*":
                            c = "stop_gained"
                        elif ref_aa == "*" and new_aa != "*":
                            c = "stop_lost"
                        elif ref_aa == new_aa:
                            c = "synonymous"
                        else:
                            c = "missense"
                        if best is None or _rank(c) < _rank(best):
                            best = c
                cls = best or "missense"
            m = by_pos.setdefault(idx, {"cls": cls, "alleles": set(), "ids": []})
            m["alleles"] |= set(tx_alleles)
            if _rank(cls) < _rank(m["cls"]):
                m["cls"] = cls
            if vid and vid not in m["ids"]:
                m["ids"].append(vid)
            used += 1
        else:
            a, b = g2c(gs), g2c(ge)
            hits = [x for x in (a, b) if x is not None]
            if not hits:
                continue
            lo, hi = min(hits), max(hits)
            if hi - lo > 300:
                continue
            ref = alleles[0]
            lr = 0 if ref == "-" else len(ref)
            best = None
            for alt in alleles[1:]:
                diff = (0 if alt == "-" else len(alt)) - lr
                if diff % 3 != 0:
                    c = "frameshift"
                elif diff > 0:
                    c = "inframe_insertion"
                elif diff < 0:
                    c = "inframe_deletion"
                else:
                    c = "protein_altering"
                if best is None or _rank(c) < _rank(best):
                    best = c
            for idx in range(lo, hi + 1):
                cls = (best or "protein_altering") if coding(idx) else utr_cls(idx)
                m = by_pos.setdefault(idx, {"cls": cls, "alleles": set(), "ids": []})
                if _rank(cls) < _rank(m["cls"]):
                    m["cls"] = cls
                m["alleles"].add("indel")
                if vid and vid not in m["ids"]:
                    m["ids"].append(vid)
            used += 1

    marks = [{"i": idx, "cls": m["cls"], "ids": m["ids"][:4],
              "alleles": "/".join(sorted(m["alleles"]))}
             for idx, m in sorted(by_pos.items())]
    return marks, used


def build_cdna_view(tx_id, var_groups=None):
    tx = ens_get(f"/lookup/id/{tx_id}", {"expand": 1})
    if not tx or tx.get("object_type") != "Transcript":
        return None

    cdna_json = ens_get(f"/sequence/id/{tx_id}", {"type": "cdna"})
    if not cdna_json or not cdna_json.get("seq"):
        return None
    cdna = cdna_json["seq"].upper()

    # Exon boundaries in cDNA coordinates (transcript 5'->3' order, 1-based inclusive)
    strand = tx.get("strand", 1)
    exons = sorted(tx.get("Exon", []), key=lambda x: x["start"], reverse=(strand == -1))
    exon_map, segs, pos = [], [], 0
    for i, ex in enumerate(exons):
        length = ex["end"] - ex["start"] + 1
        exon_map.append({"n": i + 1, "id": ex["id"], "start": pos + 1, "end": pos + length})
        segs.append((ex["start"], ex["end"], pos))
        pos += length
    cds = None

    result = {
        "transcript_id": tx_id,
        "gene_name": gene_name_for_transcript(tx),
        "biotype": tx.get("biotype"),
        "cdna": cdna,
        "length": len(cdna),
        "exons": exon_map,
        "exon_total_matches": pos == len(cdna),
        "version": tx.get("version"),
        "strand": strand,
        "chrom": tx.get("seq_region_name"),
        "segs": [list(x) for x in segs],
        "exon_count": len(exon_map),
        "intron_count": max(len(exon_map) - 1, 0),
        "cds_start": None,
        "cds_end": None,
        "protein": "",
        "protein_id": None,
        "variants": [],
        "variant_count": 0,
        "variant_error": None,
    }

    # Coding sequence (only for protein-coding transcripts)
    if tx.get("Translation"):
        result["protein_id"] = tx["Translation"].get("id")
        cds_json = ens_get(f"/sequence/id/{tx_id}", {"type": "cds"})
        if cds_json and cds_json.get("seq"):
            cds = cds_json["seq"].upper()
            idx = cdna.find(cds)
            if idx >= 0:
                result["cds_start"] = idx + 1
                result["cds_end"] = idx + len(cds)
                result["protein"] = translate_cds(cds)
            else:
                cds = None

    # Variants (optional)
    terms = [t for g in (var_groups or []) for t in SO_MAP.get(g, [])]
    if terms:
        params = [("feature", "variation")] + [("so_term", t) for t in terms]
        try:
            records = ens_get(f"/overlap/id/{tx_id}", params) or []
            cds0 = result["cds_start"] - 1 if result["cds_start"] else None
            cds1 = result["cds_end"] - 1 if result["cds_end"] else None
            marks, used = compute_variant_marks(records, segs, strand, cdna, cds, cds0, cds1)
            result["variants"], result["variant_count"] = marks, used
        except EnsemblError as e:
            result["variant_error"] = str(e)

    return result


@app.route('/api/cdna-view', methods=['GET'])
@login_required
def cdna_view():
    tx_id = request.args.get('transcript_id', '').strip()
    if not tx_id:
        return jsonify({"error": "Parameter 'transcript_id' is required"}), 400
    groups = [g for g in request.args.get('variants', '').split(',') if g in SO_MAP]
    try:
        data = build_cdna_view(tx_id, groups)
    except EnsemblError as e:
        return jsonify({"error": str(e)}), 502
    if not data:
        return jsonify({"error": "Could not retrieve the cDNA for this transcript"}), 404
    return jsonify(data)


# ------------------------------------------------------------------
# Locate a position (gene + position + nucleotide) in the transcript
# ------------------------------------------------------------------
AA3 = {"A": "Ala", "R": "Arg", "N": "Asn", "D": "Asp", "C": "Cys", "Q": "Gln", "E": "Glu",
       "G": "Gly", "H": "His", "I": "Ile", "L": "Leu", "K": "Lys", "M": "Met", "F": "Phe",
       "P": "Pro", "S": "Ser", "T": "Thr", "W": "Trp", "Y": "Tyr", "V": "Val", "*": "Ter", "X": "Xaa"}


def resolve_transcript(gene, tx_in):
    """Ensembl transcript ID to use: the one given, otherwise the gene's canonical transcript."""
    if tx_in:
        if re.fullmatch(r"N[MR]_\d+", tx_in):          # RefSeq accession -> linked Ensembl transcript
            for x in (ens_get(f"/xrefs/name/homo_sapiens/{tx_in}") or []):
                if x.get("type") == "transcript" and x.get("id"):
                    return x["id"]
            raise ValueError(f"Ensembl has no transcript linked to {tx_in}.")
        if not re.fullmatch(r"ENST\d+", tx_in):
            raise ValueError("Enter an Ensembl transcript ID (ENST...) or a RefSeq ID (NM_...), or leave it empty.")
        return tx_in
    if not re.fullmatch(r"[A-Za-z0-9._-]+", gene or ""):
        raise ValueError("Enter a valid gene symbol, for example DMD.")
    gd = ens_get(f"/lookup/symbol/homo_sapiens/{gene}", {"expand": 1})
    if not gd or not gd.get("Transcript"):
        raise ValueError(f"Gene '{gene}' was not found in Ensembl.")
    trs = gd["Transcript"]
    canon = [t for t in trs if t.get("is_canonical") == 1]
    if canon:
        return canon[0]["id"]
    pool = [t for t in trs if t.get("Translation")] or trs
    best = max(pool, key=lambda x: (x.get("Translation") or {}).get("length", x["end"] - x["start"]))
    return best["id"]


def _parse_position(pos_str, coord):
    s = (pos_str or "").strip().replace(",", "").replace(" ", "")
    if coord == "g":
        if not re.fullmatch(r"\d+", s):
            raise ValueError("A genomic position must be a plain number, for example 32389644.")
        return {"g": int(s)}
    m = re.fullmatch(r"([-*])?(\d+)(?:([+-])(\d+))?", s)
    if not m:
        raise ValueError("Position not recognised. Examples: 1234, 1234+5, 1234-2, -20, *50")
    pre, num, sign, off = m.groups()
    if coord == "n" and pre:
        raise ValueError("For transcript (n.) positions use a plain number, for example 250.")
    if int(num) < 1:
        raise ValueError("Positions start at 1.")
    return {"pre": pre, "num": int(num), "sign": sign, "off": int(off) if off else 0}


def _classify_snv(idx, alt, cdna, cds, cds0):
    """Consequence of changing cDNA base idx (inside the CDS) to alt."""
    rel = idx - cds0
    ci, p = rel // 3, rel % 3
    codon = cds[ci * 3: ci * 3 + 3]
    if len(codon) < 3:
        return None
    new = codon[:p] + alt + codon[p + 1:]
    ref_aa, new_aa = CODON_TABLE.get(codon, "X"), CODON_TABLE.get(new, "X")
    if ci == 0 and new_aa != "M":
        cls, pn = "start_lost", "p.Met1?"
    elif ref_aa != "*" and new_aa == "*":
        cls, pn = "stop_gained", f"p.{AA3[ref_aa]}{ci + 1}Ter"
    elif ref_aa == "*" and new_aa != "*":
        cls, pn = "stop_lost", f"p.Ter{ci + 1}{AA3[new_aa]}ext*?"
    elif ref_aa == new_aa:
        cls, pn = "synonymous", f"p.{AA3[ref_aa]}{ci + 1}="
    else:
        cls, pn = "missense", f"p.{AA3[ref_aa]}{ci + 1}{AA3[new_aa]}"
    return {"cls": cls, "p": pn, "alt_codon": new, "alt_aa": new_aa}


def locate_in_transcript(view, coord, pos_str, ref, alt):
    """Place a position in the transcript. `view` comes from Ensembl (has genomic info) or from the
    NCBI fallback (view['segs'] is None: exon positions in the cDNA only)."""
    cdna, n = view["cdna"], len(view["cdna"])
    strand, chrom = view.get("strand"), view.get("chrom")
    segs = view.get("segs")                       # [[gstart, gend, cdna_offset], ...] or None
    exons = view["exons"]
    n_ex = len(exons)
    if n_ex == 0:
        raise ValueError("The exon layout is not available for this transcript, so the position cannot be placed in an exon.")
    rng = [(e["start"] - 1, e["end"] - 1) for e in exons]     # 0-based, inclusive, cDNA
    cds0 = view["cds_start"] - 1 if view["cds_start"] else None
    cds1 = view["cds_end"] - 1 if view["cds_end"] else None
    cds = cdna[cds0:cds1 + 1] if cds0 is not None else None
    comp = lambda b: b.translate(_COMP)
    p = _parse_position(pos_str, coord)
    if coord == "g" and not segs:
        raise ValueError("Genomic (g.) positions need Ensembl data, which could not be reached. Use c. or n. instead.")

    def idx_to_g(i):
        if not segs:
            return None
        for st, en, cum in segs:
            if cum <= i < cum + (en - st + 1):
                o = i - cum
                return st + o if strand == 1 else en - o
        return None

    def g_to_idx(g):
        for st, en, cum in segs:
            if st <= g <= en:
                return cum + (g - st if strand == 1 else en - g)
        return None

    def exon_of(i):
        for k, (a, b) in enumerate(rng):
            if a <= i <= b:
                return k
        return None

    def intron_bounds(k):      # genomic lo..hi of the intron after exon k (0-based)
        a, b = segs[k], segs[k + 1]
        return (a[1] + 1, b[0] - 1) if strand == 1 else (b[1] + 1, a[0] - 1)

    def c_notation(i):
        if cds0 is None:
            return None
        if i < cds0:
            return f"c.-{cds0 - i}"
        if i > cds1:
            return f"c.*{i - cds1}"
        return f"c.{i - cds0 + 1}"

    # nucleotides given with g. are forward-strand -> convert to transcript orientation
    ref_t = comp(ref) if (coord == "g" and strand == -1 and ref) else ref
    alt_t = comp(alt) if (coord == "g" and strand == -1 and alt) else alt

    idx, intr = None, None
    if coord == "g":
        g = p["g"]
        idx = g_to_idx(g)
        if idx is None:
            for k in range(n_ex - 1):
                lo, hi = intron_bounds(k)
                if lo <= g <= hi:
                    d_don, d_acc = ((g - segs[k][1], segs[k + 1][0] - g) if strand == 1
                                    else (segs[k][0] - g, g - segs[k + 1][1]))
                    side, dist = ("donor", d_don) if d_don <= d_acc else ("acceptor", d_acc)
                    intr = {"k": k, "side": side, "dist": dist, "g": g}
                    break
            if intr is None:
                raise ValueError("That genomic position is outside the exons and introns of this transcript.")
    else:
        if coord == "c":
            if cds0 is None:
                raise ValueError("This transcript has no coding sequence. Use the n. (transcript position) option.")
            base = (cds0 - p["num"]) if p["pre"] == "-" else (cds1 + p["num"]) if p["pre"] == "*" else cds0 + p["num"] - 1
        else:
            base = p["num"] - 1
        if not 0 <= base < n:
            raise ValueError(f"Position is outside the cDNA (this transcript is {n} bp long).")
        if p["off"]:
            k = exon_of(base)
            a, b = rng[k]
            g = None
            if p["sign"] == "+":
                if base != b:
                    raise ValueError("A '+' offset must follow the last base of an exon.")
                if k == n_ex - 1:
                    raise ValueError("The last exon has no intron after it.")
                kk, side = k, "donor"
                if segs:
                    g = segs[k][1] + p["off"] if strand == 1 else segs[k][0] - p["off"]
            else:
                if base != a:
                    raise ValueError("A '-' offset must precede the first base of an exon.")
                if k == 0:
                    raise ValueError("The first exon has no intron before it.")
                kk, side = k - 1, "acceptor"
                if segs:
                    g = segs[k][0] - p["off"] if strand == 1 else segs[k][1] + p["off"]
            if segs:
                lo, hi = intron_bounds(kk)
                if p["off"] > hi - lo + 1:
                    raise ValueError(f"That offset is longer than the intron ({hi - lo + 1} bp).")
            intr = {"k": kk, "side": side, "dist": p["off"], "g": g}
        else:
            idx = base

    rows, warnings, marks = [], [], []
    hit_exon = hit_intron = None
    strand_txt = "+" if strand == 1 else "\u2212"

    if intr is None:                                   # ---------- exonic ----------
        k = exon_of(idx)
        hit_exon = k
        a, b = rng[k]
        ln = b - a + 1
        pos_in_exon = idx - a + 1
        base = cdna[idx]
        g = idx_to_g(idx)
        if cds0 is None:
            region = "Exon of a non-coding transcript"
        elif idx < cds0:
            region = "5\u2032 UTR"
        elif idx > cds1:
            region = "3\u2032 UTR"
        else:
            region = "Coding sequence (CDS)"
        cn = c_notation(idx)
        rows += [
            ["Location", f"Exon {k + 1} of {n_ex}"],
            ["Position in exon", f"{pos_in_exon} of {ln} bp (exon covers cDNA {a + 1}\u2013{b + 1})"],
        ]
        if g is not None:
            rows.append(["Genomic position", f"{chrom}:{g} (strand {strand_txt})"])
        rows += [
            ["Region", region],
            ["cDNA notation", f"n.{idx + 1}" + (f"  |  {cn}" if cn else "")],
            ["Reference base (transcript)", base],
        ]
        if ref_t and ref_t != base:
            warnings.append(f"The reference nucleotide you entered ({ref}) does not match the transcript base ({base}) at this position.")
        if alt_t and alt_t == base:
            warnings.append("The alternate nucleotide is the same as the reference base.")
        if pos_in_exon <= 3 and k > 0:
            warnings.append(f"Within 3 bp of the splice acceptor (boundary with intron {k}); may affect splicing.")
        if ln - pos_in_exon + 1 <= 3 and k < n_ex - 1:
            warnings.append(f"Within 3 bp of the splice donor (boundary with intron {k + 1}); may affect splicing.")

        cls = "query"
        if cds0 is not None and cds0 <= idx <= cds1:
            rel = idx - cds0
            ci, cp = rel // 3, rel % 3
            codon = cds[ci * 3: ci * 3 + 3]
            ref_aa = CODON_TABLE.get(codon, "X")
            rows.append(["Codon", f"codon {ci + 1}, base {cp + 1} of 3: {codon} = {AA3.get(ref_aa, ref_aa)} ({ref_aa})"])
            if alt_t and alt_t != base:
                r = _classify_snv(idx, alt_t, cdna, cds, cds0)
                if r:
                    cls = r["cls"]
                    rows.append(["Change", f"{codon} \u2192 {r['alt_codon']} : {AA3.get(ref_aa, ref_aa)} \u2192 {AA3.get(r['alt_aa'], r['alt_aa'])}"])
                    rows.append(["Consequence", f"{r['cls'].replace('_', ' ')}  |  {r['p']}"])
        elif alt_t and alt_t != base:
            cls = "noncoding" if cds0 is None else ("5utr" if idx < cds0 else "3utr")
            rows.append(["Consequence", {"noncoding": "non-coding transcript exon variant",
                                         "5utr": "5 prime UTR variant", "3utr": "3 prime UTR variant"}[cls]])
        marks.append({"i": idx, "cls": cls, "ids": ["your position"],
                      "alleles": f"{base}/{alt_t}" if alt_t else base})

    else:                                              # ---------- intronic ----------
        k, side, dist, g = intr["k"], intr["side"], intr["dist"], intr["g"]
        hit_intron = k
        if side == "donor":
            anchor = rng[k][1]
            where = f"{dist} bp after the end of exon {k + 1} (donor side)"
            sign = "+"
        else:
            anchor = rng[k + 1][0]
            where = f"{dist} bp before the start of exon {k + 2} (acceptor side)"
            sign = "-"
        cn = c_notation(anchor)
        rows += [
            ["Location", f"Intron {k + 1} of {n_ex - 1} (between exon {k + 1} and exon {k + 2})"],
            ["Position in intron", where],
        ]
        if segs:
            lo, hi = intron_bounds(k)
            rows.append(["Intron size", f"{hi - lo + 1:,} bp (genomic {lo}\u2013{hi})"])
        if g is not None:
            rows.append(["Genomic position", f"{chrom}:{g} (strand {strand_txt})"])
        rows.append(["cDNA notation", f"n.{anchor + 1}{sign}{dist}" + (f"  |  {cn}{sign}{dist}" if cn else "")])
        if dist <= 2:
            cons = "splice donor variant" if side == "donor" else "splice acceptor variant"
        elif dist <= 8:
            cons = "splice region variant (3\u20138 bp into the intron)"
        else:
            cons = "intron variant"
        rows.append(["Consequence", cons])
        if segs and chrom and g is not None:
            try:   # reference base in the intron (forward strand -> transcript orientation)
                seq = ens_get(f"/sequence/region/human/{chrom}:{g}..{g}:1")
                fwd = (seq or {}).get("seq", "").upper()
                if fwd:
                    tb = comp(fwd) if strand == -1 else fwd
                    rows.append(["Reference base (transcript)", tb])
                    if ref_t and ref_t != tb:
                        warnings.append(f"The reference nucleotide you entered ({ref}) does not match the base at this position ({tb}).")
                    if alt_t and alt_t == tb:
                        warnings.append("The alternate nucleotide is the same as the reference base.")
            except EnsemblError:
                pass
        else:
            warnings.append("Intron size, genomic position and the intron base are not available from the RefSeq record.")
        warnings.append("Intronic positions are not part of the cDNA, so nothing is highlighted in the cDNA sequence.")

    structure = []
    for k, (a, b) in enumerate(rng):
        gtxt = f"{min(segs[k][0], segs[k][1])}\u2013{max(segs[k][0], segs[k][1])}" if segs else "\u2013"
        structure.append({"type": f"Exon {k + 1}", "id": exons[k]["id"], "cdna": f"{a + 1}\u2013{b + 1}",
                          "g": gtxt, "length": b - a + 1, "hit": hit_exon == k})
        if k < n_ex - 1:
            if segs:
                lo, hi = intron_bounds(k)
                itxt, ilen = f"{lo}\u2013{hi}", max(hi - lo + 1, 0)
            else:
                itxt, ilen = "\u2013", "n/a"
            structure.append({"type": f"Intron {k + 1}", "id": f"exon {k + 1} / exon {k + 2}", "cdna": "\u2013",
                              "g": itxt, "length": ilen, "hit": hit_intron == k})
    return {"rows": rows, "warnings": warnings, "marks": marks,
            "highlight": idx, "structure": structure}


# ------------------------------------------------------------------
# NCBI RefSeq fallback (used only when Ensembl cannot be reached)
# ------------------------------------------------------------------
class NcbiError(Exception):
    """NCBI could not be reached or had no suitable record."""


NCBI_EUTILS = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils"


def _ncbi_get(path, params):
    key = f"ncbi|{path}|{sorted(params.items())!r}"
    err = None
    try:
        res = requests.get(NCBI_EUTILS + path, params=params, timeout=TIMEOUT)
        if res.status_code == 200:
            cache_put(key, {"text": res.text})
            return res.text
        err = NcbiError(f"NCBI returned HTTP {res.status_code}.")
    except requests.Timeout:
        err = NcbiError("NCBI took too long to respond (timeout).")
    except requests.RequestException as e:
        err = NcbiError(f"Could not reach NCBI ({e.__class__.__name__}).")
    saved, ts = cache_get(key)
    if saved is not None:
        _note_cache_use(ts)
        return saved["text"]
    raise err


def _spans(loc):
    return [(int(a), int(b)) for a, b in re.findall(r"<?(\d+)\.\.>?(\d+)", loc or "")]


def parse_genbank(text):
    """Very small GenBank flat-file reader: version, keywords, features and sequence."""
    records = []
    for block in re.split(r"\n//[ \t]*\n?", text):
        if "\nFEATURES" not in block or "\nORIGIN" not in block:
            continue
        rec = {"version": None, "keywords": "", "features": [], "seq": ""}
        m = re.search(r"^VERSION\s+(\S+)", block, re.M)
        rec["version"] = m.group(1) if m else None
        m = re.search(r"^KEYWORDS\s+(.*?)(?=^\S)", block, re.M | re.S)
        rec["keywords"] = " ".join(m.group(1).split()) if m else ""
        feat_txt = block.split("\nFEATURES", 1)[1].split("\nORIGIN", 1)[0]
        cur, qual = None, None
        for line in feat_txt.split("\n")[1:]:
            if not line.strip():
                continue
            if len(line) > 5 and line[:5] == "     " and line[5] != " ":     # new feature key
                cur = {"type": line[5:21].strip(), "loc": line[21:].strip(), "q": {}}
                rec["features"].append(cur)
                qual = None
            elif cur is not None:
                body = line.strip()
                if body.startswith("/"):
                    name, _, val = body[1:].partition("=")
                    cur["q"][name] = val.strip().strip('"')
                    qual = name
                elif qual is not None:
                    cur["q"][qual] += body.strip('"')
                else:
                    cur["loc"] += body
        rec["seq"] = re.sub(r"[^A-Za-z]", "", block.split("\nORIGIN", 1)[1]).upper()
        records.append(rec)
    return records


def ncbi_fallback_view(gene, tx_in):
    """Build a cDNA view from an NCBI RefSeq mRNA record (no introns / genomic coordinates)."""
    if tx_in and re.fullmatch(r"N[MR]_\d+", tx_in):
        text = _ncbi_get("/efetch.fcgi", {"db": "nuccore", "id": tx_in, "rettype": "gb", "retmode": "text"})
    else:
        if not gene or not re.fullmatch(r"[A-Za-z0-9._-]+", gene):
            raise NcbiError("Enter a gene name (or a RefSeq ID such as NM_000546) so NCBI can be used instead.")
        term = f'{gene}[Gene Name] AND "Homo sapiens"[Organism] AND srcdb_refseq_known[Properties]'
        found = json.loads(_ncbi_get("/esearch.fcgi", {"db": "nuccore", "term": term, "retmax": 40, "retmode": "json"}))
        ids = found.get("esearchresult", {}).get("idlist", [])
        if not ids:
            raise NcbiError(f"NCBI has no human RefSeq transcript for '{gene}'.")
        text = _ncbi_get("/efetch.fcgi", {"db": "nuccore", "id": ",".join(ids), "rettype": "gb", "retmode": "text"})
    recs = [r for r in parse_genbank(text) if r["seq"]]
    if not recs:
        raise NcbiError("NCBI returned no usable transcript record.")

    def cds_len(r):
        c = next((f for f in r["features"] if f["type"] == "CDS"), None)
        return sum(b - a + 1 for a, b in _spans(c["loc"])) if c else 0

    rec = max(recs, key=lambda r: ("MANE Select" in r["keywords"], "RefSeq Select" in r["keywords"], cds_len(r)))
    cdna = rec["seq"]
    cds_f = next((f for f in rec["features"] if f["type"] == "CDS"), None)
    gene_f = next((f for f in rec["features"] if f["type"] == "gene"), None)
    cds_start = cds_end = None
    protein, protein_id = "", None
    if cds_f and _spans(cds_f["loc"]):
        sp = _spans(cds_f["loc"])
        cds_start, cds_end = min(a for a, _ in sp), max(b for _, b in sp)
        protein = translate_cds(cdna[cds_start - 1:cds_end])
        protein_id = cds_f["q"].get("protein_id")
    ex_spans = sorted({_spans(f["loc"])[0] for f in rec["features"] if f["type"] == "exon" and _spans(f["loc"])})
    exon_map = [{"n": i + 1, "id": f"exon {i + 1}", "start": a, "end": b} for i, (a, b) in enumerate(ex_spans)]
    acc, _, ver = (rec["version"] or "").partition(".")
    return {
        "transcript_id": acc, "version": int(ver) if ver.isdigit() else None,
        "gene_name": (gene_f["q"].get("gene") if gene_f else None) or gene or "",
        "biotype": "RefSeq transcript", "cdna": cdna, "length": len(cdna),
        "exons": exon_map, "exon_total_matches": True,
        "strand": None, "chrom": None, "segs": None,
        "exon_count": len(exon_map), "intron_count": max(len(exon_map) - 1, 0),
        "cds_start": cds_start, "cds_end": cds_end, "protein": protein, "protein_id": protein_id,
        "variants": [], "variant_count": 0, "variant_error": None,
        "source": "NCBI RefSeq",
    }


def get_view_with_fallback(gene, tx_in):
    """Ensembl first (GRCh38/GRCh37, then saved answers); NCBI RefSeq if Ensembl cannot be reached."""
    try:
        tx_id = resolve_transcript(gene, tx_in)
        view = build_cdna_view(tx_id)
        if not view:
            raise ValueError("Could not retrieve this transcript from Ensembl.")
        view["source"] = "Ensembl " + current_assembly()
        view["source_note"] = None
        return view
    except EnsemblError as e:
        ens_reason = str(e)
    try:
        view = ncbi_fallback_view(gene, tx_in)
    except NcbiError as e2:
        raise EnsemblError(f"{ens_reason} The NCBI RefSeq fallback also failed: {e2}")
    view["source_note"] = (f"Ensembl could not be reached ({ens_reason}) so NCBI RefSeq record "
                           f"{view['transcript_id']}.{view['version']} is shown instead: cDNA, CDS, protein and "
                           "exon positions only (no intron sequences or genomic coordinates).")
    return view


@app.route('/api/locate-variant', methods=['GET'])
@login_required
def locate_variant():
    gene = request.args.get('gene', '').strip()
    tx_in = request.args.get('transcript', '').strip().split('.')[0].upper()
    coord = request.args.get('coord', 'c').lower()
    pos = request.args.get('pos', '')
    ref = request.args.get('ref', '').strip().upper()
    alt = request.args.get('alt', '').strip().upper()

    if coord not in ("c", "n", "g"):
        return jsonify({"error": "Position type must be c., n. or g."}), 400
    if not gene and not tx_in:
        return jsonify({"error": "Enter a gene name (or an Ensembl transcript ID)."}), 400
    if not pos.strip():
        return jsonify({"error": "Enter the position."}), 400
    for label, nt in (("Reference", ref), ("Alternate", alt)):
        if nt and not re.fullmatch(r"[ACGT]", nt):
            return jsonify({"error": f"{label} nucleotide must be a single A, C, G or T."}), 400

    try:
        view = get_view_with_fallback(gene, tx_in)
        res = locate_in_transcript(view, coord, pos, ref, alt)
    except ValueError as e:
        return jsonify({"error": str(e)}), 400
    except EnsemblError as e:
        return jsonify({"error": str(e)}), 502

    view["variants"] = res["marks"]
    view["variant_count"] = len(res["marks"])
    view["highlight"] = res["highlight"]
    return jsonify({
        "transcript": {"id": view["transcript_id"], "version": view.get("version"), "gene": view.get("gene_name"),
                       "biotype": view.get("biotype"), "source": view.get("source"),
                       "canonical_used": (not tx_in) and view.get("source", "").startswith("Ensembl")},
        "location": res["rows"], "warnings": res["warnings"],
        "structure": res["structure"], "view": view,
    })


# ------------------------------------------------------------------
# Excel / Word download: cDNA + exon / intron sequences
# ------------------------------------------------------------------
def _fit_cell(seq, limit=32000):
    """Excel cells hold at most 32,767 characters."""
    if len(seq) <= limit:
        return seq
    half = (limit - 80) // 2
    return f"{seq[:half]} ...[{len(seq) - 2 * half:,} bp not shown: Excel cell limit]... {seq[-half:]}"


def _protein_line(view, s, e):
    cds0 = view["cds_start"] - 1 if view["cds_start"] else -1
    cds1 = view["cds_end"] - 1 if view["cds_end"] else -1
    out = []
    for i in range(s, e):
        if cds0 < 0:
            out.append(" ")
        elif i < cds0 or i > cds1:
            out.append(".")
        else:
            rel = i - cds0
            out.append((view["protein"][rel // 3] if rel // 3 < len(view["protein"]) else "-") if rel % 3 == 1 else "-")
    return "".join(out)


def build_export_data(view):
    """Rows for every exon and intron with sequences. Introns need Ensembl genomic data."""
    cdna, exons, segs = view["cdna"], view["exons"], view.get("segs")
    table, note = None, None
    if segs:
        try:
            table, _ = build_structural_table(view["transcript_id"])
        except EnsemblError as e:
            note = f"Intron sequences could not be retrieved ({e})."
        if table is None and note is None:
            note = "Intron sequences could not be retrieved from Ensembl."
    else:
        note = "Intron sequences and genomic coordinates are not available from the NCBI RefSeq record."
    t_ex = [r for r in (table or []) if r["type"].startswith("Exon")]
    t_in = [r for r in (table or []) if r["type"].startswith("Intron")]

    rows = []
    for k, e in enumerate(exons):
        te = t_ex[k] if k < len(t_ex) else None
        gs = gn = None
        if segs:
            gs, gn = min(segs[k][0], segs[k][1]), max(segs[k][0], segs[k][1])
        rows.append({"type": f"Exon {e['n']}", "id": e["id"], "gstart": gs, "gend": gn,
                     "cstart": e["start"], "cend": e["end"], "length": e["end"] - e["start"] + 1,
                     "sphase": te["start_phase"] if te else "", "ephase": te["end_phase"] if te else "",
                     "sequence": cdna[e["start"] - 1:e["end"]]})
        if k < len(exons) - 1:
            ti = t_in[k] if k < len(t_in) else None
            rows.append({"type": f"Intron {k + 1}", "id": f"exon {k + 1} / exon {k + 2}",
                         "gstart": min(ti["start"], ti["end"]) if ti else None,
                         "gend": max(ti["start"], ti["end"]) if ti else None,
                         "cstart": None, "cend": None, "length": ti["length"] if ti else None,
                         "sphase": "", "ephase": "", "sequence": ti["sequence"] if ti else ""})
    return rows, note


def _summary_rows(view, refseq, note):
    ens_id = view["transcript_id"] + (f".{view['version']}" if view.get("version") else "")
    rs = []
    if refseq:
        rs = [("RefSeq mRNA", ", ".join(refseq.get("mrna", [])) or "none linked"),
              ("RefSeq ncRNA", ", ".join(refseq.get("ncrna", [])) or "none linked"),
              ("RefSeq protein", ", ".join(refseq.get("protein", [])) or "none linked")]
        if refseq.get("mane"):
            rs.append(("MANE", ", ".join(f"{m['type']} ({m['refseq']})" for m in refseq["mane"])))
    rows = [("Transcript", ens_id), ("Gene", view.get("gene_name") or ""), ("Data source", view.get("source", ""))]
    if view.get("strand"):
        rows.append(("Chromosome / strand", f"{view.get('chrom')} / {'+' if view['strand'] == 1 else '-'}"))
    rows += rs
    rows += [("Exons", view["exon_count"]), ("Introns", view["intron_count"]), ("cDNA length (bp)", len(view["cdna"]))]
    if view.get("cds_start"):
        rows.append(("CDS (cDNA positions)", f"{view['cds_start']}\u2013{view['cds_end']}"))
        rows.append(("Protein", f"{view.get('protein_id') or ''}  {len(view['protein'].replace('*', ''))} aa"))
    if view.get("source_note"):
        rows.append(("Note", view["source_note"]))
    if note:
        rows.append(("Intron note", note))
    rows.append(("Created (UTC)", datetime.datetime.utcnow().strftime("%Y-%m-%d %H:%M")))
    return rows


def make_workbook(view, elements, loc_rows, warnings, refseq, note):
    wb = Workbook()
    bold = Font(bold=True)
    mono = Font(name="Courier New", size=10)
    wrap = Alignment(wrap_text=True, vertical="top")

    ws = wb.active
    ws.title = "Summary"
    for k, v in _summary_rows(view, refseq, note):
        ws.append([k, v])
        ws.cell(row=ws.max_row, column=1).font = bold
        ws.cell(row=ws.max_row, column=2).alignment = wrap
    if any(len(e["sequence"]) > 32000 for e in elements):
        ws.append(["Note", "Some intron sequences are longer than an Excel cell allows; the middle part is not shown."])
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 90

    if loc_rows:
        w2 = wb.create_sheet("Your position")
        for k, v in loc_rows:
            w2.append([k, v])
            w2.cell(row=w2.max_row, column=1).font = bold
        for wmsg in warnings:
            w2.append(["Warning", wmsg])
        w2.column_dimensions["A"].width = 30
        w2.column_dimensions["B"].width = 90

    w3 = wb.create_sheet("Exon-Intron")
    w3.append(["Element", "ID", "Genomic start", "Genomic end", "cDNA start", "cDNA end", "Length (bp)",
               "Start phase", "End phase", "Sequence (exons UPPERCASE, introns lowercase)"])
    for c in w3[1]:
        c.font = bold
    for r in elements:
        w3.append([r["type"], r["id"], r["gstart"], r["gend"], r["cstart"], r["cend"], r["length"],
                   r["sphase"], r["ephase"], _fit_cell(r["sequence"])])
        w3.cell(row=w3.max_row, column=10).font = mono
    for col, wd in zip("ABCDEFGHI", (12, 22, 14, 14, 11, 11, 12, 11, 10)):
        w3.column_dimensions[col].width = wd
    w3.column_dimensions["J"].width = 100
    w3.freeze_panes = "A2"

    w4 = wb.create_sheet("cDNA")
    w4.append(["cDNA position", "Sequence (60 bp per row; protein row below: letter under middle base of each codon, . = UTR)", "End"])
    for c in w4[1]:
        c.font = bold
    cdna, n = view["cdna"], len(view["cdna"])
    for s0 in range(0, n, 60):
        e0 = min(s0 + 60, n)
        w4.append([s0 + 1, cdna[s0:e0], e0])
        w4.append(["protein", _protein_line(view, s0, e0), ""])
        for r in (w4.max_row - 1, w4.max_row):
            w4.cell(row=r, column=2).font = mono
    w4.column_dimensions["A"].width = 14
    w4.column_dimensions["B"].width = 90
    w4.freeze_panes = "A2"

    w5 = wb.create_sheet("Full sequences")
    w5.append(["cDNA (full)", _fit_cell(cdna)])
    if view.get("cds_start"):
        w5.append(["CDS", _fit_cell(cdna[view["cds_start"] - 1:view["cds_end"]])])
        w5.append(["Protein", view["protein"].replace("*", "")])
    for r in range(1, w5.max_row + 1):
        w5.cell(row=r, column=1).font = bold
        w5.cell(row=r, column=2).font = mono
    w5.column_dimensions["A"].width = 14
    w5.column_dimensions["B"].width = 120
    return wb


def make_word(view, elements, loc_rows, warnings, refseq, note):
    doc = Document()
    ens_id = view["transcript_id"] + (f".{view['version']}" if view.get("version") else "")
    doc.add_heading(f"{ens_id}  ({view.get('gene_name') or ''})", level=1)

    def kv_table(pairs):
        t = doc.add_table(rows=0, cols=2)
        t.style = "Table Grid"
        for k, v in pairs:
            c = t.add_row().cells
            c[0].text = str(k)
            c[1].text = str(v)
            for run in c[0].paragraphs[0].runs:
                run.font.bold = True

    kv_table(_summary_rows(view, refseq, note))

    if loc_rows:
        doc.add_heading("Your position", level=2)
        kv_table(loc_rows)
        for wmsg in warnings:
            doc.add_paragraph("Warning: " + wmsg)

    doc.add_heading("Exons and introns", level=2)
    t = doc.add_table(rows=1, cols=6)
    t.style = "Table Grid"
    for i, h in enumerate(["Element", "ID", "Genomic position", "cDNA position", "Length (bp)", "Phase (start/end)"]):
        t.rows[0].cells[i].text = h
        for run in t.rows[0].cells[i].paragraphs[0].runs:
            run.font.bold = True
    for r in elements:
        c = t.add_row().cells
        c[0].text = r["type"]
        c[1].text = str(r["id"])
        c[2].text = f"{r['gstart']}\u2013{r['gend']}" if r["gstart"] is not None else "\u2013"
        c[3].text = f"{r['cstart']}\u2013{r['cend']}" if r["cstart"] is not None else "\u2013"
        c[4].text = f"{r['length']:,}" if isinstance(r["length"], int) else "n/a"
        c[5].text = f"{r['sphase']} / {r['ephase']}" if r["sphase"] != "" else "\u2013"

    def mono_par(text, size=8, color=None, bold=False, highlight=False):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.space_before = Pt(0)
        run = p.add_run(text)
        run.font.name = "Courier New"
        run.font.size = Pt(size)
        run.font.bold = bold
        if color:
            run.font.color.rgb = color
        if highlight:
            run.font.highlight_color = WD_COLOR_INDEX.YELLOW
        return p

    doc.add_heading("Exon and intron sequences", level=2)
    for r in elements:
        label = f"{r['type']}  ({r['length']:,} bp)" if isinstance(r["length"], int) else r["type"]
        doc.add_paragraph().add_run(label).bold = True
        seq = r["sequence"]
        if not seq:
            mono_par("(sequence not available)")
            continue
        if r["type"].startswith("Intron") and len(seq) > 1000:
            seq = f"{seq[:100]}  [...{len(seq) - 200:,} bp not shown...]  {seq[-100:]}"
        color = RGBColor(0x00, 0x33, 0x66) if r["type"].startswith("Exon") else RGBColor(0x80, 0x80, 0x80)
        for i in range(0, len(seq), 60):
            mono_par(seq[i:i + 60], color=color, bold=r["type"].startswith("Exon"))

    doc.add_heading("cDNA sequence with protein", level=2)
    doc.add_paragraph("Exons alternate black and blue; UTR is highlighted yellow; the protein letter sits under the "
                      "middle base of each codon (. = UTR).")
    cdna, n = view["cdna"], len(view["cdna"])
    cds0 = view["cds_start"] - 1 if view["cds_start"] else -1
    cds1 = view["cds_end"] - 1 if view["cds_end"] else -1
    exon_of = [0] * n
    for e in view["exons"]:
        for i in range(e["start"] - 1, min(e["end"], n)):
            exon_of[i] = e["n"]
    for s0 in range(0, n, 60):
        e0 = min(s0 + 60, n)
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.space_before = Pt(6)
        r0 = p.add_run(f"{s0 + 1:>7} ")
        r0.font.name, r0.font.size = "Courier New", Pt(8)
        r0.font.color.rgb = RGBColor(0x7F, 0x8C, 0x8D)
        i = s0
        while i < e0:                      # group runs with the same exon colour / UTR state
            utr = cds0 >= 0 and not (cds0 <= i <= cds1)
            blue = exon_of[i] % 2 == 0
            j = i
            while j < e0 and (cds0 >= 0 and not (cds0 <= j <= cds1)) == utr and (exon_of[j] % 2 == 0) == blue:
                j += 1
            run = p.add_run(cdna[i:j])
            run.font.name, run.font.size = "Courier New", Pt(8)
            run.font.color.rgb = RGBColor(0x1F, 0x3F, 0xFF) if blue else RGBColor(0, 0, 0)
            if utr:
                run.font.highlight_color = WD_COLOR_INDEX.YELLOW
            i = j
        r1 = p.add_run(f" {e0}")
        r1.font.name, r1.font.size = "Courier New", Pt(8)
        r1.font.color.rgb = RGBColor(0x7F, 0x8C, 0x8D)
        mono_par(f"{'':>7} {_protein_line(view, s0, e0)}", color=RGBColor(0xB0, 0x3A, 0x2E), bold=True)
    return doc


@app.route('/api/export-cdna', methods=['GET'])
@login_required
def export_cdna():
    fmt = request.args.get('format', 'excel').lower()
    gene = request.args.get('gene', '').strip()
    tx_in = request.args.get('transcript_id', request.args.get('transcript', '')).strip().split('.')[0].upper()
    coord = request.args.get('coord', 'c').lower()
    pos = request.args.get('pos', '')
    ref = request.args.get('ref', '').strip().upper()
    alt = request.args.get('alt', '').strip().upper()

    if fmt not in ("excel", "word"):
        return jsonify({"error": "format must be 'excel' or 'word'"}), 400
    if fmt == "excel" and not Workbook:
        return jsonify({"error": "openpyxl is not installed (pip install openpyxl)"}), 501
    if fmt == "word" and not Document:
        return jsonify({"error": "python-docx is not installed (pip install python-docx)"}), 501
    if not gene and not tx_in:
        return jsonify({"error": "Provide a gene name or a transcript ID."}), 400
    if coord not in ("c", "n", "g"):
        return jsonify({"error": "Position type must be c., n. or g."}), 400

    try:
        view = get_view_with_fallback(gene, tx_in)
        loc_rows, warnings = [], []
        if pos.strip():
            res = locate_in_transcript(view, coord, pos, ref, alt)
            loc_rows, warnings = res["rows"], res["warnings"]
        elements, note = build_export_data(view)
        refseq = get_refseq_info(view["transcript_id"]) if view.get("source", "").startswith("Ensembl") else None
    except ValueError as e:
        return jsonify({"error": str(e)}), 400
    except EnsemblError as e:
        return jsonify({"error": str(e)}), 502

    stream = io.BytesIO()
    base = f"{view['transcript_id']}_cDNA_exons_introns"
    if fmt == "excel":
        make_workbook(view, elements, loc_rows, warnings, refseq, note).save(stream)
        stream.seek(0)
        return send_file(stream, as_attachment=True, download_name=f"{base}.xlsx",
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    make_word(view, elements, loc_rows, warnings, refseq, note).save(stream)
    stream.seek(0)
    return send_file(stream, as_attachment=True, download_name=f"{base}.docx",
                     mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


if __name__ == '__main__':
    # Debug mode only for local development
    app.run(debug=os.environ.get("FLASK_DEBUG", "1") == "1", port=5000)
