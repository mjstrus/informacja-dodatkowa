"""
Automatyzator Informacji Dodatkowej do sprawozdania finansowego
Streamlit app z Claude 3.5 Sonnet + LlamaParse + python-docx
"""

import streamlit as st
import anthropic
import os
import io
import json
import re
import tempfile
from pathlib import Path
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import base64
import requests
from datetime import date
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker

# ─── PAGE CONFIG ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Generator Informacji Dodatkowej",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ─── STYLES ─────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #1e3a5f 0%, #2d6a9f 100%);
        color: white;
        padding: 2rem;
        border-radius: 12px;
        margin-bottom: 2rem;
        text-align: center;
    }
    .step-card {
        background: #f8f9fa;
        border-left: 4px solid #2d6a9f;
        padding: 1rem 1.5rem;
        border-radius: 0 8px 8px 0;
        margin: 0.5rem 0;
    }
    .validation-ok { color: #28a745; font-weight: bold; }
    .validation-warn { color: #ffc107; font-weight: bold; }
    .validation-err { color: #dc3545; font-weight: bold; }
    .metric-box {
        background: white;
        border: 1px solid #dee2e6;
        border-radius: 8px;
        padding: 1rem;
        text-align: center;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }
    .stProgress > div > div { background-color: #2d6a9f; }
</style>
""", unsafe_allow_html=True)

# ─── HEADER ─────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="main-header">
    <h1>📊 Generator Informacji Dodatkowej</h1>
    <p>Automatyczne tworzenie not do sprawozdania finansowego zgodnie z Ustawą o Rachunkowości</p>
</div>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
# MODUŁ 1: PARSOWANIE PDF
# ═══════════════════════════════════════════════════════════════════════════════

def extract_text_from_pdf_basic(pdf_bytes: bytes, filename: str) -> str:
    """Ekstrakcja tekstu z PDF przy użyciu pypdf (fallback bez LlamaParse)."""
    try:
        import pypdf
        reader = pypdf.PdfReader(io.BytesIO(pdf_bytes))
        text_parts = [f"=== DOKUMENT: {filename} ===\n"]
        for i, page in enumerate(reader.pages):
            page_text = page.extract_text() or ""
            if page_text.strip():
                text_parts.append(f"\n--- Strona {i+1} ---\n{page_text}")
        return "\n".join(text_parts)
    except Exception as e:
        return f"[BŁĄD ekstrakcji {filename}: {e}]"


def parse_documents_llamaparse(pdf_files: list, llama_api_key: str, progress_callback=None) -> dict:
    """
    Krok 1 & 2: Parsowanie PDF przez LlamaParse + identyfikacja dokumentów.
    Zwraca słownik: {nazwa_pliku: tekst_markdown}
    """
    try:
        from llama_parse import LlamaParse
        parser = LlamaParse(
            api_key=llama_api_key,
            result_type="markdown",
            language="pl",
            parsing_instruction=(
                "Dokument to sprawozdanie finansowe polskiej spółki. "
                "Zachowaj strukturę tabel finansowych. "
                "Oznacz wyraźnie: BILANS, RACHUNEK ZYSKÓW I STRAT, NOTY."
            )
        )
        results = {}
        for idx, uploaded_file in enumerate(pdf_files):
            if progress_callback:
                progress_callback(idx / len(pdf_files), f"Parsowanie: {uploaded_file.name}")
            with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
                tmp.write(uploaded_file.getvalue())
                tmp_path = tmp.name
            try:
                docs = parser.load_data(tmp_path)
                results[uploaded_file.name] = "\n\n".join(d.text for d in docs)
            finally:
                os.unlink(tmp_path)
        return results
    except ImportError:
        st.warning("⚠️ LlamaParse niedostępny – używam pypdf jako fallback.")
        return parse_documents_fallback(pdf_files, progress_callback)
    except Exception as e:
        st.warning(f"⚠️ Błąd LlamaParse ({e}) – używam pypdf jako fallback.")
        return parse_documents_fallback(pdf_files, progress_callback)


def parse_documents_fallback(pdf_files: list, progress_callback=None) -> dict:
    """Fallback: ekstrakcja przez pypdf."""
    results = {}
    for idx, uploaded_file in enumerate(pdf_files):
        if progress_callback:
            progress_callback(idx / len(pdf_files), f"Ekstrakcja: {uploaded_file.name}")
        results[uploaded_file.name] = extract_text_from_pdf_basic(
            uploaded_file.getvalue(), uploaded_file.name
        )
    return results


def extract_text_from_docx(docx_bytes: bytes) -> str:
    """Wyciąga tekst z pliku DOCX (paragraphs + tables)."""
    from docx import Document as DocxDocument
    doc = DocxDocument(io.BytesIO(docx_bytes))
    parts = []
    for para in doc.paragraphs:
        if para.text.strip():
            parts.append(para.text)
    for table in doc.tables:
        for row in table.rows:
            row_text = " | ".join(cell.text.strip() for cell in row.cells)
            if row_text.strip():
                parts.append(row_text)
    return "\n".join(parts)


def extract_text_from_xlsx(xlsx_bytes: bytes) -> str:
    """Wyciąga tekst z pliku XLSX (wszystkie arkusze, wiersze jako tekst)."""
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes), data_only=True)
    parts = []
    for ws in wb.worksheets:
        for row in ws.iter_rows(values_only=True):
            vals = [str(v).strip() if v is not None else "" for v in row]
            line = " | ".join(v for v in vals if v)
            if line.strip():
                parts.append(line)
    return "\n".join(parts)


# ═══════════════════════════════════════════════════════════════════════════════
# MODUŁ KRS: POBIERANIE DANYCH Z OFICJALNEGO API MINISTERSTWA SPRAWIEDLIWOŚCI
# ═══════════════════════════════════════════════════════════════════════════════
#
# Oficjalne API KRS (api-krs.ms.gov.pl) działa TYLKO po numerze KRS.
# Endpoint: GET /api/krs/OdpisAktualny/{nrKRS}?rejestr=P&format=json
# Bezpłatne, bez klucza API.

def fetch_krs_by_krs_nr(krs_nr: str) -> dict | None:
    """
    Pobiera dane spółki z oficjalnego API KRS po numerze KRS (10 cyfr).
    """
    krs_clean = re.sub(r"[^0-9]", "", krs_nr).zfill(10)
    if len(krs_clean) != 10:
        return None

    headers = {
        "Accept": "application/json",
        "User-Agent": "Mozilla/5.0 (compatible; InformacjaDodatkowa/1.0)"
    }
    try:
        url = f"https://api-krs.ms.gov.pl/api/krs/OdpisAktualny/{krs_clean}"
        r = requests.get(url, params={"rejestr": "P", "format": "json"},
                         headers=headers, timeout=20)
        if r.status_code == 200:
            return _parse_odpis(r.json(), krs_clean)
        # Spróbuj rejestr S (stowarzyszenia)
        r2 = requests.get(url, params={"rejestr": "S", "format": "json"},
                          headers=headers, timeout=20)
        if r2.status_code == 200:
            return _parse_odpis(r2.json(), krs_clean)
    except requests.exceptions.ConnectionError:
        raise ConnectionError("Brak połączenia z API KRS")
    except requests.exceptions.Timeout:
        raise TimeoutError("API KRS nie odpowiada (timeout)")
    except Exception as e:
        raise RuntimeError(f"Błąd API KRS: {e}")
    return None


def fetch_krs_by_krs_nr_debug(krs_nr: str) -> tuple:
    """Wersja diagnostyczna — zwraca (dane, log)."""
    import json as _json
    krs_clean = re.sub(r"[^0-9]", "", krs_nr).zfill(10)
    log = [f"Numer KRS po oczyszczeniu: {krs_clean}"]
    headers = {
        "Accept": "application/json",
        "User-Agent": "Mozilla/5.0 (compatible; InformacjaDodatkowa/1.0)"
    }
    url = f"https://api-krs.ms.gov.pl/api/krs/OdpisAktualny/{krs_clean}"
    for rejestr in ["P", "S"]:
        try:
            r = requests.get(url, params={"rejestr": rejestr, "format": "json"},
                             headers=headers, timeout=20)
            log.append(f"\n→ {url}?rejestr={rejestr}")
            log.append(f"  Status: {r.status_code}")
            if r.status_code == 200:
                data = r.json()
                preview = _json.dumps(data, ensure_ascii=False, indent=2)[:2000]
                log.append(f"  Odpowiedź:\n{preview}")
                parsed = _parse_odpis(data, krs_clean)
                if parsed:
                    log.append("\n✅ Dane sparsowane pomyślnie")
                    return parsed, "\n".join(log)
            else:
                log.append(f"  Błąd: {r.text[:200]}")
        except Exception as e:
            log.append(f"  Wyjątek: {e}")
    log.append("\n❌ Nie udało się pobrać danych.")
    return None, "\n".join(log)


def _parse_odpis(data: dict, krs_nr: str = "") -> dict | None:
    """Wyciąga potrzebne pola z odpisu JSON zwróconego przez API KRS.
    Struktura rzeczywista: odpis.dane.dzial1.danePodmiotu / siedzibaIAdres / przedmiotDzialalnosci
    """
    try:
        odpis = data.get("odpis", data)
        naglowek = odpis.get("naglowekA", {})
        dane = odpis.get("dane", {})
        dzial1 = dane.get("dzial1", {})
        dane_p = dzial1.get("danePodmiotu", {})

        # ── Nazwa ─────────────────────────────────────────────────────────
        nazwa = dane_p.get("nazwa", "")

        # ── Identyfikatory: NIP i REGON są w osobnym polu ─────────────────
        ident = dane_p.get("identyfikatory", {})
        nip_val = ident.get("nip", "")
        regon_raw = ident.get("regon", "")
        # REGON może być 14-cyfrowy (z zerami) — przytnij do 9
        regon_val = regon_raw[:9] if regon_raw else ""

        # ── Forma prawna ──────────────────────────────────────────────────
        forma = dane_p.get("formaPrawna", "")

        # ── Siedziba i adres są w dzial1.siedzibaIAdres ───────────────────
        siedz_blok = dzial1.get("siedzibaIAdres", {})
        adres = siedz_blok.get("adres", {})
        ulica = adres.get("ulica", "").replace("UL. ", "ul. ").replace("UL.", "ul.")
        nr_domu = adres.get("nrDomu", "")
        nr_lok = adres.get("nrLokalu", "")
        kod = adres.get("kodPocztowy", "")
        miasto = adres.get("miejscowosc", "")
        siedziba = f"{ulica} {nr_domu}".strip()
        if nr_lok:
            siedziba += f"/{nr_lok}"
        if kod and miasto:
            siedziba += f", {kod} {miasto}"

        # ── Numer KRS z nagłówka ──────────────────────────────────────────
        krs_val = naglowek.get("numerKRS", krs_nr)

        # ── Data rejestracji ──────────────────────────────────────────────
        data_rej = naglowek.get("dataRejestracjiWKRS", "")

        # ── PKD — w dzial1.przedmiotDzialalnosci ─────────────────────────
        # PKD może być w dzial1 lub dzial3 — sprawdzamy oba
        pkd = ""
        def _wyciagnij_pkd(sekcja):
            lista = (sekcja.get("przedmiotPrzewazajacejDzialalnosci") or
                     sekcja.get("przedmiotDzialalnosci") or [])
            if isinstance(lista, list) and lista:
                p0 = lista[0]
                return f"{p0.get('kodDzialalnosci','')} {p0.get('opis','')}".strip()
            if isinstance(lista, dict):
                return f"{lista.get('kodDzialalnosci','')} {lista.get('opis','')}".strip()
            return ""

        for dzial_key in ["dzial1", "dzial3", "dzial2"]:
            dzial = dane.get(dzial_key, {})
            pkd_section = dzial.get("przedmiotDzialalnosci", {})
            if pkd_section:
                pkd = _wyciagnij_pkd(pkd_section)
                if pkd:
                    break

        # ── Kapitał podstawowy ─────────────────────────────────────────────
        kapital_blok = dzial1.get("kapital", {})
        kapital_podstawowy = ""
        if kapital_blok:
            kp = kapital_blok.get("wysokoscKapitaluZakladowego", {})
            if isinstance(kp, dict):
                kapital_podstawowy = kp.get("wartosc", "")
            elif isinstance(kp, (str, int, float)):
                kapital_podstawowy = str(kp)

        # ── Wspólnicy — szukaj w wielu lokalizacjach ──────────────────────
        wspolnicy = []
        wspolnicy_raw = None
        for dk in ["dzial1", "dzial2", "dzial3"]:
            dz = dane.get(dk, {})
            for wspol_key in ["wspolnicySpzoo", "wspolnicySpozoo", "wspolnicy",
                               "informacjaOWspolnikach",
                               "komplementariusze", "komandytariusze",
                               "wspolnicySpolkiKomandytowej"]:
                candidate = dz.get(wspol_key)
                if candidate and isinstance(candidate, list) and len(candidate) > 0:
                    wspolnicy_raw = candidate
                    break
                if isinstance(candidate, dict):
                    for sub_key in ["wspolnik", "listaWspolnikow", "dane"]:
                        sub = candidate.get(sub_key)
                        if isinstance(sub, list) and len(sub) > 0:
                            wspolnicy_raw = sub
                            break
            if wspolnicy_raw:
                break

        # Dla spółek komandytowych — szukaj OBIE grupy wspólników
        if not wspolnicy_raw:
            dzial1_keys = dzial1.keys()
            for dk_key in dzial1_keys:
                if "komplement" in dk_key.lower() or "komandyt" in dk_key.lower() or "wspol" in dk_key.lower():
                    candidate = dzial1.get(dk_key)
                    if isinstance(candidate, list) and len(candidate) > 0:
                        if wspolnicy_raw is None:
                            wspolnicy_raw = []
                        wspolnicy_raw.extend(candidate)

        if wspolnicy_raw and isinstance(wspolnicy_raw, list):
            for w in wspolnicy_raw:
                if not isinstance(w, dict):
                    continue
                nazwa_w = w.get("nazwa", "")
                if not nazwa_w:
                    imie = w.get("imiona", "")
                    if isinstance(imie, dict):
                        imie = imie.get("imie", "")
                    elif isinstance(imie, list) and imie:
                        imie = imie[0] if isinstance(imie[0], str) else imie[0].get("imie", "")
                    if not imie:
                        imie = w.get("imie", "")
                    nazwisko = w.get("nazwisko", "")
                    nazwa_w = f"{imie} {nazwisko}".strip()
                if not nazwa_w:
                    continue
                # Udziały (sp. z o.o.) lub wkład (komandytowa)
                udzialy = w.get("posiadaneUdzialy", w.get("udzialy", {}))
                liczba = ""
                wartosc = ""
                if isinstance(udzialy, dict):
                    liczba = udzialy.get("iloscUdzialow", udzialy.get("liczba", ""))
                    wartosc_ud = udzialy.get("wartoscUdzialow", udzialy.get("wartosc", ""))
                    if isinstance(wartosc_ud, dict):
                        wartosc = wartosc_ud.get("wartosc", "")
                    elif isinstance(wartosc_ud, (str, int, float)):
                        wartosc = str(wartosc_ud)
                # Wkład (spółka komandytowa)
                if not wartosc:
                    wklad = w.get("wartoscWkladu", w.get("wklad", {}))
                    if isinstance(wklad, dict):
                        wartosc = wklad.get("wartosc", "")
                    elif isinstance(wklad, (str, int, float)):
                        wartosc = str(wklad)
                    if not liczba:
                        liczba = "wkład"
                wspolnicy.append({
                    "nazwa": nazwa_w,
                    "liczba_udzialow": str(liczba),
                    "wartosc_udzialow": str(wartosc),
                })

        return {
            "nazwa": nazwa,
            "siedziba": siedziba,
            "nip": nip_val,
            "krs": krs_val,
            "regon": regon_val,
            "pkd": pkd,
            "data_rejestracji": data_rej,
            "forma_prawna": forma,
            "kapital_podstawowy": str(kapital_podstawowy),
            "wspolnicy": wspolnicy,
        }
    except Exception as e:
        return None

# ═══════════════════════════════════════════════════════════════════════════════
# MODUŁ 2: IDENTYFIKACJA I MAPOWANIE DOKUMENTÓW
# ═══════════════════════════════════════════════════════════════════════════════

# Definicja wszystkich obsługiwanych typów dokumentów
REQUIRED_DOC_TYPES = {
    "BILANS": {
        "label": "Bilans",
        "icon": "🏦",
        "desc": "Zestawienie aktywów i pasywów na dzień bilansowy",
        "keywords": ["aktywa trwałe", "aktywa obrotowe", "pasywa", "kapitał własny", "zobowiązania"],
    },
    "RZiS": {
        "label": "Rachunek Zysków i Strat",
        "icon": "📈",
        "desc": "Przychody, koszty i wynik finansowy za rok obrotowy",
        "keywords": ["przychody ze sprzedaży", "koszty działalności", "zysk netto", "wynik finansowy", "amortyzacja"],
    },
    "ŚRODKI TRWAŁE": {
        "label": "Tabela środków trwałych",
        "icon": "🏗️",
        "desc": "Wartość brutto, umorzenia i wartość netto środków trwałych",
        "keywords": ["środki trwałe", "wartość brutto", "umorzenie", "odpisy amortyzacyjne"],
    },
    "PRZEPŁYWY PIENIĘŻNE": {
        "label": "Rachunek przepływów pieniężnych",
        "icon": "💸",
        "desc": "Cash flow: operacyjny, inwestycyjny, finansowy",
        "keywords": ["przepływy", "działalność operacyjna", "działalność inwestycyjna"],
    },
    "POLITYKA RACHUNKOWOŚCI": {
        "label": "Polityka rachunkowości",
        "icon": "📜",
        "desc": "Przyjęte zasady rachunkowości, metody wyceny, okresy amortyzacji",
        "keywords": ["polityka rachunkowości", "zasady rachunkowości", "metody wyceny",
                     "okres amortyzacji", "przyjęte zasady", "opis przyjętych"],
    },
    "ZOiS": {
        "label": "Zestawienie Obrotów i Sald",
        "icon": "⚖️",
        "desc": "Obroty i salda kont księgi głównej za rok obrotowy",
        "keywords": ["zestawienie obrotów", "obroty i salda", "salda końcowe",
                     "salda otwarcia", "obroty narastająco", "konta syntetyczne",
                     "księga główna", "salda debetowe", "salda kredytowe",
                     "obroty kont", "w układzie zois", "zois za miesiąc",
                     "obroty kont aktywnych", "obroty kont pasywnych",
                     "saldo wn", "saldo ma", "obroty wn", "obroty ma",
                     "bilans otwarcia", "konto", "ob. wn", "ob. ma"],
    },
    "ANKIETA BILANSOWA": {
        "label": "Ankieta bilansowa",
        "icon": "📋",
        "desc": "Wypełniona ankieta bilansowa od klienta",
        "keywords": ["ankieta bilansowa", "kwestionariusz", "pytania do klienta",
                     "kontynuacja działalności", "zdarzenia po dniu bilansowym",
                     "zobowiązania warunkowe", "podział zysku", "pokrycie straty",
                     "powiązane", "gwarancji i poręczeń", "pożyczek",
                     "nakłady", "transakcje", "postępowani",
                     "sytuacja finansowa jest", "prognoza rozwoju"],
        "required": False,
    },
}


def identify_document_type(text: str) -> str:
    """Heurystyczna identyfikacja typu dokumentu finansowego."""
    text_lower = text.lower()
    scores = {}
    for doc_type, info in REQUIRED_DOC_TYPES.items():
        scores[doc_type] = sum(text_lower.count(kw) for kw in info["keywords"])
    best = max(scores, key=scores.get)
    return best if scores[best] > 0 else "INNY"


def check_missing_documents(doc_mapping: dict) -> list[str]:
    """Zwraca listę typów dokumentów których brakuje wśród wgranych plików."""
    types_found = {d["type"] for d in doc_mapping.values()}
    return [dt for dt in REQUIRED_DOC_TYPES if dt not in types_found]


def map_documents(parsed_docs: dict) -> dict:
    """Mapuje dokumenty do kategorii finansowych."""
    mapping = {}
    for filename, text in parsed_docs.items():
        doc_type = identify_document_type(text)
        mapping[filename] = {
            "type": doc_type,
            "text": text,
            "length": len(text)
        }
    return mapping


# ═══════════════════════════════════════════════════════════════════════════════
# MODUŁ 3: WALIDACJA SPÓJNOŚCI DANYCH
# ═══════════════════════════════════════════════════════════════════════════════

def extract_financial_number(text: str, pattern: str) -> float | None:
    """Wyciąga liczbę z tekstu na podstawie wzorca."""
    try:
        matches = re.findall(
            rf"{pattern}[:\s]+([+-]?\d[\d\s.,]*)",
            text, re.IGNORECASE
        )
        if matches:
            num_str = matches[0].replace(" ", "").replace(",", ".")
            return float(num_str)
    except Exception:
        pass
    return None


def validate_data_consistency(doc_mapping: dict) -> list:
    """
    Krok 3: Sprawdza spójność danych między dokumentami.
    Zwraca listę komunikatów walidacji.
    """
    issues = []
    all_text = "\n".join(d["text"] for d in doc_mapping.values())

    # Sprawdź sumy bilansowe
    aktywne = extract_financial_number(all_text, r"AKTYWA\s+RAZEM|suma\s+aktywów")
    pasywa = extract_financial_number(all_text, r"PASYWA\s+RAZEM|suma\s+pasywów")

    if aktywne and pasywa:
        diff = abs(aktywne - pasywa)
        if diff < 1:
            issues.append({"level": "OK", "msg": f"✅ Bilans zbilansowany: A=P={aktywne:,.0f} PLN"})
        elif diff < aktywne * 0.001:
            issues.append({"level": "WARN", "msg": f"⚠️ Drobna różnica bilansowa: {diff:,.2f} PLN (prawdopodobnie zaokrąglenia)"})
        else:
            issues.append({"level": "ERR", "msg": f"❌ Bilans NIE jest zbilansowany! Różnica: {diff:,.0f} PLN"})
    else:
        issues.append({"level": "WARN", "msg": "⚠️ Nie znaleziono sum bilansowych do porównania"})

    # Sprawdź zysk netto
    zysk_rzis = extract_financial_number(all_text, r"zysk\s+(?:netto|na\s+koniec)")
    zysk_bilans = extract_financial_number(all_text, r"wynik\s+finansowy\s+netto|zysk.*roku\s+obrotowego")
    if zysk_rzis and zysk_bilans:
        diff = abs(zysk_rzis - zysk_bilans)
        if diff < zysk_rzis * 0.001:
            issues.append({"level": "OK", "msg": f"✅ Zysk netto spójny: {zysk_rzis:,.0f} PLN"})
        else:
            issues.append({"level": "WARN", "msg": f"⚠️ Różnica zysku netto między RZiS a Bilansem: {diff:,.0f} PLN"})

    # Typy dokumentów — sprawdź wszystkie zdefiniowane
    types_found = [d["type"] for d in doc_mapping.values()]
    for doc_type, info in REQUIRED_DOC_TYPES.items():
        if doc_type in types_found:
            issues.append({"level": "OK", "msg": f"✅ Znaleziono: {info['icon']} {info['label']}"})
        else:
            issues.append({"level": "WARN", "msg": f"⚠️ Brak dokumentu: {info['icon']} {info['label']}"}) 

    return issues


# ═══════════════════════════════════════════════════════════════════════════════
# ═══════════════════════════════════════════════════════════════════════════════
# MODUŁ 3B: DOBÓR NOT OBJAŚNIAJĄCYCH
# ═══════════════════════════════════════════════════════════════════════════════

NOTA_RULES = {
    1:  {"name": "Zmiana wartości początkowej i umorzenia ŚT", "source": ["ŚRODKI TRWAŁE", "ZOiS"], "category": "auto", "priority": 1},
    2:  {"name": "Zmiana wartości początkowej i umorzenia WNiP", "source": ["ŚRODKI TRWAŁE", "ZOiS"], "category": "auto", "priority": 1},
    3:  {"name": "Zmiana wartości inwestycji długoterminowych", "source": ["ZOiS"], "category": "warunkowe", "priority": 2, "zois_keywords": ["inwestycje długoterminowe", "03"]},
    6:  {"name": "Koszty zakończonych prac rozwojowych oraz wartość firmy", "source": ["ZOiS"], "category": "warunkowe", "priority": 2, "zois_keywords": ["prace rozwojowe", "wartość firmy", "011"]},
    10: {"name": "Odpisy aktualizujące wartość należności", "source": ["ZOiS"], "category": "auto", "priority": 1, "zois_keywords": ["290", "odpis", "należności"]},
    12: {"name": "Struktura własności kapitału podstawowego (sp. z o.o.)", "source": [], "category": "auto", "priority": 1, "forma_prawna": ["SPÓŁKA Z OGRANICZONĄ ODPOWIEDZIALNOŚCIĄ"]},
    13: {"name": "Zmiany stanów kapitałów zapasowego i rezerwowego", "source": ["ZOiS", "BILANS"], "category": "auto", "priority": 1},
    14: {"name": "Zmiany w stanie kapitału z aktualizacji wyceny", "source": ["ZOiS"], "category": "warunkowe", "priority": 2, "zois_keywords": ["803", "aktualizacja wyceny"]},
    15: {"name": "Propozycja podziału zysku za rok obrotowy", "source": ["ANKIETA BILANSOWA"], "category": "ankieta", "priority": 1, "ankieta_trigger": "q6_zysk"},
    16: {"name": "Propozycja pokrycia straty za rok obrotowy", "source": ["ANKIETA BILANSOWA"], "category": "ankieta", "priority": 1, "ankieta_trigger": "q7_strata"},
    17: {"name": "Rezerwy na koszty i zobowiązania", "source": ["ZOiS", "BILANS"], "category": "auto", "priority": 1},
    18: {"name": "Odroczony podatek dochodowy", "source": ["ZOiS"], "category": "auto", "priority": 1, "zois_keywords": ["650", "841", "odroczony"]},
    19: {"name": "Zobowiązania według okresów wymagalności", "source": ["ZOiS", "BILANS"], "category": "auto", "priority": 1},
    21: {"name": "Czynne rozliczenia międzyokresowe", "source": ["ZOiS"], "category": "auto", "priority": 1, "zois_keywords": ["640", "rozliczenia międzyokresowe"]},
    22: {"name": "Rozliczenia międzyokresowe przychodów", "source": ["ZOiS"], "category": "warunkowe", "priority": 2, "zois_keywords": ["840", "845", "rozliczenia międzyokresowe przychod"]},
    29: {"name": "Struktura rzeczowa i terytorialna przychodów", "source": ["RZiS", "ZOiS"], "category": "auto", "priority": 1},
    31: {"name": "Koszty rodzajowe (porównanie rok bieżący vs poprzedni)", "source": ["RZiS"], "category": "auto", "priority": 1},
    35: {"name": "Rozliczenie różnicy CIT vs wynik finansowy", "source": ["RZiS", "ZOiS"], "category": "auto", "priority": 1},
    40: {"name": "Kursy walut przyjęte do wyceny", "source": ["ZOiS"], "category": "warunkowe", "priority": 2, "zois_keywords": ["walut", "kursow", "EUR", "USD", "GBP"]},
    41: {"name": "Struktura środków pieniężnych", "source": ["ZOiS"], "category": "auto", "priority": 1},
    57: {"name": "Różnica zobowiązań krótkoterminowych (bilans vs przepływy)", "source": ["BILANS", "PRZEPŁYWY PIENIĘŻNE"], "category": "auto", "priority": 2, "require_all_sources": True},
    58: {"name": "Różnica zapasów (bilans vs przepływy)", "source": ["BILANS", "PRZEPŁYWY PIENIĘŻNE"], "category": "auto", "priority": 2, "require_all_sources": True},
    60: {"name": "Struktura należności", "source": ["ZOiS"], "category": "auto", "priority": 1},
    61: {"name": "Należności według okresów wymagalności", "source": ["ZOiS"], "category": "auto", "priority": 1},
    63: {"name": "Środki pieniężne na rachunku VAT", "source": ["ZOiS"], "category": "warunkowe", "priority": 2, "zois_keywords": ["VAT", "rachunek VAT"]},
    73: {"name": "Zobowiązania długoterminowe > 5 lat", "source": ["ZOiS", "BILANS"], "category": "warunkowe", "priority": 2},
    76: {"name": "Informacje o transakcjach z jednostkami powiązanymi", "source": ["ANKIETA BILANSOWA"], "category": "ankieta", "priority": 2, "ankieta_trigger": "q14_powiazane"},
}

ANKIETA_TRIGGERS = {
    "q6_zysk": {"positive": ["przeznaczenie zysku", "wypłata dywidendy", "kapitał zapasowy", "podwyższenie kapitału"]},
    "q7_strata": {"positive": ["pokrycie straty", "zyskiem z lat", "kapitale zapasowym"]},
    "q14_powiazane": {"question": "transakcje ze stronami powiązanymi", "positive_answer": "tak"},
}


def _check_ankieta_trigger(trigger_key: str, ankieta_text: str) -> bool:
    if not ankieta_text:
        return False
    trigger = ANKIETA_TRIGGERS.get(trigger_key, {})
    text_lower = ankieta_text.lower()
    if "question" in trigger:
        q_pos = text_lower.find(trigger["question"])
        if q_pos == -1:
            return False
        answer_region = text_lower[q_pos:q_pos + 100]
        pos_pos = answer_region.find(trigger["positive_answer"])
        if pos_pos == -1:
            return False
        import re as _re
        nie_matches = list(_re.finditer(r'\bnie\b', answer_region))
        if nie_matches and nie_matches[0].start() < pos_pos:
            return False
        return True
    if "positive" in trigger:
        return any(kw in text_lower for kw in trigger["positive"])
    return False


def select_applicable_notes(doc_mapping: dict, company_info: dict = None) -> list:
    info = company_info or {}
    types_found = {d["type"] for d in doc_mapping.values()}
    ankieta_text = ""
    zois_text = ""
    for doc_data in doc_mapping.values():
        if doc_data["type"] == "ANKIETA BILANSOWA":
            ankieta_text = doc_data["text"]
        if doc_data["type"] == "ZOiS":
            zois_text = doc_data["text"].lower()

    selected = []
    for nota_nr, rule in sorted(NOTA_RULES.items()):
        reason = ""
        include = False
        # Forma prawna check
        if "forma_prawna" in rule:
            forma = (info.get("forma_prawna", "") or "").upper()
            if not any(fp.upper() in forma for fp in rule["forma_prawna"]):
                continue

        if rule["category"] == "auto":
            sources = rule.get("source", [])
            matched = [s for s in sources if s in types_found]
            if rule.get("require_all_sources"):
                if len(matched) == len(sources) and sources:
                    include = True
                    reason = f"Źródło: {', '.join(matched)}"
            elif matched:
                include = True
                reason = f"Źródło: {', '.join(matched)}"
            elif not sources:
                include = True
                reason = "Nota standardowa"
        elif rule["category"] == "ankieta":
            trigger_key = rule.get("ankieta_trigger", "")
            if ankieta_text and _check_ankieta_trigger(trigger_key, ankieta_text):
                include = True
                reason = "Trigger z ankiety bilansowej"
            elif not ankieta_text and rule["priority"] <= 1:
                include = True
                reason = "Brak ankiety — wymagane dane od klienta"
        elif rule["category"] == "warunkowe":
            sources = rule.get("source", [])
            matched = [s for s in sources if s in types_found]
            if matched:
                zois_kw = rule.get("zois_keywords", [])
                if zois_kw and zois_text:
                    if any(kw.lower() in zois_text for kw in zois_kw):
                        include = True
                        reason = f"Wykryto dane w ZOiS"
                elif not zois_kw:
                    non_zois = [s for s in matched if s != "ZOiS"]
                    if non_zois:
                        include = True
                        reason = f"Źródło: {', '.join(non_zois)}"

        if include:
            selected.append({"nr": nota_nr, "name": rule["name"], "category": rule["category"], "priority": rule["priority"], "reason": reason})

    return selected


def format_notes_for_prompt(selected_notes: list) -> str:
    if not selected_notes:
        return ""
    prio1 = [n for n in selected_notes if n["priority"] == 1]
    prio2 = [n for n in selected_notes if n["priority"] == 2]
    lines = [
        f"\n📋 NOTY DO WYGENEROWANIA ({len(selected_notes)} not):",
        "Numeruj noty SEKWENCYJNIE (Nota 1, Nota 2...). Numery GOFIN w nawiasach to referencja — NIE umieszczaj ich w dokumencie.\n",
    ]
    for n in prio1:
        lines.append(f"  ✅ {n['name']} [GOFIN {n['nr']}]")
    if prio2:
        lines.append("\nWAŻNE (jeśli dane wystarczające):")
        for n in prio2:
            lines.append(f"  📌 {n['name']} [GOFIN {n['nr']}]")
    lines.append(
        "\nJeśli nota ma SAME ZERA lub zjawisko nie wystąpiło — napisz 'Nie dotyczy.'"
        "\nNIGDY nie pisz '[DANE DO UZUPEŁNIENIA]'. Użyj myślnika '—' w pustych komórkach.\n"
    )
    return "\n".join(lines)


# ═══════════════════════════════════════════════════════════════════════════════
# MODUŁ 4: GENEROWANIE INFORMACJI DODATKOWEJ PRZEZ CLAUDE
# ═══════════════════════════════════════════════════════════════════════════════

SYSTEM_PROMPT = """Jesteś biegłym rewidentem i ekspertem ds. rachunkowości, specjalizującym się w polskim prawie bilansowym.
Twoim zadaniem jest sporządzenie profesjonalnej "Informacji Dodatkowej" do sprawozdania finansowego.

WYMAGANIA PRAWNE (Ustawa o Rachunkowości):
- Art. 48 UoR: Informacja dodatkowa obejmuje wprowadzenie i dodatkowe informacje i objaśnienia
- Stosuj Krajowe Standardy Rachunkowości (KSR)

STRUKTURA DOKUMENTU (obowiązkowa):
1. WPROWADZENIE DO SPRAWOZDANIA FINANSOWEGO
   1.1 Dane identyfikacyjne jednostki
   1.2 Zasady (polityka) rachunkowości
   1.3 Metody wyceny aktywów i pasywów
   1.4 Metody amortyzacji i stosowane stawki
   1.5 Zasady rozliczania przychodów i kosztów
   1.6 Korekty błędów i zmiany polityki rachunkowości

2. DODATKOWE INFORMACJE I OBJAŚNIENIA
   Generuj TYLKO noty wymienione w sekcji "NOTY DO WYGENEROWANIA".

STYL: Profesjonalne słownictwo, PLN z dokładnością do groszy, tryb oznajmujący.
- Tabele generuj w formacie MARKDOWN (| kolumna1 | kolumna2 |)
- ZASADA ZEROWYCH WARTOŚCI: Jeśli dla noty WSZYSTKIE wartości = 0, napisz "Nie dotyczy." NIE generuj tabeli.
- NUMERACJA NOT: Numeruj SEKWENCYJNIE (Nota 1, Nota 2, Nota 3...) NIE używaj numerów GOFIN.

╔══════════════════════════════════════════════════════════════════════════════╗
║  BEZWZGLĘDNY ZAKAZ UŻYWANIA "[DANE DO UZUPEŁNIENIA]"                        ║
╚══════════════════════════════════════════════════════════════════════════════╝
NIGDY nie wstawiaj "[DANE DO UZUPEŁNIENIA]". Zamiast tego:
A) Zjawisko nie wystąpiło → "Nie dotyczy."
B) Masz kwotę łączną bez rozbicia → podaj łączną, puste komórki wypełnij "—"
C) Brak danych w dokumentach → "Na dzień sporządzenia sprawozdania Spółka nie przedłożyła danych w tym zakresie."
D) Nota z samymi zerami → "Nie dotyczy."

Jeśli dostarczono Politykę Rachunkowości – sekcja 1.2–1.5 oparta WYŁĄCZNIE na jej treści.

DANE DO WYKRESÓW: Na końcu dokumentu dodaj blok JSON w znacznikach <!--CHART_DATA_START--> i <!--CHART_DATA_END-->:
<!--CHART_DATA_START-->
{"aktywa_trwale":0,"aktywa_obrotowe":0,"kapital_wlasny":0,"zobowiazania_dlugoterminowe":0,"zobowiazania_krotkoterminowe":0,"przychody_ze_sprzedazy":0,"koszty_dzialalnosci":0,"wynik_finansowy_netto":0,"srodki_trwale_brutto":0,"srodki_trwale_umorzenie":0,"srodki_trwale_netto":0,"naleznosci_krotkoterminowe":0,"srodki_pieniezne":0,"zapasy":0,"przychody_rok_poprzedni":0,"koszty_rok_poprzedni":0,"wynik_rok_poprzedni":0}
<!--CHART_DATA_END-->
Wypełnij TYLKO pola z konkretnymi danymi."""


def generate_accounting_notes(doc_mapping: dict, anthropic_api_key: str,
                               company_name: str, year: int,
                               company_info: dict = None,
                               progress_callback=None) -> str:
    client = anthropic.Anthropic(api_key=anthropic_api_key)
    info = company_info or {}

    # Polityka rachunkowości
    polityka_blok = ""
    pa = info.get("polityka_answers", {})
    if pa:
        polityka_blok = """
📋 ZASADY RACHUNKOWOŚCI (odpowiedzi użytkownika — brak Polityki Rachunkowości):
- Zasady ustalania wyniku finansowego: {wynik}
- Wycena zapasów: {zapasy}
- Amortyzacja środków trwałych: {amort}
- Wycena należności: {nal}
- Sposób sporządzania sprawozdania: {spr}
- Podatek odroczony: {pod}
- Ujęcie leasingu: {leas}
{uwagi_blok}
Na podstawie powyższych wypełnij sekcje 1.2–1.5.""".format(
            wynik=pa.get("wynik_finansowy", ""), zapasy=pa.get("wycena_zapasow", ""),
            amort=pa.get("amortyzacja", ""), nal=pa.get("wycena_naleznosci", ""),
            spr=pa.get("sposob_sprawozdania", ""),
            pod="TAK" if pa.get("podatek_odroczony") else "NIE",
            leas=pa.get("leasing", ""),
            uwagi_blok=f"- Uwagi: {pa['uwagi']}" if pa.get("uwagi") else ""
        )

    zagrozenie_blok = ""
    if info.get("zagrozenie_kontynuacji"):
        zagrozenie_blok = (
            "\n⚠️ OKOLICZNOŚCI ZAGROŻENIA KONTYNUOWANIA DZIAŁALNOŚCI (art. 5 ust. 2 UoR). "
            f"Opis: {info.get('zagrozenie_opis', '')}\n"
        )

    context_parts = [
        f"NAZWA JEDNOSTKI: {info.get('nazwa') or company_name}",
        f"FORMA PRAWNA: {info.get('forma_prawna', '')}",
        f"SIEDZIBA: {info.get('siedziba', '')}",
        f"NIP: {info.get('nip', '')}",
        f"NR KRS: {info.get('krs', '')}",
        f"REGON: {info.get('regon', '')}",
        f"GŁÓWNY PKD: {info.get('pkd', '')}",
        f"DATA REJESTRACJI W KRS: {info.get('data_rejestracji', '')}",
        f"OKRES SPRAWOZDAWCZY: od {info.get('okres_od', '')} do {info.get('okres_do', '')}",
        f"ZATRUDNIENIE ŚREDNIE ROK BIEŻĄCY: {info.get('zatrudnienie_biezacy', 0)} etatów",
        f"ZATRUDNIENIE ŚREDNIE ROK POPRZEDNI: {info.get('zatrudnienie_poprzedni', 0)} etatów",
        f"UWAGI DO ZATRUDNIENIA: {info.get('zatrudnienie_uwagi', '')}",
        f"ROK OBROTOWY: {year}",
    ]

    # Wspólnicy z KRS
    wspolnicy = info.get("wspolnicy", [])
    kapital_podst = info.get("kapital_podstawowy", "")
    if wspolnicy or kapital_podst:
        context_parts.append("\n👥 STRUKTURA WŁASNOŚCI KAPITAŁU (z KRS):")
        if kapital_podst:
            context_parts.append(f"Kapitał podstawowy: {kapital_podst} PLN")
        for w in wspolnicy:
            context_parts.append(f"  - {w.get('nazwa','?')}: {w.get('liczba_udzialow','?')} udziałów, wartość {w.get('wartosc_udzialow','?')} PLN")

    # Wynagrodzenie audytora
    wyn_aud = info.get("wynagrodzenie_audytora", "")
    if wyn_aud:
        context_parts.append(f"💰 WYNAGRODZENIE FIRMY AUDYTORSKIEJ: {wyn_aud} PLN")

    context_parts.append(polityka_blok)
    context_parts.append(zagrozenie_blok)
    context_parts.append("=" * 60)

    # Lista wybranych not
    selected_notes = info.get("selected_notes", [])
    if selected_notes:
        context_parts.append(format_notes_for_prompt(selected_notes))

    context_parts.append("=" * 60)

    # Ankieta bilansowa
    ankieta_found = False
    for filename, doc_data in doc_mapping.items():
        if doc_data["type"] == "ANKIETA BILANSOWA":
            ankieta_found = True
            context_parts.append("\n📋 ANKIETA BILANSOWA:")
            context_parts.append(doc_data["text"][:12000])
            break
    if not ankieta_found:
        context_parts.append(
            "\n⚠️ BRAK ANKIETY BILANSOWEJ. W sekcjach podziału wyniku, zdarzeń po dniu bilansowym "
            "— napisz: 'Na dzień sporządzenia sprawozdania Spółka nie przedłożyła danych w tym zakresie.'"
        )

    # Dokumenty finansowe
    context_parts.append("\nWYCIĄGI Z DOKUMENTÓW FINANSOWYCH:")
    for filename, doc_data in doc_mapping.items():
        if doc_data["type"] == "ANKIETA BILANSOWA":
            continue
        context_parts.append(f"\n[{doc_data['type']}] {filename}:")
        context_parts.append(doc_data["text"][:8000])
        if len(doc_data["text"]) > 8000:
            context_parts.append("...[tekst skrócony]")

    full_context = "\n".join(context_parts)

    user_prompt = f"""Na podstawie poniższych dokumentów sporządź kompletną "Informację Dodatkową" za rok {year}.

{full_context}

Wygeneruj pełną Informację Dodatkową zgodnie z UoR.
Gdzie masz dane — użyj konkretnych liczb. Gdzie brakuje — napisz "Nie dotyczy." lub użyj myślnika "—".
NIGDY nie pisz "[DANE DO UZUPEŁNIENIA]".
Formatuj nagłówkami i akapitami."""

    if progress_callback:
        progress_callback(0.7, "Generowanie przez Claude...")

    response = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=8000,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": user_prompt}]
    )
    return response.content[0].text


# ═══════════════════════════════════════════════════════════════════════════════
# MODUŁ 5: EKSPORT DO WORD (.docx) — PROFESJONALNA SZATA + WYKRESY
# ═══════════════════════════════════════════════════════════════════════════════

_CLR_NAVY = RGBColor(0x1B, 0x2A, 0x4A)
_CLR_BLUE = RGBColor(0x2D, 0x6A, 0x9F)
_CLR_GRAY = RGBColor(0x66, 0x66, 0x66)
_CLR_LIGHT = RGBColor(0x99, 0x99, 0x99)
_CLR_BLACK = RGBColor(0x33, 0x33, 0x33)
_TBL_HDR = "1B2A4A"
_TBL_ALT = "F2F6FA"


def _extract_chart_data(text):
    m = re.search(r"<!--CHART_DATA_START-->\s*(\{.*?\})\s*<!--CHART_DATA_END-->", text, re.DOTALL)
    if not m: return {}
    try: return json.loads(m.group(1))
    except: return {}


def _strip_chart_data(text):
    return re.sub(r"\s*<!--CHART_DATA_START-->.*?<!--CHART_DATA_END-->\s*", "", text, flags=re.DOTALL).strip()


def _fmt_pln(v):
    if abs(v) >= 1_000_000: return f"{v/1_000_000:,.1f} mln"
    if abs(v) >= 1_000: return f"{v/1_000:,.0f} tys."
    return f"{v:,.0f}"

def _setup_chart():
    plt.rcParams.update({"font.family":"sans-serif","font.size":9,"axes.titlesize":11,
        "axes.titleweight":"bold","axes.spines.top":False,"axes.spines.right":False,
        "figure.facecolor":"white","figure.dpi":150})

_COLORS = ["#1B2A4A","#2D6A9F","#3A86C8","#6BAED6","#9ECAE1","#C6DBEF"]


def generate_charts(data, year):
    if not data: return []
    charts = []
    # 1. Aktywa pie
    vals = [data.get("aktywa_trwale",0), data.get("aktywa_obrotowe",0)]
    if all(v>0 for v in vals):
        _setup_chart(); fig,ax = plt.subplots(figsize=(5,3.5))
        ax.pie(vals,autopct="%1.1f%%",colors=_COLORS[:2],startangle=90,pctdistance=0.75,wedgeprops={"linewidth":1,"edgecolor":"white"})
        ax.legend([f"Aktywa trwałe ({_fmt_pln(vals[0])} PLN)",f"Aktywa obrotowe ({_fmt_pln(vals[1])} PLN)"],loc="center left",bbox_to_anchor=(1,0.5),fontsize=8,frameon=False)
        ax.set_title(f"Struktura aktywów — {year}",color="#1B2A4A")
        buf=io.BytesIO(); fig.savefig(buf,format="png",bbox_inches="tight"); plt.close(fig); buf.seek(0)
        charts.append(("Struktura aktywów",buf.getvalue()))
    # 2. Pasywa pie
    pv = [data.get("kapital_wlasny",0),data.get("zobowiazania_dlugoterminowe",0),data.get("zobowiazania_krotkoterminowe",0)]
    pl = ["Kapitał własny","Zob. długoterm.","Zob. krótkoterm."]
    filt = [(l,v) for l,v in zip(pl,pv) if v>0]
    if len(filt)>=2:
        _setup_chart(); fig,ax = plt.subplots(figsize=(5,3.5)); ls,vs=zip(*filt)
        ax.pie(vs,autopct="%1.1f%%",colors=_COLORS[:len(vs)],startangle=90,pctdistance=0.75,wedgeprops={"linewidth":1,"edgecolor":"white"})
        ax.legend([f"{l} ({_fmt_pln(v)} PLN)" for l,v in zip(ls,vs)],loc="center left",bbox_to_anchor=(1,0.5),fontsize=8,frameon=False)
        ax.set_title(f"Struktura pasywów — {year}",color="#1B2A4A")
        buf=io.BytesIO(); fig.savefig(buf,format="png",bbox_inches="tight"); plt.close(fig); buf.seek(0)
        charts.append(("Struktura pasywów",buf.getvalue()))
    # 3. Przychody vs koszty bar
    cur=[data.get("przychody_ze_sprzedazy",0),data.get("koszty_dzialalnosci",0),data.get("wynik_finansowy_netto",0)]
    prev=[data.get("przychody_rok_poprzedni",0),data.get("koszty_rok_poprzedni",0),data.get("wynik_rok_poprzedni",0)]
    if cur[0]>0:
        _setup_chart(); fig,ax = plt.subplots(figsize=(6,3.5)); x=range(3); w=0.35
        if sum(prev)>0:
            ax.bar([i-w/2 for i in x],prev,w,label="Rok poprzedni",color="#9ECAE1")
            ax.bar([i+w/2 for i in x],cur,w,label="Rok bieżący",color="#1B2A4A")
            ax.legend(fontsize=8,frameon=False)
        else: ax.bar(x,cur,w*1.5,color=["#2D6A9F","#6BAED6","#27AE60" if cur[2]>=0 else "#E74C3C"])
        ax.set_xticks(x); ax.set_xticklabels(["Przychody","Koszty","Wynik netto"])
        ax.set_title(f"Przychody, koszty i wynik — {year}",color="#1B2A4A")
        ax.yaxis.set_major_formatter(ticker.FuncFormatter(lambda v,p:_fmt_pln(v)))
        buf=io.BytesIO(); fig.savefig(buf,format="png",bbox_inches="tight"); plt.close(fig); buf.seek(0)
        charts.append(("Analiza wyniku finansowego",buf.getvalue()))
    # 4. Środki trwałe bar
    sb,su,sn = data.get("srodki_trwale_brutto",0),data.get("srodki_trwale_umorzenie",0),data.get("srodki_trwale_netto",0)
    if sb>0:
        _setup_chart(); fig,ax = plt.subplots(figsize=(5,3))
        bars=ax.bar(["Wart. brutto","Umorzenie","Wart. netto"],[sb,su,sn],color=["#1B2A4A","#E74C3C","#27AE60"],width=0.5)
        for bar,v in zip(bars,[sb,su,sn]): ax.text(bar.get_x()+bar.get_width()/2,bar.get_height(),_fmt_pln(v),ha="center",va="bottom",fontsize=8)
        ax.set_title(f"Środki trwałe — {year}",color="#1B2A4A")
        ax.yaxis.set_major_formatter(ticker.FuncFormatter(lambda v,p:_fmt_pln(v)))
        buf=io.BytesIO(); fig.savefig(buf,format="png",bbox_inches="tight"); plt.close(fig); buf.seek(0)
        charts.append(("Środki trwałe",buf.getvalue()))
    return charts


def _add_separator(doc,color="2D6A9F",sz="6"):
    p=doc.add_paragraph(); pPr=p._p.get_or_add_pPr(); pBdr=OxmlElement("w:pBdr")
    b=OxmlElement("w:bottom"); b.set(qn("w:val"),"single"); b.set(qn("w:sz"),sz)
    b.set(qn("w:space"),"1"); b.set(qn("w:color"),color); pBdr.append(b); pPr.append(pBdr)

def _set_cell_shading(cell,color_hex):
    shd=OxmlElement("w:shd"); shd.set(qn("w:fill"),color_hex); shd.set(qn("w:val"),"clear")
    cell._tc.get_or_add_tcPr().append(shd)

def _add_page_number(doc):
    for section in doc.sections:
        footer=section.footer; footer.is_linked_to_previous=False
        p=footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        p.alignment=WD_ALIGN_PARAGRAPH.CENTER
        run=p.add_run("Strona "); run.font.size=Pt(8); run.font.color.rgb=_CLR_LIGHT
        f1=OxmlElement("w:fldChar"); f1.set(qn("w:fldCharType"),"begin"); p.add_run()._r.append(f1)
        ins=OxmlElement("w:instrText"); ins.set(qn("xml:space"),"preserve"); ins.text=" PAGE "
        r3=p.add_run(); r3.font.size=Pt(8); r3._r.append(ins)
        f2=OxmlElement("w:fldChar"); f2.set(qn("w:fldCharType"),"end"); p.add_run()._r.append(f2)

def add_markdown_table_to_doc(doc,table_lines):
    data_lines=[l for l in table_lines if not re.match(r"^\|[\s\-:|]+$",l)]
    if not data_lines: return
    rows=[[c.strip() for c in l.strip("|").split("|")] for l in data_lines]
    if not rows: return
    nc=max(len(r) for r in rows)
    for r in rows:
        while len(r)<nc: r.append("")
    table=doc.add_table(rows=len(rows),cols=nc)
    try: table.style="Table Grid"
    except: pass
    table.autofit=True
    tblPr=table._tbl.tblPr if table._tbl.tblPr else OxmlElement("w:tblPr")
    borders=OxmlElement("w:tblBorders")
    for edge in ["top","left","bottom","right","insideH","insideV"]:
        el=OxmlElement(f"w:{edge}"); el.set(qn("w:val"),"single"); el.set(qn("w:sz"),"4")
        el.set(qn("w:space"),"0"); el.set(qn("w:color"),"B0B0B0"); borders.append(el)
    tblPr.append(borders)
    for i,rd in enumerate(rows):
        for j,ct in enumerate(rd):
            cell=table.rows[i].cells[j]; cell.text=""; p=cell.paragraphs[0]
            p.paragraph_format.space_before=Pt(2); p.paragraph_format.space_after=Pt(2)
            if i==0:
                _set_cell_shading(cell,_TBL_HDR)
                run=p.add_run(ct.replace("**","")); run.bold=True; run.font.size=Pt(8)
                run.font.color.rgb=RGBColor(0xFF,0xFF,0xFF)
            else:
                if i%2==0: _set_cell_shading(cell,_TBL_ALT)
                parts=re.split(r"(\*\*[^*]+\*\*)",ct)
                for part in parts:
                    if part.startswith("**") and part.endswith("**"): run=p.add_run(part[2:-2]); run.bold=True
                    else: run=p.add_run(part)
                    run.font.size=Pt(8); run.font.color.rgb=_CLR_BLACK
    doc.add_paragraph()


def save_to_word(generated_text,company_name,year,company_info=None):
    doc=Document()
    chart_data=_extract_chart_data(generated_text)
    clean_text=_strip_chart_data(generated_text)
    charts=generate_charts(chart_data,year)
    info=company_info or {}
    style=doc.styles["Normal"]; style.font.name="Calibri"; style.font.size=Pt(10)
    style.font.color.rgb=_CLR_BLACK; style.paragraph_format.line_spacing=1.15
    for lev,sz,clr in [(1,14,_CLR_NAVY),(2,12,_CLR_BLUE),(3,11,_CLR_NAVY),(4,10,_CLR_BLUE)]:
        h=doc.styles[f"Heading {lev}"]; h.font.name="Calibri"; h.font.size=Pt(sz)
        h.font.bold=True; h.font.color.rgb=clr
    for section in doc.sections:
        section.left_margin=Inches(1.0); section.right_margin=Inches(1.0)
        section.top_margin=Inches(0.8); section.bottom_margin=Inches(0.7)
    _add_page_number(doc)
    # Strona tytułowa
    for _ in range(4): doc.add_paragraph()
    _add_separator(doc,"2D6A9F","12")
    p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    run=p.add_run("INFORMACJA DODATKOWA"); run.font.size=Pt(22); run.font.bold=True; run.font.color.rgb=_CLR_NAVY
    p2=doc.add_paragraph(); p2.alignment=WD_ALIGN_PARAGRAPH.CENTER
    run2=p2.add_run(f"do sprawozdania finansowego za rok obrotowy {year}")
    run2.font.size=Pt(14); run2.font.color.rgb=_CLR_BLUE
    _add_separator(doc,"2D6A9F","12"); doc.add_paragraph()
    fields=[("Jednostka",info.get("nazwa") or company_name),("Forma prawna",info.get("forma_prawna","")),
            ("Siedziba",info.get("siedziba","")),("NIP",info.get("nip","")),
            ("KRS",info.get("krs","")),("REGON",info.get("regon","")),
            ("Okres sprawozdawczy",f"{info.get('okres_od','')} — {info.get('okres_do','')}"  )]
    for label,value in fields:
        if not value or value==" — ": continue
        p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
        rl=p.add_run(f"{label}: "); rl.font.size=Pt(10); rl.font.color.rgb=_CLR_GRAY
        rv=p.add_run(value); rv.font.size=Pt(10); rv.font.bold=True; rv.font.color.rgb=_CLR_NAVY
    doc.add_paragraph()
    pd=doc.add_paragraph(); pd.alignment=WD_ALIGN_PARAGRAPH.CENTER
    rd=pd.add_run(f"Wygenerowano: {date.today().strftime('%d.%m.%Y')}")
    rd.font.size=Pt(9); rd.font.color.rgb=_CLR_LIGHT; rd.font.italic=True
    doc.add_page_break()
    # Parsowanie
    lines=clean_text.split("\n"); i=0
    while i<len(lines):
        line=lines[i].strip()
        if line.startswith("|") and "|" in line[1:]:
            tl=[]
            while i<len(lines) and lines[i].strip().startswith("|"):
                tl.append(lines[i].strip()); i+=1
            add_markdown_table_to_doc(doc,tl); continue
        i+=1
        if not line: continue
        if line.startswith("#### "): doc.add_heading(line[5:],level=4)
        elif line.startswith("### "): doc.add_heading(line[4:],level=3)
        elif line.startswith("## "): doc.add_heading(line[3:],level=2)
        elif line.startswith("# "): doc.add_heading(line[2:],level=1)
        elif line.startswith("---"): _add_separator(doc,"CCCCCC","2")
        elif line.startswith("**") and line.endswith("**"):
            p=doc.add_paragraph(); run=p.add_run(line.strip("*")); run.bold=True; run.font.color.rgb=_CLR_NAVY
        elif line.startswith("- ") or line.startswith("* "): doc.add_paragraph(line[2:],style="List Bullet")
        elif re.match(r"^\d+\.\s",line): doc.add_paragraph(line,style="List Number")
        else:
            p=doc.add_paragraph()
            for part in re.split(r"(\*\*[^*]+\*\*)",line):
                if part.startswith("**") and part.endswith("**"): run=p.add_run(part[2:-2]); run.bold=True
                else: p.add_run(part)
    # Wykresy
    if charts:
        doc.add_page_break(); doc.add_heading("Analiza graficzna",level=1)
        for idx,(title,png) in enumerate(charts):
            pt=doc.add_paragraph(); run=pt.add_run(f"Wykres {idx+1}. {title}")
            run.bold=True; run.font.color.rgb=_CLR_BLUE
            pi=doc.add_paragraph(); pi.alignment=WD_ALIGN_PARAGRAPH.CENTER
            pi.add_run().add_picture(io.BytesIO(png),width=Inches(5.0))
    # Stopka
    _add_separator(doc,"2D6A9F","4")
    fp=doc.add_paragraph(); fp.alignment=WD_ALIGN_PARAGRAPH.CENTER
    fr=fp.add_run(f"Informacja Dodatkowa | {company_name} | {year} | {date.today().strftime('%d.%m.%Y')}")
    fr.font.size=Pt(8); fr.font.color.rgb=_CLR_LIGHT; fr.font.italic=True
    buffer=io.BytesIO(); doc.save(buffer); return buffer.getvalue()


# ═══════════════════════════════════════════════════════════════════════════════
# GŁÓWNY INTERFEJS STREAMLIT
# ═══════════════════════════════════════════════════════════════════════════════

# ── Odczyt kluczy z Streamlit Secrets (jeśli ustawione) ─────────────────────
_anthropic_from_secrets = st.secrets.get("ANTHROPIC_API_KEY", "")
_llama_from_secrets = st.secrets.get("LLAMA_API_KEY", "")
_app_password = st.secrets.get("APP_PASSWORD", "")

# ── Ochrona hasłem ───────────────────────────────────────────────────────────────────────────
if _app_password:
    if not st.session_state.get("authenticated"):
        st.markdown("""
        <div style="max-width:400px; margin: 4rem auto; padding: 2rem;
                    border:1px solid #dee2e6; border-radius:12px;
                    box-shadow: 0 4px 12px rgba(0,0,0,0.1); text-align:center;">
            <h2>🔐 Dostęp chroniony</h2>
            <p style="color:#666;">Wprowadź hasło aby kontynuować</p>
        </div>
        """, unsafe_allow_html=True)
        col_c, col_mid, col_d = st.columns([1, 2, 1])
        with col_mid:
            entered = st.text_input("Hasło", type="password",
                                    label_visibility="collapsed",
                                    placeholder="Wpisz hasło dostępu...")
            if st.button("Zaloguj →", use_container_width=True, type="primary"):
                if entered == _app_password:
                    st.session_state["authenticated"] = True
                    st.rerun()
                else:
                    st.error("❌ Nieprawiłowe hasło")
        st.stop()

# ── Sidebar: Konfiguracja ────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Konfiguracja")

    if _anthropic_from_secrets:
        st.success("🔑 Klucz API Anthropic: wczytany automatycznie")
        anthropic_key = _anthropic_from_secrets
    else:
        anthropic_key = st.text_input(
            "🔑 Klucz API Anthropic",
            type="password",
            placeholder="sk-ant-...",
            help="Wymagany do generowania przez Claude 3.5 Sonnet"
        )

    if _llama_from_secrets:
        st.success("🦙 Klucz LlamaParse: wczytany automatycznie")
        llama_key = _llama_from_secrets
    else:
        llama_key = st.text_input(
            "🦙 Klucz API LlamaParse (opcjonalny)",
            type="password",
            placeholder="llx-...",
            help="Dla lepszej ekstrakcji tabel. Bez niego użyty zostanie pypdf."
        )

    st.divider()
    st.subheader("🏢 Dane jednostki")

    # ── Pobieranie z KRS po NIP ──────────────────────────────────────────
    krs_input = st.text_input(
        "🔍 Numer KRS spółki",
        placeholder="0000123456",
        help="Wpisz 10-cyfrowy numer KRS. Znajdziesz go na prs.ms.gov.pl lub w dokumentach spółki."
    )
    st.caption("ℹ️ Oficjalne API KRS działa po numerze KRS (nie NIP). NIP uzupełnisz ręcznie.")
    debug_krs = st.checkbox("🔍 Tryb diagnostyczny KRS", value=False,
                             help="Pokaż surową odpowiedź API KRS — pomocne przy błędach")

    if st.button("⬇️ Pobierz dane z KRS", use_container_width=True):
        if krs_input:
            with st.spinner("Pobieranie z API KRS Ministerstwa Sprawiedliwości..."):
                try:
                    if debug_krs:
                        krs_data, krs_debug_log = fetch_krs_by_krs_nr_debug(krs_input)
                        st.code(krs_debug_log)
                    else:
                        krs_data = fetch_krs_by_krs_nr(krs_input)
                    if krs_data:
                        st.session_state["krs_data"] = krs_data
                        st.success("✅ Dane pobrane z KRS!")
                    else:
                        st.error("❌ Nie znaleziono. Sprawdź numer KRS lub uzupełnij ręcznie.")
                except ConnectionError:
                    st.error("❌ Brak połączenia z API KRS.")
                except TimeoutError:
                    st.error("❌ API KRS nie odpowiada. Spróbuj za chwilę.")
                except Exception as e:
                    st.error(f"❌ Błąd: {e}")
        else:
            st.warning("Wpisz numer KRS aby pobrać dane.")

    krs = st.session_state.get("krs_data", {})

    company_name = st.text_input("Nazwa spółki",
                                  value=krs.get("nazwa", ""),
                                  placeholder="XYZ Sp. z o.o.")
    company_siedziba = st.text_input("Siedziba",
                                      value=krs.get("siedziba", ""),
                                      placeholder="ul. Przykładowa 1, 00-001 Warszawa")
    company_nip = st.text_input("NIP",
                                 value=krs.get("nip", ""),
                                 placeholder="1234567890")
    company_krs = st.text_input("Nr KRS",
                                 value=krs.get("krs", krs_input if krs_input else ""),
                                 placeholder="0000000000")
    company_regon = st.text_input("REGON",
                                   value=krs.get("regon", ""),
                                   placeholder="000000000")
    company_pkd = st.text_input("Główne PKD",
                                 value=krs.get("pkd", ""),
                                 placeholder="np. 69.20.Z Działalność rachunkowo-księgowa")
    company_data_rej = st.text_input("Data rejestracji w KRS",
                                      value=krs.get("data_rejestracji", ""),
                                      placeholder="RRRR-MM-DD")
    company_forma = st.text_input("Forma prawna",
                                   value=krs.get("forma_prawna", ""),
                                   placeholder="np. SPÓŁKA Z OGRANICZONĄ ODPOWIEDZIALNOŚCIĄ")

    # Wyświetl wspólników z KRS
    krs_wspolnicy = krs.get("wspolnicy", [])
    krs_kapital = krs.get("kapital_podstawowy", "")
    if krs_kapital or krs_wspolnicy:
        st.divider()
        st.subheader("👥 Dane wspólników (z KRS)")
        if krs_kapital:
            st.markdown(f"**Kapitał podstawowy:** {krs_kapital} PLN")
        if krs_wspolnicy:
            for w in krs_wspolnicy:
                l = w['liczba_udzialow']
                v = w['wartosc_udzialow']
                if l == "wkład":
                    st.markdown(f"- **{w['nazwa']}** — wkład {v} PLN")
                elif l and v:
                    st.markdown(f"- **{w['nazwa']}** — {l} udziałów, wartość {v} PLN")
                elif v:
                    st.markdown(f"- **{w['nazwa']}** — wartość {v} PLN")
                else:
                    st.markdown(f"- **{w['nazwa']}**")
        else:
            st.warning("⚠️ API KRS nie zwróciło danych wspólników.")

    st.divider()
    st.subheader("📅 Okres sprawozdawczy")
    okres_od = st.date_input("Od", value=date(date.today().year - 1, 1, 1))
    okres_do = st.date_input("Do", value=date(date.today().year - 1, 12, 31))
    fiscal_year = okres_do.year

    st.divider()
    st.subheader("⚠️ Kontynuacja działalności")
    zagrozenie_kontynuacji = st.checkbox(
        "Istnieją okoliczności wskazujące na zagrożenie kontynuowania działalności "
        "w okresie co najmniej 12 miesięcy od dnia bilansowego",
        value=False,
        help="Art. 5 ust. 2 UoR — zaznacz jeśli istnieją takie okoliczności"
    )
    if zagrozenie_kontynuacji:
        zagrozenie_opis = st.text_area(
            "Opis okoliczności zagrożenia:",
            placeholder="Opisz okoliczności wskazujące na zagrożenie kontynuowania działalności...",
            height=100
        )
    else:
        zagrozenie_opis = ""

    st.divider()
    st.subheader("👥 Zatrudnienie")
    zatrudnienie_biezacy = st.number_input(
        f"Średnie zatrudnienie — rok bieżący (etaty)",
        min_value=0, value=0, step=1,
        help="Średnia liczba pracowników w etatach przeliczeniowych w roku obrotowym"
    )
    zatrudnienie_poprzedni = st.number_input(
        f"Średnie zatrudnienie — rok poprzedni (etaty)",
        min_value=0, value=0, step=1,
        help="Średnia liczba pracowników w etatach przeliczeniowych w roku poprzednim"
    )
    zatrudnienie_uwagi = st.text_input(
        "Uwagi do zatrudnienia (opcjonalnie)",
        placeholder="np. w tym 2 osoby na umowie zlecenie"
    )
    wynagrodzenie_audytora = st.text_input(
        "Wynagrodzenie firmy audytorskiej (PLN)",
        placeholder="np. 5000 (zostaw puste jeśli nie dotyczy)",
        help="Kwota za badanie sprawozdania — zostaw puste jeśli spółka nie podlega badaniu"
    )

    st.divider()
    st.markdown("""
    **📋 Obsługiwane dokumenty:**
    - 🏦 Bilans
    - 📈 Rachunek Zysków i Strat
    - 🏗️ Tabela środków trwałych
    - 💸 Przepływy pieniężne
    - 📜 Polityka rachunkowości
    - ⚖️ Zestawienie Obrotów i Sald (ZOiS)
    - 📋 Ankieta bilansowa (od klienta)
    """)


# ── Główna sekcja ────────────────────────────────────────────────────────────
col1, col2 = st.columns([1, 1])

with col1:
    st.markdown('<div class="step-card"><b>📁 Krok 1:</b> Wgraj dokumenty sprawozdania (PDF / DOCX)</div>',
                unsafe_allow_html=True)
    uploaded_files = st.file_uploader(
        "Wybierz pliki PDF / DOCX / XLSX",
        type=["pdf", "docx", "xlsx"],
        accept_multiple_files=True,
        help="Wgraj dokumenty: bilans, RZiS, ŚT, ZOiS, ankietę bilansową itp."
    )

    if uploaded_files:
        st.success(f"✅ Wgrano {len(uploaded_files)} plik(ów)")
        for f in uploaded_files:
            size_kb = len(f.getvalue()) // 1024
            st.caption(f"📄 {f.name} ({size_kb} KB)")

with col2:
    st.markdown('<div class="step-card"><b>🔍 Krok 2:</b> Walidacja i mapowanie dokumentów</div>',
                unsafe_allow_html=True)

    if not anthropic_key:
        st.info("👈 Wprowadź klucz API Anthropic w panelu bocznym, aby kontynuować.")
    elif not uploaded_files:
        st.info("👈 Wgraj pliki PDF, aby rozpocząć.")
    elif not company_name:
        st.warning("⚠️ Wprowadź nazwę spółki w panelu bocznym.")

# ── Przycisk uruchomienia ────────────────────────────────────────────────────
st.divider()

run_disabled = not (anthropic_key and uploaded_files and company_name)
if st.button("🚀 Generuj Informację Dodatkową", type="primary",
              disabled=run_disabled, use_container_width=True):

    progress_bar = st.progress(0)
    status_text = st.empty()
    results_container = st.container()

    try:
        # ── KROK 1: Parsowanie ──────────────────────────────────────────────
        status_text.info("📄 Krok 1/5: Parsowanie dokumentów...")
        progress_bar.progress(10)

        def update_progress(val, msg):
            progress_bar.progress(int(10 + val * 20))
            status_text.info(f"📄 {msg}")

        # Rozdziel PDF, DOCX i XLSX
        pdf_files = [f for f in uploaded_files if f.name.lower().endswith(".pdf")]
        docx_files = [f for f in uploaded_files if f.name.lower().endswith(".docx")]
        xlsx_files = [f for f in uploaded_files if f.name.lower().endswith(".xlsx")]

        parsed = {}
        if pdf_files:
            if llama_key:
                parsed = parse_documents_llamaparse(pdf_files, llama_key, update_progress)
            else:
                parsed = parse_documents_fallback(pdf_files, update_progress)
        for df in docx_files:
            parsed[df.name] = extract_text_from_docx(df.getvalue())
        for xf in xlsx_files:
            parsed[xf.name] = extract_text_from_xlsx(xf.getvalue())

        progress_bar.progress(30)
        st.session_state["parsed_docs"] = parsed

        # ── KROK 2: Mapowanie ───────────────────────────────────────────────
        status_text.info("🗂️ Krok 2/5: Mapowanie i identyfikacja dokumentów...")
        progress_bar.progress(40)
        doc_mapping = map_documents(parsed)
        st.session_state["doc_mapping"] = doc_mapping

        # ── SPRAWDZENIE BRAKUJĄCYCH DOKUMENTÓW ─────────────────────────────
        missing = check_missing_documents(doc_mapping)
        if missing:
            progress_bar.empty()
            status_text.empty()

            st.warning("⚠️ Nie znaleziono wszystkich dokumentów w wgranych plikach.")

            st.markdown("**Brakujące dokumenty:**")
            for dt in missing:
                info = REQUIRED_DOC_TYPES[dt]
                st.markdown(
                    f"- {info['icon']} **{info['label']}** — {info['desc']}"
                )

            st.markdown("---")
            st.markdown("**Co chcesz zrobić?**")

            col_a, col_b = st.columns(2)
            with col_a:
                if st.button("▶️ Kontynuuj bez brakujących dokumentów",
                              use_container_width=True, type="primary"):
                    st.session_state["missing_decision"] = "continue"
                    st.rerun()
            with col_b:
                if st.button("📁 Anuluj — chcę dodać brakujące pliki",
                              use_container_width=True):
                    st.session_state["missing_decision"] = "cancel"
                    st.rerun()

            st.info(
                "💡 Wskazówka: Jeśli plik zawiera kilka dokumentów w jednym PDF "
                "(np. Bilans + RZiS razem), aplikacja może nie rozpoznać drugiego. "
                "Spróbuj wgrać je jako osobne pliki."
            )
            st.stop()

        # Jeśli użytkownik wrócił po decyzji
        decision = st.session_state.pop("missing_decision", None)
        if decision == "cancel":
            st.info("Wgraj brakujące pliki i uruchom ponownie.")
            st.stop()

        # (decision == "continue" lub brak braków — kontynuuj normalnie)

        # ── PYTANIA O POLITYKĘ RACHUNKOWOŚCI (gdy brak dokumentu) ──────────
        types_found = {d["type"] for d in doc_mapping.values()}
        if "POLITYKA RACHUNKOWOŚCI" not in types_found:
            # Sprawdź czy pytania już zostały wypełnione
            if not st.session_state.get("polityka_answers"):
                progress_bar.empty()
                status_text.empty()

                st.warning(
                    "📜 Nie załączono dokumentu **Polityki Rachunkowości**. "
                    "Odpowiedz na poniższe pytania — zostaną one wykorzystane "
                    "przy sporządzaniu sekcji 1.2–1.5 Informacji Dodatkowej."
                )

                with st.form("polityka_form"):
                    st.subheader("📋 Zasady rachunkowości — pytania uzupełniające")

                    q1 = st.selectbox(
                        "1. Zasady ustalania wyniku finansowego:",
                        options=[
                            "Wariant porównawczy (układ rodzajowy kosztów)",
                            "Wariant kalkulacyjny (układ funkcjonalny kosztów)",
                        ],
                        help="Dotyczy formy Rachunku Zysków i Strat (art. 47 UoR)"
                    )

                    q2_wycena = st.selectbox(
                        "2a. Metoda wyceny zapasów:",
                        options=[
                            "FIFO (pierwsze weszło, pierwsze wyszło)",
                            "LIFO (ostatnie weszło, pierwsze wyszło)",
                            "Cena przeciętna (średnia ważona)",
                            "Ceny ewidencyjne z odchyleniami",
                            "Nie dotyczy (brak zapasów)",
                        ]
                    )

                    q2_st = st.selectbox(
                        "2b. Metoda amortyzacji środków trwałych:",
                        options=[
                            "Liniowa (równomierne odpisy przez cały okres)",
                            "Degresywna (przyspieszone odpisy na początku)",
                            "Jednorazowy odpis (niskocenne ST do 10 000 zł)",
                            "Mieszana (liniowa i jednorazowa)",
                        ]
                    )

                    q2_nal = st.selectbox(
                        "2c. Wycena należności:",
                        options=[
                            "W wartości nominalnej z odpisami aktualizującymi",
                            "W wartości nominalnej bez odpisów aktualizujących",
                            "W wartości godziwej",
                        ]
                    )

                    q3 = st.selectbox(
                        "3. Sposób sporządzania sprawozdania finansowego:",
                        options=[
                            "Pełne sprawozdanie finansowe (standardowe)",
                            "Uproszczone sprawozdanie finansowe (art. 46 ust. 5 UoR — jednostki małe)",
                            "Sprawozdanie według Załącznika nr 4 UoR (mikro jednostki)",
                            "Sprawozdanie według Załącznika nr 5 UoR (małe jednostki NGO)",
                        ],
                        help="Jednostki małe mogą stosować uproszczenia zgodnie z art. 46–50 UoR"
                    )

                    q4_podatek = st.checkbox(
                        "Jednostka tworzy rezerwę i aktywa z tytułu odroczonego podatku dochodowego",
                        value=True
                    )

                    q5_leasing = st.selectbox(
                        "Ujęcie leasingu:",
                        options=[
                            "Według UoR (leasing operacyjny/finansowy wg ekonomicznej treści)",
                            "Leasing operacyjny — wszystkie umowy traktowane jako operacyjny",
                            "Nie dotyczy (brak umów leasingowych)",
                        ]
                    )

                    uwagi = st.text_area(
                        "Dodatkowe uwagi dotyczące polityki rachunkowości (opcjonalnie):",
                        placeholder="np. szczególne zasady wyceny, zmiany polityki w roku obrotowym...",
                        height=80
                    )

                    submitted = st.form_submit_button(
                        "✅ Zatwierdź i kontynuuj generowanie",
                        use_container_width=True, type="primary"
                    )

                if submitted:
                    st.session_state["polityka_answers"] = {
                        "wynik_finansowy": q1,
                        "wycena_zapasow": q2_wycena,
                        "amortyzacja": q2_st,
                        "wycena_naleznosci": q2_nal,
                        "sposob_sprawozdania": q3,
                        "podatek_odroczony": q4_podatek,
                        "leasing": q5_leasing,
                        "uwagi": uwagi,
                    }
                    st.rerun()
                else:
                    st.stop()

        # Pobierz odpowiedzi na pytania (jeśli były zadane)
        polityka_answers = st.session_state.get("polityka_answers", {})

        # ── KROK 3: Walidacja ───────────────────────────────────────────────
        status_text.info("✅ Krok 3/5: Walidacja spójności danych...")
        progress_bar.progress(55)
        validation_issues = validate_data_consistency(doc_mapping)

        with results_container:
            st.subheader("📋 Raport mapowania i walidacji")
            map_cols = st.columns(min(len(doc_mapping), 4))
            for i, (fname, ddata) in enumerate(doc_mapping.items()):
                with map_cols[i % len(map_cols)]:
                    st.markdown(f"""
                    <div class="metric-box">
                        <b>{ddata['type']}</b><br>
                        <small>{fname}</small><br>
                        <small>{ddata['length']:,} znaków</small>
                    </div>""", unsafe_allow_html=True)

            st.subheader("🔎 Walidacja danych")
            for issue in validation_issues:
                css = {"OK": "validation-ok", "WARN": "validation-warn", "ERR": "validation-err"}
                st.markdown(f'<span class="{css.get(issue["level"], "")}">{issue["msg"]}</span>',
                            unsafe_allow_html=True)

        # ── KROK 3B: Dobór not objaśniających ────────────────────────────
        status_text.info("📋 Krok 3b/5: Dobór not objaśniających...")
        progress_bar.progress(58)
        _ci_for_notes = {"forma_prawna": company_forma}
        selected_notes = select_applicable_notes(doc_mapping, _ci_for_notes)

        with results_container:
            st.subheader(f"📝 Dobrano {len(selected_notes)} not objaśniających")
            for n in selected_notes:
                prio = {1: "🔴", 2: "🟡", 3: "⚪"}.get(n["priority"], "")
                st.markdown(f'{prio} Nota {n["nr"]}: {n["name"]} — {n["reason"]}')

        # ── KROK 4: Generowanie ─────────────────────────────────────────────
        status_text.info("🤖 Krok 4/5: Generowanie przez Claude (może potrwać 1-2 min)...")
        progress_bar.progress(65)

        company_info = {
            "nazwa": company_name,
            "siedziba": company_siedziba,
            "nip": company_nip,
            "krs": company_krs,
            "regon": company_regon,
            "pkd": company_pkd,
            "data_rejestracji": company_data_rej,
            "forma_prawna": company_forma,
            "okres_od": str(okres_od),
            "okres_do": str(okres_do),
            "zagrozenie_kontynuacji": zagrozenie_kontynuacji,
            "zagrozenie_opis": zagrozenie_opis,
            "polityka_answers": polityka_answers,
            "zatrudnienie_biezacy": zatrudnienie_biezacy,
            "zatrudnienie_poprzedni": zatrudnienie_poprzedni,
            "zatrudnienie_uwagi": zatrudnienie_uwagi,
            "selected_notes": selected_notes,
            "wynagrodzenie_audytora": wynagrodzenie_audytora,
            "kapital_podstawowy": krs.get("kapital_podstawowy", ""),
            "wspolnicy": krs.get("wspolnicy", []),
        }
        generated_text = generate_accounting_notes(
            doc_mapping=doc_mapping,
            anthropic_api_key=anthropic_key,
            company_name=company_name,
            year=fiscal_year,
            company_info=company_info,
            progress_callback=lambda v, m: progress_bar.progress(int(65 + v * 20))
        )
        st.session_state["generated_text"] = _strip_chart_data(generated_text)
        progress_bar.progress(88)

        # ── KROK 5: Eksport do Word ─────────────────────────────────────────
        status_text.info("💾 Krok 5/5: Generowanie pliku Word...")
        docx_bytes = save_to_word(generated_text, company_name, fiscal_year, company_info)
        st.session_state["docx_bytes"] = docx_bytes
        progress_bar.progress(100)
        status_text.success("✅ Informacja Dodatkowa wygenerowana pomyślnie!")

        # ── Podgląd i pobieranie ────────────────────────────────────────────
        with results_container:
            st.divider()
            dl_col, _ = st.columns([1, 2])
            with dl_col:
                st.download_button(
                    label="⬇️ Pobierz Informację Dodatkową (.docx)",
                    data=docx_bytes,
                    file_name=f"informacja_dodatkowa_{company_name.replace(' ', '_')}_{fiscal_year}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    type="primary",
                    use_container_width=True
                )

            with st.expander("👁️ Podgląd wygenerowanej treści", expanded=True):
                st.markdown(_strip_chart_data(generated_text))

    except anthropic.AuthenticationError:
        st.session_state.pop("run_generation", None)
        st.error("❌ Nieprawidłowy klucz API Anthropic. Sprawdź wartość w panelu bocznym.")
    except anthropic.RateLimitError:
        st.session_state.pop("run_generation", None)
        st.error("❌ Przekroczono limit zapytań API. Poczekaj chwilę i spróbuj ponownie.")
    except Exception as e:
        st.session_state.pop("run_generation", None)
        st.error(f"❌ Błąd: {e}")
        st.exception(e)

# ── Jeśli wyniki już są w sesji ──────────────────────────────────────────────
elif "generated_text" in st.session_state:
    st.info("📝 Wyniki z poprzedniego uruchomienia (wgraj nowe pliki lub wciśnij Generuj ponownie).")
    if st.session_state.get("docx_bytes"):
        st.download_button(
            label="⬇️ Pobierz poprzedni wynik (.docx)",
            data=st.session_state["docx_bytes"],
            file_name=f"informacja_dodatkowa_{fiscal_year}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    with st.expander("👁️ Poprzednio wygenerowana treść"):
        st.markdown(st.session_state["generated_text"])
