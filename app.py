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
matplotlib.use("Agg")  # Backend bez GUI
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


def extract_text_from_docx(docx_bytes: bytes, filename: str) -> str:
    """Ekstrakcja tekstu z pliku DOCX."""
    try:
        doc = Document(io.BytesIO(docx_bytes))
        text_parts = [f"=== DOKUMENT: {filename} ===\n"]
        for para in doc.paragraphs:
            if para.text.strip():
                text_parts.append(para.text)
        # Wyciągnij też tekst z tabel
        for table in doc.tables:
            for row in table.rows:
                row_text = " | ".join(cell.text.strip() for cell in row.cells if cell.text.strip())
                if row_text:
                    text_parts.append(row_text)
        return "\n".join(text_parts)
    except Exception as e:
        return f"[BŁĄD ekstrakcji DOCX {filename}: {e}]"


def parse_documents_llamaparse(pdf_files: list, llama_api_key: str, progress_callback=None) -> dict:
    """
    Krok 1 & 2: Parsowanie PDF przez LlamaParse + identyfikacja dokumentów.
    Pliki DOCX parsowane są bezpośrednio (bez LlamaParse).
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

            # DOCX — parsuj bezpośrednio
            if uploaded_file.name.lower().endswith(".docx"):
                results[uploaded_file.name] = extract_text_from_docx(
                    uploaded_file.getvalue(), uploaded_file.name
                )
                continue

            # PDF — przez LlamaParse
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
    """Fallback: ekstrakcja przez pypdf (PDF) lub python-docx (DOCX)."""
    results = {}
    for idx, uploaded_file in enumerate(pdf_files):
        if progress_callback:
            progress_callback(idx / len(pdf_files), f"Ekstrakcja: {uploaded_file.name}")

        if uploaded_file.name.lower().endswith(".docx"):
            results[uploaded_file.name] = extract_text_from_docx(
                uploaded_file.getvalue(), uploaded_file.name
            )
        else:
            results[uploaded_file.name] = extract_text_from_pdf_basic(
                uploaded_file.getvalue(), uploaded_file.name
            )
    return results



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

        return {
            "nazwa": nazwa,
            "siedziba": siedziba,
            "nip": nip_val,
            "krs": krs_val,
            "regon": regon_val,
            "pkd": pkd,
            "data_rejestracji": data_rej,
            "forma_prawna": forma,
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
        "required": True,
    },
    "RZiS": {
        "label": "Rachunek Zysków i Strat",
        "icon": "📈",
        "desc": "Przychody, koszty i wynik finansowy za rok obrotowy",
        "keywords": ["przychody ze sprzedaży", "koszty działalności", "zysk netto", "wynik finansowy", "amortyzacja"],
        "required": True,
    },
    "ŚRODKI TRWAŁE": {
        "label": "Tabela środków trwałych",
        "icon": "🏗️",
        "desc": "Wartość brutto, umorzenia i wartość netto środków trwałych",
        "keywords": ["środki trwałe", "wartość brutto", "umorzenie", "odpisy amortyzacyjne"],
        "required": True,
    },
    "PRZEPŁYWY PIENIĘŻNE": {
        "label": "Rachunek przepływów pieniężnych",
        "icon": "💸",
        "desc": "Cash flow: operacyjny, inwestycyjny, finansowy",
        "keywords": ["przepływy", "działalność operacyjna", "działalność inwestycyjna"],
        "required": False,
    },
    "POLITYKA RACHUNKOWOŚCI": {
        "label": "Polityka rachunkowości",
        "icon": "📜",
        "desc": "Przyjęte zasady rachunkowości, metody wyceny, okresy amortyzacji",
        "keywords": ["polityka rachunkowości", "zasady rachunkowości", "metody wyceny",
                     "okres amortyzacji", "przyjęte zasady", "opis przyjętych"],
        "required": False,
    },
    "ZOiS": {
        "label": "Zestawienie Obrotów i Sald",
        "icon": "⚖️",
        "desc": "Obroty i salda kont księgi głównej za rok obrotowy",
        "keywords": ["zestawienie obrotów", "obroty i salda", "salda końcowe",
                     "salda otwarcia", "obroty narastająco", "konta syntetyczne",
                     "księga główna", "salda debetowe", "salda kredytowe",
                     "obroty wn", "obroty ma", "saldo wn", "saldo ma",
                     "bilans otwarcia", "obroty za okres", "saldo końcowe",
                     "konto", "zespół 0", "zespół 1", "zespół 2", "zespół 4",
                     "zespół 5", "zespół 7",
                     "rozrachunki", "koszty według rodzajów",
                     "bo wn", "bo ma"],
        "required": True,
    },
    "ANKIETA BILANSOWA": {
        "label": "Ankieta bilansowa",
        "icon": "📝",
        "desc": "Odpowiedzi klienta dot. zobowiązań warunkowych, kontynuacji działalności, podziału wyniku itp.",
        "keywords": ["ankieta bilansowa", "propozycja podziału zysku",
                     "propozycja pokrycia straty", "zobowiązania warunkowe",
                     "gwarancji i poręczeń", "kontynuować działalność",
                     "postępowaniu egzekucyjnym", "transakcje ze stronami powiązanymi",
                     "nakłady na niefinansowe aktywa trwałe",
                     "zdarzenia istotnie wpływające",
                     "należności wątpliwe", "odsetki zwłoki",
                     "pożyczek i świadczeń o podobnym charakterze",
                     "zabezpieczenia majątkowe",
                     "organy nadzorujące i zarządzające",
                     "prognoza rozwoju spółki", "sytuacja finansowa jest"],
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
    """Zwraca listę typów dokumentów których brakuje wśród wgranych plików.
    Uwzględnia tylko dokumenty oznaczone jako required=True."""
    types_found = {d["type"] for d in doc_mapping.values()}
    return [dt for dt, info in REQUIRED_DOC_TYPES.items()
            if info.get("required", True) and dt not in types_found]


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
        elif info.get("required", True):
            issues.append({"level": "WARN", "msg": f"⚠️ Brak dokumentu: {info['icon']} {info['label']}"})
        else:
            issues.append({"level": "WARN", "msg": f"ℹ️ Opcjonalny, nie wgrano: {info['icon']} {info['label']}"}) 

    return issues


# ═══════════════════════════════════════════════════════════════════════════════
# MODUŁ 3B: DOBÓR NOT OBJAŚNIAJĄCYCH
# ═══════════════════════════════════════════════════════════════════════════════

# Reguły doboru not: każda nota ma warunki generowania
# "source": wymagane typy dokumentów (OR - wystarczy jeden)
# "trigger": funkcja sprawdzająca czy nota ma być generowana
# "category": "auto" | "ankieta" | "warunkowe" | "specjalne"
# "priority": 1=rdzeń, 2=ważne, 3=opcjonalne

NOTA_RULES = {
    1:  {"name": "Zmiana wartości początkowej i umorzenia ŚT",
         "source": ["ŚRODKI TRWAŁE", "ZOiS"], "category": "auto", "priority": 1},
    2:  {"name": "Zmiana wartości początkowej i umorzenia WNiP",
         "source": ["ŚRODKI TRWAŁE", "ZOiS"], "category": "auto", "priority": 1},
    3:  {"name": "Zmiana wartości inwestycji długoterminowych",
         "source": ["ZOiS"], "category": "warunkowe", "priority": 2,
         "zois_keywords": ["inwestycje długoterminowe", "03"]},
    4:  {"name": "Odpisy aktualizujące wartość długoterminowych aktywów niefinansowych",
         "source": ["ZOiS"], "category": "warunkowe", "priority": 3,
         "zois_keywords": ["odpis", "aktualizuj"]},
    5:  {"name": "Odpisy aktualizujące wartość długoterminowych aktywów finansowych",
         "source": ["ZOiS"], "category": "warunkowe", "priority": 3,
         "zois_keywords": ["odpis", "aktyw", "finansow"]},
    6:  {"name": "Koszty zakończonych prac rozwojowych oraz wartość firmy",
         "source": ["ZOiS"], "category": "warunkowe", "priority": 2,
         "zois_keywords": ["prace rozwojowe", "wartość firmy", "011"]},
    7:  {"name": "Grunty użytkowane wieczyście",
         "source": [], "category": "warunkowe", "priority": 3},
    8:  {"name": "Środki trwałe nieamortyzowane (pozabilansowe)",
         "source": [], "category": "warunkowe", "priority": 3},
    9:  {"name": "Papiery wartościowe lub prawa",
         "source": ["ZOiS"], "category": "warunkowe", "priority": 3,
         "zois_keywords": ["papiery wartościowe", "03"]},
    10: {"name": "Odpisy aktualizujące wartość należności",
         "source": ["ZOiS"], "category": "auto", "priority": 1,
         "zois_keywords": ["290", "odpis", "należności"]},
    11: {"name": "Struktura własności kapitału podstawowego (S.A.)",
         "source": [], "category": "warunkowe", "priority": 2,
         "forma_prawna": ["SPÓŁKA AKCYJNA", "PROSTA SPÓŁKA AKCYJNA"]},
    12: {"name": "Struktura własności kapitału podstawowego (sp. z o.o.)",
         "source": [], "category": "auto", "priority": 1,
         "forma_prawna": ["SPÓŁKA Z OGRANICZONĄ ODPOWIEDZIALNOŚCIĄ"]},
    13: {"name": "Zmiany stanów kapitałów zapasowego i rezerwowego",
         "source": ["ZOiS", "BILANS"], "category": "auto", "priority": 1},
    14: {"name": "Zmiany w stanie kapitału z aktualizacji wyceny",
         "source": ["ZOiS"], "category": "warunkowe", "priority": 2,
         "zois_keywords": ["803", "aktualizacja wyceny"]},
    15: {"name": "Propozycja podziału zysku za rok obrotowy",
         "source": ["ANKIETA BILANSOWA"], "category": "ankieta", "priority": 1,
         "ankieta_trigger": "q6_zysk"},
    16: {"name": "Propozycja pokrycia straty za rok obrotowy",
         "source": ["ANKIETA BILANSOWA"], "category": "ankieta", "priority": 1,
         "ankieta_trigger": "q7_strata"},
    17: {"name": "Rezerwy na koszty i zobowiązania",
         "source": ["ZOiS", "BILANS"], "category": "auto", "priority": 1},
    18: {"name": "Odroczony podatek dochodowy",
         "source": ["ZOiS"], "category": "auto", "priority": 1,
         "zois_keywords": ["650", "841", "odroczony"]},
    19: {"name": "Zobowiązania według okresów wymagalności",
         "source": ["ZOiS", "BILANS"], "category": "auto", "priority": 1},
    20: {"name": "Wykaz zobowiązań zabezpieczonych na majątku",
         "source": ["ANKIETA BILANSOWA"], "category": "ankieta", "priority": 2,
         "ankieta_trigger": "q8_zobowiazania_warunkowe"},
    21: {"name": "Czynne rozliczenia międzyokresowe",
         "source": ["ZOiS"], "category": "auto", "priority": 1,
         "zois_keywords": ["640", "rozliczenia międzyokresowe"]},
    22: {"name": "Rozliczenia międzyokresowe przychodów",
         "source": ["ZOiS"], "category": "auto", "priority": 1,
         "zois_keywords": ["840", "845", "rozliczenia międzyokresowe przychod"]},
    23: {"name": "Składniki aktywów w więcej niż jednej pozycji bilansu",
         "source": ["ZOiS"], "category": "warunkowe", "priority": 3},
    24: {"name": "Składniki pasywów w więcej niż jednej pozycji bilansu",
         "source": ["ZOiS"], "category": "warunkowe", "priority": 3},
    25: {"name": "Wykaz zobowiązań warunkowych",
         "source": ["ANKIETA BILANSOWA"], "category": "ankieta", "priority": 2,
         "ankieta_trigger": "q8_zobowiazania_warunkowe"},
    26: {"name": "Wykaz zobowiązań warunkowych zabezpieczonych na majątku",
         "source": ["ANKIETA BILANSOWA"], "category": "ankieta", "priority": 3,
         "ankieta_trigger": "q8_zobowiazania_warunkowe"},
    29: {"name": "Struktura rzeczowa i terytorialna przychodów",
         "source": ["RZiS", "ZOiS"], "category": "auto", "priority": 1},
    31: {"name": "Koszty rodzajowe (wariant kalkulacyjny)",
         "source": ["ZOiS"], "category": "warunkowe", "priority": 2},
    32: {"name": "Odpisy aktualizujące wartość środków trwałych",
         "source": ["ŚRODKI TRWAŁE"], "category": "warunkowe", "priority": 2},
    33: {"name": "Odpisy aktualizujące wartość zapasów",
         "source": ["ZOiS"], "category": "warunkowe", "priority": 2,
         "zois_keywords": ["340", "odpis", "zapas"]},
    35: {"name": "Rozliczenie różnicy CIT vs wynik finansowy",
         "source": ["RZiS", "ZOiS"], "category": "auto", "priority": 1},
    36: {"name": "Koszt wytworzenia środków trwałych w budowie",
         "source": ["ZOiS"], "category": "warunkowe", "priority": 2,
         "zois_keywords": ["080", "środki trwałe w budowie"]},
    38: {"name": "Nakłady na niefinansowe aktywa trwałe",
         "source": ["ANKIETA BILANSOWA", "ZOiS"], "category": "ankieta", "priority": 2,
         "ankieta_trigger": "q18_naklady"},
    40: {"name": "Kursy walut przyjęte do wyceny",
         "source": ["ZOiS"], "category": "warunkowe", "priority": 2,
         "zois_keywords": ["walut", "kursow", "EUR", "USD", "GBP"]},
    41: {"name": "Struktura środków pieniężnych",
         "source": ["PRZEPŁYWY PIENIĘŻNE", "ZOiS"], "category": "auto", "priority": 1},
    42: {"name": "Przepływy pieniężne netto — metoda pośrednia",
         "source": ["PRZEPŁYWY PIENIĘŻNE"], "category": "auto", "priority": 2},
    43: {"name": "Przeciętne zatrudnienie w podziale na grupy zawodowe",
         "source": [], "category": "warunkowe", "priority": 2},
    44: {"name": "Wynagrodzenia organów spółki",
         "source": [], "category": "warunkowe", "priority": 2},
    46: {"name": "Zaliczki, kredyty, pożyczki dla organów",
         "source": ["ANKIETA BILANSOWA"], "category": "ankieta", "priority": 2,
         "ankieta_trigger": "q16_pozyczki"},
    47: {"name": "Wynagrodzenie firmy audytorskiej",
         "source": [], "category": "warunkowe", "priority": 2},
    48: {"name": "Błędy lat ubiegłych odnoszone na kapitał",
         "source": [], "category": "warunkowe", "priority": 3},
    49: {"name": "Skutki zmian polityki rachunkowości",
         "source": [], "category": "warunkowe", "priority": 3},
    57: {"name": "Różnica zobowiązań krótkoterminowych (bilans vs przepływy)",
         "source": ["BILANS", "PRZEPŁYWY PIENIĘŻNE"], "category": "auto", "priority": 2},
    58: {"name": "Różnica zapasów (bilans vs przepływy)",
         "source": ["BILANS", "PRZEPŁYWY PIENIĘŻNE"], "category": "auto", "priority": 2},
    59: {"name": "Ustalenie faktycznie zapłaconego podatku dochodowego",
         "source": ["ZOiS", "RZiS"], "category": "auto", "priority": 2},
    60: {"name": "Struktura należności",
         "source": ["ZOiS"], "category": "auto", "priority": 1},
    61: {"name": "Należności według okresów wymagalności",
         "source": ["ZOiS"], "category": "auto", "priority": 1},
    63: {"name": "Środki pieniężne na rachunku VAT",
         "source": ["ZOiS"], "category": "warunkowe", "priority": 2,
         "zois_keywords": ["VAT", "rachunek VAT"]},
    68: {"name": "Udziały (akcje) własne",
         "source": [], "category": "warunkowe", "priority": 3},
    72: {"name": "Gwarancje i poręczenia dla organów",
         "source": ["ANKIETA BILANSOWA"], "category": "ankieta", "priority": 2,
         "ankieta_trigger": "q9_gwarancje"},
    73: {"name": "Zobowiązania długoterminowe > 5 lat",
         "source": ["ZOiS", "BILANS"], "category": "warunkowe", "priority": 2},
    74: {"name": "Informacja o dochodach z tytułu ukrytych zysków",
         "source": [], "category": "warunkowe", "priority": 3},
    76: {"name": "Informacje o transakcjach z jednostkami powiązanymi",
         "source": ["ANKIETA BILANSOWA"], "category": "ankieta", "priority": 2,
         "ankieta_trigger": "q14_powiazane"},
}

# Mapowanie triggerów ankiety na słowa kluczowe w tekście ankiety
ANKIETA_TRIGGERS = {
    "q6_zysk": {
        "positive": ["przeznaczenie zysku", "wypłata dywidendy", "kapitał zapasowy",
                      "podwyższenie kapitału"],
        "negative": [],
    },
    "q7_strata": {
        "positive": ["pokrycie straty", "zyskiem z lat", "kapitale zapasowym",
                      "dopłat wniesionych"],
        "negative": [],
    },
    "q8_zobowiazania_warunkowe": {
        "question": "zobowiązania warunkowe",
        "positive_answer": "tak",
    },
    "q9_gwarancje": {
        "question": "gwarancji i poręczeń",
        "positive_answer": "tak",
    },
    "q14_powiazane": {
        "question": "transakcje ze stronami powiązanymi",
        "positive_answer": "tak",
    },
    "q16_pozyczki": {
        "question": "pożyczek i świadczeń o podobnym charakterze",
        "positive_answer": "tak",
    },
    "q18_naklady": {
        "question": "nakłady na niefinansowe aktywa trwałe",
        "positive_answer": "planowane",
    },
}


def _check_ankieta_trigger(trigger_key: str, ankieta_text: str) -> bool:
    """Sprawdza czy ankieta bilansowa triggeruje daną notę."""
    if not ankieta_text:
        return False

    trigger = ANKIETA_TRIGGERS.get(trigger_key, {})
    text_lower = ankieta_text.lower()

    # Dla pytań Tak/Nie — szukamy pytania i odpowiedzi tuż po nim
    if "question" in trigger:
        q_pos = text_lower.find(trigger["question"])
        if q_pos == -1:
            return False
        # Sprawdź fragment tekstu po pytaniu (następne 100 znaków)
        answer_region = text_lower[q_pos:q_pos + 100]
        pos_answer = trigger["positive_answer"]
        pos_pos = answer_region.find(pos_answer)

        if pos_pos == -1:
            return False

        # Jeśli "nie" pojawia się PRZED pozytywną odpowiedzią jako samodzielne słowo — to negacja
        # Uwaga: "nie" jako część innego słowa (np. "niefinansowe") nie liczy się
        import re as _re
        # Szukaj samodzielnego "nie" (z granicami słów) przed pozytywną odpowiedzią
        nie_matches = list(_re.finditer(r'\bnie\b', answer_region))
        if nie_matches:
            earliest_nie = nie_matches[0].start()
            if earliest_nie < pos_pos:
                # Samodzielne "nie" jest przed pozytywną odpowiedzią — to negacja
                return False

        return True

    # Dla pytań z wieloma opcjami (Q6, Q7) — wystarczy że któraś opcja jest zaznaczona
    if "positive" in trigger:
        return any(kw in text_lower for kw in trigger["positive"])

    return False


def _check_forma_prawna(nota_rule: dict, company_info: dict) -> bool:
    """Sprawdza czy forma prawna pasuje do noty."""
    if "forma_prawna" not in nota_rule:
        return True  # Brak ograniczenia
    forma = (company_info.get("forma_prawna", "") or "").upper()
    return any(fp.upper() in forma for fp in nota_rule["forma_prawna"])


def select_applicable_notes(doc_mapping: dict, company_info: dict = None) -> list:
    """
    Na podstawie wgranych dokumentów i ankiety bilansowej
    dobiera listę not objaśniających do wygenerowania.

    Zwraca listę: [{"nr": 1, "name": "...", "category": "auto", "reason": "..."}]
    """
    info = company_info or {}
    types_found = {d["type"] for d in doc_mapping.values()}

    # Wyciągnij tekst ankiety bilansowej (jeśli wgrana)
    ankieta_text = ""
    for doc_data in doc_mapping.values():
        if doc_data["type"] == "ANKIETA BILANSOWA":
            ankieta_text = doc_data["text"]
            break

    # Wyciągnij cały tekst ZOiS (do sprawdzania słów kluczowych)
    zois_text = ""
    for doc_data in doc_mapping.values():
        if doc_data["type"] == "ZOiS":
            zois_text = doc_data["text"].lower()
            break

    selected = []

    for nota_nr, rule in sorted(NOTA_RULES.items()):
        reason = ""
        include = False

        # 1. Sprawdź formę prawną (jeśli nota jest ograniczona do SA/sp.z o.o.)
        if not _check_forma_prawna(rule, info):
            continue

        # 2. Kategoria "auto" — generuj jeśli mamy wymagane dokumenty
        if rule["category"] == "auto":
            sources = rule.get("source", [])
            matched_sources = [s for s in sources if s in types_found]
            if matched_sources:
                include = True
                reason = f"Źródło: {', '.join(matched_sources)}"
            elif not sources:
                # Brak wymaganych źródeł = zawsze generuj
                include = True
                reason = "Nota standardowa"

        # 3. Kategoria "ankieta" — generuj jeśli ankieta triggeruje
        elif rule["category"] == "ankieta":
            trigger_key = rule.get("ankieta_trigger", "")
            if ankieta_text and _check_ankieta_trigger(trigger_key, ankieta_text):
                include = True
                reason = "Trigger z ankiety bilansowej"
            elif not ankieta_text and rule["priority"] <= 1:
                # Brak ankiety — przy priorytetowych notach zaznacz jako "do uzupełnienia"
                include = True
                reason = "Brak ankiety — wymagane dane od klienta"

        # 4. Kategoria "warunkowe" — generuj jeśli mamy źródło + ewentualnie słowa kluczowe w ZOiS
        elif rule["category"] == "warunkowe":
            sources = rule.get("source", [])
            matched_sources = [s for s in sources if s in types_found]
            if matched_sources:
                # Jeśli nota ma słowa kluczowe ZOiS — sprawdź czy występują
                zois_kw = rule.get("zois_keywords", [])
                if zois_kw and zois_text:
                    if any(kw.lower() in zois_text for kw in zois_kw):
                        include = True
                        reason = f"Wykryto dane w ZOiS ({', '.join(matched_sources)})"
                elif not zois_kw:
                    # Warunkowe bez zois_keywords — generuj tylko jeśli
                    # źródło to NIE sam ZOiS (np. ŚRODKI TRWAŁE)
                    non_zois_sources = [s for s in matched_sources if s != "ZOiS"]
                    if non_zois_sources:
                        include = True
                        reason = f"Źródło dostępne: {', '.join(non_zois_sources)}"

        if include:
            selected.append({
                "nr": nota_nr,
                "name": rule["name"],
                "category": rule["category"],
                "priority": rule["priority"],
                "reason": reason,
            })

    return selected


def format_notes_for_prompt(selected_notes: list) -> str:
    """Formatuje listę wybranych not do wstawienia w prompt Claude."""
    if not selected_notes:
        return "\nNie wybrano żadnych not objaśniających do wygenerowania.\n"

    lines = [
        "\n📋 NOTY OBJAŚNIAJĄCE DO WYGENEROWANIA:",
        f"Na podstawie wgranych dokumentów i ankiety bilansowej wybrano {len(selected_notes)} not.",
        "WAŻNE: Numeruj noty SEKWENCYJNIE (Nota 1, Nota 2, Nota 3...) w kolejności poniższej listy.",
        "Numery w nawiasach [GOFIN XX] służą tylko do identyfikacji wzoru — NIE umieszczaj ich w dokumencie.\n",
        "OBLIGATORYJNE (generuj ZAWSZE z danymi):"
    ]

    # Podziel na priorytetowe i opcjonalne
    prio1 = [n for n in selected_notes if n["priority"] == 1]
    prio2 = [n for n in selected_notes if n["priority"] == 2]
    prio3 = [n for n in selected_notes if n["priority"] == 3]

    for idx, n in enumerate(prio1, 1):
        lines.append(f"  ✅ {n['name']} [GOFIN {n['nr']}] — {n['reason']}")

    if prio2:
        lines.append("\nWAŻNE (generuj jeśli dane wystarczające):")
        for n in prio2:
            lines.append(f"  📌 {n['name']} [GOFIN {n['nr']}] — {n['reason']}")

    if prio3:
        lines.append("\nOPCJONALNE (generuj jeśli dane dostępne, pomiń jeśli brak):")
        for n in prio3:
            lines.append(f"  📎 {n['name']} [GOFIN {n['nr']}] — {n['reason']}")

    lines.append(
        "\nINSTRUKCJA: Wygeneruj KAŻDĄ notę z powyższej listy w formie tabeli markdown. "
        "Noty obligatoryjne MUSZĄ być wypełnione danymi z dokumentów. "
        "Jeśli brakuje danych dla noty — wstaw [DANE DO UZUPEŁNIENIA]. "
        "Noty, których NIE MA na liście — NIE generuj.\n"
        "WAŻNE: Jeśli dla danej noty WSZYSTKIE wartości liczbowe wynoszą 0 (zero), "
        "NIE generuj tabeli — zamiast tego napisz krótko: "
        "\"Nota [numer sekwencyjny] — [tytuł noty]: Nie dotyczy.\"\n"
    )

    return "\n".join(lines)


def format_notes_for_display(selected_notes: list) -> list:
    """Formatuje listę not do wyświetlenia w UI Streamlit."""
    display = []
    for n in selected_notes:
        icon = {"auto": "🔄", "ankieta": "📋", "warunkowe": "❓"}.get(n["category"], "📝")
        prio = {"1": "🔴", "2": "🟡", "3": "⚪"}.get(str(n["priority"]), "")
        display.append({
            "level": "OK",
            "msg": f"{prio} {icon} Nota {n['nr']}: {n['name']} — {n['reason']}"
        })
    return display


# ═══════════════════════════════════════════════════════════════════════════════
# MODUŁ 4: GENEROWANIE INFORMACJI DODATKOWEJ PRZEZ CLAUDE
# ═══════════════════════════════════════════════════════════════════════════════

SYSTEM_PROMPT = """Jesteś biegłym rewidentem i ekspertem ds. rachunkowości, specjalizującym się w polskim prawie bilansowym.
Twoim zadaniem jest sporządzenie profesjonalnej "Informacji Dodatkowej" do sprawozdania finansowego.

WYMAGANIA PRAWNE (Ustawa o Rachunkowości, Dz.U. z 2023 r. poz. 120):
- Art. 48 UoR: Informacja dodatkowa obejmuje wprowadzenie i dodatkowe informacje i objaśnienia
- Stosuj Krajowe Standardy Rachunkowości (KSR)
- Używaj zasad wyceny zgodnych z UoR

STRUKTURA DOKUMENTU (obowiązkowa):
1. WPROWADZENIE DO SPRAWOZDANIA FINANSOWEGO
   1.1 Dane identyfikacyjne jednostki
   1.2 Zasady (polityka) rachunkowości — opisz DOKŁADNIE przyjęte zasady z dostarczonego dokumentu polityki
   1.3 Metody wyceny aktywów i pasywów (środki trwałe, zapasy, należności, zobowiązania)
   1.4 Metody amortyzacji i stosowane stawki/okresy ekonomicznej użyteczności
   1.5 Zasady rozliczania przychodów i kosztów
   1.6 Korekty błędów i zmiany polityki rachunkowości

2. DODATKOWE INFORMACJE I OBJAŚNIENIA
   2.1 Szczegółowy zakres zmian wartości grup rodzajowych środków trwałych
       (wartość brutto, odpisy amortyzacyjne/umorzeniowe, wartość netto)
   2.2 Wartości niematerialne i prawne
   2.3 Należności długoterminowe
   2.4 Inwestycje długoterminowe
   2.5 Zapasy (surowce, WIP, wyroby gotowe)
   2.6 Należności krótkoterminowe (z podziałem na tytuły)
   2.7 Środki pieniężne i ich ekwiwalenty
   2.8 Kapitały własne (zmiany w roku obrotowym)
   2.9 Zobowiązania długo- i krótkoterminowe
   2.10 Rozliczenia międzyokresowe
   2.11 Przychody i koszty operacyjne (analiza)
   2.12 Wynik finansowy i jego podział
   2.13 Zobowiązania podatkowe (podatek odroczony)
   2.14 Zatrudnienie (średnie w roku)
   2.15 Wynagrodzenia organów spółki
   2.16 Zdarzenia po dniu bilansowym
   2.17 Inne istotne informacje

ANKIETA BILANSOWA — ZASADY WYKORZYSTANIA:
Jeśli dostarczono wypełnioną Ankietę Bilansową od klienta, OBOWIĄZKOWO uwzględnij odpowiedzi:
- Kontynuacja działalności (pyt. 12) → sekcja 1.1 (oświadczenie o kontynuacji) oraz 2.17
- Postępowania sądowe/egzekucyjne (pyt. 5) → sekcja 2.17, nota o rezerwach
- Propozycja podziału zysku (pyt. 6) → sekcja 2.12 (wynik finansowy i jego podział)
- Propozycja pokrycia straty (pyt. 7) → sekcja 2.12
- Zobowiązania warunkowe i zabezpieczenia (pyt. 8) → osobna nota / sekcja 2.9
- Gwarancje i poręczenia udzielone (pyt. 9) i otrzymane (pyt. 10) → sekcja 2.17
- Należności wątpliwe (pyt. 11) → sekcja 2.6 (nota o odpisach aktualizujących)
- Prognozy rozwoju (pyt. 13) → sekcja 2.17
- Transakcje z powiązanymi (pyt. 14-15) → osobna nota / sekcja 2.17
- Pożyczki dla organów (pyt. 16-17) → sekcja 2.15
- Planowane nakłady inwestycyjne (pyt. 18) → sekcja 2.17
- Odsetki od należności (pyt. 19) → wpływa na wycenę należności w sekcji 1.3
- Zdarzenia po dniu bilansowym (pyt. 20) → sekcja 2.16

Jeśli odpowiedź na pytanie z ankiety brzmi "Tak" — ROZWIŃ temat profesjonalnie.
Jeśli odpowiedź brzmi "Nie" — krótko stwierdź brak wystąpienia danego zjawiska.
W przypadku Q6/Q7 (podział zysku/pokrycie straty) — opisz wybraną propozycję.

STYL I JĘZYK:
- Profesjonalne słownictwo: "wartość netto aktywów", "odpisy amortyzacyjne", "kapitał własny"
- Liczby w PLN z dokładnością do groszy lub w tysiącach PLN (konsekwentnie)
- Tryb oznajmujący, strona bierna, czas przeszły dla zdarzeń roku
- Odesłania do konkretnych not i pozycji bilansu
- Tabele generuj w formacie MARKDOWN (| kolumna1 | kolumna2 |) — zostaną skonwertowane na tabele Word
- ZASADA ZEROWYCH WARTOŚCI: Jeśli dla danej noty objaśniającej WSZYSTKIE wartości liczbowe wynoszą 0 (zero), NIE generuj tabeli. Zamiast tego napisz: "Nota X — [tytuł]: Nie dotyczy." Dotyczy to zarówno tabel, jak i opisów liczbowych.
- NUMERACJA NOT: Noty objaśniające numeruj SEKWENCYJNIE (Nota 1, Nota 2, Nota 3...) w kolejności ich występowania w dokumencie. NIE używaj numerów katalogowych GOFIN (np. Nota 17, Nota 35). Każda nota w wygenerowanym dokumencie dostaje kolejny numer od 1.

WAŻNE: Jeśli dane finansowe są dostępne w dokumentach – cytuj je dokładnie.
Jeśli brakuje danych – zaznacz "[DANE DO UZUPEŁNIENIA]" i opisz co powinno się znaleźć.
Jeśli dostarczono dokument Polityki Rachunkowości – sekcja 1.2–1.5 musi być oparta WYŁĄCZNIE na jego treści.

DANE DO WYKRESÓW (OBOWIĄZKOWE):
Na samym końcu dokumentu, PO całej treści Informacji Dodatkowej, dodaj blok danych w formacie JSON otoczony znacznikami <!--CHART_DATA_START--> i <!--CHART_DATA_END-->.
Blok ten NIE będzie widoczny w dokumencie Word — służy wyłącznie do automatycznego generowania wykresów.
Wypełnij TYLKO te pola, dla których masz KONKRETNE dane liczbowe z dokumentów. Pomiń pola bez danych.

Format:
<!--CHART_DATA_START-->
{
  "aktywa_trwale": 0,
  "aktywa_obrotowe": 0,
  "kapital_wlasny": 0,
  "zobowiazania_dlugoterminowe": 0,
  "zobowiazania_krotkoterminowe": 0,
  "przychody_ze_sprzedazy": 0,
  "koszty_dzialalnosci": 0,
  "wynik_finansowy_netto": 0,
  "amortyzacja": 0,
  "srodki_trwale_brutto": 0,
  "srodki_trwale_umorzenie": 0,
  "srodki_trwale_netto": 0,
  "naleznosci_krotkoterminowe": 0,
  "srodki_pieniezne": 0,
  "zapasy": 0,
  "przychody_rok_poprzedni": 0,
  "koszty_rok_poprzedni": 0,
  "wynik_rok_poprzedni": 0
}
<!--CHART_DATA_END-->"""


def generate_accounting_notes(doc_mapping: dict, anthropic_api_key: str,
                               company_name: str, year: int,
                               company_info: dict = None,
                               progress_callback=None) -> str:
    """
    Krok 4: Wywołuje Claude 3.5 Sonnet do generowania Informacji Dodatkowej.
    """
    client = anthropic.Anthropic(api_key=anthropic_api_key)
    info = company_info or {}

    # Sekcja polityki rachunkowości z odpowiedzi na pytania
    polityka_blok = ""
    pa = info.get("polityka_answers", {})
    if pa:
        # Stałe bloki wyceny (pozycje bez wyboru — zawsze identyczne)
        blok_wnip = (
            "1. WARTOŚCI NIEMATERIALNE I PRAWNE (bez wyboru):\n"
            "   - Wycena początkowa: według cen nabycia.\n"
            "   - Amortyzacja: odpisy amortyzacyjne metodą liniową. Okresy amortyzacji odzwierciedlają "
            "przewidywany czas ekonomicznej użyteczności (licencje na oprogramowanie — 2-5 lat, "
            "koszty zakończonych prac rozwojowych — max. 5 lat).\n"
            "   - Weryfikacja: raz w roku przegląd stawek amortyzacyjnych oraz ocena przesłanek "
            "do odpisów aktualizujących z tytułu trwałej utraty wartości (KSR 4)."
        )

        blok_rat_stale = (
            "   - Koszty finansowania: cena nabycia/koszt wytworzenia obejmuje koszty obsługi zobowiązań "
            "(odsetki, prowizje) oraz różnice kursowe od zobowiązań w walutach obcych, "
            "poniesione do momentu oddania środka do używania.\n"
            "   - Komponenty: w przypadku istotnych środków trwałych o różnych okresach użytkowania "
            "części składowych — podejście komponentowe (osobna amortyzacja).\n"
            "   - Niskocenne składniki: o wartości poniżej 10 000 PLN odpisywane jednorazowo w koszty."
        )

        blok_aktywa_fin_stale = (
            "   Inne aktywa finansowe:\n"
            "   - Przeznaczone do obrotu: wyceniane w wartości godziwej przez wynik finansowy.\n"
            "   - Utrzymywane do terminu wymagalności: wyceniane wg skorygowanej ceny nabycia (efektywna stopa procentowa).\n"
            "   - Pożyczki udzielone i należności własne: w kwocie wymaganej zapłaty z zachowaniem zasady ostrożności."
        )

        blok_zapasy_stale = (
            "   - Koszty wytworzenia: obejmują koszty bezpośrednie oraz uzasadnioną część pośrednich kosztów produkcji. "
            "Koszty niewykorzystanych zdolności produkcyjnych odnoszone bezpośrednio w wynik finansowy.\n"
            "   - Rozchód zapasów: ustalany metodą FIFO.\n"
            "   - Odpisy aktualizujące: tworzone na zapasy wolnorotujące (>12 miesięcy) oraz o obniżonej przydatności."
        )

        blok_naleznosci = (
            "6. NALEŻNOŚCI I ZOBOWIĄZANIA (bez wyboru):\n"
            "   - Wycena: w kwocie wymaganej zapłaty (wraz z odsetkami na dzień bilansowy).\n"
            "   - Odpisy aktualizujące należności: metoda indywidualna dla przeterminowanych >180 dni "
            "oraz metoda ogólna (portfelowa) na podstawie historycznych wskaźników ściągalności.\n"
            "   - Wycena walutowa: aktywa i pasywa w walutach obcych wg średniego kursu NBP "
            "z dnia poprzedzającego dzień bilansowy."
        )

        blok_rezerwy = (
            "7. REZERWY NA ŚWIADCZENIA PRACOWNICZE I INNE ZOBOWIĄZANIA (bez wyboru):\n"
            "   - Rezerwy aktuarialne: na odprawy emerytalne i nagrody jubileuszowe wyceniane "
            "metodą prognozowanych uprawnień jednostkowych.\n"
            "   - Rezerwy na niewykorzystane urlopy: iloczyn dni niewykorzystanego urlopu "
            "i średniej stawki dziennego wynagrodzenia powiększonej o narzuty ZUS.\n"
            "   - Pozostałe rezerwy: na znane ryzyka (postępowania sądowe, naprawy gwarancyjne) "
            "w kwocie wiarygodnie oszacowanej."
        )

        blok_podatek = (
            "8. PODATEK ODROCZONY (bez wyboru):\n"
            "   - Aktywa i rezerwy z tytułu odroczonego podatku dochodowego ustalane w związku "
            "z przejściowymi różnicami między wartością bilansową a podatkową aktywów i pasywów.\n"
            "   - Aktywa z tytułu podatku odroczonego rozpoznawane tylko do wysokości prawdopodobnego "
            "dochodu podatkowego pozwalającego na ich potrącenie."
        )

        blok_rmp = (
            "9. ROZLICZENIA MIĘDZYOKRESOWE PRZYCHODÓW (bez wyboru):\n"
            "   - Obejmują m.in. otrzymane dotacje na sfinansowanie nabycia środków trwałych, "
            "rozliczane równolegle do odpisów amortyzacyjnych tych środków."
        )

        polityka_blok = """
📋 ZASADY RACHUNKOWOŚCI (odpowiedzi udzielone przez użytkownika — brak załączonej Polityki Rachunkowości):

=== A. ZASADY OGÓLNE ===
- Zasady ustalania wyniku finansowego: {wynik}
- Sposób sporządzania sprawozdania: {spr}
- Ujęcie leasingu: {leas}
- Rezerwa/aktywa z tytułu podatku odroczonego: {pod}

=== B. METODY WYCENY AKTYWÓW I PASYWÓW ===

{wnip}

2. RZECZOWE AKTYWA TRWAŁE (wybór użytkownika):
   - Wycena początkowa: {rat}
{rat_stale}

3. INWESTYCJE W NIERUCHOMOŚCI (wybór użytkownika):
   - Wycena: {inwest}

4. AKTYWA I PASYWA FINANSOWE (wybór użytkownika):
   - Udziały w jednostkach podporządkowanych: {udzialy}
{aktywa_fin_stale}

5. ZAPASY (wybór użytkownika):
   - Wycena bilansowa: {zapasy}
{zapasy_stale}

{naleznosci}

{rezerwy}

{podatek}

{rmp}
{uwagi_blok}

Na podstawie powyższych odpowiedzi wypełnij DOKŁADNIE sekcje 1.2–1.5 Informacji Dodatkowej,
opisując WSZYSTKIE metody wyceny aktywów i pasywów (punkty 1-9) w sposób profesjonalny.
""".format(
            wynik=pa.get("wynik_finansowy", ""),
            spr=pa.get("sposob_sprawozdania", ""),
            leas=pa.get("leasing", ""),
            pod="TAK" if pa.get("podatek_odroczony") else "NIE",
            wnip=blok_wnip,
            rat=pa.get("rat_wycena", ""),
            rat_stale=blok_rat_stale,
            inwest=pa.get("inwestycje_nieruchomosci", ""),
            udzialy=pa.get("udzialy_wycena", ""),
            aktywa_fin_stale=blok_aktywa_fin_stale,
            zapasy=pa.get("zapasy_wycena", ""),
            zapasy_stale=blok_zapasy_stale,
            naleznosci=blok_naleznosci,
            rezerwy=blok_rezerwy,
            podatek=blok_podatek,
            rmp=blok_rmp,
            uwagi_blok=f"- Dodatkowe uwagi: {pa['uwagi']}" if pa.get("uwagi") else ""
        )

    # Sekcja zagrożenia kontynuacji
    zagrozenie_blok = ""
    if info.get("zagrozenie_kontynuacji"):
        zagrozenie_blok = (
            "\n⚠️ WAŻNE: Jednostka zidentyfikowała OKOLICZNOŚCI ZAGROŻENIA KONTYNUOWANIA "
            "DZIAŁALNOŚCI (art. 5 ust. 2 UoR). W sekcji dotyczącej zasad rachunkowości "
            "OBOWIĄZKOWO opisz te okoliczności i ich wpływ na wycenę aktywów i pasywów.\n"
            f"Opis okoliczności: {info.get('zagrozenie_opis', '')}\n"
        )

    # Przygotuj kontekst z dokumentów
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
        f"ROK OBROTOWY: {year}",
        polityka_blok,
        zagrozenie_blok,
        "=" * 60,
    ]

    # Ankieta bilansowa — wyodrębnij i podaj z wyróżnieniem
    ankieta_found = False
    for filename, doc_data in doc_mapping.items():
        if doc_data["type"] == "ANKIETA BILANSOWA":
            ankieta_found = True
            context_parts.append("\n📋 ANKIETA BILANSOWA OD KLIENTA (odpowiedzi na pytania):")
            context_parts.append("=" * 40)
            context_parts.append(doc_data["text"][:12000])  # Ankieta jest krótka — dajemy więcej
            context_parts.append("=" * 40)
            context_parts.append(
                "INSTRUKCJA: Powyższa ankieta zawiera odpowiedzi klienta. "
                "Na ich podstawie OBOWIĄZKOWO uwzględnij w Informacji Dodatkowej: "
                "kontynuację działalności, podział wyniku, zobowiązania warunkowe, "
                "gwarancje/poręczenia, należności wątpliwe, transakcje z powiązanymi, "
                "pożyczki dla organów, planowane nakłady, zdarzenia po dniu bilansowym. "
                "Przy odpowiedzi 'Tak' — ROZWIŃ profesjonalnie. "
                "Przy odpowiedzi 'Nie' — krótko stwierdź brak wystąpienia."
            )
            break

    if not ankieta_found:
        context_parts.append(
            "\n⚠️ BRAK ANKIETY BILANSOWEJ: Nie dostarczono ankiety bilansowej od klienta. "
            "W sekcjach 2.12 (podział wyniku), 2.16 (zdarzenia po dniu bilansowym), "
            "2.17 (zobowiązania warunkowe, gwarancje, transakcje z powiązanymi) "
            "wpisz [DANE DO UZUPEŁNIENIA — wymagana ankieta bilansowa od klienta]."
        )

    context_parts.append("\n" + "=" * 60)

    # Dodaj listę wybranych not objaśniających
    selected_notes = info.get("selected_notes", [])
    if selected_notes:
        context_parts.append(format_notes_for_prompt(selected_notes))

    context_parts.append("=" * 60)
    context_parts.append("WYCIĄGI Z DOKUMENTÓW FINANSOWYCH:")

    for filename, doc_data in doc_mapping.items():
        if doc_data["type"] == "ANKIETA BILANSOWA":
            continue  # Już dodana wyżej z wyróżnieniem
        context_parts.append(f"\n[{doc_data['type']}] {filename}:")
        # Ogranicz do 8000 znaków na dokument
        context_parts.append(doc_data["text"][:8000])
        if len(doc_data["text"]) > 8000:
            context_parts.append("...[tekst skrócony]")

    full_context = "\n".join(context_parts)

    user_prompt = f"""Na podstawie poniższych dokumentów finansowych sporządź kompletną "Informację Dodatkową" do sprawozdania finansowego za rok {year}.

{full_context}

Wygeneruj pełną Informację Dodatkową zgodnie z polską Ustawą o Rachunkowości.
Gdzie masz dane – użyj konkretnych liczb. Gdzie brakuje – napisz [DANE DO UZUPEŁNIENIA].
Formatuj wyraźnie nagłówkami i akapitami."""

    if progress_callback:
        progress_callback(0.7, "Generowanie przez Claude 3.5 Sonnet...")

    response = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=8000,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": user_prompt}]
    )

    return response.content[0].text


# ═══════════════════════════════════════════════════════════════════════════════
# MODUŁ 5: EKSPORT DO WORD (.docx) — PROFESJONALNA SZATA GRAFICZNA
# ═══════════════════════════════════════════════════════════════════════════════

# Kolory firmowe
_CLR_NAVY = RGBColor(0x1B, 0x2A, 0x4A)      # Nagłówki główne
_CLR_BLUE = RGBColor(0x2D, 0x6A, 0x9F)       # Nagłówki sekcji
_CLR_ACCENT = RGBColor(0x3A, 0x86, 0xC8)     # Akcent, linie
_CLR_GRAY = RGBColor(0x66, 0x66, 0x66)       # Tekst pomocniczy
_CLR_LIGHT = RGBColor(0x99, 0x99, 0x99)      # Stopka
_CLR_BLACK = RGBColor(0x33, 0x33, 0x33)      # Tekst główny
_CLR_TABLE_HEADER = "1B2A4A"                  # Tło nagłówka tabeli (hex)
_CLR_TABLE_ALT = "F2F6FA"                     # Naprzemienne wiersze tabeli


def _setup_styles(doc: Document):
    """Konfiguruje style dokumentu."""
    # Normal
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(10)
    style.font.color.rgb = _CLR_BLACK
    pf = style.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(4)
    pf.line_spacing = 1.15

    # Heading 1 — sekcje główne (np. "1. WPROWADZENIE...")
    h1 = doc.styles["Heading 1"]
    h1.font.name = "Calibri"
    h1.font.size = Pt(14)
    h1.font.bold = True
    h1.font.color.rgb = _CLR_NAVY
    h1.paragraph_format.space_before = Pt(18)
    h1.paragraph_format.space_after = Pt(8)
    h1.paragraph_format.keep_with_next = True

    # Heading 2 — podsekcje (np. "1.1 Dane identyfikacyjne")
    h2 = doc.styles["Heading 2"]
    h2.font.name = "Calibri"
    h2.font.size = Pt(12)
    h2.font.bold = True
    h2.font.color.rgb = _CLR_BLUE
    h2.paragraph_format.space_before = Pt(12)
    h2.paragraph_format.space_after = Pt(6)
    h2.paragraph_format.keep_with_next = True

    # Heading 3 — pod-podsekcje
    h3 = doc.styles["Heading 3"]
    h3.font.name = "Calibri"
    h3.font.size = Pt(11)
    h3.font.bold = True
    h3.font.color.rgb = _CLR_NAVY
    h3.paragraph_format.space_before = Pt(8)
    h3.paragraph_format.space_after = Pt(4)

    # Heading 4 — noty
    h4 = doc.styles["Heading 4"]
    h4.font.name = "Calibri"
    h4.font.size = Pt(10)
    h4.font.bold = True
    h4.font.italic = True
    h4.font.color.rgb = _CLR_BLUE
    h4.paragraph_format.space_before = Pt(6)
    h4.paragraph_format.space_after = Pt(3)


def _add_separator(doc: Document, color: str = "2D6A9F", thickness: int = 6):
    """Dodaje profesjonalną linię separatora."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), str(thickness))
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), color)
    pBdr.append(bottom)
    pPr.append(pBdr)


def _add_page_number(doc: Document):
    """Dodaje numerację stron w stopce."""
    for section in doc.sections:
        footer = section.footer
        footer.is_linked_to_previous = False
        p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)

        # Linia nad stopką
        pPr = p._p.get_or_add_pPr()
        pBdr = OxmlElement("w:pBdr")
        top = OxmlElement("w:top")
        top.set(qn("w:val"), "single")
        top.set(qn("w:sz"), "4")
        top.set(qn("w:space"), "4")
        top.set(qn("w:color"), "CCCCCC")
        pBdr.append(top)
        pPr.append(pBdr)

        run = p.add_run("Strona ")
        run.font.size = Pt(8)
        run.font.color.rgb = _CLR_LIGHT
        run.font.name = "Calibri"

        # Pole numeru strony
        fldChar1 = OxmlElement("w:fldChar")
        fldChar1.set(qn("w:fldCharType"), "begin")
        run2 = p.add_run()
        run2._r.append(fldChar1)

        instrText = OxmlElement("w:instrText")
        instrText.set(qn("xml:space"), "preserve")
        instrText.text = " PAGE "
        run3 = p.add_run()
        run3.font.size = Pt(8)
        run3.font.color.rgb = _CLR_LIGHT
        run3._r.append(instrText)

        fldChar2 = OxmlElement("w:fldChar")
        fldChar2.set(qn("w:fldCharType"), "end")
        run4 = p.add_run()
        run4._r.append(fldChar2)


def _add_header(doc: Document, company_name: str, year: int):
    """Dodaje nagłówek dokumentu z nazwą firmy."""
    for section in doc.sections:
        header = section.header
        header.is_linked_to_previous = False
        p = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)

        run = p.add_run(f"{company_name} | Informacja Dodatkowa {year}")
        run.font.size = Pt(7)
        run.font.color.rgb = _CLR_LIGHT
        run.font.name = "Calibri"
        run.font.italic = True

        # Linia pod nagłówkiem
        pPr = p._p.get_or_add_pPr()
        pBdr = OxmlElement("w:pBdr")
        bottom = OxmlElement("w:bottom")
        bottom.set(qn("w:val"), "single")
        bottom.set(qn("w:sz"), "4")
        bottom.set(qn("w:space"), "4")
        bottom.set(qn("w:color"), "CCCCCC")
        pBdr.append(bottom)
        pPr.append(pBdr)


def _set_cell_shading(cell, color_hex: str):
    """Ustawia kolor tła komórki."""
    shading = OxmlElement("w:shd")
    shading.set(qn("w:fill"), color_hex)
    shading.set(qn("w:val"), "clear")
    cell._tc.get_or_add_tcPr().append(shading)


def add_markdown_table_to_doc(doc: Document, table_lines: list):
    """Konwertuje linie markdown table na profesjonalnie sformatowaną tabelę Word."""
    data_lines = [l for l in table_lines if not re.match(r"^\|[\s\-:|]+$", l)]
    if not data_lines:
        return

    rows = []
    for line in data_lines:
        cells = [c.strip() for c in line.strip("|").split("|")]
        rows.append(cells)
    if not rows:
        return

    num_cols = max(len(r) for r in rows)
    for r in rows:
        while len(r) < num_cols:
            r.append("")

    table = doc.add_table(rows=len(rows), cols=num_cols)
    try:
        table.style = "Table Grid"
    except KeyError:
        pass
    table.autofit = True

    # Ustaw obramowanie na delikatne szare linie
    tbl = table._tbl
    tblPr = tbl.tblPr if tbl.tblPr is not None else OxmlElement("w:tblPr")
    borders = OxmlElement("w:tblBorders")
    for edge in ["top", "left", "bottom", "right", "insideH", "insideV"]:
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "4")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "B0B0B0")
        borders.append(el)
    tblPr.append(borders)

    for i, row_data in enumerate(rows):
        row = table.rows[i]
        for j, cell_text in enumerate(row_data):
            cell = row.cells[j]
            cell.text = ""
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)

            # Nagłówek (wiersz 0) — białe litery na ciemnym tle
            if i == 0:
                _set_cell_shading(cell, _CLR_TABLE_HEADER)
                parts = re.split(r"(\*\*[^*]+\*\*)", cell_text)
                for part in parts:
                    clean = part.strip("*") if part.startswith("**") else part
                    run = p.add_run(clean)
                    run.bold = True
                    run.font.size = Pt(8)
                    run.font.name = "Calibri"
                    run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            else:
                # Naprzemienne wiersze
                if i % 2 == 0:
                    _set_cell_shading(cell, _CLR_TABLE_ALT)

                parts = re.split(r"(\*\*[^*]+\*\*)", cell_text)
                for part in parts:
                    if part.startswith("**") and part.endswith("**"):
                        run = p.add_run(part[2:-2])
                        run.bold = True
                    else:
                        run = p.add_run(part)
                    run.font.size = Pt(8)
                    run.font.name = "Calibri"
                    run.font.color.rgb = _CLR_BLACK

    doc.add_paragraph()  # Odstęp po tabeli


def _extract_chart_data(generated_text: str) -> dict:
    """Wyciąga blok JSON z danymi do wykresów z odpowiedzi Claude."""
    match = re.search(
        r"<!--CHART_DATA_START-->\s*(\{.*?\})\s*<!--CHART_DATA_END-->",
        generated_text, re.DOTALL
    )
    if not match:
        return {}
    try:
        return json.loads(match.group(1))
    except (json.JSONDecodeError, ValueError):
        return {}


def _strip_chart_data(generated_text: str) -> str:
    """Usuwa blok chart data z tekstu przed wyświetleniem/eksportem."""
    return re.sub(
        r"\s*<!--CHART_DATA_START-->.*?<!--CHART_DATA_END-->\s*",
        "", generated_text, flags=re.DOTALL
    ).strip()


def _setup_chart_style():
    """Konfiguruje styl wykresów — profesjonalny, spójny z dokumentem."""
    plt.rcParams.update({
        "font.family": "sans-serif",
        "font.sans-serif": ["Calibri", "DejaVu Sans", "Arial"],
        "font.size": 9,
        "axes.titlesize": 11,
        "axes.titleweight": "bold",
        "axes.labelsize": 9,
        "axes.spines.top": False,
        "axes.spines.right": False,
        "figure.facecolor": "white",
        "figure.dpi": 150,
        "savefig.dpi": 150,
        "savefig.bbox": "tight",
        "savefig.pad_inches": 0.2,
    })


# Paleta kolorów spójna z dokumentem
_CHART_COLORS = ["#1B2A4A", "#2D6A9F", "#3A86C8", "#6BAED6", "#9ECAE1",
                  "#C6DBEF", "#4A90D9", "#7FB3E0", "#A8D0E8", "#D1E8F5"]
_CHART_ACCENT = "#E74C3C"  # Czerwony akcent (np. strata)


def _fmt_pln(value: float) -> str:
    """Formatuje kwotę PLN do czytelnego formatu."""
    if abs(value) >= 1_000_000:
        return f"{value/1_000_000:,.1f} mln"
    elif abs(value) >= 1_000:
        return f"{value/1_000:,.0f} tys."
    return f"{value:,.0f}"


def _generate_pie_chart(data: dict, keys: list, labels: list,
                         title: str) -> bytes | None:
    """Generuje wykres kołowy. Zwraca PNG jako bytes."""
    values = [data.get(k, 0) for k in keys]
    # Filtruj zera
    filtered = [(l, v) for l, v in zip(labels, values) if v > 0]
    if len(filtered) < 2:
        return None

    labels_f, values_f = zip(*filtered)

    _setup_chart_style()
    fig, ax = plt.subplots(figsize=(5, 3.5))

    colors = _CHART_COLORS[:len(values_f)]
    wedges, texts, autotexts = ax.pie(
        values_f, labels=None, autopct="%1.1f%%",
        colors=colors, startangle=90,
        pctdistance=0.75, wedgeprops={"linewidth": 1, "edgecolor": "white"}
    )
    for at in autotexts:
        at.set_fontsize(8)
        at.set_color("white")
        at.set_fontweight("bold")

    # Legenda z kwotami
    legend_labels = [f"{l} ({_fmt_pln(v)} PLN)" for l, v in zip(labels_f, values_f)]
    ax.legend(wedges, legend_labels, loc="center left", bbox_to_anchor=(1, 0.5),
              fontsize=8, frameon=False)

    ax.set_title(title, pad=12, color="#1B2A4A")

    buf = io.BytesIO()
    fig.savefig(buf, format="png")
    plt.close(fig)
    buf.seek(0)
    return buf.getvalue()


def _generate_bar_comparison(data: dict, title: str) -> bytes | None:
    """Generuje wykres słupkowy porównujący rok bieżący vs poprzedni."""
    categories = ["Przychody", "Koszty", "Wynik netto"]
    current = [
        data.get("przychody_ze_sprzedazy", 0),
        data.get("koszty_dzialalnosci", 0),
        data.get("wynik_finansowy_netto", 0),
    ]
    previous = [
        data.get("przychody_rok_poprzedni", 0),
        data.get("koszty_rok_poprzedni", 0),
        data.get("wynik_rok_poprzedni", 0),
    ]

    # Sprawdź czy mamy dane za oba lata
    if sum(current) == 0:
        return None

    _setup_chart_style()
    fig, ax = plt.subplots(figsize=(6, 3.5))

    x = range(len(categories))
    width = 0.35

    has_previous = sum(previous) != 0

    if has_previous:
        bars1 = ax.bar([i - width/2 for i in x], previous, width,
                        label="Rok poprzedni", color="#9ECAE1", edgecolor="white")
        bars2 = ax.bar([i + width/2 for i in x], current, width,
                        label="Rok bieżący", color="#1B2A4A", edgecolor="white")
    else:
        colors = ["#2D6A9F", "#6BAED6",
                  "#27AE60" if current[2] >= 0 else _CHART_ACCENT]
        bars2 = ax.bar(x, current, width * 1.5, color=colors, edgecolor="white")

    ax.set_xticks(x)
    ax.set_xticklabels(categories)
    ax.set_title(title, pad=12, color="#1B2A4A")

    # Formatowanie osi Y
    ax.yaxis.set_major_formatter(ticker.FuncFormatter(
        lambda v, p: _fmt_pln(v)
    ))

    # Etykiety wartości nad słupkami
    for bar in (bars2 if not has_previous else list(bars1) + list(bars2)):
        h = bar.get_height()
        if h != 0:
            ax.text(bar.get_x() + bar.get_width()/2, h,
                    _fmt_pln(h), ha="center", va="bottom", fontsize=7,
                    color="#333333")

    if has_previous:
        ax.legend(fontsize=8, frameon=False)

    ax.axhline(y=0, color="#CCCCCC", linewidth=0.5)
    fig.tight_layout()

    buf = io.BytesIO()
    fig.savefig(buf, format="png")
    plt.close(fig)
    buf.seek(0)
    return buf.getvalue()


def _generate_asset_structure_bar(data: dict) -> bytes | None:
    """Generuje wykres struktury aktywów obrotowych."""
    keys = ["naleznosci_krotkoterminowe", "zapasy", "srodki_pieniezne"]
    labels = ["Należności", "Zapasy", "Środki pieniężne"]
    values = [data.get(k, 0) for k in keys]

    filtered = [(l, v) for l, v in zip(labels, values) if v > 0]
    if len(filtered) < 2:
        return None

    labels_f, values_f = zip(*filtered)

    _setup_chart_style()
    fig, ax = plt.subplots(figsize=(5, 3))

    colors = ["#2D6A9F", "#3A86C8", "#6BAED6"][:len(values_f)]
    bars = ax.barh(labels_f, values_f, color=colors, edgecolor="white", height=0.5)

    for bar, v in zip(bars, values_f):
        ax.text(bar.get_width() + max(values_f) * 0.02, bar.get_y() + bar.get_height()/2,
                f"{_fmt_pln(v)} PLN", va="center", fontsize=8, color="#333333")

    ax.set_title("Struktura aktywów obrotowych", pad=12, color="#1B2A4A")
    ax.xaxis.set_major_formatter(ticker.FuncFormatter(lambda v, p: _fmt_pln(v)))
    ax.invert_yaxis()
    fig.tight_layout()

    buf = io.BytesIO()
    fig.savefig(buf, format="png")
    plt.close(fig)
    buf.seek(0)
    return buf.getvalue()


def generate_charts(chart_data: dict, year: int) -> list:
    """
    Generuje wszystkie wykresy na podstawie danych.
    Zwraca listę: [(tytuł, bytes_png), ...]
    """
    if not chart_data:
        return []

    charts = []

    # 1. Struktura aktywów (pie)
    png = _generate_pie_chart(
        chart_data,
        keys=["aktywa_trwale", "aktywa_obrotowe"],
        labels=["Aktywa trwałe", "Aktywa obrotowe"],
        title=f"Struktura aktywów — {year}"
    )
    if png:
        charts.append(("Struktura aktywów", png))

    # 2. Struktura pasywów (pie)
    png = _generate_pie_chart(
        chart_data,
        keys=["kapital_wlasny", "zobowiazania_dlugoterminowe", "zobowiazania_krotkoterminowe"],
        labels=["Kapitał własny", "Zobowiązania długoterminowe", "Zobowiązania krótkoterminowe"],
        title=f"Struktura pasywów — {year}"
    )
    if png:
        charts.append(("Struktura pasywów", png))

    # 3. Przychody vs koszty vs wynik (bar)
    png = _generate_bar_comparison(
        chart_data,
        title=f"Przychody, koszty i wynik finansowy — {year}"
    )
    if png:
        charts.append(("Analiza wyniku finansowego", png))

    # 4. Struktura aktywów obrotowych (horizontal bar)
    png = _generate_asset_structure_bar(chart_data)
    if png:
        charts.append(("Struktura aktywów obrotowych", png))

    # 5. Środki trwałe — brutto vs umorzenie vs netto (bar)
    st_data = {
        "brutto": chart_data.get("srodki_trwale_brutto", 0),
        "umorzenie": chart_data.get("srodki_trwale_umorzenie", 0),
        "netto": chart_data.get("srodki_trwale_netto", 0),
    }
    if st_data["brutto"] > 0:
        _setup_chart_style()
        fig, ax = plt.subplots(figsize=(5, 3))
        cats = ["Wartość brutto", "Umorzenie", "Wartość netto"]
        vals = [st_data["brutto"], st_data["umorzenie"], st_data["netto"]]
        colors = ["#1B2A4A", "#E74C3C", "#27AE60"]
        bars = ax.bar(cats, vals, color=colors, edgecolor="white", width=0.5)
        for bar, v in zip(bars, vals):
            ax.text(bar.get_x() + bar.get_width()/2, bar.get_height(),
                    _fmt_pln(v), ha="center", va="bottom", fontsize=8)
        ax.set_title(f"Środki trwałe — {year}", pad=12, color="#1B2A4A")
        ax.yaxis.set_major_formatter(ticker.FuncFormatter(lambda v, p: _fmt_pln(v)))
        fig.tight_layout()
        buf = io.BytesIO()
        fig.savefig(buf, format="png")
        plt.close(fig)
        buf.seek(0)
        charts.append(("Środki trwałe", buf.getvalue()))

    return charts


def _add_charts_section(doc, charts: list):
    """Dodaje sekcję z wykresami do dokumentu Word."""
    if not charts:
        return

    doc.add_page_break()
    h = doc.add_heading("Analiza graficzna", level=1)

    p_intro = doc.add_paragraph()
    run = p_intro.add_run(
        "Poniższe wykresy przedstawiają graficzną analizę kluczowych danych "
        "finansowych na podstawie sprawozdania finansowego."
    )
    run.font.size = Pt(9)
    run.font.color.rgb = _CLR_GRAY
    run.font.italic = True

    for i, (title, png_bytes) in enumerate(charts):
        # Tytuł wykresu
        p_title = doc.add_paragraph()
        p_title.paragraph_format.space_before = Pt(12)
        run = p_title.add_run(f"Wykres {i+1}. {title}")
        run.bold = True
        run.font.size = Pt(10)
        run.font.color.rgb = _CLR_BLUE

        # Wstaw obraz
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_img = p_img.add_run()
        run_img.add_picture(io.BytesIO(png_bytes), width=Inches(5.0))

        # Separator między wykresami
        if i < len(charts) - 1:
            _add_separator(doc, "EEEEEE", 2)


def _add_title_page(doc: Document, company_name: str, year: int, company_info: dict = None):
    """Tworzy profesjonalną stronę tytułową."""
    info = company_info or {}

    # Kilka pustych akapitów na górze dla wycentrowania
    for _ in range(4):
        doc.add_paragraph().paragraph_format.space_after = Pt(0)

    # Linia dekoracyjna
    _add_separator(doc, "2D6A9F", 12)

    # Tytuł
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(16)
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run("INFORMACJA DODATKOWA")
    run.font.name = "Calibri"
    run.font.size = Pt(22)
    run.font.bold = True
    run.font.color.rgb = _CLR_NAVY

    # Podtytuł
    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.paragraph_format.space_after = Pt(4)
    run2 = p2.add_run("do sprawozdania finansowego")
    run2.font.name = "Calibri"
    run2.font.size = Pt(14)
    run2.font.color.rgb = _CLR_BLUE

    p3 = doc.add_paragraph()
    p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p3.paragraph_format.space_after = Pt(12)
    run3 = p3.add_run(f"za rok obrotowy {year}")
    run3.font.name = "Calibri"
    run3.font.size = Pt(14)
    run3.font.color.rgb = _CLR_BLUE

    _add_separator(doc, "2D6A9F", 12)

    # Blok z danymi spółki
    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    fields = [
        ("Jednostka", info.get("nazwa") or company_name),
        ("Forma prawna", info.get("forma_prawna", "")),
        ("Siedziba", info.get("siedziba", "")),
        ("NIP", info.get("nip", "")),
        ("KRS", info.get("krs", "")),
        ("REGON", info.get("regon", "")),
        ("PKD", info.get("pkd", "")),
        ("Okres sprawozdawczy", f"{info.get('okres_od', '')} — {info.get('okres_do', '')}"),
    ]

    for label, value in fields:
        if not value or value == " — ":
            continue
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        run_label = p.add_run(f"{label}: ")
        run_label.font.name = "Calibri"
        run_label.font.size = Pt(10)
        run_label.font.color.rgb = _CLR_GRAY
        run_val = p.add_run(value)
        run_val.font.name = "Calibri"
        run_val.font.size = Pt(10)
        run_val.font.bold = True
        run_val.font.color.rgb = _CLR_NAVY

    # Data generowania
    doc.add_paragraph().paragraph_format.space_after = Pt(24)
    p_date = doc.add_paragraph()
    p_date.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_date = p_date.add_run(f"Wygenerowano: {date.today().strftime('%d.%m.%Y')}")
    run_date.font.name = "Calibri"
    run_date.font.size = Pt(9)
    run_date.font.color.rgb = _CLR_LIGHT
    run_date.font.italic = True

    doc.add_page_break()


def _add_rich_paragraph(doc: Document, line: str):
    """Dodaje akapit z obsługą **bold** i zachowaniem formatowania."""
    p = doc.add_paragraph()
    parts = re.split(r"(\*\*[^*]+\*\*)", line)
    for part in parts:
        if part.startswith("**") and part.endswith("**"):
            run = p.add_run(part[2:-2])
            run.bold = True
        else:
            p.add_run(part)
    return p


def save_to_word(generated_text: str, company_name: str, year: int,
                 company_info: dict = None) -> bytes:
    """Konwertuje wygenerowaną treść AI na profesjonalny plik .docx z wykresami."""
    doc = Document()

    # Wyciągnij dane do wykresów i oczyść tekst
    chart_data = _extract_chart_data(generated_text)
    clean_text = _strip_chart_data(generated_text)

    # Generuj wykresy
    charts = generate_charts(chart_data, year)
    doc = Document()

    # Konfiguracja stylów
    _setup_styles(doc)

    # Marginesy i rozmiar strony
    for section in doc.sections:
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.7)
        section.header_distance = Inches(0.3)
        section.footer_distance = Inches(0.3)

    # Nagłówek i stopka (numeracja stron)
    _add_header(doc, company_name, year)
    _add_page_number(doc)

    # Strona tytułowa
    _add_title_page(doc, company_name, year, company_info)

    # Parsowanie i formatowanie treści
    lines = clean_text.split("\n")
    i = 0
    while i < len(lines):
        line = lines[i].strip()

        # Wykryj tabelę markdown
        if line.startswith("|") and "|" in line[1:]:
            table_lines = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                table_lines.append(lines[i].strip())
                i += 1
            add_markdown_table_to_doc(doc, table_lines)
            continue

        i += 1

        if not line:
            continue  # Pomijamy puste linie (spacing robi swoje)

        # Nagłówki markdown
        if line.startswith("#### "):
            doc.add_heading(line[5:], level=4)
        elif line.startswith("### "):
            doc.add_heading(line[4:], level=3)
        elif line.startswith("## "):
            h = doc.add_heading(line[3:], level=2)
        elif line.startswith("# "):
            h = doc.add_heading(line[2:], level=1)
        elif line.startswith("---"):
            _add_separator(doc, "CCCCCC", 2)
        elif line.startswith("**") and line.endswith("**"):
            p = doc.add_paragraph()
            run = p.add_run(line.strip("*"))
            run.bold = True
            run.font.color.rgb = _CLR_NAVY
        elif line.startswith("- ") or line.startswith("* "):
            doc.add_paragraph(line[2:], style="List Bullet")
        elif re.match(r"^\d+\.\s", line):
            doc.add_paragraph(line, style="List Number")
        else:
            _add_rich_paragraph(doc, line)

    # Sekcja wykresów (jeśli dane dostępne)
    _add_charts_section(doc, charts)

    # Stopka dokumentu
    _add_separator(doc, "2D6A9F", 4)
    footer_p = doc.add_paragraph()
    footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = footer_p.add_run(
        f"Informacja Dodatkowa | {company_name} | Rok obrotowy {year} | "
        f"Wygenerowano {date.today().strftime('%d.%m.%Y')}"
    )
    run.font.size = Pt(8)
    run.font.color.rgb = _CLR_LIGHT
    run.font.name = "Calibri"
    run.font.italic = True

    # Zapis
    buffer = io.BytesIO()
    doc.save(buffer)
    return buffer.getvalue()


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
    st.markdown("""
    **📋 Obsługiwane dokumenty:**
    - 🏦 Bilans
    - 📈 Rachunek Zysków i Strat
    - 🏗️ Tabela środków trwałych
    - 💸 Przepływy pieniężne
    - 📜 Polityka rachunkowości
    - ⚖️ Zestawienie Obrotów i Sald (ZOiS)
    - 📋 Ankieta bilansowa (wypełniona przez klienta)
    """)


# ── Główna sekcja ────────────────────────────────────────────────────────────
col1, col2 = st.columns([1, 1])

with col1:
    st.markdown('<div class="step-card"><b>📁 Krok 1:</b> Wgraj dokumenty sprawozdania (PDF / DOCX)</div>',
                unsafe_allow_html=True)
    uploaded_files = st.file_uploader(
        "Wybierz pliki PDF lub DOCX",
        type=["pdf", "docx"],
        accept_multiple_files=True,
        help="Wgraj dokumenty: bilans, RZiS, noty, ZOiS, tabela ŚT, ankieta bilansowa (PDF lub DOCX)"
    )

    if uploaded_files:
        st.success(f"✅ Wgrano {len(uploaded_files)} plik(ów)")
        for f in uploaded_files:
            size_kb = len(f.getvalue()) // 1024
            ext = "DOCX" if f.name.lower().endswith(".docx") else "PDF"
            st.caption(f"📄 {f.name} ({size_kb} KB, {ext})")

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
    # Wyczyść poprzednie stany przy nowym uruchomieniu
    st.session_state["run_generation"] = True
    st.session_state.pop("missing_decision", None)
    st.session_state.pop("polityka_answers", None)
    st.session_state.pop("parsed_docs", None)
    st.session_state.pop("doc_mapping", None)
    st.rerun()

# ── Pipeline generowania (działa na podstawie session_state) ─────────────────
if st.session_state.get("run_generation") and anthropic_key and uploaded_files and company_name:

    progress_bar = st.progress(0)
    status_text = st.empty()
    results_container = st.container()

    try:
        # ── KROK 1: Parsowanie (tylko raz, wynik w session_state) ────────
        if "parsed_docs" not in st.session_state:
            status_text.info("📄 Krok 1/5: Parsowanie dokumentów PDF...")
            progress_bar.progress(10)

            def update_progress(val, msg):
                progress_bar.progress(int(10 + val * 20))
                status_text.info(f"📄 {msg}")

            if llama_key:
                parsed = parse_documents_llamaparse(uploaded_files, llama_key, update_progress)
            else:
                parsed = parse_documents_fallback(uploaded_files, update_progress)

            st.session_state["parsed_docs"] = parsed

        parsed = st.session_state["parsed_docs"]
        progress_bar.progress(30)

        # ── KROK 2: Mapowanie (tylko raz) ────────────────────────────────
        if "doc_mapping" not in st.session_state:
            status_text.info("🗂️ Krok 2/5: Mapowanie i identyfikacja dokumentów...")
            doc_mapping = map_documents(parsed)
            st.session_state["doc_mapping"] = doc_mapping

        doc_mapping = st.session_state["doc_mapping"]
        progress_bar.progress(40)

        # ── SPRAWDZENIE BRAKUJĄCYCH DOKUMENTÓW ─────────────────────────────
        missing = check_missing_documents(doc_mapping)
        missing_decision = st.session_state.get("missing_decision")

        if missing and missing_decision is None:
            # Jeszcze nie podjęto decyzji — pokaż pytanie
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

        if missing_decision == "cancel":
            st.session_state.pop("run_generation", None)
            st.session_state.pop("missing_decision", None)
            st.info("Wgraj brakujące pliki i uruchom ponownie.")
            st.stop()

        # (missing_decision == "continue" lub brak braków — kontynuuj normalnie)

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

                    # ── A. ZASADY OGÓLNE ──────────────────────────────────────
                    st.markdown("---")
                    st.markdown("#### A. Zasady ogólne")

                    q1 = st.selectbox(
                        "1. Zasady ustalania wyniku finansowego:",
                        options=[
                            "Wariant porównawczy (układ rodzajowy kosztów)",
                            "Wariant kalkulacyjny (układ funkcjonalny kosztów)",
                        ],
                        help="Dotyczy formy Rachunku Zysków i Strat (art. 47 UoR)"
                    )

                    q_sprawozdanie = st.selectbox(
                        "2. Sposób sporządzania sprawozdania finansowego:",
                        options=[
                            "Wariant pełny (duże jednostki) — pełny bilans, RZiS (porównawczy lub kalkulacyjny), informacja dodatkowa, zestawienie zmian w kapitale, rachunek przepływów pieniężnych",
                            "Wariant dla jednostek małych — uproszczony bilans i RZiS, brak obowiązku zestawienia zmian w kapitale i przepływów pieniężnych (o ile nie podlegają badaniu)",
                            "Wariant dla jednostek mikro — minimalistyczny zakres danych, informacja dodatkowa ograniczona do absolutnego minimum",
                        ],
                        help="Art. 45–50 UoR — zakres sprawozdania zależy od wielkości jednostki"
                    )

                    q_leasing = st.selectbox(
                        "3. Ujęcie leasingu:",
                        options=[
                            "Według UoR (leasing operacyjny/finansowy wg ekonomicznej treści)",
                            "Leasing operacyjny — wszystkie umowy traktowane jako operacyjny",
                            "Nie dotyczy (brak umów leasingowych)",
                        ]
                    )

                    # ── B. METODY WYCENY AKTYWÓW (z wyborem) ──────────────────
                    st.markdown("---")
                    st.markdown("#### B. Metody wyceny aktywów i pasywów")
                    st.info(
                        "Poniższe pytania dotyczą metod wyceny wymaganych przez UoR. "
                        "Pozycje oznaczone *(z wyborem)* wymagają wskazania stosowanej metody. "
                        "Pozycje bez wyboru zostaną uzupełnione automatycznie wg standardowych zasad UoR."
                    )

                    # 2. Rzeczowe aktywa trwałe (z wyborem)
                    st.markdown("**2. Rzeczowe aktywa trwałe** *(z wyborem)*")
                    q_rat_wycena = st.selectbox(
                        "Wycena początkowa rzeczowych aktywów trwałych:",
                        options=[
                            "Według cen nabycia, pomniejszonych o skumulowane odpisy amortyzacyjne oraz odpisy z tytułu trwałej utraty wartości",
                            "Według kosztów wytworzenia, pomniejszonych o skumulowane odpisy amortyzacyjne oraz odpisy z tytułu trwałej utraty wartości",
                        ],
                        help="Art. 28 ust. 1 pkt 1 UoR"
                    )

                    # 3. Inwestycje w nieruchomości (z wyborem)
                    st.markdown("**3. Inwestycje w nieruchomości** *(z wyborem)*")
                    q_inwest_nieruch = st.selectbox(
                        "Wycena inwestycji w nieruchomości:",
                        options=[
                            "Według cen nabycia (zasady jak dla środków trwałych)",
                            "Według wartości godziwej (skutki przeszacowania odnoszone do pozostałych przychodów/kosztów operacyjnych)",
                        ],
                        help="Art. 28 ust. 1 pkt 1a UoR"
                    )

                    # 4. Aktywa i pasywa finansowe (z wyborem)
                    st.markdown("**4. Udziały w jednostkach podporządkowanych** *(z wyborem)*")
                    q_udzialy = st.selectbox(
                        "Wycena udziałów w jednostkach podporządkowanych:",
                        options=[
                            "Metodą ceny nabycia pomniejszonej o odpisy z tytułu trwałej utraty wartości",
                            "Metodą praw własności",
                        ],
                        help="Art. 28 ust. 1 pkt 4 UoR"
                    )

                    # 5. Zapasy (z wyborem)
                    st.markdown("**5. Zapasy** *(z wyborem)*")
                    q_zapasy_wycena = st.selectbox(
                        "Wycena bilansowa zapasów:",
                        options=[
                            "Według cen nabycia, nie wyższych od cen sprzedaży netto",
                            "Według kosztów wytworzenia, nie wyższych od cen sprzedaży netto",
                        ],
                        help="Art. 28 ust. 1 pkt 6 UoR"
                    )

                    # ── C. POZYCJE BEZ WYBORU (informacja) ────────────────────
                    st.markdown("---")
                    st.markdown("#### C. Pozycje wyceniane wg stałych zasad UoR (bez wyboru)")
                    st.caption(
                        "Poniższe pozycje zostaną automatycznie opisane w Informacji Dodatkowej "
                        "zgodnie ze standardowymi zasadami wynikającymi z Ustawy o Rachunkowości:"
                    )
                    st.markdown("""
- **1. Wartości niematerialne i prawne** — wycena wg cen nabycia, amortyzacja liniowa, przegląd stawek raz w roku (KSR 4)
- **6. Należności i zobowiązania** — w kwocie wymaganej zapłaty, odpisy aktualizujące indywidualnie (>180 dni) i portfelowo, wycena walutowa wg kursu NBP
- **7. Rezerwy na świadczenia pracownicze** — rezerwy aktuarialne (odprawy, jubileusze), rezerwa na niewykorzystane urlopy, pozostałe rezerwy
- **8. Podatek odroczony** — aktywa i rezerwy z tytułu różnic przejściowych
- **9. Rozliczenia międzyokresowe przychodów** — w tym dotacje rozliczane równolegle do amortyzacji
                    """)

                    q_podatek = st.checkbox(
                        "Jednostka tworzy rezerwę i aktywa z tytułu odroczonego podatku dochodowego",
                        value=True
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
                        "sposob_sprawozdania": q_sprawozdanie,
                        "leasing": q_leasing,
                        "rat_wycena": q_rat_wycena,
                        "inwestycje_nieruchomosci": q_inwest_nieruch,
                        "udzialy_wycena": q_udzialy,
                        "zapasy_wycena": q_zapasy_wycena,
                        "podatek_odroczony": q_podatek,
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
            map_cols = st.columns(len(doc_mapping))
            for i, (fname, ddata) in enumerate(doc_mapping.items()):
                with map_cols[i]:
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

        # Potrzebujemy company_info do sprawdzenia formy prawnej
        _ci_for_notes = {"forma_prawna": company_forma}
        selected_notes = select_applicable_notes(doc_mapping, _ci_for_notes)
        st.session_state["selected_notes"] = selected_notes

        with results_container:
            st.subheader(f"📝 Dobrano {len(selected_notes)} not objaśniających")
            notes_display = format_notes_for_display(selected_notes)
            for item in notes_display:
                st.markdown(f'<span class="validation-ok">{item["msg"]}</span>',
                            unsafe_allow_html=True)
            if not selected_notes:
                st.warning("Nie dobrano żadnych not — sprawdź wgrane dokumenty.")

        # ── KROK 4: Generowanie ─────────────────────────────────────────────
        status_text.info("🤖 Krok 4/5: Generowanie przez Claude 3.5 Sonnet (może potrwać 1-2 min)...")
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
            "selected_notes": selected_notes,
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

        # Wyczyść flagi pipeline'u (generowanie zakończone)
        st.session_state.pop("run_generation", None)
        st.session_state.pop("missing_decision", None)

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
