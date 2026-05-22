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
                     "konta aktywne", "obroty debetowe", "obroty kredytowe",
                     "saldo dt", "saldo ct", "stan na", "obroty za okres",
                     "zestawienie kont", "plan kont"],
    },
}


def identify_document_type(text: str) -> str:
    """
    Identyfikacja typu dokumentu finansowego.
    Krok 1: szuka nagłówka w pierwszych 500 znakach (tytuł dokumentu).
    Krok 2: jeśli brak — liczy słowa kluczowe w całym tekście.
    """
    # Nagłówki które jednoznacznie identyfikują dokument
    HEADER_RULES = [
        ("ZOiS",                    ["zestawienie obrotów i sald", "obroty i salda", "zois"]),
        ("BILANS",                   ["bilans na dzień", "bilans jednostki", "aktywa i pasywa"]),
        ("RZiS",                     ["rachunek zysków i strat", "rachunek wyników",
                                       "wynik finansowy netto"]),
        ("ŚRODKI TRWAŁE",            ["tabela środków trwałych", "środki trwałe i wartości",
                                       "zestawienie środków trwałych"]),
        ("POLITYKA RACHUNKOWOŚCI",   ["polityka rachunkowości", "zasady rachunkowości przyjęte"]),
        ("PRZEPŁYWY PIENIĘŻNE",      ["rachunek przepływów pieniężnych", "cash flow"]),
    ]

    header = text[:500].lower()
    for doc_type, phrases in HEADER_RULES:
        if any(phrase in header for phrase in phrases):
            return doc_type

    # Fallback: zliczanie słów kluczowych w całym tekście
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
# MODUŁ 4: GENEROWANIE INFORMACJI DODATKOWEJ PRZEZ CLAUDE
# ═══════════════════════════════════════════════════════════════════════════════

SYSTEM_PROMPT = """Jesteś biegłym rewidentem i ekspertem ds. rachunkowości, specjalizującym się w polskim prawie bilansowym.
Sporządzasz "Informację Dodatkową" do sprawozdania finansowego zgodnie z Ustawą o Rachunkowości.

ZASADA NADRZĘDNA: Generuj WYŁĄCZNIE noty które mają niezerowe wartości wynikające z dokumentów.
Jeśli pozycja wynosi 0 lub nie ma danych — POMIJASZ całą notę. Żadnych "Nie dotyczy", żadnych zer.

═══════════════════════════════════════════════════════════════
CZĘŚĆ 1. WPROWADZENIE DO SPRAWOZDANIA FINANSOWEGO
═══════════════════════════════════════════════════════════════

1.1 Dane identyfikacyjne jednostki
Wypełnij na podstawie danych z panelu: nazwa, forma prawna, siedziba, NIP, KRS, REGON, PKD,
data rejestracji, okres sprawozdawczy. Dodaj zdanie o kontynuacji działalności (lub zagrożeniu).

1.2 Zasady (polityka) rachunkowości
Na podstawie dokumentu Polityki lub odpowiedzi z ankiety — opisz:
- podstawę prawną, wariant RZiS, rodzaj sprawozdania
- leasing, podatek odroczony, inne istotne zasady

1.3 Metody wyceny aktywów i pasywów
Opisz wycenę: WNiP, ST, inwestycji, zapasów (metoda FIFO/LIFO/śr.ważona),
należności i zobowiązań, walut obcych.

1.4 Metody amortyzacji i stosowane stawki
Metoda (liniowa/degresywna), okresy dla grup ST i WNiP, próg niskocennych.

1.5 Zasady rozliczania przychodów i kosztów
Moment ujęcia przychodów, rezerwy na urlopy/emerytury, RMK.

1.6 Korekty błędów i zmiany polityki rachunkowości
Jeśli brak zmian — napisz jedno zdanie. Jeśli są — opisz.

═══════════════════════════════════════════════════════════════
CZĘŚĆ 2. NOTY — GENERUJ TYLKO TE Z NIEZEROWYMI WARTOŚCIAMI
═══════════════════════════════════════════════════════════════

Dla każdej noty: sprawdź czy w dokumentach (Bilans, RZiS, ZOiS, ST) są niezerowe wartości.
Jeśli tak — wygeneruj notę w formie tabeli Markdown. Jeśli nie — POMIŃ bez komentarza.

DOSTĘPNE NOTY (1-79):

NOTY O AKTYWACH TRWAŁYCH:
Nota 1  — Zmiana wartości początkowej i umorzenia ŚRODKÓW TRWAŁYCH
          Tabela: grupy ST × (wartość brutto BO, zwiększenia, zmniejszenia, wartość brutto BZ,
          umorzenie BO, amortyzacja roku, umorzenie BZ, wartość netto BZ)
          → generuj gdy ST > 0

Nota 2  — Zmiana wartości WARTOŚCI NIEMATERIALNYCH I PRAWNYCH (analogiczna struktura jak Nota 1)
          → generuj gdy WNiP > 0

Nota 3  — Zmiana wartości INWESTYCJI DŁUGOTERMINOWYCH
          → generuj gdy inwestycje długoterminowe > 0

Nota 4  — Odpisy aktualizujące wartość DŁUGOTERMINOWYCH AKTYWÓW NIEFINANSOWYCH
          → generuj gdy są odpisy

Nota 5  — Odpisy aktualizujące wartość DŁUGOTERMINOWYCH AKTYWÓW FINANSOWYCH
          → generuj gdy są odpisy

Nota 6  — Koszty zakończonych prac rozwojowych oraz WARTOŚĆ FIRMY
          → generuj gdy wartość firmy lub prace rozwojowe > 0

Nota 7  — GRUNTY użytkowane wieczyście
          → generuj gdy grunty w wieczystym użytkowaniu > 0

Nota 8  — Środki trwałe NIEAMORTYZOWANE (ewidencja pozabilansowa)
          → generuj gdy są

Nota 9  — PAPIERY WARTOŚCIOWE lub prawa
          → generuj gdy są

NOTY O NALEŻNOŚCIACH:
Nota 10 — Odpisy aktualizujące wartość NALEŻNOŚCI
          Tabela: grupy należności × (stan BO, zwiększenia, wykorzystanie, uznanie za zbędne, stan BZ)
          → generuj gdy odpisy > 0

Nota 60 — STRUKTURA NALEŻNOŚCI (przeterminowane vs nieprzeterminowane, wg dni)
          → generuj gdy należności > 0

Nota 61 — NALEŻNOŚCI według okresów wymagalności (do 30 dni, 31-90, 91-180, >181, >12 mies.)
          → generuj gdy należności > 0

NOTY O KAPITAŁACH:
Nota 11 — Struktura własności KAPITAŁU PODSTAWOWEGO — spółki akcyjne (serie akcji)
          → generuj dla SA i PSA

Nota 12 — Struktura własności KAPITAŁU PODSTAWOWEGO — spółka z o.o.
          Tabela: wspólnik, wartość nominalna udziałów, % udziału
          → generuj zawsze dla sp. z o.o.

Nota 13 — Zmiany stanów KAPITAŁÓW zapasowego i rezerwowego
          Tabela: stan BO, zwiększenia (agio, podział zysku, dopłaty), zmniejszenia, stan BZ
          → generuj gdy kapitał zapasowy lub rezerwowy ≠ 0

Nota 14 — Zmiany w stanie KAPITAŁU Z AKTUALIZACJI WYCENY
          → generuj gdy kapitał z aktualizacji > 0

Nota 15 — Propozycja PODZIAŁU ZYSKU za rok obrotowy
          Tabela: zysk netto, nierozliczony wynik lat ubiegłych, razem do podziału,
          proponowany podział (dywidenda, kapitał zapasowy, pokrycie straty, inne)
          → generuj gdy zysk netto > 0

Nota 16 — Propozycja POKRYCIA STRATY za rok obrotowy
          → generuj gdy strata netto

NOTY O ZOBOWIĄZANIACH I REZERWACH:
Nota 17 — REZERWY na koszty i zobowiązania
          Tabela: rodzaj rezerwy × (stan BO, zwiększenia, wykorzystanie, rozwiązanie, stan BZ)
          → generuj gdy rezerwy > 0

Nota 18 — ODROCZONY PODATEK DOCHODOWY
          Tabela: aktywa i rezerwy z tyt. CIT odroczonego, różnice przejściowe
          → generuj gdy podatek odroczony ≠ 0

Nota 19 — ZOBOWIĄZANIA według okresów wymagalności
          Tabela: rodzaje zobowiązań × (do 1 roku, 1-3 lata, 3-5 lat, >5 lat) — BO i BZ
          → generuj gdy zobowiązania > 0

Nota 20 — Wykaz ZOBOWIĄZAŃ ZABEZPIECZONYCH na majątku
          → generuj gdy są zabezpieczenia (hipoteka, zastaw, przewłaszczenie)

Nota 25 — Wykaz ZOBOWIĄZAŃ WARUNKOWYCH
          → generuj gdy są zobowiązania warunkowe (poręczenia, gwarancje)

Nota 26 — Zobowiązania warunkowe ZABEZPIECZONE na majątku
          → generuj gdy są

Nota 45 — Zobowiązania z tytułu EMERYTUR i podobnych świadczeń
          → generuj gdy są

Nota 72 — Zobowiązania z tytułu GWARANCJI I PORĘCZEŃ w imieniu osób trzecich
          → generuj gdy są

Nota 73 — Zobowiązania długoterminowe o pozostałym OKRESIE WYMAGALNOŚCI > 5 lat
          → generuj gdy są

NOTY O ROZLICZENIACH MIĘDZYOKRESOWYCH:
Nota 21 — CZYNNE rozliczenia międzyokresowe
          Tabela: wyszczególnienie × (stan BO, zwiększenia, zmniejszenia, stan BZ)
          → generuj gdy RMK czynne > 0

Nota 22 — ROZLICZENIA MIĘDZYOKRESOWE PRZYCHODÓW
          → generuj gdy RMP > 0

NOTY O PRZYCHODACH I KOSZTACH:
Nota 29 — STRUKTURA RZECZOWA I TERYTORIALNA przychodów (kraj/eksport/WDT)
          Tabela: wyroby/usługi/towary × (kraj rok poprz., kraj rok bież., eksport, WDT)
          → generuj gdy przychody > 0

Nota 31 — KOSZTY RODZAJOWE i koszt wytworzenia produktów na własne potrzeby
          Tabela: amortyzacja, zużycie mat., usługi obce, podatki, wynagrodzenia,
          ubezpieczenia, pozostałe, razem × (rok poprzedni, rok bieżący)
          → generuj zawsze gdy RZiS dostępny

Nota 34 — Przychody i koszty DZIAŁALNOŚCI ZANIECHANEJ
          → generuj gdy były

Nota 39 — Pozycje przychodów/kosztów o NADZWYCZAJNEJ WARTOŚCI lub charakterze
          → generuj gdy są

NOTY PODATKOWE:
Nota 35 — ROZLICZENIE RÓŻNICY między podstawą CIT a wynikiem brutto
          Tabela pełna: zysk brutto, trwałe różnice (przychody zwolnione, NKUP),
          przejściowe różnice, dochód/strata podatkowa, podatek należny
          → generuj zawsze gdy CIT > 0

Nota 59 — FAKTYCZNIE ZAPŁACONY podatek dochodowy
          Tabela: CIT z RZiS, zmiana rezerwy, CIT wg deklaracji, zmiana należności, CIT zapłacony
          → generuj zawsze gdy CIT > 0

Nota 18 — Podatek ODROCZONY (aktywa i rezerwy)
          → generuj gdy podatek odroczony ≠ 0

NOTY O ŚRODKACH PIENIĘŻNYCH:
Nota 41 — STRUKTURA ŚRODKÓW PIENIĘŻNYCH przyjęta do rachunku przepływów
          Tabela: kasa, rachunki bankowe, inne × (rok poprz., rok bież., zmiana, ograniczone)
          → generuj gdy są środki pieniężne

Nota 63 — ŚRODKI NA RACHUNKU VAT (split payment)
          → generuj gdy są środki na rachunku VAT

NOTY O ZATRUDNIENIU I WYNAGRODZENIACH:
Nota 43 — PRZECIĘTNE ZATRUDNIENIE w podziale na grupy zawodowe
          Tabela: pracownicy umysłowi, robotnicy, zagranica, uczniowie, urlopy, razem
          → generuj gdy zatrudnienie > 0

Nota 44 — WYNAGRODZENIA organów spółki (zarząd, rada nadzorcza)
          → generuj gdy były wynagrodzenia

Nota 46 — ZALICZKI, KREDYTY, POŻYCZKI dla członków organów
          → generuj gdy były

Nota 70 — Zaliczki i pożyczki dla OSÓB Z ORGANÓW jednostki
          → generuj gdy były

NOTY O JEDNOSTKACH POWIĄZANYCH:
Nota 52 — Spółki z ZAANGAŻOWANIEM KAPITAŁOWYM jednostki
          → generuj gdy są udziały/akcje w innych spółkach

Nota 76 — TRANSAKCJE Z JEDNOSTKAMI POWIĄZANYMI
          Tabela: jednostka × (charakter transakcji, należności, zobowiązania, przychody, koszty)
          → generuj gdy są transakcje z podmiotami powiązanymi

NOTY SZCZEGÓŁOWE:
Nota 27 — Aktywa niebędące inst. finansowymi WYCENIANE WG WARTOŚCI GODZIWEJ
Nota 28 — Zmiany kapitału z aktualizacji wyceny AKTYWÓW NIEFINANSOWYCH
Nota 30 — UMOWY O USŁUGI DŁUGOTERMINOWE
Nota 32 — Odpisy aktualizujące ST
Nota 33 — Odpisy aktualizujące ZAPASY
Nota 36 — Koszt wytworzenia ST W BUDOWIE
Nota 37 — Odsetki i różnice kursowe w CENIE NABYCIA
Nota 38 — NAKŁADY na niefinansowe aktywa trwałe
Nota 40 — KURSY WALUT do wyceny bilansu i RZiS
Nota 42 — PRZEPŁYWY PIENIĘŻNE z działalności operacyjnej (metoda pośrednia)
Nota 47 — Wynagrodzenie FIRMY AUDYTORSKIEJ
Nota 48 — Błędy lat ubiegłych odniesione na KAPITAŁ
Nota 49 — Skutki ZMIAN POLITYKI RACHUNKOWOŚCI
Nota 50 — Dane zapewniające PORÓWNYWALNOŚĆ
Nota 57 — Różnica zmiana stanu ZOBOWIĄZAŃ KT bilans vs przepływy
Nota 58 — Różnica zmiana stanu ZAPASÓW bilans vs przepływy
Nota 65 — Instrumenty finansowe wg WARTOŚCI GODZIWEJ
Nota 68 — UDZIAŁY/AKCJE WŁASNE
Nota 69 — WARTOŚĆ FIRMY
Nota 74 — UKRYTE ZYSKI (estoński CIT)
Nota 77 — Prezentacja pozycji bilansu wg rodzajów DZIAŁALNOŚCI
Nota 78 — Prezentacja RZiS wariant PORÓWNAWCZY wg działalności
Nota 79 — Prezentacja RZiS wariant KALKULACYJNY wg działalności

═══════════════════════════════════════════════════════════════
CZĘŚĆ 3. POZOSTAŁE INFORMACJE
═══════════════════════════════════════════════════════════════

3.1 Przeciętne zatrudnienie — jeśli nie wygenerowano Noty 43
3.2 Wynagrodzenia organów — jeśli nie wygenerowano Noty 44
3.3 Zdarzenia po dniu bilansowym
3.4 Inne istotne informacje

═══════════════════════════════════════════════════════════════
CZĘŚĆ 4. ANALIZA GRAFICZNA (na końcu dokumentu)
═══════════════════════════════════════════════════════════════

Wygeneruj opisy trzech wykresów do umieszczenia na końcu dokumentu:
[WYKRES 1: Struktura pasywów — opis danych do wykresu kołowego]
[WYKRES 2: Analiza wyniku finansowego — opis danych do wykresu słupkowego]
[WYKRES 3: Struktura aktywów obrotowych — opis danych do wykresu kołowego]

FORMAT ODPOWIEDZI:
- Używaj nagłówków Markdown (##, ###)
- Tabele w formacie Markdown z wyrównaniem
- Liczby z separatorem tysięcy i dokładnością do 2 miejsc po przecinku (np. 1 234 567,89 PLN)
- Przy brakujących danych szczegółowych pisz konkretnie czego brakuje w nawiasie kwadratowym
- NIE pisz "Nie dotyczy", NIE generuj not z samymi zerami
"""


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
        polityka_blok = """
\n📋 ZASADY RACHUNKOWOŚCI (odpowiedzi udzielone przez użytkownika — brak załączonej Polityki Rachunkowości):
- Zasady ustalania wyniku finansowego: {wynik}
- Wycena zapasów: {zapasy}
- Amortyzacja środków trwałych: {amort}
- Wycena należności: {nal}
- Sposób sporządzania sprawozdania: {spr}
- Rezerwa/aktywa z tytułu podatku odroczonego: {pod}
- Ujęcie leasingu: {leas}
{uwagi_blok}
Na podstawie powyższych odpowiedzi wypełnij DOKŁADNIE sekcje 1.2–1.5 Informacji Dodatkowej.\n""".format(
            wynik=pa.get("wynik_finansowy", ""),
            zapasy=pa.get("wycena_zapasow", ""),
            amort=pa.get("amortyzacja", ""),
            nal=pa.get("wycena_naleznosci", ""),
            spr=pa.get("sposob_sprawozdania", ""),
            pod="TAK" if pa.get("podatek_odroczony") else "NIE",
            leas=pa.get("leasing", ""),
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
        f"ZATRUDNIENIE ŚREDNIE ROK BIEŻĄCY: {info.get('zatrudnienie_biezacy', 0)} etatów",
        f"ZATRUDNIENIE ŚREDNIE ROK POPRZEDNI: {info.get('zatrudnienie_poprzedni', 0)} etatów",
        f"UWAGI DO ZATRUDNIENIA: {info.get('zatrudnienie_uwagi', '')}",
        f"ROK OBROTOWY: {year}",
        polityka_blok,
        zagrozenie_blok,
        "=" * 60,
        "WYCIĄGI Z DOKUMENTÓW FINANSOWYCH:",
    ]
    for filename, doc_data in doc_mapping.items():
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
Formatuj wyraźnie nagłówkami i akapitami.

SZCZEGÓLNA UWAGA — NOTA KOSZTY RODZAJOWE (pkt 2.12):
Przeszukaj dokument RZiS (Rachunek Zysków i Strat) i wyciągnij WSZYSTKIE pozycje kosztów rodzajowych
dla roku bieżącego ORAZ poprzedniego (kolumna "poprzedni rok" lub "rok ubiegły" w RZiS).
Sporządź tabelę Markdown z dokładnymi kwotami. Jeśli RZiS jest w wariancie kalkulacyjnym
a nie porównawczym — zaznacz to i użyj dostępnych danych kosztowych."""

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
# MODUŁ 5: EKSPORT DO WORD (.docx)
# ═══════════════════════════════════════════════════════════════════════════════

def add_horizontal_rule(doc: Document):
    """Dodaje poziomą linię jako separator."""
    p = doc.add_paragraph()
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "2d6a9f")
    pBdr.append(bottom)
    pPr.append(pBdr)


# ─── KOLORY (wzór z 28.03.2026) ─────────────────────────────────────────────
NAVY  = RGBColor(0x1B, 0x2A, 0x4A)   # nagłówki, tło tabel
BLUE  = RGBColor(0x2D, 0x6A, 0x9F)   # Heading 2, akcenty
WHITE = RGBColor(0xFF, 0xFF, 0xFF)   # tekst na ciemnym tle
DARK  = RGBColor(0x22, 0x22, 0x22)   # tekst normalny
GRAY6 = RGBColor(0x66, 0x66, 0x66)   # dane na stronie tytułowej
GRAY9 = RGBColor(0x99, 0x99, 0x99)   # stopka, data
LIGHT = "EBF3FB"                      # naprzemienne wiersze tabel


def _tcPr_shading(cell, fill_hex: str):
    """Ustawia kolor tła komórki."""
    from docx.oxml.ns import qn as _qn
    from docx.oxml import OxmlElement as _OE
    tcPr = cell._tc.get_or_add_tcPr()
    shd = _OE("w:shd")
    shd.set(_qn("w:val"),   "clear")
    shd.set(_qn("w:color"), "auto")
    shd.set(_qn("w:fill"),  fill_hex)
    tcPr.append(shd)


def _cell_margins(cell, val: int = 80):
    """Ustawia marginesy wewnętrzne komórki."""
    from docx.oxml.ns import qn as _qn
    from docx.oxml import OxmlElement as _OE
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = _OE("w:tcMar")
    for side in ("top", "bottom", "left", "right"):
        el = _OE(f"w:{side}")
        el.set(_qn("w:w"),    str(val))
        el.set(_qn("w:type"), "dxa")
        tcMar.append(el)
    tcPr.append(tcMar)


def _add_runs(paragraph, text: str, bold=False, color=None, size_pt=None):
    """Dodaje run z opcjonalnym formatowaniem."""
    run = paragraph.add_run(text)
    run.font.name = "Calibri"
    if bold is not None:
        run.font.bold = bold
    if color:
        run.font.color.rgb = color
    if size_pt:
        run.font.size = Pt(size_pt)
    return run


def _add_inline_text(doc, line: str, style=None):
    """Akapit z obsługą **bold** inline."""
    p = doc.add_paragraph(style=style) if style else doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after  = Pt(2)
    parts = re.split(r"(\*\*[^*]+\*\*)", line)
    for part in parts:
        if part.startswith("**") and part.endswith("**"):
            _add_runs(p, part[2:-2], bold=True)
        else:
            _add_runs(p, part, bold=False)
    return p


def _render_md_table(doc, table_lines: list):
    """Renderuje tabelę Markdown jako tabelę Word — styl wzoru z 28.03.2026."""
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.oxml.ns import qn as _qn

    # Odfiltruj separator (|---|)
    rows_raw = [l for l in table_lines
                if not re.match(r"^\|[\s\-:|]+\|$", l.strip())]
    if not rows_raw:
        return

    rows = []
    for line in rows_raw:
        cells = [c.strip() for c in line.strip().strip("|").split("|")]
        rows.append(cells)

    ncols = max(len(r) for r in rows)
    rows  = [r + [""] * (ncols - len(r)) for r in rows]

    # Szerokości: łącznie ~8500 DXA (A4 z marginesami 2.5cm)
    col_w = 8500 // ncols

    table = doc.add_table(rows=len(rows), cols=ncols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"

    for ri, row_data in enumerate(rows):
        row = table.rows[ri]
        is_header = (ri == 0)
        for ci, raw_text in enumerate(row_data):
            cell = row.cells[ci]
            cell.width = Pt(col_w)

            # Tło
            if is_header:
                _tcPr_shading(cell, "1B2A4A")
            elif ri % 2 == 1:
                _tcPr_shading(cell, LIGHT)

            _cell_margins(cell, 80)

            # Tekst — usuń bold markdown
            clean = re.sub(r"[*][*]([^*]+)[*][*]", r"\1", raw_text)
            clean = "".join(ch for ch in clean if ord(ch) >= 32 or ord(ch) in (9,10,13))
            is_bold_md = "**" in raw_text

            p = cell.paragraphs[0]
            p.clear()
            run = p.add_run(clean)
            run.font.name = "Calibri"
            run.font.size = Pt(9)
            run.font.bold = is_header or is_bold_md
            run.font.color.rgb = WHITE if is_header else DARK

            # Liczby do prawej, nagłówki do środka
            if is_header:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            elif re.search(r"\d", clean):
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    doc.add_paragraph()  # odstęp po tabeli


def save_to_word(generated_text: str, company_name: str, year: int) -> bytes:
    """
    Generuje .docx w formacie identycznym z wzorem z 28.03.2026.
    Calibri, marginesy 2.5cm, nagłówki NAVY/BLUE, tabele z granatowym nagłówkiem.
    """
    from docx.oxml.ns import qn as _qn
    from docx.oxml import OxmlElement as _OE

    doc = Document()

    # ── Style globalne — Calibri 10pt ─────────────────────────────────────
    normal = doc.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(10)

    # Heading 1: 14pt, bold, NAVY
    h1s = doc.styles["Heading 1"]
    h1s.font.name  = "Calibri"
    h1s.font.size  = Pt(14)
    h1s.font.bold  = True
    h1s.font.color.rgb = NAVY

    # Heading 2: 12pt, bold, BLUE
    h2s = doc.styles["Heading 2"]
    h2s.font.name  = "Calibri"
    h2s.font.size  = Pt(12)
    h2s.font.bold  = True
    h2s.font.color.rgb = BLUE

    # Heading 3: 11pt, bold, NAVY
    h3s = doc.styles["Heading 3"]
    h3s.font.name  = "Calibri"
    h3s.font.size  = Pt(11)
    h3s.font.bold  = True
    h3s.font.color.rgb = NAVY

    # ── Marginesy (wzór: 2.5cm lewy/prawy, 2.0cm górny, 1.8cm dolny) ─────
    from docx.shared import Cm
    for section in doc.sections:
        section.left_margin   = Cm(2.5)
        section.right_margin  = Cm(2.5)
        section.top_margin    = Cm(2.0)
        section.bottom_margin = Cm(1.8)

    # ── STRONA TYTUŁOWA ───────────────────────────────────────────────────
    # Wiersz z nazwą spółki i tytułem w stopce (kursywa, szara)
    header_line = doc.add_paragraph()
    _add_runs(header_line,
              f"{company_name} | Informacja Dodatkowa {year}",
              bold=False, color=GRAY6, size_pt=9)
    header_line.paragraph_format.space_after = Pt(12)

    # Tytuł główny
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _add_runs(title_p, "INFORMACJA DODATKOWA",
              bold=True, color=NAVY, size_pt=22)
    title_p.paragraph_format.space_before = Pt(24)
    title_p.paragraph_format.space_after  = Pt(4)

    sub1 = doc.add_paragraph()
    sub1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _add_runs(sub1, "do sprawozdania finansowego",
              bold=False, color=BLUE, size_pt=14)

    sub2 = doc.add_paragraph()
    sub2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _add_runs(sub2, f"za rok obrotowy {year}",
              bold=False, color=BLUE, size_pt=14)
    sub2.paragraph_format.space_after = Pt(16)

    # Dane spółki
    from datetime import datetime
    dane = {
        "Jednostka":         company_name,
        "Okres sprawozdawczy": f"01.01.{year} — 31.12.{year}",
    }
    for label, value in dane.items():
        p = doc.add_paragraph()
        _add_runs(p, f"{label}: ", bold=False, color=GRAY6, size_pt=10)
        _add_runs(p, value, bold=True, color=GRAY6, size_pt=10)

    # Data wygenerowania
    p_date = doc.add_paragraph()
    _add_runs(p_date, f"Wygenerowano: {datetime.now().strftime('%d.%m.%Y')}",
              bold=False, color=GRAY9, size_pt=9)
    p_date.paragraph_format.space_before = Pt(8)

    doc.add_page_break()

    # ── TREŚĆ ─────────────────────────────────────────────────────────────
    lines = generated_text.split("\n")
    i = 0
    while i < len(lines):
        line  = lines[i]
        strip = line.strip()

        if not strip:
            i += 1
            continue

        # Tabela Markdown
        if strip.startswith("|"):
            tbl = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                tbl.append(lines[i].strip())
                i += 1
            _render_md_table(doc, tbl)
            continue

        # Nagłówki
        if strip.startswith("#### "):
            h = doc.add_heading(strip[5:], level=4)
            for r in h.runs:
                r.font.name = "Calibri"
        elif strip.startswith("### "):
            doc.add_heading(strip[4:], level=3)
        elif strip.startswith("## "):
            doc.add_heading(strip[3:], level=2)
        elif strip.startswith("# "):
            doc.add_heading(strip[2:], level=1)

        # Linia pozioma
        elif strip.startswith("---"):
            add_horizontal_rule(doc)

        # Lista
        elif strip.startswith("- ") or strip.startswith("* "):
            _add_inline_text(doc, strip[2:], style="List Bullet")

        elif re.match(r"^\d+\.\s", strip):
            _add_inline_text(doc, strip, style="List Number")

        # Zwykły tekst
        else:
            _add_inline_text(doc, strip)

        i += 1

    # ── STOPKA ────────────────────────────────────────────────────────────
    add_horizontal_rule(doc)
    foot = doc.add_paragraph(
        f"Informacja Dodatkowa | {company_name} | Rok {year}"
    )
    foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for r in foot.runs:
        r.font.name  = "Calibri"
        r.font.size  = Pt(8)
        r.font.color.rgb = GRAY9

    # ── ZAPIS ─────────────────────────────────────────────────────────────
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def _sanitize_text(text: str) -> str:
    """Usuwa znaki kontrolne i NULL które są niezgodne z XML/docx."""
    import unicodedata
    result = []
    for ch in text:
        cp = ord(ch)
        # Dozwolone: tab (9), LF (10), CR (13), oraz znaki >= 32
        if cp in (9, 10, 13) or cp >= 32:
            # Dodatkowo wyklucz surrogaty i znaki specjalne XML
            cat = unicodedata.category(ch)
            if cat != "Cs":  # Cs = surrogate
                result.append(ch)
    return "".join(result)

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

    st.divider()
    st.markdown("""
    **📋 Obsługiwane dokumenty:**
    - 🏦 Bilans
    - 📈 Rachunek Zysków i Strat
    - 🏗️ Tabela środków trwałych
    - 💸 Przepływy pieniężne
    - 📜 Polityka rachunkowości
    - ⚖️ Zestawienie Obrotów i Sald (ZOiS)
    """)


# ── Główna sekcja ────────────────────────────────────────────────────────────
col1, col2 = st.columns([1, 1])

with col1:
    st.markdown('<div class="step-card"><b>📁 Krok 1:</b> Wgraj dokumenty PDF sprawozdania</div>',
                unsafe_allow_html=True)
    uploaded_files = st.file_uploader(
        "Wybierz pliki PDF",
        type=["pdf"],
        accept_multiple_files=True,
        help="Wgraj 3-6 dokumentów: bilans, RZiS, noty, przepływy pieniężne"
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

# ═══════════════════════════════════════════════════════════════════════════════
# MASZYNA STANÓW GENEROWANIA
# ═══════════════════════════════════════════════════════════════════════════════
#
# Stany (session_state["app_state"]):
#   "idle"            → czeka na kliknięcie Generuj
#   "parsing"         → parsuje PDF i mapuje dokumenty
#   "confirm_missing" → pyta o brakujące dokumenty
#   "polityka"        → pyta o zasady rachunkowości
#   "generating"      → wywołuje Claude i zapisuje docx
#   "done"            → pokazuje wyniki
#   "error"           → pokazuje błąd
#
# Zasada: st.rerun() zawsze wraca do TEGO bloku który
# sprawdza stan i wykonuje właściwy krok.
# ═══════════════════════════════════════════════════════════════════════════════

st.divider()

def _reset_state():
    """Resetuje maszynę stanów do początku."""
    for key in ["app_state", "parsed_docs", "doc_mapping", "missing_docs",
                "polityka_answers", "generated_text", "docx_bytes"]:
        st.session_state.pop(key, None)

def _set_state(state: str):
    st.session_state["app_state"] = state

def _get_state() -> str:
    return st.session_state.get("app_state", "idle")

# ── Przycisk Generuj ─────────────────────────────────────────────────────────
run_disabled = not (anthropic_key and uploaded_files and company_name)
if st.button("🚀 Generuj Informację Dodatkową", type="primary",
             disabled=run_disabled, use_container_width=True):
    _reset_state()
    _set_state("parsing")
    st.rerun()

# ── STAN: parsing ─────────────────────────────────────────────────────────────
if _get_state() == "parsing":
    progress_bar = st.progress(0)
    status_text = st.empty()
    try:
        status_text.info("📄 Krok 1/4: Parsowanie dokumentów PDF...")
        progress_bar.progress(10)

        def update_progress(val, msg):
            progress_bar.progress(int(10 + val * 20))
            status_text.info(f"📄 {msg}")

        if llama_key:
            parsed = parse_documents_llamaparse(uploaded_files, llama_key, update_progress)
        else:
            parsed = parse_documents_fallback(uploaded_files, update_progress)

        status_text.info("🗂️ Krok 2/4: Mapowanie dokumentów...")
        progress_bar.progress(40)
        doc_mapping = map_documents(parsed)

        st.session_state["parsed_docs"] = parsed
        st.session_state["doc_mapping"] = doc_mapping

        missing = check_missing_documents(doc_mapping)
        st.session_state["missing_docs"] = missing

        if missing:
            _set_state("confirm_missing")
        else:
            _set_state("polityka")

        progress_bar.empty()
        status_text.empty()
        st.rerun()

    except Exception as e:
        progress_bar.empty()
        status_text.empty()
        _set_state("error")
        st.session_state["error_msg"] = str(e)
        st.rerun()

# ── STAN: confirm_missing ─────────────────────────────────────────────────────
elif _get_state() == "confirm_missing":
    missing = st.session_state.get("missing_docs", [])
    doc_mapping = st.session_state.get("doc_mapping", {})

    # Pokaż raport mapowania
    st.subheader("📋 Rozpoznane dokumenty")
    cols = st.columns(max(len(doc_mapping), 1))
    for i, (fname, ddata) in enumerate(doc_mapping.items()):
        with cols[i % len(cols)]:
            st.markdown(f"""<div class="metric-box">
                <b>{ddata['type']}</b><br>
                <small>{fname}</small>
            </div>""", unsafe_allow_html=True)

    st.warning("⚠️ Nie znaleziono wszystkich dokumentów w wgranych plikach.")
    st.markdown("**Brakujące dokumenty:**")
    for dt in missing:
        info_dt = REQUIRED_DOC_TYPES[dt]
        st.markdown(f"- {info_dt['icon']} **{info_dt['label']}** — {info_dt['desc']}")

    st.info("💡 Jeśli plik zawiera kilka dokumentów w jednym PDF — spróbuj wgrać je jako osobne pliki.")
    st.markdown("---")
    st.markdown("**Co chcesz zrobić?**")

    col_a, col_b = st.columns(2)
    with col_a:
        if st.button("▶️ Kontynuuj bez brakujących dokumentów",
                     use_container_width=True, type="primary"):
            _set_state("polityka")
            st.rerun()
    with col_b:
        if st.button("📁 Anuluj — chcę dodać brakujące pliki",
                     use_container_width=True):
            _reset_state()
            st.rerun()

# ── STAN: polityka ────────────────────────────────────────────────────────────
elif _get_state() == "polityka":
    doc_mapping = st.session_state.get("doc_mapping", {})
    types_found = {d["type"] for d in doc_mapping.values()}

    if "POLITYKA RACHUNKOWOŚCI" not in types_found:
        st.warning("📜 Nie załączono dokumentu **Polityki Rachunkowości**. Wypełnij poniższe pytania.")

        with st.form("polityka_form"):
            st.subheader("📋 Zasady rachunkowości — pytania uzupełniające")

            q1 = st.selectbox("1. Zasady ustalania wyniku finansowego:", options=[
                "Wariant porównawczy (układ rodzajowy kosztów)",
                "Wariant kalkulacyjny (układ funkcjonalny kosztów)",
            ])
            q2_wycena = st.selectbox("2a. Metoda wyceny zapasów:", options=[
                "FIFO (pierwsze weszło, pierwsze wyszło)",
                "LIFO (ostatnie weszło, pierwsze wyszło)",
                "Cena przeciętna (średnia ważona)",
                "Ceny ewidencyjne z odchyleniami",
                "Nie dotyczy (brak zapasów)",
            ])
            q2_st = st.selectbox("2b. Metoda amortyzacji środków trwałych:", options=[
                "Liniowa (równomierne odpisy przez cały okres)",
                "Degresywna (przyspieszone odpisy na początku)",
                "Jednorazowy odpis (niskocenne ST do 10 000 zł)",
                "Mieszana (liniowa i jednorazowa)",
            ])
            q2_nal = st.selectbox("2c. Wycena należności:", options=[
                "W wartości nominalnej z odpisami aktualizującymi",
                "W wartości nominalnej bez odpisów aktualizujących",
                "W wartości godziwej",
            ])
            q3 = st.selectbox("3. Sposób sporządzania sprawozdania finansowego:", options=[
                "Pełne sprawozdanie finansowe (standardowe)",
                "Uproszczone sprawozdanie finansowe (art. 46 ust. 5 UoR — jednostki małe)",
                "Sprawozdanie według Załącznika nr 4 UoR (mikro jednostki)",
                "Sprawozdanie według Załącznika nr 5 UoR (małe jednostki NGO)",
            ])
            q4_podatek = st.checkbox(
                "Jednostka tworzy rezerwę i aktywa z tytułu odroczonego podatku dochodowego",
                value=True
            )
            q5_leasing = st.selectbox("Ujęcie leasingu:", options=[
                "Według UoR (leasing operacyjny/finansowy wg ekonomicznej treści)",
                "Leasing operacyjny — wszystkie umowy traktowane jako operacyjny",
                "Nie dotyczy (brak umów leasingowych)",
            ])
            uwagi = st.text_area("Dodatkowe uwagi (opcjonalnie):", height=80)

            if st.form_submit_button("✅ Zatwierdź i generuj", use_container_width=True, type="primary"):
                st.session_state["polityka_answers"] = {
                    "wynik_finansowy": q1, "wycena_zapasow": q2_wycena,
                    "amortyzacja": q2_st, "wycena_naleznosci": q2_nal,
                    "sposob_sprawozdania": q3, "podatek_odroczony": q4_podatek,
                    "leasing": q5_leasing, "uwagi": uwagi,
                }
                _set_state("generating")
                st.rerun()
    else:
        st.session_state["polityka_answers"] = {}
        _set_state("generating")
        st.rerun()

# ── STAN: generating ──────────────────────────────────────────────────────────
elif _get_state() == "generating":
    doc_mapping = st.session_state.get("doc_mapping", {})
    polityka_answers = st.session_state.get("polityka_answers", {})

    progress_bar = st.progress(0)
    status_text = st.empty()
    results_container = st.container()

    try:
        # Walidacja
        status_text.info("✅ Krok 3/4: Walidacja spójności danych...")
        progress_bar.progress(55)
        validation_issues = validate_data_consistency(doc_mapping)

        with results_container:
            st.subheader("📋 Raport mapowania i walidacji")
            map_cols = st.columns(max(len(doc_mapping), 1))
            for i, (fname, ddata) in enumerate(doc_mapping.items()):
                with map_cols[i % len(map_cols)]:
                    st.markdown(f"""<div class="metric-box">
                        <b>{ddata['type']}</b><br>
                        <small>{fname}</small><br>
                        <small>{ddata['length']:,} znaków</small>
                    </div>""", unsafe_allow_html=True)

            st.subheader("🔎 Walidacja danych")
            css = {"OK": "validation-ok", "WARN": "validation-warn", "ERR": "validation-err"}
            for issue in validation_issues:
                st.markdown(f'<span class="{css.get(issue["level"], "")}">{issue["msg"]}</span>',
                            unsafe_allow_html=True)

        # Generowanie przez Claude
        status_text.info("🤖 Krok 4/4: Generowanie przez Claude 3.5 Sonnet...")
        progress_bar.progress(65)

        company_info = {
            "nazwa": company_name, "siedziba": company_siedziba,
            "nip": company_nip, "krs": company_krs,
            "regon": company_regon, "pkd": company_pkd,
            "data_rejestracji": company_data_rej, "forma_prawna": company_forma,
            "okres_od": str(okres_od), "okres_do": str(okres_do),
            "zagrozenie_kontynuacji": zagrozenie_kontynuacji,
            "zagrozenie_opis": zagrozenie_opis,
            "polityka_answers": polityka_answers,
            "zatrudnienie_biezacy": zatrudnienie_biezacy,
            "zatrudnienie_poprzedni": zatrudnienie_poprzedni,
            "zatrudnienie_uwagi": zatrudnienie_uwagi,
        }

        generated_text = generate_accounting_notes(
            doc_mapping=doc_mapping,
            anthropic_api_key=anthropic_key,
            company_name=company_name,
            year=fiscal_year,
            company_info=company_info,
            progress_callback=lambda v, m: progress_bar.progress(int(65 + v * 20))
        )

        # Zapis do Word
        status_text.info("💾 Generowanie pliku Word...")
        try:
            docx_bytes = save_to_word(
                _sanitize_text(generated_text), company_name, fiscal_year
            )
        except Exception as docx_err:
            import traceback
            st.error(f"❌ Błąd zapisu Word: {docx_err}")
            st.code(traceback.format_exc())
            st.text_area("Wygenerowany tekst (do wklejenia ręcznego):",
                         generated_text, height=400)
            st.stop()

        st.session_state["generated_text"] = generated_text
        st.session_state["docx_bytes"] = docx_bytes
        progress_bar.progress(100)
        status_text.success("✅ Informacja Dodatkowa wygenerowana pomyślnie!")
        _set_state("done")
        st.rerun()

    except anthropic.AuthenticationError:
        _set_state("error")
        st.session_state["error_msg"] = "Nieprawidłowy klucz API Anthropic."
        st.rerun()
    except anthropic.RateLimitError:
        _set_state("error")
        st.session_state["error_msg"] = "Przekroczono limit zapytań API. Poczekaj chwilę."
        st.rerun()
    except Exception as e:
        _set_state("error")
        st.session_state["error_msg"] = str(e)
        st.rerun()

# ── STAN: done ────────────────────────────────────────────────────────────────
elif _get_state() == "done":
    st.success("✅ Informacja Dodatkowa wygenerowana pomyślnie!")

    dl_col, _ = st.columns([1, 2])
    with dl_col:
        st.download_button(
            label="⬇️ Pobierz Informację Dodatkową (.docx)",
            data=st.session_state["docx_bytes"],
            file_name=f"informacja_dodatkowa_{company_name.replace(' ', '_')}_{fiscal_year}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary", use_container_width=True
        )
    if st.button("🔄 Generuj dla innej spółki", use_container_width=True):
        _reset_state()
        st.rerun()

    with st.expander("👁️ Podgląd wygenerowanej treści", expanded=True):
        st.markdown(st.session_state["generated_text"])

# ── STAN: error ───────────────────────────────────────────────────────────────
elif _get_state() == "error":
    st.error(f"❌ Błąd: {st.session_state.get('error_msg', 'Nieznany błąd')}")
    if st.button("🔄 Spróbuj ponownie"):
        _reset_state()
        st.rerun()
