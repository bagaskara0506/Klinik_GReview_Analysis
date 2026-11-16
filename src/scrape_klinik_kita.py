import os
import random
import re
import sys
import time
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
from dateutil.relativedelta import relativedelta
import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from playwright.sync_api import Error as PlaywrightError, sync_playwright

WIB = ZoneInfo("Asia/Jakarta")
IS_CI = os.getenv("GITHUB_ACTIONS", "").lower() == "true"

PLACE_FID = "0x2e69ef91b21b2455:0x1876c0007d3a7a25"

# Gunakan nama lengkap klinik + hl=id & gl=id agar Google langsung membuka pop-up ulasan dalam Bahasa Indonesia
TARGET_URL = (
    "https://www.google.com/search?q=Klinik+Kita+1+dan+Klinik+Kita+2+Depok"
    "&hl=id&gl=id&ie=UTF-8"
    f"#lrd={PLACE_FID},1,,,,"
)

FALLBACK_URLS = [
    f"https://www.google.com/search?q=Klinik+Kita+1+dan+Klinik+Kita+2&hl=id&gl=id#lrd={PLACE_FID},1,,,,",
    f"https://www.google.com/search?q=Klinik+Kita+Depok&hl=id&gl=id#lrd={PLACE_FID},1,,,,",
]

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUTPUT_FILE = os.path.join(BASE_DIR, "raw_data", "hasil_review_klinik_kita.xlsx")
PROFILE_DIR = os.path.join(BASE_DIR, "browser_profile")

RAW_POS_FILE = os.path.join(BASE_DIR, "Bank_Ulasan_Negative_Positive", "raw", "positive.txt")
RAW_NEG_FILE = os.path.join(BASE_DIR, "Bank_Ulasan_Negative_Positive", "raw", "negative.txt")
FINAL_POS_FILE = os.path.join(BASE_DIR, "Bank_Ulasan_Negative_Positive", "final", "bank_data_kata_positif.txt")
FINAL_NEG_FILE = os.path.join(BASE_DIR, "Bank_Ulasan_Negative_Positive", "final", "bank_data_kata_negatif.txt")

BULAN_ID = {
    1: "Januari",
    2: "Februari",
    3: "Maret",
    4: "April",
    5: "Mei",
    6: "Juni",
    7: "Juli",
    8: "Agustus",
    9: "September",
    10: "Oktober",
    11: "November",
    12: "Desember",
}

NEGATION_WORDS = {
    "tidak", "tdk", "tak", "bukan", "blm", "belum", "ga", "gak", "gk", "g",
    "enggak", "nggak", "ngga", "kagak", "kurang", "tanpa", "jangan",
    "jgn", "minim", "susah", "sulit", "not", "never", "no",
}

CONTEXT_NEUTRAL_SINGLE_WORDS = {
    "sangat", "banget", "sekali", "terlalu", "cukup", "agak",
    "memutuskan", "banyak", "sedikit", "rendah", "tinggi", "besar", "kecil",
    "yg", "yang", "dan", "di", "ke", "dari", "untuk", "utk", "ini", "itu",
    "nya", "aku", "saya", "kami", "kita", "dia", "mereka", "sudah", "udah",
    "akan", "bisa", "dapat", "ada", "sama", "dengan", "dgn", "pada", "dalam",
    "saat", "waktu", "hari", "klinik", "dokter", "bidan", "perawat",
    "suster", "admin", "pasien",
    "jujur", "oke", "okey", "okeh", "ok", "okay", "paham", "normal",
    "pas", "aman", "jelas", "lengkap", "detail", "baik", "selesai",
    "perhatian", "penting", "harap", "berharap", "semoga", "tolong",
    "mohon", "terima", "kasih", "terimakasih", "makasih", "ya", "yah",
    "iya", "siap", "sesuai", "lumayan", "standar", "biasa", "wajar",
    "mungkin", "seharusnya",
}

NEUTRAL_PHRASES = [
    "mohon perhatian",
    "untuk perhatiannya",
    "atas perhatiannya",
    "terima kasih",
    "terimakasih",
    "sesuai info",
    "sesuai informasi",
    "sesuai prosedur",
    "apakah ada keluhan",
]

EXTRA_POS_PHRASES = [
    "sangat ramah", "sangat membantu", "sangat memuaskan", "sangat informatif",
    "sangat jelas", "sangat detail", "sangat bersih", "sangat nyaman",
    "sangat baik", "sangat bagus", "highly recommended", "sangat recommended",
    "lebih mendetail", "pelayanan bagus", "pelayanan ramah", "pelayanan memuaskan",
    "dokter ramah", "bidan ramah", "suster ramah", "perawat ramah",
    "penjelasan detail", "penjelasannya detail", "tidak sakit", "gak sakit",
    "ga sakit", "tidak antri", "gak antri", "ga lama", "gak lama", "tidak lama",
]

EXTRA_POSITIVE_WORDS = {
    "best", "good", "great", "nice", "helpful", "recommended", "rekomended",
    "mantap", "ramah", "informatif", "komunikatif", "menyenangkan",
    "menenangkan", "membantu", "clean", "friendly", "polite", "excellent",
    "rapi", "wangi", "harum", "dingin", "sejuk", "sabar", "telaten",
    "profesional", "terjangkau", "murah", "solutif", "top", "terbaik",
    "juara", "suka", "senang", "alhamdulillah", "asik", "asyik", "lembut",
    "gentle", "tenang", "disiplin", "rinci", "mendetail",
}

EXTRA_NEG_PHRASES = [
    "sangat kurang", "kurang sekali", "kurang menyenangkan", "tidak menyenangkan",
    "tdk menyenangkan", "gak menyenangkan", "ga menyenangkan", "kurang asik",
    "kurang informatif", "tidak informatif", "tdk informatif", "gak informatif",
    "tidak detail", "tdk detail", "gak detail", "tidak jelas", "tdk jelas",
    "gak jelas", "kurang jelas", "tidak ramah", "kurang ramah", "gak ramah",
    "tidak memuaskan", "kurang memuaskan", "gak memuaskan", "tidak nyaman",
    "kurang nyaman", "gak nyaman", "tidak membantu", "gak membantu",
    "tidak sesuai", "gak sesuai", "tidak ada informasi", "tidak ada penjelasan",
    "gak dingin", "tidak dingin", "g rekomended", "gak rekomended",
    "tidak rekomended", "tidak recommended", "terburu buru", "buru buru",
    "bad attitude", "tidak enak", "gak enak", "tidak sopan", "kurang sopan",
    "salah hitung", "sangat disayangkan", "disayangkan", "tidak akan kembali",
    "gak akan kembali", "ga akan kembali", "tidak pernah", "gak direspon",
    "ga direspon", "tidak direspon", "belum datang", "belum dateng",
    "terlalu lama", "kecewa sekali", "sangat kecewa", "pengalaman buruk",
]

EXTRA_NEGATIVE_WORDS = {
    "kecewa", "mengecewakan", "buruk", "jelek", "parah", "lama", "lelet",
    "lambat", "slow", "judes", "jutek", "galak", "ketus", "kasar", "cuek",
    "kotor", "bau", "apek", "panas", "gerah", "rusak", "mahal", "expensive",
    "disappointed", "uncomfortable", "bad", "keluhan", "komplain", "protes",
    "rugi", "dirugikan", "salah", "keliru", "bingung", "ribet", "antri",
    "antre", "telat", "terlambat", "molor", "ngaret", "males", "malas",
    "menakutkan", "shock", "kaget", "nyeri", "sakit", "terlepas", "lepas",
    "berongga", "cekung", "kejepit", "ketarik", "surcharge", "melanggar",
    "pelit", "singkat", "acuh",
}

STAR_SCORE_MAP = {
    5: 4,
    4: 2,
    3: 0,
    2: -2,
    1: -4,
}

MORE_SUFFIX_PATTERN = re.compile(
    r"(?:\.{1,3}|…)\s*(?:More|Lainnya|Lebih\s+banyak|Selengkapnya)\s*$",
    re.IGNORECASE,
)
UI_ARTIFACT_PATTERN = re.compile(
    r"^(?:Report review|Laporkan ulasan|Laporkan|Flag as inappropriate|Tandai sebagai tidak pantas)$",
    re.IGNORECASE,
)


def parse_review_dates(raw_time: str, ref_date: datetime | None = None) -> tuple[str, str, str]:
    """Mengonversi waktu relatif Google Review menjadi 3 format:
    1. Waktu_Review_Tanggal : 'DD Bulan YYYY' (misal: '13 November 2026')
    2. Waktu_Review_Bulan   : 'Bulan YYYY'    (misal: 'November 2026')
    3. Waktu_Review_Real    : teks asli Google Review (misal: '3 hari lalu')
    """
    if ref_date is None:
        ref_date = datetime.now(WIB).replace(tzinfo=None)

    real_text = (raw_time or "").strip() if isinstance(raw_time, str) else ""
    if not real_text:
        return "", "", ""

    text = re.sub(r"^(?:diedit|edited)\s+", "", real_text.lower()).strip()

    if any(
        k in text
        for k in (
            "baru saja", "just now", "detik", "menit", "jam",
            "hari ini", "today", "second", "minute", "hour",
        )
    ):
        target_dt = ref_date
    elif text in ("kemarin", "yesterday"):
        target_dt = ref_date - timedelta(days=1)
    else:
        match = re.search(
            r"^(se|a|an|\d+)\s*(hari|minggu|bulan|tahun|day|days|week|weeks|month|months|year|years)\b",
            text,
        )
        if not match:
            return real_text, real_text, real_text

        qty_raw = match.group(1)
        qty = 1 if qty_raw in ("se", "a", "an") else int(qty_raw)
        unit = match.group(2)

        if unit in ("hari", "day", "days"):
            target_dt = ref_date - timedelta(days=qty)
        elif unit in ("minggu", "week", "weeks"):
            target_dt = ref_date - timedelta(weeks=qty)
        elif unit in ("bulan", "month", "months"):
            target_dt = ref_date - relativedelta(months=qty)
        elif unit in ("tahun", "year", "years"):
            target_dt = ref_date - relativedelta(years=qty)
        else:
            return real_text, real_text, real_text

    bulan_tahun = f"{BULAN_ID[target_dt.month]} {target_dt.year}"
    tanggal_lengkap = f"{target_dt.day} {bulan_tahun}"
    return tanggal_lengkap, bulan_tahun, real_text


def resolve_latest_lexicon(raw_path: str, final_path: str) -> str:
    """Memilih file kamus sentimen terbaru yang tersedia dan tidak kosong."""
    candidates = [
        p for p in (raw_path, final_path)
        if os.path.exists(p) and os.path.getsize(p) > 0
    ]
    if not candidates:
        return raw_path
    return max(candidates, key=os.path.getmtime)


def load_lexicon(
    filepath: str,
    extra_phrases: list[str] | None = None,
    extra_words: set[str] | None = None,
) -> tuple[list[str], set[str]]:
    """Memuat daftar frasa (multi-kata) dan kata tunggal dari file bank data sentimen."""
    phrases = list(extra_phrases) if extra_phrases else []
    words = set(extra_words) if extra_words else set()

    if not os.path.exists(filepath):
        print(f"Peringatan: File '{filepath}' tidak ditemukan.")
    else:
        with open(filepath, "r", encoding="utf-8") as f:
            for line in f:
                term = line.strip().lower()
                if not term or term.startswith("#"):
                    continue
                if " " in term:
                    phrases.append(term)
                elif term not in CONTEXT_NEUTRAL_SINGLE_WORDS and term not in NEGATION_WORDS:
                    words.add(term)

    phrases = sorted(set(phrases), key=len, reverse=True)
    return phrases, words


def normalize_token_variants(token: str) -> list[str]:
    """Menghasilkan variasi kata untuk menangani huruf yang dipanjangkan (misal: 'besttt' -> 'best')."""
    variants = [token]
    collapsed_triple = re.sub(r"(.)\1{2,}", r"\1", token)
    if collapsed_triple not in variants:
        variants.append(collapsed_triple)
    collapsed_double = re.sub(r"(.)\1+", r"\1", token)
    if collapsed_double not in variants:
        variants.append(collapsed_double)
    return variants


def analyze_sentiment(
    text: str,
    rating: int,
    pos_phrases: list[str],
    pos_words: set[str],
    neg_phrases: list[str],
    neg_words: set[str],
) -> tuple[str, int, int, int, str, str]:
    """Menghitung skor gabungan berbasis teks (dengan deteksi negasi) serta jumlah bintang."""
    rating_int = int(rating or 0)
    star_score = STAR_SCORE_MAP.get(rating_int, 0)

    matched_pos = []
    matched_neg = []
    text_score = 0

    if text:
        clean_text = re.sub(r"[^\w\s]", " ", text.lower())
        clean_text = re.sub(r"\s+", " ", clean_text).strip()
        working_text = clean_text

        for n_phrase in NEUTRAL_PHRASES:
            working_text = re.sub(rf"\b{re.escape(n_phrase)}\b", " ", working_text)

        for phrase in neg_phrases:
            pattern = rf"\b{re.escape(phrase)}\b"
            matches = re.findall(pattern, working_text)
            if matches:
                matched_neg.extend(matches)
                working_text = re.sub(pattern, " ", working_text)

        for phrase in pos_phrases:
            pattern = rf"\b{re.escape(phrase)}\b"
            matches = re.findall(pattern, working_text)
            if matches:
                matched_pos.extend(matches)
                working_text = re.sub(pattern, " ", working_text)

        tokens = working_text.split()
        for idx, token in enumerate(tokens):
            if token in CONTEXT_NEUTRAL_SINGLE_WORDS or token in NEGATION_WORDS:
                continue

            prev_window = tokens[max(0, idx - 2):idx]
            is_negated = any(w in NEGATION_WORDS for w in prev_window)

            for variant in normalize_token_variants(token):
                if variant in CONTEXT_NEUTRAL_SINGLE_WORDS or variant in NEGATION_WORDS:
                    break

                if variant in neg_words and variant not in pos_words:
                    if is_negated:
                        matched_pos.append(f"tidak {variant}")
                    else:
                        matched_neg.append(variant)
                    break

                if variant in pos_words:
                    if is_negated:
                        matched_neg.append(f"tidak/kurang {variant}")
                    else:
                        matched_pos.append(variant)
                    break

        text_score = len(matched_pos) - len(matched_neg)

    total_score = text_score + star_score

    if rating_int == 1 and text_score <= 1:
        label = "Negatif"
    elif rating_int == 2 and text_score <= 0:
        label = "Negatif"
    elif rating_int == 5 and text_score >= -1:
        label = "Positif"
    elif rating_int == 4 and text_score >= 0:
        label = "Positif"
    elif total_score > 0:
        label = "Positif"
    elif total_score < 0:
        label = "Negatif"
    else:
        if rating_int in (1, 2):
            label = "Negatif"
        elif rating_int in (4, 5):
            label = "Positif"
        else:
            label = "Netral"

    return (
        label,
        total_score,
        text_score,
        star_score,
        ", ".join(dict.fromkeys(matched_pos)),
        ", ".join(dict.fromkeys(matched_neg)),
    )


def try_open_reviews_modal(page) -> bool:
    """Mencoba memicu pembukaan pop-up ulasan jika Google tidak membukanya secara otomatis."""
    try:
        return page.evaluate(
            r"""(fid) => {
                const selectors = [
                    `a[data-fid="${fid}"]`,
                    `[data-fid="${fid}"] a`,
                    'a[href*="#lrd="]',
                    '[data-async-trigger="reviewDialog"]',
                ];
                for (const sel of selectors) {
                    const el = document.querySelector(sel);
                    if (el) {
                        el.click();
                        return true;
                    }
                }

                const links = Array.from(document.querySelectorAll("a, button, span[role='button']"));
                for (const el of links) {
                    const txt = (el.textContent || "").trim();
                    if (/^\(?\d[\d.,]*\)?\s*(?:ulasan\s+Google|Google\s+reviews|ulasan|reviews)?$/i.test(txt) && /ulasan|review/i.test(txt)) {
                        el.click();
                        return true;
                    }
                }

                const targetHash = `#lrd=${fid},1,,,,`;
                if (window.location.hash !== targetHash) {
                    window.location.hash = targetHash;
                } else {
                    window.location.hash = "";
                    setTimeout(() => {
                        window.location.hash = targetHash;
                    }, 150);
                }
                return false;
            }""",
            PLACE_FID,
        )
    except PlaywrightError:
        return False


def wait_for_reviews_dialog(page, ci_timeout_sec: int = 90):
    """Menunggu pop-up ulasan muncul dengan auto-fallback anti-stuck & batas waktu mode GitHub Actions."""
    captcha_notified = False
    tried_trigger = False
    fallback_idx = 0
    start_wait_time = time.time()
    last_action_time = start_wait_time

    while True:
        if page.is_closed():
            print("\nJendela browser ditutup. Proses scraping dibatalkan.")
            sys.exit(0)

        if IS_CI and (time.time() - start_wait_time) > ci_timeout_sec:
            raise TimeoutError(
                f"Pop-up ulasan Google tidak muncul dalam {ci_timeout_sec} detik di runner GitHub Actions "
                "(kemungkinan terblokir verifikasi/reCAPTCHA dari IP cloud)."
            )

        try:
            is_sorry_page = "google.com/sorry" in page.url
            has_recaptcha = page.locator("iframe[src*='recaptcha']").count() > 0

            if is_sorry_page or has_recaptcha:
                if IS_CI:
                    raise RuntimeError(
                        "Google menampilkan halaman verifikasi/reCAPTCHA di runner GitHub Actions. "
                        "Jalankan scraping secara lokal atau ulangi workflow beberapa saat lagi."
                    )
                if not captcha_notified:
                    print(
                        "\n[VERIFIKASI DIBUTUHKAN] Google menampilkan halaman reCAPTCHA ('I'm not a robot').\n"
                        "Silakan centang dan selesaikan verifikasi langsung di jendela browser.\n"
                        "Script akan otomatis melanjutkan setelah halaman ulasan muncul...\n"
                    )
                    captcha_notified = True
                last_action_time = time.time()
                time.sleep(2)
                continue

            dialog = page.locator("div.QcQKfe")
            if dialog.count() > 0 and dialog.first.is_visible():
                print("Pop-up ulasan berhasil terbuka dalam Bahasa Indonesia.")
                return dialog.first

            elapsed = time.time() - last_action_time

            if elapsed >= 5 and not tried_trigger:
                print("Pop-up belum terbuka otomatis, mencoba memicu modal ulasan...")
                try_open_reviews_modal(page)
                tried_trigger = True
                last_action_time = time.time()
            elif elapsed >= 6 and fallback_idx < len(FALLBACK_URLS):
                next_url = FALLBACK_URLS[fallback_idx]
                fallback_idx += 1
                print(f"Membuka URL cadangan ({fallback_idx}/{len(FALLBACK_URLS)}) dalam Bahasa Indonesia...")
                page.goto(next_url, wait_until="domcontentloaded")
                tried_trigger = False
                last_action_time = time.time()

        except PlaywrightError:
            if page.is_closed():
                print("\nJendela browser ditutup. Proses scraping dibatalkan.")
                sys.exit(0)

        time.sleep(1)


def read_card_snapshot(card) -> dict:
    """Membaca seluruh field dari satu card ulasan serta status apakah tombol 'Lebih banyak' masih ada."""
    return card.evaluate(
        r"""(card) => {
            const uiNoiseRegex = /^(?:See translation.*|Translated by Google|See original.*|Lihat terjemahan.*|Diterjemahkan oleh Google|Lihat aslinya.*|Like|Share|Suka|Bagikan|New|Baru|More|Lainnya|Lebih\s+banyak|Selengkapnya|Less|Lebih\s+sedikit|Lebih\s+ringkas|Response from the owner|Tanggapan dari pemilik|Report review|Laporkan ulasan|Laporkan|Flag as inappropriate|Tandai sebagai tidak pantas|·|\d+)$/i;
            const expandBtnRegex = /^(?:\.{1,3}|…)?\s*(?:More|Lainnya|Lebih\s+banyak|Selengkapnya)$/i;
            const truncSuffixRegex = /(?:\.{1,3}|…)\s*(?:More|Lainnya|Lebih\s+banyak|Selengkapnya)\s*$/i;

            function isVisible(el) {
                if (!el) return false;
                const style = window.getComputedStyle(el);
                if (style.display === "none" || style.visibility === "hidden") return false;
                return el.getClientRects().length > 0;
            }

            function cleanText(str) {
                let cleaned = (str || "")
                    .replace(/\s*(?:\.\.)?\s*(?:See translation|Translated by Google|Lihat terjemahan|Diterjemahkan oleh Google).*$/is, "")
                    .replace(/\s*(?:Less|Lebih\s+sedikit|Lebih\s+ringkas)\s*$/i, "")
                    .trim();
                return uiNoiseRegex.test(cleaned) ? "" : cleaned;
            }

            function pickBestCandidate(items) {
                const valid = items
                    .map(it => ({ text: cleanText(it.text), visible: Boolean(it.visible) }))
                    .filter(it => Boolean(it.text));
                if (valid.length === 0) return "";

                const isTrunc = (t) => truncSuffixRegex.test(t);
                const visibleUntrunc = valid.filter(it => it.visible && !isTrunc(it.text));
                if (visibleUntrunc.length > 0) {
                    return visibleUntrunc.reduce((a, b) => (b.text.length > a.text.length ? b : a)).text;
                }

                const anyUntrunc = valid.filter(it => !isTrunc(it.text));
                if (anyUntrunc.length > 0) {
                    return anyUntrunc.reduce((a, b) => (b.text.length > a.text.length ? b : a)).text;
                }

                const visibleAny = valid.filter(it => it.visible);
                const pool = visibleAny.length > 0 ? visibleAny : valid;
                return pool.reduce((a, b) => (b.text.length > a.text.length ? b : a)).text;
            }

            function extractFromContainer(containerEl) {
                if (!containerEl) return "";
                const items = [];
                const selectors = ["div.gbdaZe div.d83Iyc", "div.d83Iyc", "div.gbdaZe", "div.EJ8fFf"];
                for (const sel of selectors) {
                    containerEl.querySelectorAll(sel).forEach(el => {
                        items.push({ text: el.textContent, visible: isVisible(el) });
                    });
                }
                if (items.length === 0) {
                    items.push({ text: containerEl.textContent, visible: isVisible(containerEl) });
                }
                return pickBestCandidate(items);
            }

            function inZone(el, zone, ownerHeader) {
                if (!ownerHeader) return zone === "reviewer";
                const isAfterOwner = Boolean(
                    ownerHeader.compareDocumentPosition(el) & Node.DOCUMENT_POSITION_FOLLOWING
                );
                return zone === "owner" ? isAfterOwner : !isAfterOwner;
            }

            function hasVisibleExpandButton(zone, ownerHeader) {
                const triggers = Array.from(
                    card.querySelectorAll('span.w8nwRe, a.MtCSLb, [role="button"], button, a, span')
                ).filter(el => inZone(el, zone, ownerHeader) && isVisible(el));

                return triggers.some(el => {
                    const txt = (el.textContent || "").trim();
                    return expandBtnRegex.test(txt);
                });
            }

            function extractZoneText(zone, ownerHeader) {
                const mjwikbEls = Array.from(card.querySelectorAll("div.Mjwikb"))
                    .filter(el => inZone(el, zone, ownerHeader));
                if (mjwikbEls.length > 0) {
                    const visibleMj = mjwikbEls.filter(isVisible);
                    const targetContainer = visibleMj.length > 0 ? visibleMj[0] : mjwikbEls[0];
                    const txt = extractFromContainer(targetContainer);
                    if (txt) return txt;
                }

                const d83Els = Array.from(card.querySelectorAll("div.d83Iyc"))
                    .filter(el => inZone(el, zone, ownerHeader));
                if (d83Els.length > 0) {
                    const txt = pickBestCandidate(
                        d83Els.map(el => ({ text: el.textContent, visible: isVisible(el) }))
                    );
                    if (txt) return txt;
                }

                const expandEls = Array.from(card.querySelectorAll("div.EuXh6e, div.gbdaZe, div.EJ8fFf"))
                    .filter(el => inZone(el, zone, ownerHeader));
                if (expandEls.length > 0) {
                    const txt = pickBestCandidate(
                        expandEls.map(el => ({ text: el.textContent, visible: isVisible(el) }))
                    );
                    if (txt) return txt;
                }

                const excludeSelector = [
                    "a.yC3ZMb", "div.rhtdWc", "div.YE6B9b", "div.h3PQJ",
                    "span.m6Mr5d", "span.M2wKwd", "div.svzjne",
                    "button", "[role='button']", "[role='radiogroup']",
                    "[role='menu']", "[role='menuitem']", "[aria-haspopup='menu']",
                    "[aria-label*='Report' i]", "[aria-label*='Laporkan' i]",
                    "svg", "img"
                ].join(", ");

                const leafTexts = [];
                const allNodes = Array.from(card.querySelectorAll("div, span"))
                    .filter(el => inZone(el, zone, ownerHeader) && isVisible(el));

                for (const el of allNodes) {
                    if (el.closest(excludeSelector)) continue;
                    const hasElementChildren = Array.from(el.children).some(ch => ch.tagName !== "BR");
                    if (hasElementChildren) continue;

                    const txt = cleanText(el.textContent);
                    if (!txt) continue;
                    if (!leafTexts.includes(txt)) {
                        leafTexts.push(txt);
                    }
                }

                return leafTexts.join(" ").trim();
            }

            const reviewId = card.getAttribute("data-id") || "";
            const authorHeaders = Array.from(card.querySelectorAll("div.rhtdWc"));
            const reviewerName = authorHeaders.length > 0
                ? (authorHeaders[0].textContent || "").trim()
                : "";

            let ownerHeader = null;
            if (authorHeaders.length > 1) {
                const secondAuthor = (authorHeaders[1].textContent || "").trim().toLowerCase();
                if (secondAuthor.includes("owner") || secondAuthor.includes("pemilik")) {
                    ownerHeader = authorHeaders[1];
                }
            }

            const reviewText = extractZoneText("reviewer", ownerHeader);
            const ownerResponse = extractZoneText("owner", ownerHeader);
            const hasMoreReview = hasVisibleExpandButton("reviewer", ownerHeader);
            const hasMoreOwner = hasVisibleExpandButton("owner", ownerHeader);

            const rating = card.querySelectorAll('div.h3PQJ svg path[fill="#fabb05"]').length;

            const timeEls = card.querySelectorAll("span.m6Mr5d");
            const reviewTime = timeEls.length > 0
                ? (timeEls[0].textContent || "").trim()
                : "";

            const badge = card.querySelector("span.M2wKwd") ? "Baru" : "";

            const statsEls = card.querySelectorAll("div.YE6B9b");
            const reviewerStats = statsEls.length > 0
                ? (statsEls[0].textContent || "").trim()
                : "";

            const profileLink = card.querySelector("a.yC3ZMb");
            const profileUrl = profileLink
                ? (profileLink.getAttribute("href") || "")
                : "";

            return {
                review_id: reviewId,
                reviewer_name: reviewerName,
                rating: rating,
                review_time: reviewTime,
                badge: badge,
                review_text: reviewText,
                owner_response: ownerResponse,
                has_more_review: hasMoreReview,
                has_more_owner: hasMoreOwner,
                reviewer_stats: reviewerStats,
                profile_url: profileUrl,
            };
        }"""
    )


def sanitize_extracted_text(text: str) -> str:
    """Membersihkan sisa label tombol UI dan blok terjemahan otomatis Google."""
    cleaned = re.sub(
        r"\s*(?:\.\.)?\s*(?:See translation|Translated by Google|Lihat terjemahan|Diterjemahkan oleh Google).*$",
        "",
        text or "",
        flags=re.IGNORECASE | re.DOTALL,
    )
    cleaned = re.sub(
        r"\s*(?:\.{1,3}|…)?\s*(?:More|Lainnya|Lebih\s+banyak|Selengkapnya|Less|Lebih\s+sedikit|Lebih\s+ringkas)\s*$",
        "",
        cleaned,
        flags=re.IGNORECASE,
    ).strip()
    if UI_ARTIFACT_PATTERN.match(cleaned):
        return ""
    return cleaned


def switch_card_to_original_language(card):
    """Memastikan teks tampil dalam bahasa aslinya jika ada tombol 'Lihat aslinya' / 'See original'."""
    has_see_original = card.evaluate(
        r"""(card) => /See original|Lihat aslinya/i.test(card.textContent || "")"""
    )
    if not has_see_original:
        return

    try:
        card.scroll_into_view_if_needed(timeout=2000)
        time.sleep(0.1)
    except Exception:
        pass

    clicked_count = card.evaluate(
        r"""(card) => {
            function isVisible(el) {
                const style = window.getComputedStyle(el);
                return style.display !== "none" && style.visibility !== "hidden" && el.getClientRects().length > 0;
            }

            const candidates = Array.from(
                card.querySelectorAll('span, a, button, [role="button"]')
            ).filter(el => {
                if (!isVisible(el)) return false;
                const txt = (el.textContent || "").trim();
                return /^(?:See original|Lihat aslinya)(?:\s*\([^)]+\))?$/i.test(txt);
            });

            const distinctTriggers = candidates.filter(
                el => !candidates.some(other => other !== el && other.contains(el))
            );

            distinctTriggers.forEach(el => el.click());
            return distinctTriggers.length;
        }"""
    )

    if clicked_count > 0:
        for _ in range(10):
            time.sleep(0.15)
            still_translated = card.evaluate(
                r"""(card) => {
                    const els = Array.from(card.querySelectorAll('span, a, button, [role="button"]'));
                    return els.some(el => {
                        const style = window.getComputedStyle(el);
                        if (style.display === "none" || style.visibility === "hidden") return false;
                        if (el.getClientRects().length === 0) return false;
                        const txt = (el.textContent || "").trim();
                        return /^(?:See original|Lihat aslinya)(?:\s*\([^)]+\))?$/i.test(txt);
                    });
                }"""
            )
            if not still_translated:
                break


def expand_and_read_card(card) -> dict:
    """Membaca ulasan dalam bahasa asli Indonesia, serta membuka 'Lebih banyak'/'Lainnya' jika terpotong."""
    switch_card_to_original_language(card)

    initial_data = read_card_snapshot(card)

    needs_review_expand = initial_data["has_more_review"] or bool(
        MORE_SUFFIX_PATTERN.search(initial_data["review_text"])
    )
    needs_owner_expand = initial_data["has_more_owner"] or bool(
        MORE_SUFFIX_PATTERN.search(initial_data["owner_response"])
    )

    if not needs_review_expand and not needs_owner_expand:
        initial_data["review_text"] = sanitize_extracted_text(initial_data["review_text"])
        initial_data["owner_response"] = sanitize_extracted_text(initial_data["owner_response"])
        return initial_data

    try:
        card.scroll_into_view_if_needed(timeout=2000)
        time.sleep(0.15)
    except Exception:
        pass

    card.evaluate(
        r"""(card, opts) => {
            const expandBtnRegex = /^(?:\.{1,3}|…)?\s*(?:More|Lainnya|Lebih\s+banyak|Selengkapnya)$/i;
            const authorHeaders = Array.from(card.querySelectorAll("div.rhtdWc"));
            let ownerHeader = null;
            if (authorHeaders.length > 1) {
                const secondAuthor = (authorHeaders[1].textContent || "").trim().toLowerCase();
                if (secondAuthor.includes("owner") || secondAuthor.includes("pemilik")) {
                    ownerHeader = authorHeaders[1];
                }
            }

            function isVisible(el) {
                const style = window.getComputedStyle(el);
                return style.display !== "none" && style.visibility !== "hidden" && el.getClientRects().length > 0;
            }

            function inZone(el, zone) {
                if (!ownerHeader) return zone === "reviewer";
                const isAfterOwner = Boolean(
                    ownerHeader.compareDocumentPosition(el) & Node.DOCUMENT_POSITION_FOLLOWING
                );
                return zone === "owner" ? isAfterOwner : !isAfterOwner;
            }

            function clickMoreInZone(zone) {
                const candidates = Array.from(
                    card.querySelectorAll('span.w8nwRe, a.MtCSLb, [role="button"], button, a, span')
                ).filter(el => inZone(el, zone) && isVisible(el));

                const matches = candidates.filter(el => {
                    const txt = (el.textContent || "").trim();
                    return expandBtnRegex.test(txt);
                });

                const distinct = matches.filter(
                    el => !matches.some(other => other !== el && other.contains(el))
                );
                if (distinct.length > 0) {
                    distinct[0].click();
                }
            }

            if (opts.expandReview) clickMoreInZone("reviewer");
            if (opts.expandOwner) clickMoreInZone("owner");
        }""",
        {"expandReview": needs_review_expand, "expandOwner": needs_owner_expand},
    )

    best_data = dict(initial_data)
    for _ in range(10):
        time.sleep(0.2)
        updated = read_card_snapshot(card)

        if len(updated["review_text"]) >= len(best_data["review_text"]):
            best_data["review_text"] = updated["review_text"]
        if len(updated["owner_response"]) >= len(best_data["owner_response"]):
            best_data["owner_response"] = updated["owner_response"]

        still_truncated_review = updated["has_more_review"] or bool(
            MORE_SUFFIX_PATTERN.search(best_data["review_text"])
        )
        still_truncated_owner = updated["has_more_owner"] or bool(
            MORE_SUFFIX_PATTERN.search(best_data["owner_response"])
        )
        if not still_truncated_review and not still_truncated_owner:
            break

    switch_card_to_original_language(card)
    final_check = read_card_snapshot(card)
    if final_check["review_text"] and not MORE_SUFFIX_PATTERN.search(final_check["review_text"]):
        if len(final_check["review_text"]) >= len(sanitize_extracted_text(best_data["review_text"])):
            best_data["review_text"] = final_check["review_text"]
    if final_check["owner_response"] and not MORE_SUFFIX_PATTERN.search(final_check["owner_response"]):
        if len(final_check["owner_response"]) >= len(sanitize_extracted_text(best_data["owner_response"])):
            best_data["owner_response"] = final_check["owner_response"]

    best_data["review_text"] = sanitize_extracted_text(best_data["review_text"])
    best_data["owner_response"] = sanitize_extracted_text(best_data["owner_response"])

    return best_data


def cleanup_old_excel_files(output_path: str):
    """Menghapus file Excel lama (termasuk file cadangan lama) sebelum menyimpan hasil scraping baru."""
    output_dir = os.path.dirname(output_path)
    if not os.path.exists(output_dir):
        return

    base_prefix = os.path.splitext(os.path.basename(output_path))[0]
    for fname in os.listdir(output_dir):
        if fname.startswith(base_prefix) and fname.endswith(".xlsx"):
            old_file = os.path.join(output_dir, fname)
            try:
                os.remove(old_file)
                print(f"File Excel lama dihapus: '{old_file}'")
            except PermissionError:
                print(
                    f"\n[ERROR] File lama '{old_file}' sedang dibuka di Excel.\n"
                    f"Tutup file tersebut terlebih dahulu lalu jalankan ulang script."
                )
                sys.exit(1)


def save_formatted_excel(df: pd.DataFrame, output_path: str) -> str:
    """Menghapus file Excel lama dan menyimpan hasil scraping terbaru dengan format rapi."""
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    cleanup_old_excel_files(output_path)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Google_Reviews")
        ws = writer.sheets["Google_Reviews"]

        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)

        column_widths = {
            "A": 6,   # No
            "B": 24,  # Review_ID
            "C": 24,  # Nama_Reviewer
            "D": 10,  # Bintang
            "E": 22,  # Waktu_Review_Tanggal
            "F": 20,  # Waktu_Review_Bulan
            "G": 20,  # Waktu_Review_Real
            "H": 10,  # Badge
            "I": 65,  # Isi_Ulasan
            "J": 50,  # Tanggapan_Pemilik
            "K": 12,  # Sentimen
            "L": 15,  # Skor_Sentimen
            "M": 12,  # Skor_Teks
            "N": 14,  # Skor_Bintang
            "O": 28,  # Kata_Positif
            "P": 28,  # Kata_Negatif
            "Q": 24,  # Info_Reviewer
            "R": 40,  # URL_Profil
            "S": 22,  # Waktu_Scraping
        }

        for col_letter, width in column_widths.items():
            ws.column_dimensions[col_letter].width = width

        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            for cell in row:
                cell.alignment = Alignment(vertical="top", wrap_text=True)

        ws.freeze_panes = "A2"

    return output_path


def scrape_all_google_reviews():
    pos_path = resolve_latest_lexicon(RAW_POS_FILE, FINAL_POS_FILE)
    neg_path = resolve_latest_lexicon(RAW_NEG_FILE, FINAL_NEG_FILE)
    print(f"Memuat kamus positif dari : {pos_path}")
    print(f"Memuat kamus negatif dari : {neg_path}")

    pos_phrases, pos_words = load_lexicon(
        pos_path, extra_phrases=EXTRA_POS_PHRASES, extra_words=EXTRA_POSITIVE_WORDS
    )
    neg_phrases, neg_words = load_lexicon(
        neg_path, extra_phrases=EXTRA_NEG_PHRASES, extra_words=EXTRA_NEGATIVE_WORDS
    )
    print(f"Total aturan positif: {len(pos_phrases)} frasa, {len(pos_words)} kata")
    print(f"Total aturan negatif: {len(neg_phrases)} frasa, {len(neg_words)} kata")

    os.makedirs(PROFILE_DIR, exist_ok=True)

    with sync_playwright() as p:
        launch_args = [
            "--disable-blink-features=AutomationControlled",
            "--start-maximized",
            "--lang=id",
        ]
        if IS_CI:
            launch_args.extend(["--no-sandbox", "--disable-dev-shm-usage"])

        launch_kwargs = dict(
            user_data_dir=PROFILE_DIR,
            headless=False,
            no_viewport=True,
            locale="id-ID",
            timezone_id="Asia/Jakarta",
            extra_http_headers={"Accept-Language": "id-ID,id;q=0.9"},
            args=launch_args,
        )

        if IS_CI:
            context = p.chromium.launch_persistent_context(**launch_kwargs)
        else:
            try:
                context = p.chromium.launch_persistent_context(channel="chrome", **launch_kwargs)
            except Exception:
                context = p.chromium.launch_persistent_context(**launch_kwargs)

        context.add_init_script(
            "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
        )

        page = context.pages[0] if context.pages else context.new_page()
        page.set_default_timeout(60000)

        print("Membuka halaman Google Review Klinik Kita (Bahasa Indonesia - WIB)...")
        page.goto(TARGET_URL, wait_until="domcontentloaded")

        dialog = wait_for_reviews_dialog(page)

        print("Mengurutkan ulasan berdasarkan 'Terbaru'...")
        newest_btn = dialog.locator('div[role="radiogroup"] div[role="radio"][data-sort="2"]')
        newest_btn.wait_for(state="visible")
        time.sleep(random.uniform(0.8, 1.5))
        newest_btn.click()

        page.wait_for_selector(
            'div[role="radiogroup"] div[role="radio"][data-sort="2"][aria-checked="true"]'
        )
        time.sleep(random.uniform(2.5, 3.5))

        print("Memulai auto-scroll tanpa batas hingga seluruh ulasan selesai dimuat...")
        previous_count = 0
        stagnant_rounds = 0
        max_stagnant_rounds = 8

        while True:
            cards = dialog.locator('div.bwb7ce[jsname="ShBeI"]')
            current_count = cards.count()
            print(f"Total ulasan dimuat sementara: {current_count}")

            if current_count == previous_count:
                stagnant_rounds += 1
                print(f"Tidak ada penambahan ulasan baru (percobaan {stagnant_rounds}/{max_stagnant_rounds})...")
                if stagnant_rounds >= max_stagnant_rounds:
                    print("Seluruh ulasan sudah habis dimuat sampai bagian paling bawah.")
                    break
            else:
                stagnant_rounds = 0
                previous_count = current_count

            scroll_body = dialog.locator("div.review-dialog-body")
            if scroll_body.count() > 0:
                scroll_body.first.evaluate("el => { el.scrollTop = el.scrollHeight; }")

            if current_count > 0:
                last_card = cards.nth(current_count - 1)
                last_card.scroll_into_view_if_needed()
                box = last_card.bounding_box()
                if box:
                    page.mouse.move(box["x"] + box["width"] / 2, box["y"] + box["height"] / 2)
                    page.mouse.wheel(0, random.randint(1200, 1800))

            time.sleep(random.uniform(2.2, 3.6))

        cards = dialog.locator('div.bwb7ce[jsname="ShBeI"]')
        total_cards = cards.count()
        print(f"Mengekstrak lengkap dan menganalisis sentimen dari {total_cards} card ulasan...")

        results = []
        seen_ids = set()
        scrape_timestamp = datetime.now(WIB).replace(tzinfo=None)
        scrape_timestamp_str = scrape_timestamp.strftime("%Y-%m-%d %H:%M:%S")

        for i in range(total_cards):
            card = cards.nth(i)
            item = expand_and_read_card(card)

            review_id = item["review_id"] or f"row_{i}"
            if review_id in seen_ids:
                continue
            seen_ids.add(review_id)

            review_text = item["review_text"]
            rating = item["rating"]
            review_tanggal, review_bulan, review_real = parse_review_dates(
                item["review_time"], scrape_timestamp
            )

            (
                sentiment_label,
                total_score,
                text_score,
                star_score,
                pos_found,
                neg_found,
            ) = analyze_sentiment(
                review_text, rating, pos_phrases, pos_words, neg_phrases, neg_words
            )

            results.append(
                {
                    "No": len(results) + 1,
                    "Review_ID": review_id,
                    "Nama_Reviewer": item["reviewer_name"],
                    "Bintang": rating,
                    "Waktu_Review_Tanggal": review_tanggal,
                    "Waktu_Review_Bulan": review_bulan,
                    "Waktu_Review_Real": review_real,
                    "Badge": item["badge"],
                    "Isi_Ulasan": review_text,
                    "Tanggapan_Pemilik": item["owner_response"],
                    "Sentimen": sentiment_label,
                    "Skor_Sentimen": total_score,
                    "Skor_Teks": text_score,
                    "Skor_Bintang": star_score,
                    "Kata_Positif": pos_found,
                    "Kata_Negatif": neg_found,
                    "Info_Reviewer": item["reviewer_stats"],
                    "URL_Profil": item["profile_url"],
                    "Waktu_Scraping": scrape_timestamp_str,
                }
            )

            if (i + 1) % 50 == 0 or (i + 1) == total_cards:
                print(f"Progress ekstraksi: {i + 1}/{total_cards} ulasan diproses...")

        context.close()

        saved_path = save_formatted_excel(pd.DataFrame(results), OUTPUT_FILE)
        print(
            f"Selesai! {len(results)} ulasan lengkap (waktu penarikan: {scrape_timestamp_str} WIB) "
            f"berhasil disimpan ke '{saved_path}'."
        )


if __name__ == "__main__":
    scrape_all_google_reviews()