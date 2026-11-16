import os
import sys
from datetime import datetime
from urllib.parse import parse_qsl, urlencode, urlparse, urlunparse
import pandas as pd
from sqlalchemy import create_engine, text

try:
    from dotenv import load_dotenv
except ImportError:
    load_dotenv = None

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
INPUT_FILE = os.path.join(BASE_DIR, "raw_data", "hasil_review_klinik_kita.xlsx")
ENV_FILE = os.path.join(BASE_DIR, ".env")

BULAN_MAP = {
    "Januari": "Jan", "Februari": "Feb", "Maret": "Mar", "April": "Apr",
    "Mei": "May", "Juni": "Jun", "Juli": "Jul", "Agustus": "Aug",
    "September": "Sep", "Oktober": "Oct", "November": "Nov", "Desember": "Dec",
}


def load_database_url() -> str:
    """Membaca DATABASE_URL dari file .env di folder root project."""
    if load_dotenv is not None and os.path.exists(ENV_FILE):
        load_dotenv(ENV_FILE, override=True)
    elif os.path.exists(ENV_FILE):
        with open(ENV_FILE, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                k, v = line.split("=", 1)
                k = k.strip()
                v = v.strip().strip('"').strip("'")
                if k:
                    os.environ[k] = v

    raw_url = (
        os.environ.get("DATABASE_URL")
        or os.environ.get("NEON_DATABASE_URL")
        or ""
    ).strip()

    if not raw_url:
        print(
            "[ERROR] DATABASE_URL tidak ditemukan di file .env.\n"
            f"Pastikan file '{ENV_FILE}' berisi:\n"
            "DATABASE_URL=postgresql://user:password@host/dbname?sslmode=require"
        )
        sys.exit(1)

    if raw_url.startswith("postgres://"):
        raw_url = "postgresql://" + raw_url[len("postgres://"):]

    has_psycopg3 = False
    has_psycopg2 = False
    try:
        import psycopg  # noqa: F401
        has_psycopg3 = True
    except ImportError:
        pass
    try:
        import psycopg2  # noqa: F401
        has_psycopg2 = True
    except ImportError:
        pass

    if not has_psycopg3 and not has_psycopg2:
        print(
            "[ERROR] Driver PostgreSQL belum terpasang di venv.\n"
            'Jalankan: pip install "psycopg[binary]" sqlalchemy'
        )
        sys.exit(1)

    if raw_url.startswith("postgresql+psycopg://") and not has_psycopg3 and has_psycopg2:
        raw_url = raw_url.replace("postgresql+psycopg://", "postgresql+psycopg2://", 1)
    elif raw_url.startswith("postgresql+psycopg2://") and not has_psycopg2 and has_psycopg3:
        raw_url = raw_url.replace("postgresql+psycopg2://", "postgresql+psycopg://", 1)
    elif raw_url.startswith("postgresql://"):
        if has_psycopg3:
            raw_url = raw_url.replace("postgresql://", "postgresql+psycopg://", 1)
        elif has_psycopg2:
            raw_url = raw_url.replace("postgresql://", "postgresql+psycopg2://", 1)

    parsed = urlparse(raw_url)
    query_params = dict(parse_qsl(parsed.query))
    if "sslmode" not in query_params:
        query_params["sslmode"] = "require"
    return urlunparse(parsed._replace(query=urlencode(query_params)))


def clean_nullable_str(val) -> str | None:
    if pd.isna(val):
        return None
    s = str(val).strip()
    return s if s else None


def build_raw_dataset(excel_path: str) -> tuple[pd.DataFrame, datetime]:
    """Membaca file Excel hasil scraper murni (19 kolom) sesuai tipe data asli."""
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"File sumber tidak ditemukan: {excel_path}")

    xls = pd.ExcelFile(excel_path)
    sheet_name = "Google_Reviews" if "Google_Reviews" in xls.sheet_names else xls.sheet_names[0]
    df = pd.read_excel(excel_path, sheet_name=sheet_name)

    fallback_ts = datetime.now()

    if "Waktu_Scraping" in df.columns:
        scrape_dt_series = pd.to_datetime(
            df["Waktu_Scraping"].astype(str).replace(BULAN_MAP, regex=True),
            errors="coerce",
        ).fillna(fallback_ts)
    else:
        scrape_dt_series = pd.Series([fallback_ts] * len(df))

    batch_scrape_ts = scrape_dt_series.iloc[0].to_pydatetime()

    dt_tanggal = pd.to_datetime(
        df["Waktu_Review_Tanggal"].astype(str).replace(BULAN_MAP, regex=True),
        format="%d %b %Y",
        errors="coerce",
    )
    dt_bulan = pd.to_datetime(
        "1 " + df["Waktu_Review_Bulan"].astype(str).replace(BULAN_MAP, regex=True),
        format="%d %b %Y",
        errors="coerce",
    )

    reviews_df = pd.DataFrame(
        {
            "no_urut": df["No"].astype(int),
            "review_id": df["Review_ID"].astype(str).str.strip(),
            "nama_reviewer": df["Nama_Reviewer"].apply(clean_nullable_str),
            "bintang": df["Bintang"].astype(int),
            "waktu_review_tanggal": dt_tanggal.dt.date,
            "waktu_review_bulan": dt_bulan.dt.date,
            "waktu_review_real": df["Waktu_Review_Real"].apply(clean_nullable_str),
            "badge": df["Badge"].apply(clean_nullable_str),
            "isi_ulasan": df["Isi_Ulasan"].apply(clean_nullable_str),
            "tanggapan_pemilik": df["Tanggapan_Pemilik"].apply(clean_nullable_str),
            "sentimen": df["Sentimen"].astype(str).str.strip(),
            "skor_sentimen": df["Skor_Sentimen"].astype(int),
            "skor_teks": df["Skor_Teks"].astype(int),
            "skor_bintang": df["Skor_Bintang"].astype(int),
            "kata_positif": df["Kata_Positif"].apply(clean_nullable_str),
            "kata_negatif": df["Kata_Negatif"].apply(clean_nullable_str),
            "info_reviewer": df["Info_Reviewer"].apply(clean_nullable_str),
            "url_profil": df["URL_Profil"].apply(clean_nullable_str),
            "waktu_scraping": [ts.to_pydatetime() for ts in scrape_dt_series],
        }
    )

    return reviews_df, batch_scrape_ts


def upload_to_neon():
    db_url = load_database_url()
    print(f"Membaca dataset dari: {INPUT_FILE}")
    reviews_df, batch_ts = build_raw_dataset(INPUT_FILE)

    print(
        f"Data siap diunggah (Waktu Scraping: {batch_ts.strftime('%Y-%m-%d %H:%M:%S')}) -> "
        f"{len(reviews_df)} baris, {len(reviews_df.columns)} kolom murni."
    )

    engine = create_engine(db_url, pool_pre_ping=True)

    setup_ddl = [
        "DROP TABLE IF EXISTS klinik_kita_reviews CASCADE;",
        """
        DO $$
        BEGIN
            IF NOT EXISTS (
                SELECT 1
                FROM pg_type t
                JOIN pg_namespace n ON n.oid = t.typnamespace
                WHERE t.typname = 'link' AND n.nspname = 'public'
            ) THEN
                CREATE DOMAIN link AS VARCHAR
                CHECK (VALUE IS NULL OR VALUE ~* '^https?://');
            END IF;
        END$$;
        """,
        """
        CREATE TABLE klinik_kita_reviews (
            no_urut INTEGER NOT NULL,
            review_id VARCHAR PRIMARY KEY,
            nama_reviewer VARCHAR,
            bintang INTEGER NOT NULL,
            waktu_review_tanggal DATE NOT NULL,
            waktu_review_bulan DATE NOT NULL,
            waktu_review_real VARCHAR,
            badge VARCHAR,
            isi_ulasan VARCHAR,
            tanggapan_pemilik VARCHAR,
            sentimen VARCHAR NOT NULL,
            skor_sentimen INTEGER NOT NULL,
            skor_teks INTEGER NOT NULL,
            skor_bintang INTEGER NOT NULL,
            kata_positif VARCHAR,
            kata_negatif VARCHAR,
            info_reviewer VARCHAR,
            url_profil link,
            waktu_scraping TIMESTAMP NOT NULL
        );
        """,
        "CREATE INDEX idx_kk_reviews_tanggal ON klinik_kita_reviews (waktu_review_tanggal);",
        "CREATE INDEX idx_kk_reviews_bulan ON klinik_kita_reviews (waktu_review_bulan);",
        "CREATE INDEX idx_kk_reviews_sentimen ON klinik_kita_reviews (sentimen);",
        "CREATE INDEX idx_kk_reviews_bintang ON klinik_kita_reviews (bintang);",
    ]

    insert_reviews_sql = text(
        """
        INSERT INTO klinik_kita_reviews (
            no_urut, review_id, nama_reviewer, bintang,
            waktu_review_tanggal, waktu_review_bulan, waktu_review_real,
            badge, isi_ulasan, tanggapan_pemilik,
            sentimen, skor_sentimen, skor_teks, skor_bintang,
            kata_positif, kata_negatif, info_reviewer, url_profil, waktu_scraping
        )
        VALUES (
            :no_urut, :review_id, :nama_reviewer, :bintang,
            :waktu_review_tanggal, :waktu_review_bulan, :waktu_review_real,
            :badge, :isi_ulasan, :tanggapan_pemilik,
            :sentimen, :skor_sentimen, :skor_teks, :skor_bintang,
            :kata_positif, :kata_negatif, :info_reviewer, :url_profil, :waktu_scraping
        );
        """
    )

    print("Menghubungkan ke Neon PostgreSQL dan memperbarui tabel klinik_kita_reviews...")
    with engine.begin() as conn:
        for ddl in setup_ddl:
            conn.execute(text(ddl))

        conn.execute(insert_reviews_sql, reviews_df.to_dict(orient="records"))

    print(
        f"\nUpload ke Neon PostgreSQL selesai! (Waktu Scraping: {batch_ts.strftime('%Y-%m-%d %H:%M:%S')})\n"
        f"Tabel klinik_kita_reviews ter-update dengan {len(reviews_df)} baris (waktu_review_real = VARCHAR)."
    )


if __name__ == "__main__":
    upload_to_neon()