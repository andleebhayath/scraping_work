"""
Create core tables (if missing) and load from raw_institutions_officers_social_raw.

1) organization              — one row per institution (iau_id)
2) person                    — one row per distinct officer name
3) organization_person       — link table: which person belongs to which organization (M2M)
4) organization_officer      — one row per raw row: designation + URLs + search metadata

Each run: if DROP_REBUILD is True, those four tables are dropped and recreated, then loaded
from raw again (no duplicate rows). Raw table is never touched.

Edit settings below, then run:
  python scripts/load_core_organization_from_raw_mysql.py
"""

from __future__ import annotations

import mysql.connector

# --- edit ---
HOST = "localhost"
USER = "root"
PASSWORD = "root"
DATABASE = "organization_data"

RAW = "raw_institutions_officers_social_raw"
ORG = "organization"
PERSON = "person"
ORG_PERSON = "organization_person"  # link: organization <-> person (many-to-many)
LINK = "organization_officer"  # designation + URLs per raw row (still keyed by org + person)

# True = drop the four tables above and reload (clean, no duplicates).
# False = keep tables; inserts use upserts only (OK if keys never change).
DROP_REBUILD = True
# When False, tables are kept and CREATE uses IF NOT EXISTS (upserts still avoid dup keys).
IFNE = "IF NOT EXISTS " if not DROP_REBUILD else ""
# ---


def drop_core_tables(cur) -> None:
    """Drop derived tables (children first). Raw table is never dropped."""
    cur.execute("SET FOREIGN_KEY_CHECKS = 0")
    for tbl in (LINK, ORG_PERSON, ORG, PERSON):
        cur.execute(f"DROP TABLE IF EXISTS `{tbl}`")
    cur.execute("SET FOREIGN_KEY_CHECKS = 1")
    print("Dropped (if existed):", ", ".join((LINK, ORG_PERSON, ORG, PERSON)))


CREATE_ORG = f"""
CREATE TABLE {IFNE}`{ORG}` (
    organization_id      BIGINT AUTO_INCREMENT PRIMARY KEY,
    source_system        VARCHAR(100) NOT NULL,
    source_org_id        VARCHAR(100) NULL,
    org_name             VARCHAR(500) NOT NULL,
    org_type             VARCHAR(50) NOT NULL,
    official_website     VARCHAR(1000) NULL,
    country              VARCHAR(100) NULL,
    city                 VARCHAR(200) NULL,
    is_active            TINYINT(1) NOT NULL DEFAULT 1,
    created_at_utc       DATETIME(3) NOT NULL DEFAULT CURRENT_TIMESTAMP(3),
    updated_at_utc       DATETIME(3) NOT NULL DEFAULT CURRENT_TIMESTAMP(3) ON UPDATE CURRENT_TIMESTAMP(3),
    UNIQUE KEY ux_org_source_id (source_system, source_org_id)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci
"""

CREATE_PERSON = f"""
CREATE TABLE {IFNE}`{PERSON}` (
    person_id            BIGINT AUTO_INCREMENT PRIMARY KEY,
    full_name            VARCHAR(300) NOT NULL,
    created_at_utc       DATETIME(3) NOT NULL DEFAULT CURRENT_TIMESTAMP(3),
    UNIQUE KEY ux_person_full_name (full_name)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci
"""

CREATE_ORG_PERSON = f"""
CREATE TABLE {IFNE}`{ORG_PERSON}` (
    organization_person_id BIGINT AUTO_INCREMENT PRIMARY KEY,
    organization_id        BIGINT NOT NULL,
    person_id              BIGINT NOT NULL,
    created_at_utc         DATETIME(3) NOT NULL DEFAULT CURRENT_TIMESTAMP(3),
    UNIQUE KEY ux_org_person (organization_id, person_id),
    KEY ix_op_person (person_id),
    CONSTRAINT fk_op_org FOREIGN KEY (organization_id)
        REFERENCES `{ORG}`(organization_id) ON DELETE CASCADE ON UPDATE CASCADE,
    CONSTRAINT fk_op_person FOREIGN KEY (person_id)
        REFERENCES `{PERSON}`(person_id) ON DELETE CASCADE ON UPDATE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci
"""

CREATE_LINK = f"""
CREATE TABLE {IFNE}`{LINK}` (
    link_id              BIGINT AUTO_INCREMENT PRIMARY KEY,
    raw_id               BIGINT NOT NULL,
    organization_id      BIGINT NOT NULL,
    person_id            BIGINT NOT NULL,
    officer_role         VARCHAR(200) NULL,
    job_title            VARCHAR(300) NULL,
    linkedin_url         VARCHAR(2000) NULL,
    other_social_urls    LONGTEXT NULL,
    university_page_url  VARCHAR(2000) NULL,
    top_urls             LONGTEXT NULL,
    search_engine        VARCHAR(100) NULL,
    search_query         VARCHAR(1000) NULL,
    error_text           VARCHAR(2000) NULL,
    created_at_utc       DATETIME(3) NOT NULL DEFAULT CURRENT_TIMESTAMP(3),
    UNIQUE KEY ux_officer_raw_id (raw_id),
    KEY ix_officer_org (organization_id),
    KEY ix_officer_person (person_id),
    CONSTRAINT fk_officer_org FOREIGN KEY (organization_id)
        REFERENCES `{ORG}`(organization_id) ON DELETE CASCADE ON UPDATE CASCADE,
    CONSTRAINT fk_officer_person FOREIGN KEY (person_id)
        REFERENCES `{PERSON}`(person_id) ON DELETE CASCADE ON UPDATE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci
"""

LOAD_ORG = f"""
INSERT INTO `{ORG}` (
    source_system,
    source_org_id,
    org_name,
    org_type,
    official_website
)
SELECT
    'WHED',
    NULLIF(TRIM(iau_id), ''),
    TRIM(university_name),
    'university',
    MAX(NULLIF(TRIM(official_www), ''))
FROM `{RAW}`
WHERE TRIM(COALESCE(university_name, '')) <> ''
  AND TRIM(COALESCE(iau_id, '')) <> ''
GROUP BY
    NULLIF(TRIM(iau_id), ''),
    TRIM(university_name)
ON DUPLICATE KEY UPDATE
    org_name = VALUES(org_name),
    org_type = VALUES(org_type),
    official_website = VALUES(official_website),
    updated_at_utc = CURRENT_TIMESTAMP(3)
"""

LOAD_PERSON = f"""
INSERT INTO `{PERSON}` (full_name)
SELECT DISTINCT TRIM(officer_name)
FROM `{RAW}`
WHERE TRIM(COALESCE(officer_name, '')) <> ''
ON DUPLICATE KEY UPDATE full_name = VALUES(full_name)
"""

# Distinct (organization, person) pairs from raw — the actual M2M link.
LOAD_ORG_PERSON = f"""
INSERT INTO `{ORG_PERSON}` (organization_id, person_id)
SELECT DISTINCT
    o.organization_id,
    p.person_id
FROM `{RAW}` AS r
INNER JOIN `{ORG}` AS o
    ON o.source_system = 'WHED'
   AND o.source_org_id = NULLIF(TRIM(r.iau_id), '')
INNER JOIN `{PERSON}` AS p
    ON p.full_name = TRIM(r.officer_name)
WHERE TRIM(COALESCE(r.officer_name, '')) <> ''
  AND TRIM(COALESCE(r.iau_id, '')) <> ''
ON DUPLICATE KEY UPDATE
    organization_id = VALUES(organization_id)
"""

# One row per raw_id: joins org (WHED + iau_id) and person (officer_name).
LOAD_LINK = f"""
INSERT INTO `{LINK}` (
    raw_id,
    organization_id,
    person_id,
    officer_role,
    job_title,
    linkedin_url,
    other_social_urls,
    university_page_url,
    top_urls,
    search_engine,
    search_query,
    error_text
)
SELECT
    r.raw_id,
    o.organization_id,
    p.person_id,
    NULLIF(TRIM(r.officer_role), ''),
    NULLIF(TRIM(r.job_title), ''),
    NULLIF(TRIM(r.linkedin_url), ''),
    NULLIF(TRIM(r.other_social_urls), ''),
    NULLIF(TRIM(r.university_page_url), ''),
    NULLIF(TRIM(r.top_urls), ''),
    NULLIF(TRIM(r.search_engine), ''),
    NULLIF(TRIM(r.search_query), ''),
    NULLIF(TRIM(r.error_text), '')
FROM `{RAW}` AS r
INNER JOIN `{ORG}` AS o
    ON o.source_system = 'WHED'
   AND o.source_org_id = NULLIF(TRIM(r.iau_id), '')
INNER JOIN `{PERSON}` AS p
    ON p.full_name = TRIM(r.officer_name)
WHERE TRIM(COALESCE(r.officer_name, '')) <> ''
  AND TRIM(COALESCE(r.iau_id, '')) <> ''
ON DUPLICATE KEY UPDATE
    organization_id = VALUES(organization_id),
    person_id = VALUES(person_id),
    officer_role = VALUES(officer_role),
    job_title = VALUES(job_title),
    linkedin_url = VALUES(linkedin_url),
    other_social_urls = VALUES(other_social_urls),
    university_page_url = VALUES(university_page_url),
    top_urls = VALUES(top_urls),
    search_engine = VALUES(search_engine),
    search_query = VALUES(search_query),
    error_text = VALUES(error_text)
"""


def main() -> None:
    conn = mysql.connector.connect(
        host=HOST,
        user=USER,
        password=PASSWORD,
        database=DATABASE,
    )
    conn.autocommit = False
    cur = conn.cursor()

    try:
        if DROP_REBUILD:
            drop_core_tables(cur)
            conn.commit()

        cur.execute(CREATE_ORG)
        print("Table ready:", ORG)

        cur.execute(CREATE_PERSON)
        print("Table ready:", PERSON)

        cur.execute(CREATE_ORG_PERSON)
        print("Table ready:", ORG_PERSON)

        cur.execute(CREATE_LINK)
        print("Table ready:", LINK)

        cur.execute(LOAD_ORG)
        print("Organizations loaded. rowcount:", cur.rowcount)

        cur.execute(LOAD_PERSON)
        print("Persons loaded. rowcount:", cur.rowcount)

        cur.execute(LOAD_ORG_PERSON)
        print("Organization–person links loaded. rowcount:", cur.rowcount)

        cur.execute(LOAD_LINK)
        print("Officer links (designation + URLs) loaded. rowcount:", cur.rowcount)

        conn.commit()
        print("Done.")
    except Exception:
        conn.rollback()
        raise
    finally:
        cur.close()
        conn.close()


if __name__ == "__main__":
    main()
