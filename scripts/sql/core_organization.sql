-- Run in MySQL (set your database name).
-- Example:  USE organization_data;

CREATE TABLE IF NOT EXISTS core_organization (
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
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;

-- Optional: fill from raw WHED export (one row per institution).
-- Requires non-empty iau_id for UNIQUE (source_system, source_org_id) to dedupe.
-- If some rows have empty iau_id, use the Python loader instead (it handles that).
INSERT INTO core_organization (
    source_system,
    source_org_id,
    org_name,
    org_type,
    official_website,
    country,
    city
)
SELECT
    'WHED' AS source_system,
    NULLIF(TRIM(r.iau_id), '') AS source_org_id,
    TRIM(r.university_name) AS org_name,
    'university' AS org_type,
    NULLIF(TRIM(r.official_www), '') AS official_website,
    NULL AS country,
    NULL AS city
FROM raw_institutions_officers_social_raw AS r
INNER JOIN (
    SELECT
        NULLIF(TRIM(iau_id), '') AS iau_key,
        MAX(raw_id) AS max_raw_id
    FROM raw_institutions_officers_social_raw
    WHERE TRIM(COALESCE(university_name, '')) <> ''
      AND TRIM(COALESCE(iau_id, '')) <> ''
    GROUP BY NULLIF(TRIM(iau_id), '')
) AS pick
    ON pick.max_raw_id = r.raw_id
   AND pick.iau_key = NULLIF(TRIM(r.iau_id), '')
ON DUPLICATE KEY UPDATE
    org_name = VALUES(org_name),
    org_type = VALUES(org_type),
    official_website = VALUES(official_website),
    updated_at_utc = CURRENT_TIMESTAMP(3);
