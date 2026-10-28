-- =====================================================================
-- Metadata import: one function + one mapping table.
--
-- Replaces:
--   preview_metadata_import, metadata_timestamp_ist,
--   metadata_direction_code, metadata_duration_seconds
-- (their logic now lives inside import_metadata_rows).
--
-- How Excel columns are matched to "TblMst_Metadata_2026":
--   1. Headers listed in metadata_column_map go to the column(s) named
--      there (e.g. "Agentid" -> dealer_id and agent_id).
--   2. Any other header is matched to the table column with the same
--      name, ignoring case, spaces and punctuation
--      ("Personal Name", "personal_name", "PERSONALNAME" -> personal_name).
--   3. Headers that match nothing are ignored and reported back.
-- So a new Excel column needs no code change: if its header matches a
-- table column it is stored; if it is named differently, add one row to
-- metadata_column_map.
-- =====================================================================


-- ---------------------------------------------------------------------
-- 1. Header aliases (only for headers whose name differs from the column)
-- ---------------------------------------------------------------------

CREATE TABLE IF NOT EXISTS public.metadata_column_map (
    source_key    text NOT NULL,  -- Excel header: lower-case, letters/digits only
    target_column text NOT NULL,  -- exact column name in TblMst_Metadata_2026
    PRIMARY KEY (source_key, target_column)
);

INSERT INTO public.metadata_column_map (source_key, target_column) VALUES
    ('agentid',     'dealer_id'),
    ('agentid',     'agent_id'),
    ('dnis',         'dnis_code'),
    ('connecttime', 'local_start_time'),
    ('connecttime', 'local_end_time')   -- same as the old preview function; check this is intended
ON CONFLICT DO NOTHING;


-- ---------------------------------------------------------------------
-- 2. Import function
--
--   p_rows      JSON array of Excel rows, keys = original Excel headers.
--               Optional key "__row" = Excel row number (used in errors).
--   p_defaults  Values for columns the Excel file does not fill,
--               e.g. '{"branch_id": "BR_0007"}'. Excel values win.
-- ---------------------------------------------------------------------



CREATE OR REPLACE FUNCTION public.import_metadata_rows(
    p_rows     jsonb,
    p_defaults jsonb DEFAULT '{}'::jsonb
)

RETURNS jsonb
LANGUAGE plpgsql
AS $BODY$

DECLARE
    v_received     integer;
    v_message      text;
    v_error_count  integer;
    v_insert_cols  text;
    v_select_cols  text;
    v_inserted     integer;
    v_unique_rows  integer;
    v_repeated     integer;

BEGIN
    IF p_rows IS NULL OR jsonb_typeof(p_rows) <> 'array' THEN
        RAISE EXCEPTION 'Input must be a JSON array.';
    END IF;


    IF p_defaults IS NULL OR jsonb_typeof(p_defaults) <> 'object' THEN
        RAISE EXCEPTION 'p_defaults must be a JSON object.';
    END IF;


    v_received := jsonb_array_length(p_rows);

    IF v_received = 0 THEN
        RETURN jsonb_build_object(
            'status', 'success',
            'rows_received', 0,
            'rows_inserted', 0,
            'existing_database_duplicates', 0,
            'repeated_excel_call_ids', 0,
            'rows_skipped', 0,
            'mapped_columns', '[]'::jsonb,
            'unmapped_headers', '[]'::jsonb
        );
    END IF;



    DROP TABLE IF EXISTS _meta_cols, _meta_headers, _meta_map,
                         _meta_values, _meta_rows;

    -- Writable columns of the target table (identity, serial and
    -- generated columns are left to the database).
    CREATE TEMP TABLE _meta_cols ON COMMIT DROP AS
    SELECT

        a.attname::text AS column_name,
        t.typname::text AS type_name,
        CASE
            WHEN t.typname IN ('varchar', 'bpchar') AND a.atttypmod > 4
            THEN a.atttypmod - 4
        END AS max_len,

        lower(regexp_replace(a.attname, '[^a-zA-Z0-9]', '', 'g')) AS name_key
    FROM pg_attribute a
    JOIN pg_type t ON t.oid = a.atttypid



    LEFT JOIN pg_attrdef d
           ON d.adrelid = a.attrelid AND d.adnum = a.attnum
    WHERE a.attrelid = 'public."TblMst_Metadata_2026"'::regclass
      AND a.attnum > 0
      AND NOT a.attisdropped
      AND a.attidentity = ''
      AND a.attgenerated = ''
      AND COALESCE(pg_get_expr(d.adbin, d.adrelid), '') NOT LIKE 'nextval(%';

    -- Excel headers. Every row carries every header (the importer sends
    -- empty cells as null), so the first row is enough.



    CREATE TEMP TABLE _meta_headers ON COMMIT DROP AS
    SELECT
        k AS header,
        lower(regexp_replace(k, '[^a-zA-Z0-9]', '', 'g')) AS header_key
    FROM jsonb_object_keys(p_rows -> 0) AS k
    WHERE k <> '__row';

    -- Header -> column: alias table first, otherwise same normalized name.


    CREATE TEMP TABLE _meta_map ON COMMIT DROP AS
    SELECT h.header, c.column_name, c.type_name, c.max_len
    FROM _meta_headers h
    JOIN _meta_cols c
      ON c.column_name IN (
             SELECT m.target_column
             FROM public.metadata_column_map m
             WHERE m.source_key = h.header_key
         )
      OR (
             c.name_key = h.header_key
             AND NOT EXISTS (
                 SELECT 1
                 FROM public.metadata_column_map m
                 WHERE m.source_key = h.header_key
             )
         );

    IF NOT EXISTS (SELECT 1 FROM _meta_map) THEN
        RAISE EXCEPTION
            'None of the Excel headers match a column of TblMst_Metadata_2026: %',
            (SELECT string_agg(header, ', ') FROM _meta_headers);
    END IF;

    SELECT string_agg(format('%s <- %s', column_name, headers), '; ')//detect 2 csv file is trying to map to same column
    INTO v_message
    FROM (
        SELECT column_name, string_agg(format('"%s"', header), ', ') AS headers
        FROM _meta_map
        GROUP BY column_name
        HAVING count(*) > 1
    ) d;


    IF v_message IS NOT NULL THEN
        RAISE EXCEPTION
            'More than one Excel header maps to the same column: %. '
            'Remove or rename one of them.', v_message;
    END IF;

    SELECT string_agg(k, ', ')
    INTO v_message
    FROM jsonb_object_keys(p_defaults) AS k
    WHERE k NOT IN (SELECT column_name FROM _meta_cols);

    IF v_message IS NOT NULL THEN
        RAISE EXCEPTION 'Unknown column(s) in defaults: %', v_message;
    END IF;

    -- One row per (Excel row, column), converted to the column's type.
    -- A value that cannot be converted becomes NULL here and is reported
    -- by the check below instead of failing with no row number.


    CREATE TEMP TABLE _meta_values ON COMMIT DROP AS
    SELECT
        v.row_num,
        v.excel_row,
        v.column_name,
        v.raw,
        v.max_len,
        CASE
            WHEN v.raw IS NULL THEN NULL

            -- Columns with their own rules ---------------------------
            WHEN v.column_name = 'direction' THEN
                CASE
                    WHEN lower(v.raw) = 'inbound'     THEN to_jsonb(0)
                    WHEN lower(v.raw) = 'outbound'    THEN to_jsonb(1)
                    WHEN lower(v.raw) = 'transferred' THEN to_jsonb(2)
                    WHEN v.raw ~ '^[0-2]$' THEN to_jsonb(v.raw::integer)
                END


            WHEN v.column_name = 'duration' THEN
                CASE
                    -- seconds
                    WHEN v.raw ~ '^\d{1,9}$' THEN to_jsonb(v.raw::integer)

                    -- H:MM:SS
                    WHEN v.raw ~ '^\d{1,5}:[0-5]\d:[0-5]\d$' THEN to_jsonb(
                          split_part(v.raw, ':', 1)::integer * 3600
                        + split_part(v.raw, ':', 2)::integer * 60
                        + split_part(v.raw, ':', 3)::integer)
                    -- Excel time cell, sent as an ISO timestamp
                    WHEN v.raw ~ '^\d{4}-\d{2}-\d{2}T.*(Z|[+-]\d{2}:?\d{2})$' THEN to_jsonb(
                        floor(extract(epoch FROM
                            (v.raw::timestamptz AT TIME ZONE 'Asia/Kolkata')::time
                        ))::integer)
                END


            -- "aws s3 ls a.mp3, b.mp3" -> [{"path": "a.mp3"}, {"path": "b.mp3"}]

            WHEN v.column_name = 'file_location' THEN (
                SELECT COALESCE(
                    jsonb_agg(jsonb_build_object('path', trim(p.path)) ORDER BY p.ord),
                    '[]'::jsonb)
                FROM unnest(string_to_array(
                         regexp_replace(v.raw, '^aws\s+s3\s+ls\s+', '', 'i'), ','))
                     WITH ORDINALITY AS p(path, ord)
                WHERE trim(p.path) <> ''
            )

            -- Everything else, by column type ------------------------

            WHEN v.type_name IN ('int2', 'int4', 'int8') THEN
                CASE
                    WHEN v.raw ~ '^[+-]?\d{1,18}(\.0+)?$'
                    THEN to_jsonb(split_part(v.raw, '.', 1)::bigint)
                END

            WHEN v.type_name IN ('numeric', 'float4', 'float8') THEN
                CASE
                    WHEN v.raw ~ '^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$'
                    THEN to_jsonb(v.raw::numeric)
                END

            WHEN v.type_name = 'bool' THEN
                CASE
                    WHEN lower(v.raw) IN ('true', 't', 'yes', 'y', '1')  THEN 'true'::jsonb
                    WHEN lower(v.raw) IN ('false', 'f', 'no', 'n', '0') THEN 'false'::jsonb
                END

            -- Timestamps: "2026-09-01 08:00:00" is taken as IST and stored
            -- as written; one with a zone ("...Z", "...+05:30") is
            -- converted to IST.
            WHEN v.type_name IN ('timestamp', 'timestamptz', 'date') THEN
                CASE
                    WHEN v.raw !~ '^\d{4}-\d{2}-\d{2}([ T]\d{2}:\d{2}(:\d{2}(\.\d+)?)?)?\s*(Z|[+-]\d{2}(:?\d{2})?)?$'
                        THEN NULL
                    WHEN v.raw ~ '\d{2}:\d{2}(:\d{2}(\.\d+)?)?\s*(Z|[+-]\d{2}(:?\d{2})?)$' THEN
                        CASE v.type_name
                            WHEN 'timestamptz' THEN to_jsonb(v.raw::timestamptz)
                            WHEN 'date' THEN to_jsonb((v.raw::timestamptz AT TIME ZONE 'Asia/Kolkata')::date)
                            ELSE to_jsonb(v.raw::timestamptz AT TIME ZONE 'Asia/Kolkata')
                        END
                    ELSE
                        CASE v.type_name
                            WHEN 'timestamptz' THEN to_jsonb(v.raw::timestamp AT TIME ZONE 'Asia/Kolkata')
                            WHEN 'date' THEN to_jsonb(v.raw::timestamp::date)
                            ELSE to_jsonb(v.raw::timestamp)
                        END
                END

            WHEN v.type_name IN ('json', 'jsonb') THEN v.src

            -- text, varchar, uuid, ...
            ELSE to_jsonb(v.raw)
        END AS val
    FROM (
        SELECT
            r.ord AS row_num,
            COALESCE((r.value ->> '__row')::integer, r.ord::integer + 1) AS excel_row,
            m.column_name,
            m.type_name,
            m.max_len,
            r.value -> m.header AS src,
            NULLIF(trim(r.value ->> m.header), '') AS raw
        FROM jsonb_array_elements(p_rows) WITH ORDINALITY AS r(value, ord)
        CROSS JOIN _meta_map m
    ) v;

    SELECT count(*)::integer
    INTO v_error_count
    FROM _meta_values
    WHERE raw IS NOT NULL
      AND (val IS NULL OR length(raw) > max_len);

    IF v_error_count > 0 THEN
        SELECT string_agg(
                   format('  Excel row %s, column %s: %L', excel_row, column_name, raw),
                   E'\n' ORDER BY row_num, column_name)
        INTO v_message
        FROM (
            SELECT row_num, excel_row, column_name, raw
            FROM _meta_values
            WHERE raw IS NOT NULL
              AND (val IS NULL OR length(raw) > max_len)
            ORDER BY row_num, column_name
            LIMIT 20
        ) e;

        RAISE EXCEPTION E'% invalid value(s) found. Nothing was imported.\n%',
            v_error_count, v_message;
    END IF;

    -- Rebuild one JSON object per Excel row, keyed by table column.


    CREATE TEMP TABLE _meta_rows ON COMMIT DROP AS
    SELECT
        row_num,
        p_defaults || COALESCE(
            jsonb_object_agg(column_name, val) FILTER (WHERE val IS NOT NULL),
            '{}'::jsonb
        ) AS obj
    FROM _meta_values
    GROUP BY row_num;

    -- Insert only the columns that came from Excel or defaults; every
    -- other column keeps its database default.
    SELECT
        string_agg(quote_ident(column_name), ', ' ORDER BY column_name),
        string_agg('rec.' || quote_ident(column_name), ', ' ORDER BY column_name)
    INTO v_insert_cols, v_select_cols
    FROM (
        SELECT column_name FROM _meta_map
        UNION
        SELECT jsonb_object_keys(p_defaults)
        
    ) c;

    EXECUTE format($sql$
        WITH ranked AS MATERIALIZED (
            SELECT
                row_num,
                obj,
                row_number() OVER (
                    PARTITION BY
                        obj ->> 'call_id',
                        CASE WHEN obj ->> 'call_id' IS NULL THEN row_num END
                    ORDER BY row_num
                ) AS dup_rank
            FROM _meta_rows
        ),


        inserted AS (
            INSERT INTO public."TblMst_Metadata_2026" (%s)
            SELECT %s
            FROM ranked
            CROSS JOIN LATERAL jsonb_populate_record(
                NULL::public."TblMst_Metadata_2026", ranked.obj) AS rec
            WHERE ranked.dup_rank = 1
            ORDER BY ranked.row_num   -- interaction_id follows Excel order
            ON CONFLICT (call_id) DO NOTHING
            RETURNING 1
        )
        SELECT
            (SELECT count(*) FROM inserted)::integer,
            count(*) FILTER (WHERE dup_rank = 1)::integer,
            count(*) FILTER (WHERE dup_rank > 1)::integer
        FROM ranked
    $sql$, v_insert_cols, v_select_cols)
    INTO v_inserted, v_unique_rows, v_repeated;


    RETURN jsonb_build_object(
        'status', 'success',
        'rows_received', v_received,
        'rows_inserted', v_inserted,
        'existing_database_duplicates', v_unique_rows - v_inserted,
        'repeated_excel_call_ids', v_repeated,
        'rows_skipped', v_unique_rows - v_inserted + v_repeated,
        'mapped_columns', (
            SELECT jsonb_agg(
                       jsonb_build_object('excel_header', header, 'column', column_name)
                       ORDER BY header, column_name)
            FROM _meta_map
        ),
        'unmapped_headers', (
            SELECT COALESCE(jsonb_agg(header ORDER BY header), '[]'::jsonb)
            FROM _meta_headers h
            WHERE NOT EXISTS (SELECT 1 FROM _meta_map m WHERE m.header = h.header)
        )
    );
END;
$BODY$;


ALTER FUNCTION public.import_metadata_rows(jsonb, jsonb) OWNER TO postgres;

-- The old one-argument version would clash with calls like
-- import_metadata_rows($1::jsonb), so remove it.
DROP FUNCTION IF EXISTS public.import_metadata_rows(jsonb);


-- ---------------------------------------------------------------------
-- 3. Old helper functions — no longer used by the import.
--    Run these only once you are sure nothing else calls them.
-- ---------------------------------------------------------------------
-- DROP FUNCTION IF EXISTS public.preview_metadata_import(jsonb);
-- DROP FUNCTION IF EXISTS public.metadata_timestamp_ist(text);
-- DROP FUNCTION IF EXISTS public.metadata_direction_code(text);
-- DROP FUNCTION IF EXISTS public.metadata_duration_seconds(text);
