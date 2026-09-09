-- ============================================================================
-- Phase counts for the CCAT work
--
-- Written 2026-09-08 from the phase diagram.
--
-- Every CTE selects ONLY the columns something downstream actually uses.  The
-- first version carried SERIALNUMBER, CAT_MAKE_CODE, OUR_OWNERSHIP_TYPE,
-- OUR_REQUEST_TYPE and CCAT_HAS_PENDING that nothing referenced - all of them
-- failure surface for no benefit, and CAT_MAKE_CODE duly turned out not to
-- exist on the view.
--
-- Query 0 checks the names before you trust any of the numbers.  A wrong
-- column name fails loud; a wrong VALUE (a verdict spelled differently from
-- 'MISSING') fails silent and just reports zero, which is the dangerous one.
--
-- Query 1 is the grid - it says whether the phase taxonomy holds.
-- Query 2 rolls the grid up into your phase labels.
-- Query 3 is the fleet-type breakdown for the inventory question.
-- ============================================================================


-- ----------------------------------------------------------------------------
-- 0. NAME CHECK.  Run this first - it takes a second and it is the difference
--    between a number you can quote and a number you cannot.
-- ----------------------------------------------------------------------------
SELECT 'EQUIPMENT_NAXT_VW' AS object, COLUMN_NAME
FROM DEV.INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_SCHEMA = 'EDW_ENT' AND TABLE_NAME = 'EQUIPMENT_NAXT_VW'
  AND (COLUMN_NAME ILIKE '%CUSTOMER%' OR COLUMN_NAME ILIKE '%FLEET%'
       OR COLUMN_NAME ILIKE '%EQUIPMENT%' OR COLUMN_NAME ILIKE '%ACTIVE%'
       OR COLUMN_NAME ILIKE '%MAKE%'      OR COLUMN_NAME ILIKE '%SERIAL%')
UNION ALL
SELECT 'CCAT_CHECK_STATE', COLUMN_NAME
FROM DEV.INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_SCHEMA = 'STD_UMT' AND TABLE_NAME = 'CCAT_CHECK_STATE'
ORDER BY 1, 2;

-- and the values, which are the silent ones:
-- SELECT VERDICT, COUNT(*) FROM DEV.STD_UMT.CCAT_CHECK_STATE GROUP BY 1 ORDER BY 2 DESC;


-- ----------------------------------------------------------------------------
-- 1. THE GRID.  Every machine, by what NAXT says and what CCAT says.
--
--    Your phases are cells in this grid.  Reading it this way is what shows
--    that 1C and 3C are the same CCAT situation seen from two different NAXT
--    sides, rather than two separate phases.
-- ----------------------------------------------------------------------------
WITH naxt AS (
    SELECT EQUIPMENTNUMBER,
           NULLIF(TRIM(CUSTOMERNUMBER), '') AS customer_number
    FROM DEV.EDW_ENT.EQUIPMENT_NAXT_VW
    WHERE ISACTIVE = TRUE           -- deliberately the whole active fleet, not
),                                  -- just Phase-1-eligible machines: this grid
                                    -- is meant to account for everything.  The
                                    -- eligibility rule lives in query 3.
state AS (
    SELECT EQUIPMENTNUMBER, VERDICT, OUR_DCN, CCAT_OTHER_DEALERS
    FROM DEV.STD_UMT.CCAT_CHECK_STATE
)
SELECT
    CASE WHEN n.customer_number IS NULL
         THEN 'stock (no customer)' ELSE 'customer' END        AS naxt_side,

    CASE
        WHEN s.EQUIPMENTNUMBER IS NULL               THEN 'never checked'
        WHEN s.VERDICT = 'MISSING'                   THEN 'absent'
        WHEN s.VERDICT = 'MISSING_OTHER_DEALER'      THEN 'another dealer only'
        WHEN s.OUR_DCN = 'INT00495'                  THEN 'ours, on INT00495'
        WHEN COALESCE(s.CCAT_OTHER_DEALERS,'') <> '' THEN 'ours + another dealer'
        WHEN s.OUR_DCN IS NOT NULL                   THEN 'ours, on a customer DCN'
        ELSE 'other'
    END                                                        AS ccat_side,

    -- does CCAT agree with NAXT about WHICH customer?
    COALESCE(
        CASE
            WHEN n.customer_number IS NULL OR s.OUR_DCN IS NULL THEN NULL
            WHEN s.OUR_DCN = n.customer_number                  THEN 'same DCN'
            ELSE 'different DCN'
        END, '-')                                              AS dcn_agreement,

    COUNT(*)                                                   AS machines
FROM naxt n
LEFT JOIN state s ON s.EQUIPMENTNUMBER = n.EQUIPMENTNUMBER
GROUP BY 1, 2, 3
ORDER BY 1, 2, 3;


-- ----------------------------------------------------------------------------
-- 2. THE SAME THING, ROLLED UP INTO YOUR PHASE LABELS.
--
--    Order matters: the CASE stops at the first match, so the phases are
--    written most-specific first.  Every machine lands in exactly one phase,
--    which is the property that makes these numbers add up to the fleet.
-- ----------------------------------------------------------------------------
WITH naxt AS (
    SELECT EQUIPMENTNUMBER,
           NULLIF(TRIM(CUSTOMERNUMBER), '') AS customer_number
    FROM DEV.EDW_ENT.EQUIPMENT_NAXT_VW
    WHERE ISACTIVE = TRUE
),
state AS (
    SELECT EQUIPMENTNUMBER, VERDICT, OUR_DCN, CCAT_OTHER_DEALERS
    FROM DEV.STD_UMT.CCAT_CHECK_STATE
)
SELECT phase, COUNT(*) AS machines
FROM (
    SELECT CASE
        -- Not looked at yet.  Broken out on purpose: an unchecked machine is
        -- not a phase, it is an absence of information, and folding it into
        -- 1A would inflate the add backlog with things that may already exist.
        WHEN s.EQUIPMENTNUMBER IS NULL
            THEN '0  - not yet checked'

        -- ---- dealer stock (NAXT customer blank) ----
        WHEN n.customer_number IS NULL AND s.VERDICT = 'MISSING'
            THEN '1A - stock, absent from CCAT          -> add on INT00495'
        WHEN n.customer_number IS NULL
             AND (s.VERDICT = 'MISSING_OTHER_DEALER'
                  OR COALESCE(s.CCAT_OTHER_DEALERS,'') <> '')
            THEN '1C - stock, another dealer holds it   -> transfer'
        WHEN n.customer_number IS NULL AND s.OUR_DCN = 'INT00495'
            THEN '1B - stock, on INT00495               -> nothing until sold'
        WHEN n.customer_number IS NULL AND s.OUR_DCN IS NOT NULL
            THEN '1B*- stock, but on a customer DCN     -> check, likely sold'

        -- ---- customer-owned (NAXT customer populated) ----
        WHEN s.VERDICT = 'MISSING'
            THEN '2A - customer, absent from CCAT       -> add'
        WHEN s.VERDICT = 'MISSING_OTHER_DEALER'
             OR COALESCE(s.CCAT_OTHER_DEALERS,'') <> ''
            THEN '3C - customer, another dealer holds it -> transfer'
        WHEN s.OUR_DCN = 'INT00495'
            THEN '1B - sold, still on INT00495          -> move to customer'
        WHEN s.OUR_DCN <> n.customer_number
            THEN '3C - customer, CCAT has another DCN   -> reassign'
        WHEN s.OUR_DCN = n.customer_number
            THEN 'OK - agrees'
        ELSE 'other'
    END AS phase
    FROM naxt n
    LEFT JOIN state s ON s.EQUIPMENTNUMBER = n.EQUIPMENTNUMBER
)
GROUP BY phase
ORDER BY phase;


-- ----------------------------------------------------------------------------
-- 3. WHAT IS LEFT IN INVENTORY, BY FLEET TYPE - ELIGIBLE MACHINES ONLY.
--
--    The Phase-1 expansion question: Lindsay approved 'New' only, and these
--    are the counts that say what each additional fleet type would cost.
--
--    The WHERE clause is the Phase-1 eligibility rule, so every machine
--    counted here is one the add proc would actually accept.  That is the
--    number worth quoting - an unfiltered count includes attachments and
--    non-Cat iron that were never candidates, and overstates the work.
-- ----------------------------------------------------------------------------
SELECT n.FLEETTYPE,
       COUNT(*)                                          AS machines,
       COUNT_IF(s.EQUIPMENTNUMBER IS NULL)               AS never_checked,
       COUNT_IF(s.VERDICT = 'MISSING')                   AS absent_from_ccat,
       COUNT_IF(s.OUR_DCN = 'INT00495')                  AS already_on_int00495,
       COUNT_IF(COALESCE(s.CCAT_OTHER_DEALERS,'') <> '') AS other_dealer_holds
FROM DEV.EDW_ENT.EQUIPMENT_NAXT_VW n
LEFT JOIN DEV.STD_UMT.CCAT_CHECK_STATE s
       ON s.EQUIPMENTNUMBER = n.EQUIPMENTNUMBER
WHERE n.ISACTIVE        = TRUE
  AND n.ATTACHMENT      = 'No'
  AND n.MAKE            = 'CAT'
  AND n.EXCLUDEFROMDDSW = 0
  AND NULLIF(TRIM(n.CUSTOMERNUMBER), '') IS NULL     -- dealer stock only
GROUP BY n.FLEETTYPE
ORDER BY machines DESC;


-- ----------------------------------------------------------------------------
-- 4. THE SAME FILTER, AS A FUNNEL.
--
--    How much each eligibility rule removes.  Worth having when somebody asks
--    why the inventory number is smaller than the fleet number - the answer
--    is a line in this table rather than an argument.
-- ----------------------------------------------------------------------------
SELECT COUNT(*)                                                AS all_equipment,
       COUNT_IF(ISACTIVE = TRUE)                               AS active,
       COUNT_IF(ISACTIVE = TRUE
                AND NULLIF(TRIM(CUSTOMERNUMBER),'') IS NULL)   AS active_stock,
       COUNT_IF(ISACTIVE = TRUE
                AND NULLIF(TRIM(CUSTOMERNUMBER),'') IS NULL
                AND MAKE = 'CAT')                              AS plus_cat_make,
       COUNT_IF(ISACTIVE = TRUE
                AND NULLIF(TRIM(CUSTOMERNUMBER),'') IS NULL
                AND MAKE = 'CAT'
                AND ATTACHMENT = 'No')                         AS plus_no_attachment,
       COUNT_IF(ISACTIVE = TRUE
                AND NULLIF(TRIM(CUSTOMERNUMBER),'') IS NULL
                AND MAKE = 'CAT'
                AND ATTACHMENT = 'No'
                AND EXCLUDEFROMDDSW = 0)                       AS eligible
FROM DEV.EDW_ENT.EQUIPMENT_NAXT_VW;


-- ============================================================================
-- WHAT DEV.ERD.CCAT_ADD_INVENTORY_BATCH HAS PROCESSED
--
-- Rewritten 2026-09-09 against the actual source of file 50 / file 53, which
-- corrected three guesses:
--
--   * there is no RUN_ID and no CREATED_AT.  The inner proc writes SHIFT_DATE
--     = current_date(), a DATE, and that is the only time column it sets.
--   * SOURCE is hard-coded 'MANUAL' on every row, batch or single, so it
--     cannot tell them apart.
--   * NOTES is passed straight through from the batch's own NOTES parameter.
--     It is the ONLY column a caller controls, so it is the only thing that
--     can mark a run - and only if a value was passed at call time.
--
-- So: ACTION_TYPE = 'ADD_INVENTORY' is every inventory add ever made, batch or
-- single.  Nothing in the row says which proc made it.  If NOTES was set per
-- run, filter on that; otherwise SHIFT_DATE clusters the rounds.
--
-- Dry runs really are absent: the inner proc returns before the INSERT when
-- DRY_RUN is true, so DRY_RUN is FALSE on every row that exists.
-- ============================================================================


-- ----------------------------------------------------------------------------
-- 5. TRY THE VIEW FIRST.
--
--    The batch's own closing note points at V_CCAT_INVENTORY_ADDS, so this was
--    already solved once.  If it exists, use it rather than any of the below.
-- ----------------------------------------------------------------------------
SHOW VIEWS LIKE 'V_CCAT_INVENTORY%' IN DATABASE DEV;
-- SELECT * FROM DEV.ERD.V_CCAT_INVENTORY_ADDS LIMIT 20;


-- ----------------------------------------------------------------------------
-- 6. THE ROLL-UP.  Every inventory add, by day and outcome.
--
--    PHASE is worth having beside it: the inner proc sets 'P1' when CCAT had
--    no record at all and 'P2' when it had some (just none of ours), so the
--    split says how much of the backlog is genuinely new iron.
-- ----------------------------------------------------------------------------
SELECT a.SHIFT_DATE,
       a.PHASE,
       a.OUTCOME,
       COUNT(*)                              AS machines,
       COUNT(DISTINCT a.NOTES)               AS distinct_notes,
       MAX(a.NOTES)                          AS sample_note
FROM DEV.STD_UMT.CCAT_AUDIT a
WHERE a.ACTION_TYPE = 'ADD_INVENTORY'
GROUP BY 1, 2, 3
ORDER BY 1 DESC, 2, 3;


-- ----------------------------------------------------------------------------
-- 7. THE LIST - every machine we have put into CCAT.  *** THIS IS THE ONE ***
--
--    ADD_INVENTORY is the whole upload history: inventory adds are the only
--    writes that have been executed, so this is not a subset of the work, it
--    is the work.  OUTCOME = 'EXECUTED' keeps it to machines that really
--    landed; drop that line to see the failed attempts alongside them.
--
--    Fleet type and customer come out of EQUIPMENT_SNAPSHOT - the NAXT row
--    captured at the moment of the add - not from the view as it stands today.
--    A machine added in August may since have been sold, and joining live NAXT
--    would quietly rewrite history to say it never was inventory.
-- ----------------------------------------------------------------------------
--    MODEL YEAR.  The audit holds both halves, so the placeholders are
--    identifiable exactly rather than inferred:
--
--      EQUIPMENT_SNAPSHOT:MANUFACTURERYEAR = what NAXT had at the time
--      REQUEST_PAYLOAD:body:modelYear      = what was actually sent to Cat
--
--    The test is TRY_TO_NUMBER ... BETWEEN 1000 AND 9999, not IS NOT NULL,
--    for two reasons: NAXT stores a missing year as 0, never NULL, and the
--    snapshot is an OBJECT_CONSTRUCT so the value can arrive as 2006.000000000.
--    Deriving "missing" from NAXT's own value rather than from a sent year of
--    1900 also avoids libelling any machine genuinely recorded as 1900.
WITH adds AS (
    SELECT a.SHIFT_DATE                                        AS processed_on,
           a.EQUIPMENTNUMBER,
           a.SERIALNUMBER,
           a.EQUIPMENT_SNAPSHOT:FLEETTYPE::string              AS fleettype_then,
           a.MODEL,
           a.OWNERSHIP_TYPE,

           -- the three years: then, sent, and now
           TRY_TO_NUMBER(a.EQUIPMENT_SNAPSHOT:MANUFACTURERYEAR::string)
                                                               AS naxt_year_then,
           a.REQUEST_PAYLOAD:body:modelYear::string            AS year_sent_to_cat,
           TRY_TO_NUMBER(n.MANUFACTURERYEAR::string)           AS naxt_year_now,

           a.PHASE,
           a.OUTCOME,
           a.HTTP_STATUS,
           a.DCN_AFTER,
           a.RAW_RESPONSE:code::string                         AS error_code,
           a.NOTES,
           a.CCAT_TRACKING_ID,
           -- where it sits now, so anything that has moved on is visible
           s.VERDICT                                           AS ccat_verdict_now,
           s.OUR_DCN                                           AS ccat_dcn_now,
           NULLIF(TRIM(n.CUSTOMERNUMBER), '')                  AS naxt_customer_now
    FROM DEV.STD_UMT.CCAT_AUDIT a
    LEFT JOIN DEV.EDW_ENT.EQUIPMENT_NAXT_VW n
           ON n.EQUIPMENTNUMBER = a.EQUIPMENTNUMBER
    LEFT JOIN DEV.STD_UMT.CCAT_CHECK_STATE s
           ON s.EQUIPMENTNUMBER = a.EQUIPMENTNUMBER
    WHERE a.ACTION_TYPE = 'ADD_INVENTORY'
      AND a.OUTCOME     = 'EXECUTED'    -- drop this line to include failures
)
SELECT CASE
           -- COALESCE to 0 first: a NULL year would make NOT BETWEEN evaluate
           -- to NULL and fall through to 'from NAXT', which is the wrong way
           -- round for a missing value.
           WHEN COALESCE(naxt_year_then, 0) NOT BETWEEN 1000 AND 9999
               THEN 'MISSING - sent 1900 placeholder'
           WHEN year_sent_to_cat <> TO_VARCHAR(naxt_year_then)
               THEN 'override used'
           ELSE 'from NAXT'
       END                                                     AS year_status,

       -- a placeholder NAXT can now answer is a backfill waiting to happen
       (COALESCE(naxt_year_then, 0) NOT BETWEEN 1000 AND 9999
        AND COALESCE(naxt_year_now, 0) BETWEEN 1000 AND 9999)  AS backfillable_now,

       a.*
FROM adds a
ORDER BY CASE WHEN COALESCE(naxt_year_then, 0) NOT BETWEEN 1000 AND 9999
              THEN 0 ELSE 1 END,        -- missing years first
         backfillable_now DESC,
         processed_on DESC,
         SERIALNUMBER;


-- ----------------------------------------------------------------------------
-- 7b. THE SAME THING AS A COUNT - how big is the year problem.
-- ----------------------------------------------------------------------------
SELECT a.SHIFT_DATE,
       COUNT(*)                                                       AS uploaded,
       COUNT_IF(COALESCE(TRY_TO_NUMBER(a.EQUIPMENT_SNAPSHOT:MANUFACTURERYEAR::string), 0)
                NOT BETWEEN 1000 AND 9999)                            AS no_year_sent_1900,
       COUNT_IF(COALESCE(TRY_TO_NUMBER(a.EQUIPMENT_SNAPSHOT:MANUFACTURERYEAR::string), 0)
                    NOT BETWEEN 1000 AND 9999
                AND COALESCE(TRY_TO_NUMBER(n.MANUFACTURERYEAR::string), 0)
                    BETWEEN 1000 AND 9999)                            AS backfillable_now
FROM DEV.STD_UMT.CCAT_AUDIT a
LEFT JOIN DEV.EDW_ENT.EQUIPMENT_NAXT_VW n
       ON n.EQUIPMENTNUMBER = a.EQUIPMENTNUMBER
WHERE a.ACTION_TYPE = 'ADD_INVENTORY'
  AND a.OUTCOME     = 'EXECUTED'
GROUP BY a.SHIFT_DATE
ORDER BY a.SHIFT_DATE DESC;


-- ----------------------------------------------------------------------------
-- 8. THE REMAINING BACKLOG - the batch's own qualifying rules, exactly.
--
--    My earlier version of this was wrong and overstated the backlog.  It
--    missed three rules the proc actually applies:
--
--      * FIN_DIVISION = 'G' only          (scope rule agreed 2026-08-17)
--      * MANUFACTURERYEAR between 1000 and 9999
--        (2026-08-24: NAXT stores "missing" as 0, never NULL, so IS NOT NULL
--         alone lets every one of them through)
--      * the history-based skips - serials whose latest CCAT snapshot already
--        shows one of ours, or shows another dealer holding OWNED/ACTIVE
--
--    It also keyed the audit check on EQUIPMENTNUMBER; the proc keys it on
--    SERIALNUMBER.  Set FLEET_TYPE below to the lane you are pacing.
-- ----------------------------------------------------------------------------
WITH latest_hist AS (
    SELECT SERIALNUMBER, CCAT_RECORDS
    FROM DEV.STD_UMT.CCAT_HISTORY
    QUALIFY ROW_NUMBER() OVER (PARTITION BY SERIALNUMBER
                               ORDER BY OBSERVED_AT DESC, HISTORY_ID DESC) = 1
),
known_ours AS (
    SELECT DISTINCT l.SERIALNUMBER
    FROM latest_hist l, LATERAL FLATTEN(input => l.CCAT_RECORDS) f
    WHERE UPPER(f.value:ownership:dealerAssociation:dealerCode::string) = 'B150'
),
known_blocked AS (
    -- another dealer holds OWNED/ACTIVE: a quiet add is impossible, these are
    -- the transfer worklist rather than the add backlog
    SELECT DISTINCT l.SERIALNUMBER
    FROM latest_hist l, LATERAL FLATTEN(input => l.CCAT_RECORDS) f
    WHERE UPPER(f.value:ownership:dealerAssociation:dealerCode::string) <> 'B150'
      AND UPPER(f.value:ownership:dealerAssociation:dcnOwnershipType:code::string) = 'OWNED'
      AND UPPER(f.value:ownership:dealerAssociation:dcnRelationStatus:code::string) = 'ACTIVE'
)
SELECT v.FLEETTYPE,
       COUNT(DISTINCT v.SERIALNUMBER) AS still_to_do
FROM DEV.EDW_ENT.EQUIPMENT_NAXT_VW v
WHERE v.ATTACHMENT = 'No'
  AND UPPER(v.MAKE) = 'CAT'
  AND UPPER(TRIM(v.FIN_DIVISION)) = 'G'
  AND v.MANUFACTURERYEAR IS NOT NULL
  AND v.MANUFACTURERYEAR BETWEEN 1000 AND 9999
  AND v.EXCLUDEFROMDDSW = 0
  AND v.ISACTIVE = TRUE
  AND (v.CUSTOMERNUMBER IS NULL OR TRIM(v.CUSTOMERNUMBER) = '')
  AND v.SERIALNUMBER IS NOT NULL AND TRIM(v.SERIALNUMBER) <> ''
  AND NOT EXISTS (SELECT 1 FROM DEV.STD_UMT.CCAT_AUDIT a
                  WHERE a.SERIALNUMBER = v.SERIALNUMBER
                    AND a.ACTION_TYPE IN ('ADD', 'ADD_INVENTORY')
                    AND a.OUTCOME = 'EXECUTED')
  AND NOT EXISTS (SELECT 1 FROM DEV.STD_UMT.CCAT_AUDIT a
                  WHERE a.SERIALNUMBER = v.SERIALNUMBER
                    AND a.OUTCOME = 'FAILED'
                    AND a.RAW_RESPONSE:code::string = '400.203')
  AND NOT EXISTS (SELECT 1 FROM known_ours k    WHERE k.SERIALNUMBER = v.SERIALNUMBER)
  AND NOT EXISTS (SELECT 1 FROM known_blocked b WHERE b.SERIALNUMBER = v.SERIALNUMBER)
GROUP BY v.FLEETTYPE
ORDER BY still_to_do DESC;


-- ----------------------------------------------------------------------------
-- 9. WHAT THE SCOPE RULES COST, per fleet type.
--
--    Same population, but each rule relaxed one at a time - so when somebody
--    asks "why is New only showing 40 left when we had 718", the answer is a
--    column here.  Run it before widening any of the constants in file 53.
-- ----------------------------------------------------------------------------
SELECT v.FLEETTYPE,
       COUNT(DISTINCT v.SERIALNUMBER)                                AS stock_total,
       COUNT(DISTINCT CASE WHEN UPPER(TRIM(v.FIN_DIVISION)) = 'G'
                           THEN v.SERIALNUMBER END)                  AS in_division_g,
       COUNT(DISTINCT CASE WHEN v.MANUFACTURERYEAR NOT BETWEEN 1000 AND 9999
                             OR v.MANUFACTURERYEAR IS NULL
                           THEN v.SERIALNUMBER END)                  AS dropped_no_year,
       COUNT(DISTINCT CASE WHEN UPPER(TRIM(v.FIN_DIVISION)) <> 'G'
                           THEN v.SERIALNUMBER END)                  AS dropped_other_division
FROM DEV.EDW_ENT.EQUIPMENT_NAXT_VW v
WHERE v.ATTACHMENT = 'No'
  AND UPPER(v.MAKE) = 'CAT'
  AND v.EXCLUDEFROMDDSW = 0
  AND v.ISACTIVE = TRUE
  AND (v.CUSTOMERNUMBER IS NULL OR TRIM(v.CUSTOMERNUMBER) = '')
  AND v.SERIALNUMBER IS NOT NULL AND TRIM(v.SERIALNUMBER) <> ''
GROUP BY v.FLEETTYPE
ORDER BY stock_total DESC;
