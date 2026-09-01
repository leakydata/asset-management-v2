-- ============================================================================
-- Machines held back from Phase-1 inventory ONLY because of the model-year rule
-- ----------------------------------------------------------------------------
-- Mirrors the candidate filter in DEV.ERD.CCAT_ADD_INVENTORY_BATCH (v5) exactly,
-- with the year predicate INVERTED. So every row here passes every other test --
-- division G, CAT, active, unassigned, not already added, not blocked -- and is
-- excluded for one reason: no usable MANUFACTURERYEAR.
--
-- NAXT stores a missing year as 0, never NULL. An `IS NULL` test finds nothing;
-- the real test is the 4-digit range, matching the wrapper's own check in
-- file 50 (which would otherwise substitute the 1900 placeholder).
--
-- Fix the year in NAXT and the machine re-enters the backlog automatically --
-- nothing here needs to be re-run or un-flagged.
--
-- Scope note: the serial-length guard is set to 8 to match the batch call in
-- use (MAX_SERIAL_LEN => 8). Comment that line out to see the full set.
-- ============================================================================

with latest_hist as (
    select SERIALNUMBER, CCAT_RECORDS
    from DEV.STD_UMT.CCAT_HISTORY
    qualify row_number() over (partition by SERIALNUMBER
                               order by OBSERVED_AT desc, HISTORY_ID desc) = 1
),
known_ours as (
    select distinct l.SERIALNUMBER
    from latest_hist l, lateral flatten(input => l.CCAT_RECORDS) f
    where upper(f.value:ownership:dealerAssociation:dealerCode::string) = 'B150'
),
known_blocked as (
    select distinct l.SERIALNUMBER
    from latest_hist l, lateral flatten(input => l.CCAT_RECORDS) f
    where upper(f.value:ownership:dealerAssociation:dealerCode::string) <> 'B150'
      and upper(f.value:ownership:dealerAssociation:dcnOwnershipType:code::string) = 'OWNED'
      and upper(f.value:ownership:dealerAssociation:dcnRelationStatus:code::string) = 'ACTIVE'
)
select
      v.FLEETTYPE
    , v.EQUIPMENTNUMBER
    , v.SERIALNUMBER
    , v.MODEL
    , v.MANUFACTURERYEAR
    , case
          when v.MANUFACTURERYEAR is null then 'NULL'
          when v.MANUFACTURERYEAR = 0     then 'ZERO (NAXT "missing")'
          else 'OUT OF RANGE'
      end                                     as YEAR_PROBLEM
    , v.MODIFIEDDATETIME                      as MODIFIED
from DEV.EDW_ENT.EQUIPMENT_NAXT_VW v
where v.ATTACHMENT = 'No'
  and upper(v.MAKE) = 'CAT'
  and upper(trim(v.FIN_DIVISION)) = 'G'
  -- the rule, inverted: these are the ones the batch now skips
  and (v.MANUFACTURERYEAR is null
       or v.MANUFACTURERYEAR not between 1000 and 9999)
  and v.EXCLUDEFROMDDSW = 0
  and v.ISACTIVE = TRUE
  and (v.CUSTOMERNUMBER is null or trim(v.CUSTOMERNUMBER) = '')
  and v.SERIALNUMBER is not null and trim(v.SERIALNUMBER) <> ''
  and length(trim(v.SERIALNUMBER)) <= 8          -- matches MAX_SERIAL_LEN => 8
  and not exists (select 1 from DEV.STD_UMT.CCAT_AUDIT a
                  where a.SERIALNUMBER = v.SERIALNUMBER
                    and a.ACTION_TYPE in ('ADD', 'ADD_INVENTORY')
                    and a.OUTCOME = 'EXECUTED')
  and not exists (select 1 from DEV.STD_UMT.CCAT_AUDIT a
                  where a.SERIALNUMBER = v.SERIALNUMBER
                    and a.OUTCOME = 'FAILED'
                    and a.RAW_RESPONSE:code::string = '400.203')
  and not exists (select 1 from known_ours k    where k.SERIALNUMBER = v.SERIALNUMBER)
  and not exists (select 1 from known_blocked b where b.SERIALNUMBER = v.SERIALNUMBER)
qualify row_number() over (partition by v.SERIALNUMBER
                           order by v.MODIFIEDDATETIME asc, v.EQUIPMENTNUMBER) = 1
order by v.FLEETTYPE, v.SERIALNUMBER;


-- ----------------------------------------------------------------------------
-- Counts only, if you just want the shape of it. Swap the SELECT above for:
--
--   select v.FLEETTYPE, count(distinct v.SERIALNUMBER) as N ... group by 1 order by 2 desc
--
-- Expected (as of 2026-09-01): RENTAL 21, CANCELLED 6, LEASE 4, RE-RENT 4 = 35.
-- ----------------------------------------------------------------------------
