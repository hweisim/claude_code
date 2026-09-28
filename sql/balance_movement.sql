/* ====================================================================
   Balance movement  —  Increase / Stable / Decrease

   Classifies a customer from two balance snapshots into exactly three
   groups, using a tolerance band around the prior balance.

       Increase :  gained  more than the band
       Stable   :  moved   within  the band, either direction
       Decrease :  lost    more than the band

   Band = the LARGER of 5% of prior balance and THB 100.
   The absolute floor matters: 5% of a THB 143 balance is THB 7, so a
   percentage band alone makes dormant accounts swing between Increase
   and Decrease on a single small deposit.

   MATERIALITY IS A SEPARATE COLUMN, NOT A FOURTH GROUP.
   On the current base, 154,561 customers — 48% of the base — hold under
   THB 500 and together account for 0.79% of the book. Left in, they
   dominate every headcount percentage while carrying no money. Filter
   on is_material for any funding conclusion; keep them for reach and
   reactivation work.

   Parameters live in the params CTE. They are heuristics, not fitted.
   ==================================================================== */

WITH params AS (
    SELECT
        0.05    AS pct_band,      /* relative half-width of Stable      */
        100     AS abs_floor,     /* minimum band in THB                */
        500     AS material_min   /* below this, balance is immaterial  */
),

bal AS (
    SELECT
        ccd_id,
        COALESCE(prev_bal,   0) AS prev_bal,
        COALESCE(latest_bal, 0) AS latest_bal
    FROM customer_balance
)

SELECT
    b.ccd_id,
    b.prev_bal,
    b.latest_bal,
    b.latest_bal - b.prev_bal                              AS bal_change,

    /* the band actually applied to this customer */
    GREATEST(p.pct_band * b.prev_bal, p.abs_floor)         AS band_thb,

    /* does this customer carry enough balance to matter? */
    CASE WHEN b.prev_bal   >= p.material_min
           OR b.latest_bal >= p.material_min
         THEN 1 ELSE 0 END                                 AS is_material,

    CASE
        WHEN b.latest_bal - b.prev_bal
             >  GREATEST(p.pct_band * b.prev_bal, p.abs_floor)
            THEN 'Increase'

        WHEN b.prev_bal - b.latest_bal
             >  GREATEST(p.pct_band * b.prev_bal, p.abs_floor)
            THEN 'Decrease'

        ELSE 'Stable'
    END                                                    AS bal_movement

FROM bal b
CROSS JOIN params p;


/* ====================================================================
   Summary — always read headcount and baht side by side.

   Every revision of the sticky flag so far has moved the headcount
   number while the money behind it went unreported. Customers and baht
   disagree sharply on this base: half the population holds 0.79% of it.
   ==================================================================== */

-- SELECT
--     bal_movement,
--     is_material,
--     COUNT(*)                                                    AS customers,
--     ROUND(100.0 * COUNT(*) / SUM(COUNT(*)) OVER (), 1)          AS pct_customers,
--     SUM(prev_bal)                                               AS prev_bal,
--     ROUND(100.0 * SUM(prev_bal)
--           / NULLIF(SUM(SUM(prev_bal)) OVER (), 0), 1)           AS pct_prev_bal,
--     SUM(latest_bal)                                             AS latest_bal,
--     SUM(latest_bal - prev_bal)                                  AS bal_change,
--     ROUND(AVG(prev_bal), 0)                                     AS avg_prev_bal
-- FROM (<the query above>) t
-- GROUP BY bal_movement, is_material
-- ORDER BY is_material DESC, prev_bal DESC;


/* --------------------------------------------------------------------
   OPTIONAL — a band that widens with account size.

   A flat 5% is a rounding error on a dormant account and a real
   constraint on a large one. On the current base the over-THB 100k
   group, holding 73% of the book, is the one most often pushed out of
   Stable. If Decrease is dominated by large balances, swap the band
   expression above for this and re-check the distribution.
   -------------------------------------------------------------------- */
-- CASE
--     WHEN prev_bal <   10000 THEN GREATEST(0.05 * prev_bal, 100)
--     WHEN prev_bal <  100000 THEN 0.07 * prev_bal
--     ELSE                         0.10 * prev_bal
-- END AS band_thb


/* --------------------------------------------------------------------
   OPTIONAL — map the three groups back to a sticky flag, if the
   downstream model needs a binary. Increase and Stable are both
   retention; only Decrease is balance leaving.
   -------------------------------------------------------------------- */
-- CASE WHEN bal_movement IN ('Increase', 'Stable') THEN 'Sticky'
--      ELSE 'Non-sticky'
-- END AS retention_flag
