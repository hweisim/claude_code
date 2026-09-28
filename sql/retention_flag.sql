/* ====================================================================
   Balance retention flag  —  sticky / non-sticky from two snapshots

   Classifies a customer using only prev_bal and latest_bal.

   Design choices, and why:
     - One-sided.   Growth is never instability. A customer who gained
                    20% is retained, not volatile.
     - Floor.       5% of a THB 143 balance is THB 7, so a percentage
                    band alone makes tiny accounts fail on noise. The
                    tolerance is the LARGER of 5% and THB 100.
     - Gate.        Accounts below THB 500 prior have nothing to retain.
                    Scoring them as "non-sticky" is what put ~68k dormant
                    accounts in the risk corner of the old framework.
                    They are labelled, not scored.

   Known limitation: two snapshots cannot see the path between them.
   A customer who went 100k -> 0 -> 100k is classified Sticky here.
   Replace with a trough ratio once a daily balance panel exists.

   Parameters below are heuristics. Fit them against observed runoff.
   ==================================================================== */

SELECT
    ccd_id,
    prev_bal,
    latest_bal,
    latest_bal - prev_bal                                   AS bal_change,

    CASE
        /* 1. unusable input -------------------------------------- */
        WHEN prev_bal IS NULL OR latest_bal IS NULL
            THEN 'Unknown'

        /* 2. nothing to retain ----------------------------------- */
        WHEN prev_bal < 500
            THEN CASE WHEN latest_bal >= 500
                      THEN 'Newly funded'
                      ELSE 'Dormant'
                 END

        /* 3. retained: lost no more than 5%, or THB 100 ----------- */
        WHEN latest_bal >= prev_bal - GREATEST(0.05 * prev_bal, 100)
            THEN 'Sticky'

        /* 4. everything else lost material balance --------------- */
        ELSE 'Non-sticky'
    END                                                     AS retention_flag

FROM customer_balance;


/* --------------------------------------------------------------------
   Strict binary, if only two classes are allowed.
   Dormant accounts fall to Non-sticky here, which is the behaviour the
   gate above exists to avoid — use only when the schema forces it.
   -------------------------------------------------------------------- */
-- CASE
--     WHEN COALESCE(latest_bal, 0)
--          >= COALESCE(prev_bal, 0) - GREATEST(0.05 * COALESCE(prev_bal, 0), 100)
--          THEN 'Sticky'
--     ELSE 'Non-sticky'
-- END AS retention_flag


/* --------------------------------------------------------------------
   Symmetric variant — balance held WITHIN +/-5%, so growth also breaks
   the band. Flags customers whose balance barely moves in either
   direction. Useful for spotting parked money, not for retention.
   -------------------------------------------------------------------- */
-- CASE
--     WHEN prev_bal IS NULL OR latest_bal IS NULL      THEN 'Unknown'
--     WHEN prev_bal < 500                              THEN 'Not scored'
--     WHEN ABS(latest_bal - prev_bal)
--          <= GREATEST(0.05 * prev_bal, 100)           THEN 'Stable'
--     WHEN latest_bal > prev_bal                       THEN 'Growing'
--     ELSE 'Declining'
-- END AS balance_movement


/* ====================================================================
   Distribution check — run this before trusting the cut-offs.
   If 'Sticky' is over ~90% of the funded base the tolerance is too
   loose; if 'Dormant' is large, that is the population the retention
   campaign should be suppressing.
   ==================================================================== */

-- SELECT
--     retention_flag,
--     COUNT(*)                                          AS customers,
--     ROUND(100.0 * COUNT(*) / SUM(COUNT(*)) OVER (), 1) AS pct_customers,
--     SUM(prev_bal)                                     AS prev_bal,
--     SUM(latest_bal)                                   AS latest_bal,
--     ROUND(100.0 * SUM(latest_bal) / NULLIF(SUM(prev_bal), 0), 1) AS pct_retained
-- FROM (<the query above>) t
-- GROUP BY retention_flag
-- ORDER BY prev_bal DESC;
