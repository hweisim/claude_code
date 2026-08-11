-- =============================================================================
-- Account-closure reason grouping (Redshift)
-- -----------------------------------------------------------------------------
-- Buckets free-text closure reasons (Thai + English + junk input) into a small
-- set of English key reasons.
--
-- IMPORTANT: the CASE arms are ordered by priority, not by volume.
-- More specific patterns must stay ABOVE more generic ones, e.g.
--   'ระบบห่วย'  -> System / technical issues   (not "Dissatisfied", ห่วย)
--   'ระบบไม่ดี' -> System / technical issues   (not "Dissatisfied", ไม่ดี)
--   'กู้ยาก'    -> Loan application rejected   (not "Hard to use", ยาก)
--   'ไม่มีประโยชน์' -> No longer needed        (not "No reason given", ไม่มี)
-- Reordering the arms will change the results.
-- =============================================================================

WITH base AS (
    SELECT
        rsn,
        BTRIM(LOWER(rsn)) AS rsn_clean
    FROM your_schema.your_table          -- <<< replace with the real source
)
SELECT
    rsn,
    CASE
        ---------------------------------------------------------------------
        -- 0. Junk / non-answers: blank, punctuation only, digits only,
        --    politeness particles, single stray characters.
        ---------------------------------------------------------------------
        WHEN rsn IS NULL
             OR rsn_clean = ''
             OR rsn_clean ~ '^[[:space:][:punct:][:digit:]]*$'   -- '.', '-', '..', '_', '1', '0', '....'
             OR rsn_clean IN ('n/a','na','none','no','nil','-','_',
                              'ไม่มี','ไม่','ไม่มีเหตุผล','ค่ะ','คะ','ครับ','ค','ๆ')
            THEN 'No reason given / invalid response'

        ---------------------------------------------------------------------
        -- 1. Explicit "other" option picked in the survey.
        ---------------------------------------------------------------------
        WHEN rsn_clean IN ('อื่นๆ','อื่น ๆ','other','others')
            THEN 'Other (customer selected "Other")'

        ---------------------------------------------------------------------
        -- 2. Fees too high (before "better offer": a fee complaint can also
        --    mention ดอกเบี้ย / interest).
        ---------------------------------------------------------------------
        WHEN rsn_clean LIKE '%ค่าธรรมเนียม%'      -- fee
             OR rsn_clean LIKE '%ค่าบริการ%'      -- service charge
             OR rsn_clean LIKE '%fee%'
             OR rsn_clean LIKE '%charge%'
            THEN 'Fees too high'

        ---------------------------------------------------------------------
        -- 3. Better offer / rate elsewhere.
        ---------------------------------------------------------------------
        WHEN rsn_clean LIKE '%ข้อเสนอ%'           -- offer
             OR rsn_clean LIKE '%ดอกเบี้ย%'        -- interest rate
             OR rsn_clean LIKE '%ที่อื่น%'          -- elsewhere
             OR rsn_clean LIKE '%ที่ดีกว่า%'        -- better
             OR rsn_clean LIKE '%better offer%'
             OR rsn_clean LIKE '%interest rate%'
            THEN 'Found a better offer or rate elsewhere'

        ---------------------------------------------------------------------
        -- 4. Security / fraud concerns.
        ---------------------------------------------------------------------
        WHEN rsn_clean LIKE '%ปลอดภัย%'           -- safe / security
             OR rsn_clean LIKE '%กังวล%'           -- worried
             OR rsn_clean LIKE '%กลัว%'            -- afraid
             OR rsn_clean LIKE '%หลอก%'            -- scam / deceive
             OR rsn_clean LIKE '%โกง%'             -- cheat
             OR rsn_clean LIKE '%security%'
             OR rsn_clean LIKE '%scam%'
             OR rsn_clean LIKE '%fraud%'
            THEN 'Security / trust concerns'

        ---------------------------------------------------------------------
        -- 5. Loan / credit application rejected or hard to get approved.
        --    Covers สินเชื่อ, กู้, สมัคร...ไม่ผ่าน, อนุมัติ, ผ่านยาก, bare 'ไม่ผ่าน'.
        ---------------------------------------------------------------------
        WHEN rsn_clean LIKE '%สินเชื่อ%'           -- loan / credit
             OR rsn_clean LIKE '%กู้%'             -- borrow
             OR rsn_clean LIKE '%อนุมัติ%'          -- approve
             OR rsn_clean LIKE '%ไม่ผ่าน%'          -- not approved / rejected
             OR rsn_clean LIKE '%ผ่านยาก%'          -- hard to get approved
             OR rsn_clean LIKE '%วงเงิน%'           -- credit limit
             OR rsn_clean LIKE '%loan%'
             OR rsn_clean LIKE '%credit%'
             OR rsn_clean LIKE '%reject%'
             OR rsn_clean LIKE '%approve%'
            THEN 'Loan / credit application rejected'

        ---------------------------------------------------------------------
        -- 6. System / technical problems and slowness.
        --    MUST stay above "Dissatisfied" (ระบบห่วย, ระบบแย่, ระบบไม่ดี).
        ---------------------------------------------------------------------
        WHEN rsn_clean LIKE '%ระบบ%'              -- system
             OR rsn_clean LIKE '%ช้า%'             -- slow
             OR rsn_clean LIKE '%ล่าช้า%'           -- delayed
             OR rsn_clean LIKE '%ล่ม%'             -- crash / down
             OR rsn_clean LIKE '%ค้าง%'            -- freeze / hang
             OR rsn_clean LIKE '%รวน%'             -- glitchy
             OR rsn_clean LIKE '%เสถียร%'           -- (un)stable
             OR rsn_clean LIKE '%ประมวลผล%'         -- processing
             OR rsn_clean LIKE '%ข้อผิดพลาด%'        -- error
             OR rsn_clean LIKE '%รอนาน%'            -- long wait
             OR rsn_clean LIKE '%เสียเวลา%'          -- waste of time
             OR rsn_clean LIKE '%แอป%'             -- app
             OR rsn_clean LIKE '%error%'
             OR rsn_clean LIKE '%bug%'
             OR rsn_clean LIKE '%slow%'
             OR rsn_clean LIKE '%crash%'
            THEN 'System / technical issues'

        ---------------------------------------------------------------------
        -- 7. Hard / confusing to use (after loan: กู้ยาก, สมัครสินเชื่อยาก).
        ---------------------------------------------------------------------
        WHEN rsn_clean LIKE '%ยุ่งยาก%'            -- troublesome
             OR rsn_clean LIKE '%ซับซ้อน%'          -- complicated
             OR rsn_clean LIKE '%งง%'              -- confusing
             OR rsn_clean LIKE '%ยาก%'             -- difficult
            THEN 'Product hard / confusing to use'

        ---------------------------------------------------------------------
        -- 8. Dissatisfied with product or service (incl. one-word gripes
        --    and profanity, which is a dissatisfaction signal).
        ---------------------------------------------------------------------
        WHEN rsn_clean LIKE '%ไม่พอใจ%'            -- dissatisfied
             OR rsn_clean LIKE '%ประสบการณ์%'        -- experience
             OR rsn_clean LIKE '%บริการ%'           -- service
             OR rsn_clean LIKE '%แย่%'              -- bad
             OR rsn_clean LIKE '%ห่วย%'             -- lousy
             OR rsn_clean LIKE '%กาก%'              -- rubbish
             OR rsn_clean LIKE '%ไม่ดี%'             -- not good
             OR rsn_clean LIKE '%ไม่ชอบ%'            -- dislike
             OR rsn_clean LIKE '%ไม่โอเค%'           -- not ok
             OR rsn_clean LIKE '%เบื่อ%'             -- fed up
             OR rsn_clean LIKE '%ผิดหวัง%'           -- disappointed
             OR rsn_clean LIKE '%ควย%'              -- profanity
             OR rsn_clean LIKE '%เหี้ย%'             -- profanity
             OR rsn_clean LIKE '%สัส%'              -- profanity
             OR rsn_clean LIKE '%ควาย%'             -- profanity
            THEN 'Dissatisfied with product / service'

        ---------------------------------------------------------------------
        -- 9. Wants a different account number (before "no longer needed" and
        --    "close account": these strings also contain บัญชี / ปิด).
        ---------------------------------------------------------------------
        WHEN rsn_clean LIKE '%เลขบัญชี%'           -- account number
             OR rsn_clean LIKE '%เลขที่บัญชี%'
             OR rsn_clean LIKE '%account number%'
            THEN 'Wants to change account number'

        ---------------------------------------------------------------------
        -- 10. No longer needs the account (largest bucket).
        ---------------------------------------------------------------------
        WHEN rsn_clean LIKE '%ไม่จำเป็น%'           -- not necessary
             OR rsn_clean LIKE '%ไม่ได้ใช้%'          -- not using it
             OR rsn_clean LIKE '%ไม่ใช้%'            -- don't use
             OR rsn_clean LIKE '%เลิกใช้%'           -- stopped using
             OR rsn_clean LIKE '%ไม่มีประโยชน์%'       -- no benefit
             OR rsn_clean LIKE '%no longer need%'
             OR rsn_clean LIKE '%not need%'
             OR rsn_clean LIKE '%do not use%'
             OR rsn_clean LIKE '%don''t use%'
            THEN 'No longer needs the account'

        ---------------------------------------------------------------------
        -- 11. Just wants the account closed, no reason attached.
        ---------------------------------------------------------------------
        WHEN rsn_clean LIKE '%ปิดบัญชี%'
             OR rsn_clean LIKE '%close%account%'
             OR rsn_clean IN ('ปิด','ต้องการปิด','อยากปิด','close')
            THEN 'Wants to close account (no reason stated)'

        ELSE 'Other / uncategorised'
    END AS churn_reason_group
FROM base;


-- =============================================================================
-- Summary roll-up
-- =============================================================================
-- WITH grouped AS ( <paste the SELECT above> )
-- SELECT churn_reason_group,
--        COUNT(*) AS cnt,
--        ROUND(100.0 * COUNT(*) / SUM(COUNT(*)) OVER (), 2) AS pct
-- FROM grouped
-- GROUP BY 1
-- ORDER BY cnt DESC;
--
-- QA: find anything the rules missed, biggest first
-- SELECT rsn, COUNT(*) FROM grouped
-- WHERE churn_reason_group = 'Other / uncategorised'
-- GROUP BY 1 ORDER BY 2 DESC LIMIT 50;
