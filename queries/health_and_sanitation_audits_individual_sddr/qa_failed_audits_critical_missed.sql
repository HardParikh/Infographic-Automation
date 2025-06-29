-- CRITICAL MISSED QUESTIONS FROM FAILED QA AUDITS
WITH  CTE_HS_UNIT_SCORES_BELOW_90 AS ( 
SELECT RESPONSE_ID, UNIT_NAME, UNIT_SAP_NUMBER, UNIT_INDUSTRY, HIERARCHY_SCOTT_DAVIS_DIRECT_REPORTS AS SCOTT_DAVIS_DIRECT_REPORTS, 'H&S (QA) Unit Audit Score' AS METRIC, ROUND(AVG(CATEGORY_SCORE),2) AS STATISTIC,
CASE WHEN STATISTIC < 90 THEN 'FAIL'
ELSE 'PASS'
END AS "PASS/FAIL"
FROM FLIK_ANALYTICS.CURIOSITY.SURVEYS_COMBINED
WHERE (UNIT_BUSINESS_PORTFOLIO <> 'Test' or UNIT_BUSINESS_PORTFOLIO is null)  -- excludes 'test' portfolios but includes 'null' portfolios
AND SCOTT_DAVIS_DIRECT_REPORTS = '{SDDR_NAME}'
AND SURVEY_ID = 'SV_cBggdZ7uWwWjmgl'
AND AUDIT_DATE BETWEEN DATEADD('month', -1, DATE_TRUNC('month', CURRENT_DATE()))
AND DATE_TRUNC('month', CURRENT_DATE())
AND UNIT_SAP_NUMBER NOT IN ('1', '4440', '33940') -- removes test B&I account & Insights test accounts
AND RESPONSE_VALID = 'Yes'
AND CATEGORY = 'Total Score'
AND CATEGORY_SCORE_ID = 'PercentScore'
AND ISGOLDLISTACTIVE = TRUE
GROUP BY RESPONSE_ID, UNIT_NAME, UNIT_SAP_NUMBER, UNIT_INDUSTRY, SCOTT_DAVIS_DIRECT_REPORTS
HAVING STATISTIC < 90
)
SELECT HIERARCHY_SCOTT_DAVIS_DIRECT_REPORTS AS SCOTT_DAVIS_DIRECT_REPORTS,UNIT_SAP_NUMBER, UNIT_NAME, RESPONSE_ID,YEAR(AUDIT_DATE) AS YR, MONTH(AUDIT_DATE) AS MO, DAY(AUDIT_DATE) AS DAY,
 QUESTION_TEXT, FOLLOW_UP_QUESTION_RESPONSE
FROM FLIK_ANALYTICS.CURIOSITY.SURVEYS_COMBINED
WHERE (UNIT_BUSINESS_PORTFOLIO <> 'Test' or UNIT_BUSINESS_PORTFOLIO is null)  -- excludes 'test' portfolios but includes 'null' portfolios
AND SURVEY_ID = 'SV_cBggdZ7uWwWjmgl'
AND AUDIT_DATE BETWEEN DATEADD('month', -1, DATE_TRUNC('month', CURRENT_DATE()))
AND DATE_TRUNC('month', CURRENT_DATE()) -- Need window function
AND UNIT_SAP_NUMBER NOT IN ('1', '4440', '33940') -- removes test B&I account & Insights test accounts
AND RESPONSE_ID IN (SELECT RESPONSE_ID FROM CTE_HS_UNIT_SCORES_BELOW_90)
AND RESPONSE_VALID = 'Yes'
AND DATATYPE = 'Responses'
AND QUESTION_TYPE = 'MC'
AND ISGOLDLISTACTIVE = TRUE
AND FOLLOW_UP_QUESTION_RESPONSE LIKE '%(CRITICAL)%'
AND QUESTION_TEXT <> 'Reason for selecting NO'
ORDER BY HIERARCHY_SCOTT_DAVIS_DIRECT_REPORTS,YR DESC, MO DESC, DAY DESC, RESPONSE_ID, QUESTION_TEXT
;