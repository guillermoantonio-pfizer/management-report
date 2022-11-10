SELECT NUMBERPRGN AS RFC_REFERENCE,REQUEST_DATE AS DATE_OPENED, BRIEF_DESCRIPTION AS DESCRIPTION, PLANNED_START AS DATE_OF_EVENT
FROM ITSM_OWNER.SRC_SC_CM3RM1
WHERE ASSIGN_DEPT = 'GBL-NETWORK UC'
AND (LOWER(BRIEF_DESCRIPTION) LIKE '%migration%' OR LOWER(BRIEF_DESCRIPTION) LIKE '%upgrade%' OR LOWER(BRIEF_DESCRIPTION) LIKE '%replacement%')
AND PLANNED_START > TO_DATE('{}','MM/DD/YYYY HH24:MI:SS')
AND STATUS != 'Completed'
ORDER BY REQUEST_DATE ASC
