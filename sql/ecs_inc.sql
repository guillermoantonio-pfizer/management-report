SELECT NUMBERPRGN AS INCIDENT#,OPEN_TIME AS DATE_OPENED, BRIEF_DESCRIPTION AS DESCRIPTION, PRIORITY, UPDATE_ACTION AS CURRENT_STATUS, PROBLEM_STATUS AS STATUS, LOCATION
FROM ITSM_OWNER.ITSM_PROBSUMMARYM1_V WHERE ASSIGNMENT = 'GBL-NETWORK ECS' AND PRIORITY IN ('1 - Critical','2 - High')
AND OPEN_TIME >= TO_DATE('{}','MM/DD/YYYY HH24:MI:SS') AND OPEN_TIME <= TO_DATE('{}','MM/DD/YYYY HH24:MI:SS')
AND (PROBLEM_STATUS != 'Closed' AND PROBLEM_STATUS != 'Resolved')
ORDER BY OPEN_TIME ASC
