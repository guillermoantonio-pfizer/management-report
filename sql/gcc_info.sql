SELECT ID AS INCIDENT_NUMBER, IMPACT_START_UTC AS START_TIME, TITLE AS BRIEF_DESCRIPTION, STATUS
FROM ITSM_OWNER.SRC_GCC_EVENTS 
WHERE SEVERITY = 'Informational' and STATUS = 'Active' ORDER BY SUBMITTED_ON_UTC DESC