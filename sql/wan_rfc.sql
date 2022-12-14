SELECT A.NUMBERPRGN AS RFC_REFERENCE,A.REQUEST_DATE AS DATE_OPENED, A.PLANNED_START, A.BRIEF_DESCRIPTION AS MAINTENANCE_DESCRIPTION, A.PLANNED_END, B.LOCATION
FROM ITSM_OWNER.SRC_SC_CM3RM1 A LEFT JOIN ITSM_OWNER.SRC_SC_DEVICEM1 B ON (A.NETWORK_NAME = B.NETWORK_NAME)
WHERE A.ASSIGN_DEPT = 'GBL-NETWORK WAN'
AND LOWER(A.BRIEF_DESCRIPTION) LIKE '%maintenance%'
AND B.LOCATION IN ('SOMERSET RDC','STERLING DATA CENTER','DURHAM RDC','EUROPEAN DATA CENTER','SINGAPORE RDC','BEIJING IDC')
AND A.PLANNED_START > TO_DATE('{}','MM/DD/YYYY HH24:MI:SS')
AND A.STATUS != 'Completed'
ORDER BY A.PLANNED_START ASC
