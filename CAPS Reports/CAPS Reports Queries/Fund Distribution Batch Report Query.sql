/*This query is used in Funding Distribution Report. Pulling data from draw request, actual draws, batch, and project for completed batch. */
SELECT 
b.caps_paymentdescription as PaymentDescription
,b.caps_name as Batch
,c.caps_name as ProjectCode
,c.caps_project as ProjectCodeProject
,c.caps_responsibilitycentre as RespCentre
,c.caps_serviceline as ServiceLine
,c.caps_stob as Stob
,b.caps_signedbyname as ExpenseAuthority
,cast(b.caps_signedon as DATE) as SignedOn
--,d.caps_projectname as Project
,s.edu_commonname as SDName
,s.edu_number as SDNumber
,sum(a.caps_amount) as Amount
,CAST(SUBSTRING(s.edu_number, 3, LEN(s.edu_number)-2) AS INT) AS SDRevised
,CONCAT(
        CAST(SUBSTRING(s.edu_number, 3, LEN(s.edu_number)-2) AS INT), 
        ',', 
        CAST(sum(a.caps_amount) AS INT)
    ) AS CDS
,    SUM(SUM(a.caps_amount)) OVER (PARTITION BY s.edu_number) AS TotalPerSDNumber
,    SUM(SUM(a.caps_amount)) OVER (PARTITION BY b.caps_name) AS TotalPerBatch
FROM caps_actualdraw a
JOIN caps_drawrequest d on a.caps_originatingdrawrequest = d.caps_drawrequestid
JOIN caps_batch b on d.caps_batch = b.caps_batchid
JOIN caps_projectcode c on c.caps_projectcodeid = b.caps_projectcode
JOIN caps_projecttracker p on d.caps_project = p.caps_projecttrackerid
JOIN edu_schooldistrict s on p.caps_schooldistrict = s.edu_schooldistrictid

WHERE 
d.statuscode = 200870003 AND
b.statuscode = 200870003 
AND b.caps_batchid = @Batch

GROUP BY 
b.caps_paymentdescription
,b.caps_name
,c.caps_name
,c.caps_project
,c.caps_responsibilitycentre
,c.caps_serviceline
,c.caps_stob
,b.caps_signedbyname
,b.caps_signedon
--,d.caps_projectname
,s.edu_commonname
,s.edu_number