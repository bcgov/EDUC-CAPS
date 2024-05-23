/*This query is used in Funding Distribution Report. Pulling data from draw request, actual draws, batch, and project for completed batch. */
SELECT 
p.caps_projectnumber as ProjectNumber
,

FROM caps_actualdraw a
JOIN caps_drawrequest d on a.caps_originatingdrawrequest = d.caps_drawrequestid
JOIN caps_batch b on d.caps_batch = b.caps_batchid
JOIN caps_projectcode c on c.caps_projectcodeid = b.caps_projectcode
JOIN caps_projecttracker p on d.caps_project = p.caps_projecttrackerid
JOIN edu_schooldistrict s on p.caps_schooldistrict = s.edu_schooldistrictid

WHERE 
d.statuscode = 200870003 AND
b.statuscode = 200870003

GROUP BY p.caps_projectnumber