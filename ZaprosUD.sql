select
   DNum."Num",
   (select "Names" from "GetCaseParticipantsNames"( CM."OID", 27, 0 )),
   (select MAIN_ARTICLES from GETMAINCRIMARTICLE(CM."OID")),
   CM."OID" as "OID",
   CM."JudicalUid",
   (select "INNUM" from getinnum(CM.oid)),
   (select dn."IntNum" from "DocNum" dn where dn."Id"=Jic."DocNum"),
   Doc."CreateDate",
   DNum."IntNum",
   BaseCase."CanIteratively",
   ti."Id",
   ti."Name",
   tg."ExecutionDate",
   CM."PlanDate",
   CM."FinishingDate",
   CM."OldNum",
   Doc."Comment",
   CM."IsRawMaterial",
   fd."DecisionDate",
   fd."DecisionYear",
   pub_data."Data",
   (select result from "GetCrimCaseIsOutOfTime"(CM."OID")),
   CM."ResultOfTheReview", 
   CM."SittingDate",
   Reason."DocumentType",
   CM."CaseMaterialType"
from "CaseMaterial" CM
left outer join "Document" Reason on CM."DocumentReason"=Reason."OID"
left outer join "Document" Doc on CM."OID"=Doc."OID"
left outer join "JicDocument" Jic on Reason."Jic"=Jic."OID"
left outer join "DocNum" DNum on CM."DocNum"=DNum."Id"
left outer join "BaseCourtCase" BaseCase on CM."OID"=BaseCase."OID"
join "TaskGerm" tg on tg."OID"=CM."CurrentTask"
join "TaskInfo" ti on ti."Id"=tg."Info"
left join "CrimFinalDecisionsData" fd on fd."CaseId" = cm.oid
left join "CasePublicationData" pub_data on pub_data."CaseId" = cm.oid
where CM."ProcKind"=3 and DNum."Num" like '%:paramYear'
order by DNum."IntNum"