select
  DNum."Num", 
 (select first(1) pp."PersonName" from "ParticipantPerson" pp where pp."CaseId"=CM."OID" and pp."Kind"=31) as "Names",
   a1."Code" as "MainArticle22", 
   CM."OID",
   CM."JudicalUid",
   DNum."Num",
   (select dn."Num" from "DocNum" dn where dn."Id"=Jic."DocNum"),
   (select dn."IntNum" from "DocNum" dn where dn."Id"=Jic."DocNum"),
   Doc."CreateDate",
   DNum."IntNum",
   (select first(1) pp."PersonName" from "ParticipantPerson" pp where pp."CaseId"=CM."OID" and pp."Kind"=31) as "Names",
  a1."Code" as "MainArticle22", 
  BaseCase."NormativeLegalDocument",
   BaseCase."APCodesList",
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
   (select result from "GetAdmCaseIsOutOfTime"(CM."OID")),
   cm."ResultOfTheReview",
   cm."SittingDate",
   CM."CaseMaterialType",
   agency."Name"
from "CaseMaterial" CM
left outer join "Document" Reason on CM."DocumentReason"=Reason."OID"
left outer join "AdmArticle" a1 on CM."OID"=a1."AdmCase"
left outer join "Document" Doc on CM."OID"=Doc."OID"
left outer join "JicDocument" Jic on Reason."Jic"=Jic."OID"
left outer join "DocNum" DNum on CM."DocNum"=DNum."Id"
left outer join "BaseCourtCase" BaseCase on CM."OID"=BaseCase."OID"
join "TaskGerm" tg on tg."OID"=CM."CurrentTask"
join "TaskInfo" ti on ti."Id"=tg."Info"
left join "AdmFinalDecisionsData" fd on fd."CaseId" = cm.oid
left join "CasePublicationData" pub_data on pub_data."CaseId" = cm.oid
left join "AdmDocProtocol" protocol on protocol."OID" = Jic."Document"
left join "Agency" agency on agency."OID" = protocol."ProtocolAuthorAgency"
where CM."ProcKind"=1 and DNum."Num" like '%:paramYear'
order by DNum."IntNum"