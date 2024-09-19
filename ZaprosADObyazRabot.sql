select
   DNum."Num",
   CM."JudicalUid",
   (select first(1) pp."PersonName" from "ParticipantPerson" pp where pp."CaseId"=CM."OID" and pp."Kind"=31) as "Names",
   ti."Name",
   AP."OID",
   BP."SimplePenalty",
   BaseCase."NormativeLegalDocument",
   BaseCase."APCodesList",
Doc."CreateDate",
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
left join "AdmPunishment" AP on AP."Case" = BaseCase."OID"
left join "BasePunishment" BP on BP."OID" = AP."OID"
where CM."ProcKind"=1 and DNum."Num" like '%:paramYear' and fd."DecisionDate" is not null and BP."SimplePenalty" = 38
order by DNum."IntNum"