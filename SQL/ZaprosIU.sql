select
   DNum."Num",
case
    when EType."Category" is null then
         ' '
    when EType."Category" = 1 then
           (select "Names" from "GetCaseParticipantsNames"( CM."OID", 35, 0 ))
    when EType."Category" = 2 then
           (select "Names" from "GetCaseParticipantsNames"( CM."OID", 31, 0 ))
    when EType."Category" = 3 then
           (select "Names" from "GetCaseParticipantsNames"( CM."OID", 5, 7 ))
  end as "NAMES3",
  (select "Names" from GETMERGEDCASEPARTNAMES( CM."OID", CM."CaseMaterialType", '5, 8, 26' )),
   (select "Names" from GETMERGEDCASEPARTNAMES( CM."OID", CM."CaseMaterialType", '7, 10' )),
      ti."Name",
   CM."OID",
   CM."JudicalUid",
   (select dn."Num" from "DocNum" dn where dn."Id"=Jic."DocNum"),
   (select dn."IntNum" from "DocNum" dn where dn."Id"=Jic."DocNum"),
   Doc."CreateDate",
   DNum."IntNum",
   Stage."Name",
   ti."Id",
   tg."ExecutionDate",
   CM."PlanDate",
   CM."FinishingDate",
   CM."SittingDate",
   CM."ResultOfTheReview",
   CM."OldNum",
   Doc."Comment",
   CM."IsRawMaterial",
   fd."DecisionDate",
   fd."DecisionYear",
   pub_data."Data",
	0,
	EMT."Name"
from "CaseMaterial" CM
left outer join "Document" Reason on CM."DocumentReason"=Reason."OID"
left outer join "Document" Doc on CM."OID"=Doc."OID"
left outer join "JicDocument" Jic on Reason."Jic"=Jic."OID"
left outer join "DocNum" DNum on CM."DocNum"=DNum."Id"
left outer join "Stage" Stage on CM."CurrentStage"=Stage."Id"
left outer join "BaseCourtCase" BaseCase on CM."OID"=BaseCase."OID"
join "ExecMaterial" EM on CM."OID"=EM."OID"
join "TaskGerm" tg on tg."OID"=CM."CurrentTask"
join "TaskInfo" ti on ti."Id"=tg."Info"
left outer join "ExecMaterialType" EType on EM."MaterialType" = EType."Id"
left join "ExecMatFinalDecisionsData" fd on fd."CaseId" = cm.oid
left join "CasePublicationData" pub_data on pub_data."CaseId" = cm.oid
left join "ExecMaterialType" EMT on EMT."Id" = EM."MaterialType"
where (CM."ProcKind"=6 or "ProcKind"=8) and (EType."Category"=1) and DNum."Num" like '%:paramYear'
order by DNum."IntNum"