select
  case
  when CM."CaseMaterialType" = 3 then
        (select "CASESNUMS" from GETCASENUMWITHMERGEDCASESNUMS(CM.oid))
      else
       DNum."Num"
   end as "Num",
(select "Type" from "Participant" PC where ( PC."OID"=CM."OID")),
   (select "Names" from GETMERGEDCASEPARTNAMES( CM."OID", CM."CaseMaterialType", '7, 10' )) as "Names", 
(select "Names" from GETMERGEDCASEPARTNAMES( CM."OID", CM."CaseMaterialType", '5, 8, 26' )) as "Names1",
   CM."OID",
   CM."JudicalUid",
   (select "INNUM" from getinnum(CM.oid)),
   (select dn."IntNum" from "DocNum" dn where dn."Id"=Jic."DocNum"),
   Doc."CreateDate",
   case
      when CM."CaseMaterialType" = 3 then
        (select "CASESNUMS" from GETCASENUMWITHMERGEDCASESNUMS(CM.oid))
      else
        DNum."Num"
   end,
   DNum."IntNum",
   CC."CanIteratively",
   CM."CaseMaterialType",
   ti."Id",
   ti."Name",
   tg."ExecutionDate",
   CM."PlanDate",
   CM."FinishingDate",
   CM."OldNum",
   Doc."Comment",
   DNumRaw."Num",
   CM."IsRawMaterial",
   merge_nums."Num",
   merge_nums."OldNum",
   merge_nums."RegNum",
   fd."DecisionDate",
   fd."DecisionYear",
   pub_data."Data",
   oft."IsOutOfTime",
   CC."CategoryNamePersistent",
   CM."ResultOfTheReview",
   CM."SittingDate"
from "CaseMaterial" CM
left join "Document" Reason on CM."DocumentReason"=Reason."OID"
left join "Document" Doc on CM."OID"=Doc."OID"
left join "JicDocument" Jic on Reason."Jic"=Jic."OID"
left join "DocNum" DNum on CM."DocNum"=DNum."Id"
left join "CivilCase" CC on CM."OID"=CC."OID"
left join "DocNum" DNumRaw on CC."DocNumRaw"=DNumRaw."Id"
left join "GetMergedCaseNum"(CM.oid, CM."CaseMaterialType") merge_nums on 1=1
join "TaskGerm" tg on tg."OID"=CM."CurrentTask"
join "TaskInfo" ti on ti."Id"=tg."Info"
left join "CivilFinalDecisionsData" fd on fd."CaseId" = cm.oid
left join "CasePublicationData" pub_data on pub_data."CaseId" = cm.oid
join "CivilCaseIsOutOfTime" oft on oft."CaseId" = CM.oid
where CM."ProcKind"=7 and DNum."Num" like '%:paramYear'
order by DNum."IntNum"