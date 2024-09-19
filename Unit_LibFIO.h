#include <DateUtils.hpp>
AnsiString DS (AnsiString data)
        { AnsiString m;
        if (!data.IsEmpty())
                {switch (StrToInt(data.SubString(4,2)))
                        {
                        case 1: m = "€нвар€"; break;
                        case 2: m = "феврал€"; break;
                        case 3: m = "марта"; break;
                        case 4: m = "апрел€"; break;
                        case 5: m = "ма€"; break;
                        case 6: m = "июн€"; break;
                        case 7: m = "июл€"; break;
                        case 8: m = "августа"; break;
                        case 9: m = "сент€бр€"; break;
                        case 10: m = "окт€бр€"; break;
                        case 11: m = "но€бр€"; break;
                        case 12: m = "декабр€"; break;
                        };
                return data.SubString(1,2)+" "+m+" "+data.SubString(7,4)+" года";
                }
                else
                return "";
        };

TDateTime StringToDT (AnsiString sdate)
        {Word day, month, year;
        TDateTime res;
        if (!sdate.IsEmpty())
                {day = StrToInt(sdate.SubString(1,2));
                month= StrToInt(sdate.SubString(4,2));
                year = StrToInt(sdate.SubString(7,4));
                res=RecodeDate(res, year, month, day);
                res=RecodeTime(res, 23, 0, 0, 0);
                //AnsiString ss;
                //ss="day="+IntToStr(day)+" month="+IntToStr(month)+" year="+IntToStr(year);
                //MessageBox(NULL, ss.c_str(), "Amirs", MB_OK | MB_ICONERROR);
                };
        return res;
        };

AnsiString DataToString (TDateTime datetime)
        {AnsiString m;
         Word year, month, day;
         datetime.DecodeDate(&year,&month,&day);
         switch (month)
              {
              case 1: m = "€нвар€"; break;
              case 2: m = "феврал€"; break;
              case 3: m = "марта"; break;
              case 4: m = "апрел€"; break;
              case 5: m = "ма€"; break;
              case 6: m = "июн€"; break;
              case 7: m = "июл€"; break;
              case 8: m = "августа"; break;
              case 9: m = "сент€бр€"; break;
              case 10: m = "окт€бр€"; break;
              case 11: m = "но€бр€"; break;
              case 12: m = "декабр€"; break;
              };
         m = IntToStr(day)+" "+m+" "+IntToStr(year)+" года";
         return m;
         };

AnsiString TS (AnsiString time)
        { AnsiString h, m;
        if (!time.IsEmpty())
                {
                int pos=time.Pos(":"); h=time.SubString(1,pos-1);
                time = time.SubString(pos+1, time.Length());
                pos=time.Pos(":");
                if (pos==0) m=time; else m=time.SubString(1,pos-1);
                return h+" час. "+m+" мин.";
                }
                else
                return "";
        };

class FIO {
        public:
        int sex;
        AnsiString fio, fiodat, fiorod, fiovin, fiotvo,
        fiolit, fiolitdat, fiolitrod, fiolitvin, fiolittvo,
        fam, famdat, famrod, famvin, famtvo,
        name, namedat, namerod, namevin, nametvo,
        otch, otchdat, otchrod, otchvin, otchtvo;
        FIO (AnsiString fio1, int sex1);
        };

FIO::FIO (AnsiString fio1, int sex1)
        {fio=fio1;
         AnsiString lc=fio1.LowerCase();
         int posfam=lc.Pos(" "); fam=lc.SubString(1,posfam-1);
         lc=lc.SubString(posfam+1, lc.Length()).TrimLeft();
         int posname=lc.Pos(" "); name=lc.SubString(1,posname-1);
         otch=lc.SubString(posname+1, lc.Length()).TrimLeft();
         otch=otch.TrimRight();
         //Boolean sex;
         sex=sex1;
         if (sex == 0) if (otch.SubString(otch.Length(), 1)=="ч") sex = 1; else sex=2;
         if (otch.SubString(otch.Length()-2, 3)=="глы") sex = 1;
        //‘амилии
        AnsiString ss,e2,n1,n2; ss = fam.SubString(fam.Length(), 1); e2 = fam.SubString(fam.Length()-1, 2);
                      n1 = name.SubString(name.Length(), 1); n2 = name.SubString(name.Length()-1, 2);
        if (sex==1) {if (ss =="о"|ss=="и"|ss=="€"|ss=="а"|ss=="е" |ss=="х") {famdat=fam;famrod=fam;famvin=fam;famtvo=fam;}
           else
             if (ss=="й") {famdat=fam.SubString(0, fam.Length()-2)+"ому";
                         famrod=fam.SubString(0, fam.Length()-2)+"ого";
                         famvin=famrod;
                         famtvo=fam.SubString(0, fam.Length()-2)+"им";}
             else  {famdat=fam+"у"; famrod=fam+"а"; famvin=fam+"а"; famtvo=fam+"ым";}
           }
        else      {if (e2=="ва"|e2=="на") {famdat=fam.SubString(0, fam.Length()-1)+"ой";
                           famrod=fam.SubString(0, fam.Length()-1)+"ой";
                           famvin=fam.SubString(0, fam.Length()-1)+"у";
                           famtvo=fam.SubString(0, fam.Length()-1)+"ой";}
              else
              if (e2=="а€") {famdat=fam.SubString(0, fam.Length()-2)+"ой";
                           famrod=fam.SubString(0, fam.Length()-2)+"ой";
                           famvin=fam.SubString(0, fam.Length()-2)+"ую";
                           famtvo=fam.SubString(0, fam.Length()-2)+"ой";}
              else {famdat=fam;famrod=fam;famvin=fam;famtvo=fam;};
           }
        //имена
        if (sex==1) {if (n1 =="й"|n1=="ь") {namedat=name.SubString(0, name.Length()-1)+"ю";
                                  namerod=name.SubString(0, name.Length()-1)+"€";
                                  namevin=name.SubString(0, name.Length()-1)+"€";
                                  nametvo=name.SubString(0, name.Length()-1)+"ем";}
                  else
                  if (n2=="ев") {namedat=name.SubString(0, name.Length()-2)+"ьву";
                                  namerod=name.SubString(0, name.Length()-2)+"ьва";
                                  namevin=name.SubString(0, name.Length()-2)+"ьва";
                                  nametvo=name.SubString(0, name.Length()-2)+"ьвом";}
                  else
                  if (n2=="ел") {namedat=name.SubString(0, name.Length()-2)+"лу";
                                  namerod=name.SubString(0, name.Length()-2)+"ла";
                                  namevin=name.SubString(0, name.Length()-2)+"ла";
                                  nametvo=name.SubString(0, name.Length()-2)+"лом";}
                  else
                  if (n1 =="а") {namedat=name.SubString(0, name.Length()-1)+"е";
                                  namerod=name.SubString(0, name.Length()-1)+"ы";
                                  namevin=name.SubString(0, name.Length()-1)+"у";
                                  nametvo=name.SubString(0, name.Length()-1)+"ой";}
                  else  {namedat=name+"у"; namerod=name+"а"; namevin=name+"а"; nametvo=name+"ом";};
                  }
        else
                {if (n2=="га" | n2=="ка") {namedat=name.SubString(0, name.Length()-1)+"е";
                                  namerod=name.SubString(0, name.Length()-1)+"и";
                                  namevin=name.SubString(0, name.Length()-1)+"у";
                                  nametvo=name.SubString(0, name.Length()-1)+"ой";}
                 else
                 if (n1 =="а") {namedat=name.SubString(0, name.Length()-1)+"е";
                                  namerod=name.SubString(0, name.Length()-1)+"ы";
                                  namevin=name.SubString(0, name.Length()-1)+"у";
                                  nametvo=name.SubString(0, name.Length()-1)+"ой";}
                 else
                 if (n2=="ь€"|n2=="о€")
                                {namedat=name.SubString(0, name.Length()-1)+"е";
                                  namerod=name.SubString(0, name.Length()-1)+"и";
                                  namevin=name.SubString(0, name.Length()-1)+"ю";
                                  nametvo=name.SubString(0, name.Length()-1)+"ей";}
                 else
                 if (n2=="и€") {namedat=name.SubString(0, name.Length()-1)+"и";
                                  namerod=name.SubString(0, name.Length()-1)+"и";
                                  namevin=name.SubString(0, name.Length()-1)+"ю";
                                  nametvo=name.SubString(0, name.Length()-1)+"ей";}
                 else

                 if (n1=="ь") {namedat=name.SubString(0, name.Length()-1)+"и";
                                  namerod=name.SubString(0, name.Length()-1)+"и";
                                  namevin=name.SubString(0, name.Length()-1)+"ь";
                                  nametvo=name.SubString(0, name.Length()-1)+"ью";}
                 else  {namedat=name; namerod=name; namevin=name; nametvo=name;};
        }
        //отчества
        if (sex==1) if (otch.SubString(otch.Length()-2, 3)=="глы" | otch.SubString(otch.Length()-2, 3)=="гли") {otchdat=otch; otchrod=otch; otchvin=otch; otchtvo=otch;}
                else
                {otchdat=otch+"у"; otchrod=otch+"а"; otchvin=otch+"а"; otchtvo=otch+"ем";}
        else
           if (otch.SubString(otch.Length()-2, 3)=="ызы") {otchdat=otch; otchrod=otch; otchvin=otch; otchtvo=otch;}
            else
               {otchdat=otch.SubString(0, otch.Length()-1)+"е";
                                  otchrod=otch.SubString(0, otch.Length()-1)+"ы";
                                  otchvin=otch.SubString(0, otch.Length()-1)+"у";
                                  otchtvo=otch.SubString(0, otch.Length()-1)+"ой";}
        famdat=famdat.SubString(1,1).UpperCase()+famdat.SubString(2,famdat.Length());
        namedat=namedat.SubString(1,1).UpperCase()+namedat.SubString(2,namedat.Length());
        otchdat=otchdat.SubString(1,1).UpperCase()+otchdat.SubString(2,otchdat.Length());
        famrod=famrod.SubString(1,1).UpperCase()+famrod.SubString(2,famrod.Length());
        namerod=namerod.SubString(1,1).UpperCase()+namerod.SubString(2,namerod.Length());
        otchrod=otchrod.SubString(1,1).UpperCase()+otchrod.SubString(2,otchrod.Length());
        famvin=famvin.SubString(1,1).UpperCase()+famvin.SubString(2,famvin.Length());
        namevin=namevin.SubString(1,1).UpperCase()+namevin.SubString(2,namevin.Length());
        otchvin=otchvin.SubString(1,1).UpperCase()+otchvin.SubString(2,otchvin.Length());
        famtvo=famtvo.SubString(1,1).UpperCase()+famtvo.SubString(2,famtvo.Length());
        nametvo=nametvo.SubString(1,1).UpperCase()+nametvo.SubString(2,nametvo.Length());
        otchtvo=otchtvo.SubString(1,1).UpperCase()+otchtvo.SubString(2,otchtvo.Length());
        fiodat=famdat+" "+namedat+" "+otchdat;
        fiorod=famrod+" "+namerod+" "+otchrod;
        fiovin=famvin+" "+namevin+" "+otchvin;
        fiotvo=famtvo+" "+nametvo+" "+otchtvo;
        fam=fam.SubString(1,1).UpperCase()+fam.SubString(2,fam.Length());
        fiolit   =fam+" "+name.SubString(1,1).UpperCase()+"."+otch.SubString(1,1).UpperCase()+".";
        fiolitdat=famdat+" "+name.SubString(1,1).UpperCase()+"."+otch.SubString(1,1).UpperCase()+".";
        fiolitrod=famrod+" "+name.SubString(1,1).UpperCase()+"."+otch.SubString(1,1).UpperCase()+".";
        fiolitvin=famvin+" "+name.SubString(1,1).UpperCase()+"."+otch.SubString(1,1).UpperCase()+".";
        fiolittvo=famtvo+" "+name.SubString(1,1).UpperCase()+"."+otch.SubString(1,1).UpperCase()+".";
        };

