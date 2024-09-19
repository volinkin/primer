#include <DateUtils.hpp>
AnsiString DS (AnsiString data)
        { AnsiString m;
        if (!data.IsEmpty())
                {switch (StrToInt(data.SubString(4,2)))
                        {
                        case 1: m = "������"; break;
                        case 2: m = "�������"; break;
                        case 3: m = "�����"; break;
                        case 4: m = "������"; break;
                        case 5: m = "���"; break;
                        case 6: m = "����"; break;
                        case 7: m = "����"; break;
                        case 8: m = "�������"; break;
                        case 9: m = "��������"; break;
                        case 10: m = "�������"; break;
                        case 11: m = "������"; break;
                        case 12: m = "�������"; break;
                        };
                return data.SubString(1,2)+" "+m+" "+data.SubString(7,4)+" ����";
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
              case 1: m = "������"; break;
              case 2: m = "�������"; break;
              case 3: m = "�����"; break;
              case 4: m = "������"; break;
              case 5: m = "���"; break;
              case 6: m = "����"; break;
              case 7: m = "����"; break;
              case 8: m = "�������"; break;
              case 9: m = "��������"; break;
              case 10: m = "�������"; break;
              case 11: m = "������"; break;
              case 12: m = "�������"; break;
              };
         m = IntToStr(day)+" "+m+" "+IntToStr(year)+" ����";
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
                return h+" ���. "+m+" ���.";
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
         if (sex == 0) if (otch.SubString(otch.Length(), 1)=="�") sex = 1; else sex=2;
         if (otch.SubString(otch.Length()-2, 3)=="���") sex = 1;
        //�������
        AnsiString ss,e2,n1,n2; ss = fam.SubString(fam.Length(), 1); e2 = fam.SubString(fam.Length()-1, 2);
                      n1 = name.SubString(name.Length(), 1); n2 = name.SubString(name.Length()-1, 2);
        if (sex==1) {if (ss =="�"|ss=="�"|ss=="�"|ss=="�"|ss=="�" |ss=="�") {famdat=fam;famrod=fam;famvin=fam;famtvo=fam;}
           else
             if (ss=="�") {famdat=fam.SubString(0, fam.Length()-2)+"���";
                         famrod=fam.SubString(0, fam.Length()-2)+"���";
                         famvin=famrod;
                         famtvo=fam.SubString(0, fam.Length()-2)+"��";}
             else  {famdat=fam+"�"; famrod=fam+"�"; famvin=fam+"�"; famtvo=fam+"��";}
           }
        else      {if (e2=="��"|e2=="��") {famdat=fam.SubString(0, fam.Length()-1)+"��";
                           famrod=fam.SubString(0, fam.Length()-1)+"��";
                           famvin=fam.SubString(0, fam.Length()-1)+"�";
                           famtvo=fam.SubString(0, fam.Length()-1)+"��";}
              else
              if (e2=="��") {famdat=fam.SubString(0, fam.Length()-2)+"��";
                           famrod=fam.SubString(0, fam.Length()-2)+"��";
                           famvin=fam.SubString(0, fam.Length()-2)+"��";
                           famtvo=fam.SubString(0, fam.Length()-2)+"��";}
              else {famdat=fam;famrod=fam;famvin=fam;famtvo=fam;};
           }
        //�����
        if (sex==1) {if (n1 =="�"|n1=="�") {namedat=name.SubString(0, name.Length()-1)+"�";
                                  namerod=name.SubString(0, name.Length()-1)+"�";
                                  namevin=name.SubString(0, name.Length()-1)+"�";
                                  nametvo=name.SubString(0, name.Length()-1)+"��";}
                  else
                  if (n2=="��") {namedat=name.SubString(0, name.Length()-2)+"���";
                                  namerod=name.SubString(0, name.Length()-2)+"���";
                                  namevin=name.SubString(0, name.Length()-2)+"���";
                                  nametvo=name.SubString(0, name.Length()-2)+"����";}
                  else
                  if (n2=="��") {namedat=name.SubString(0, name.Length()-2)+"��";
                                  namerod=name.SubString(0, name.Length()-2)+"��";
                                  namevin=name.SubString(0, name.Length()-2)+"��";
                                  nametvo=name.SubString(0, name.Length()-2)+"���";}
                  else
                  if (n1 =="�") {namedat=name.SubString(0, name.Length()-1)+"�";
                                  namerod=name.SubString(0, name.Length()-1)+"�";
                                  namevin=name.SubString(0, name.Length()-1)+"�";
                                  nametvo=name.SubString(0, name.Length()-1)+"��";}
                  else  {namedat=name+"�"; namerod=name+"�"; namevin=name+"�"; nametvo=name+"��";};
                  }
        else
                {if (n2=="��" | n2=="��") {namedat=name.SubString(0, name.Length()-1)+"�";
                                  namerod=name.SubString(0, name.Length()-1)+"�";
                                  namevin=name.SubString(0, name.Length()-1)+"�";
                                  nametvo=name.SubString(0, name.Length()-1)+"��";}
                 else
                 if (n1 =="�") {namedat=name.SubString(0, name.Length()-1)+"�";
                                  namerod=name.SubString(0, name.Length()-1)+"�";
                                  namevin=name.SubString(0, name.Length()-1)+"�";
                                  nametvo=name.SubString(0, name.Length()-1)+"��";}
                 else
                 if (n2=="��"|n2=="��")
                                {namedat=name.SubString(0, name.Length()-1)+"�";
                                  namerod=name.SubString(0, name.Length()-1)+"�";
                                  namevin=name.SubString(0, name.Length()-1)+"�";
                                  nametvo=name.SubString(0, name.Length()-1)+"��";}
                 else
                 if (n2=="��") {namedat=name.SubString(0, name.Length()-1)+"�";
                                  namerod=name.SubString(0, name.Length()-1)+"�";
                                  namevin=name.SubString(0, name.Length()-1)+"�";
                                  nametvo=name.SubString(0, name.Length()-1)+"��";}
                 else

                 if (n1=="�") {namedat=name.SubString(0, name.Length()-1)+"�";
                                  namerod=name.SubString(0, name.Length()-1)+"�";
                                  namevin=name.SubString(0, name.Length()-1)+"�";
                                  nametvo=name.SubString(0, name.Length()-1)+"��";}
                 else  {namedat=name; namerod=name; namevin=name; nametvo=name;};
        }
        //��������
        if (sex==1) if (otch.SubString(otch.Length()-2, 3)=="���" | otch.SubString(otch.Length()-2, 3)=="���") {otchdat=otch; otchrod=otch; otchvin=otch; otchtvo=otch;}
                else
                {otchdat=otch+"�"; otchrod=otch+"�"; otchvin=otch+"�"; otchtvo=otch+"��";}
        else
           if (otch.SubString(otch.Length()-2, 3)=="���") {otchdat=otch; otchrod=otch; otchvin=otch; otchtvo=otch;}
            else
               {otchdat=otch.SubString(0, otch.Length()-1)+"�";
                                  otchrod=otch.SubString(0, otch.Length()-1)+"�";
                                  otchvin=otch.SubString(0, otch.Length()-1)+"�";
                                  otchtvo=otch.SubString(0, otch.Length()-1)+"��";}
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

