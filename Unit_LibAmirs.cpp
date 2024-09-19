#include "Unit_LibAmirs.h"
//#include <vcl.h>
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include <DB.hpp>
#include <IBDatabase.hpp>
#include <IBQuery.hpp>
#include <IBTable.hpp>
#include <ComCtrls.hpp>
//#include "CSPIN.h"
#include <DBCtrls.hpp>
#include <DBGrids.hpp>
#include <ExtCtrls.hpp>
#include <Grids.hpp>
#include "Excel_2K_SRVR.h"
#include "Word_2K_SRVR.h"
#include <OleServer.hpp>
#include <IBCustomDataSet.hpp>

TAmirs::TAmirs(TForm * Form)
        {
        Errors=0;
        slErrors = new TStringList;
        Form1=Form;
        //NameApp=NameApp1;
        IBDatabase = new TIBDatabase(Form1);
        IBTransaction = new TIBTransaction(Form1);
        IBQuery = new TIBQuery(Form1);
        IBQuery1 = new TIBQuery(Form1);
        DataSource = new TDataSource(Form1);
        slProfileFile = new TStringList;
        slProfileFile = new TStringList;
        slRoamingFile = new TStringList;
        slName = new TStringList; //���������� � ��������
        slTextZapros = new TStringList; //���������� � ��������
        //Participant
        slParticipant = new TStringList;
        slParticipantPerson = new TStringList;
        slParticipantKind = new TStringList;
        slParticipantType = new TStringList;
          
        rf = new TRoamingFile("AmirsDop");
        ProfileFile     = this->rf->Get(TRoamingFile::ProfileFile);
        ProfileFileLast = this->rf->Get(TRoamingFile::ProfileFileLast);
        this->ProfileZagruz();
        this->Connect();


        AnsiString asp = rf->Get(TRoamingFile::page);
        if (asp=="�������") {};
        if (asp=="������")  {};
        if (asp=="��������"){}
                       else {};

        };

TAmirs::TAmirs(TForm * Form, AnsiString NameApp1)
        {
        Errors=0;
        slErrors = new TStringList;
        Form1=Form;
        NameApp=NameApp1;
        IBDatabase = new TIBDatabase(Form1);
        IBTransaction = new TIBTransaction(Form1);
        IBQuery = new TIBQuery(Form1);
        IBQuery1 = new TIBQuery(Form1);
        DataSource = new TDataSource(Form1);
        slProfileFile = new TStringList;
        slProfileFile = new TStringList;
        slRoamingFile = new TStringList;
        slName = new TStringList; //���������� � ��������
        slTextZapros = new TStringList; //���������� � ��������
        //Participant
        slParticipant = new TStringList;
        slParticipantPerson = new TStringList;
        slParticipantKind = new TStringList;
        slParticipantType = new TStringList;

        //��������� Word
        //WordNazv = new TStringList;
        WordName = new TStringList;
        WordVal  = new TStringList;
      try
        {
        //��������� �� ����� �������� �������� � �������
        TFileStream *tfile =new TFileStream(GetSpecialFolderPath(CSIDL_APPDATA)+"\\AmirsDop\\"+NameApp+".ini",fmOpenReadWrite);
        slRoamingFile->LoadFromStream(tfile);
        delete(tfile);
        ProfileFile = slRoamingFile->Strings[0];
        ProfileFileLast = slRoamingFile->Strings[1];
        //DataSourceDel->OnDataChange=OnClickX;
        this->ProfileZagruz();
        this->Connect();
        }
      catch(Exception &ex)
          {
           Errors++;
           slErrors->Add("������ �������� �������. ��������� �������");
          //MessageBox(NULL, ex.Message.c_str(), "Amirs", MB_OK | MB_ICONERROR);
          }
        };

void TAmirs::ProfileSelect()
        {
          AnsiString strLast; strLast = "��������� �������:" + ProfileFileLast + ". ��������� ���?";
          if (!ProfileFileLast.IsEmpty() && (MessageBox(NULL, strLast.c_str(), "", MB_YESNO) == IDYES))
           {AnsiString st1; st1=ProfileFile; ProfileFile=ProfileFileLast; ProfileFileLast=st1;}
           else
           {TOpenDialog *OpenDialog1 = new TOpenDialog(Form1);
            OpenDialog1->InitialDir= ExtractFilePath(Application->ExeName);
            if (OpenDialog1->Execute()){ProfileFileLast=ProfileFile; ProfileFile=OpenDialog1->FileName;};
           };
          this->rf->Put(TRoamingFile::ProfileFile, ProfileFile);
          this->rf->Put(TRoamingFile::ProfileFileLast, ProfileFileLast);
          this->ProfileZagruz();
         };

void TAmirs::ProfileZagruz()
        {
        //MessageBox(NULL, "7", "ProfileZagruz", MB_OK);
        try
         {TFileStream *tfile =new TFileStream(ProfileFile,fmOpenReadWrite);
          slProfileFile->LoadFromStream(tfile);
          delete(tfile);
         }
         catch(Exception &ex)
         {Errors++;
         slErrors->Add("������ �������� �������. ��������� �������");
        MessageBox(NULL, ex.Message.c_str(), "ProfileZagruz", MB_OK | MB_ICONERROR);
         }
        };

void TAmirs::ActivateTabSheetShablon(TDBNavigator * DBN1)
        {
        this->DataSource->DataSet=this->IBQuery;
        DBN1->DataSource=this->DataSource;
        this->Connect();
        this->VidDela=1;
        this->Year=2024;
        this->GetDela();
        };

void TAmirs::Connect()
        {
        try
        {
         TFileStream *tfile =new TFileStream(ProfileFile,fmOpenReadWrite);
         slProfileFile->LoadFromStream(tfile);
         delete(tfile);
         IBDatabase->Connected=false;
         AnsiString DBName, st; st = FindSL(slProfileFile, "database:"); //������ ������� �������
         for (int pos=1; pos<=st.Length(); pos++) {if (st[pos]=='\\') DBName = DBName+st[pos]+st[pos]; else DBName = DBName+st[pos];};
         IBDatabase->DatabaseName = DBName;
         IBDatabase->Params->Clear();
         IBDatabase->Params->Add("user_name=sysdba");
         IBDatabase->Params->Add("password=masterkey");
         IBDatabase->Params->Add("lc_ctype=CYRL"); //WIN1251
         IBDatabase->LoginPrompt=false;
         IBDatabase->Connected=true;
         IBTransaction->DefaultDatabase=IBDatabase;
         IBTransaction->Active=true;
         }
      catch(Exception &ex)
          {
           Errors++;
           slErrors->Add("������ ���������� � ����� Connect. ��������� �������");
          //MessageBox(NULL, ex.Message.c_str(), "Connect", MB_OK | MB_ICONERROR);
          }
        };

void TAmirs::GetDela()
        {
        try
        {
         IBTransaction->DefaultDatabase=IBDatabase;
         IBQuery->Database=IBDatabase;
         IBQuery->SQL->Clear();
         //IBQuery1->SQL->Add("select * from ZAPROS");
         TStringList * sl = new TStringList;
         switch (this->VidDela)
                {
                case 1: GetTextFromResourse(sl, "zaprosUD"); break;
                case 2: GetTextFromResourse(sl, "zaprosGD"); break;
                case 3: GetTextFromResourse(sl, "zaprosAD"); break;
                case 4: GetTextFromResourse(sl, "zaprosIU"); break;
                case 5: GetTextFromResourse(sl, "zaprosIG"); break;
                case 6: GetTextFromResourse(sl, "zaprosGD"); break;
                case 7: GetTextFromResourse(sl, "zaprosAD_SH"); break;
                case 8: GetTextFromResourse(sl, "zaprosAD_OR"); break;
                case 9: GetTextFromResourse(sl, "zaprosGDIsk"); break;
                };
         IBQuery->SQL->Add(ZaprosYear(sl->Text, ":paramYear", IntToStr(this->Year)));
         IBQuery->Active=true;
         IBQuery->FetchAll();
        }
      catch(Exception &ex)
          {
           Errors++;
           slErrors->Add("������ ��������� ������ ��� �� ���� GetDela. ��������� �������");
          //MessageBox(NULL, ex.Message.c_str(), "Connect", MB_OK | MB_ICONERROR);
          }
        };


void TAmirs::GetDelaZapros(int year, AnsiString NameZapros)
        {
        try
        {
         IBTransaction->DefaultDatabase=IBDatabase;
         IBQuery1->Database=IBDatabase;
         IBQuery1->SQL->Clear();
         TStringList * sl = new TStringList;
         GetTextFromResourse(sl, NameZapros);
         IBQuery1->SQL->Add(ZaprosYear(sl->Text, ":paramYear", IntToStr(year)));
         IBQuery1->Active=true;
         IBQuery1->FetchAll();
        }
      catch(Exception &ex)
          {
           Errors++;
           slErrors->Add("������ ��������� ������ ��� �� ���� GetDelaZapros. ��������� �������");
          //MessageBox(NULL, ex.Message.c_str(), "Connect", MB_OK | MB_ICONERROR);
          }
        };



//������ � Word
void TAmirs::WordClear() {/*WordNazv->Clear();*/ WordName->Clear(); WordVal->Clear();};
void TAmirs::WordAdd(/*AnsiString Nazv1, */ AnsiString Name1, AnsiString Val1) {/*WordNazv->Add(Nazv1);*/ WordName->Add(Name1); WordVal->Add(Val1);};

void TAmirs::SplitSLToWord (TStringList * SL)
        {for (int j=0; j<SL->Count; j++)
                {int pos=SL->Strings[j].Pos(":");
                if (pos) //return SL->Strings[j].SubString(pos+tfind.Length(),SL->Strings[j].Length()).TrimLeft();
                //Memo->Lines->Add(SL->Strings[j].SubString(pos+tfind.Length(),SL->Strings[j].Length()).TrimLeft(););
                this->WordAdd(SL->Strings[j].SubString(1,pos-1).TrimLeft(), SL->Strings[j].SubString(pos+1,SL->Strings[j].Length()).TrimLeft());
                };
        };

void TAmirs::GetTextToWord()
        {
        WordClear();
        WordAdd("����� ����", this->IBQuery->Fields->FieldByName("NUM")->Value);
        if (!this->IBQuery->Fields->FieldByName("OID")->IsNull) this->OID = AnsiString(this->IBQuery->Fields->FieldByName("OID")->Value);
        if (!this->IBQuery->Fields->FieldByName("JudicalUid")->IsNull) this->UniqueNum = AnsiString(this->IBQuery->Fields->FieldByName("JudicalUid")->Value);
        this->WordAdd("���",this->UniqueNum);
        this->GetParticipant();
        //this->PhysPerson("1640", "�����__________");
        //MessageBox(NULL, "3", "TAmirs::zaprosDelInfo", MB_OK);
        AnsiString res;
        for (int j=0; j<slParticipantPerson->Count; j++)
                {res += " OID="+this->OID;
                 res += " Participant="+slParticipant->Strings[j];
                 res += " ParticipantPerson="+slParticipantPerson->Strings[j];
                 res += " ParticipantKind="+slParticipantKind->Strings[j];
                 res += " ParticipantType="+slParticipantType->Strings[j];
                 };
        //MessageBox(NULL,res.c_str(), "", MB_OK);
         AnsiString sdata;
                if (!this->IBQuery->Fields->FieldByName("CreateDate")->IsNull) sdata=this->IBQuery->Fields->FieldByName("CreateDate")->Value;
                WordAdd("���� ��������",DS(sdata.SubString(1,10)));
                if (!this->IBQuery->Fields->FieldByName("SittingDate")->IsNull) sdata=this->IBQuery->Fields->FieldByName("SittingDate")->Value;
                WordAdd("��������",  DS(sdata.SubString(1,10)));
                WordAdd("���������", TS(sdata.SubString(12,5)));

        switch (this->VidDela)
                {
         case 1:{
                AnsiString st; if (!this->IBQuery->Fields->FieldByName("MAIN_Articles")->IsNull) st = this->IBQuery->Fields->FieldByName("MAIN_Articles")->Value;
                st = "��." + st.SubString(7, st.Length());
                WordAdd("������", st);
                TStringList * slComment = new TStringList;
                slComment->Text = GetFirstRecForTable ("Document", "Comment", "\"OID\" = "+this->OID);
                this->SplitSLToWord(slComment);
                }; break;
         case 2: {
                         //��������� ������ ������
                         AnsiString Claim = GetFirstRecForTable ("Claim", "OID", "\"CivilCase\" = "+this->OID);
                         res = GetFirstRecForTable ("CivilCaseCategoryLink", "VnKod", "\"Claim\" = "+Claim);
                         if (res == "100110")  res="2"; if (res == "100210") res="3";
                         if (res == "210010")  res="114";
                         if (res == "50001001") res="150";
                         if (res == "510020") res="151";
                         if (res == "29001101") res="152";
                         if (res == "29001102") res="153";
                         if (res == "3101103") res="168";
                         if (res == "3102101" | res == "310310") res="169";
                         if (res == "510010")  res="203";
                         if (res == "520300")  res="209";
                         if (res == "270610")  res="75";
                         this->WordAdd("������������",res);
                         res = GetFirstRecForTable ("ClaimInAction", "Essence", "\"Claim\" = "+Claim);
                         this->WordAdd("�������",res);
                         res = GetFirstRecForTable ("ClaimInAction", "Sum", "\"Claim\" = "+Claim);
                         this->WordAdd("�����",res);
                         //res = GetFirstRecForTable ("ClaimInAction", "Comment", "\"Claim\" = "+Claim);
                         //this->WordAdd("�����������",res);
                         TStringList * slComment = new TStringList;
                         slComment->Text = GetFirstRecForTable ("Document", "Comment", "\"OID\" = "+this->OID);
                         this->SplitSLToWord(slComment);
                         //this->WordAdd("�����������2",res);

                         //this->WordAdd("���� � ���������",GetFirstRecForTable ("Document", "CreateDate", "\"OID\" = "+this->OID));
                         //  this->WordAdd("Pflfxf",GetFirstRecForTable ("Document", "CreateDate", "\"OID\" = "+this->OID));
                         //AnsiString s1; s1="115552"; int t1=115552;
                          //s1=this->GetFirstRecForTable ("CivilRawRegData", "ToJudgeDate", "\"OID\" = "+s1);
                        // AnsiString M1 = GetFirstRecForTable ("TaskGerm", "ExecutionDate", "\"OID\" = 195713");
                          //this->WordAdd("������", M1);
                            /*
                          AnsiString MaterialData = GetFirstRecForTable ("CaseTask", "Data", "\"Material\" = "+this->OID+" and \"Stage\"=32");
                          this->WordAdd("���� ���������", MaterialData);
                          TIBQuery * IBQuery2 = new TIBQuery(Form1);
                          this->GetAllRecForTable (IBQuery2, "CivilRawRegData", "\"OID\" =" + MaterialData); //"+s1);
                          IBQuery2->First();
                          TDateTime datetime1;
                          datetime1 = IBQuery2->Fields->FieldByName("ToJudgeDate")->Value.VDate;
                          this->WordAdd("���� ��������", DataToString(datetime1));

                           this->GetAllRecForTable (IBQuery2, "TaskGerm", "\"OID\" = 118649"); //"+s1);
                          IBQuery2->First();
                          //TDateTime datetime1;
                          datetime1 = IBQuery2->Fields->FieldByName("ExecutionDate")->Value.VDate;
                          this->WordAdd("������", DataToString(datetime1));
                           */
                         }; break;
          case 3: {
                /*
                AnsiString sdata;
                if (!this->IBQuery->Fields->FieldByName("CreateDate")->IsNull)
                { sdata=this->IBQuery->Fields->FieldByName("CreateDate")->Value;};
                this->WordAdd("���� ��������", DS(sdata.SubString(1,10)));
                if (!this->IBQuery->Fields->FieldByName("SittingDate")->IsNull)
                { sdata=this->IBQuery->Fields->FieldByName("SittingDate")->Value;};
                this->WordAdd("��������", DS(sdata.SubString(1,10)));
                this->WordAdd("���������", sdata.SubString(12,5));
                */
                TStringList * slComment = new TStringList;
                slComment->Text = GetFirstRecForTable ("Document", "Comment", "\"OID\" = "+this->OID);
                this->SplitSLToWord(slComment);
                AnsiString st, st2, st3; st =  GetFirstRecForTable ("AdmArticle", "Code", "\"AdmCase\" = "+this->OID);
                //this->WordAdd("������", st);
                st = st.SubString(7,st.Length());
                if (st.Length()>0)
                {if (st[1]=='0') st=st.SubString(2,st.Length());
                 if (st[1]=='0') st=st.SubString(2,st.Length());
                 int pos = st.Pos(","); st2 = st.SubString(pos+1, st.Length());
                 if (st2[1]=='0') st2 = st2.SubString(2, st2.Length());
                 st = "��. " + st.SubString(1, pos-1);
                 pos = st2.Pos("."); st3 = st2.SubString(pos+1, st2.Length());
                 st2 = st2.SubString(1, pos-1);
                 if (st2 == "0") st2=""; else st2=" �. "+st2;
                 if (st3 == "0") st3=""; else st3="."+st3;
                 st+=st2+st3;
                 };
                this->WordAdd("������", st);
                        }; break;
                case 4: {}; break;
          case 5: {
                //���������� �����������

                this->WordAdd("����� ���� ����������", this->Num());
                this->WordAdd("���������", this->Zayvitel);
                this->WordAdd("������� ���� ��������", this->DolznikDataRozd);
                this->WordAdd("������� �����", this->DolznikAddr);
                this->WordAdd("������ ���", this->StarOID);

                TStringList * slComment = new TStringList;
                slComment->Text = GetFirstRecForTable ("Document", "Comment", "\"OID\" = "+this->OID);
                this->SplitSLToWord(slComment);
                //��������������
                this->WordAdd("���������", FindSL(slComment, "���������="));
                this->WordAdd("�������", FindSL(slComment, "�������="));
                this->WordAdd("������������������", FindSL(slComment, "����="));
                //����������
                this->WordAdd("������", FindSL(slComment, "������="));
                this->WordAdd("�����", FindSL(slComment, "�����="));
                this->WordAdd("������������", FindSL(slComment, "������������="));
               };
               break;
        };//����� switch
        //MessageBox(NULL, "6", "TAmirs::GetTextToWord", MB_OK);
        SplitSLToWord(slProfileFile);
        };

void TAmirs::TextToMemo(TMemo * Memo)
        {
        //Memo->Lines->Clear();
        Memo->Clear();
        AnsiString st="������: "; if (!FileNameShablon.IsEmpty()) st+=FileNameShablon; else st+="�� ������";
        Memo->Lines->Add(st);
        for (int i = 0; i < WordName->Count; i++)
                {
                  Memo->Lines->Add("%"+this->WordName->Strings[i]+"% : "+this->WordVal->Strings[i]);
                }

        };

void TAmirs::TextToWord()
        {
        //��������� ����� ������ ����� ���� � � ����� ���������
        AnsiString st1;
        st1 = this->IBQuery->Fields->FieldByName("NUM")->Value;
        for (int pos=1; pos<=st1.Length(); pos++) {if (st1[pos]==47) st1[pos]=45;};
        AnsiString file11 = GetFileFromFileName(FileNameShablon);
        WideString ASfilename=GetSpecialFolderPath(CSIDL_PERSONAL)+"\\"+st1+"-"+file11;
        Variant app ;
        Variant my_docs ;
        Variant active_doc ;
        //Variant bm; Variant bm2;
        //Variant vVarParagraphs;
        //Variant vVarParagraph;
        //AnsiString mark="name";
        app = CreateOleObject("Word.Application");
        app.OlePropertySet("Visible", 1);
        my_docs = app.OlePropertyGet("Documents");
        active_doc = my_docs.OleFunction("open", this->FileNameShablon);
        active_doc.OleFunction("SaveAs",ASfilename);
        for (int i = 0; i < WordName->Count; i++)
                {
                  WideString s; s="%"+WordName->Strings[i]+"%";
                  WideString s1; s1=WordVal->Strings[i];
                 try
                {
                         app.OlePropertyGet("Selection").OlePropertyGet("Find").OleProcedure
                        ("Execute", /*FindText=*/ s, /*MatchCase=*/false,
                        /*MatchWholeWord=*/ false, /*MatchWildcards=*/false,
                        /*MatchSoundsLike=*/false, /*MatchAllWordForms=*/false,
                        /*Forward=*/true, /*Wrap=*/1, /*Format=*/false,
                        /*ReplaceWith=*/ s1, /*Replace=*/ 2);
                 }
                 catch(Exception &ex)
                 {
                       // MessageBox(NULL, "������� ������� ��������",ex.Message.c_str(), MB_OK | MB_ICONERROR);
                 }
               // AnsiString s3,s4; s3=s; s4=s1;
                //s3+="--"+s4;
                //MessageBox(NULL, s3.c_str(), "TAmirs::TextToWord", MB_OK);
                }
        active_doc.OleFunction("Save"); //FileName.c_str());
        //active_doc.OleFunction("PrintOut");
        };

void TAmirs::Search(AnsiString Text)
        {
  // MessageBox(NULL, "1", "TAmirs::Search", MB_OK);
        try
        {
         TLocateOptions Opts; Opts.Clear(); Opts << loPartialKey;
         Variant locvalues[2];
         AnsiString poisk;
         switch (this->VidDela)
                {
                case 1: poisk = "1-"+Text+"/"+IntToStr(this->Year); break;
                case 2: poisk = "2-"+Text+"/"+IntToStr(this->Year); break;
                case 3: poisk = "3-"+Text+"/"+IntToStr(this->Year); break;
                case 4: poisk = "14-"+Text+"/"+IntToStr(this->Year); break;
                case 5: poisk = "13-"+Text+"/"+IntToStr(this->Year); break;
                case 6: poisk = "2�-"+Text+"/"+IntToStr(this->Year); break;
                };
         //poisk = "1-"+Text+"/"+IntToStr(this->Year);
         //poisk = Text;
         locvalues[0] = Variant(poisk);//"13-264/2022");
         this->IBQuery->Locate("Num", VarArrayOf(locvalues, 0), Opts);
         //MessageBox(NULL, "2", "TAmirs::Search", MB_OK);
         }
      catch(Exception &ex)
          {Errors++; slErrors->Add("������ ������ ���. ��������� �������");
          };
        };

//������ � �����

AnsiString TAmirs::GetFirstRecForTable (AnsiString Table, AnsiString Field, AnsiString Where)
        {
          AnsiString res="";
          try
          {
           if (!Table.IsEmpty() & !Field.IsEmpty() & !Where.IsEmpty())
           {TIBQuery *IBQuery = new TIBQuery(Form1);
           IBQuery->Active=false;
           IBQuery->Database=IBDatabase;
           IBQuery->Transaction=IBTransaction;
           IBQuery->SQL->Clear();
           IBQuery->SQL->Add("select * from \"" + Table + "\" where "+Where);
           IBQuery->Active=true;
           IBQuery->FetchAll();
           IBQuery->First();
           if (!IBQuery->Fields->FieldByName(Field)->IsNull)
                res = AnsiString(IBQuery->Fields->FieldByName(Field)->Value);
           };
          }
          catch(Exception &ex)
          { Errors++; slErrors->Add("������ ������� ��� �� ����. ��������� GetFirstRecForTable"); }
        return res;
        };

void TAmirs::GetAllRecForTable (TIBQuery * IBQuery, AnsiString Table, AnsiString Where)
        {
          try
          {
           if (!Table.IsEmpty() & !Where.IsEmpty())
           {IBQuery->Active=false;
           IBQuery->Database=IBDatabase;
           IBQuery->Transaction=IBTransaction;
           IBQuery->SQL->Clear();
           IBQuery->SQL->Add("select * from \"" + Table + "\" where "+Where);
           IBQuery->Active=true;
           IBQuery->FetchAll();
          }
          }
          catch(Exception &ex)
          { Errors++; slErrors->Add("������ ������� ��� �� ����. ��������� GetAllRecForTable"); }
        };

void TAmirs::GetParticipant()
        {
         try
          {
           TIBQuery *IBQuery = new TIBQuery(Form1);
           this->GetAllRecForTable (IBQuery, "Participant", "\"CaseMaterial\"="+this->OID);
           //IBQueryAgency->Active=false;
           //IBQueryAgency->Database=IBDatabase;
           //IBQueryAgency->Transaction=IBTransaction;
           //IBQueryAgency->SQL->Clear();
           //IBQueryAgency->SQL->Add("select * from \"Participant\" where \"CaseMaterial\"="+this->OID);
           //IBQueryAgency->Active=true;
           TStringList * slIstec = new TStringList; slIstec->Clear();
           TStringList * slOtvetchik = new TStringList; slOtvetchik->Clear();
           slIstecRod = new TStringList; slIstecRod->Clear();
           slOtvetDat = new TStringList; slOtvetDat->Clear();
           slParticipant->Clear();
           slParticipantPerson->Clear();
           slParticipantKind->Clear();
           slParticipantType->Clear();
           //IBQueryAgency->FetchAll();
           IBQuery->First();
           while (!IBQuery->Eof)
                {
                AnsiString asPrefix, asPerson, ParticipantType; int ParticipantKind;
                if (!IBQuery->Fields->FieldByName("ParticipantKind")->IsNull)
                        {ParticipantKind = StrToInt(IBQuery->Fields->FieldByName("ParticipantKind")->Value);
                        slParticipantKind->Add(AnsiString(IBQuery->Fields->FieldByName("ParticipantKind")->Value));
                        asPrefix = GetFirstRecForTable ("ParticipantKind", "Name", "\"Id\" = "+(AnsiString(IBQuery->Fields->FieldByName("ParticipantKind")->Value)));
                        } else slParticipantKind->Add("���");

                switch (ParticipantKind)
                         {
                         case  8: asPrefix = "�����"; break;
                         case  31: asPrefix = "���������������"; break;
                         };
                if (!IBQuery->Fields->FieldByName("OID")->IsNull)
                        slParticipant->Add(AnsiString(IBQuery->Fields->FieldByName("OID")->Value)); else slParticipant->Add("���");
                if (!IBQuery->Fields->FieldByName("Person")->IsNull)
                        {slParticipantPerson->Add(AnsiString(IBQuery->Fields->FieldByName("Person")->Value));
                        asPerson = AnsiString(IBQuery->Fields->FieldByName("Person")->Value);
                        } else slParticipantPerson->Add("���");
                if (!IBQuery->Fields->FieldByName("Type")->IsNull)
                         {ParticipantType=IBQuery->Fields->FieldByName("Type")->Value;
                         slParticipantType->Add(AnsiString(IBQuery->Fields->FieldByName("Type")->Value));}
                        else slParticipantType->Add("���");
                //���������� ��� ��� ����
               // if (this->VidDela==2) if (ParticipantType=="6") this->PhysPerson(asPerson, asPrefix); else this->JuridicalPerson(asPerson, asPrefix);
                if (this->VidDela==1) this->PhysPerson(asPerson, asPrefix);
                if (ParticipantType=="6" | ParticipantType=="3") this->PhysPerson(asPerson, asPrefix);
                if (ParticipantType=="11" | ParticipantType=="4") this->JuridicalPerson(asPerson, asPrefix);
                if (this->VidDela==5 & ParticipantKind==38)
                        if (ParticipantType=="6") this->PhysPerson(asPerson, asPrefix); else this->JuridicalPerson(asPerson, asPrefix);
                //this->PhysPerson("1346", "�����");
                IBQuery->Next(); // i++;
                }
           if (this->VidDela==2)
                {AnsiString IstecAll=""; for (int j=0; j<slIstecRod->Count; j++) {IstecAll += slIstecRod->Strings[j]; if (j<(slIstecRod->Count-1)) IstecAll += ", ";}
                AnsiString OtvetAll=""; for (int j=0; j<slOtvetDat->Count; j++) {OtvetAll += slOtvetDat->Strings[j]; if (j<(slOtvetDat->Count-1)) OtvetAll += ", ";}
                this->WordAdd("�����������", IstecAll);
                this->WordAdd("���������������", OtvetAll);
                };
          }
          catch(Exception &ex)
         { Errors++; slErrors->Add("������ ������� ��� �� ����. ��������� GetParticipant"); }

         };

void TAmirs::Agency(int OID, AnsiString *Name, AnsiString *Address)
        {TIBQuery *IBQueryAgency = new TIBQuery(Form1);
         IBQueryAgency->Active=false;
         IBQueryAgency->Database=IBDatabase;
         IBQueryAgency->Transaction=IBTransaction;
         IBQueryAgency->SQL->Clear();
         IBQueryAgency->SQL->Add("select * from \"Agency\" where \"OID\"="+IntToStr(OID));
         IBQueryAgency->Active=true;
         *Name=""; *Address="";
         IBQueryAgency->FetchAll();
         IBQueryAgency->First();
         if (!IBQueryAgency->Fields->FieldByName("Name")->IsNull)
                {*Name = IBQueryAgency->Fields->FieldByName("Name")->Value;};
         //if (!IBQueryAgency->Fields->FieldByName("Address")->IsNull)
         //       {*Address = IBQueryAgency->Fields->FieldByName("Address")->Value;};
         };

void TAmirs::Jurik(int OID, AnsiString *Name, AnsiString *Address)
        {
         try
          {
         TIBQuery *IBQueryAgency = new TIBQuery(Form1);
         IBQueryAgency->Active=false;
         IBQueryAgency->Database=IBDatabase;
         IBQueryAgency->Transaction=IBTransaction;
         IBQueryAgency->SQL->Clear();
         IBQueryAgency->SQL->Add("select * from \"JuridicalPerson\" where \"OID\"="+IntToStr(OID));
         IBQueryAgency->Active=true;
         *Name=""; *Address="";
         IBQueryAgency->FetchAll();
         IBQueryAgency->First();
         if (!IBQueryAgency->Fields->FieldByName("Name")->IsNull)
                {*Name = IBQueryAgency->Fields->FieldByName("Name")->Value;};
         //if (!IBQueryAgency->Fields->FieldByName("Address")->IsNull)
         //       {*Address = IBQueryAgency->Fields->FieldByName("Address")->Value;};
         }
          catch(Exception &ex)
         // catch(EDivByZero &ex)
          {
           MessageBox(NULL, ex.Message.c_str(), "Jurik", MB_OK | MB_ICONERROR);
          }
         };



void TAmirs::fizik(AnsiString OID)
        {if (!OID.IsEmpty())
         {TIBQuery *IBQueryAgency = new TIBQuery(Form1);
         IBQueryAgency->Active=false;
         IBQueryAgency->Database=IBDatabase;
         IBQueryAgency->Transaction=IBTransaction;
         IBQueryAgency->SQL->Clear();
         IBQueryAgency->SQL->Add("select * from \"PhysPerson\" where \"OID\"="+OID);
         IBQueryAgency->Active=true;
         IBQueryAgency->FetchAll();
         IBQueryAgency->First();
         if (!IBQueryAgency->Fields->FieldByName("BirthDate")->IsNull)
         this->DolznikDataRozd = AnsiString(IBQueryAgency->Fields->FieldByName("BirthDate")->Value); else this->DolznikDataRozd="";
         };};

void TAmirs::PhysPerson(AnsiString OID, AnsiString Prefix)
         {
         try
          {
           if (!OID.IsEmpty())
           {TIBQuery *IBQuery = new TIBQuery(Form1);
           this->GetAllRecForTable (IBQuery, "PhysPerson", "\"OID\"="+OID);
           IBQuery->First();
           AnsiString name=""; int Gender=0;
           if (!IBQuery->Fields->FieldByName("Surname")->IsNull) name += IBQuery->Fields->FieldByName("Surname")->Value;
           if (!IBQuery->Fields->FieldByName("Name")->IsNull) name += " "+IBQuery->Fields->FieldByName("Name")->Value;
           if (!IBQuery->Fields->FieldByName("Patronymic")->IsNull) name += " "+IBQuery->Fields->FieldByName("Patronymic")->Value;
           if (!IBQuery->Fields->FieldByName("Gender")->IsNull) Gender = IBQuery->Fields->FieldByName("Gender")->Value;
           this->WordAdd(Prefix+"_�", name);
           FIO fio1(name, Gender);
           this->WordAdd(Prefix+"_�", fio1.fiodat);
           this->WordAdd(Prefix+"_�", fio1.fiorod);
           this->WordAdd(Prefix+"_�", fio1.fiovin);
           this->WordAdd(Prefix+"_�", fio1.fiotvo);
           this->WordAdd(Prefix+"_����", fio1.fiolit);
           this->WordAdd(Prefix+"_����", fio1.fiolitdat);
           this->WordAdd(Prefix+"_����", fio1.fiolitrod);
           this->WordAdd(Prefix+"_����", fio1.fiolitvin);
           this->WordAdd(Prefix+"_����", fio1.fiolittvo);
           this->WordAdd(Prefix+"���", IntToStr(fio1.sex)); //Gender));
           //this->WordAdd(Prefix+"���Gender", IntToStr(Gender));
           //this->WordAdd(Prefix+"���Sex", IntToStr(fio1.sex));
           if (this->VidDela==2) if (fio1.sex == 1) this->WordAdd("�������_�", fio1.fiotvo); else this->WordAdd("�������_�", fio1.fiotvo);
           if (Prefix=="�����" | Prefix=="����������") this->slIstecRod->Add(fio1.fiorod);
           if (Prefix=="��������" | Prefix=="�������") this->slOtvetDat->Add(fio1.fiodat);
           if (!IBQuery->Fields->FieldByName("BirthDate")->IsNull) this->WordAdd(Prefix+"������������", AnsiString(IBQuery->Fields->FieldByName("BirthDate")->Value));
           if (!IBQuery->Fields->FieldByName("BirthPlace")->IsNull) this->WordAdd(Prefix+"�������������", AnsiString(IBQuery->Fields->FieldByName("BirthPlace")->Value));
           if (!IBQuery->Fields->FieldByName("Snils")->IsNull) this->WordAdd(Prefix+"�����",AnsiString(IBQuery->Fields->FieldByName("Snils")->Value));
           if (!IBQuery->Fields->FieldByName("Phone")->IsNull) this->WordAdd(Prefix+"�������",AnsiString(IBQuery->Fields->FieldByName("Phone")->Value));
           this->Address(OID, Prefix);
           this->PhysPersonDocumentGetAll(OID, Prefix);
           };
          }
          catch(Exception &ex)
          { Errors++; slErrors->Add("������ ������� ��� �� ����. ��������� PhysPerson"); }
         };

void TAmirs::JuridicalPerson(AnsiString OID, AnsiString Prefix)
         {
         try
          {
           if (!OID.IsEmpty())
           {TIBQuery *IBQuery = new TIBQuery(Form1);
           this->GetAllRecForTable (IBQuery, "JuridicalPerson", "\"OID\"="+OID);
           IBQuery->First();
           AnsiString name, nameIm, nameDat, nameRod, nameTvo, nameVin;
           if (!IBQuery->Fields->FieldByName("Name")->IsNull) name = IBQuery->Fields->FieldByName("Name")->Value;
           nameIm=name; nameDat=name; nameRod=name; nameTvo=name; nameVin=name;
           if (name.SubString(1,3)=="���")
                {nameIm="�������� � ������������ ����������������"+name.SubString(4,name.Length());
                 nameDat="�������� � ������������ ����������������"+name.SubString(4,name.Length());
                 nameRod="�������� � ������������ ����������������"+name.SubString(4,name.Length());
                 nameTvo="��������� � ������������ ����������������"+name.SubString(4,name.Length());
                 nameVin="�������� � ������������ ����������������"+name.SubString(4,name.Length());
                };
           if (name.SubString(1,3)=="���")
                {nameIm="��������� ����������� ��������"+name.SubString(4,name.Length());
                 nameDat="���������� ������������ ��������"+name.SubString(4,name.Length());
                 nameRod="���������� ������������ ��������"+name.SubString(4,name.Length());
                 nameTvo="��������� ����������� ���������"+name.SubString(4,name.Length());
                 nameVin="��������� ����������� ��������"+name.SubString(4,name.Length());
                };
           if (name.SubString(1,3)=="���")
                {nameIm="�������� ����������� ��������"+name.SubString(4,name.Length());
                 nameDat="��������� ������������ ��������"+name.SubString(4,name.Length());
                 nameRod="��������� ������������ ��������"+name.SubString(4,name.Length());
                 nameTvo="�������� ����������� ���������"+name.SubString(4,name.Length());
                 nameVin="�������� ����������� ��������"+name.SubString(4,name.Length());
                };
           if (name.SubString(1,3)=="�� ")
                {nameIm="����������� �������� "+name.SubString(4,name.Length());
                 nameDat="������������ �������� "+name.SubString(4,name.Length());
                 nameRod="������������ �������� "+name.SubString(4,name.Length());
                 nameTvo="����������� ��������� "+name.SubString(4,name.Length());
                 nameVin="����������� �������� "+name.SubString(4,name.Length());
                };
           if (name.SubString(1,3)=="�� ")
                {nameIm="�������������� ��������������� "+name.SubString(4,name.Length());
                 nameDat="��������������� ��������������� "+name.SubString(4,name.Length());
                 nameRod="��������������� ��������������� "+name.SubString(4,name.Length());
                 nameTvo="�������������� ���������������� "+name.SubString(4,name.Length());
                 nameVin="��������������� ��������������� "+name.SubString(4,name.Length());
                };
           this->WordAdd(Prefix+"_�", nameIm);
           this->WordAdd(Prefix+"_�", nameDat);
           this->WordAdd(Prefix+"_�", nameRod);
           this->WordAdd(Prefix+"_�", nameTvo);
           this->WordAdd(Prefix+"_�", nameVin);
           this->WordAdd(Prefix+"_����", name);
           this->WordAdd(Prefix+"_����", name);
           this->WordAdd(Prefix+"_����", name);
           this->WordAdd(Prefix+"_����", name);
           this->WordAdd(Prefix+"_����", name);
           if (Prefix=="�����" | Prefix=="����������") this->slIstecRod->Add(nameRod);
           if (Prefix=="��������" | Prefix=="�������") this->slOtvetDat->Add(nameDat);
           //if (!IBQuery->Fields->FieldByName("BirthDate")->IsNull) this->WordAdd(Prefix+"������������", AnsiString(IBQuery->Fields->FieldByName("BirthDate")->Value));
           //if (!IBQuery->Fields->FieldByName("BirthPlace")->IsNull) this->WordAdd(Prefix+"�������������", AnsiString(IBQuery->Fields->FieldByName("BirthPlace")->Value));
           //if (!IBQuery->Fields->FieldByName("Snils")->IsNull) this->WordAdd(Prefix+"�����",AnsiString(IBQuery->Fields->FieldByName("Snils")->Value));
           this->Address(OID, Prefix);
           };
          }
          catch(Exception &ex)
          { Errors++; slErrors->Add("������ ������� ��� �� ����. ��������� JuridicalPerson"); }
         };

AnsiString TAmirs::Num()
        {TIBQuery *IBQueryAgency = new TIBQuery(Form1);
         IBQueryAgency->Active=false;
         IBQueryAgency->Database=IBDatabase;
         IBQueryAgency->Transaction=IBTransaction;
         IBQueryAgency->SQL->Clear();
         IBQueryAgency->SQL->Add("select * from \"ExecMaterial\" where \"OID\"="+this->OID);
         IBQueryAgency->Active=true;
         IBQueryAgency->FetchAll();
         IBQueryAgency->First();
         AnsiString res;
         if (!IBQueryAgency->Fields->FieldByName("CourtCase")->IsNull)
                res = AnsiString(IBQueryAgency->Fields->FieldByName("CourtCase")->Value); else res="";
         this->StarOID=res;
         IBQueryAgency->Active=false;
         IBQueryAgency->SQL->Clear();
         IBQueryAgency->SQL->Add("select * from \"CaseMaterial\" where \"OID\"="+res);
         IBQueryAgency->Active=true;
         IBQueryAgency->FetchAll();
         IBQueryAgency->First();
         AnsiString DocNum;
         if (!IBQueryAgency->Fields->FieldByName("DocNum")->IsNull)
                DocNum = AnsiString(IBQueryAgency->Fields->FieldByName("DocNum")->Value); else res="";
         IBQueryAgency->Active=false;
         IBQueryAgency->SQL->Clear();
         IBQueryAgency->SQL->Add("select * from \"DocNum\" where \"Id\"="+DocNum);
         IBQueryAgency->Active=true;
         IBQueryAgency->FetchAll();
         IBQueryAgency->First();
         if (!IBQueryAgency->Fields->FieldByName("Num")->IsNull)
                res = AnsiString(IBQueryAgency->Fields->FieldByName("Num")->Value); else res="";

         return res;
         };

AnsiString TAmirs::address(AnsiString OID)
        { AnsiString res="";
        if (!OID.IsEmpty())
         {TIBQuery *IBQueryAgency = new TIBQuery(Form1);
         IBQueryAgency->Active=false;
         IBQueryAgency->Database=IBDatabase;
         IBQueryAgency->Transaction=IBTransaction;
         IBQueryAgency->SQL->Clear();
         IBQueryAgency->SQL->Add("select * from \"Address\" where \"OID\" = "+OID);
         IBQueryAgency->Active=true;
         IBQueryAgency->FetchAll();
         IBQueryAgency->First();

                if (!IBQueryAgency->Fields->FieldByName("Index")->IsNull)
                res = AnsiString(IBQueryAgency->Fields->FieldByName("Index")->Value)+ ", ";
                if (!IBQueryAgency->Fields->FieldByName("RfRegion")->IsNull)
                res += GetFirstRecForTable ("RfRegion", "Name", "\"Code\" = "+(AnsiString(IBQueryAgency->Fields->FieldByName("RfRegion")->Value)))+", ";
                if (!IBQueryAgency->Fields->FieldByName("District")->IsNull)
                res += AnsiString(IBQueryAgency->Fields->FieldByName("District")->Value)+" �����,";
                if (!IBQueryAgency->Fields->FieldByName("SettlementKind3")->IsNull)
                res += GetFirstRecForTable ("AddressObjectType", "ShortName", "\"OID\" = "+(AnsiString(IBQueryAgency->Fields->FieldByName("SettlementKind3")->Value)));
                if (!IBQueryAgency->Fields->FieldByName("Settlement")->IsNull)
                res += " "+AnsiString(IBQueryAgency->Fields->FieldByName("Settlement")->Value)+", ";
                if (!IBQueryAgency->Fields->FieldByName("SettlementKind4")->IsNull)
                res += GetFirstRecForTable ("AddressObjectType", "ShortName", "\"OID\" = "+(AnsiString(IBQueryAgency->Fields->FieldByName("SettlementKind4")->Value)))+".";
                if (!IBQueryAgency->Fields->FieldByName("SettlementBySettlement")->IsNull)
                res += " "+AnsiString(IBQueryAgency->Fields->FieldByName("SettlementBySettlement")->Value)+", ";
                if (!IBQueryAgency->Fields->FieldByName("StreetKind")->IsNull)
                res += GetFirstRecForTable ("AddressObjectType", "ShortName", "\"OID\" = "+(AnsiString(IBQueryAgency->Fields->FieldByName("StreetKind")->Value)))+". ";
                if (!IBQueryAgency->Fields->FieldByName("Street")->IsNull)
                res += AnsiString(IBQueryAgency->Fields->FieldByName("Street")->Value)+", ";
                if (!IBQueryAgency->Fields->FieldByName("Building")->IsNull)
                res += "�. "+AnsiString(IBQueryAgency->Fields->FieldByName("Building")->Value)+", ";
                if (!IBQueryAgency->Fields->FieldByName("AdditionalBuilding")->IsNull)
                res += AnsiString(IBQueryAgency->Fields->FieldByName("AdditionalBuilding")->Value)+", ";
                if (!IBQueryAgency->Fields->FieldByName("Room")->IsNull)
                res += "��. "+AnsiString(IBQueryAgency->Fields->FieldByName("Room")->Value);
         };
         return res;
         };


void TAmirs::Address(AnsiString Subject, AnsiString Prefix)
        {

        if (!Subject.IsEmpty())
         {TIBQuery *IBQueryAgency = new TIBQuery(Form1);
         IBQueryAgency->Active=false;
         IBQueryAgency->Database=IBDatabase;
         IBQueryAgency->Transaction=IBTransaction;
         IBQueryAgency->SQL->Clear();
         IBQueryAgency->SQL->Add("select * from \"Address\" where \"Subject\" = "+Subject);
         IBQueryAgency->Active=true;
         IBQueryAgency->FetchAll();
         IBQueryAgency->First();
         //MessageBox(NULL, "addr0", "���������", MB_OK);
         AnsiString AllAddress="";
         for (int j=0; j<IBQueryAgency->RecordCount; j++)
            {
             //MessageBox(NULL, "addrIF", "���������", MB_OK);
             AnsiString res="";
             if (!IBQueryAgency->Fields->FieldByName("Index")->IsNull)
             res = AnsiString(IBQueryAgency->Fields->FieldByName("Index")->Value)+ ", ";
             if (!IBQueryAgency->Fields->FieldByName("RfRegion")->IsNull)
             res += GetFirstRecForTable ("RfRegion", "Name", "\"Code\" = "+(AnsiString(IBQueryAgency->Fields->FieldByName("RfRegion")->Value)))+", ";
             if (!IBQueryAgency->Fields->FieldByName("District")->IsNull)
             res += AnsiString(IBQueryAgency->Fields->FieldByName("District")->Value)+" �����, ";
             if (!IBQueryAgency->Fields->FieldByName("SettlementKind3")->IsNull)
             res += GetFirstRecForTable ("AddressObjectType", "ShortName", "\"OID\" = "+(AnsiString(IBQueryAgency->Fields->FieldByName("SettlementKind3")->Value)));
             if (!IBQueryAgency->Fields->FieldByName("Settlement")->IsNull)
                if (!AnsiString(IBQueryAgency->Fields->FieldByName("Settlement")->Value).IsEmpty())
                res += AnsiString(IBQueryAgency->Fields->FieldByName("Settlement")->Value)+", ";
             if (!IBQueryAgency->Fields->FieldByName("SettlementKind4")->IsNull)
             res += GetFirstRecForTable ("AddressObjectType", "ShortName", "\"OID\" = "+(AnsiString(IBQueryAgency->Fields->FieldByName("SettlementKind4")->Value)))+".";
             if (!IBQueryAgency->Fields->FieldByName("SettlementBySettlement")->IsNull)
             res += AnsiString(IBQueryAgency->Fields->FieldByName("SettlementBySettlement")->Value)+", ";
             if (!IBQueryAgency->Fields->FieldByName("StreetKind")->IsNull)
             res += GetFirstRecForTable ("AddressObjectType", "ShortName", "\"OID\" = "+(AnsiString(IBQueryAgency->Fields->FieldByName("StreetKind")->Value)))+". ";
             if (!IBQueryAgency->Fields->FieldByName("Street")->IsNull)
                if (!AnsiString(IBQueryAgency->Fields->FieldByName("Street")->Value).IsEmpty())
                res += AnsiString(IBQueryAgency->Fields->FieldByName("Street")->Value)+", ";
             if (!IBQueryAgency->Fields->FieldByName("Building")->IsNull)
             res += "�. "+AnsiString(IBQueryAgency->Fields->FieldByName("Building")->Value)+", ";
             if (!IBQueryAgency->Fields->FieldByName("AdditionalBuilding")->IsNull)
             res += AnsiString(IBQueryAgency->Fields->FieldByName("AdditionalBuilding")->Value)+", ";
             if (!IBQueryAgency->Fields->FieldByName("Room")->IsNull)
             res += "��. "+AnsiString(IBQueryAgency->Fields->FieldByName("Room")->Value);

             if (!IBQueryAgency->Fields->FieldByName("AddressType")->IsNull)
               if (AnsiString(IBQueryAgency->Fields->FieldByName("AddressType")->Value) == "3")
                 this->WordAdd(Prefix+"����������������", res);
             if (!IBQueryAgency->Fields->FieldByName("AddressType")->IsNull)
               if (AnsiString(IBQueryAgency->Fields->FieldByName("AddressType")->Value) == "5")
                 this->WordAdd(Prefix+"���������������", res);
             AllAddress +=res;
             IBQueryAgency->Next();
            };
            this->WordAdd(Prefix+"��������", AllAddress);
         };
         };

void TAmirs::PhysPersonDocumentGetAll(AnsiString Subject, AnsiString Prefix)
        {AnsiString AllDocuments="";
        if (!Subject.IsEmpty())
         {TIBQuery *IBQuery = new TIBQuery(Form1);
         GetAllRecForTable (IBQuery, "PhysPersonDocument", "\"PhysicalPerson\"="+Subject);
         IBQuery->First();
         for (int j=0; j<IBQuery->RecordCount; j++)
            {AnsiString res="";
             if (!IBQuery->Fields->FieldByName("PhysPersonDocumentType")->IsNull)
               if (AnsiString(IBQuery->Fields->FieldByName("PhysPersonDocumentType")->Value) == "1")
                { res = "������� ���������� ���������� ��������� ";
                if (!IBQuery->Fields->FieldByName("Serie")->IsNull)
                res += AnsiString(IBQuery->Fields->FieldByName("Serie")->Value)+ " ";
                if (!IBQuery->Fields->FieldByName("Number")->IsNull)
                res += AnsiString(IBQuery->Fields->FieldByName("Number")->Value);
                if ((!IBQuery->Fields->FieldByName("IssueDate")->IsNull) | (!IBQuery->Fields->FieldByName("PassportIssued")->IsNull))
                        res += ", �������� ";
                if (!IBQuery->Fields->FieldByName("IssueDate")->IsNull)
                        if (AnsiString(IBQuery->Fields->FieldByName("IssueDate")->Value) != "")
                        res += " " + AnsiString(IBQuery->Fields->FieldByName("IssueDate")->Value);
                if (!IBQuery->Fields->FieldByName("PassportIssued")->IsNull)
                        if (AnsiString(IBQuery->Fields->FieldByName("PassportIssued")->Value) != "")
                        res += " " + AnsiString(IBQuery->Fields->FieldByName("PassportIssued")->Value);
                if (!IBQuery->Fields->FieldByName("PassportIssuedCode")->IsNull)
                        if (AnsiString(IBQuery->Fields->FieldByName("PassportIssuedCode")->Value) != "")
                        res += ", ��� ������������� " + AnsiString(IBQuery->Fields->FieldByName("PassportIssuedCode")->Value);
                this->WordAdd(Prefix+"���������", res);
                AllDocuments += res+", ";
                };

             if (!IBQuery->Fields->FieldByName("PhysPersonDocumentType")->IsNull)
               if (AnsiString(IBQuery->Fields->FieldByName("PhysPersonDocumentType")->Value) == "4")
                { res = "������������ ������������� ";
                if (!IBQuery->Fields->FieldByName("Serie")->IsNull)
                res += AnsiString(IBQuery->Fields->FieldByName("Serie")->Value)+ " ";
                if (!IBQuery->Fields->FieldByName("Number")->IsNull)
                res += AnsiString(IBQuery->Fields->FieldByName("Number")->Value)+ ", �������� ";
                if (!IBQuery->Fields->FieldByName("IssueDate")->IsNull)
                        if (AnsiString(IBQuery->Fields->FieldByName("IssueDate")->Value) != "")
                        res += " " + AnsiString(IBQuery->Fields->FieldByName("IssueDate")->Value);
                if (!IBQuery->Fields->FieldByName("PassportIssued")->IsNull)
                        if (AnsiString(IBQuery->Fields->FieldByName("PassportIssued")->Value) != "")
                        res += " " + AnsiString(IBQuery->Fields->FieldByName("PassportIssued")->Value);
                if (!IBQuery->Fields->FieldByName("EndDate")->IsNull)
                        if (AnsiString(IBQuery->Fields->FieldByName("EndDate")->Value) != "")
                        res += ", ������ �� " + AnsiString(IBQuery->Fields->FieldByName("EndDate")->Value)+"�.";
                this->WordAdd(Prefix+"�������������������������", res);
                AllDocuments += res+", ";
                };

                IBQuery->Next();
            };
            //this->word->Add(Prefix+"������������","%"+Prefix+"������������%",AllDocuments);
         };
         };

AnsiString TAmirs::stardolznik(AnsiString OID1)
        {AnsiString res;
         TIBQuery *IBQueryAgency = new TIBQuery(Form1);
         IBQueryAgency->Active=false;
         IBQueryAgency->Database=IBDatabase;
         IBQueryAgency->Transaction=IBTransaction;
         IBQueryAgency->SQL->Clear();
         IBQueryAgency->SQL->Add("select * from \"Participant\" where (\"CaseMaterial\"= "+OID1+" and \"ParticipantKind\"=7)");
         IBQueryAgency->Active=true;
         IBQueryAgency->FetchAll();
         IBQueryAgency->First();
         if (!IBQueryAgency->Fields->FieldByName("OID")->IsNull)
                res = AnsiString(IBQueryAgency->Fields->FieldByName("OID")->Value); else res="";
         IBQueryAgency->Active=false;
         IBQueryAgency->SQL->Clear();
         IBQueryAgency->SQL->Add("select * from \"ParticipantAddress\" where (\"Participant\"="+res+" and \"IsBasic\"=1)");
         IBQueryAgency->Active=true;
         IBQueryAgency->FetchAll();
         IBQueryAgency->First();
         if (!IBQueryAgency->Fields->FieldByName("Address")->IsNull)
                res = AnsiString(IBQueryAgency->Fields->FieldByName("Address")->Value); else res="";
         return res;
         };

class TAgency
        {public:
        AnsiString name, address;
        TAgency (TAmirs * Amirs, int OID) {Amirs->Agency(OID, &name, &address);};
        };

void TAmirs::zapros(int i, TIBQuery IBQuery)
        {

        };

int TAmirs::zapros(int i)
        {


        };

void TAmirs::TabIndexStat(int i, int * vol, TStrings * l)
        {
        IBQuery->Database=IBDatabase;
        IBQuery->SQL->Clear();
        IBQuery->SQL->Add(ZaprosYear(this->slTextZapros->Strings[i], ":paramYear", IntToStr(this->Year)));
        IBQuery->Active=true;
        IBQuery->FetchAll();
        *vol=IBQuery->RecordCount;
        };
/*
void TAmirs::zaprosDel()
        {
        this->Connect();

        IBQuery->Database=IBDatabase;
        IBQuery->SQL->Clear();
        //IBQuery1->SQL->Add("select * from ZAPROS");
        TStringList * sl = new TStringList;
        switch (Amirs->VidDela)
                {
                case 1: GetTextFromResourse(sl, "zaprosUD"); break;
                case 2: GetTextFromResourse(sl, "zaprosGD"); break;
                case 3: GetTextFromResourse(sl, "zaprosAD"); break;
                case 4: GetTextFromResourse(sl, "zaprosIU"); break;
                case 5: GetTextFromResourse(sl, "zaprosIG"); break;
                };
        //AnsiString ZaprosYear(AnsiString Text, AnsiString ParYear, AnsiString Year)
        IBQuery->SQL->Add(ZaprosYear(sl->Text, ":paramYear", IntToStr(Amirs->Year)));
        IBQuery->Active=true;
        IBQuery->FetchAll();
        };   */

/*AnsiString TAmirs::zaprosDelInfo()
        {
        //MessageBox(NULL, "1", "TAmirs::zaprosDelInfo", MB_OK);
        AnsiString res="";  //cut
        if (!this->IBQuery->Fields->FieldByName("OID")->IsNull) this->OID = AnsiString(this->IBQuery->Fields->FieldByName("OID")->Value);
        if (!this->IBQuery->Fields->FieldByName("JudicalUid")->IsNull) this->UniqueNum = AnsiString(this->IBQuery->Fields->FieldByName("JudicalUid")->Value);
        //MessageBox(NULL, "2", "TAmirs::zaprosDelInfo", MB_OK);
        this->word->Add("���","%���%",this->UniqueNum);
        this->Participant();
        //MessageBox(NULL, "3", "TAmirs::zaprosDelInfo", MB_OK);
        for (int j=0; j<slParticipantPerson->Count; j++)
                {res += " OID="+this->OID;
                 res += " Participant="+slParticipant->Strings[j];
                 res += " ParticipantPerson="+slParticipantPerson->Strings[j];
                 res += " ParticipantKind="+slParticipantKind->Strings[j];
                 res += " ParticipantType="+slParticipantType->Strings[j];
                 };
        //MessageBox(NULL, "4", "TAmirs::zaprosDelInfo", MB_OK);
        for (int j=0; j<slParticipantPerson->Count; j++)
                {
                 };

        switch (this->VidDela)
                {
                case 1: {}; break;
                case 2: {
                         //��������� ������ ������
                         AnsiString Claim = GetFirstRecForTable ("Claim", "OID", "\"CivilCase\" = "+this->OID);
                         res = GetFirstRecForTable ("CivilCaseCategoryLink", "VnKod", "\"Claim\" = "+Claim);
                         if (res == "100110")  res="2"; if (res == "100210") res="3";
                         if (res == "210010")  res="114";
                         if (res == "3101103") res="168";
                         if (res == "510010")  res="203";
                         if (res == "520300")  res="209";
                         this->word->Add("������������","%������������%",res);
                         res = GetFirstRecForTable ("ClaimInAction", "Essence", "\"Claim\" = "+Claim);
                         this->word->Add("�������","%�������%",res);
                         //MessageBox(NULL, res.c_str(), "���������", MB_OK);
                        }; break;
                case 3: {
                         for (int j=0; j<slParticipantPerson->Count; j++)
                          {if (slParticipantKind->Strings[j]=="31")
                                {this->PhysPerson(slParticipantPerson->Strings[j], "���������������");};
                                //("select * from \"ParticipantAddress\" where (\"Participant\"="+res+" and \"IsBasic\"=1)")
                                //res = GetFirstRecForTable ("ParticipantAddress", "OID", "\"Participant\"=274179 and \"IsBasic\"=1");

                                this->Address(slParticipantPerson->Strings[j],"���������������");
                                this->PhysPersonDocumentGetAll(slParticipantPerson->Strings[j],"���������������");
                                };
                         //this->word->Add("���������������������������","%���������������������������%",this->DolznikDataRozd);
                         //this->word->Add("�������������������������������","%�������������������������������%",this->DolznikAddr);

                        }; break;
                case 4: {}; break;
                case 5: {
                //���������� �����������
        //this->Participant();
        res += " slParticipantPerson->Count="+IntToStr(slParticipantPerson->Count);
        //MessageBox(NULL, "5", "TAmirs::zaprosDelInfo", MB_OK);
        for (int j=0; j<slParticipantPerson->Count; j++)
                {res += " ParticipantPerson="+slParticipantPerson->Strings[j];
                 res += " ParticipantKind="+slParticipantKind->Strings[j];
                 res += " ParticipantType="+slParticipantType->Strings[j];
                 if (slParticipantType->Strings[j]=="6") res+="�����";
                 if (slParticipantType->Strings[j]=="11")
                        {res+="����"; AnsiString name; AnsiString addr; Amirs->Jurik(StrToInt(slParticipantPerson->Strings[j]),&name,&addr);
                                                res+=name;}
                 if (slParticipantKind->Strings[j]=="7") {this->fizik(slParticipantPerson->Strings[j]);
                                                        //this->DolznikAddr = this->address(slParticipant->Strings[j]);
                                                        Amirs->Num();
                                                        //this->DolznikAddr = this->address(this->stardolznik(this->StarOID));
                                                         this->Address(slParticipantPerson->Strings[j],"�������");
                                                        //this->DolznikAddr = this->stardolznik("155945");
                                                        };
                 //Zayvitel
                 if (slParticipantKind->Strings[j]=="38") {AnsiString addr; Amirs->Jurik(StrToInt(slParticipantPerson->Strings[j]),&Amirs->Zayvitel,&addr);};
                };
                //MessageBox(NULL, "6", "TAmirs::zaprosDelInfo", MB_OK);
         res+="����� ���� = "+this->Num();
                //MessageBox(NULL, "7", "TAmirs::zaprosDelInfo", MB_OK);
                //����� ���������� ����������
                }; break;
                };
        return res;
        };  */






void TAmirs::ExcelIndex(AnsiString FileName)
        {
       // this->zaprosDel();
        //Excel
        Variant app ;
        Variant books ;
        Variant book ;
        Variant sheet ;
        app = CreateOleObject("Excel.Application");
        app.OlePropertySet("Visible", 0);
        //app.OlePropertySet("Visible", 1);
        books = app.OlePropertyGet("Workbooks");
        books.Exec(Procedure("Add"));
        book = books.OlePropertyGet("item",1);
        sheet= book.OlePropertyGet("WorkSheets",1);
        sheet.OleProcedure("Activate");
        IBQuery->First();
        for (int i = 0; i < IBQuery->RecordCount; i++)
                {for (int j=0; j<3; j++)
                        {
                        Variant cell = sheet.OlePropertyGet("Cells", i+1, j+1);
                        cell.OlePropertySet("NumberFormat","@");
                        if (!IBQuery->Fields->FieldByNumber(j+1)->Value.IsNull())
                        cell.OlePropertySet("Value", WideString(String(IBQuery->Fields->FieldByNumber(j+1)->Value)));
                        }
                IBQuery->Next();
                }
        app.OlePropertyGet("WorkBooks",1).OleProcedure("SaveAs",WideString(FileName));
        app.OleProcedure("Quit");
        }





void TValidator::open (TExcelApp1 * excel1)
        {
        cols=0; rows=0;
        while (!excel1->CellGet(rows+1, 1).IsEmpty()) rows++;
        while (!excel1->CellGet(1, cols+1).IsEmpty()) cols++;
        matrix = new String * [rows];
        for (int i = 0; i < rows; i++) matrix[i] = new String [cols];
        for (int i = 0; i < rows; i++) for (int j = 0; j < cols; j++) matrix[i][j]=excel1->CellGet(i+1, j+1);
        };

int TValidator::poisk (AnsiString NomDela)
        {
        int j=0; //MessageBox(NULL, AnsiString(IntToStr(this->rows)+"-"+NomDela).c_str(), "Amirs", MB_OK | MB_ICONERROR);
        boolean poisk=true; while (j<this->rows & poisk) if (this->matrix[j][0]==NomDela) poisk=false; else j++;
        if (j< this->rows) return j; else return 0;
        };

bool TValidator::proverka (int vid, int row)
        {
        int j=1; int nom;
        switch (vid)
                        {case 0: nom=10; break;
                         case 1: nom=6; break;
                         case 2: nom=6; break;
                         case 3: nom=1; break;
                         case 4: nom=1; break;};
        boolean res=true; while (j<=nom & res) if (this->matrix[row][j].IsEmpty()) res=false; else j++;
        return res;
        };

void TAmirs::ProverkaFile()
        {
        TOpenDialog *OpenDialog1 = new TOpenDialog(Form1);
        OpenDialog1->InitialDir= GetSpecialFolderPath(CSIDL_PERSONAL);
        if (OpenDialog1->Execute())
        if (!OpenDialog1->FileName.IsEmpty()) {this->ProvFile=GetPathDirFromFileName(OpenDialog1->FileName); this->rf->Put(TRoamingFile::ProvFile,this->ProvFile);};
        };

void TAmirs::ProverkaStart(int year)
        {
        ProvRichEdit->Clear();
        if (bProvCheckListBoxGet("������� ���")) this->Proverka(year-1);
        this->Proverka(year);//this->YearProv);
        };

void TAmirs::ProverkaPrint()
        {this->ProvRichEdit->Print("�������� �������");
        };

TDateTime TAmirs::DateToExecution (TDateTime DateFinal, int vid)
        {
        switch (vid)
                        {case 0: return DateFinal+15; break; //���������
                         case 1: return DateFinal+30; break; //����������� �������
                         case 2: return DateFinal+15; break;//����������������
                         case 3: return DateFinal+15; break;//���������� �� ��
                         case 4: return DateFinal+15; break;//���������� �� ��
                        };
        };

void TAmirs::Proverka (int year)
        {
        this->Connect();
        if (bProvCheckListBoxGet("�������� �� ����� Excel"))
        {this->ProvFile = this->rf->Get(TRoamingFile::ProvFile);//MessageBox(NULL, this->ProvFile.c_str(), "Amirs", MB_OK | MB_ICONERROR);
         if (!this->ProvFile.IsEmpty())
         {//AnsiString ss=this->ProvFile+"\\��������"+StrToInt(year)+".xlsx";
         //MessageBox(NULL, ss.c_str(), "Amirs", MB_OK | MB_ICONERROR);
         TExcelApp1 excel1(this->ProvFile+"\\��������"+StrToInt(year)+".xlsx", false);
         excel1.OpenSheet("��������� ����");
         this->validator[0].open(& excel1);
         excel1.OpenSheet("����������� ����");
         this->validator[1].open(& excel1);
         excel1.OpenSheet("���� �� ��");
         this->validator[2].open(& excel1);
         excel1.OpenSheet("���������� �� ��");
         this->validator[3].open(& excel1);
         excel1.OpenSheet("���������� �� ��");
         this->validator[4].open(& excel1);
         excel1.Quit();
         };
        };
        //slDelta = new TStringList;
        //slProverka = new TStringList;

        TStringList * slNumX5[5];
        for (int i = 0; i < 6; i++) slNumX5[i] = new TStringList;

        //ProvRichEdit->Clear();
        AnsiString NotResult="��� ���������� ������������: ";
        AnsiString NotExecution="��� ���� ����������: ";
        AnsiString NotExcel="��� ������� � �������� � �����: ";
        for (int vid = 0; vid < 5; vid++)
                {
                 switch (vid)
                        {case 0: {this->GetDelaZapros(year,"zaprosUD");    NotResult +=" \n��������� ����: ";}   break;
                         case 1: {this->GetDelaZapros(year,"zaprosGDIsk"); NotResult +=" \n����������� ����: ";} break;
                         case 2: {this->GetDelaZapros(year,"zaprosAD");    NotResult +=" \n���� �� ��: ";}       break;
                         case 3: {this->GetDelaZapros(year,"zaprosIU");    NotResult +=" \n���������� �� ��: ";} break;
                         case 4: {this->GetDelaZapros(year,"zaprosIG");    NotResult +=" \n���������� �� ��: ";} break;
                        };
                 switch (vid)
                        {case 0: NotExecution+=" \n��������� ����: ";   break;
                         case 1: NotExecution+=" \n����������� ����: "; break;
                         case 2: NotExecution+=" \n���� �� ��: ";       break;
                         case 3: NotExecution+=" \n���������� �� ��: "; break;
                         case 4: NotExecution+=" \n���������� �� ��: "; break;
                        };
                 switch (vid)
                        {case 0: NotExcel+=" \n��������� ����: ";   break;
                         case 1: NotExcel+=" \n����������� ����: "; break;
                         case 2: NotExcel+=" \n���� �� ��: ";       break;
                         case 3: NotExcel+=" \n���������� �� ��: "; break;
                         case 4: NotExcel+=" \n���������� �� ��: "; break;
                        };
                 NotResult+="��������� "+IntToStr(this->IBQuery1->RecordCount)+" ���: ";
                 NotExecution+="��������� "+IntToStr(this->IBQuery1->RecordCount)+" ���: ";
                 AnsiString sdata,OID, Num, res, ispol; TDateTime SittingDate;
                 for (int i = 0; i < this->IBQuery1->RecordCount; i++)
                    {
                    Boolean bResReview=false, bSittingDate=false; sdata="";
                    if (!this->IBQuery1->Fields->FieldByName("SittingDate")->IsNull)
                        {bSittingDate=true; sdata=this->IBQuery1->Fields->FieldByName("SittingDate")->Value; SittingDate=StringToDT(sdata);}
                    if (!this->IBQuery1->Fields->FieldByName("OID")->IsNull) OID=this->IBQuery1->Fields->FieldByName("OID")->Value;
                    if (!this->IBQuery1->Fields->FieldByName("NUM")->IsNull) Num=this->IBQuery1->Fields->FieldByName("NUM")->Value;
                    if (!this->IBQuery1->Fields->FieldByName("ResultOfTheReview")->IsNull)
                        {res=this->IBQuery1->Fields->FieldByName("ResultOfTheReview")->Value; if (!res.IsEmpty()) bResReview=true; };
                    if (!this->IBQuery1->Fields->FieldByName("ExecutionDate")->IsNull) ispol=this->IBQuery1->Fields->FieldByName("ExecutionDate")->Value;
                    //if (vid==2) { ProvRichEdit->Lines->Add(Num+"__________");
                    //             if (bResReview) ProvRichEdit->Lines->Add("bResReview=true"); else ProvRichEdit->Lines->Add("bResReview=false");
                    //             if (bSittingDate) ProvRichEdit->Lines->Add("bSittingDate=true"); else ProvRichEdit->Lines->Add("bSittingDate=false");
                    //             ProvRichEdit->Lines->Add(sdata);
                    //             if (((SittingDate<=Now()) & !bResReview) | (!bSittingDate & !bResReview)) ProvRichEdit->Lines->Add("��"); else ProvRichEdit->Lines->Add("���");
                    //            };
                    if (((SittingDate<Now()) & !bResReview) | (!bSittingDate & !bResReview)) {NotResult += " "+Num+" ("+sdata.SubString(1,10)+"); ";//ProvRichEdit->Lines->Add("1111-���"); else ProvRichEdit->Lines->Add("1111-��� ���������");
                                                                               //if (!bSittingDate & !bResReview) NotResult += " (!bSittingDate & !bResReview)==true ";
                                                                               //if (SittingDate<Now()) NotResult += " (SittingDate<Now==true "; else NotResult += " (SittingDate<Now==false ";;
                                                                               //if (bResReview) NotResult += " bResReview=true; "; else NotResult += " bResReview=false; ";
                                                                               //if (bSittingDate) NotResult += " bSittingDate=true "; else NotResult += " bSittingDate=false ";
                                                                                };
                    //if (bResReview) ProvRichEdit->Lines->Add("bResReview-��� ���������"); else ProvRichEdit->Lines->Add("bResReview-��� ����������");
                    //if (bSittingDate) ProvRichEdit->Lines->Add("bResSittingDate-��� ���������"); else ProvRichEdit->Lines->Add("bResSittingDate-��� ����������");
                    //ProvRichEdit->Lines->Add(Num +" sdata="+ sdata+" res="+ res+" ispol="+ ispol);
                    //NotResult=+ ���� ��������� ���� ��� ����������

                    //if (j) {if (!this->validator[vid].proverka(vid,j)) ProvRichEdit->Lines->Add(this->GetQuery1TextField("NUM")+ "___���");} else slProverka->Add(this->GetQuery1TextField("NUM")+ "++++���");
                   
                    slNumX5[vid]->Add(Num);
                    //��� ���� ����������
                    if (bProvCheckListBoxGet("��� ���� ����������"))
                     {
                      if (( (DateToExecution(SittingDate, vid)<=Now()) & bResReview)) NotExecution += " "+Num+" ("+sdata.SubString(1,10)+"); ";
                     };
                     //�������� �� ����� Excel
                     if (bProvCheckListBoxGet("�������� �� ����� Excel"))
                     if (vid==0 | vid==1 | vid==2 | vid==3 | vid==4)
                     {
                      if (( (DateToExecution(SittingDate, vid)<=Now()) & bResReview))
                      {
                      bool ResProv=true;
                      int j= this->validator[vid].poisk(Num);
                      if (j) for (int ii = 1; (ii < this->validator[vid].cols) & ResProv; ii++) ResProv=!this->validator[vid].matrix[j][ii].IsEmpty();
                         else ResProv=false;
                      if (!ResProv) NotExcel += " "+Num+" ("+sdata.SubString(1,10)+"); ";
                      };
                     };
                    this->IBQuery1->Next();
                    };//����� for (int i = 0; i < this->IBQuery1->RecordCount; i++)

                }; //for (int vid = 0; vid < 5; vid++)

        ProvRichEdit->Lines->Add(IntToStr(year)+" ���");
        if (bProvCheckListBoxGet("��� ����������"))     {ProvRichEdit->Lines->Add(NotResult);};
        if (bProvCheckListBoxGet("��� ���� ����������")){ProvRichEdit->Lines->Add(NotExecution);};
        if (bProvCheckListBoxGet("�������� �� ����� Excel")){ProvRichEdit->Lines->Add(NotExcel);};
        if (bProvCheckListBoxGet("�������� ����� ������� �� ����� Excel"))
                {int total=0; AnsiString NotShtrafExcel; TStringList * slProverka = new TStringList;
                this->ProvFile = this->rf->Get(TRoamingFile::ProvFile);
                if (!this->ProvFile.IsEmpty())
                 {//AnsiString ss=this->ProvFile+"\\��������"+StrToInt(year)+".xlsx";
                 //MessageBox(NULL, ss.c_str(), "Amirs", MB_OK | MB_ICONERROR);
                 TExcelApp1 excel1(this->ProvFile+"\\�����"+StrToInt(year)+".xlsx", false);
                 excel1.OpenSheet("����1");
                 int rows=0;
                 while (!excel1.CellGet(rows+1, 1).IsEmpty()) rows++;
                 for (int i = 0; i < rows; i++) slProverka->Add(excel1.CellGet(i+1, 1));
                 excel1.Quit();
                 this->GetDelaZapros(year,"zaprosAD_SH");
                  for (int i = 0; i < this->IBQuery1->RecordCount; i++)
                    {AnsiString Num; if (!this->IBQuery1->Fields->FieldByName("NUM")->IsNull) Num=this->IBQuery1->Fields->FieldByName("NUM")->Value;
                     bool pres=false;
                     for (int j = 0; j <slProverka->Count ; j++)
                     {int pos=slProverka->Strings[j].Pos(Num);
                     if (pos) {pres=true; j=slProverka->Count;};};
                    if (!pres) {NotShtrafExcel += Num+"; "; total++;};
                    this->IBQuery1->Next();
                    };
                 };
                ProvRichEdit->Lines->Add("\n��� ������ �� ������ ������ �� "+IntToStr(total)+" �����: "+NotShtrafExcel);
                };
        if (bProvCheckListBoxGet("�������� ���� ����� �� ����� Excel"))
                {int total=0; AnsiString NotShtrafExcel; TStringList * slProverka = new TStringList;
                this->ProvFile = this->rf->Get(TRoamingFile::ProvFile);
                if (!this->ProvFile.IsEmpty())
                 {//AnsiString ss=this->ProvFile+"\\��������"+StrToInt(year)+".xlsx";
                 //MessageBox(NULL, ss.c_str(), "Amirs", MB_OK | MB_ICONERROR);
                 TExcelApp1 excel1(this->ProvFile+"\\��"+StrToInt(year)+".xlsx", false);
                 excel1.OpenSheet("����1");
                 int rows=0;
                 while (!excel1.CellGet(rows+1, 1).IsEmpty()) rows++;
                 for (int i = 0; i < rows; i++) slProverka->Add(excel1.CellGet(i+1, 1));
                 excel1.Quit();
                 this->GetDelaZapros(year,"zaprosAD_OR");
                  for (int i = 0; i < this->IBQuery1->RecordCount; i++)
                    {AnsiString Num; if (!this->IBQuery1->Fields->FieldByName("NUM")->IsNull) Num=this->IBQuery1->Fields->FieldByName("NUM")->Value;
                     bool pres=false;
                     for (int j = 0; j <slProverka->Count ; j++)
                     {int pos=slProverka->Strings[j].Pos(Num);
                     if (pos) {pres=true; j=slProverka->Count;};};
                    if (!pres) {NotShtrafExcel += Num+"; "; total++;};
                    this->IBQuery1->Next();
                    };
                 };
                ProvRichEdit->Lines->Add("\n��� ������ �� ���� ������� �� "+IntToStr(total)+" �����: "+NotShtrafExcel);
                };
        //excel2.OpenSheet("������");
       // for (int i = 0; i < this->slProverka->Count; i++) {} //excel2.CellSet(i+1, 1, this->slProverka->Strings[i]);
         //       };
        if (bProvCheckListBoxGet("�������� ������ ����� Excel"))
               {
                 TExcelApp1 excel2(false);
                 String s[]= {"��������� ����", "����������� ����", "���� �� ��", "���������� �� ��", "���������� �� ��", "��������", "������"};
                 excel2.NewBook(7, s);
                 for (int vid = 0; vid < 5; vid++)
                 {switch (vid)
                        {case 0: {excel2.OpenSheet("��������� ����");
                                String s1[]= {"�����", "�����", "����������", "���������", "��������", "��������", "����������", "�����6", "���������", "�������", "������"};
                                for (int i = 0; i < 11; i++) excel2.CellSet(1, i+1, s1[i]); //������ ������
                                }; break;
                         case 1: {excel2.OpenSheet("����������� ����");
                                String s1[]= {"�����", "�����", "�����������", "���������", "��������", "��������", "����������"};
                                for (int i = 0; i < 8; i++) excel2.CellSet(1, i+1, s1[i]); //������ ������
                                }; break;
                         case 2: {excel2.OpenSheet("���� �� ��");
                                String s1[]= {"�����", "�����", "����������", "���������", "��������", "����������", "�����"};
                                for (int i = 0; i < 8; i++) excel2.CellSet(1, i+1, s1[i]); //������ ������
                                }; break;
                         case 3: {excel2.OpenSheet("���������� �� ��");
                                String s1[]= {"�����", "��������"};
                                for (int i = 0; i < 3; i++) excel2.CellSet(1, i+1, s1[i]); //������ ������
                                }; break;
                         case 4: {excel2.OpenSheet("���������� �� ��");
                                String s1[]= {"�����", "��������"};
                                for (int i = 0; i < 3; i++) excel2.CellSet(1, i+1, s1[i]); //������ ������
                                }; break;
                        };
                 for (int i = 0; i < slNumX5[vid]->Count; i++)
                        {excel2.CellSet(i+2, 1, slNumX5[vid]->Strings[i]);
                         int j= this->validator[vid].poisk(slNumX5[vid]->Strings[i]);
                         //AnsiString ss; ss=" "+slNumX5[vid]->Strings[i]+" "+IntToStr(j);
                         //MessageBox(NULL, ss.c_str(), "Amirs", MB_OK);
                         if (j) for (int ii = 1; ii < this->validator[vid].cols; ii++) excel2.CellSet(i+2, ii+1, this->validator[vid].matrix[j][ii]);
                         else for (int ii = 1; ii < this->validator[vid].cols; ii++) excel2.CellSet(i+2, ii+1, "");
                         };
                 excel2.Zagolovok();
                 };//for (int vid = 0; vid < 5; vid++)
               excel2.Show();
               };
        };

//�������
        void TAmirs::SelectStart(int year, int vid, TRichEdit * RichEdit1)
        {
        SelectRichEdit=RichEdit1;
        this->Connect();
        switch (vid)
                {
                case 0: this->GetDelaZapros(year,"zaprosKAS");    break;
                case 1: this->GetDelaZapros(year,"zaprosAD_OR");    break;
                case 2: this->GetDelaZapros(year,"zaprosAD_SH"); break;
                case 3: this->GetDelaZapros(year,"zaprosAD_LP");    break;
                case 4: this->GetDelaZapros(year,"zaprosIU");    break;
                case 5: this->GetDelaZapros(year,"zaprosIG");    break;
                };
        AnsiString sdata,OID, Num, Names, res, ispol;
        AnsiString sSelect="";
        sSelect+="����� "+IntToStr(this->IBQuery1->RecordCount)+" ���: ";
        for (int i = 0; i < this->IBQuery1->RecordCount; i++)
                {
                if (!this->IBQuery1->Fields->FieldByName("NUM")->IsNull) Num=this->IBQuery1->Fields->FieldByName("NUM")->Value;
                if (!this->IBQuery1->Fields->FieldByName("Names")->IsNull) Names=this->IBQuery1->Fields->FieldByName("Names")->Value;
                sSelect+= "\n"+Num+", "+Names+";";
                this->IBQuery1->Next();
                };
        SelectRichEdit->Lines->Add(sSelect);
        };

void TAmirs::SelectPrint() {this->SelectRichEdit->Print("������� ���");};

