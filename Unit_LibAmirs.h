#define NO_WIN32_LEAN_AND_MEAN //SHGetSpecialFolderPath
//---------------------------------------------------------------------------
#ifndef Unit_LibAmirs
#define Unit_LibAmirs
//---------------------------------------------------------------------------
#endif

#include <vcl.h>
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
#include <CheckLst.hpp>
#include <IBCustomDataSet.hpp>
#include "Excel_2K_SRVR.h"
#include "Word_2K_SRVR.h"
#include <OleServer.hpp>

#include "Unit_LibCommon.cpp"
#include "Unit_LibFIO.h"
#include "Unit_LibExcel.h"
//#include "Unit_LibExcel.cpp"


class TValidator
        {public:
         String **matrix;
         int cols; int rows;
         TValidator() {cols=0; rows=0;};
         void open (TExcelApp1 * excel1);
         String delta ();
         int poisk (AnsiString NomDela);
         bool proverka (int vid, int row);};

//������ ��� ���� ������ �� ��� ��� ��������� � �����������
//� ������������ ������� ������ PhisPerson, ������� ������� ������ FIO
//���������� TextToWord + ObrazDoc, ������������ ��� ����������
class TAmirs
        {
        public:
        TRoamingFile * rf; //1111 ����� ���� ��������
        TForm * Form1;
        AnsiString NameApp;

        TAmirs(TForm * Form); //�������� �����������
        //������������ ��� �������� ����������
        TAmirs1(TMemo * Memo);//�������
        //������
        //��������
        TAmirs(TMemo * Memo, TProgressBar ProgressBar);//����������� ��������
        TCheckListBox * ProvChecklistbox; //������� �� ��������
        AnsiString ProvCheckListBoxGet() {AnsiString res; res=""; for(int i = 0; i < ProvChecklistbox->Items->Count; i++)
                               {if (ProvChecklistbox->Checked[i]) res+="1"; else res+="0";}; return res;}; //������� � ����������
        void ProvCheckListBoxPut(AnsiString txt1) {for(int i = 0; i < txt1.Length(); i++)
                        if (txt1[i+1]=='1') ProvChecklistbox->State[i]=cbChecked; else ProvChecklistbox->State[i]=cbUnchecked;};//������� �� ����������
        bool bProvCheckListBoxGet(AnsiString txt1) {bool res=false; for(int i = 0; i < ProvChecklistbox->Items->Count; i++)
                                       if (ProvChecklistbox->Items->Strings[i]==txt1) {res=ProvChecklistbox->Checked[i]; break;}; return res;}; //������ �� �����
        TRichEdit * ProvRichEdit;
        
        AnsiString ProvFile;
        void ProverkaFile();
        void ProverkaStart(int year);
        void ProverkaPrint();
        //�������
        TRichEdit * SelectRichEdit;
        void SelectStart(int year, int vid, TRichEdit * RichEdit1);
        void SelectPrint();
        TAmirs(TForm * Form, AnsiString NameApp1); //������� ����������� ��������
        TStringList * slRoamingFile; //���� ��������
        //�������
        AnsiString ProfileFile;
        AnsiString ProfileFileLast;
        AnsiString ProfileShablon;
        void ProfileSelect();
        TStringList * slProfileFile;

        void ProfileZagruz();
        void ActivateTabSheetShablon(TDBNavigator * DBN1);

        int Errors;
        TStringList * slErrors;
        void ErrorsInMemo(TMemo * Memo){if (this->Errors) Memo->Lines = slErrors;};
        int Year; //, YearProv;
        int VidDela;
        //int TabSheet; //����� �������
        AnsiString OID; //
        AnsiString UniqueNum; //
        AnsiString NomDela; //


        TIBDatabase * IBDatabase;
        TIBTransaction * IBTransaction;
        TIBQuery * IBQuery; //�������� ������, ��������� � ���������� ����������
        TIBQuery * IBQuery1; //��������������� ������ ��� ������� ����������  + GetDelaZapros
        TDataSource * DataSource;

        void Connect();
        void Reconnect(){};
        void GetDela();                            //��������� ������ ��� � IBQuery
        void GetDelaZapros(int year, AnsiString NameZapros); //��������� ������ ��� � IBQuery1 �� ����� ������� � ����� ��������
        WideString GetQuery1TextField (AnsiString FieldName) {if (!this->IBQuery1->Fields->FieldByName(FieldName)->IsNull) return WideString(this->IBQuery1->Fields->FieldByName(FieldName)->Value); else return "";};
        void Search(AnsiString Text);

        //TStringList * WordNazv;
        TStringList * WordName;
        TStringList * WordVal;
        void WordClear();
        //void WordAdd(AnsiString Nazv1, AnsiString Name1, AnsiString Val1);
        void WordAdd(AnsiString Name1, AnsiString Val1);
        void SplitSLToWord (TStringList * SL);
        void GetTextToWord(); //����� �������� ������ ���������� ����� ��� ���
        void TextToMemo(TMemo * Memo); //WordName+WordVal -> Memo
        void TextToWord(); //WordName+WordVal -> Word
        WideString FileNameShablon;
        int TextToWordZavershenie; //0-�������� �������� ����, 1-�����������, 2-��������� � �������
        //������� ����
        AnsiString GetFirstRecForTable (AnsiString Table, AnsiString Field, AnsiString Where);
        void GetAllRecForTable (TIBQuery * IBQuery, AnsiString Table, AnsiString Where);

        //������� �� ����
        void GetParticipant();
        TStringList * slParticipant;
        TStringList * slParticipantPerson;
        TStringList * slParticipantKind;
        TStringList * slParticipantType;
        TStringList * slIstecRod;
        TStringList * slOtvetDat;

        TStringList * slName; //����� �������
        TStringList * slTextZapros; // ������ �������� �������
        void zapros(int i, TIBQuery IBQuery); //������� �������
        int zapros(int i);
        //void zaprosDel();
        AnsiString DolznikAddr;
        AnsiString DolznikDataRozd;
        AnsiString AdminAddr;
        AnsiString AdminDataRozd;
        AnsiString StarOID;
        AnsiString Zayvitel;
        //AnsiString zaprosDelInfo();

        //void (*callback)();
        //void setcallb(void (*callback1)()) {this->callback=callback1;};
        //void z1(TForm * Form) {};
        void TabIndexStat(int i, int * vol, TStrings * l);
        void ExcelIndex(AnsiString FileName);
        void Agency(int OID, AnsiString *Name, AnsiString *Address);
        void Jurik(int OID, AnsiString *Name, AnsiString *Address);
        void fizik(AnsiString OID); //������ ������ ���� ����� PhysPerson
        void PhysPerson(AnsiString OID, AnsiString Prefix);
        void JuridicalPerson(AnsiString OID, AnsiString Prefix);
        AnsiString stardolznik(AnsiString OID1);
        AnsiString address(AnsiString OID);
        void Address(AnsiString Subject, AnsiString Prefix);
        void PhysPersonDocumentGetAll(AnsiString Subject, AnsiString Prefix);
        AnsiString Num();
        TValidator validator[5]; //�������� ���
        TStringList * slDelta;
        TStringList * slProverka;
        void Proverka (int year);
        TDateTime DateToExecution (TDateTime DateFinal, int vid);
        };

TAmirs * Amirs;


