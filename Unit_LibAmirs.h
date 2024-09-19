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

//Запрос дел Инфо убрать Из нее все перенести в Партиципант
//В партиципанте сделать вызовы PhisPerson, отттуда сделать вызовы FIO
//Объединить TextToWord + ObrazDoc, использовать код завершения
class TAmirs
        {
        public:
        TRoamingFile * rf; //1111 Новый файл роуминга
        TForm * Form1;
        AnsiString NameApp;

        TAmirs(TForm * Form); //основной конструктор
        //Конструкторы для создания интерфейса
        TAmirs1(TMemo * Memo);//Шаблоны
        //Пакеты
        //Проверка
        TAmirs(TMemo * Memo, TProgressBar ProgressBar);//Конструктор проверка
        TCheckListBox * ProvChecklistbox; //задание на проверку
        AnsiString ProvCheckListBoxGet() {AnsiString res; res=""; for(int i = 0; i < ProvChecklistbox->Items->Count; i++)
                               {if (ProvChecklistbox->Checked[i]) res+="1"; else res+="0";}; return res;}; //задание в сохраненку
        void ProvCheckListBoxPut(AnsiString txt1) {for(int i = 0; i < txt1.Length(); i++)
                        if (txt1[i+1]=='1') ProvChecklistbox->State[i]=cbChecked; else ProvChecklistbox->State[i]=cbUnchecked;};//задание из сохраненки
        bool bProvCheckListBoxGet(AnsiString txt1) {bool res=false; for(int i = 0; i < ProvChecklistbox->Items->Count; i++)
                                       if (ProvChecklistbox->Items->Strings[i]==txt1) {res=ProvChecklistbox->Checked[i]; break;}; return res;}; //нажато по имени
        TRichEdit * ProvRichEdit;
        
        AnsiString ProvFile;
        void ProverkaFile();
        void ProverkaStart(int year);
        void ProverkaPrint();
        //Выборки
        TRichEdit * SelectRichEdit;
        void SelectStart(int year, int vid, TRichEdit * RichEdit1);
        void SelectPrint();
        TAmirs(TForm * Form, AnsiString NameApp1); //убирать конструктор рудимент
        TStringList * slRoamingFile; //файл роуминга
        //Профиль
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
        //int TabSheet; //Номер вкладки
        AnsiString OID; //
        AnsiString UniqueNum; //
        AnsiString NomDela; //


        TIBDatabase * IBDatabase;
        TIBTransaction * IBTransaction;
        TIBQuery * IBQuery; //Основной запрос, связанный с элементами управления
        TIBQuery * IBQuery1; //Вспомогательный запрос для выборок статистики  + GetDelaZapros
        TDataSource * DataSource;

        void Connect();
        void Reconnect(){};
        void GetDela();                            //Получение списка дел в IBQuery
        void GetDelaZapros(int year, AnsiString NameZapros); //Получение списка дел в IBQuery1 по имени запроса в файле ресурсов
        WideString GetQuery1TextField (AnsiString FieldName) {if (!this->IBQuery1->Fields->FieldByName(FieldName)->IsNull) return WideString(this->IBQuery1->Fields->FieldByName(FieldName)->Value); else return "";};
        void Search(AnsiString Text);

        //TStringList * WordNazv;
        TStringList * WordName;
        TStringList * WordVal;
        void WordClear();
        //void WordAdd(AnsiString Nazv1, AnsiString Name1, AnsiString Val1);
        void WordAdd(AnsiString Name1, AnsiString Val1);
        void SplitSLToWord (TStringList * SL);
        void GetTextToWord(); //Здесь основная логика заполнения полей для дел
        void TextToMemo(TMemo * Memo); //WordName+WordVal -> Memo
        void TextToWord(); //WordName+WordVal -> Word
        WideString FileNameShablon;
        int TextToWordZavershenie; //0-оставить открытым ворд, 1-распечатать, 2-сохранить и закрыть
        //Обслуга базы
        AnsiString GetFirstRecForTable (AnsiString Table, AnsiString Field, AnsiString Where);
        void GetAllRecForTable (TIBQuery * IBQuery, AnsiString Table, AnsiString Where);

        //Выборки из базы
        void GetParticipant();
        TStringList * slParticipant;
        TStringList * slParticipantPerson;
        TStringList * slParticipantKind;
        TStringList * slParticipantType;
        TStringList * slIstecRod;
        TStringList * slOtvetDat;

        TStringList * slName; //имена выборок
        TStringList * slTextZapros; // тексты запросов выборок
        void zapros(int i, TIBQuery IBQuery); //сделать выборку
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
        void fizik(AnsiString OID); //убрать вместо него будет PhysPerson
        void PhysPerson(AnsiString OID, AnsiString Prefix);
        void JuridicalPerson(AnsiString OID, AnsiString Prefix);
        AnsiString stardolznik(AnsiString OID1);
        AnsiString address(AnsiString OID);
        void Address(AnsiString Subject, AnsiString Prefix);
        void PhysPersonDocumentGetAll(AnsiString Subject, AnsiString Prefix);
        AnsiString Num();
        TValidator validator[5]; //Проверка дел
        TStringList * slDelta;
        TStringList * slProverka;
        void Proverka (int year);
        TDateTime DateToExecution (TDateTime DateFinal, int vid);
        };

TAmirs * Amirs;


