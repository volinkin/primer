//-----____----------------------------------------------------------------------
//#define NO_WIN32_LEAN_AND_MEAN
#ifndef UnitAmirsDopH
#define UnitAmirsDopH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include <ComCtrls.hpp>
#include <Grids.hpp>
#include <ValEdit.hpp>
#include <CheckLst.hpp>
#include <DBCtrls.hpp>
#include <ExtCtrls.hpp>

//---------------------------------------------------------------------------
class TForm1 : public TForm
{
__published:	// IDE-managed Components
        TPageControl *PageControl1;
        TTabSheet *TabSheetShablon;
        TTabSheet *TabSheet2;
        TTabSheet *TabSheetProv;
        TButton *Button1;
        TButton *Button2;
        TDateTimePicker *DateTimePickerProv;
        TTabSheet *TabSheetSelect;
        TButton *Button4;
        TPageControl *PageControlProv;
        TTabSheet *TabSheetProvZad;
        TTabSheet *TabSheet6;
        TRichEdit *ProvRichEdit;
        TCheckListBox *CheckListBox1;
        TButton *Button3;
        TDateTimePicker *DateTimePickerSelect;
        TComboBox *ComboBoxSelect;
        TButton *ButtonSelect;
        TButton *ButtonPrintSelect;
        TButton *ButtonSelectProfile;
        TRichEdit *RichEditSelect;
        TDateTimePicker *DateTimePickerShablon;
        TButton *ButtonWordShablon;
        TComboBox *ComboBoxShablon;
        TEdit *Edit1;
        TDBNavigator *DBNavigatorShablon;
        TButton *ButtonSelectShablon;
        TTabSheet *TabSheetSelectProfile;
        TButton *ButtonRefreshShablon;
        TMemo *Memo1;
        TProgressBar *ProgressBarShablon;
        TListView *ListView1;
        TButton *Button5;
        TListBox *ListBox1;
        void __fastcall FormActivate(TObject *Sender);
        void __fastcall Button2Click(TObject *Sender);
        void __fastcall Button1Click(TObject *Sender);
        void __fastcall Button3Click(TObject *Sender);
        void __fastcall Button4Click(TObject *Sender);
        void __fastcall CheckListBox1ClickCheck(TObject *Sender);
        void __fastcall TabSheetProvShow(TObject *Sender);
        void __fastcall TabSheetSelectShow(TObject *Sender);
        void __fastcall ButtonSelectClick(TObject *Sender);
        void __fastcall ButtonSelectProfileClick(TObject *Sender);
        void __fastcall ButtonPrintSelectClick(TObject *Sender);
        void __fastcall TabSheetShablonShow(TObject *Sender);
        void __fastcall Button5Click(TObject *Sender);
private:	// User declarations
public:		// User declarations
        __fastcall TForm1(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TForm1 *Form1;
//---------------------------------------------------------------------------
#endif
