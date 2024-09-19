//---------------------------------------------------------------------------
#define NO_WIN32_LEAN_AND_MEAN
#include <vcl.h>
#pragma hdrstop

#include "UnitAmirsDop.h"
#include "Unit_LibAmirs.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TForm1 *Form1;
//---------------------------------------------------------------------------
__fastcall TForm1::TForm1(TComponent* Owner)
        : TForm(Owner)
{
}
//---------------------------------------------------------------------------
void __fastcall TForm1::FormActivate(TObject *Sender)
{
Amirs = new TAmirs(Form1);
//Amirs->Connect();
CheckListBox1->Items->Text="������� ���\n��� ����������\n��� ���� ����������\n�������� �� ����� Excel\n�������� ������ ����� Excel\n�������� ����� ������� �� ����� Excel\n�������� ���� ����� �� ����� Excel";
Amirs->ProvRichEdit = ProvRichEdit;
Amirs->ProvChecklistbox = CheckListBox1;
Amirs->ProvCheckListBoxPut(Amirs->rf->Get(TRoamingFile::ProvZadanie));
//Shablon
switch (PageControl1->TabIndex)
                {
                case 0: Amirs->ActivateTabSheetShablon(DBNavigatorShablon); break;
                };
Amirs->ActivateTabSheetShablon(DBNavigatorShablon);
}
//---------------------------------------------------------------------------
void __fastcall TForm1::Button2Click(TObject *Sender)
{
Amirs->ProfileSelect();
}
//---------------------------------------------------------------------------
void __fastcall TForm1::Button1Click(TObject *Sender)
{
PageControlProv->ActivePageIndex=1;
ProvRichEdit->Lines->Clear();
Word year, month, day;
DateTimePickerProv->Date.DecodeDate(&year,&month, &day);
Amirs->ProverkaStart(year);

//ProvRichEdit->Lines->Add(Amirs->ProvCheckListBoxGet());
//ProvZadanie
//this->ProvFile = this->rf->Get(TRoamingFile::ProvFile);
//Amirs->rf->Put(TRoamingFile::ProvZadanie,Amirs->ProvCheckListBoxGet());
//Amirs->ProvCheckListBoxPut(Amirs->rf->Get(TRoamingFile::ProvZadanie));
//Amirs->ProvChecklistbox->State[0]=cbChecked;
/*ProvRichEdit->Lines->Clear();
if (Amirs->bProvCheckListBoxGet("�����1")) ProvRichEdit->Lines->Add("�����1=��"); else ProvRichEdit->Lines->Add("�����1=���");
if (Amirs->bProvCheckListBoxGet("�����2")) ProvRichEdit->Lines->Add("�����2=��"); else ProvRichEdit->Lines->Add("�����2=���");
if (Amirs->bProvCheckListBoxGet("�����3")) ProvRichEdit->Lines->Add("�����3=��"); else ProvRichEdit->Lines->Add("�����3=���");
if (Amirs->bProvCheckListBoxGet("��������������Excel�����")) ProvRichEdit->Lines->Add("�����4=��"); else ProvRichEdit->Lines->Add("�����4=���");
*/
}
//---------------------------------------------------------------------------
void __fastcall TForm1::Button3Click(TObject *Sender)
{
Amirs->ProverkaFile();
}
//---------------------------------------------------------------------------
void __fastcall TForm1::Button4Click(TObject *Sender)
{
 Amirs->ProverkaPrint();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::CheckListBox1ClickCheck(TObject *Sender)
{
Amirs->rf->Put(TRoamingFile::ProvZadanie,Amirs->ProvCheckListBoxGet());
}

//---------------------------------------------------------------------------

void __fastcall TForm1::TabSheetProvShow(TObject *Sender)
{
PageControlProv->ActivePageIndex=0;
DateTimePickerProv->Date=Now();
DateTimePickerProv->MaxDate=Now();
PageControlProv->ActivePageIndex=0;
//Amirs->ProvRichEdit = ProvRichEdit;
//Amirs->ProvChecklistbox = CheckListBox1;
//Amirs->ProvCheckListBoxPut(Amirs->rf->Get(TRoamingFile::ProvZadanie));
//MessageBox(NULL, /*this->ProvFile.c_str()*/ "TabSheetProvShow", "Amirs", MB_OK);
}
//---------------------------------------------------------------------------

void __fastcall TForm1::TabSheetSelectShow(TObject *Sender)
{
RichEditSelect->Clear();
ComboBoxSelect->Items->Text="���������������� ���� (���)\n���� �� ��. ���� ������\n���� �� ��. �����\n���� �� ��. ������� ���� ���������� ��";
DateTimePickerSelect->Date=Now();
DateTimePickerSelect->MaxDate=Now();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::ButtonSelectClick(TObject *Sender)
{
Word year, month, day;
DateTimePickerSelect->Date.DecodeDate(&year,&month, &day);
RichEditSelect->Clear();
Amirs->SelectStart(year, ComboBoxSelect->ItemIndex, RichEditSelect);
}
//---------------------------------------------------------------------------

void __fastcall TForm1::ButtonSelectProfileClick(TObject *Sender)
{
Amirs->ProfileSelect();        
}
//---------------------------------------------------------------------------

void __fastcall TForm1::ButtonPrintSelectClick(TObject *Sender)
{
Amirs->SelectPrint();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::TabSheetShablonShow(TObject *Sender)
{
DateTimePickerShablon->Date=Now();
DateTimePickerShablon->MaxDate=Now();
ComboBoxShablon->Items->Text="��������� ����\n����������� ����\n�����. ���� (���)\n���� �� �����. ��������.\n�����. ��������� ���\n�����. ����������� ���";
ComboBoxShablon->ItemIndex=0;
}
//---------------------------------------------------------------------------

void __fastcall TForm1::Button5Click(TObject *Sender)
{
ListBox1->Items->Add("cfm nmcvorm,fvtg[vby;.[.bv'g/bv'/ghhnuuuuuuuuuuxscp");
ListBox1->Items->Add("cfvgbyuttttttttttttttttttttttttttttttttttttttttttttthnjumi");
}
//---------------------------------------------------------------------------

