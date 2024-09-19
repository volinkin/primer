//---------------------------------------------------------------------------
//#pragma hdrstop
//#include "Excel_2K_SRVR.h"
#include "Unit_LibExcel.h"
//---------------------------------------------------------------------------
//#pragma package(smart_init)

TExcelApp1::TExcelApp1(boolean Visible)
        {
        app = CreateOleObject("Excel.Application");
        app.OlePropertySet("Visible", Visible);
        books = app.OlePropertyGet("Workbooks");
        };

TExcelApp1::TExcelApp1(AnsiString filename, boolean Visible)
        {
        //TExcelApp1::TExcelApp1(Visible);
        app = CreateOleObject("Excel.Application");
        app.OlePropertySet("Visible", Visible);
        books = app.OlePropertyGet("Workbooks");
        books.Exec(Procedure("Open")<<filename);
        book = books.OlePropertyGet("item",1);
        };

TExcelApp1::NewBook(int NomSheets, String NameSheets[]) //Имена листов 3, { "One", "Two", "Three" };
        {
        app.OlePropertySet("SheetsInNewWorkbook", NomSheets);
        books.Exec(Procedure("Add"));
        book = books.OlePropertyGet("item",1);
        for (int i = 0; i < NomSheets; i++)
                { sheet= book.OlePropertyGet("WorkSheets",i+1);
                  sheet.OlePropertySet("Name", WideString(NameSheets[i]));
                }

        };