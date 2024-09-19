#ifndef _Unit_LibExcel_H
#define _Unit_LibExcel_H
#endif // _Unit_LibExcel_H

//#pragma hdrstop
#include "Excel_2K_SRVR.h"
//#pragma package(smart_init)
//#include "Unit_LibExcel.cpp"

class TExcelApp1 {
        public:
        Variant app ;
        Variant books ;
        Variant book ;
        Variant sheet ;
        Variant range;
        TExcelApp1 (boolean Visible);
        TExcelApp1 (AnsiString filename, boolean Visible);
        void Show() {app.OlePropertySet("Visible", true);};
        NewBook (int NomSheets, String NameSheets[]); //Имена листов 3, { "One", "Two", "Three" };
        OpenSheet (WideString sheetname) {sheet = book.OlePropertyGet("WorkSheets", sheetname); sheet.OleProcedure("Activate");};
        AnsiString CellGet (int i, int j) {Variant cell = sheet.OlePropertyGet("Cells", i, j); return cell.OlePropertyGet("Value");};
        void CellSet (int i, int j, WideString textcell) {Variant cell = sheet.OlePropertyGet("Cells", i, j);
                                                          cell.OlePropertySet("NumberFormat","@"); cell.OlePropertySet("Value",textcell);};
        void Quit() {app.OleProcedure("Quit");};
        void Zagolovok() {app.OlePropertyGet("ActiveWindow").OlePropertySet("SplitRow", 1);
                 app.OlePropertyGet("ActiveWindow").OlePropertySet("FreezePanes", true);
                 sheet.OlePropertyGet("PageSetup").OlePropertySet("PrintTitleRows", "$1:$1");
                 range=sheet.OlePropertyGet("Range","A:K");
                 range.OlePropertyGet("Borders").OlePropertyGet("Item", 12).OlePropertySet("LineStyle", 1); //xlInsideHorizontal 12
                 range.OlePropertyGet("Borders").OlePropertyGet("Item", 11).OlePropertySet("LineStyle", 1); //xlInsideVertical 11
                 //range.OlePropertyGet("Borders").OlePropertyGet("Item", 5).OlePropertySet("LineStyle", 1); //xlDiagonalDown 5
                 //range.OlePropertyGet("Borders").OlePropertyGet("Item", 6).OlePropertySet("LineStyle", 1); //xlDiagonalUp 6
                 range.OlePropertyGet("Borders").OlePropertyGet("Item", 9).OlePropertySet("LineStyle", 1); //xlEdgeBottom 9
                 range.OlePropertyGet("Borders").OlePropertyGet("Item", 7).OlePropertySet("LineStyle", 1); //xlEdgeLeft 7
                 range.OlePropertyGet("Borders").OlePropertyGet("Item", 10).OlePropertySet("LineStyle", 1); //xlEdgeRight 10
                 range.OlePropertyGet("Borders").OlePropertyGet("Item", 8).OlePropertySet("LineStyle", 1); //xlEdgeTop 8
                 range=sheet.OlePropertyGet("Range","A:A");
                 //vVarCell.OlePropertySet("RowHeight", 20);
                 range.OlePropertySet("ColumnWidth",10);
                 range=sheet.OlePropertyGet("Range","B:K");
                 range.OlePropertySet("ColumnWidth",6);
                 };

        };


 