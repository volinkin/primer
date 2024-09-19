//#define NO_WIN32_LEAN_AND_MEAN //SHGetSpecialFolderPath
/*#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <ComCtrls.hpp>
#include <Forms.hpp> */

AnsiString GetSpecialFolderPath(int Path){char Buff[MAX_PATH]; SHGetSpecialFolderPath(0, Buff, Path, 0); return AnsiString(Buff);};
//перед ней в Unit.cpp проекта нужно в самом верху прописать #define NO_WIN32_LEAN_AND_MEAN
// CSIDL_PERSONAL , CSIDL_APPDATA

AnsiString FindSL (TStringList *SL, AnsiString tfind)
        {for (int j=0; j<SL->Count; j++)
                {int pos=SL->Strings[j].Pos(tfind);
                if (pos) return SL->Strings[j].SubString(pos+tfind.Length(),SL->Strings[j].Length()).TrimLeft();
                };
        };

AnsiString FindInFile (AnsiString file, AnsiString tfind)
        {AnsiString file1, res;
        file1=ExtractFilePath(Application->ExeName)+"\\"+file;
        TFileStream *tfile =new TFileStream(file1,fmOpenReadWrite);
        TStringList *SL = new TStringList;
        SL->LoadFromStream(tfile);
        res=FindSL(SL, tfind);
        delete SL; delete tfile;
        return res;
        };

void GetTextFromResourse (TStringList * sl, AnsiString name)
        {TResourceStream* ptRes = new TResourceStream((int)HInstance, name,"RT_RCDATA");
        //Можно и сохранить ptRes->SaveToFile("readme.txt");
        //Загружаем ресурс
        //TStringList *st = new TStringList;
        sl->LoadFromStream(ptRes);
        delete ptRes;
        };

/*AnsiString GetTextFromResourse (TStringList * sl, AnsiString name)
        {TResourceStream* ptRes =
        new TResourceStream((int)HInstance, name,"RT_RCDATA");
        //Можно и сохранить ptRes->SaveToFile("readme.txt");
        //Загружаем ресурс
        //TStringList *st = new TStringList;
        sl->LoadFromStream(ptRes);
        delete ptRes;

        }; */

AnsiString GetPathDirFromFileName (AnsiString FileName)
        {for (int pos=FileName.Length()-1; pos>=1; pos--) {if (FileName[pos]==92) {return FileName.SubString(1, pos-1); pos=1;};};
        };

AnsiString GetFileFromFileName (AnsiString FileName)
        {for (int pos=FileName.Length()-1; pos>=1; pos--) {if (FileName[pos]==92) {return FileName.SubString(pos+1, FileName.Length()); pos=1;};};
        };

AnsiString ZaprosYear(AnsiString Text, AnsiString ParYear, AnsiString Year) //заменяет в Text тег ParYear на Year - ZaprosYear(sl->Text, ":paramYear", IntToStr(this->Year));
        {
        int pos = Text.Pos(ParYear); int con=pos+ParYear.Length();
        if (pos) return Text.SubString(1, pos-1)+Year+Text.SubString(con,Text.Length());
        else return Text;
        };

AnsiString FormJuric(AnsiString name)
        {AnsiString res="";
         res=name;
         if (name.SubString(1,3)=="ООО") res="общества с ограниченной ответственностью"+name.SubString(4,name.Length());
         if (name.SubString(1,3)=="ПАО") res="публичного акционерного общества"+name.SubString(4,name.Length());
         if (name.SubString(1,3)=="АО ") res="акционерного общества "+name.SubString(4,name.Length());
         if (name.SubString(1,3)=="ИП ") res="индивидуального предпринимателя "+name.SubString(4,name.Length());
         return res;
        };

class TRoamingFile {
        public:
        AnsiString Vers;
        AnsiString NameApp;
        TStringList *SL;
        enum TRoamingFileRows {vers, page, ProfileFile, ProfileFileLast, r2, r3, ProvFile,ProvZadanie, max} rows;
        //enum rows {vers=1, profile, max};
        //int max=0;
        TRoamingFile (AnsiString NameApp1);
        AnsiString Get(rows) {return SL->Strings[rows];};
        void Put(rows, AnsiString text) {SL->Strings[rows]=text; Save();};
        void Clear();
        void Save();
        void Show() {AnsiString mess;
                mess+=" Count="+IntToStr(SL->Count)+"; ";
                for (int j=0; j<SL->Count; j++){mess+=" s["+IntToStr(j)+"]="+SL->Strings[j]+";";};
                MessageBox(NULL, mess.c_str(), "TRoamingFile::Show", MB_OK);
                }
        };

TRoamingFile::TRoamingFile(AnsiString NameApp1)
        {Vers="12-02-2024";
        NameApp=NameApp1;
        SL = new TStringList;
        try
         {
         TFileStream *tfile =new TFileStream(GetSpecialFolderPath(CSIDL_APPDATA)+"\\"+NameApp+"\\option"+".ini",fmOpenReadWrite);
         SL->LoadFromStream(tfile);
         delete(tfile);
         if (!(SL->Strings[vers] == Vers &  max==SL->Count)) {
                //MessageBox(NULL, "vers, count", "TRoamingFile::Show", MB_OK);
                Clear();}
         }
        catch(Exception &ex){//MessageBox(NULL, "Нет файла", "TRoamingFile::Show", MB_OK);
                        Clear();
                        }
        };
void TRoamingFile::Clear() {SL->Clear(); for (int j=0; j<max; j++){SL->Add("");}; SL->Strings[vers]=Vers;};
void TRoamingFile::Save()
        {
        try
          {
          CreateDir(GetSpecialFolderPath(CSIDL_APPDATA)+"\\"+NameApp);
          TFileStream *tfile =new TFileStream(GetSpecialFolderPath(CSIDL_APPDATA)+"\\"+NameApp+"\\option"+".ini",fmCreate);
          SL->SaveToStream(tfile);
          delete(tfile);
          }
          catch(Exception &ex){};
        };
