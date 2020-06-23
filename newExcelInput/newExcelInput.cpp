#include "deflate.h"
#include "Header.h"
#include "timeget.h"
#include "slidewindow.h"
#include "CRC.h"
#include "encode.h"
#include "decord.h"
#include "compressdata.h"
#include "SearchItemNUm.h"
#include "ShipDataRead.h"
#include "ZipFilewrite.h"
#include "shareRandW.h"
#include "RowColumn.h"

#include <codecvt>

#include "atlbase.h"
#include "atlstr.h"
#include "comutil.h"

#include "StyleWrite.h"

#include <locale.h>

//#include <crtdbg.h>//メモリリーク用
#define BUFSIZE  256

char* SJIStoUTF8(char* szSJIS, char* bufUTF8, int size) {
    wchar_t bufUnicode[BUFSIZE];
    int lenUnicode = MultiByteToWideChar(CP_ACP, 0, szSJIS, strlen(szSJIS) + 1, bufUnicode, BUFSIZE);
    WideCharToMultiByte(CP_UTF8, 0, bufUnicode, lenUnicode, bufUTF8, size, NULL, NULL);
    return bufUTF8;
}

char* convchar(wchar_t* a) {

    size_t origsize = wcslen(a) + 1;
    size_t convertedChars = 0;

    const size_t newsize = origsize * 2;
    // The new string will contain a converted copy of the original
    // string plus the type of string appended to it.
    char* nstring = new char[newsize];

    // Put a copy of the converted string into nstring
    wcstombs_s(&convertedChars, nstring, newsize, a, _TRUNCATE);
    // append the type of string to the new string.
    //_mbscat_s((unsigned char*)nstring, newsize + strConcatsize, (unsigned char*)strConcat);

    return nstring;
}

int main(char* fname[], int i) {

    // メモリリーク検出
    //_CrtSetDbgFlag(_CRTDBG_ALLOC_MEM_DF | _CRTDBG_LEAK_CHECK_DF);
    char Zfn[] = "C:/Users/Bee/Desktop/発注てんぷれ.xlsx";
    char mfile[] = "C:/Users/Bee/Desktop/test[システム]test.xlsx"; //テスト書き出し
    char recd[] = "C:/Users/Bee/Desktop/Centraldata"; //テスト書き出し
    
    char shares[] = "sharedStrings.xml";
    char stylefn[] = "xl/styles.xml";

    /*-----------------------
    shareシート読み込み
    -----------------------*/
    CenterDerect* cddata = nullptr;
    std::ifstream Zr(Zfn, std::ios::in | std::ios::binary);

    HeaderRead* hr2 = new HeaderRead(Zfn);
    hr2->endread(&Zr);//終端コードの読み込み
    
    DeflateDecode *decShare = new DeflateDecode(&Zr);

    while (hr2->filenum < hr2->ER->centralsum) {
        cddata = hr2->centeroneread(hr2->readpos, hr2->ER->size, hr2->ER->centralnum, shares, &Zr);
        if (cddata)
            break;
    }
    if (cddata) {//ファイル名が合えばローカルヘッダー読み込み
        hr2->localread(cddata->localheader, &Zr);//sharesstringsの読み込み
        decShare->dataread(hr2->LH->pos, cddata->nonsize);
    }

    shareRandD *sharray = new shareRandD(decShare->ReadV, decShare->readlen);//share 再初期化 検索用に呼び出し
    sharray->getSicount();//get si count
    sharray->ReadShare();//文字列読み込み
 
    delete hr2;
    delete decShare;

     /*-----------------------
    スタイルシート読み込み
    -----------------------*/
    hr2=new HeaderRead(Zfn);
    hr2->endread(&Zr);//終端コードの読み込み

    DeflateDecode* Sdeco=new DeflateDecode(&Zr);

    while (hr2->filenum < hr2->ER->centralsum) {
        cddata = hr2->centeroneread(hr2->readpos, hr2->ER->size, hr2->ER->centralnum, stylefn, &Zr);
        if (cddata)
            break;
    }
    if (cddata) {//ファイル名が合えばローカルヘッダー読み込み
        hr2->localread(cddata->localheader, &Zr);//sharesstringsの読み込み
        Sdeco->dataread(hr2->LH->pos, cddata->nonsize);
    }

    FILE* f=nullptr;

    StyleWrite* sr = new StyleWrite(f);
    
    sr->readstyle(Sdeco->ReadV, Sdeco->readlen);

    delete Sdeco;
  
     /*-----------------------
    発注到着シート読み込み
    -----------------------*/
 
    DeflateDecode* Hdeco;
    char sheet[] = "worksheets/sheet1";

    bool t = false;
    Ctags* mh;//発注到着　cell データ読み込み
    CDdataes* slideCDdata = hr2->saveCD;//ファイル名検索用

    hr2->readpos = hr2->ER->position;//読み込み位置初期化
    hr2->filenum = 0;//レコード数初期化
    int result = 0;


    while (hr2->filenum < hr2->ER->centralsum) {
        //ファイル名 sheet 部分一致検索
        cddata = hr2->centeroneread(hr2->readpos, hr2->ER->size, hr2->ER->centralnum, sheet, &Zr);
        if (cddata) {
            std::cout << "sheet一致ファイルネーム：" << cddata->filename << std::endl;
            hr2->localread(cddata->localheader, &Zr);//"worksheets/sheet"に一致するファイルの中身検索

            Hdeco = new DeflateDecode(&Zr);//解凍
            Hdeco->dataread(hr2->LH->pos, cddata->nonsize);//解凍　データ読み込み

            mh = new Ctags(Hdeco->ReadV, Hdeco->readlen, sharray);//シートデータ読み込み
            mh->sheetread();

            //style 検索
            //177 検索
            int stylec = 0;

            //177 bee
            //24 shoplist
            //26 zozo
            //28 smarby
            //30 magaseek
            while (stylec < 24 && sr->cellXfsRoot) {
                sr->cellXfsRoot = sr->cellXfsRoot->next;
                stylec++;
            }
            ArrayNumber* changeStr = new ArrayNumber;
            std::cout << "style 24 " << std::endl;

            if (sr->cellXfsRoot->numFmtId)
                std::cout << "numFmtId : " << sr->cellXfsRoot->numFmtId << std::endl;
            if (sr->cellXfsRoot->fontId) {
                std::cout << "fontId : " << sr->cellXfsRoot->fontId << std::endl;
            }
            if (sr->cellXfsRoot->fillId)
                std::cout << "fillId : " << sr->cellXfsRoot->fillId << std::endl;
            if (sr->cellXfsRoot->borderId)
                std::cout << "borderId : " << sr->cellXfsRoot->borderId << std::endl;
            if (sr->cellXfsRoot->xfId)
                std::cout << "xfId : " << sr->cellXfsRoot->xfId << std::endl;
            if (sr->cellXfsRoot->applyNumberFormat)
                std::cout << "applyNumberFormat : " << sr->cellXfsRoot->applyNumberFormat << std::endl;
            if (sr->cellXfsRoot->applyFont)
                std::cout << "applyFont : " << sr->cellXfsRoot->applyFont << std::endl;
            if (sr->cellXfsRoot->applyFill)
                std::cout << "applyFill : " << sr->cellXfsRoot->applyFill << std::endl;
            if (sr->cellXfsRoot->Avertical)
                std::cout << "vertical : " << sr->cellXfsRoot->Avertical << std::endl;
            if (sr->cellXfsRoot->horizontal)
                std::cout << "horizontal : " << sr->cellXfsRoot->horizontal << std::endl;
            if (sr->cellXfsRoot->AwrapText)
                std::cout << "wrapText : " << sr->cellXfsRoot->AwrapText << std::endl;

            int place = 0;

            //fontid 検索
            while (sr->cellXfsRoot->fontId[place] != '\0')
                place++;
            UINT32 fontnum = changeStr->RowArraytoNum(sr->cellXfsRoot->fontId, place);

            stylec = 0;
            while (stylec < fontnum) {
                sr->fontRoot = sr->fontRoot->next;
                stylec++;
            }
            std::cout << std::endl;
            std::cout << "フォント検索 " << fontnum << std::endl;
            std::cout << std::endl;
            if (sr->fontRoot->sz)
                std::cout << "sz : " << sr->fontRoot->sz << std::endl;
            if (sr->fontRoot->color)
                std::cout << "theme : " << sr->fontRoot->color << std::endl;
            if (sr->fontRoot->rgb)
                std::cout << "rgb : " << sr->fontRoot->rgb << std::endl;
            if (sr->fontRoot->name)
                std::cout << "name : " << sr->fontRoot->name << std::endl;
            if (sr->fontRoot->family)
                std::cout << "family : " << sr->fontRoot->family << std::endl;
            if (sr->fontRoot->charset)
                std::cout << "charset : " << sr->fontRoot->charset << std::endl;
            if (sr->fontRoot->scheme)
                std::cout << "scheme : " << sr->fontRoot->scheme << std::endl;

            //fillid 検索
            place = 0;
            while (sr->cellXfsRoot->fillId[place] != '\0')
                place++;
            fontnum = changeStr->RowArraytoNum(sr->cellXfsRoot->fillId, place);
            stylec = 0;
            while (stylec < fontnum) {
                sr->fillroot = sr->fillroot->next;
                stylec++;
            }
            std::cout << std::endl;
            std::cout << "fill検索 " << fontnum << std::endl;
            std::cout << std::endl;
            if (sr->fillroot->patten)
                std::cout << "patternType : " << sr->fillroot->patten->patternType << std::endl;
            if (sr->fillroot->fg){
                if (sr->fillroot->fg->rgb)
                    std::cout << "fgColor rgb : " << sr->fillroot->fg->rgb << std::endl;
                if (sr->fillroot->fg->theme)
                    std::cout << "fgColor theme : " << sr->fillroot->fg->theme << std::endl;
        }

            //ボーダー 検索
            place = 0;
            while (sr->cellXfsRoot->borderId[place] != '\0')
                place++;
            fontnum = changeStr->RowArraytoNum(sr->cellXfsRoot->borderId, place);
            stylec = 0;
            while (stylec < fontnum) {
                sr->fillroot = sr->fillroot->next;
                stylec++;
            }
            std::cout << std::endl;
            std::cout << "border検索 " << fontnum << std::endl;
            std::cout << std::endl;
            if (sr->BorderRoot->left)
                std::cout << "left style : " << sr->BorderRoot->left->style << std::endl;
            else
                std::cout << "left 何もなし " << std::endl;
            if (sr->BorderRoot->right)
                std::cout << "right style : " << sr->BorderRoot->right->style << std::endl;
            else
                std::cout << "right 何もなし " << std::endl;
            if (sr->BorderRoot->bottom)
                std::cout << "bottom style : " << sr->BorderRoot->top->style << std::endl;
            else
                std::cout << "bottom 何もなし " << std::endl;
            if (sr->BorderRoot->top)
                std::cout << "top style : " << sr->BorderRoot->top->style << std::endl;
            else
                std::cout << "top 何もなし " << std::endl;

            //xfId 検索
            place = 0;
            while (sr->cellXfsRoot->xfId[place] != '\0')
                place++;
            fontnum = changeStr->RowArraytoNum(sr->cellXfsRoot->xfId, place);
            stylec = 0;
            while (stylec < fontnum) {
                sr->cellstyleXfsRoot = sr->cellstyleXfsRoot->next;
                stylec++;
            }
            std::cout << std::endl;
            std::cout << "xfId検索 " << fontnum << std::endl;
            std::cout << std::endl;
            if (sr->cellstyleXfsRoot->numFmtId)
                std::cout << "numFmtId : " << sr->cellstyleXfsRoot->numFmtId << std::endl;
            if (sr->cellstyleXfsRoot->fontId)
                std::cout << "fontId : " << sr->cellstyleXfsRoot->fontId << std::endl;
            if (sr->cellstyleXfsRoot->fillId)
                std::cout << "fillId : " << sr->cellstyleXfsRoot->fillId << std::endl;
            if (sr->cellstyleXfsRoot->borderId)
                std::cout << "borderId : " << sr->cellstyleXfsRoot->borderId << std::endl;
            if (sr->cellstyleXfsRoot->applyFont)
                std::cout << "applyFont : " << sr->cellstyleXfsRoot->applyFont << std::endl;
            if (sr->cellstyleXfsRoot->applyBorder)
                std::cout << "applyBorder : " << sr->cellstyleXfsRoot->applyBorder << std::endl;
            if (sr->cellstyleXfsRoot->applyAlignment)
                std::cout << "applyAlignment : " << sr->cellstyleXfsRoot->applyAlignment << std::endl;
            if (sr->cellstyleXfsRoot->applyProtection)
                std::cout << "applyProtection : " << sr->cellstyleXfsRoot->applyProtection << std::endl;
            if (sr->cellstyleXfsRoot->Avertical)
                std::cout << "vertical : " << sr->cellstyleXfsRoot->Avertical << std::endl;

            //numfmid 検索
            place = 0;
            stylec = 0;
            int res = 0;
            while (sr->numFmtsRoot && sr->cellstyleXfsRoot->numFmtId[0]!='0') {

                res = strcmp((const char*)sr->numFmtsRoot->Id, (const char*)(sr->cellstyleXfsRoot->numFmtId));
                if (res == 0)
                    break;
                sr->numFmtsRoot = sr->numFmtsRoot->next;
            }
            std::cout << std::endl;
            std::cout << "numFmtId検索 " << sr->cellstyleXfsRoot->numFmtId << std::endl;
            std::cout << std::endl;
            if (sr->numFmtsRoot->Id)
                std::cout << "numFmtId : " << sr->numFmtsRoot->Id << std::endl;
            if (sr->numFmtsRoot->Code)
                std::cout << "code : " << sr->numFmtsRoot->Code << std::endl;
            
            delete Hdeco;//デコードデータ　削除
            delete mh;
        }
    }
    std::cout << "end item search" << std::endl;
    
    delete sharray;

    delete hr2;
         
    Zr.close();
    
    return 0;
}