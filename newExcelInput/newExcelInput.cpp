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

#include "excel_style.h"

#include <locale.h>

//#include <crtdbg.h>//メモリリーク用
#define BUFSIZE  256

std::wstring multi_to_wide_capi(std::string const& src)
{
    std::size_t converted{};
    std::vector<wchar_t> dest(src.size(), L'\0');
    if (::_mbstowcs_s_l(&converted, dest.data(), dest.size(), src.data(), _TRUNCATE, ::_create_locale(LC_ALL, "jpn")) != 0) {
        throw std::system_error{ errno, std::system_category() };
    }
    dest.resize(std::char_traits<wchar_t>::length(dest.data()));
    dest.shrink_to_fit();
    return std::wstring(dest.begin(), dest.end());
}

std::string wide_to_utf8_cppapi(std::wstring const& src)
{
    std::wstring_convert<std::codecvt_utf8_utf16<wchar_t>> converter;
    return converter.to_bytes(src);
}

std::string multi_to_utf8_cppapi(std::string const& src)
{
    auto const wide = multi_to_wide_capi(src);
    return wide_to_utf8_cppapi(wide);
}

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
    char* nstring = new char[newsize];

    wcstombs_s(&convertedChars, nstring, newsize, a, _TRUNCATE);

    return nstring;
}

int main(char* fname[], int i) {

    // メモリリーク検出
    //_CrtSetDbgFlag(_CRTDBG_ALLOC_MEM_DF | _CRTDBG_LEAK_CHECK_DF);

    char PLfn[] = "C:/Users/Bee/Desktop/LY200623-A ゾゾ.xlsx"; //読み込むファイルの指定 /LY20191127-SHIP BEE だけ.xlsx
    char Zfn[] = "C:/Users/Bee/Desktop/【1 縦売り 夏】在庫状況.xlsx";
    char mfile[] = "C:/Users/Bee/Desktop/test[システム]test.xlsx"; //テスト書き出し
    char recd[] = "C:/Users/Bee/Desktop/Centraldata"; //テスト書き出し
    
    char inputStr[] = "2020,ship,5/19着";//shift-jis

   //ワイド文字列(WCHAR*)をマルチバイト文字列(char*)に変換
//変換前文字列
    WCHAR wStrW[] = _T("2020,ship,5/19着");
    //変換文字列格納バッファ

    wchar_t intxto[] = _T("2020,shiptyt着");

    std::string conv = wide_to_utf8_cppapi(intxto);//utf8へ
    const char* newinput = conv.c_str();

    char* Mnstring = convchar(intxto);

    //変換文字列格納バッファ
   //char	inMainstr[50];
    char	wStrC[250];
    size_t wLen = 0;
    errno_t err = 0;

    //ロケール指定
    setlocale(LC_ALL, "japanese");

    //変換
    err = wcstombs_s(&wLen, wStrC, 250, intxto, _TRUNCATE);

    char* sMstr = (char*)malloc(100);
    char* inMainstr = (char*)malloc(100);

    sMstr = SJIStoUTF8(wStrC, sMstr, 100);
    
    int wlen = 0;
    while (sMstr[wlen] != '\0') {
        inMainstr[wlen] = sMstr[wlen];
        wlen++;
    }
    inMainstr[wlen] = '\0';
    
    std::cout << "wchar to char : " << inMainstr <<" "<<wLen<< std::endl;

    /*-----------------------------
    入力文字チェック
    -------------------------------*/

    /*-----------------------
     シェアー文字列読み込み
     -----------------------*/
        
    char shares[] = "sharedStrings.xml";//ファイル位置　ファイル名部分一致検索　sharedStrings.xml

    std::ifstream PLR(PLfn, std::ios::in | std::ios::binary);

    HeaderRead* hr = new HeaderRead(PLfn);
    CenterDerect* cddata = nullptr;//セントラルディレクトのデータ
    hr->endread(&PLR);//終端コードの読み込み

    DeflateDecode* decShare = new DeflateDecode(&PLR);//sharestring ファイルの保存用
    //share セントラル取得
    while (hr->filenum < hr->ER->centralsum) {//cddata mallocなし
        cddata = hr->centeroneread(hr->readpos, hr->ER->size, hr->ER->centralsum, shares, &PLR);
        if (cddata) {
            break;
        }
    }
    if (cddata) {//ファイル名が合えばローカルヘッダー読み込み
        hr->localread(cddata->localheader,&PLR);//sharesstringsの読み込み
        decShare->dataread(hr->LH->pos, cddata->nonsize);
    }

    shareRandD* sharray=new shareRandD(decShare->ReadV, decShare->readlen);//share string read to array

    sharray->getSicount();//get si count

    sharray->ReadShare();//文字列読み込み デコードデータ削除OK

    delete decShare;

     /*-----------------------
    パッキングリストシート読み込み
    -----------------------*/

    char sheetname[] = "worksheets/sheet1.xml";//ファイル位置　ファイル名部分一致検索
    DeflateDecode* decsheet = new DeflateDecode(&PLR);

    hr->filenum = 0;//レコード数初期化
    hr->readpos = hr->ER->position;
    cddata = nullptr;
    while (hr->filenum < hr->ER->centralsum) {
        cddata = hr->centeroneread(hr->readpos, hr->ER->size, hr->ER->centralsum, sheetname, &PLR);//セントラルディレクトのデータ
        if (cddata) {
            break;
        }
    }
    if (cddata) {//ファイル名が合えばローカルヘッダー読み込み
        hr->localread(cddata->localheader, &PLR);//sharesstringsの読み込み
        decsheet->dataread(hr->LH->pos, cddata->nonsize);
    }

    Ctags* ms=new Ctags(decsheet->ReadV, decsheet->readlen, sharray);

    ms->sheetread();

    delete decsheet;//デコードデータ削除

    shipinfo* sg=new shipinfo(ms->rows);

    sg->GetItems();//mallocなし　シートとセット

    PLR.close();

    /*-----------------------
    shareシート読み込み
    -----------------------*/
    
    std::ifstream Zr(Zfn, std::ios::in | std::ios::binary);

    HeaderRead* hr2 = new HeaderRead(Zfn);
    hr2->endread(&Zr);//終端コードの読み込み
    
    decShare = new DeflateDecode(&Zr);

    while (hr2->filenum < hr2->ER->centralsum) {
        cddata = hr2->centeroneread(hr2->readpos, hr2->ER->size, hr2->ER->centralnum, shares, &Zr);
        if (cddata)
            break;
    }
    if (cddata) {//ファイル名が合えばローカルヘッダー読み込み
        hr2->localread(cddata->localheader, &Zr);//sharesstringsの読み込み
        decShare->dataread(hr2->LH->pos, cddata->nonsize);
    }

    //入力テキスト文字分ける
    char** splstr = nullptr;
    inputtxt* inptxt = nullptr;
    if(strlen(Mnstring)>0)
        inptxt = hr2->slipInputText((char*)Mnstring, inptxt);//分割テキスト
    else
    {
        inptxt = (inputtxt*)malloc(sizeof(inputtxt));
        inptxt = nullptr;
    }
    inputtxt* rootintxt = inptxt;

    int txtnum = hr2->intxtCount;//テキスト数

    delete sharray;
    delete hr;
    
    sharray = new shareRandD(decShare->ReadV, decShare->readlen);//share 再初期化 検索用に呼び出し
    sharray->getSicount();//get si count
    sharray->ReadShare();//文字列読み込み

    //delete decShare;
    //share 検索　あったら番号取得

    int ma = 0;
    //UINT8** instrFlag = (UINT8**)malloc(txtnum);

    rootintxt = sharray->searchSi(rootintxt);//マッチした文字列のSi番号取得　なければnullptr

    
    //シェアー書き込み
    UINT8* sharedata = nullptr;
    
    std::cout << "day : " << sg->day << std::endl;
    sharedata= sharray->writeshare(sg->day, sg->daylen, inptxt, txtnum);//share文字列書き込み share data更新
    UINT64 shrelength = sharray->writeleng;
    

    /*--------------------------
    share data書き込み　圧縮
    ---------------------------*/

    //ファイル書き出し用
    _Post_ _Notnull_ FILE* cdf;
    fopen_s(&cdf, recd, "wb");
    if (!cdf)
        printf("ファイルを開けませんでした");

    _Post_ _Notnull_ FILE* fpr;
    fopen_s(&fpr, mfile, "wb");
    if (!fpr)
        printf("ファイルを開けませんでした");

    LHmake dd;
    dd.getday();//時間の取得
    zipwrite zip(dd.times);//zipデータを作る
    HeaderWrite hw;//ローカルヘッダの書き込み
    UINT32 localSize = 0;
    UINT32 CDp = 0;

    encoding* shenc = new encoding;//sharestring 圧縮
    shenc->compress(sharedata, shrelength);
    free(sharedata);

    if (cddata) {
        localSize = hw.localwrite(fpr, dd.times, sharray->buckcrc, shenc->datalen, sharray->writeleng, cddata->filenameleng, cddata->filename, cddata->version, cddata->bitflag, cddata->method);
        CDp = (shenc->datalen) + localSize;//ローカルデータサイズ

        //セントラル情報の書き換え//
        cddata->crc = sharray->Crc.crc32;
        cddata->localheader = zip.writeposition;
        cddata->day = (dd.times) & 0xFFFF;
        cddata->time = (dd.times >> 16) & 0xFFFF;
        cddata->size = shenc->datalen;
        cddata->nonsize = sharray->writeleng;//内容変更したら　更新必要

        hw.centralwrite(cdf, *cddata);
    }

    zip.writeposition += CDp;//データ書き込み位置更新

    shenc->write(fpr);//圧縮データの書き込み
    
    delete decShare;//シェアー解凍データ削除
    delete shenc;//share圧縮データ削除

     /*-----------------------
    スタイルシート読み込み
    -----------------------*/

    hr2=new HeaderRead(Zfn);
    hr2->endread(&Zr);//終端コードの読み込み

    DeflateDecode* Sdeco=new DeflateDecode(&Zr);

    char stylefn[] = "xl/styles.xml";

    while (hr2->filenum < hr2->ER->centralsum) {
        cddata = hr2->centeroneread(hr2->readpos, hr2->ER->size, hr2->ER->centralnum, stylefn, &Zr);
        if (cddata)
            break;
    }
    if (cddata) {//ファイル名が合えばローカルヘッダー読み込み
        hr2->localread(cddata->localheader, &Zr);//sharesstringsの読み込み
        Sdeco->dataread(hr2->LH->pos, cddata->nonsize);
    }

    styleread* sr = new styleread();
    
    sr->readstyle(Sdeco->ReadV, Sdeco->readlen);

    delete Sdeco;
    delete sr;
  
     /*-----------------------
    発注到着シート読み込み
    -----------------------*/
 
    DeflateDecode* Hdeco;
    char sheet[] = "worksheets/sheet";
    const char sharefn[] = "xl/sharedStrings.xml";

    bool t = false;
    Ctags* mh;//発注到着　cell データ読み込み
    searchItemNum* sI=nullptr;//品番検索　＆　書き込み
    CDdataes* slideCDdata = hr2->saveCD;//ファイル名検索用

    hr2->readpos = hr2->ER->position;//読み込み位置初期化
    hr2->filenum = 0;//レコード数初期化
    int result = 0;

    //品番、カラーエラー用
    MatchColrs* matchs = (MatchColrs*)malloc(sizeof(MatchColrs));
    matchs = nullptr;
    MatchColrs* matchsroot = nullptr;


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

            sI = new searchItemNum(sg->its, mh);

            t = sI->searchitemNumber(sharray->uniqstr, inptxt,txtnum);//品番検索　＆　セルデータ追加　シェアー消去（入れる場合は引数に）

            if (t)
            {//一致品番あった場合
                if (sI->rootMat) {//一致アイテムの保存
                    while (sI->rootMat) {
                        matchs = sI->addmatches(matchs, sI->rootMat->itemnum, sI->rootMat->color);
                        sI->rootMat = sI->rootMat->next;
                    }
                }

                mh->writesheetdata();//シートデータ書き込み
                crc CRC;
                CRC.rescrc();
                CRC.mcrc(mh->wd, mh->p);//crc 計算

                encoding* enc = new encoding;//圧縮
                enc->compress(mh->wd, mh->p);//データ圧縮

                cddata->zokusei = 0x00;
                cddata->gaibuzokusei = 0x00;
                cddata->bitflag = 0x06;
                localSize = hw.localwrite(fpr, dd.times, CRC.crc32, enc->datalen & 0xFFFFFFFF, mh->p & 0xFFFFFFFF, cddata->filenameleng, cddata->filename, cddata->version, cddata->bitflag, cddata->method);//ローカルヘッダー書き込み
                CDp = (enc->datalen) + localSize;//ローカルデータサイズ
                cddata->crc = CRC.crc32;//セントラル情報の書き換え
                cddata->localheader = zip.writeposition;
                cddata->time = (dd.times) & 0xFFFF;
                cddata->day = (dd.times >> 16) & 0xFFFF;
                cddata->size = enc->datalen;
                cddata->nonsize = mh->p & 0xFFFFFFFF;//内容変更したら　更新必要
                //cddata->bitflag = 0x07;//0B00000111 超高速 高速 cddata->bitflag
                zip.writeposition += CDp;//データ書き込み位置更新

                enc->write(fpr);//圧縮データの書き込み

                //cddata一旦書き込み
                hw.centralwrite(cdf, *cddata);

                delete enc;
            }
            else {
                //cddata一旦書き込み
                UINT32 LHposstock = zip.writeposition;//ローカルヘッダーの位置更新用
                zip.LoclheadAndDatacopy(cddata->localheader, fpr, Zfn);//ローカルヘッダー検索＆書き込み
                cddata->localheader = LHposstock;//ローカルヘッダー相対位置のみ変更
                hw.centralwrite(cdf, *cddata);
            }
            delete sI;
            delete Hdeco;//デコードデータ　削除
            delete mh;
        }
        else {
            result = strcmp(hr2->scd->filename, sharefn);
            if (result != 0) {
                //cddata一旦書き込み
                UINT32 LHposstock = zip.writeposition;//ローカルヘッダーの位置更新用
                zip.LoclheadAndDatacopy(hr2->scd->localheader, fpr, Zfn);//ローカルヘッダー検索＆書き込み
                hr2->scd->localheader = LHposstock;//ローカルヘッダー相対位置のみ変更
                hw.centralwrite(cdf, *hr2->scd);
            }
        }
    }
    std::cout << "end item search" << std::endl;
    hr2->freetxt(inptxt);//入力文字列の削除
    
    delete(sharray);

    Items* errorItem = (Items*)malloc(sizeof(Items));
    errorItem = nullptr;

    sI = new searchItemNum(nullptr, nullptr);

    matchsroot = matchs;
    unitC un;
    int matchor = 1;
    int matchcol = 1;

    //シートがない品番、カラーを比較
    if (matchs) {
        while (sg->its) {
            sg->its->col = un.slipNum(sg->its->col);
            while (matchs) {
                matchor = strcmp((const char*)sg->its->itn, (const char*)matchs->itemnum);
                if (matchor == 0) {//品番、いろ一致入力
                    matchcol = strcmp((const char*)sg->its->col, (const char*)matchs->color);
                    if (matchcol == 0) {
                        //std::cout << "保存　一致データあり" << std::endl;
                        break;
                    }
                }
                else {
                    matchcol = 1;
                }
                matchs = matchs->next;
            }
            if (matchor == 0 && matchcol == 0) {//品番、いろ一致入力
            }
            else {
                //エラー　シートなし
                errorItem = sI->addItems(errorItem, sg->its);
                char* shiftj = sI->CharChenge(sg->its->col);
                std::cout << "シートなし" << sg->its->itn << std::endl;
            }
            matchs = matchsroot;//初期化
            sg->its = sg->its->next;
        }
    }
    else {
        std::cout << "sheet no error" << std::endl;
    }
    free(errorItem);
    free(matchsroot);

    delete sg;//アイテム　文字データ シートとセット削除
    delete ms;//シート　セル削除 PLシートデータ

    hw.eocdwrite(cdf, hr2->ER->discnum, hr2->ER->disccentral, hr2->ER->centralnum, hr2->ER->centralsum, zip.writeposition, hw.cdsize);

    if (cdf)
        fclose(cdf);

    std::ifstream fin(recd, std::ios::in | std::ios::binary);
    if (!fin) {
        std::cout << "not file open" << std::endl;
    }
    //std::cout << hw.cdsize << " zip " <<zip.writeposition << std::endl;
    char rd = 0;

    if (fpr) {
        while (!fin.eof()) {
            fin.read((char*)&rd, sizeof(char));
            if (!fin.eof())
                fwrite(&rd, sizeof(char), 1, fpr);
        }
    }

    
    delete hr2;

    remove(recd);

    if (fpr)
        fclose(fpr);
        

    Zr.close();
    
    return 0;
}