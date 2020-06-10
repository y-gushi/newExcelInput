#include "RowColumn.h"

//セルデータの追加　row span更新
void Ctags::addcelldata(UINT8* row, UINT8* col, UINT8* t, UINT8* s, UINT8* v, F* f, UINT8* si) {
    Row* ro = nullptr;
    bool replace = false;
    int rp = 0;//桁数
    while (row[rp] != '\0')
        rp++;
    UINT32 rn = NA.RowArraytoNum(row, rp);//行配列から数字

    rp = 0;//桁数
    while (col[rp] != '\0')
        rp++;
    UINT32 cn = NA.ColumnArraytoNumber(col, rp);//spanよう　列文字数字

    if (cn > maxcol) {//diment 最大値更新
        maxcol = cn;
        dm->eC = NA.ColumnNumtoArray(maxcol, &rp);
        replace = true;
    }

    ro = searchRow(rows, rn);//row位置検索
    std::cout << " 追加 v : " << v << std::endl;
    ro->cells = addCtable(ro->cells, t, s, si, cn, v, f);//セル情報検索
    cn = NA.ColumnCharnumtoNumber(cn);//列番号連番へ　入れた列

    if (replace) {//最大値更新で　cols style確認
        cols* ncols = cls;
        while (ncols->next)//cols　最後参照
            ncols = ncols->next;
        int j = 0;
        while (ncols->max[j] != '\0')
            j++;
        UINT32 ncn = NA.RowArraytoNum(ncols->max, j);//最終max文字列 数字に変換
        if (cn > ncn) {//cols max 小さい
            UINT8 wi[] = "9"; UINT8 styl[] = "1";
            cls = addcolatyle(cls, col, col, wi, styl, nullptr, nullptr, nullptr);//cols 追加
        }
    }
    rp = 0;//桁数
    while (ro->spanE[rp] != '\0')
        rp++;
    UINT32 nnc = NA.RowArraytoNum(ro->spanE, rp);//span 文字列数字に
    if (cn > nnc) {//新しい文字列が大きい場合　更新
        UINT8* newspan = NA.InttoChar(cn, &rp);
        ro->spanE = newspan;
        //std::cout << " new span : " << ro->spanE << std::endl;
    }
}

void Ctags::writesheetdata() {
    size_t msize= size_t(dlen) + 6000;

    wd = (UINT8*)malloc(msize);//メモリサイズ変更　書き込み用
    //std::cout << "データ更新" << p << std::endl;
    p = 0;

    if (wd) {
        while (headXML[p] != '\0') {
            wd[p] = headXML[p];
            p++;
        }
    }
    
    writeDiment();
    writeSelection();
    writecols();
    writecells();
    writefinal();

    std::cout << "size : " << p << " dlen : " << dlen << std::endl;
}
//diment書き込み
void Ctags::writeDiment() {
    const char* dimstart = "<dimension ref=\"";
    std::cout << "ディメント更新" << p << std::endl;
    while (dimstart[writep] != '\0') {
        wd[p] = dimstart[writep]; p++; writep++;
    }writep = 0;
    while (dm->sC[writep] != '\0') {//スタート列
        wd[p] = dm->sC[writep]; p++; writep++;
    }writep = 0;
    while (dm->sR[writep] != '\0') {//スタート行
        wd[p] = dm->sR[writep]; p++; writep++;
    }writep = 0;
    wd[p] = ':'; p++;
    while (dm->eC[writep] != '\0') {//終了列
        wd[p] = dm->eC[writep]; p++; writep++;
    }writep = 0;
    while (dm->eR[writep] != '\0') {//終了行
        wd[p] = dm->eR[writep]; p++; writep++;
    }writep = 0;
    std::cout << "ディメント更新" << dm->eC << std::endl;
}
//cols書き込み
void Ctags::writecols() {
    const char* startcols = "<cols>";
    const char* colstr[] = { "<col min=\"","\" max=\"","\" width=\"","\" style=\""
        ,"\" bestFit=\"","\" hidden=\"","\" customWidth=\"" };

    std::cout << "コルス更新" << p << std::endl;
    const char* endcols = "</cols>";
    int writep = 0;
    while (startcols[writep] != '\0') {
        wd[p] = startcols[writep]; p++; writep++;
    }
    writep = 0;
    while (cls) {
        while (colstr[0][writep] != '\0') {
            wd[p] = colstr[0][writep]; p++; writep++;
        }writep = 0;
        while (cls->min[writep] != '\0') {//min+head
            wd[p] = cls->min[writep]; p++; writep++;
        }writep = 0;
        while (colstr[1][writep] != '\0') {
            wd[p] = colstr[1][writep]; p++; writep++;
        }writep = 0;
        while (cls->max[writep] != '\0') {//max
            wd[p] = cls->max[writep]; p++; writep++;
        }writep = 0;
        while (colstr[2][writep] != '\0') {
            wd[p] = colstr[2][writep]; p++; writep++;
        }writep = 0;
        while (cls->width[writep] != '\0') {//width
            wd[p] = cls->width[writep]; p++; writep++;
        }writep = 0;
        if (cls->style) {
            while (colstr[3][writep] != '\0') {
                wd[p] = colstr[3][writep]; p++; writep++;
            }writep = 0;
            while (cls->style[writep] != '\0') {//style
                wd[p] = cls->style[writep]; p++; writep++;
            }writep = 0;
        }
        if (cls->hidden) {
            while (colstr[5][writep] != '\0') {
                wd[p] = colstr[5][writep]; p++; writep++;
            }writep = 0;
            while (cls->hidden[writep] != '\0') {//hidden
                wd[p] = cls->hidden[writep]; p++; writep++;
            }writep = 0;
        }
        if (cls->bestffit) {
            while (colstr[4][writep] != '\0') {
                wd[p] = colstr[4][writep]; p++; writep++;
            }writep = 0;
            while (cls->bestffit[writep] != '\0') {//hidden
                wd[p] = cls->bestffit[writep]; p++; writep++;
            }writep = 0;
        }
        if (cls->customwidth) {
            while (colstr[6][writep] != '\0') {
                wd[p] = colstr[6][writep]; p++; writep++;
            }writep = 0;
            while (cls->customwidth[writep] != '\0') {//hidden
                wd[p] = cls->customwidth[writep]; p++; writep++;
            }writep = 0;
        }
        while (closetag[writep] != '\0') {
            wd[p] = closetag[writep]; p++; writep++;
        }writep = 0;
        cls = cls->next;
    }
    while (endcols[writep] != '\0') {
        wd[p] = endcols[writep]; p++; writep++;
    }writep = 0;
}
//selection pane書きこみ
void Ctags::writeSelection() {
    const char* selpane[] = { "<selection pane=\"","\" activeCell=\"","\" sqref=\"","</sheetView>","</sheetViews>" };
    const char* panes[] = { "<pane"," xSplit=\""," ySplit=\""," topLeftCell=\""," activePane=\""," state=\"" };
    std::cout << "セレクション更新" << p << std::endl;

    // <sheetview書き込み
    while (dimtopane[writep] != '\0') {
        wd[p] = dimtopane[writep]; p++; writep++;
    }writep = 0;

    // <pane 書き込み
    // <pane xSplit="176" ySplit="11" topLeftCell="HI75" activePane="bottomRight" state="frozen"/>
    while (Panes) {
        while (panes[0][writep] != '\0') {// <pane
            wd[p] = panes[0][writep]; p++; writep++;
        }writep = 0;

        if (Panes->xSp) {
            while (panes[1][writep] != '\0') {// <pane
                wd[p] = panes[1][writep]; p++; writep++;
            }writep = 0;
            while (Panes->xSp[writep] != '\0') {//pane
                wd[p] = Panes->xSp[writep]; p++; writep++;
            }writep = 0;
            wd[p] = '"'; p++;
        }
        if (Panes->ySp) {
            while (panes[2][writep] != '\0') {// <pane
                wd[p] = panes[2][writep]; p++; writep++;
            }writep = 0;
            while (Panes->ySp[writep] != '\0') {//pane
                wd[p] = Panes->ySp[writep]; p++; writep++;
            }writep = 0;
            wd[p] = '"'; p++;
        }
        if (Panes->tLeftC) {
            while (panes[3][writep] != '\0') {// <pane
                wd[p] = panes[3][writep]; p++; writep++;
            }writep = 0;
            while (Panes->tLeftC[writep] != '\0') {//pane
                wd[p] = Panes->tLeftC[writep]; p++; writep++;
            }writep = 0;
            wd[p] = '"'; p++;
        }
        if (Panes->activP) {
            while (panes[4][writep] != '\0') {// <pane
                wd[p] = panes[4][writep]; p++; writep++;
            }writep = 0;
            while (Panes->activP[writep] != '\0') {//pane
                wd[p] = Panes->activP[writep]; p++; writep++;
            }writep = 0;
            wd[p] = '"'; p++;
        }
        if (Panes->state) {
            while (panes[5][writep] != '\0') {// <pane
                wd[p] = panes[5][writep]; p++; writep++;
            }writep = 0;
            while (Panes->state[writep] != '\0') {//pane
                wd[p] = Panes->state[writep]; p++; writep++;
            }writep = 0;
            wd[p] = '"'; p++;
        }
        wd[p] = '/'; p++;
        wd[p] = '>'; p++;
        Panes = Panes->next;
    }

    // <selection 書き込み
    while (sct) {
        while (selpane[0][writep] != '\0') {
            wd[p] = selpane[0][writep]; p++; writep++;
        }writep = 0;
        if (sct->p) {
            while (sct->p[writep] != '\0') {//pane
                wd[p] = sct->p[writep]; p++; writep++;
            }writep = 0;
        }
        
        while (selpane[1][writep] != '\0') {
            wd[p] = selpane[1][writep]; p++; writep++;
        }writep = 0;
        if (sct->a) {
            while (sct->a[writep] != '\0') {//activcell
                wd[p] = sct->a[writep]; p++; writep++;
            }writep = 0;
        }
        
        while (selpane[2][writep] != '\0') {
            wd[p] = selpane[2][writep]; p++; writep++;
        }writep = 0;
        if (sct->s) {
            while (sct->s[writep] != '\0') {//sqref
                wd[p] = sct->s[writep]; p++; writep++;
            }writep = 0;
        }
        
        while (closetag[writep] != '\0') {
            wd[p] = closetag[writep]; p++; writep++;
        }writep = 0;
        sct = sct->next;
    }

    while (selpane[3][writep] != '\0') {
        wd[p] = selpane[3][writep]; p++; writep++;
    }writep = 0;
    while (selpane[4][writep] != '\0') {
        wd[p] = selpane[4][writep]; p++; writep++;
    }writep = 0;
    // <sheetFormatPr write
    while (sFPr[writep] != '\0') {
        wd[p] = sFPr[writep]; p++; writep++;
    }writep = 0;
}
//セルデータ書き込み
void Ctags::writecells() {
    const char* sheetstart = "<sheetData>";
    const char* shend = "</sheetData>";
    std::cout << "セル更新" << p << std::endl;

    int Place = 0;
    while (sheetstart[writep] != '\0') {
        wd[p] = sheetstart[writep]; p++; writep++;
    }writep = 0;

    const char* rstart = "<row r=\"";
    const char* rspa = "\" spans=\"";
    const char* rS = "\" s=\"";
    const char* rtaght = "\" ht=\"";
    const char* rcH = "\" customHeight=\"";
    const char* rcF = "\" customFormat=\"";
    const char* rtB = "\" thickBot=\"";
    const char* rend = "\">";
    const char* Rend = "</row>";

    while (rows) {
        while (rstart[writep] != '\0') {
            wd[p] = rstart[writep]; p++; writep++;
        }writep = 0;
        UINT8* RN = NA.InttoChar(rows->r, &Place);//数字　配列変換
        while (RN[writep] != '\0') {// r
            wd[p] = RN[writep]; p++; writep++;
        }writep = 0;
        if (rows->spanS) {//ない場合ある
            while (rspa[writep] != '\0') {//span
                wd[p] = rspa[writep]; p++; writep++;
            }writep = 0;

            while (rows->spanS[writep] != '\0') {// spanstart
                wd[p] = rows->spanS[writep]; p++; writep++;
            }writep = 0;
            wd[p] = ':'; p++;
        }
        if (rows->spanE) {
            while (rows->spanE[writep] != '\0') {// spanend
                wd[p] = rows->spanE[writep]; p++; writep++;
            }writep = 0;
        }
        if (rows->s) {
            while (rS[writep] != '\0') {
                wd[p] = rS[writep]; p++; writep++;
            }writep = 0;
            while (rows->s[writep] != '\0') {// s
                wd[p] = rows->s[writep]; p++; writep++;
            }writep = 0;
        }
        if (rows->customFormat) {
            while (rcF[writep] != '\0') {// customformat
                wd[p] = rcF[writep]; p++; writep++;
            }writep = 0;
            while (rows->customFormat[writep] != '\0') {//r
                wd[p] = rows->customFormat[writep]; p++; writep++;
            }writep = 0;
        }
        if (rows->ht) {
            while (rtaght[writep] != '\0') {// ht
                wd[p] = rtaght[writep]; p++; writep++;
            }writep = 0;
            while (rows->ht[writep] != '\0') {//r
                wd[p] = rows->ht[writep]; p++; writep++;
            }writep = 0;
        }
        if (rows->customHeight) {
            while (rcH[writep] != '\0') {//custumhigh
                wd[p] = rcH[writep]; p++; writep++;
            }writep = 0;
            while (rows->customHeight[writep] != '\0') {//r
                wd[p] = rows->customHeight[writep]; p++; writep++;
            }writep = 0;
        }
        if (rows->thickBot) {
            while (rtB[writep] != '\0') {// thickbot
                wd[p] = rtB[writep]; p++; writep++;
            }writep = 0;
            while (rows->thickBot[writep] != '\0') {//r
                wd[p] = rows->thickBot[writep]; p++; writep++;
            }writep = 0;
        }
        while (rend[writep] != '\0') {// ">
            wd[p] = rend[writep]; p++; writep++;
        }writep = 0;

        UINT8* rne = NA.InttoChar(rows->r, &Place);//配列に変更
        writec(rows->cells, rne);

        while (Rend[writep] != '\0') {// ">
            wd[p] = Rend[writep]; p++; writep++;
        }writep = 0;

        rows = rows->next;
    }
    while (shend[writep] != '\0') {
        wd[p] = shend[writep]; p++; writep++;
    }writep = 0;
}
// cタグ　書き込み
void Ctags::writec(C* ctag, UINT8* ROW) {
    const char* cstart = "<c r=\"";
    const char* ctags = "\" s=\"";
    const char* ctagt = "\" t=\"";
    const char* ftag = "<f";
    const char* ftagt = " t=\"";
    const char* ftagref = "\" ref=\"";
    const char* ftagsi = "\" si=\"";
    const char* ftagend = "</f>";
    const char* vtagend = "</v>";
    const char* ctagend = "</c>";
    const char* cvend = "\">";
    int place = 0;

    while (ctag) {
        while (cstart[writep] != '\0') {// thickbot
            wd[p] = cstart[writep]; p++; writep++;
        }writep = 0;
        UINT8* Col = NA.ColumnNumtoArray(ctag->col, &place);
        while (Col[writep] != '\0') {//r 列
            wd[p] = Col[writep]; p++; writep++;
        }writep = 0;
        while (ROW[writep] != '\0') {//r 行
            wd[p] = ROW[writep]; p++; writep++;
        }writep = 0;
        if (ctag->s) {
            while (ctags[writep] != '\0') {// s
                wd[p] = ctags[writep]; p++; writep++;
            }writep = 0;
            while (ctag->s[writep] != '\0') {// s
                wd[p] = ctag->s[writep]; p++; writep++;
            }writep = 0;
        }
        if (ctag->t) {//t あり
            while (ctagt[writep] != '\0') {// t
                wd[p] = ctagt[writep]; p++; writep++;
            }writep = 0;
            while (ctag->t[writep] != '\0') {// t
                wd[p] = ctag->t[writep]; p++; writep++;
            }writep = 0;
        }
        if (ctag->f) {//Fあり
            while (cvend[writep] != '\0') {// ">
                wd[p] = cvend[writep]; p++; writep++;
            }writep = 0;
            while (ftag[writep] != '\0') {// f
                wd[p] = ftag[writep];
                p++; writep++;
            }writep = 0;
            if (ctag->f->t) {
                while (ftagt[writep] != '\0') {// t
                    wd[p] = ftagt[writep]; p++; writep++;
                }writep = 0;
                while (ctag->f->t[writep] != '\0') {// t
                    wd[p] = ctag->f->t[writep];
                    p++; writep++;
                }writep = 0;
            }
            if (ctag->f->ref) {
                while (ftagref[writep] != '\0') {// f ref
                    wd[p] = ftagref[writep]; p++; writep++;
                }writep = 0;
                while (ctag->f->ref[writep] != '\0') {// f ref
                    wd[p] = ctag->f->ref[writep]; p++; writep++;
                }writep = 0;
            }
            if (ctag->f->si) {
                while (ftagsi[writep] != '\0') {// f si
                    wd[p] = ftagsi[writep]; p++; writep++;
                }writep = 0;
                while (ctag->f->si[writep] != '\0') {// f si
                    wd[p] = ctag->f->si[writep]; p++; writep++;
                }writep = 0;
            }
            if (ctag->f->val) {//f vあり 計算式
                if (ctag->f->t) {//tあり ">閉じ v
                    wd[p] = '"';
                    p++;
                }
                wd[p] = '>';//tなし
                p++;

                while (ctag->f->val[writep] != '\0') {//f val
                    wd[p] = ctag->f->val[writep]; p++; writep++;
                }writep = 0;
                while (ftagend[writep] != '\0') {// f end
                    wd[p] = ftagend[writep]; p++; writep++;
                }writep = 0;
            }
            else {
                while (closetag[writep] != '\0') {//vなし "/>とじ
                    wd[p] = closetag[writep]; p++; writep++;
                }writep = 0;
            }
        }
        if (ctag->val) {//V あり
            if (!ctag->f) {//fない場合閉じる
                while (cvend[writep] != '\0') {// ">
                    wd[p] = cvend[writep]; p++; writep++;
                }writep = 0;
            }
            while (Vtag[writep] != '\0') {// v
                wd[p] = Vtag[writep]; p++; writep++;
            }writep = 0;
            while (ctag->val[writep] != '\0') {// v val
                wd[p] = ctag->val[writep]; p++; writep++;
            }writep = 0;
            while (vtagend[writep] != '\0') {// /v
                wd[p] = vtagend[writep]; p++; writep++;
            }writep = 0;

            while (ctagend[writep] != '\0') {// /c end
                wd[p] = ctagend[writep]; p++; writep++;
            }writep = 0;
        }
        else {//V なし　クローズ
            while (closetag[writep] != '\0') {
                wd[p] = closetag[writep]; p++; writep++;
            }writep = 0;
        }
        ctag = ctag->next;
    }
}
//最終文字列書き込み
void Ctags::writefinal() {
    while (fstr[writep] != '\0') {// /v
        wd[p] = fstr[writep]; p++; writep++;
    }writep = 0;
}

UINT8* Ctags::StrInit()
{
    UINT8* p = (UINT8*)malloc(1);
    p = nullptr;
    return p;
}
