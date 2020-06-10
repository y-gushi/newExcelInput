#include "RowColumn.h"

//�Z���f�[�^�̒ǉ��@row span�X�V
void Ctags::addcelldata(UINT8* row, UINT8* col, UINT8* t, UINT8* s, UINT8* v, F* f, UINT8* si) {
    Row* ro = nullptr;
    bool replace = false;
    int rp = 0;//����
    while (row[rp] != '\0')
        rp++;
    UINT32 rn = NA.RowArraytoNum(row, rp);//�s�z�񂩂琔��

    rp = 0;//����
    while (col[rp] != '\0')
        rp++;
    UINT32 cn = NA.ColumnArraytoNumber(col, rp);//span�悤�@�񕶎�����

    if (cn > maxcol) {//diment �ő�l�X�V
        maxcol = cn;
        dm->eC = NA.ColumnNumtoArray(maxcol, &rp);
        replace = true;
    }

    ro = searchRow(rows, rn);//row�ʒu����
    std::cout << " �ǉ� v : " << v << std::endl;
    ro->cells = addCtable(ro->cells, t, s, si, cn, v, f);//�Z����񌟍�
    cn = NA.ColumnCharnumtoNumber(cn);//��ԍ��A�Ԃց@���ꂽ��

    if (replace) {//�ő�l�X�V�Ł@cols style�m�F
        cols* ncols = cls;
        while (ncols->next)//cols�@�Ō�Q��
            ncols = ncols->next;
        int j = 0;
        while (ncols->max[j] != '\0')
            j++;
        UINT32 ncn = NA.RowArraytoNum(ncols->max, j);//�ŏImax������ �����ɕϊ�
        if (cn > ncn) {//cols max ������
            UINT8 wi[] = "9"; UINT8 styl[] = "1";
            cls = addcolatyle(cls, col, col, wi, styl, nullptr, nullptr, nullptr);//cols �ǉ�
        }
    }
    rp = 0;//����
    while (ro->spanE[rp] != '\0')
        rp++;
    UINT32 nnc = NA.RowArraytoNum(ro->spanE, rp);//span �����񐔎���
    if (cn > nnc) {//�V���������񂪑傫���ꍇ�@�X�V
        UINT8* newspan = NA.InttoChar(cn, &rp);
        ro->spanE = newspan;
        //std::cout << " new span : " << ro->spanE << std::endl;
    }
}

void Ctags::writesheetdata() {
    size_t msize= size_t(dlen) + 6000;

    wd = (UINT8*)malloc(msize);//�������T�C�Y�ύX�@�������ݗp
    //std::cout << "�f�[�^�X�V" << p << std::endl;
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
//diment��������
void Ctags::writeDiment() {
    const char* dimstart = "<dimension ref=\"";
    std::cout << "�f�B�����g�X�V" << p << std::endl;
    while (dimstart[writep] != '\0') {
        wd[p] = dimstart[writep]; p++; writep++;
    }writep = 0;
    while (dm->sC[writep] != '\0') {//�X�^�[�g��
        wd[p] = dm->sC[writep]; p++; writep++;
    }writep = 0;
    while (dm->sR[writep] != '\0') {//�X�^�[�g�s
        wd[p] = dm->sR[writep]; p++; writep++;
    }writep = 0;
    wd[p] = ':'; p++;
    while (dm->eC[writep] != '\0') {//�I����
        wd[p] = dm->eC[writep]; p++; writep++;
    }writep = 0;
    while (dm->eR[writep] != '\0') {//�I���s
        wd[p] = dm->eR[writep]; p++; writep++;
    }writep = 0;
    std::cout << "�f�B�����g�X�V" << dm->eC << std::endl;
}
//cols��������
void Ctags::writecols() {
    const char* startcols = "<cols>";
    const char* colstr[] = { "<col min=\"","\" max=\"","\" width=\"","\" style=\""
        ,"\" bestFit=\"","\" hidden=\"","\" customWidth=\"" };

    std::cout << "�R���X�X�V" << p << std::endl;
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
//selection pane��������
void Ctags::writeSelection() {
    const char* selpane[] = { "<selection pane=\"","\" activeCell=\"","\" sqref=\"","</sheetView>","</sheetViews>" };
    const char* panes[] = { "<pane"," xSplit=\""," ySplit=\""," topLeftCell=\""," activePane=\""," state=\"" };
    std::cout << "�Z���N�V�����X�V" << p << std::endl;

    // <sheetview��������
    while (dimtopane[writep] != '\0') {
        wd[p] = dimtopane[writep]; p++; writep++;
    }writep = 0;

    // <pane ��������
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

    // <selection ��������
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
//�Z���f�[�^��������
void Ctags::writecells() {
    const char* sheetstart = "<sheetData>";
    const char* shend = "</sheetData>";
    std::cout << "�Z���X�V" << p << std::endl;

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
        UINT8* RN = NA.InttoChar(rows->r, &Place);//�����@�z��ϊ�
        while (RN[writep] != '\0') {// r
            wd[p] = RN[writep]; p++; writep++;
        }writep = 0;
        if (rows->spanS) {//�Ȃ��ꍇ����
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

        UINT8* rne = NA.InttoChar(rows->r, &Place);//�z��ɕύX
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
// c�^�O�@��������
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
        while (Col[writep] != '\0') {//r ��
            wd[p] = Col[writep]; p++; writep++;
        }writep = 0;
        while (ROW[writep] != '\0') {//r �s
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
        if (ctag->t) {//t ����
            while (ctagt[writep] != '\0') {// t
                wd[p] = ctagt[writep]; p++; writep++;
            }writep = 0;
            while (ctag->t[writep] != '\0') {// t
                wd[p] = ctag->t[writep]; p++; writep++;
            }writep = 0;
        }
        if (ctag->f) {//F����
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
            if (ctag->f->val) {//f v���� �v�Z��
                if (ctag->f->t) {//t���� ">�� v
                    wd[p] = '"';
                    p++;
                }
                wd[p] = '>';//t�Ȃ�
                p++;

                while (ctag->f->val[writep] != '\0') {//f val
                    wd[p] = ctag->f->val[writep]; p++; writep++;
                }writep = 0;
                while (ftagend[writep] != '\0') {// f end
                    wd[p] = ftagend[writep]; p++; writep++;
                }writep = 0;
            }
            else {
                while (closetag[writep] != '\0') {//v�Ȃ� "/>�Ƃ�
                    wd[p] = closetag[writep]; p++; writep++;
                }writep = 0;
            }
        }
        if (ctag->val) {//V ����
            if (!ctag->f) {//f�Ȃ��ꍇ����
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
        else {//V �Ȃ��@�N���[�Y
            while (closetag[writep] != '\0') {
                wd[p] = closetag[writep]; p++; writep++;
            }writep = 0;
        }
        ctag = ctag->next;
    }
}
//�ŏI�����񏑂�����
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
