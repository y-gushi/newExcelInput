#include "RowColumn.h"

void Ctags::sheetread() {
    GetSheetPr();
    GetDiment();
    GetSelectionPane();
    Getcols();
    GetCtagValue();
    getfinalstr();
}
//c tag add table
C* Ctags::addCtable(C* c, UINT8* tv, UINT8* sv, UINT8* si, UINT32 col, UINT8* v, F* fv) {
    if (!c) {
        c = (C*)malloc(sizeof(C));
        if (c) {
            c->val = v;
            c->t = tv;
            c->s = sv;
            c->si = si;
            c->col = col;
            c->f = fv;
            c->next = nullptr;
        }
    }
    else if (col == c->col) {//��ԍ������@���X�V
        c->val = v;
        c->t = tv;
        c->s = sv;
        c->si = si;
        c->col = col;
        c->f = fv;
    }
    else if (col < c->col) {
        C* p = (C*)malloc(sizeof(C));
        if (p && c) {
            p->val = c->val;
            p->t = c->t;
            p->s = c->s;
            p->si = c->si;
            p->col = c->col;
            p->f = c->f;
            p->next = c->next;
        }

        c = (C*)malloc(sizeof(C));
        if (c) {
            c->val = v;
            c->t = tv;
            c->s = sv;
            c->si = si;
            c->col = col;
            c->f = fv;
            c->next = p;
        }
    }
    else {
        c->next = addCtable(c->next, tv, sv, si, col, v, fv);
    }
    return c;
}

Ctags::Ctags(UINT8* decorddata, UINT64 datalen, shareRandD* shdata)
{
    data = decorddata;
    dlen = datalen;
    sh = shdata;
    coden = nullptr;
    wd = nullptr;
}
//�f�X�g���N�^
Ctags::~Ctags()
{
    Rowtablefree();
    selectfree();
    colfree();
    panefree();

    free(headXML);
    free(dimtopane);
    free(sFPr);
    //free(sct);
    free(dm);
    //free(cls);
    //free(Panes);
    //free(MC);

    free(wd);
}

//c tag v��z���
void Ctags::GetCtagValue() {

    int slen = 0; int tlen = 0; int flen = 0;
    UINT8* Tv = (UINT8*)malloc(1); Tv = nullptr;
    UINT8* Sv = (UINT8*)malloc(1); Sv = nullptr;
    UINT8* Vv = (UINT8*)malloc(1); Vv = nullptr;
    UINT32 Vnum = 0;
    UINT8* col = (UINT8*)malloc(1); col = nullptr;
    UINT32 colnum = 0; UINT32 rownum = 0;
    UINT8* row = (UINT8*)malloc(1); row = nullptr;

    int collen = 0; int rowlen = 0; UINT32 vlen = 0; int endresult = 1;

    UINT8* fv = (UINT8*)malloc(1); fv = nullptr;//f tag value
    UINT8* ft = (UINT8*)malloc(1); ft = nullptr;// f tag t
    UINT8* fre = (UINT8*)malloc(1); fre = nullptr;// f tag ref
    UINT8* fsi = (UINT8*)malloc(1); fsi = nullptr;// f tag si

    UINT32 limcol = NA.ColumnArraytoNumber(dm->eC, eClen);//�������l�ϊ�
    limcol = NA.ColumnCharnumtoNumber(limcol);

    Row* Srow = (Row*)malloc(sizeof(Row));//row���� c����
    UINT8* rs = (UINT8*)malloc(4);// <row �����p
    UINT8* sr = (UINT8*)malloc(4);// <c �����p
    UINT8* ftag = (UINT8*)malloc(2);//f �����X���C�h�p
    UINT8* shend = (UINT8*)malloc(12);//f �����X���C�h�p
    UINT8* slt = (UINT8*)malloc(3);//f �����X���C�h�p
    UINT8* fsl = (UINT8*)malloc(2);//f tag �����X���C�h�p

    F* fvs = nullptr;

    while (endresult != 0) {
        for (int j = 0; j < 12 - 1; j++) {
            shend[j] = shend[j + 1];
            if (j < 4 - 1)
                rs[j] = rs[j + 1];
            if (j < (3 - 1))
                sr[j] = sr[j + 1];
        }
        shend[12 - 1] = rs[4 - 1] = sr[3 - 1] = data[p];
        p++;//�ʒu�ړ�

        endresult = strncmp((char const*)shend, sheetend, 12);// </sheetdata>

        if (strncmp((char const*)rs, Rowtag, 4) == 0) {
            Getrow();//row tag get
        }

        if (strncmp((char const*)sr, Ctag, 3) == 0) {// <c match
            while (data[p - 1] != '"')
                p++;
            collen = 0; rowlen = 0;
            while (data[p + collen + rowlen] != '"') {//��s get
                if (data[p + collen + rowlen] > 64)//��
                    collen++;//�������J�E���g
                else//�s
                    rowlen++;//�������J�E���g
            }
            UINT32 csize = (UINT32)collen + 1;
            UINT32 rsize = (UINT32)rowlen + 1;
            col = (UINT8*)malloc(csize);
            row = (UINT8*)malloc(rsize);

            for (int i = 0; i < collen; i++) {
                col[i] = data[p]; p++;
            }
            for (int i = 0; i < rowlen; i++) {
                row[i] = data[p]; p++;
            }
            col[collen] = '\0'; row[rowlen] = '\0';
            colnum = NA.ColumnArraytoNumber(col, collen);//�������l�ϊ�
            rownum = NA.RowArraytoNum(row, rowlen);//�����ϊ�
            //std::cout << " col : " << col << " row : " << row << std::endl;

            while (data[p] != '>') {//t s ����
                for (int j = 0; j < 3 - 1; j++) {
                    slt[j] = slt[j + 1];
                }
                slt[3 - 1] = data[p];
                p++;//�ʒu�ړ�

                if (strncmp((char const*)slt, Rs, 3) == 0) {//s get
                    slen = 0;
                    while (data[p + slen] != '"')
                        slen++;
                    free(Sv);
                    UINT32 ssize = (UINT32)slen + 1;
                    Sv = (UINT8*)malloc(ssize);
                    for (int i = 0; i < slen; i++) {
                        Sv[i] = data[p]; p++;
                    }
                    Sv[slen] = '\0';
                }
                if (strncmp((char const*)slt, Ft, 3) == 0) {//t get
                    tlen = 0;
                    while (data[p + tlen] != '"')
                        tlen++;
                    UINT32 tsize = (UINT32)tlen + 1;
                    free(Tv);
                    Tv = (UINT8*)malloc(tsize);
                    for (int i = 0; i < tlen; i++) {
                        Tv[i] = data[p]; p++;
                    }
                    Tv[tlen] = '\0';
                }
            }

            //v get
            if (data[p - 1] == '/') {//no v data
                Srow = searchRow(rows, rownum);
                if (Srow) {
                    Srow->cells = addCtable(Srow->cells, Tv, Sv, nullptr, colnum, Vv, fvs);
                }
                else {
                    //need make new row
                }
                free(col); free(row);
                col = nullptr; row = nullptr;

                Tv = (UINT8*)malloc(1); Tv = nullptr;
                Sv = (UINT8*)malloc(1); Sv = nullptr;
                ft = (UINT8*)malloc(1); ft = nullptr;
                fre = (UINT8*)malloc(1); fre = nullptr;
                fsi = (UINT8*)malloc(1); fsi = nullptr;
                fv = (UINT8*)malloc(1); fv = nullptr;
                Vv = (UINT8*)malloc(1); Vv = nullptr;
                fvs = (F*)malloc(sizeof(F)); fvs = nullptr;
                ft = 0;
            }
            else {//get v tag
                while (strncmp((char const*)sr, endC, 3) != 0) {// c ���܂�
                    for (int j = 0; j < 3 - 1; j++) {
                        sr[j] = sr[j + 1];
                        if (j < 2 - 1) {
                            fsl[j] = fsl[j + 1];
                        }
                    }
                    fsl[2 - 1] = sr[3 - 1] = data[p];//
                    p++;

                    if (strncmp((char const*)sr, Vtag, 3) == 0) {//get v tag
                        vlen = 0;
                        while (data[p + vlen] != '<')
                            vlen++;
                        UINT32 vsize = vlen + 1;
                        free(Vv);
                        Vv = (UINT8*)malloc(vsize);
                        for (UINT32 i = 0; i < vlen; i++) {
                            Vv[i] = data[p];
                            p++;
                        }
                        Vv[vlen] = '\0';

                        if (Tv && tlen == 1 && Tv[0]!='e') {//share data new str get & t=e�ȊO
                            Vnum = NA.RowArraytoNum(Vv, vlen);//�����ɕϊ�
                            UINT32 sistrlen = 0;//si �^�O�� <t>�����̑�������������  //UINT8* Vsi = sh->sis[Vnum]->Ts;//share str into struct�@<t>��������
                            UINT32 Tstrlen = 0;

                            Si* RootSi = sh->sis[Vnum];//�e�@�Q��
                            UINT8* Vsi = nullptr;
                            
                            while (RootSi) {//t �^�O�@����������@�����m�F
                                Tstrlen = 0;
                                while (RootSi->Ts[Tstrlen] != '\0') {
                                    sistrlen++; Tstrlen++;
                                }
                                RootSi = RootSi->next;
                            }

                            RootSi = sh->sis[Vnum];

                            if (sistrlen > 0) {
                                sistrlen = sistrlen+1;
                                //free(Vsi);
                                Vsi = (UINT8*)malloc(sistrlen);//�������m��

                                sistrlen = 0;

                                while (RootSi) {//t�^�O�������킹��
                                    Tstrlen = 0;//������
                                    while (RootSi->Ts[Tstrlen] != '\0') {
                                        Vsi[sistrlen] = RootSi->Ts[Tstrlen];
                                        sistrlen++; Tstrlen++;
                                    }
                                    RootSi = RootSi->next;
                                }
                                Vsi[sistrlen] = '\0';
                            }                          

                            Srow = searchRow(rows, rownum);
                            if (Srow) {
                                Srow->cells = addCtable(Srow->cells, Tv, Sv, Vsi, colnum, Vv, fvs);
                               // Vsi = (UINT8*)malloc(1);
                                Vsi = nullptr;
                                RootSi = nullptr;
                            }
                            else {
                                //need make new row
                            }
                        }
                        else {
                            //not share data str or v in
                            Srow = searchRow(rows, rownum);
                            if (Srow) {
                                Srow->cells = addCtable(Srow->cells, Tv, Sv, nullptr, colnum, Vv, fvs);
                            }
                            else {
                                //need make new row
                            }
                        }
                        free(col); free(row);
                        col = nullptr; row = nullptr;
                        Vv = (UINT8*)malloc(1); Vv = nullptr;
                        Tv = (UINT8*)malloc(1); Tv = nullptr;
                        Sv = (UINT8*)malloc(1); Sv = nullptr;
                        ft = (UINT8*)malloc(1); ft = nullptr;
                        fre = (UINT8*)malloc(1); fre = nullptr;
                        fsi = (UINT8*)malloc(1); fsi = nullptr;
                        fv = (UINT8*)malloc(1); fv = nullptr;
                        fvs = (F*)malloc(sizeof(F)); fvs = nullptr;
                        ft = 0;
                    }

                    if (strncmp((char const*)fsl, Ftag, 2) == 0) {// <f t="shared" si="13"/>
                        // <f t="shared" ref="G12:G29" si="0">F12-C12</f>
                        free(fvs);
                        fvs = (F*)malloc(sizeof(F));
                        while (data[p] != '>') {
                            for (int j = 0; j < 2 - 1; j++) {
                                ftag[j] = ftag[j + 1];
                            }
                            ftag[2 - 1] = data[p]; p++;//�ʒu�ړ�

                            if (strncmp((char const*)ftag, Fref, 2) == 0) {// ref macth
                                flen = 0;
                                while (data[p - 1] != '"')// f ref
                                    p++;
                                while (data[p + flen] != '"')// f ref v get
                                    flen++;
                                UINT32 fsize = (UINT32)flen + 1;
                                free(fre);
                                fre = (UINT8*)malloc(fsize);
                                for (int i = 0; i < flen; i++) {
                                    fre[i] = data[p]; p++;
                                }
                                fre[flen] = '\0';
                                fvs->ref = fre;
                            }

                            if (strncmp((char const*)ftag, Fsi, 2) == 0) {// si macth
                                flen = 0;
                                while (data[p - 1] != '"')// f si
                                    p++;
                                while (data[p + flen] != '"')// f si v get
                                    flen++;
                                UINT32 fsize = (UINT32)flen + 1;
                                free(fsi);
                                fsi = (UINT8*)malloc(fsize);
                                for (int i = 0; i < flen; i++) {
                                    fsi[i] = data[p]; p++;
                                }
                                fsi[flen] = '\0';
                                fvs->si = fsi;
                            }

                            if (strncmp((char const*)ftag, Ft, 2) == 0) {// t macth
                                flen = 0;
                                while (data[p - 1] != '"')// f ref
                                    p++;
                                while (data[p + flen] != '"')// f ref v get
                                    flen++;
                                UINT32 fsize = (UINT32)flen + 1;
                                free(ft);
                                ft = (UINT8*)malloc(fsize);
                                for (int i = 0; i < flen; i++) {
                                    ft[i] = data[p]; p++;
                                }
                                ft[flen] = '\0';
                                fvs->t = ft;
                            }
                        }
                        p++;//> �@�Ƃ΂�
                        // </f>�܂ł̒l
                        flen = 0;
                        while (data[p + flen] != '<')
                            flen++;
                        if (flen > 0) {
                            free(fv);
                            UINT32 fsize = (UINT32)flen + 1;
                            fv = (UINT8*)malloc(fsize);
                            for (int i = 0; i < flen; i++) {
                                fv[i] = data[p]; p++;// f ta v get
                            }
                            fv[flen] = '\0';

                            fvs->val = fv;
                        }
                        if (!fre)
                            fvs->ref = nullptr;
                        if (!fsi)
                            fvs->si = nullptr;
                        if (!ft)
                            fvs->t = nullptr;
                        if (!fv)
                            fvs->val = nullptr;
                    }
                }
            }
        }
    }
    free(rs);// <row �����p
    free(sr);// <c �����p
    free(ftag);//f �����X���C�h�p
    free(shend);//f �����X���C�h�p
    free(slt);//f �����X���C�h�p
    free(fsl);//f tag �����X���C�h�p
}
// <sheetPr get
void Ctags::GetSheetPr() {
    /* <sheetPr>
<tabColor tint="0.4" theme="7"/>
<pageSetUpPr fitToPage="1"/>
</sheetPr>*/
    //const char* shPr = "<sheetPr";//8
    //const char* shPrEnd = "</sheetPr>";//10
    const char* diment = "<dimension";//10
    UINT8* pr = (UINT8*)malloc(10);
    p = 0;//sheetdata�ŏ�����

    int ucode = 0;//dimention�܂� <sheetPr �Ƃ΂�
    
    if (pr) {
        int result = 0;
        do{
            for (int j = 0; j < 10 - 1; j++) {
                pr[j] = pr[j + 1];
            }
            pr[10 - 1] = data[p + ucode];
            ucode++;

            result = strncmp((char const*)pr, diment, 10);
        } while (result != 0);
        ucode -= 10;
        UINT32 msize = (UINT32)ucode + 1;
        headXML = (UINT8*)malloc(msize);
        if (headXML) {
            for (int i = 0; i < ucode; i++) {//dimention�O�܂ŃR�s�[
                headXML[i] = data[p];
                p++;
            }
            headXML[ucode] = '\0';
        }
        free(pr);
    }
    
}

//row �e�[�u���ǉ�
Row* Ctags::addrows(Row* row, UINT32 r, UINT8* spanS, UINT8* spanE, UINT8* ht, UINT8* thickBot, UINT8* s, UINT8* customFormat, UINT8* customHeight)
{
    if (!row) {
        row = (Row*)malloc(sizeof(Row));
        if (row) {
            row->s = s;
            row->spanS = spanS;
            row->spanE = spanE;
            row->ht = ht;
            row->thickBot = thickBot;
            row->r = r;
            row->customFormat = customFormat;
            row->customHeight = customHeight;
            row->cells = nullptr;
            row->next = nullptr;
        }
    }
    else if (r == row->r) {//�s�ԍ������@�X�V
        row->s = s;
        row->spanS = spanS;
        row->spanE = spanE;
        row->ht = ht;
        row->thickBot = thickBot;
        row->r = r;
        row->customFormat = customFormat;
        row->customHeight = customHeight;
    }
    else {
        row->next = addrows(row->next, r, spanS, spanE, ht, thickBot, s, customFormat, customHeight);
    }
    return row;
}
//row�ԍ��i������j�Ńe�[�u������
Row* Ctags::searchRow(Row* r, UINT32 newrow) {
    Row* sr = r;
    while (sr) {
        if (newrow == sr->r) {
            return sr;
        }
        sr = sr->next;
    }
    return nullptr;
}
//row���e�擾 �z��� p >�܂�
void Ctags::Getrow() {
    // <row r="11" spans="1:21" s="4" customFormat="1" ht="48.75" customHeight="1" thickBot="1">
    UINT8* cusf = (UINT8*)malloc(15);
    UINT8* rb = (UINT8*)malloc(11);
    UINT8* spa = (UINT8*)malloc(8);
    UINT8* ht = (UINT8*)malloc(5);
    UINT8* rs = (UINT8*)malloc(4);
    int vlen = 0; UINT32 rownum = 0;

    //�l�p
    UINT8* s = (UINT8*)malloc(1); s = nullptr;
    UINT8* r = (UINT8*)malloc(1); r= nullptr;
    UINT8* h = (UINT8*)malloc(1); h= nullptr;
    UINT8* sps = (UINT8*)malloc(1); sps= nullptr;
    UINT8* spe = (UINT8*)malloc(1); spe= nullptr;
    UINT8* tb = (UINT8*)malloc(1); tb = nullptr;
    UINT8* ch = (UINT8*)malloc(1); ch= nullptr;
    UINT8* cf = (UINT8*)malloc(1); cf= nullptr;

    while (data[p] != '>') {
        for (int j = 0; j < 14 - 1; j++) {
            cusf[j] = cusf[j + 1];
            if (j < 10 - 1)
                rb[j] = rb[j + 1];
            if (j < 7 - 1)
                spa[j] = spa[j + 1];
            if (j < 4 - 1)
                ht[j] = ht[j + 1];
            if (j < 3 - 1)
                rs[j] = rs[j + 1];
        }
        cusf[14 - 1] = rb[10 - 1] = spa[7 - 1] = ht[4 - 1] = rs[3 - 1] = data[p];
        rs[3] = '\0';
        p++;
        if (strncmp((char const*)rs, Rs, 3) == 0 && strncmp((char const*)spa, Rspans, 7) != 0) {
            vlen = 0;
            while (data[p + vlen] != '"')
                vlen++;
            UINT32 vsize = (UINT32)vlen + 1;
            free(s);
            s = (UINT8*)malloc(vsize);
            for (int i = 0; i < vlen; i++) {
                s[i] = data[p];
                p++;
            }
            s[vlen] = '\0';
        }
        if (strncmp((char const*)rs, Rr, 3) == 0) {
            vlen = 0;
            while (data[p + vlen] != '"')
                vlen++;
            UINT32 vsize = (UINT32)vlen + 1;
            free(r);
            r = (UINT8*)malloc(vsize);
            for (int i = 0; i < vlen; i++) {
                r[i] = data[p];
                p++;
            }
            r[vlen] = '\0';
            //std::cout << " row : " << r << std::endl;
            rownum = NA.RowArraytoNum(r, vlen);
        }
        if (strncmp((char const*)ht, Rht, 4) == 0 && strncmp((char const*)cusf, RcustomHeight, 14) != 0) {
            vlen = 0;
            while (data[p + vlen] != '"')
                vlen++;
            free(h);
            UINT32 vsize = (UINT32)vlen + 1;
            h = (UINT8*)malloc(vsize);
            for (int i = 0; i < vlen; i++) {
                h[i] = data[p];
                p++;
            }
            h[vlen] = '\0';
        }
        if (strncmp((char const*)spa, Rspans, 7) == 0) {
            vlen = 0;
            while (data[p + vlen] != ':')//span �ŏ�
                vlen++;
            free(sps);
            UINT32 vsize = (UINT32)vlen + 1;
            sps = (UINT8*)malloc(vsize);
            for (int i = 0; i < vlen; i++) {
                sps[i] = data[p];
                p++;
            }
            sps[vlen] = '\0';
            p++;
            
            vlen = 0;
            while (data[p + vlen] != '"')//span �I���
                vlen++;
            free(spe);
            vsize = (UINT32)vlen + 1;
            spe = (UINT8*)malloc(vsize);
            for (int i = 0; i < vlen; i++) {
                spe[i] = data[p];
                p++;
            }
            spe[vlen] = '\0';
            //std::cout << " span : " << sps << " : " << spe << std::endl;
        }
        if (strncmp((char const*)rb, RthickBot, 10) == 0) {
            vlen = 0;
            while (data[p + vlen] != '"')
                vlen++;
            free(tb);
            UINT32 vsize = (UINT32)vlen + 1;
            tb = (UINT8*)malloc(vsize);
            for (int i = 0; i < vlen; i++) {
                tb[i] = data[p];
                p++;
            }
            tb[vlen] = '\0';
        }
        if (strncmp((char const*)cusf, RcustomHeight, 14) == 0) {
            vlen = 0;
            while (data[p + vlen] != '"')
                vlen++;
            free(ch);
            UINT32 vsize = (UINT32)vlen + 1;
            ch = (UINT8*)malloc(vsize);
            for (int i = 0; i < vlen; i++) {
                ch[i] = data[p];
                p++;
            }
            ch[vlen] = '\0';
        }
        if (strncmp((char const*)cusf, RcustomFormat, 14) == 0) {
            vlen = 0;
            while (data[p + vlen] != '"')
                vlen++;
            free(cf);
            UINT32 vsize = (UINT32)vlen + 1;
            cf = (UINT8*)malloc(vsize);
            for (int i = 0; i < vlen; i++) {
                cf[i] = data[p];
                p++;
            }
            cf[vlen] = '\0';
        }
    }
    rows = addrows(rows, rownum, sps, spe, h, tb, s, cf, ch);
    free(cusf); free(spa); free(ht);
    free(rb); free(rs);
}
//demention get
void Ctags::GetDiment() {

    UINT8* sd = (UINT8*)malloc(17);
    int spos = 0;

    dm = (demention*)malloc(sizeof(demention));
    dm->sC = (UINT8*)malloc(5);//demention start ��
    dm->sR = (UINT8*)malloc(5);//demention start �s
    dm->eC = (UINT8*)malloc(5);//demention end ��
    dm->eR = (UINT8*)malloc(5);//demention end �s
    // <dimension ref="A1:EJ128"/>������
    while (spos < dlen) {// : �܂ŃR�s�[
        for (int j = 0; j < 16 - 1; j++) {
            sd[j] = sd[j + 1];
        }
        sd[16 - 1] = data[p];//�Ō�ɕt��������
        p++;

        if (strncmp((char const*)sd, dement, 16) == 0)
            break;
    }
    while (data[p] != ':') {//startcell
        if (data[p] > 64) {//��ԍ�
            dm->sC[sClen] = data[p];
            sClen++;
        }
        else {//�s�ԍ�
            dm->sR[sRlen] = data[p];
            sRlen++;
        }
        p++;
    }
    free(sd);
    dm->sC[sClen] = '\0'; dm->sR[sRlen] = '\0';
    p++;//':' ����
    while (data[p] != '"') {//endcell
        if (data[p] > 64) {//��ԍ�
            dm->eC[eClen] = data[p];
            eClen++;
        }
        else {//�s�ԍ�
            dm->eR[eRlen] = data[p];
            eRlen++;
        }
        p++;
    }
    dm->eC[eClen] = '\0'; dm->eR[eRlen] = '\0';
    maxcol = NA.ColumnArraytoNumber(dm->eC, eClen);//��max��
    //std::cout << "diment �ő��: " << dm->eC << std::endl;
}
//selection tag get�@�e�[�u����
void Ctags::GetSelectionPane() {
    const char* pane = "<pane";//5
    const char* ecp = "pane=\"";//6
    const char* active = "activeCell=\"";//12
    const char* sqref = "sqref=\"";//7
    const char* select = "<selection";//10

    int panelen = 0;
    int activeCelllen = 0;
    int sqreflen = 0;

    UINT8* at = (UINT8*)malloc(12);
    UINT8* sve = (UINT8*)malloc(13);
    UINT8* pn = (UINT8*)malloc(6);
    UINT8* sr = (UINT8*)malloc(7);
    UINT8* PN = (UINT8*)malloc(10);
    UINT8* slcs = (UINT8*)malloc(10);
    int PNlen = 0;
    UINT8* pa = (UINT8*)malloc(1);
    UINT8* ac = (UINT8*)malloc(1);
    UINT8* sq = (UINT8*)malloc(1);
    UINT8* PT = (UINT8*)malloc(5);

    int res = 1;
    int panetagres = 0;

    while (res != 0)
    {//diment /> to <sheetView �܂�
        for (int j = 0; j < 10 - 1; j++) {//�������J�E���g
            PN[j] = PN[j + 1];
        }
        PN[10 - 1] = data[p + PNlen];//�Ō�ɕt��������
        PNlen++;

        res = strncmp((char const*)PN, startSV, 10);
    }
    //std::cout << "data[p + PNlen] : " << data[p + PNlen] << std::endl;
    if (data[p + PNlen] == 's') {// <sheetViews>tag
        PNlen += 3;// <sheetViews> ���^�O�Ƃ΂�
    }
    while (data[p + PNlen - 1] != '>')//�������J�E���g sheetView >���܂�
        PNlen++;
    free(PN);
    UINT32 msize = (UINT32)PNlen + 1;
    dimtopane = (UINT8*)malloc(msize);
    if (dimtopane) {
        for (int i = 0; i < PNlen; i++) {///to <sheetView ������R�s�[
            dimtopane[i] = data[p];
            p++;
        }
        dimtopane[PNlen] = '\0';
        //std::cout << "cim to pane" << dimtopane << std::endl;
    }

    res = 1;
    int seleresult = 0;
    /*
    <selection/>
    <selection pane="topRight"/>
    <selection pane="bottomLeft"/>
    <selection pane="bottomRight" sqref="B191" activeCell="B191"/>
    */
    while (res != 0)//selectionPane vals get
    {
        for (int j = 0; j < 13 - 1; j++) {
            sve[j] = sve[j + 1];
            if (j < 10 - 1) {
                slcs[j] = slcs[j + 1];
            }
            if (j < 5 - 1) {
                PT[j] = PT[j + 1];
            }
        }
        PT[5 - 1] = sve[13 - 1] = slcs[10 - 1] = data[p];
        p++;

        seleresult = strncmp((char const*)slcs, select, 10);
        panetagres = strncmp((char const*)PT, pane, 5);
        if (panetagres == 0) {// <pane tag get
            //std::cout << "get selection  pane" << std::endl;
            GetPane();
        }
        if (seleresult == 0) {
            pa = nullptr; ac = nullptr; sq = nullptr;
            if (data[p] == '/') {//tag�I��
                sct = SLTaddtable(sct, nullptr, nullptr, nullptr);//���ׂ�nullptr�킽��
            }
            else {
                while (data[p] != '>') {
                    for (int j = 0; j < 12 - 1; j++) {
                        at[j] = at[j + 1];//active
                        if (j < (6 - 1))
                            pn[j] = pn[j + 1];//pane
                        if (j < (7 - 1))
                            sr[j] = sr[j + 1];//sqref
                    }
                    sr[7 - 1] = pn[6 - 1] = at[12 - 1] = data[p];//�Ō�ɕt��������
                    p++;

                    if (strncmp((char const*)pn, ecp, 6) == 0) {//pane=" ��v
                        panelen = 0;
                        while (data[p + panelen] != '"')
                            panelen++;//�������J�E���g
                        free(pa);
                        UINT32 pasize = (UINT32)panelen + 1;
                        pa = (UINT8*)malloc(pasize);
                        for (int i = 0; i < panelen; i++) {
                            pa[i] = data[p];//pane str
                            p++;
                        }
                        pa[panelen] = '\0';
                    }

                    if (strncmp((char const*)at, active, 12) == 0) {//activeCell=" ��v
                        activeCelllen = 0;
                        while (data[p + activeCelllen] != '"')
                            activeCelllen++;//�������J�E���g
                        free(ac);
                        UINT32 acsize = (UINT32)activeCelllen + 1;
                        ac = (UINT8*)malloc(acsize);
                        for (int i = 0; i < activeCelllen; i++) {
                            ac[i] = data[p];//pane str
                            p++;
                        }
                        ac[activeCelllen] = '\0';
                    }

                    if (strncmp((char const*)sr, sqref, 7) == 0) {//sqref=" ��v
                        sqreflen = 0;
                        while (data[p + sqreflen] != '"')
                            sqreflen++;//�������J�E���g
                        free(sq);
                        UINT32 sqsize = (UINT32)sqreflen + 1;
                        sq = (UINT8*)malloc(sqsize);
                        for (int i = 0; i < sqreflen; i++) {
                            sq[i] = data[p];//pane str
                            p++;
                        }
                        sq[sqreflen] = '\0';
                    }
                }
                if (pa || ac || sq) {
                    sct = SLTaddtable(sct, pa, ac, sq);//�\���̂փR�s�[
                    sq = (UINT8*)malloc(1);
                    pa = (UINT8*)malloc(1);
                    ac = (UINT8*)malloc(1);
                    sq = nullptr; pa = nullptr; ac = nullptr;
                }
            }
        }
        res = strncmp((char const*)sve, SVend, 13);// </sheetViews>�Ŕ�����
    }
    free(PT); free(sve); free(slcs); free(sr); free(pn); free(at);
}

void Ctags::GetPane() {
    const char* PaneTags[] = { "xSplit=\"","ySplit=\"","topLeftCell=\"","activePane=\"","state=\"" };
    //8 8 13 12 7
    UINT8* XYs = (UINT8*)malloc(8);
    UINT8* tpl = (UINT8*)malloc(13);
    UINT8* acP = (UINT8*)malloc(12);
    UINT8* stt = (UINT8*)malloc(7);

    UINT8* X = (UINT8*)malloc(1); UINT8* Y = (UINT8*)malloc(1);
    UINT8* tL = (UINT8*)malloc(1); UINT8* ap = (UINT8*)malloc(1);
    UINT8* s = (UINT8*)malloc(1);

    X = nullptr; Y = nullptr; tL = nullptr;
    ap = nullptr; s = nullptr;

    int len = 0;

    while (data[p] != '>') {
        for (int j = 0; j < 13 - 1; j++) {
            tpl[j] = tpl[j + 1];//topLeftCell
            if (j < (12 - 1))
                acP[j] = acP[j + 1];//activePane
            if (j < (8 - 1))
                XYs[j] = XYs[j + 1];//x,ysplit
            if (j < (7 - 1))
                stt[j] = stt[j + 1];//state=
        }
        tpl[13 - 1] = acP[12 - 1] = XYs[8 - 1] = stt[7 - 1] = data[p];
        p++;

        if ((strncmp((const char*)tpl, PaneTags[2], 13)) == 0) {
            len = 0;
            while (data[p + len] != '"')
                len++;
            free(tL);
            tL = (UINT8*)malloc(len + 1);
            for (int i = 0; i < len; i++) {
                tL[i] = data[p]; p++;
            }
            tL[len] = '\0';
        }
        if ((strncmp((const char*)acP, PaneTags[3], 12)) == 0) {
            len = 0;
            while (data[p + len] != '"')
                len++;
            free(ap);
            ap = (UINT8*)malloc(len + 1);
            for (int i = 0; i < len; i++) {
                ap[i] = data[p]; p++;
            }
            ap[len] = '\0';
            std::cout << "pane action : " << ap << len << std::endl;
        }
        if ((strncmp((const char*)XYs, PaneTags[0], 8)) == 0) {
            len = 0;
            while (data[p + len] != '"')
                len++;
            free(X);
            X = (UINT8*)malloc(len + 1);
            for (int i = 0; i < len; i++) {
                X[i] = data[p]; p++;
            }
            X[len] = '\0';
        }
        if ((strncmp((const char*)XYs, PaneTags[1], 8)) == 0) {
            len = 0;
            while (data[p + len] != '"')
                len++;
            free(Y);
            Y = (UINT8*)malloc(len + 1);
            for (int i = 0; i < len; i++) {
                Y[i] = data[p]; p++;
            }
            Y[len] = '\0';
        }
        if ((strncmp((const char*)stt, PaneTags[4], 7)) == 0) {
            len = 0;
            while (data[p + len] != '"')
                len++;
            free(s);
            s = (UINT8*)malloc(len + 1);
            for (int i = 0; i < len; i++) {
                s[i] = data[p]; p++;
            }
            s[len] = '\0';
        }
    }
    Panes = addpanetable(Panes, X, Y, tL, ap, s);
    free(tpl); free(acP); free(XYs); free(stt);
}

Pane* Ctags::addpanetable(Pane* p, UINT8* x, UINT8* y, UINT8* tl, UINT8* ap, UINT8* sta) {
    if (!p) {
        p = (Pane*)malloc(sizeof(Pane));
        if (p) {
            p->xSp = x;
            p->ySp = y;
            p->tLeftC = tl;
            p->activP = ap;
            p->state = sta;
            p->next = nullptr;
        }
    }
    else {
        p->next = addpanetable(p->next, x, y, tl, ap, sta);
    }
    return p;
}

selection* Ctags::SLTaddtable(selection* s, UINT8* pv, UINT8* av, UINT8* sv) {
    if (!s) {
        s = (selection*)malloc(sizeof(selection));
        s->p = pv;
        s->a = av;
        s->s = sv;
        s->next = nullptr;
    }
    else {
        s->next = SLTaddtable(s->next, pv, av, sv);
    }

    return s;
}

//col tag get
void Ctags::Getcols() {

    UINT8* search = (UINT8*)malloc(7);
    UINT8* Scoltag = (UINT8*)malloc(5);
    int result = 0;
    int coltaglen = 7;
    unsigned char* min = nullptr;
    unsigned char* max = nullptr;
    unsigned char* style = nullptr;
    unsigned char* bestf = nullptr;
    unsigned char* hidd = nullptr;
    unsigned char* colw = nullptr;
    unsigned char* cuw = nullptr;
    int endresult = 1;

    int sfplen = 0;
    //sheetFormatPr ������
    while (data[p + sfplen] != '>') {
        sfplen++;
    }
    sfplen++;
    UINT32 sfsize = (UINT32)sfplen + 1;
    sFPr = (UINT8*)malloc(sfsize);
    for (int j = 0; j < sfplen; j++) {
        sFPr[j] = data[p];
        p++;
    }
    sFPr[sfplen] = '\0';

    //�z����ꕶ���Â��炷
    while (endresult != 0) {
        for (int j = 0; j < coltaglen - 1; j++) {
            search[j] = search[j + 1];
            if (j < (5 - 1))
                Scoltag[j] = Scoltag[j + 1];
        }
        search[coltaglen - 1] = Scoltag[5 - 1] = data[p];//�Ō�ɕt��������
        p++;

        result = strncmp((char const*)Scoltag, Coltag, 5);
        endresult = strncmp((char const*)search, endtag, 7);
        if (result == 0) {
            UINT8* mimax = (UINT8*)malloc(5);
            UINT8* wist = (UINT8*)malloc(7);
            UINT8* ccw = (UINT8*)malloc(13);
            UINT8* cbf = (UINT8*)malloc(9);
            UINT8* chid = (UINT8*)malloc(8);

            min = nullptr;
            max = nullptr;
            style = nullptr;
            bestf = nullptr;
            hidd = nullptr;
            colw = nullptr;
            cuw = nullptr;
            int StrLen = 0;
            while (data[p] != '>') {
                for (int j = 0; j < 13 - 1; j++) {
                    ccw[j] = ccw[j + 1];
                    if (j < (9 - 1))
                        cbf[j] = cbf[j + 1];//bestFit
                    if (j < (8 - 1))
                        chid[j] = chid[j + 1];//hidden
                    if (j < (7 - 1))
                        wist[j] = wist[j + 1];//style width
                    if (j < (5 - 1))
                        mimax[j] = mimax[j + 1];//min max
                }
                ccw[13 - 1] = cbf[9 - 1] = chid[8 - 1] = wist[7 - 1] = mimax[5 - 1] = data[p];//�Ō�ɕt��������
                p++;
                int compere = strncmp((const char*)ccw, (const char*)ColcW, 13);
                if (compere == 0) {
                    StrLen = 0;
                    while (data[p + StrLen] != '"')//customWidth �l
                        StrLen++;
                    free(cuw);
                    UINT32 strsize = (UINT32)StrLen + 1;
                    cuw = (UINT8*)malloc(strsize);
                    for (int i = 0; i < StrLen; i++) {
                        cuw[i] = data[p]; p++;
                    }
                    cuw[StrLen] = '\0';
                }
                compere = strncmp((const char*)cbf, (const char*)Colbf, 9);
                if (compere == 0) {//bestFit get
                    StrLen = 0;
                    while (data[p + StrLen] != '"')//bestFit �l
                        StrLen++;
                    free(bestf);
                    UINT32 strsize = (UINT32)StrLen + 1;
                    bestf = (UINT8*)malloc(strsize);
                    for (int i = 0; i < StrLen; i++) {
                        bestf[i] = data[p]; p++;
                    }
                    bestf[StrLen] = '\0';
                }
                compere = strncmp((const char*)chid, (const char*)Colhid, 8);
                if (compere == 0) {//hidden get
                    StrLen = 0;
                    while (data[p + StrLen] != '"')//hidden �l
                        StrLen++;
                    free(hidd);
                    UINT32 strsize = (UINT32)StrLen + 1;
                    hidd = (UINT8*)malloc(strsize);
                    for (int i = 0; i < StrLen; i++) {
                        hidd[i] = data[p]; p++;
                    }
                    hidd[StrLen] = '\0';
                }
                compere = strncmp((const char*)wist, (const char*)ColS, 7);
                if (compere == 0) {//style get
                    StrLen = 0;
                    while (data[p + StrLen] != '"')//hidden �l
                        StrLen++;
                    free(style);
                    UINT32 strsize = (UINT32)StrLen + 1;
                    style = (UINT8*)malloc(strsize);
                    for (int i = 0; i < StrLen; i++) {
                        style[i] = data[p]; p++;
                    }
                    style[StrLen] = '\0';
                }
                compere = strncmp((const char*)wist, (const char*)Colswidth, 7);
                if (compere == 0) {//width get
                    StrLen = 0;
                    while (data[p + StrLen] != '"')//hidden �l
                        StrLen++;
                    free(colw);
                    UINT32 strsize = (UINT32)StrLen + 1;
                    colw = (UINT8*)malloc(strsize);
                    for (int i = 0; i < StrLen; i++) {
                        colw[i] = data[p]; p++;
                    }
                    colw[StrLen] = '\0';
                }
                compere = strncmp((const char*)mimax, (const char*)Colmax, 5);
                if (compere == 0) {//width get
                    StrLen = 0;
                    while (data[p + StrLen] != '"')//hidden �l
                        StrLen++;
                    free(max);
                    UINT32 strsize = (UINT32)StrLen + 1;
                    max = (UINT8*)malloc(strsize);
                    for (int i = 0; i < StrLen; i++) {
                        max[i] = data[p]; p++;
                    }
                    max[StrLen] = '\0';
                }
                compere = strncmp((const char*)mimax, Colmin, 5);
                if (compere == 0) {//width get
                    StrLen = 0;
                    while (data[p + StrLen] != '"')//hidden �l
                        StrLen++;
                    free(min);
                    UINT32 strsize = (UINT32)StrLen + 1;
                    min = (UINT8*)malloc(strsize);
                    for (int i = 0; i < StrLen; i++) {
                        min[i] = data[p]; p++;
                    }
                    min[StrLen] = '\0';
                }
            }
            cls = addcolatyle(cls, min, max, colw, style, hidd, bestf, cuw);
            min = (UINT8*)malloc(1);
            min = nullptr;
            max = (UINT8*)malloc(1);
            max = nullptr;
            style = (UINT8*)malloc(1);
            style = nullptr;
            bestf = (UINT8*)malloc(1);
            bestf = nullptr;
            hidd = (UINT8*)malloc(1);
            hidd = nullptr;
            colw = (UINT8*)malloc(1);
            colw = nullptr;
            cuw = (UINT8*)malloc(1);
            cuw = nullptr;
            free(ccw); free(cbf); free(chid); free(wist); free(mimax);
        }
    }
}

cols* Ctags::addcolatyle(cols* cs, UINT8* min, UINT8* max, UINT8* W, UINT8* sty, UINT8* hid, UINT8* bF, UINT8* cuW) {
    if (!cs) {
        cs = coltalloc();
        cs->min = min;
        cs->max = max;
        cs->width = W;
        cs->style = sty;
        cs->hidden = hid;
        cs->bestffit = bF;
        cs->customwidth = cuW;
        cs->next = nullptr;
    }
    else
        cs->next = addcolatyle(cs->next, min, max, W, sty, hid, bF, cuW);

    return cs;
}

cols* Ctags::coltalloc() {
    return (cols*)malloc(sizeof(cols));
}

void Ctags::Ctableprint(C* c) {
    C* p = c;
    while (p) {
        if (p->val)
            std::cout << " val : " << p->val;
        if (p->t)
            std::cout << " t : " << p->t;
        if (p->s)
            std::cout << " s : " << p->s;
        if (p->col)
            std::cout << " col : " << p->col;
        if (p->f)
            std::cout << " f : " << p->f;
        std::cout << std::endl;
        p = p->next;
    }
}
//�Ō�̕�����R�s�[
void Ctags::getfinalstr() {
    UINT32 s = UINT32(dlen) - UINT32(p);
    UINT64 i = 0;
    const char *margeinfo="<mergeCells count=\"";//19char
    const char *marge="<mergeCell ref=\"";//16char
    UINT8* sm=(UINT8*)malloc(20);
    UINT8* Sm=(UINT8*)malloc(17);
    UINT32 fstrsize = s + 1;
    fstr = (UINT8*)malloc(fstrsize);
    int result=0;
    int mresult=0;
    
    while (p < dlen) {
        for (int j = 0; j < 19 - 1; j++) {
            sm[j] = sm[j + 1];
            if (j < (16 - 1))
                Sm[j] = Sm[j + 1];//marge count
        }
        Sm[16-1]=sm[19-1]=fstr[i] = data[p];
        p++; i++;
        
        result=strncmp((const char*)sm, margeinfo, 19);
        mresult=strncmp((const char*)Sm, marge, 16);
        if(result==0){//�}�[�W�@�Z�����擾
            UINT32 len=UINT32(i);
            while (data[p] != '"'){
                fstr[i] = data[p];
                p++; i++;
            }
            len=UINT32(i)-len;//�������擾
            MC=(UINT8*)malloc(len);
            for(UINT32 j=0;j<len;j++)
                MC[j]=data[p-len+j];
            MC[len]='\0';
        }
        if(mresult==0){
            UINT32 len=UINT32(i);
            while (data[p] != '"'){
                fstr[i] = data[p];
                p++; i++;
            }
            len=UINT32(i)-len;//�������擾
            MC=(UINT8*)malloc(len);
            for(UINT32 j=0;j<len;j++)
                MC[j]=data[p-len+j];
            MC[len]='\0';
        }
    }
    fstr[s] = '\0';
    free(sm);
    free(Sm);
    /* <mergeCells count="15">
    <mergeCell ref="E11:G11"/>
    <mergeCell ref="E12:G12"/>
    <mergeCell ref="O13:P13"/>
    <mergeCell ref="E15:M15"/>
    <mergeCell ref="B15:B16"/>
    <mergeCell ref="C15:C16"/>
    <mergeCell ref="D15:D16"/>
    <mergeCell ref="N15:N16"/>
    <mergeCell ref="O15:O16"/>
    <mergeCell ref="P15:P16"/>
    <mergeCell ref="B2:P2"/>
    <mergeCell ref="B3:P3"/>
    <mergeCell ref="B4:P4"/>
    <mergeCell ref="B5:P5"/>
    <mergeCell ref="B6:P6"/>
    </mergeCells> */
}

void Ctags::Ctablefree(C* c) {
    C* q;

    while (c) {
        q = c->next;  /* ���ւ̃|�C���^��ۑ� */
        free(c);
        c = q;
    }
}

void Ctags::Rowtablefree() {
    Row* q;

    while (rows) {
        Ctablefree(rows->cells);
        q = rows->next;  /* ���ւ̃|�C���^��ۑ� */
        free(rows);
        rows = q;
    }
}

void Ctags::selectfree() {
    selection* q;

    while (sct) {
        q = sct->next;  /* ���ւ̃|�C���^��ۑ� */
        free(sct);
        sct = q;
    }
}

void Ctags::colfree() {
    cols* q;

    while (cls) {
        q = cls->next;  /* ���ւ̃|�C���^��ۑ� */
        free(cls);
        cls = q;
    }
}

void Ctags::panefree() {
    Pane* q;

    while (Panes) {
        q = Panes->next;  /* ���ւ̃|�C���^��ۑ� */
        free(Panes);
        Panes = q;
    }
}