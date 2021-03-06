#include "ShipDataRead.h"

shipinfo::shipinfo(Row* cel) {
    cells = cel;
    searchtag = MargeaSearch();
}

shipinfo::~shipinfo() {
    //freeItem(its);
    //free(shipday);//元で解放 no malloc
    //free(day);
    //セル引数　元で解放 cel no malloc
}

void shipinfo::shipdataRead() {
    Row* r = cells;
    UINT32 shipR = 7;
    UINT32 shipC = 'C';
    bool shipFlag = false;
    int i = 0;

    while (r) {//C7 cell get
        if (r->r == shipR) {
            while (r->cells) {
                if (r->cells->col == shipC) {
                    if (r->cells->si) {
                        shipday = r->cells->si;
                        shipFlag = true;
                    }
                    else {
                        std::cout << "this cell not have shipdata" << std::endl;
                    }
                }
                r->cells = r->cells->next;
            }
        }
        r = r->next;
    }
    if (shipFlag) {
        day = (UINT8*)malloc(7);//日付だけ
        if (day) {
            day[0] = 'P'; day[1] = 'L';

            while (shipday[i] != '-' && shipday[i] != '\0' && shipday[i] != ' ') {
                if (i > 3) {
                    day[i - 2] = shipday[i];
                }
                i++;
            }
            day[6] = '\0';
            daylen = 6;
        }
    }
}

//指定の文字までの文字列　削る
UINT8* shipinfo::outuntil(UINT8 c, UINT8* st)
{
    int j = 0;
    int i = 0;
    UINT8* p = nullptr;

    while (st[j] != c && st[j] != '\0')
        j++;
    
    const UINT32 msize = (UINT32)j + 1;
    p = (UINT8*)malloc(msize);

    if (j == 0) {//最初に　(　があったら　)が読み始め
        while (st[j] != ')')
            j++;
        int startpos = 0;
        while (st[j] != '\0') {
            if (!isspace(st[j])) {
                p[startpos] = st[j];
                startpos++;
            }
            j++;
        }
        p[startpos] = '\0';
    }
    else {
        int t = 0; int h = 0;
        

        if (p) {
            while (h < j) {
                if (!isspace(st[h])) {
                    p[t] = st[h];
                    t++;
                }
                h++;
            }
            p[t] = '\0';
        }
        else {
            p = nullptr;
        }
    }

    return p;
}

Items* shipinfo::itemtalloc() {
    return (Items*)malloc(sizeof(Items));
}

Items* shipinfo::additem(Items* r, UINT8* num, UINT8* co, UINT8* ni, UINT8* te, UINT8* el, UINT8* tw, UINT8* th, UINT8* fo, UINT8* fi, UINT8* si, UINT8* f)
{
    if (!r) {
        r = itemtalloc();
        r->itn = num;
        r->col = co;
        r->s90 = ni;
        r->s100 = te;
        r->s110 = el;
        r->s120 = tw;
        r->s130 = th;
        r->s140 = fo;
        r->s150 = fi;
        r->s160 = si;
        r->sF = f;
        r->next = nullptr;
    }
    else {
        r->next = additem(r->next, num, co, ni, te, el, tw, th, fo, fi, si, f);
    }
    return r;
}

bool shipinfo::searchitemCell() {
    Row* sr = cells;

    while (sr) {
        while (sr->cells) {
            if (sr->cells->si) {
                if (strcmp((const char*)sr->cells->si, IN[0]) == 0) {
                    ITcells[0] = sr->cells->col; ITEMrow = sr->r;
                }
                else if (strcmp((const char*)sr->cells->si, IN[1]) == 0) {
                    ITcells[1] = sr->cells->col;
                }
                else if (strcmp((const char*)sr->cells->si, IN[2]) == 0) {
                    ITcells[2] = sr->cells->col;
                }
                else if (strcmp((const char*)sr->cells->si, IN[3]) == 0) {
                    ITcells[3] = sr->cells->col;
                }
                else if (strcmp((const char*)sr->cells->si, IN[4]) == 0) {
                    ITcells[4] = sr->cells->col;
                }
                else if (strcmp((const char*)sr->cells->si, IN[5]) == 0) {
                    ITcells[5] = sr->cells->col;
                }
                else if (strcmp((const char*)sr->cells->si, IN[6]) == 0) {
                    ITcells[6] = sr->cells->col;
                }
                else if (strcmp((const char*)sr->cells->si, IN[7]) == 0) {
                    ITcells[7] = sr->cells->col;
                }
                else if (strcmp((const char*)sr->cells->si, IN[8]) == 0) {
                    ITcells[8] = sr->cells->col;
                }
                else if (strcmp((const char*)sr->cells->si, IN[9]) == 0) {
                    ITcells[9] = sr->cells->col;
                }
                else if (strcmp((const char*)sr->cells->si, IN[10]) == 0) {
                    ITcells[10] = sr->cells->col;
                    return 1;
                }
            }
            sr->cells = sr->cells->next;
        }
        sr = sr->next;
    }
    return 0;
}

void shipinfo::GetItems() {
    Row* sr = cells;
    bool itemnumFlag = false;

    UINT8* IN = (UINT8*)malloc(1); IN = nullptr;
    UINT8* Colo = (UINT8*)malloc(1); Colo = nullptr;
    UINT8* nine = (UINT8*)malloc(1); nine = nullptr;
    UINT8* ten = (UINT8*)malloc(1); ten = nullptr;
    UINT8* ele = (UINT8*)malloc(1); ele = nullptr;
    UINT8* twe = (UINT8*)malloc(1); twe = nullptr;
    UINT8* thr = (UINT8*)malloc(1); thr = nullptr;
    UINT8* four = (UINT8*)malloc(1); four = nullptr;
    UINT8* fif = (UINT8*)malloc(1); fif = nullptr;
    UINT8* six = (UINT8*)malloc(1); six = nullptr;
    UINT8* f = (UINT8*)malloc(1); f = nullptr;

    shipdataRead();//C 7 セル読み取り

    searchitemCell();

    while (sr) {
        if (sr->r == ITEMrow)
            break;//スタートまでとばす
        sr = sr->next;
    }

    std::cout << "item cell column : " << ITEMrow << std::endl;
    int j = 0;
    while (sr) {
        while (sr->cells) {
            //std::cout << "item cell Si : " << sr->cells->si << std::endl;
            if (sr->cells->col == ITcells[0] && ITcells[0] != 0) {//numbe
                j = 0;
                if (sr->cells->si) {
                    //std::cout << "item cell Si : " << sr->cells->si << std::endl;
                    //while (sr->cells->si[j] != '\0')
                    //    j++;

                    //IN = (UINT8*)malloc(j);
                    IN = outuntil('(', sr->cells->si);
                    itemnumFlag = true;
                }
            }
            if (sr->cells->col == ITcells[1] && itemnumFlag && ITcells[1] != 0) {//color
                j = 0;
                if (sr->cells->si) {
                   // while (sr->cells->si[j] != '\0')
                    //    j++;
                    //Colo = (UINT8*)malloc(j);
                    Colo = outuntil('(', sr->cells->si);
                }
                else if (sr->cells->val) {
                    Colo = sr->cells->val;
                }
            }
            if (sr->cells->col == ITcells[2] && itemnumFlag && ITcells[2] != 0) {
                j = 0;
                if (sr->cells->si) {
                    nine = sr->cells->si;
                }
                else if (sr->cells->val) {
                    nine = sr->cells->val;
                }
            }
            if (sr->cells->col == ITcells[3] && itemnumFlag && ITcells[3] != 0) {
                j = 0;
                if (sr->cells->si) {
                    ten = sr->cells->si;
                }
                else if (sr->cells->val) {
                    ten = sr->cells->val;
                }
            }
            if (sr->cells->col == ITcells[4] && itemnumFlag && ITcells[4] != 0) {
                j = 0;
                if (sr->cells->si) {
                    ele = sr->cells->si;
                }
                else if (sr->cells->val) {
                    ele = sr->cells->val;
                }
            }
            if (sr->cells->col == ITcells[5] && itemnumFlag && ITcells[5] != 0) {
                j = 0;
                if (sr->cells->si) {
                    twe = sr->cells->si;
                }
                else if (sr->cells->val) {
                    twe = sr->cells->val;
                }
            }
            if (sr->cells->col == ITcells[6] && itemnumFlag && ITcells[6] != 0) {
                j = 0;
                if (sr->cells->si) {
                    thr = sr->cells->si;
                }
                else if (sr->cells->val) {
                    thr = sr->cells->val;
                }
            }
            if (sr->cells->col == ITcells[7] && itemnumFlag && ITcells[7] != 0) {
                j = 0;
                if (sr->cells->si) {
                    four = sr->cells->si;
                }
                else if (sr->cells->val) {
                    four = sr->cells->val;
                }
            }
            if (sr->cells->col == ITcells[8] && itemnumFlag && ITcells[8] != 0) {
                j = 0;
                if (sr->cells->si) {
                    fif = sr->cells->si;
                }
                else if (sr->cells->val) {
                    fif = sr->cells->val;
                }
            }
            if (sr->cells->col == ITcells[9] && itemnumFlag && ITcells[9] != 0) {
                j = 0;
                if (sr->cells->si) {
                    six = sr->cells->si;
                }
                else if (sr->cells->val) {
                    six = sr->cells->val;
                }
            }
            if (sr->cells->col == ITcells[10] && itemnumFlag && ITcells[10] != 0) {
                j = 0;
                if (sr->cells->si) {
                    f = sr->cells->si;
                }
                else if (sr->cells->val) {
                    f = sr->cells->val;;
                }

                itemnumFlag = false;

                its = additem(its, IN, Colo, nine, ten, ele, twe, thr, four, fif, six, f);


                IN = nullptr; Colo = nullptr; nine = nullptr; ten = nullptr;; ele = nullptr;
                twe = nullptr; thr = nullptr; four = nullptr; fif = nullptr; six = nullptr; f = nullptr;
            }
            sr->cells = sr->cells->next;
        }
        sr = sr->next;
    }
}

UINT8* shipinfo::StrInit() {
    UINT8* p = (UINT8*)malloc(1);
    p = nullptr;
    return p;
}

void shipinfo::freeItem(Items* t) {
    struct Items* q;
    while (t != NULL) {
        q = t->next;  /* 次へのポインタを保存 */
        free(t);
        t = q;
    }
}