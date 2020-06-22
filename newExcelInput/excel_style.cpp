#include "excel_style.h"

styleread::styleread()
{
	readp = 0;
	fontcou = 0;
	fontCount = (UINT8*)malloc(1);
	fontCount = nullptr;
	fontRoot = nullptr;
}

styleread::~styleread()
{
}

Fonts* styleread::addfonts(Fonts* f, UINT8* sz, UINT8* co, UINT8* na, UINT8* fa, UINT8* cha, UINT8* sch) {
	if (!f) {
		f = (Fonts*)malloc(sizeof(Fonts));
		f->sz = sz;
		f->color = co;
		f->name = na;
		f->family = fa;
		f->charset = cha;
		f->scheme = sch;
		f->next = nullptr;
	}
	else {
		f->next = addfonts(f->next, sz, co, na, fa, cha, sch);
	}

	return f;
}

Fills* styleread::addfill(Fills* fi, FillPattern* p, fgcolor* f, bgcolor* b) {
	if (!fi) {
		fi = (Fills*)malloc(sizeof(Fills));
		fi->patten = p;
		fi->fg = f;
		fi->bg = b;
		fi->next = nullptr;
	}
	else {
		fi->next = addfill(fi->next, p, f, b);
	}

	return fi;
}

borders* styleread::addborder(borders* b, borderstyle* l,borderstyle* r,borderstyle* t,borderstyle* bo,borderstyle* d) {
	if (!b) {
		b = (borders*)malloc(sizeof(borders));
		b->left = l;
		b->right = r;
		b->top = t;
		b->bottom = bo;
		b->diagonal = d;
		b->next = nullptr;
	}
	else {
		b->next = addborder(b->next, l, r, t, bo, d);
	}

	return b;
}

stylexf* styleread::addstylexf(stylexf* sx, UINT8* n, UINT8* fo, UINT8* fi, UINT8* bi, UINT8* an, UINT8* ab, UINT8* aa, UINT8* ap, UINT8* av) {
	if (!sx) {
		sx = (stylexf*)malloc(sizeof(stylexf));
		sx->numFmtId = n;
		sx->fontId = fo;
		sx->fillId = fi;
		sx->borderId = bi;
		sx->applyNumberFormat = an;
		sx->applyBorder = ab;
		sx->applyAlignment = aa;
		sx->applyProtection = ap;
		sx->Avertical = av;//alignment vertical
		sx->next = nullptr;
	}
	else {
		sx->next = addstylexf(sx->next, n, fo, fi, bi, an, ab, aa, ap, av);
	}

	return sx;
}

cellxfs* styleread::addcellxfs(cellxfs* c, UINT8* n, UINT8* fo, UINT8* fi, UINT8* bi, UINT8* an, UINT8* ab, UINT8* aa, UINT8* av, UINT8* at,UINT8* ho,UINT8* afo,UINT8*afi,UINT8* xid) {
	if (!c) {
		c = (cellxfs*)malloc(sizeof(cellxfs));
		c->numFmtId = n;
		c->fontId = fo;
		c->fillId = fi;
		c->borderId = bi;
		c->applyNumberFormat = an;
		c->applyBorder = ab;
		c->applyAlignment = aa;
		c->Avertical = av;
		c->AwrapText = at;
		c->horizontal = ho;
		c->applyFont = afo;
		c->applyFill = afi;
		c->xfId = xid;
		c->next = nullptr;
	}
	else {
		c->next = addcellxfs(c->next, n, fo, fi, bi, an, ab, aa, av,at,ho,afo,afi,xid);
	}
	return c;
}

cellstyle* styleread::addcellstyle(cellstyle* c, UINT8* n, UINT8* x, UINT8* b, UINT8* cu, UINT8* xr) {
	if (!c) {
		c = (cellstyle*)malloc(sizeof(cellstyle));
		c->name = n;
		c->xfId = x;
		c->builtinId = b;
		c->customBuilt = cu;
		c->xruid = xr;
		c->next = nullptr;
	}
	else {
		c->next = addcellstyle(c->next, n, x, b, cu, xr);
	}

	return c;
}

Dxf* styleread::adddxf(Dxf* d, UINT8* c, UINT8* bc) {
	if (!d) {
		d = (Dxf*)malloc(sizeof(Dxf));
		d->color = c;
		d->bgcolor = bc;
		d->next = nullptr;
	}
	else {
		d->next = adddxf(d->next, c, bc);
	}

	return d;
}

colors* styleread::addcolor(colors* c, UINT8* co) {
	if (!c) {
		c = (colors*)malloc(sizeof(colors));
		c->color = co;
		c->next = nullptr;
	}
	else {
		c->next = addcolor(c->next, co);
	}
	return c;
}

numFMts* styleread::addnumFmts(numFMts* n, UINT8* i, UINT8* c) {
	if (!n) {
		n = (numFMts*)malloc(sizeof(numFMts));
		n->Id = i;
		n->Code = c;
		n->next = nullptr;
	}
	else {
		n->next = addnumFmts(n->next, i, c);
	}
	return n;
}

void styleread::readfontV(UINT8* dat) {
	const char sz[9] = "sz val=\"";//8
	const char colo[14] = "color theme=\"";//13
	const char name[11] = "name val=\"";//10
	const char fami[13] = "family val=\"";//12
	const char chars[14] = "charset val=\"";//13
	const char sche[13] = "scheme val=\"";//12

	const char Efont[] = "</font>";//7ji

	UINT8 endtag[8] = { 0 };//7文字
	UINT8 strsear[14] = { 0 };//13文字
	UINT8 seafam[13] = { 0 };//12文字
	UINT8 seanam[11] = { 0 };//10文字

	UINT8* SZ = (UINT8*)malloc(1); SZ = nullptr;
	UINT8* Name = (UINT8*)malloc(1); Name = nullptr;
	UINT8* Color = (UINT8*)malloc(1); Color = nullptr;
	UINT8* Family = (UINT8*)malloc(1); Color = nullptr;
	UINT8* Charset = (UINT8*)malloc(1); Charset = nullptr;
	UINT8* Scheme = (UINT8*)malloc(1); Scheme = nullptr;

	int eresult = 1;
	int txtresult = 0;

	// </font>まで読み込み
	while (eresult != 0) {
		//1つずつずらす
		for (int i = 0; i < 12; i++) {
			strsear[i] = strsear[i + 1];
			if (i < 6)
				endtag[i] = endtag[i + 1];
			if (i < 12 - 1)
				seafam[i] = seafam[i + 1];
			if (i < 10 - 1)
				seanam[i] = seanam[i + 1];
		}
		strsear[12] = endtag[6] = seafam[11] = seanam[9] = dat[readp];
		strsear[13] = endtag[7] = seafam[12] = seanam[10] = '\0';
		readp++;

		txtresult = strncmp((const char*)strsear, sche, 12);//font スタートタグ 検索

		txtresult = strncmp((const char*)strsear, sche, 12);//scheme 検索
		if (txtresult == 0) //scheme 読み込み
			Scheme = getValue(dat);

		txtresult = strncmp((const char*)strsear, chars, 13);//charset 検索
		if (txtresult == 0) //charset 読み込み
			Charset = getValue(dat);

		txtresult = strncmp((const char*)seafam, fami, 12);//family検索
		if (txtresult == 0) //family読み込み
			Family = getValue(dat);

		txtresult = strncmp((const char*)strsear, colo, 13);//color検索
		if (txtresult == 0) //color読み込み
			Color = getValue(dat);

		txtresult = strncmp((const char*)endtag, sz, 8);//sz検索
		if (txtresult == 0) //sz読み込み
			SZ = getValue(dat);

		txtresult = strncmp((const char*)seanam, name, 10);//name検索
		if (txtresult == 0) //name読み込み
			Name = getValue(dat);

		eresult = strcmp((const char*)endtag, Efont);
	}

	fontRoot = addfonts(fontRoot, SZ, Color, Name, Family, Charset, Scheme);//フォント　追加
}

UINT8* styleread::readfonts(UINT8* sd) {


	const char Sfont[] = "<font>";//6ji

	UINT8 startTag[7] = { 0 };//6文字

	UINT8 endtag[9] = { 0 };//8文字

	UINT8* SZ = (UINT8*)malloc(1); SZ = nullptr;
	UINT8* Name = (UINT8*)malloc(1); Name = nullptr;
	UINT8* Color = (UINT8*)malloc(1); Color = nullptr;
	UINT8* Family = (UINT8*)malloc(1); Color = nullptr;
	UINT8* Charset = (UINT8*)malloc(1); Charset = nullptr;
	UINT8* Scheme = (UINT8*)malloc(1); Scheme = nullptr;

	int eresult = 1;
	int txtresult = 0;
	while (eresult != 0) {
		//1つずつずらす
		for (int i = 0; i < 7; i++) {
			endtag[i] = endtag[i + 1];
			if (i < 6 - 1)
				startTag[i] = startTag[i + 1];
		}
		startTag[5] = endtag[7] = sd[readp];
		startTag[6] = endtag[8] = '\0';
		readp++;

		txtresult = strncmp((const char*)startTag, Sfont, 6);//font スタートタグ 検索
		if (txtresult == 0)
			readfontV(sd);
		
		eresult = strcmp((const char*)endtag, Efonts);
	}

	return nullptr;
}

// <fill>検索
void styleread::getfillV(UINT8* d) {
	const char type[14] = "patternType=\"";//13字
	const char Fgtag[9] = "<fgColor";//8ji
	const char Bgtag[9] = "<bgColor";//8ji
	const char EndFill[] = "</fill>";//7文字

	UINT8 fbtag[9] = { 0 };//8文字
	UINT8 endtag[8] = { 0 };//7文字
	UINT8 Pattern[14] = { 0 };//13文字

	UINT8* PaT = (UINT8*)malloc(1); PaT = nullptr;
	fgcolor* FG = (fgcolor*)malloc(sizeof(fgcolor)); FG = nullptr;
	bgcolor* BG = (bgcolor*)malloc(sizeof(bgcolor)); BG = nullptr;

	size_t msize = 0;//malloc用

	size_t vallen = 0;//tag 値
	size_t taglen = 0;//タグ　読み込み開始位置

	int eresult = 1;
	int txtresult = 0;

	while (eresult != 0) {
		//1つずつずらす
		for (int i = 0; i < 12; i++) {
			Pattern[i] = Pattern[i + 1];
			if (i < 7)
				fbtag[i] = fbtag[i + 1];
			if (i < 6)
				endtag[i] = endtag[i + 1];
		}
		Pattern[12] = fbtag[7]= endtag[6] = d[readp];
		Pattern[13] = fbtag[8]= endtag[7] = '\0';
		readp++;

		txtresult = strncmp((const char*)Pattern, type, 10);//pattern Type 検索
		if (txtresult == 0) //pattern Type 読み込み
			PaT = getValue(d);

		txtresult = strncmp((const char*)fbtag, Fgtag, 8);//pattern Type 検索
		if (txtresult == 0) {//pattern Type 読み込み
			//fgColor 検索
			free(FG);
			FG = readfillFg(d);
		}

		txtresult = strncmp((const char*)fbtag, Bgtag, 8);//pattern Type 検索
		if (txtresult == 0) {//pattern Type 読み込み
			//fgColor 検索
			free(BG);
			BG = readbgcolor(d);
		}

		eresult = strcmp((const char*)endtag, EndFill);
	}

	FillPattern* fi = (FillPattern*)malloc(sizeof(FillPattern));
	fi->patternType = PaT;

	fillroot = addfill(fillroot, fi, FG, BG);
}

// <fills>検索
UINT8* styleread::readfill(UINT8* sd) {

	const char Startfill[] = "<fill>";//6文字

	UINT8 endtag[9] = { 0 };//8文字
	UINT8 starttag[7] = { 0 };//6文字

	int eresult = 1;
	int txtresult = 0;

	while (eresult != 0) {
		//1つずつずらす
		for (int i = 0; i < 7; i++) {
			endtag[i] = endtag[i + 1];
			if (i < 5)
				starttag[i] = starttag[i + 1];
		}
		starttag[5] = endtag[7] = sd[readp];
		starttag[6] = endtag[8] = '\0';
		readp++;

		txtresult = strncmp((const char*)starttag, Startfill, 10);//pattern Type 検索
		if (txtresult == 0) //pattern Type 読み込み
			getfillV(sd);//スタートタグ　読み込み
		
		eresult = strcmp((const char*)endtag, endfil);
	}

	return nullptr;
}

fgcolor* styleread::readfillFg(UINT8* dat) {
	const char theme[] = "theme=\"";//7字
	const char fgtag[] = "tint=\"";//6ji

	UINT8 Theme[8] = { 0 };//7文字
	UINT8 Fgtag[7] = { 0 };//6文字

	UINT8* TH = (UINT8*)malloc(1); TH = nullptr;
	UINT8* SC = (UINT8*)malloc(1); SC = nullptr;

	int vallen = 0;
	size_t taglen = 0;
	size_t msize = 0;

	int result = 0;
	//fgColor 読み込み
	while (dat[readp] != '>') {

		for (int i = 0; i < 6; i++) {
			Theme[i] = Theme[i + 1];
			if (i < 5)
				Fgtag[i] = Fgtag[i + 1];
		}
		Theme[6] = Fgtag[5] = dat[readp];
		Theme[7] = Fgtag[6] = '\0';
		readp++;

		result = strncmp((const char*)Theme, theme, 7);//theme 検索
		if (result == 0)
			TH = getValue(dat);

		result = strncmp((const char*)Fgtag, fgtag, 7);//fgtag 検索
		if (result == 0)
			SC = getValue(dat);
	}

	fgcolor* FG = (fgcolor*)malloc(sizeof(fgcolor));
	FG->theme = TH;
	FG->tint = SC;

	return FG;
}

bgcolor* styleread::readbgcolor(UINT8* dat) {
	
	const char indexed[10] = "indexed=\"";//9ji

	UINT8 Index[10] = { 0 };//9文字

	UINT8* ID = (UINT8*)malloc(1); ID = nullptr;

	int vallen = 0;
	size_t taglen = 0;
	size_t msize = 0;

	int result = 0;
	//fgColor 読み込み
	while (dat[readp] != '>') {

		for (int i = 0; i < 8; i++) {
			Index[i] = Index[i + 1];
		}
		Index[8] = dat[readp];
		Index[9] = '\0';
		readp++;

		result = strncmp((const char*)Index, indexed, 7);//indexed 検索
		if (result == 0)
			ID = getValue(dat);//Index 値取得
	}

	bgcolor* BG = (bgcolor*)malloc(sizeof(bgcolor));
	BG->indexed = ID;

	return BG;
}
//border value 取得
borderstyle* styleread::getborV(UINT8* dat,UINT8* endT,size_t endTlen) {

	const char style[] = "style=\"";//7
	const char indexed[] = "indexed=\"";//9
	const char au[] = "auto=\"";//6
	const char theme[] = "theme=\"";//7
	const char tint[] = "tint=\"";//6
	const char rgb[] = "rgb=\"";//5


	UINT8 Style[8] = { 0 };
	UINT8 Index[10] = { 0 };
	UINT8 AUT[7] = { 0 };
	UINT8 RGB[6] = { 0 };

	UINT8* endtag = (UINT8*)malloc(endTlen+1);
	memset(endtag, 0, endTlen + 1);

	int result = 1;
	int txtresult = 0;

	int maxlen = (endTlen > 10) ? endTlen : 10;

	UINT8* ST = nullptr;
	UINT8* AT = nullptr;
	UINT8* IND = (UINT8*)malloc(1); IND = nullptr;
	UINT8* Theme = (UINT8*)malloc(1); Theme = nullptr;
	UINT8* Tin = (UINT8*)malloc(1); Tin = nullptr;
	UINT8* Rgb = (UINT8*)malloc(1); Rgb = nullptr;

	borderstyle* bv = nullptr;

	if (dat[readp+1] == '/') {
		//タグ値なし
		bv = (borderstyle*)malloc(sizeof(borderstyle));
		bv = nullptr;
	}
	else {
		while (result != 0) {

			for (int i = 0; i < maxlen - 1; i++) {
				if (i < 6 - 1)
					AUT[i] = AUT[i + 1];
				if (i < 5 - 1)
					RGB[i] = RGB[i + 1];
				if (i < 9 - 1)
					Index[i] = Index[i + 1];
				if (i < 7 - 1)
					Style[i] = Style[i + 1];
				if (i < endTlen - 1)
					endtag[i] = endtag[i + 1];
			}
			AUT[5] = RGB[4] = Index[8] = Style[6] = endtag[endTlen - 1] = dat[readp];
			AUT[6] = RGB[5] = Index[9] = Style[7] = endtag[endTlen] = '\0';
			readp++;

			txtresult = strncmp((const char*)AUT, tint, 5);//auto 検索
			if (txtresult == 0) //sz読み込み
				Tin = getValue(dat);

			txtresult = strncmp((const char*)RGB, rgb, 5);//auto 検索
			if (txtresult == 0) //sz読み込み
				Rgb = getValue(dat);

			txtresult = strncmp((const char*)AUT, au, 6);//auto 検索
			if (txtresult == 0) //sz読み込み
				AT = getValue(dat);

			txtresult = strncmp((const char*)Index, indexed, 9);//indexed 検索
			if (txtresult == 0) //sz読み込み
				IND = getValue(dat);

			txtresult = strncmp((const char*)Style, style, 7);//style 検索
			if (txtresult == 0) //sz読み込み
				ST = getValue(dat);

			result = strncmp((const char*)Index, indexed, endTlen);
		}

		bv = (borderstyle*)malloc(sizeof(borderstyle));
		bv->index = IND;
		bv->style = ST;
		bv->tint = Tin;
		bv->rgb = Rgb;
		bv->Auto = AT;
	}

	return bv;
}

//border 値取得
void styleread::getBorder(UINT8* d) {
	const char left[] = "<left";//5
	const char right[] = "<right";//6
	const char top[] = "<top";//4
	const char bottom[] = "<bottom";//7
	const char diag[] = "<diagonal";//9

	const char el[] = "</left>";//8
	const char er[] = "</right>";//8
	const char et[] = "</top>";//6
	const char eb[] = "</bottom>";//9
	const char ed[] = "</diagonal>";//11

	const char endtag[] = "</border>";//9

	UINT8 L[6] = { 0 };
	UINT8 R[7] = { 0 };
	UINT8 T[5] = { 0 };
	UINT8 B[8] = { 0 };
	UINT8 D[10] = { 0 };

	borderstyle* l = nullptr;
	borderstyle* r = nullptr;
	borderstyle* t = nullptr;
	borderstyle* b = nullptr;
	borderstyle* dia = nullptr;

	int eresult = 1;
	int txtresult = 0;

	while (eresult!=0)
	{
		for (int i = 0; i < 8; i++) {
			D[i] = D[i + 1];
			if (i < 6)
				B[i] = B[i + 1];
			if (i < 5)
				R[i] = R[i + 1];
			if (i < 4)
				L[i] = L[i + 1];
			if (i < 3)
				T[i] = T[i + 1];
		}
		L[4] = R[5] = T[3] = B[6] = D[8] = d[readp];
		L[5] = R[6] = T[4] = B[7] = D[9] = '\0';
		readp++;

		//全種　タグ　必ずある
		txtresult = strncmp((const char*)L, left, 5);//left検索
		if (txtresult == 0) 
			l = getborV(d,(UINT8*)el,8);

		txtresult = strncmp((const char*)R, right, 6);//right検索
		if (txtresult == 0) 
			r = getborV(d, (UINT8*)er, 8);

		txtresult = strncmp((const char*)T, top, 4);//top検索
		if (txtresult == 0)
			t = getborV(d, (UINT8*)et, 6);

		txtresult = strncmp((const char*)B, bottom, 7);//bottom検索
		if (txtresult == 0)
			b = getborV(d, (UINT8*)eb, 9);

		txtresult = strncmp((const char*)D, diag, 9);//diagonal検索
		if (txtresult == 0) {
			dia = getborV(d, (UINT8*)ed, 11);
			BorderRoot = addborder(BorderRoot, l, r, t, b, dia);
		}
		eresult = strcmp((const char*)D, endtag);
	}
}

//border 検索
UINT8* styleread::readboerder(UINT8* d) {
	const char startBor[] = "<border>";//8

	UINT8 starttag[9] = { 0 };//8文字
	UINT8 endtag[11] = { 0 };//10文字

	int eresult = 1;
	int txtresult = 0;
	while (eresult != 0) {
		//1つずつずらす
		for (int i = 0; i < 9; i++) {
			endtag[i] = endtag[i + 1];
			if(i<7)
				starttag[i] = starttag[i + 1];
		}
		starttag[7]= endtag[9] = d[readp];
		starttag[8]= endtag[10] = '\0';
		readp++;

		txtresult = strncmp((const char*)starttag, startBor, 8);//sz検索
		if (txtresult == 0) //sz読み込み
			getBorder(d);//ボーダー読み込み
		
		eresult = strcmp((const char*)endtag, enbor);
	}

	return nullptr;
}

//tag""値取得
UINT8* styleread::getValue(UINT8* d) {
	int vallen = 0;
	size_t taglen = size_t(readp);
	size_t msize = 0;
	UINT8* tagval = (UINT8*)malloc(1);
	tagval = nullptr;

	while (d[readp] != '"') {
		if (d[readp] != '"' && d[readp] != '/')//read data
			vallen++;
		readp++;
	}
	msize = vallen + 1;
	free(tagval);
	tagval = (UINT8*)malloc(msize);
	for (size_t i = 0; i < vallen; i++)
		tagval[i] = d[taglen + i];
	tagval[vallen] = '\0';

	return tagval;
}

//CellStyleXfs読み込み
void styleread::getCellStyleXfsV(UINT8* d) {
	const char numFmtId[] = "numFmtId=\"";//10文字
	const char fontId[] = "fontId=\"";//8文字
	const char fillId[] = "fillId=\"";//8文字
	const char borderId[] = "borderId=\"";//10文字
	const char vertical[] = "vertical=\"";//10文字
	const char applyNumberFormat[] = "applyNumberFormat=\"";//19文字
	const char applyBorder[] = "applyBorder=\"";//13文字
	const char applyAlignment[] = "applyAlignment=\"";//16文字
	const char applyProtection[] = "applyProtection=\"";//17文字

	const char endtag[] = "</xf>";//5文字

	UINT8 nFI[11] = { 0 };//10文字
	UINT8 Fid[9] = { 0 };//8文字
	UINT8 borid[11] = { 0 };//10文字
	UINT8 Et[6] = { 0 };//5文字
	UINT8 anum[20] = { 0 };//19文字
	UINT8 abor[14] = { 0 };//13文字
	UINT8 aali[17] = { 0 };//16文字
	UINT8 apro[18] = { 0 };//17文字

	UINT8* num = (UINT8*)malloc(1); num= nullptr;
	UINT8* fon = (UINT8*)malloc(1); fon = nullptr;
	UINT8* fil = (UINT8*)malloc(1); fil = nullptr;
	UINT8* bor = (UINT8*)malloc(1); bor = nullptr;
	UINT8* verti = (UINT8*)malloc(1); verti = nullptr;
	UINT8* Anu = (UINT8*)malloc(1); Anu = nullptr;
	UINT8* Abo = (UINT8*)malloc(1); Abo = nullptr;
	UINT8* Aal = (UINT8*)malloc(1); Aal = nullptr;
	UINT8* Apr = (UINT8*)malloc(1); Apr = nullptr;

	int result = 1;
	int sresult = 0;

	while (result != 0) {
		for (int i = 0; i < 19 - 1; i++) {
			anum[i] = anum[i + 1];
			if (i < 13 - 1)
				abor[i] = abor[i + 1];
			if (i < 16 - 1)
				aali[i] = aali[i + 1];
			if (i < 17 - 1)
				apro[i] = apro[i + 1];
			if (i < 10 - 1)
				borid[i] = borid[i + 1];
			if (i < 8 - 1)
				Fid[i] = Fid[i + 1];
			if (i < 5 - 1)
				Et[i] = Et[i + 1];
		}
		apro[16] = aali[15] = abor[12] = borid[9] = Fid[7] = Et[4] = d[readp];
		apro[17] = aali[16] = abor[13] = borid[10] = Fid[8] = Et[5] = '\0';
		readp++;

		//終了タグ　なし　/>終了
		if (d[readp - 1] == '/')
			if (d[readp] == '>')
				break;
		sresult = strncmp((const char*)apro, applyProtection, 17);
		if (sresult == 0) {
			//numFmtId　値読み込み
			free(Apr);
			Apr = getValue(d);
		}
		sresult = strncmp((const char*)aali, applyAlignment, 16);
		if (sresult == 0) {
			//numFmtId　値読み込み
			free(Aal);
			Aal = getValue(d);
		}
		sresult = strncmp((const char*)abor, applyBorder, 13);
		if (sresult == 0) {
			//numFmtId　値読み込み
			free(Abo);
			Abo = getValue(d);
		}
		sresult = strncmp((const char*)anum, applyNumberFormat, 19);
		if (sresult == 0) {
			//numFmtId　値読み込み
			free(Anu);
			Anu = getValue(d);
		}
		sresult = strncmp((const char*)borid, numFmtId, 10);
		if (sresult == 0) {
			//numFmtId　値読み込み
			free(num);
			num = getValue(d);
		}
		sresult = strncmp((const char*)borid, borderId, 10);
		if (sresult == 0) {
			//borderId　値読み込み
			free(bor);
			bor = getValue(d);
		}
		sresult = strncmp((const char*)borid, vertical, 10);
		if (sresult == 0) {
			//vertical　値読み込み
			free(verti);
			verti = getValue(d);
		}
		sresult = strncmp((const char*)Fid, fontId, 8);
		if (sresult == 0) {
			//fontId　値読み込み
			free(fon);
			fon = getValue(d);
		}
		sresult = strncmp((const char*)Fid, fillId, 8);
		if (sresult == 0) {
			//fillId　値読み込み
			free(fil);
			fil = getValue(d);
		}
		result = strncmp((const char*)Et, endtag, 5);
	}

	cellstyleXfsRoot = addstylexf(cellstyleXfsRoot, num, fon, fil, bor, Anu, Abo, Aal, Apr, verti);
}

//CellStyleXfs読み込み
void styleread::getCellXfsV(UINT8* d) {
	const char numFmtId[] = "numFmtId=\"";//10文字
	const char fontId[] = "fontId=\"";//8文字
	const char fillId[] = "fillId=\"";//8文字
	const char borderId[] = "borderId=\"";//10文字
	const char vertical[] = "vertical=\"";//10文字
	const char applyNumberFormat[] = "applyNumberFormat=\"";//19文字
	const char applyBorder[] = "applyBorder=\"";//13文字
	const char applyAlignment[] = "applyAlignment=\"";//16文字
	const char horizontal[] = "horizontal=\"";//12文字
	const char wrapText[] = "wrapText=\"";//10文字

	const char applyFont[] = "applyFont=\"";//11文字
	const char applyFill[] = "applyFill=\"";//11文字
	const char xfId[] = "xfId=\"";//6文字

	const char endtag[] = "</xf>";//5文字


	UINT8 aF[12] = { 0 };//11文字
	UINT8 Fid[9] = { 0 };//8文字
	UINT8 borid[11] = { 0 };//10文字
	UINT8 Et[6] = { 0 };//5文字
	UINT8 anum[20] = { 0 };//19文字
	UINT8 abor[14] = { 0 };//13文字
	UINT8 aali[17] = { 0 };//16文字
	UINT8 apro[18] = { 0 };//17文字
	UINT8 hori[18] = { 0 };//17文字
	UINT8 xfid[7] = { 0 };//6文字

	UINT8* num = (UINT8*)malloc(1); num = nullptr;
	UINT8* fon = (UINT8*)malloc(1); fon = nullptr;
	UINT8* fil = (UINT8*)malloc(1); fil = nullptr;
	UINT8* bor = (UINT8*)malloc(1); bor = nullptr;
	UINT8* verti = (UINT8*)malloc(1); verti = nullptr;
	UINT8* Anu = (UINT8*)malloc(1); Anu = nullptr;
	UINT8* Abo = (UINT8*)malloc(1); Abo = nullptr;
	UINT8* Aal = (UINT8*)malloc(1); Aal = nullptr;
	UINT8* Apr = (UINT8*)malloc(1); Apr = nullptr;
	UINT8* Ho = (UINT8*)malloc(1); Ho = nullptr;
	UINT8* Wt = (UINT8*)malloc(1); Wt = nullptr;
	UINT8* aFo = (UINT8*)malloc(1); aFo = nullptr;
	UINT8* aFi = (UINT8*)malloc(1); aFi = nullptr;
	UINT8* xId = (UINT8*)malloc(1); xId = nullptr;

	int result = 1;
	int sresult = 0;

	while (result != 0) {
		for (int i = 0; i < 19 - 1; i++) {
			anum[i] = anum[i + 1];
			if (i < 13 - 1)
				abor[i] = abor[i + 1];
			if (i < 16 - 1)
				aali[i] = aali[i + 1];
			if (i < 17 - 1)
				apro[i] = apro[i + 1];
			if (i < 12 - 1)
				hori[i] = hori[i + 1];
			if (i < 11 - 1)
				aF[i] = aF[i + 1];
			if (i < 10 - 1)
				borid[i] = borid[i + 1];
			if (i < 8 - 1)
				Fid[i] = Fid[i + 1];
			if (i < 6 - 1)
				xfid[i] = xfid[i + 1];
			if (i < 5 - 1)
				Et[i] = Et[i + 1];
		}
		xfid[5]=aF[10]=hori[11]= apro [16]=aali[15]=abor[12]=borid[9] = Fid[7] = Et[4] = d[readp];
		xfid[6]=aF[11]=hori[12] = apro[17] = aali[16] = abor[13]=borid[10] = Fid[8] = Et[5] = '\0';
		readp++;

		//終了タグ　なし　/>終了
		/*if (d[readp - 1] == '/')
			if (d[readp] == '>')
				break;*/
		sresult = strncmp((const char*)xfid, xfId, 10);
		if (sresult == 0) {
			//numFmtId　値読み込み
			free(xId);
			xId = getValue(d);
		}
		sresult = strncmp((const char*)borid, wrapText, 10);
		if (sresult == 0) {
			//numFmtId　値読み込み
			free(Wt);
			Wt = getValue(d);
		}
		sresult = strncmp((const char*)hori, horizontal, 12);
		if (sresult == 0) {
			//numFmtId　値読み込み
			free(Ho);
			Ho = getValue(d);
		}
		sresult = strncmp((const char*)aF, applyFont, 11);
		if (sresult == 0) {
			//numFmtId　値読み込み
			free(aFo);
			aFo = getValue(d);
		}
		sresult = strncmp((const char*)aF, applyFill, 11);
		if (sresult == 0) {
			//numFmtId　値読み込み
			free(aFi);
			aFi = getValue(d);
		}
		sresult = strncmp((const char*)aali, applyAlignment, 16);
		if (sresult == 0) {
			//numFmtId　値読み込み
			free(Aal);
			Aal = getValue(d);
		}
		sresult = strncmp((const char*)abor, applyBorder, 13);
		if (sresult == 0) {
			//numFmtId　値読み込み
			free(Abo);
			Abo = getValue(d);
		}
		sresult = strncmp((const char*)anum, applyNumberFormat, 19);
		if (sresult == 0) {
			//numFmtId　値読み込み
			free(Anu);
			Anu = getValue(d);
		}
		sresult = strncmp((const char*)borid, numFmtId, 10);
		if (sresult == 0) {
			//numFmtId　値読み込み
			free(num);
			num = getValue(d);
		}
		sresult = strncmp((const char*)borid, borderId, 10);
		if (sresult == 0) {
			//borderId　値読み込み
			free(bor);
			bor = getValue(d);
		}
		sresult = strncmp((const char*)borid, vertical, 10);
		if (sresult == 0) {
			//vertical　値読み込み
			free(verti);
			verti = getValue(d);
		}
		sresult = strncmp((const char*)Fid, fontId, 8);
		if (sresult == 0) {
			//fontId　値読み込み
			free(fon);
			fon = getValue(d);
		}
		sresult = strncmp((const char*)Fid, fillId, 8);
		if (sresult == 0) {
			//fillId　値読み込み
			free(fil);
			fil = getValue(d);
		}
		result = strncmp((const char*)Et, endtag, 5);
	}

	cellXfsRoot = addcellxfs(cellXfsRoot, num, fon, fil, bor, Anu, Abo, Aal, verti,Wt,Ho,aFo,aFi,xId);
}

//cellStyleXfs 読み取り
void styleread::readCellStyleXfs(UINT8* d) {
	const char starttag[] = "<xf";//3文字

	UINT8 stag[4] = { 0 };//3文字
	UINT8 endtag[16] = { 0 };//15文字

	int result = 1;
	int sresult = 0;

	while (result != 0) {
		for (int i = 0; i < 15 - 1; i++) {
			endtag[i] = endtag[i + 1];
			if(i<2)
				stag[i] = stag[i + 1];
		}
		endtag[14]=stag[2] = d[readp];
		endtag[15]=stag[3] = '\0';
		readp++;

		sresult = strncmp((const char*)stag, starttag, 3);
		if (result == 0) {
			//スタートタグ　値読み込み
			getCellStyleXfsV(d);
		}

		result = strncmp((const char*)endtag, EncsXfs, 15);
	}
}

//cellXfs 読み取り
void styleread::readCellXfs(UINT8* d) {
	const char starttag[] = "<xf";//3文字

	UINT8 stag[4] = { 0 };//3文字
	UINT8 endtag[14] = { 0 };//13文字

	int result = 1;
	int sresult = 0;

	while (result != 0) {
		for (int i = 0; i < 13 - 1; i++) {
			endtag[i] = endtag[i + 1];
			if (i < 2)
				stag[i] = stag[i + 1];
		}
		endtag[12] = stag[2] = d[readp];
		endtag[13] = stag[3] = '\0';
		readp++;

		sresult = strncmp((const char*)stag, starttag, 3);
		if (result == 0) {
			//スタートタグ　値読み込み
			getCellStyleXfsV(d);
		}

		result = strncmp((const char*)endtag, Ecellsty, 13);
	}
}

void styleread::readstyle(UINT8* sdata, UINT64 sdatalen)
{
	const char count[] = "count=\"";//count検索

	UINT8 sEs[14] = { 0 };//13文字
	UINT8 sExfs[16] = { 0 };//15文字
	UINT8 sfon[7] = { 0 };//6文字
	UINT8 sdxf[6] = { 0 };//5文字
	UINT8 sbor[9] = { 0 };//8文字
	UINT8 sstyle[12] = { 0 };//11文字

	UINT8 Cou[8] = { 0 };

	int result = 1;
	int otherresult = 0;

	int vallen = 0;
	int taglen = 0;

	size_t msize = 0;

	//文字検索
	while (result != 0) {
		//1つずつずらす
		for (int i = 0; i < 15-1; i++) {
			sExfs[i] = sExfs[i + 1];
			if (i < 13-1)
				sEs[i] = sEs[i + 1];//エンドタグ
			if (i < 6 - 1)
				sfon[i] = sfon[i + 1];
			if (i < 8 - 1)
				sbor[i] = sbor[i + 1];
			if (i < 11 - 1)
				sstyle[i] = sstyle[i + 1];
			if (i < 5 - 1)
				sdxf[i] = sdxf[i + 1];
		}
		sEs[12] = sExfs[14]= sfon[5]=sbor[7]=sstyle[10]=sdxf[4] = sdata[readp];
		sEs[13] = sExfs[15] = sfon[6] = sbor[8] = sstyle[11] = sdxf[5] = '\0';
		readp++;

		otherresult = strncmp((const char*)sfon, fonts, 6);
		if (otherresult == 0) {
			//フォント一致 read font
			while (sdata[readp] != '>') {
				for (int i = 0; i < 7 - 1; i++) {
					Cou[i] = Cou[i + 1];
				}
				Cou[6] = sdata[readp];
				Cou[7] = '\0';
				readp++;

				result = strncmp((const char*)Cou, count, 7);
				if (result == 0)
					fontCount = getValue(sdata);//count 取得
			}
			readfonts(sdata);//フォントタグ取得
		}

		otherresult = strncmp((const char*)sfon, fills, 6);
		if (otherresult == 0) {
			//フィル一致 read font
			while (sdata[readp] != '>') {
				for (int i = 0; i < 7 - 1; i++) {
					Cou[i] = Cou[i + 1];
				}
				Cou[6] = sdata[readp];
				Cou[7] = '\0';
				readp++;

				result = strncmp((const char*)Cou, count, 7);
				if (result == 0)
					fillCount = getValue(sdata);//count 取得
			}
			readfill(sdata);
		}

		otherresult = strncmp((const char*)sbor, border, 8);
		if (otherresult == 0) {
			//ボーダー一致 read border
			while (sdata[readp] != '>') {
				for (int i = 0; i < 7 - 1; i++) {
					Cou[i] = Cou[i + 1];
				}
				Cou[6] = sdata[readp];
				Cou[7] = '\0';
				readp++;

				result = strncmp((const char*)Cou, count, 7);
				if (result == 0)
					borderCount = getValue(sdata);//count 取得
			}
			readboerder(sdata);
		}

		otherresult = strncmp((const char*)sEs, cellstXfs, 13);
		if (otherresult == 0) {
			//セルスタイル一致 read cellstyleXfs
			while (sdata[readp] != '>') {
				for (int i = 0; i < 7 - 1; i++) {
					Cou[i] = Cou[i + 1];
				}
				Cou[6] = sdata[readp];
				Cou[7] = '\0';
				readp++;

				result = strncmp((const char*)Cou, count, 7);
				if (result == 0)
					cellStyleXfsCount = getValue(sdata);//count 取得
			}
			readCellStyleXfs(sdata);
		}

		otherresult = strncmp((const char*)sbor, Xfs, 8);
		if (otherresult == 0) {
			//cellXfs一致 read cellxfs
			while (sdata[readp] != '>') {
				for (int i = 0; i < 7 - 1; i++) {
					Cou[i] = Cou[i + 1];
				}
				Cou[6] = sdata[readp];
				Cou[7] = '\0';
				readp++;

				result = strncmp((const char*)Cou, count, 7);
				if (result == 0)
					cellXfsCount = getValue(sdata);//count 取得
			}
		}

		otherresult = strncmp((const char*)sstyle, CellStyl, 11);
		if (otherresult == 0) {
			//cellStylr一致 read cellstyle

		}

		otherresult = strncmp((const char*)sdxf, dxf, 5);
		if (otherresult == 0) {
			//dxfs一致 read dxfs

		}

		otherresult = strncmp((const char*)sExfs, color, 8);
		if (otherresult == 0) {
			//color一致 read colors

		}

		result = strcmp(EncsXfs, (const char*)sEs);
	}

}
