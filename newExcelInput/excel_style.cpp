#include "excel_style.h"
#include <iostream>

styleread::styleread()
{
	readp = 0;
	fontcou = 0;

	fontCount = nullptr;
	fillCount = nullptr;
	borderCount = nullptr;
	cellStyleXfsCount = nullptr;
	cellXfsCount = nullptr;
	numFmtsCount = nullptr;
	cellstyleCount = nullptr;
	dxfsCount = nullptr;

	fontRoot = nullptr;
	fillroot = nullptr;
	BorderRoot = nullptr;
	cellstyleXfsRoot = nullptr;
	cellXfsRoot = nullptr;
	numFmtsRoot = nullptr;
	cellStyleRoot = nullptr;
	dxfRoot = nullptr;
	colorsRoot = nullptr;

	styleSheetStr = nullptr;
	extLstStr = nullptr;
	kFonts = nullptr;

	fonts = "<fonts";
	Efonts = "</fonts>";
	fills = "<fills";
	endfil = "</fills>";
	border = "<borders";
	enbor = "</borders>";
	cellstXfs = "<cellStyleXfs";
	EncsXfs = "</cellStyleXfs>";
	color = "<colors>";
	Ecolor = "</colors>";
	Xfs = "<cellXfs";
	Exfs = "</cellXfs>";
	CellStyl = "<cellStyles";
	Ecellsty = "</cellStyles>";
	dxf = "<dxfs";
	Edxf = "</dxfs>";
	Endstyle = "</styleSheet>";
	Fmts = "<numFmts";
	Efmts = "</numFmts>";
}

styleread::~styleread()
{
	/*
	for(int i=0;i<fontNum;i++)
		freefonts(fontRoot[i]);
	for (int i = 0; i < fillNum; i++)
		freefill(fillroot[i]);
	
	freeborder(BorderRoot);
	freestylexf(cellstyleXfsRoot);
	freecellxfs(cellXfsRoot);
	freenumFMts(numFmtsRoot);
	freecellstyle(cellStyleRoot);
	freeDxf(dxfRoot);
	*/
	free(styleSheetStr);
	free(extLstStr);
}

void styleread::freefonts(Fonts* f) {
	Fonts* q;

	while (f) {
		q = f->next;  /* ���ւ̃|�C���^��ۑ� */
		free(f);
		f = q;
	}
}

void styleread::freefill(Fills* f) {
	Fills* q;

	while (f) {
		q = f->next;  /* ���ւ̃|�C���^��ۑ� */
		free(f);
		f = q;
	}
}

void styleread::freeborder(borders* f) {
	borders* q;

	while (f) {
		q = f->next;  /* ���ւ̃|�C���^��ۑ� */
		free(f);
		f = q;
	}
}

void styleread::freestylexf(stylexf* f) {
	stylexf* q;

	while (f) {
		q = f->next;  /* ���ւ̃|�C���^��ۑ� */
		free(f);
		f = q;
	}
}

void styleread::freecellxfs(cellxfs* f) {
	cellxfs* q;

	while (f) {
		q = f->next;  /* ���ւ̃|�C���^��ۑ� */
		free(f);
		f = q;
	}
}

void styleread::freenumFMts(numFMts* f) {
	numFMts* q;

	while (f) {
		q = f->next;  /* ���ւ̃|�C���^��ۑ� */
		free(f);
		f = q;
	}
}

void styleread::freecellstyle(cellstyle* f) {
	cellstyle* q;

	while (f) {
		q = f->next;  /* ���ւ̃|�C���^��ۑ� */
		free(f);
		f = q;
	}
}

void styleread::freeDxf(Dxf* f) {
	Dxf* q;

	while (f) {
		q = f->next;  /* ���ւ̃|�C���^��ۑ� */
		free(f);
		f = q;
	}
}

Fonts* styleread::addfonts(Fonts* f, UINT8* b, UINT8* sz, UINT8* co, UINT8* na, UINT8* fa, UINT8* cha, UINT8* sch, UINT8* rg, UINT8* ind) {
	if (!f) {
		f = (Fonts*)malloc(sizeof(Fonts));
		f->biu = b;
		f->sz = sz;
		f->color = co;
		f->name = na;
		f->family = fa;
		f->charset = cha;
		f->scheme = sch;
		f->rgb = rg;
		f->indexed = ind;
		f->next = nullptr;
	}
	else {
		f->next = addfonts(f->next, b, sz, co, na, fa, cha, sch, rg, ind);
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

borders* styleread::addborder(borders* b, borderstyle* l, borderstyle* r, borderstyle* t, borderstyle* bo, borderstyle* d) {
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

stylexf* styleread::addstylexf(stylexf* sx, UINT8* n, UINT8* fo, UINT8* fi, UINT8* bi, UINT8* an, UINT8* ab, UINT8* aa, UINT8* ap, UINT8* av, UINT8* af, UINT8* afil) {
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
		sx->applyFont = af;
		sx->applyFill = afil;
		sx->next = nullptr;
	}
	else {
		sx->next = addstylexf(sx->next, n, fo, fi, bi, an, ab, aa, ap, av, af, afil);
	}

	return sx;
}

cellxfs* styleread::addcellxfs(cellxfs* c, UINT8* n, UINT8* fo, UINT8* fi, UINT8* bi, UINT8* an, UINT8* ab, UINT8* aa, UINT8* av, UINT8* at, UINT8* ho, UINT8* afo, UINT8* afi, UINT8* xid) {
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
		c->next = addcellxfs(c->next, n, fo, fi, bi, an, ab, aa, av, at, ho, afo, afi, xid);
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

Dxf* styleread::adddxf(Dxf* d, dxfFont* Fo, dxfFill* Fi) {
	if (!d) {
		d = (Dxf*)malloc(sizeof(Dxf));
		d->fon = Fo;
		d->fil = Fi;
		d->next = nullptr;
	}
	else {
		d->next = adddxf(d->next, Fo, Fi);
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
	const char colorgb[12] = "color rgb=\"";//11
	const char name[11] = "name val=\"";//10
	const char fami[13] = "family val=\"";//12
	const char chars[14] = "charset val=\"";//13
	const char sche[13] = "scheme val=\"";//12
	const char inde[16] = "color indexed=\"";//15

	const char Efont[] = "</font>";//7ji

	const char btag[5] = "<b/>";
	const char itag[5] = "<i/>";
	const char utag[5] = "<u/>";//4

	UINT8 endtag[8] = { 0 };//7����
	UINT8 biutag[5] = { 0 };//4����
	UINT8 strsear[14] = { 0 };//13����
	UINT8 seafam[13] = { 0 };//12����
	UINT8 seanam[11] = { 0 };//10����
	UINT8 Sz[9] = { 0 };//8����
	UINT8 rgb[12] = { 0 };//11����
	UINT8 index[16] = { 0 };//15����

	UINT8* SZ = (UINT8*)malloc(1); SZ = nullptr;
	UINT8* Name = (UINT8*)malloc(1); Name = nullptr;
	UINT8* Color = (UINT8*)malloc(1); Color = nullptr;
	UINT8* Family = (UINT8*)malloc(1); Color = nullptr;
	UINT8* Charset = (UINT8*)malloc(1); Charset = nullptr;
	UINT8* Scheme = (UINT8*)malloc(1); Scheme = nullptr;
	UINT8* RGB = (UINT8*)malloc(1); RGB = nullptr;
	UINT8* IND = (UINT8*)malloc(1); IND = nullptr;
	UINT8* BIU = (UINT8*)malloc(1); BIU = nullptr;

	int eresult = 1;
	int txtresult = 0;

	// </font>�܂œǂݍ���
	while (eresult != 0) {
		//1�����炷
		for (int i = 0; i < 14; i++) {
			index[i] = index[i + 1];
			if (i < 12)
				strsear[i] = strsear[i + 1];
			if (i < 6)
				endtag[i] = endtag[i + 1];
			if (i < 12 - 1)
				seafam[i] = seafam[i + 1];
			if (i < 10 - 1)
				seanam[i] = seanam[i + 1];
			if (i < 8 - 1)
				Sz[i] = Sz[i + 1];
			if (i < 11 - 1)
				rgb[i] = rgb[i + 1];
			if (i < 4 - 1)
				biutag[i] = biutag[i + 1];
		}
		biutag[3] = index[14] = rgb[10] = Sz[7] = strsear[12] = endtag[6] = seafam[11] = seanam[9] = dat[readp];
		biutag[4] = index[15] = rgb[11] = Sz[8] = strsear[13] = endtag[7] = seafam[12] = seanam[10] = '\0';
		readp++;

		//txtresult = strncmp((const char*)strsear, sche, 12);//font �X�^�[�g�^�O ����

		txtresult = strncmp((const char*)biutag, btag, 4);//indexed ����
		if (txtresult == 0) { //scheme �ǂݍ���
			free(BIU);
			BIU = (UINT8*)malloc(5);
			strcpy_s((char*)BIU, 5, btag);
		}

		txtresult = strncmp((const char*)biutag, itag, 4);//indexed ����
		if (txtresult == 0) { //scheme �ǂݍ���
			free(BIU);
			BIU = (UINT8*)malloc(5);
			strcpy_s((char*)BIU, 5, itag);
		}

		txtresult = strncmp((const char*)biutag, utag, 4);//indexed ����
		if (txtresult == 0) { //scheme �ǂݍ���
			free(BIU);
			BIU = (UINT8*)malloc(5);
			strcpy_s((char*)BIU, 5, utag);
		}

		txtresult = strncmp((const char*)index, inde, 15);//indexed ����
		if (txtresult == 0) //scheme �ǂݍ���
			Scheme = getValue(dat);

		txtresult = strncmp((const char*)seafam, sche, 12);//scheme ����
		if (txtresult == 0) //scheme �ǂݍ���
			Scheme = getValue(dat);

		txtresult = strncmp((const char*)strsear, chars, 13);//charset ����
		if (txtresult == 0) //charset �ǂݍ���
			Charset = getValue(dat);

		txtresult = strncmp((const char*)seafam, fami, 12);//family����
		if (txtresult == 0) { //family�ǂݍ���
			free(Family);
			Family = getValue(dat);
		}

		txtresult = strncmp((const char*)strsear, colo, 13);//color����
		if (txtresult == 0) {//color�ǂݍ���
			free(Color);
			Color = getValue(dat);
		}

		txtresult = strncmp((const char*)rgb, colorgb, 11);//color����
		if (txtresult == 0) {//color�ǂݍ���
			free(RGB);
			RGB = getValue(dat);
		}

		txtresult = strncmp((const char*)Sz, sz, 8);//sz����
		if (txtresult == 0) //sz�ǂݍ���
			SZ = getValue(dat);

		txtresult = strncmp((const char*)seanam, name, 10);//name����
		if (txtresult == 0) //name�ǂݍ���
			Name = getValue(dat);

		eresult = strcmp((const char*)endtag, Efont);
	}

	fontRoot[frcount]= (Fonts*)malloc(sizeof(Fonts));
	fontRoot[frcount]->biu = BIU;
	fontRoot[frcount]->sz = SZ;
	fontRoot[frcount]->color = Color;
	fontRoot[frcount]->name = Name;
	fontRoot[frcount]->family = Family;
	fontRoot[frcount]->charset = Charset;
	fontRoot[frcount]->scheme = Scheme;
	fontRoot[frcount]->rgb = RGB;
	fontRoot[frcount]->indexed = IND;

	frcount++;
}

UINT8* styleread::readfonts(UINT8* sd) {


	const char Sfont[] = "<font>";//6ji

	UINT8 startTag[7] = { 0 };//6����

	UINT8 endtag[9] = { 0 };//8����

	int eresult = 1;
	int txtresult = 0;
	while (eresult != 0) {
		//1�����炷
		for (int i = 0; i < 7; i++) {
			endtag[i] = endtag[i + 1];
			if (i < 6 - 1)
				startTag[i] = startTag[i + 1];
		}
		startTag[5] = endtag[7] = sd[readp];
		startTag[6] = endtag[8] = '\0';
		readp++;

		txtresult = strncmp((const char*)startTag, Sfont, 6);//font �X�^�[�g�^�O ����
		if (txtresult == 0)
			readfontV(sd);

		eresult = strcmp((const char*)endtag, Efonts);
	}

	return nullptr;
}

// <fill>����
void styleread::getfillV(UINT8* d) {
	const char type[14] = "patternType=\"";//13��
	const char Fgtag[9] = "<fgColor";//8ji
	const char Bgtag[9] = "<bgColor";//8ji
	const char EndFill[] = "</fill>";//7����

	UINT8 fbtag[9] = { 0 };//8����
	UINT8 endtag[8] = { 0 };//7����
	UINT8 Pattern[14] = { 0 };//13����

	UINT8* PaT = (UINT8*)malloc(1); PaT = nullptr;
	fgcolor* FG = (fgcolor*)malloc(sizeof(fgcolor)); FG = nullptr;
	bgcolor* BG = (bgcolor*)malloc(sizeof(bgcolor)); BG = nullptr;

	size_t msize = 0;//malloc�p

	size_t vallen = 0;//tag �l
	size_t taglen = 0;//�^�O�@�ǂݍ��݊J�n�ʒu

	int eresult = 1;
	int txtresult = 0;

	while (eresult != 0) {
		//1�����炷
		for (int i = 0; i < 12; i++) {
			Pattern[i] = Pattern[i + 1];
			if (i < 7)
				fbtag[i] = fbtag[i + 1];
			if (i < 6)
				endtag[i] = endtag[i + 1];
		}
		Pattern[12] = fbtag[7] = endtag[6] = d[readp];
		Pattern[13] = fbtag[8] = endtag[7] = '\0';
		readp++;

		txtresult = strncmp((const char*)Pattern, type, 13);//pattern Type ����
		if (txtresult == 0) { //pattern Type �ǂݍ���
			free(PaT);
			PaT = getValue(d);
		}

		txtresult = strncmp((const char*)fbtag, Fgtag, 8);//pattern Type ����
		if (txtresult == 0) {//pattern Type �ǂݍ���
			//fgColor ����
			free(FG);
			FG = readfillFg(d);
		}

		txtresult = strncmp((const char*)fbtag, Bgtag, 8);//pattern Type ����
		if (txtresult == 0) {//pattern Type �ǂݍ���
			//fgColor ����
			free(BG);
			BG = readbgcolor(d);
		}

		eresult = strcmp((const char*)endtag, EndFill);
	}

	fillroot[ficount] = (Fills*)malloc(sizeof(Fills));
	fillroot[ficount]->patten = (FillPattern*)malloc(sizeof(FillPattern));
	fillroot[ficount]->patten->patternType = PaT;
	fillroot[ficount]->fg = FG;
	fillroot[ficount]->bg = BG;

	ficount++;
}

// <fills>����
UINT8* styleread::readfill(UINT8* sd) {

	const char Startfill[] = "<fill>";//6����

	UINT8 endtag[9] = { 0 };//8����
	UINT8 starttag[7] = { 0 };//6����

	int eresult = 1;
	int txtresult = 0;

	while (eresult != 0) {
		//1�����炷
		for (int i = 0; i < 7; i++) {
			endtag[i] = endtag[i + 1];
			if (i < 5)
				starttag[i] = starttag[i + 1];
		}
		starttag[5] = endtag[7] = sd[readp];
		starttag[6] = endtag[8] = '\0';
		readp++;

		txtresult = strncmp((const char*)starttag, Startfill, 10);//pattern Type ����
		if (txtresult == 0) //pattern Type �ǂݍ���
			getfillV(sd);//�X�^�[�g�^�O�@�ǂݍ���

		eresult = strcmp((const char*)endtag, endfil);
	}

	return nullptr;
}

fgcolor* styleread::readfillFg(UINT8* dat) {
	const char theme[] = "theme=\"";//7��
	const char rgb[] = "rgb=\"";//5��
	const char fgtag[] = "tint=\"";//6ji

	UINT8 Theme[8] = { 0 };//7����
	UINT8 Fgtag[7] = { 0 };//6����
	UINT8 Rg[6] = { 0 };//5����

	UINT8* TH = (UINT8*)malloc(1); TH = nullptr;
	UINT8* SC = (UINT8*)malloc(1); SC = nullptr;
	UINT8* RGB = (UINT8*)malloc(1);  RGB = nullptr;

	int vallen = 0;
	size_t taglen = 0;
	size_t msize = 0;

	int result = 0;
	//fgColor �ǂݍ���
	while (dat[readp] != '>') {

		for (int i = 0; i < 6; i++) {
			Theme[i] = Theme[i + 1];
			if (i < 5)
				Fgtag[i] = Fgtag[i + 1];
			if (i < 4)
				Rg[i] = Rg[i + 1];
		}
		Rg[4] = Theme[6] = Fgtag[5] = dat[readp];
		Rg[5] = Theme[7] = Fgtag[6] = '\0';
		readp++;

		result = strncmp((const char*)Theme, theme, 7);//theme ����
		if (result == 0) {
			free(TH);
			TH = getValue(dat);
		}

		result = strncmp((const char*)Rg, rgb, 5);//theme ����
		if (result == 0) {
			free(RGB);
			RGB = getValue(dat);
		}

		result = strncmp((const char*)Fgtag, fgtag, 6);//fgtag ����
		if (result == 0)
			SC = getValue(dat);
	}

	fgcolor* FG = (fgcolor*)malloc(sizeof(fgcolor));
	FG->theme = TH;
	FG->tint = SC;
	FG->rgb = RGB;

	return FG;
}

bgcolor* styleread::readbgcolor(UINT8* dat) {

	const char indexed[10] = "indexed=\"";//9ji

	UINT8 Index[10] = { 0 };//9����

	UINT8* ID = (UINT8*)malloc(1); ID = nullptr;

	int vallen = 0;
	size_t taglen = 0;
	size_t msize = 0;

	int result = 0;
	//bgColor �ǂݍ���
	while (dat[readp] != '>') {

		for (int i = 0; i < 8; i++) {
			Index[i] = Index[i + 1];
		}
		Index[8] = dat[readp];
		Index[9] = '\0';
		readp++;

		result = strncmp((const char*)Index, indexed, 9);//indexed ����
		if (result == 0)
			ID = getValue(dat);//Index �l�擾
	}

	bgcolor* BG = (bgcolor*)malloc(sizeof(bgcolor));
	BG->indexed = ID;

	return BG;
}
//border value �擾
borderstyle* styleread::getborV(UINT8* dat, UINT8* endT, size_t endTlen) {

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

	UINT8* endtag = (UINT8*)malloc(endTlen + 1);
	memset(endtag, 0, endTlen + 1);

	int result = 1;
	int txtresult = 0;

	size_t maxlen = (endTlen > 9) ? endTlen : 9;

	UINT8* ST = nullptr;
	UINT8* AT = nullptr;
	UINT8* IND = (UINT8*)malloc(1); IND = nullptr;
	UINT8* Theme = (UINT8*)malloc(1); Theme = nullptr;
	UINT8* Tin = (UINT8*)malloc(1); Tin = nullptr;
	UINT8* Rgb = (UINT8*)malloc(1); Rgb = nullptr;

	borderstyle* bv = nullptr;

	if (dat[readp] == '/') {
		//�^�O�l�Ȃ�
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

			txtresult = strncmp((const char*)AUT, tint, 5);//auto ����
			if (txtresult == 0) //sz�ǂݍ���
				Tin = getValue(dat);

			txtresult = strncmp((const char*)RGB, rgb, 5);//auto ����
			if (txtresult == 0) //sz�ǂݍ���
				Rgb = getValue(dat);

			txtresult = strncmp((const char*)AUT, au, 6);//auto ����
			if (txtresult == 0) //sz�ǂݍ���
				AT = getValue(dat);

			txtresult = strncmp((const char*)Index, indexed, 9);//indexed ����
			if (txtresult == 0) //sz�ǂݍ���
				IND = getValue(dat);

			txtresult = strncmp((const char*)Style, style, 7);//style ����
			if (txtresult == 0) //sz�ǂݍ���
				ST = getValue(dat);

			result = strncmp((const char*)endtag, (const char*)endT, endTlen);
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

//border �l�擾
void styleread::getBorder(UINT8* d) {
	const char left[] = "<left";//5
	const char right[] = "<right";//6
	const char top[] = "<top";//4
	const char bottom[] = "<bottom";//7
	const char diag[] = "<diagonal";//9

	const char el[] = "</left>";//7
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

	while (eresult != 0)
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

		//�S��@�^�O�@�K������
		txtresult = strncmp((const char*)L, left, 5);//left����
		if (txtresult == 0)
			l = getborV(d, (UINT8*)el, 7);

		txtresult = strncmp((const char*)R, right, 6);//right����
		if (txtresult == 0)
			r = getborV(d, (UINT8*)er, 8);

		txtresult = strncmp((const char*)T, top, 4);//top����
		if (txtresult == 0)
			t = getborV(d, (UINT8*)et, 6);

		txtresult = strncmp((const char*)B, bottom, 7);//bottom����
		if (txtresult == 0)
			b = getborV(d, (UINT8*)eb, 9);

		txtresult = strncmp((const char*)D, diag, 9);//diagonal����
		if (txtresult == 0) {
			dia = getborV(d, (UINT8*)ed, 11);

			BorderRoot[bocount] = (borders*)malloc(sizeof(borders));
			BorderRoot[bocount]->left = l;
			BorderRoot[bocount]->right = r;
			BorderRoot[bocount]->top = t;
			BorderRoot[bocount]->bottom = b;
			BorderRoot[bocount]->diagonal = dia;
			bocount++;
		}
		eresult = strcmp((const char*)D, endtag);
	}
}

//border ����
UINT8* styleread::readboerder(UINT8* d) {
	const char startBor[] = "<border>";//8

	UINT8 starttag[9] = { 0 };//8����
	UINT8 endtag[11] = { 0 };//10����

	int eresult = 1;
	int txtresult = 0;
	while (eresult != 0) {
		//1�����炷
		for (int i = 0; i < 9; i++) {
			endtag[i] = endtag[i + 1];
			if (i < 7)
				starttag[i] = starttag[i + 1];
		}
		starttag[7] = endtag[9] = d[readp];
		starttag[8] = endtag[10] = '\0';
		readp++;

		txtresult = strncmp((const char*)starttag, startBor, 8);//sz����
		if (txtresult == 0) //sz�ǂݍ���
			getBorder(d);//�{�[�_�[�ǂݍ���

		eresult = strcmp((const char*)endtag, enbor);
	}

	return nullptr;
}

//tag""�l�擾
UINT8* styleread::getValue(UINT8* d) {
	size_t vallen = 0;
	size_t taglen = size_t(readp);
	size_t msize = 0;
	UINT8* tagval;
	tagval = nullptr;

	while (d[readp] != '"') {
		if (d[readp] != '"' && d[readp] != '/')//read data
			vallen++;
		readp++;
	}
	strl = vallen;
	msize = vallen + 1;

	tagval = (UINT8*)malloc(msize);
	for (size_t i = 0; i < vallen; i++)
		tagval[i] = d[taglen + i];
	tagval[vallen] = '\0';

	return tagval;
}

//CellStyleXfs�ǂݍ���
void styleread::getCellStyleXfsV(UINT8* d) {
	const char numFmtId[] = "numFmtId=\"";//10����
	const char fontId[] = "fontId=\"";//8����
	const char fillId[] = "fillId=\"";//8����
	const char borderId[] = "borderId=\"";//10����
	const char vertical[] = "vertical=\"";//10����
	const char applyNumberFormat[] = "applyNumberFormat=\"";//19����
	const char applyBorder[] = "applyBorder=\"";//13����
	const char applyAlignment[] = "applyAlignment=\"";//16����
	const char applyProtection[] = "applyProtection=\"";//17����
	const char afon[] = "applyFont=\"";//11����
	const char afil[] = "applyFill=\"";//11����

	const char endtag[] = "</xf>";//5����

	UINT8 nFI[11] = { 0 };//10����
	UINT8 Fid[9] = { 0 };//8����
	UINT8 borid[11] = { 0 };//10����
	UINT8 Et[6] = { 0 };//5����
	UINT8 anum[20] = { 0 };//19����
	UINT8 abor[14] = { 0 };//13����
	UINT8 aali[17] = { 0 };//16����
	UINT8 apro[18] = { 0 };//17����
	UINT8 afo[12] = { 0 };//11����

	UINT8* num = (UINT8*)malloc(1); num = nullptr;
	UINT8* fon = (UINT8*)malloc(1); fon = nullptr;
	UINT8* fil = (UINT8*)malloc(1); fil = nullptr;
	UINT8* bor = (UINT8*)malloc(1); bor = nullptr;
	UINT8* verti = (UINT8*)malloc(1); verti = nullptr;
	UINT8* Anu = (UINT8*)malloc(1); Anu = nullptr;
	UINT8* Abo = (UINT8*)malloc(1); Abo = nullptr;
	UINT8* Aal = (UINT8*)malloc(1); Aal = nullptr;
	UINT8* Apr = (UINT8*)malloc(1); Apr = nullptr;
	UINT8* Afo = (UINT8*)malloc(1); Afo = nullptr;
	UINT8* Afi = (UINT8*)malloc(1); Afi = nullptr;

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
			if (i < 11 - 1)
				afo[i] = afo[i + 1];
		}
		afo[10] = apro[16] = aali[15] = abor[12] = borid[9] = Fid[7] = Et[4] = d[readp];
		afo[11] = apro[17] = aali[16] = abor[13] = borid[10] = Fid[8] = Et[5] = '\0';
		readp++;

		//�I���^�O�@�Ȃ��@/>�I��
		if (d[readp - 1] == '/')
			if (d[readp] == '>')
				break;
		sresult = strncmp((const char*)apro, afon, 11);
		if (sresult == 0) {
			//apply font�@�l�ǂݍ���
			free(Afo);
			Afo = getValue(d);
		}
		sresult = strncmp((const char*)apro, afil, 11);
		if (sresult == 0) {
			//numFmtId�@�l�ǂݍ���
			free(Afi);
			Afi = getValue(d);
		}
		sresult = strncmp((const char*)afo, applyProtection, 17);
		if (sresult == 0) {
			//numFmtId�@�l�ǂݍ���
			free(Apr);
			Apr = getValue(d);
		}
		sresult = strncmp((const char*)aali, applyAlignment, 16);
		if (sresult == 0) {
			//numFmtId�@�l�ǂݍ���
			free(Aal);
			Aal = getValue(d);
		}
		sresult = strncmp((const char*)abor, applyBorder, 13);
		if (sresult == 0) {
			//numFmtId�@�l�ǂݍ���
			free(Abo);
			Abo = getValue(d);
		}
		sresult = strncmp((const char*)anum, applyNumberFormat, 19);
		if (sresult == 0) {
			//numFmtId�@�l�ǂݍ���
			free(Anu);
			Anu = getValue(d);
		}
		sresult = strncmp((const char*)borid, numFmtId, 10);
		if (sresult == 0) {
			//numFmtId�@�l�ǂݍ���
			free(num);
			num = getValue(d);
		}
		sresult = strncmp((const char*)borid, borderId, 10);
		if (sresult == 0) {
			//borderId�@�l�ǂݍ���
			free(bor);
			bor = getValue(d);
		}
		sresult = strncmp((const char*)borid, vertical, 10);
		if (sresult == 0) {
			//vertical�@�l�ǂݍ���
			free(verti);
			verti = getValue(d);
		}
		sresult = strncmp((const char*)Fid, fontId, 8);
		if (sresult == 0) {
			//fontId�@�l�ǂݍ���
			free(fon);
			fon = getValue(d);
		}
		sresult = strncmp((const char*)Fid, fillId, 8);
		if (sresult == 0) {
			//fillId�@�l�ǂݍ���
			free(fil);
			fil = getValue(d);
		}
		result = strncmp((const char*)Et, endtag, 5);
	}

	cellstyleXfsRoot[csXcount] = (stylexf*)malloc(sizeof(stylexf));
	cellstyleXfsRoot[csXcount]->numFmtId = num;
	cellstyleXfsRoot[csXcount]->fontId = fon;
	cellstyleXfsRoot[csXcount]->fillId = fil;
	cellstyleXfsRoot[csXcount]->borderId = bor;
	cellstyleXfsRoot[csXcount]->applyNumberFormat = Anu;
	cellstyleXfsRoot[csXcount]->applyBorder = Abo;
	cellstyleXfsRoot[csXcount]->applyAlignment = Aal;
	cellstyleXfsRoot[csXcount]->applyProtection = Apr;
	cellstyleXfsRoot[csXcount]->Avertical = verti;//alignment vertical
	cellstyleXfsRoot[csXcount]->applyFont = Afo;
	cellstyleXfsRoot[csXcount]->applyFill = Afi;

	csXcount++;
}

//CellXfs�ǂݍ���
void styleread::getCellXfsV(UINT8* d) {
	const char numFmtId[] = "numFmtId=\"";//10����
	const char fontId[] = "fontId=\"";//8����
	const char fillId[] = "fillId=\"";//8����
	const char borderId[] = "borderId=\"";//10����
	const char vertical[] = "vertical=\"";//10����
	const char applyNumberFormat[] = "applyNumberFormat=\"";//19����
	const char applyBorder[] = "applyBorder=\"";//13����
	const char applyAlignment[] = "applyAlignment=\"";//16����
	const char horizontal[] = "horizontal=\"";//12����
	const char wrapText[] = "wrapText=\"";//10����

	const char applyFont[] = "applyFont=\"";//11����
	const char applyFill[] = "applyFill=\"";//11����
	const char xfId[] = "xfId=\"";//6����

	const char endtag[] = "</xf>";//5����


	UINT8 aF[12] = { 0 };//11����
	UINT8 Fid[9] = { 0 };//8����
	UINT8 borid[11] = { 0 };//10����
	UINT8 Et[6] = { 0 };//5����
	UINT8 anum[20] = { 0 };//19����
	UINT8 abor[14] = { 0 };//13����
	UINT8 aali[17] = { 0 };//16����
	UINT8 apro[18] = { 0 };//17����
	UINT8 hori[18] = { 0 };//17����
	UINT8 xfid[7] = { 0 };//6����

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
		anum[18] = xfid[5] = aF[10] = hori[11] = apro[16] = aali[15] = abor[12] = borid[9] = Fid[7] = Et[4] = d[readp];
		anum[19] = xfid[6] = aF[11] = hori[12] = apro[17] = aali[16] = abor[13] = borid[10] = Fid[8] = Et[5] = '\0';
		readp++;

		//�I���^�O�@�Ȃ��@/>�I��
		/*if (d[readp - 1] == '/')
			if (d[readp] == '>')
				break;*/
		sresult = strncmp((const char*)xfid, xfId, 10);
		if (sresult == 0) {
			//numFmtId�@�l�ǂݍ���
			free(xId);
			xId = getValue(d);
		}
		sresult = strncmp((const char*)borid, wrapText, 10);
		if (sresult == 0) {
			//numFmtId�@�l�ǂݍ���
			free(Wt);
			Wt = getValue(d);
		}
		sresult = strncmp((const char*)hori, horizontal, 12);
		if (sresult == 0) {
			//numFmtId�@�l�ǂݍ���
			free(Ho);
			Ho = getValue(d);
		}
		sresult = strncmp((const char*)aF, applyFont, 11);
		if (sresult == 0) {
			//numFmtId�@�l�ǂݍ���
			free(aFo);
			aFo = getValue(d);
		}
		sresult = strncmp((const char*)aF, applyFill, 11);
		if (sresult == 0) {
			//numFmtId�@�l�ǂݍ���
			free(aFi);
			aFi = getValue(d);
		}
		sresult = strncmp((const char*)aali, applyAlignment, 16);
		if (sresult == 0) {
			//numFmtId�@�l�ǂݍ���
			free(Aal);
			Aal = getValue(d);
		}
		sresult = strncmp((const char*)abor, applyBorder, 13);
		if (sresult == 0) {
			//numFmtId�@�l�ǂݍ���
			free(Abo);
			Abo = getValue(d);
		}
		sresult = strncmp((const char*)anum, applyNumberFormat, 19);
		if (sresult == 0) {
			//numFmtId�@�l�ǂݍ���
			free(Anu);
			Anu = getValue(d);
		}
		sresult = strncmp((const char*)borid, numFmtId, 10);
		if (sresult == 0) {
			//numFmtId�@�l�ǂݍ���
			free(num);
			num = getValue(d);
		}
		sresult = strncmp((const char*)borid, borderId, 10);
		if (sresult == 0) {
			//borderId�@�l�ǂݍ���
			free(bor);
			bor = getValue(d);
		}
		sresult = strncmp((const char*)borid, vertical, 10);
		if (sresult == 0) {
			//vertical�@�l�ǂݍ���
			free(verti);
			verti = getValue(d);
		}
		sresult = strncmp((const char*)Fid, fontId, 8);
		if (sresult == 0) {
			//fontId�@�l�ǂݍ���
			free(fon);
			fon = getValue(d);
		}
		sresult = strncmp((const char*)Fid, fillId, 8);
		if (sresult == 0) {
			//fillId�@�l�ǂݍ���
			free(fil);
			fil = getValue(d);
		}
		result = strncmp((const char*)Et, endtag, 5);
	}

	cellXfsRoot[cXcount] = (cellxfs*)malloc(sizeof(cellxfs));
	cellXfsRoot[cXcount]->numFmtId = num;
	cellXfsRoot[cXcount]->fontId = fon;
	cellXfsRoot[cXcount]->fillId = fil;
	cellXfsRoot[cXcount]->borderId = bor;
	cellXfsRoot[cXcount]->applyNumberFormat = Anu;
	cellXfsRoot[cXcount]->applyBorder = Abo;
	cellXfsRoot[cXcount]->applyAlignment = Aal;
	cellXfsRoot[cXcount]->Avertical = verti;
	cellXfsRoot[cXcount]->AwrapText = Wt;
	cellXfsRoot[cXcount]->horizontal = Ho;
	cellXfsRoot[cXcount]->applyFont = aFo;
	cellXfsRoot[cXcount]->applyFill = aFi;
	cellXfsRoot[cXcount]->xfId = xId;

	cXcount++;

}

//cellStyleXfs �ǂݎ��
void styleread::readCellStyleXfs(UINT8* d) {
	const char starttag[] = "<xf";//3����

	UINT8 stag[4] = { 0 };//3����
	UINT8 endtag[16] = { 0 };//15����

	int result = 1;
	int sresult = 0;

	while (result != 0) {
		for (int i = 0; i < 15 - 1; i++) {
			endtag[i] = endtag[i + 1];
			if (i < 2)
				stag[i] = stag[i + 1];
		}
		endtag[14] = stag[2] = d[readp];
		endtag[15] = stag[3] = '\0';
		readp++;

		sresult = strncmp((const char*)stag, starttag, 3);
		if (sresult == 0) {
			//�X�^�[�g�^�O�@�l�ǂݍ���
			getCellStyleXfsV(d);
		}

		result = strncmp((const char*)endtag, EncsXfs, 15);
	}
}

//cellXfs �ǂݎ��
void styleread::readCellXfs(UINT8* d) {
	const char starttag[] = "<xf";//3����

	UINT8 stag[4] = { 0 };//3����
	UINT8 endtag[11] = { 0 };//10����

	int result = 1;
	int sresult = 0;

	while (result != 0) {
		for (int i = 0; i < 10 - 1; i++) {
			endtag[i] = endtag[i + 1];
			if (i < 2)
				stag[i] = stag[i + 1];
		}
		endtag[9] = stag[2] = d[readp];
		endtag[10] = stag[3] = '\0';
		readp++;

		sresult = strncmp((const char*)stag, starttag, 3);
		if (sresult == 0) {
			//�X�^�[�g�^�O�@�l�ǂݍ���
			getCellXfsV(d);
		}

		result = strncmp((const char*)endtag, Exfs, 10);
	}
}

//numFmts �l�擾
void styleread::getnumFmts(UINT8* dat) {
	const char numFmtId[] = "numFmtId=\"";//10����
	const char formatCode[] = "formatCode=\"";//12����

	UINT8 stag[11] = { 0 };//10����
	UINT8 code[13] = { 0 };//12����

	int result = 0;
	int eresu = 0;

	UINT8* ID = (UINT8*)malloc(1); ID = nullptr;
	UINT8* COD = (UINT8*)malloc(1); COD = nullptr;

	while (dat[readp] != '>') {
		for (int i = 0; i < 11; i++) {
			code[i] = code[i + 1];
			if (i < 9)
				stag[i] = stag[i + 1];
		}
		code[11] = stag[9] = dat[readp];
		code[12] = stag[10] = '\0';
		readp++;

		result = strncmp((const char*)stag, numFmtId, 10);
		if (result == 0) {
			//read values
			free(ID);
			ID = getValue(dat);
			std::cout << "get numfmts id" << ID << std::endl;
		}

		result = strncmp((const char*)code, formatCode, 12);
		if (result == 0) {
			//read values
			free(COD);
			COD = getValue(dat);
		}
	}

	numFmtsRoot[nucount] = (numFMts*)malloc(sizeof(numFMts));
	numFmtsRoot[nucount]->Id = ID;
	numFmtsRoot[nucount]->Code = COD;

	nucount++;
}

void styleread::readnumFmt(UINT8* d) {
	const char numFmt[] = "<numFmt";//7����

	UINT8 stag[8] = { 0 };//7����
	UINT8 etag[11] = { 0 };//10����

	int result = 0;
	int eresu = 1;

	while (eresu != 0) {
		for (int i = 0; i < 9; i++) {
			etag[i] = etag[i + 1];
			if (i < 6)
				stag[i] = stag[i + 1];
		}
		etag[9] = stag[6] = d[readp];
		etag[10] = stag[7] = '\0';
		readp++;

		result = strncmp((const char*)stag, numFmt, 7);
		if (result == 0) {
			//read values
			getnumFmts(d);
		}

		eresu = strncmp((const char*)etag, Efmts, 10);
	}
}
//cellstyle �l�擾
void styleread::getcellStyle(UINT8* d) {
	//std::cout << "get cellstyle" << std::endl;
	//std::cout << std::endl;
	const char name[] = "name=\"";//6����
	const char xfId[] = "xfId=\"";//6����
	const char builtinId[] = "builtinId=\"";//11����
	const char customBuiltin[] = "customBuiltin=\"";//15����
	const char xruid[] = "xr:uid=\"";//8����

	UINT8 naid[7] = { 0 };//6����
	UINT8 build[12] = { 0 };//11����
	UINT8 cbuilt[16] = { 0 };//15����
	UINT8 xrld[9] = { 0 };//8����

	int result = 0;
	int eresu = 0;

	UINT8* NAM = (UINT8*)malloc(1); NAM = nullptr;
	UINT8* XfID = (UINT8*)malloc(1); XfID = nullptr;
	UINT8* BUID = (UINT8*)malloc(1); BUID = nullptr;
	UINT8* CUBID = (UINT8*)malloc(1); CUBID = nullptr;
	UINT8* XrID = (UINT8*)malloc(1); XrID = nullptr;

	while (d[readp] != '/') {
		for (int i = 0; i < 14; i++) {
			cbuilt[i] = cbuilt[i + 1];
			if (i < 10)
				build[i] = build[i + 1];
			if (i < 7)
				xrld[i] = xrld[i + 1];
			if (i < 5)
				naid[i] = naid[i + 1];
		}
		cbuilt[14] = build[10] = xrld[7] = naid[5] = d[readp];
		cbuilt[15] = build[11] = xrld[8] = naid[6] = '\0';
		readp++;

		result = strncmp((const char*)cbuilt, customBuiltin, 15);
		if (result == 0) {
			//read values
			free(CUBID);
			CUBID = getValue(d);
		}

		result = strncmp((const char*)build, builtinId, 11);
		if (result == 0) {
			//read values
			free(BUID);
			BUID = getValue(d);
		}

		result = strncmp((const char*)xrld, xruid, 8);
		if (result == 0) {
			//read values
			free(XrID);
			XrID = getValue(d);
		}

		result = strncmp((const char*)naid, name, 6);
		if (result == 0) {
			//read values
			free(NAM);
			NAM = getValue(d);
		}

		result = strncmp((const char*)naid, xfId, 6);
		if (result == 0) {
			//read values
			free(XfID);
			XfID = getValue(d);
		}
	}

	cellStyleRoot[cscount] = (cellstyle*)malloc(sizeof(cellstyle));
	cellStyleRoot[cscount]->name = NAM;
	cellStyleRoot[cscount]->xfId = XfID;
	cellStyleRoot[cscount]->builtinId = BUID;
	cellStyleRoot[cscount]->customBuilt = CUBID;
	cellStyleRoot[cscount]->xruid = XrID;

	cscount++;
}
//cellstyle �ǂݍ���
void styleread::readcellStyles(UINT8* d) {

	const char cellStyle[] = "<cellStyle";//10����

	UINT8 stag[11] = { 0 };//10����
	UINT8 etag[14] = { 0 };//13����

	int result = 0;
	int eresu = 1;

	while (eresu != 0) {
		for (int i = 0; i < 12; i++) {
			etag[i] = etag[i + 1];
			if (i < 9)
				stag[i] = stag[i + 1];
		}
		etag[12] = stag[9] = d[readp];
		etag[13] = stag[10] = '\0';
		readp++;

		result = strncmp((const char*)stag, cellStyle, 10);
		if (result == 0) {
			//read values
			getcellStyle(d);
		}

		eresu = strncmp((const char*)etag, Ecellsty, 13);
	}
}
//dxf font read
dxfFont* styleread::readdxffont(UINT8* d) {
	const char val[] = "val=\"";//5����
	const char rgb[] = "rgb=\"";//5����
	const char theme[] = "theme=\"";//7����
	const char etag[] = "</font>";//7����
	const char btag[] = "<b/>";//4����

	UINT8 bol[5] = { 0 };//4����
	UINT8 Val[6] = { 0 };//5����
	UINT8 Them[8] = { 0 };//7����

	int result = 0;
	int eresu = 1;

	UINT8* VAL = (UINT8*)malloc(1); VAL = nullptr;
	UINT8* THE = (UINT8*)malloc(1); THE = nullptr;
	UINT8* RGB = (UINT8*)malloc(1); RGB = nullptr;
	UINT8* B = (UINT8*)malloc(1); B = nullptr;

	while (eresu != 0) {
		for (int i = 0; i < 6; i++) {
			Them[i] = Them[i + 1];
			if (i < 4)
				Val[i] = Val[i + 1];
			if (i < 3)
				bol[i] = bol[i + 1];
		}
		bol[3] = Them[6] = Val[4] = d[readp];
		bol[4] = Them[7] = Val[5] = '\0';
		readp++;

		result = strncmp((const char*)Val, val, 5);
		if (result == 0) {
			//read values
			free(VAL);
			VAL = getValue(d);
		}

		result = strncmp((const char*)bol, btag, 4);
		if (result == 0) {
			//read values
			free(B);
			B = (UINT8*)malloc(5);
			strcpy_s((char*)B, 5, btag);
		}

		result = strncmp((const char*)Val, rgb, 5);
		if (result == 0) {
			//read values
			free(RGB);
			RGB = getValue(d);
		}

		result = strncmp((const char*)Them, theme, 7);
		if (result == 0) {
			//read values
			free(THE);
			THE = getValue(d);
		}

		eresu = strncmp((const char*)Them, etag, 7);

	}

	dxfFont* f = (dxfFont*)malloc(sizeof(dxfFont));
	f->b = B;
	f->ival = VAL;
	f->theme = THE;
	f->rgb = RGB;

	return f;
}
//dxf fill read
dxfFill* styleread::readdxffill(UINT8* d) {
	const char rgb[] = "rgb=\"";//5����
	const char etag[] = "</fill>";//7����

	UINT8 Val[6] = { 0 };//5����
	UINT8 Them[8] = { 0 };//7����

	int result = 0;
	int eresu = 1;

	UINT8* RGB = (UINT8*)malloc(1); RGB = nullptr;

	while (eresu != 0) {
		for (int i = 0; i < 6; i++) {
			Them[i] = Them[i + 1];
			if (i < 4)
				Val[i] = Val[i + 1];
		}
		Them[6] = Val[4] = d[readp];
		Them[7] = Val[5] = '\0';
		readp++;

		result = strncmp((const char*)Val, rgb, 5);
		if (result == 0) {
			//read values
			free(RGB);
			RGB = getValue(d);
		}

		eresu = strncmp((const char*)Them, etag, 7);

	}

	dxfFill* f = (dxfFill*)malloc(sizeof(dxfFill));
	f->rgb = RGB;

	return f;
}

//dXfs font fill�ǂݍ���
void styleread::getdxfs(UINT8* d) {
	const char font[] = "<font>";//6����
	const char fill[] = "<fill>";//6����
	const char etag[] = "</dxf>";//6����

	UINT8 ff[7] = { 0 };//6����

	int result = 0;
	int eresu = 1;

	dxfFont* fo = nullptr;
	dxfFill* fi = nullptr;

	while (eresu != 0) {
		for (int i = 0; i < 5; i++) {
			ff[i] = ff[i + 1];
		}
		ff[5] = d[readp];
		ff[6] = '\0';
		readp++;

		result = strncmp((const char*)ff, font, 6);
		if (result == 0) {
			//read values
			fo = readdxffont(d);
		}

		result = strncmp((const char*)ff, fill, 6);
		if (result == 0) {
			//read values
			fi = readdxffill(d);
		}

		eresu = strncmp((const char*)ff, etag, 6);
	}

	dxfRoot[dxcount] = (Dxf*)malloc(sizeof(Dxf));
	dxfRoot[dxcount]->fon = fo;
	dxfRoot[dxcount]->fil = fi;

	dxcount++;
}

//dXfs �ǂݍ���
void styleread::readdxfs(UINT8* d) {
	const char dxf[] = "<dxf>";//5����

	UINT8 stag[6] = { 0 };//5����
	UINT8 etag[8] = { 0 };//7����

	int result = 0;
	int eresu = 1;

	while (eresu != 0) {
		for (int i = 0; i < 6; i++) {
			etag[i] = etag[i + 1];
			if (i < 4)
				stag[i] = stag[i + 1];
		}
		etag[6] = stag[4] = d[readp];
		etag[7] = stag[5] = '\0';
		readp++;

		result = strncmp((const char*)stag, dxf, 5);
		if (result == 0) {
			//read values
			getdxfs(d);
		}

		eresu = strncmp((const char*)etag, Edxf, 7);
	}
}
//colors �l�擾
void styleread::getcolor(UINT8* d) {
	const char rgb[] = "rgb=\"";//5����

	UINT8 Rg[6] = { 0 };//5����

	int result = 0;

	UINT8* RGB = (UINT8*)malloc(1); RGB = nullptr;

	while (d[readp] != '/') {
		for (int i = 0; i < 4; i++) {
			Rg[i] = Rg[i + 1];
		}
		Rg[4] = d[readp];
		Rg[5] = '\0';
		readp++;

		result = strncmp((const char*)Rg, rgb, 5);
		if (result == 0) {
			//read values
			free(RGB);
			RGB = getValue(d);
		}
	}

	colorsRoot = addcolor(colorsRoot, RGB);

}
//colors �ǂݍ���
void styleread::readcolors(UINT8* d) {
	const char color[] = "<color";//6����

	UINT8 stag[7] = { 0 };//6����
	UINT8 etag[10] = { 0 };//9����

	int result = 0;
	int eresu = 1;

	while (eresu != 0) {
		for (int i = 0; i < 8; i++) {
			etag[i] = etag[i + 1];
			if (i < 5)
				stag[i] = stag[i + 1];
		}
		etag[8] = stag[5] = d[readp];
		etag[9] = stag[6] = '\0';
		readp++;

		result = strncmp((const char*)stag, color, 6);
		if (result == 0) {
			//read values
			getcolor(d);
		}

		eresu = strncmp((const char*)etag, Ecolor, 9);
	}
}
//�ŏI������
void styleread::readextLst(UINT8* d) {
	const char Endtag[] = "</extLst>";//9

	UINT8 etag[10] = { 0 };//9����

	int eresu = 1;
	size_t stockp = size_t(readp);

	while (eresu != 0) {
		for (int i = 0; i < 8; i++) {
			etag[i] = etag[i + 1];
		}
		etag[8] = d[readp];
		etag[9] = '\0';
		readp++;

		eresu = strncmp((const char*)etag, Endtag, 9);
	}

	size_t dis = size_t(readp) - stockp;//�����񒷂�
	size_t msi = dis + 1;

	extLstStr = (UINT8*)malloc(msi);

	//�ŏI������R�s�[
	for (size_t i = 0; i < dis; i++)
		extLstStr[i] = d[stockp + i];

	extLstStr[dis] = '\0';

	std::cout << "extLst : " << extLstStr << std::endl;
}

void styleread::readstyle(UINT8* sdata, UINT64 sdatalen)
{
	const char count[] = "count=\"";//count����
	const char knownFonts[] = "x14ac:knownFonts=\"";//18
	kFonts = (UINT8*)malloc(1); kFonts = nullptr;

	UINT8 sEs[14] = { 0 };//13����
	UINT8 sExfs[16] = { 0 };//15����
	UINT8 sfon[7] = { 0 };//6����
	UINT8 sdxf[6] = { 0 };//5����
	UINT8 sbor[9] = { 0 };//8����
	UINT8 sstyle[12] = { 0 };//11����

	UINT8 Cou[8] = { 0 };
	UINT8 knoF[19] = { 0 };

	int result = 1;
	int otherresult = 0;
	int exresult = 0;

	int vallen = 0;
	int taglen = 0;

	size_t msize = 0;

	//��������
	while (result != 0) {
		//1�����炷
		for (int i = 0; i < 15 - 1; i++) {
			sExfs[i] = sExfs[i + 1];//�G���h�^�O
			if (i < 13 - 1)
				sEs[i] = sEs[i + 1];
			if (i < 6 - 1)
				sfon[i] = sfon[i + 1];
			if (i < 8 - 1)
				sbor[i] = sbor[i + 1];
			if (i < 11 - 1)
				sstyle[i] = sstyle[i + 1];
			if (i < 5 - 1)
				sdxf[i] = sdxf[i + 1];
		}
		sExfs[14] = sEs[12] = sfon[5] = sbor[7] = sstyle[10] = sdxf[4] = sdata[readp];
		sExfs[15] = sEs[13] = sfon[6] = sbor[8] = sstyle[11] = sdxf[5] = '\0';
		readp++;

		otherresult = strncmp((const char*)sstyle, styleSheet, 11);
		if (otherresult == 0) {
			//stylesheet �J�n�^�O
			//���^�O�܂œǂݍ���
			while (sdata[readp] != '>')
				readp++;

			size_t slen = size_t(readp) + 1;
			styleSheetStr = (UINT8*)malloc(slen);

			for (size_t i = 0; i < slen - 1; i++)
				styleSheetStr[i] = sdata[i];

			styleSheetStr[slen - 1] = '\0';

			std::cout << "style sheet start : " << styleSheetStr << std::endl;
		}

		otherresult = strncmp((const char*)sbor, Fmts, 8);
		if (otherresult == 0) {
			//numFmt ��v
			std::cout << "read Fmts" << sbor << std::endl;
			while (sdata[readp] != '>') {
				for (int i = 0; i < 7 - 1; i++) {
					Cou[i] = Cou[i + 1];
				}
				Cou[6] = sdata[readp];
				Cou[7] = '\0';
				readp++;

				result = strncmp((const char*)Cou, count, 7);
				if (result == 0) {
					numFmtsCount = getValue(sdata);//count �擾
					break;
				}
			}
			numFmtsNum = strchange.RowArraytoNum(numFmtsCount, strl);

			numFmtsRoot = (numFMts**)malloc(sizeof(numFMts) * numFmtsNum);
			readnumFmt(sdata);
		}

		otherresult = strncmp((const char*)sfon, fonts, 6);
		if (otherresult == 0) {
			std::cout << "read fonts" << sfon << std::endl;
			//�t�H���g��v read font
			while (sdata[readp] != '>') {
				for (int i = 0; i < 18 - 1; i++) {
					knoF[i] = knoF[i + 1];
					if (i < 7 - 1)
						Cou[i] = Cou[i + 1];
				}
				knoF[18 - 1] = Cou[6] = sdata[readp];
				knoF[18] = Cou[7] = '\0';
				readp++;

				exresult = strncmp((const char*)knoF, knownFonts, 18);
				result = strncmp((const char*)Cou, count, 7);
				if (result == 0) {
					fontCount = getValue(sdata);//count �擾
					fontNum = strchange.RowArraytoNum(fontCount, strl);
					fontRoot = (Fonts**)malloc(sizeof(Fonts) * fontNum);
				}
				if (exresult == 0) {
					std::cout << "known font : " << knoF << std::endl;
					free(kFonts);
					kFonts = getValue(sdata);
				}
			}
			readfonts(sdata);//�t�H���g�^�O�擾
		}

		otherresult = strncmp((const char*)sfon, fills, 6);
		if (otherresult == 0) {
			std::cout << "read fills" << sfon << std::endl;
			//�t�B����v read font
			while (sdata[readp] != '>') {
				for (int i = 0; i < 7 - 1; i++) {
					Cou[i] = Cou[i + 1];
				}
				Cou[6] = sdata[readp];
				Cou[7] = '\0';
				readp++;

				result = strncmp((const char*)Cou, count, 7);
				if (result == 0)
					fillCount = getValue(sdata);//count �擾
			}
			fillNum = strchange.RowArraytoNum(fillCount, strl);
			fillroot = (Fills**)malloc(sizeof(Fills) * fillNum);
			readfill(sdata);
		}

		otherresult = strncmp((const char*)sbor, border, 8);
		if (otherresult == 0) {
			std::cout << "read border" << sbor << std::endl;
			//�{�[�_�[��v read border
			while (sdata[readp] != '>') {
				for (int i = 0; i < 7 - 1; i++) {
					Cou[i] = Cou[i + 1];
				}
				Cou[6] = sdata[readp];
				Cou[7] = '\0';
				readp++;

				result = strncmp((const char*)Cou, count, 7);
				if (result == 0)
					borderCount = getValue(sdata);//count �擾				
			}
			borderNum = strchange.RowArraytoNum(borderCount, strl);
			BorderRoot = (borders**)malloc(sizeof(borders)*borderNum);
			readboerder(sdata);
		}

		otherresult = strncmp((const char*)sEs, cellstXfs, 13);
		if (otherresult == 0) {
			std::cout << "read cellstXfs" << sEs << std::endl;
			//�Z���X�^�C����v read cellstyleXfs
			while (sdata[readp] != '>') {
				for (int i = 0; i < 7 - 1; i++) {
					Cou[i] = Cou[i + 1];
				}
				Cou[6] = sdata[readp];
				Cou[7] = '\0';
				readp++;

				result = strncmp((const char*)Cou, count, 7);
				if (result == 0)
					cellStyleXfsCount = getValue(sdata);//count �擾
			}
			cellstyleXfsNum = strchange.RowArraytoNum(cellStyleXfsCount, strl);
			cellstyleXfsRoot = (stylexf**)malloc(sizeof(stylexf) * cellstyleXfsNum);
			readCellStyleXfs(sdata);
		}

		otherresult = strncmp((const char*)sbor, Xfs, 8);
		if (otherresult == 0) {
			std::cout << "read Xfs" << sbor << std::endl;
			//cellXfs��v read cellxfs
			while (sdata[readp] != '>') {
				for (int i = 0; i < 7 - 1; i++) {
					Cou[i] = Cou[i + 1];
				}
				Cou[6] = sdata[readp];
				Cou[7] = '\0';
				readp++;

				result = strncmp((const char*)Cou, count, 7);
				if (result == 0) 
					cellXfsCount = getValue(sdata);//count �擾
			}
			cellXfsNum = strchange.RowArraytoNum(cellXfsCount, strl);
			cellXfsRoot = (cellxfs**)malloc(sizeof(cellxfs) * cellXfsNum);
			readCellXfs(sdata);
		}

		otherresult = strncmp((const char*)sstyle, CellStyl, 11);
		if (otherresult == 0) {
			std::cout << "read CellStyl" << sstyle << std::endl;
			//cellStyle��v read cellstyle
			while (sdata[readp] != '>') {
				for (int i = 0; i < 7 - 1; i++) {
					Cou[i] = Cou[i + 1];
				}
				Cou[6] = sdata[readp];
				Cou[7] = '\0';
				readp++;

				result = strncmp((const char*)Cou, count, 7);
				if (result == 0) 
					cellstyleCount = getValue(sdata);//count �擾
			}
			cellstyleNum = strchange.RowArraytoNum(cellstyleCount, strl);
			cellStyleRoot = (cellstyle**)malloc(sizeof(cellstyle) * cellstyleNum);
			readcellStyles(sdata);
		}

		otherresult = strncmp((const char*)sdxf, dxf, 5);
		if (otherresult == 0) {
			std::cout << "read dxf" << sdxf << std::endl;
			//dxfs��v read dxfs
			while (sdata[readp] != '>') {
				for (int i = 0; i < 7 - 1; i++) {
					Cou[i] = Cou[i + 1];
				}
				Cou[6] = sdata[readp];
				Cou[7] = '\0';
				readp++;

				result = strncmp((const char*)Cou, count, 7);
				if (result == 0) 
					dxfsCount = getValue(sdata);//count �擾
			}
			dxfsNm = strchange.RowArraytoNum(dxfsCount, strl);
			dxfRoot = (Dxf**)malloc(sizeof(Dxf) * dxfsNm);
			readdxfs(sdata);
		}

		otherresult = strncmp((const char*)sbor, color, 8);
		if (otherresult == 0) {
			//color��v read colors
			std::cout << "read color" << sbor << std::endl;
			readcolors(sdata);
		}

		otherresult = strncmp((const char*)sbor, extLst, 8);
		if (otherresult == 0) {
			//extLst ��v
			readextLst(sdata);
		}

		result = strncmp(Endstyle, (const char*)sEs, 13);
	}

}
