#pragma once

#include "excel_style.h"



class StyleWrite :public styleread
{
public:
	UINT8* wd;
	UINT64 wdlen;

	StyleWrite();
	~StyleWrite();

	void styledatawrite(UINT64 dlen);

	void Dxfswrite();

	void writeColors();

	void writefirst();

	void writeextLst();

	void CELLStyle();

	void cellXfswrite();
	void fontswrite();
	void fillwrite();
	void borderwrite();
	void cellstyleXfs();
	void numFmidwrite();

	void oneStrplusDoubleq(UINT8* str, UINT8* v);

	void oneStrwrite(UINT8* str);

	void writedata(UINT8* stag, UINT8* v, UINT8* etag);

private:

};

StyleWrite::StyleWrite()
{
	wd = nullptr;
	wdlen = 0;
}

StyleWrite::~StyleWrite()
{
	free(wd);
}

inline void StyleWrite::styledatawrite(UINT64 dlen) {
	size_t msi = (size_t)dlen + 5000;

	wd = (UINT8*)malloc(msi);

	writefirst();

	numFmidwrite();

	fontswrite();

	fillwrite();

	borderwrite();

	cellstyleXfs();

	cellXfswrite();

	CELLStyle();

	Dxfswrite();

	writeColors();

	writeextLst();

	oneStrwrite((UINT8*)Endstyle);
}

//write dxfs
inline void StyleWrite::Dxfswrite()
{
	UINT8 et[] = "/>";
	UINT8 e[] = ">";

	UINT8 stag[] = "<dxf>";
	UINT8 etag[] = "</dxf>";

	UINT8 fstag[] = "<font>";
	UINT8 fetag[] = "</font>";

	UINT8 col[] = "<color";

	UINT8 rgb[] = " rgb=\"";

	UINT8 sfil[] = "<fill>";
	UINT8 efil[] = "</fill>";

	UINT8 spat[] = "<patternFill>";
	UINT8 epat[] = "</patternFill>";

	UINT8 bgc[] = "<bgColor";
	UINT8 istr[] = "<i";

	UINT8 val[] = " val=\"";

	Dxf** dx = dxfRoot;
	size_t cou = 0;

	oneStrwrite((UINT8*)dxf);
	oneStrplusDoubleq((UINT8*)countstr, dxfsCount);
	oneStrwrite(e);

	while (cou< dxfsNm) {
		oneStrwrite(stag);

		if (dx[cou]->fon) {
			oneStrwrite(fstag);

			if (dx[cou]->fon->b)
				oneStrwrite(dx[cou]->fon->b);
			if (dx[cou]->fon->ival) {
				oneStrwrite(istr);
				oneStrplusDoubleq(val, dx[cou]->fon->ival);
				oneStrwrite(et);
			}

			if (dx[cou]->fon->rgb)
			{
				oneStrwrite(col);
				oneStrplusDoubleq(rgb, dx[cou]->fon->rgb);
				oneStrwrite(et);
			}

			oneStrwrite(fetag);
		}
		if (dx[cou]->fil) {
			oneStrwrite(sfil);

			if (dx[cou]->fil->rgb) {
				oneStrwrite(spat);
				oneStrwrite(bgc);
				oneStrplusDoubleq(rgb, dx[cou]->fil->rgb);
				oneStrwrite(et);
				oneStrwrite(epat);
			}

			oneStrwrite(efil);
		}

		oneStrwrite(etag);

		cou++;
	}

	oneStrwrite((UINT8*)Edxf);

}
//write colors
inline void StyleWrite::writeColors()
{
	UINT8 et[] = "/>";
	UINT8 e[] = ">";

	UINT8 stag[] = "<mruColors>";
	UINT8 etag[] = "</mruColors>";

	UINT8 col[] = "<color";

	UINT8 rgb[] = " rgb=\"";

	colors* c = colorsRoot;

	if (c) {
		oneStrwrite((UINT8*)color);

		oneStrwrite(stag);
		while (c) {

			oneStrwrite(col);
			oneStrplusDoubleq(rgb, c->color);
			oneStrwrite(et);

			c = c->next;
		}

		oneStrwrite(etag);

		oneStrwrite((UINT8*)Ecolor);

	}
}
//write �ŏ��̕�����
inline void StyleWrite::writefirst() {
	UINT8 e[] = ">";
	oneStrwrite(styleSheetStr);
	oneStrwrite(e);
}
//write extLst
inline void StyleWrite::writeextLst() {
	UINT8 stag[] = "<extLst>";

	oneStrwrite(stag);
	oneStrwrite(extLstStr);
}

//write cellstyle
inline void StyleWrite::CELLStyle()
{
	UINT8 et[] = "/>";
	UINT8 e[] = ">";

	UINT8 stag[] = "<cellStyle";

	UINT8 nam[] = " name=\"";
	UINT8 xfid[] = " xfId=\"";
	UINT8 xruid[] = " xr:uid=\"";
	UINT8 bui[] = " builtinId=\"";
	UINT8 cb[] = " customBuilt=\"";

	cellstyle** cs = cellStyleRoot;

	size_t cou = 0;

	oneStrwrite((UINT8*)CellStyl);
	oneStrplusDoubleq((UINT8*)countstr, cellstyleCount);
	oneStrwrite(e);

	while (cou<cellstyleNum) {
		oneStrwrite(stag);

		if (cs[cou]->name)
			oneStrplusDoubleq(nam, cs[cou]->name);
		if (cs[cou]->xfId)
			oneStrplusDoubleq(xfid, cs[cou]->xfId);
		if (cs[cou]->xruid)
			oneStrplusDoubleq(xruid, cs[cou]->xruid);
		if (cs[cou]->builtinId)
			oneStrplusDoubleq(bui, cs[cou]->builtinId);
		if (cs[cou]->customBuilt)
			oneStrplusDoubleq(cb, cs[cou]->customBuilt);

		oneStrwrite(et);// />�^�O

		cou++;
	}

	oneStrwrite((UINT8*)Ecellsty);
}
//write cellXfs
inline void StyleWrite::cellXfswrite()
{
	UINT8 et[] = "/>";
	UINT8 e[] = ">";

	UINT8 stag[] = "<xf";
	UINT8 etag[] = "</xf>";

	UINT8 num[] = " numFmtId=\"";
	UINT8 fon[] = " fontId=\"";
	UINT8 fil[] = " fillId=\"";
	UINT8 bor[] = " borderId=\"";
	UINT8 an[] = " applyNumberFormat=\"";
	UINT8 ab[] = " applyBorder=\"";
	UINT8 aa[] = " applyAlignment=\"";
	UINT8 ap[] = " applyProtection=\"";
	UINT8 af[] = " applyFill=\"";
	UINT8 align[] = "<alignment";
	UINT8 ver[] = " vertical=\"";
	UINT8 afo[] = " applyFont=\"";
	UINT8 hor[] = " horizontal=\"";
	UINT8 wt[] = " wrapText=\"";
	UINT8 xi[] = " xfId=\"";

	cellxfs** cx = cellXfsRoot;

	size_t cou = 0;

	oneStrwrite((UINT8*)Xfs);
	oneStrplusDoubleq((UINT8*)countstr, cellXfsCount);
	oneStrwrite(e);

	while (cou< cellXfsNum) {
		oneStrwrite(stag);

		if (cx[cou]->numFmtId)
			oneStrplusDoubleq(num, cx[cou]->numFmtId);
		if (cx[cou]->fontId)
			oneStrplusDoubleq(fon, cx[cou]->fontId);
		if (cx[cou]->fillId)
			oneStrplusDoubleq(fil, cx[cou]->fillId);
		if (cx[cou]->borderId)
			oneStrplusDoubleq(bor, cx[cou]->borderId);
		if (cx[cou]->applyNumberFormat)
			oneStrplusDoubleq(an, cx[cou]->applyNumberFormat);
		if (cx[cou]->applyFont)
			oneStrplusDoubleq(afo, cx[cou]->applyFont);
		if (cx[cou]->applyFill)
			oneStrplusDoubleq(af, cx[cou]->applyFill);
		if (cx[cou]->applyBorder)
			oneStrplusDoubleq(ab, cx[cou]->applyBorder);
		if (cx[cou]->applyAlignment)
			oneStrplusDoubleq(aa, cx[cou]->applyAlignment);
		if (cx[cou]->xfId)
			oneStrplusDoubleq(xi, cx[cou]->xfId);

		if (cx[cou]->Avertical || cx[cou]->horizontal || cx[cou]->AwrapText) {
			oneStrwrite(e);// >�^�O
			oneStrwrite(align);
			if (cx[cou]->Avertical)
				oneStrplusDoubleq(ver, cx[cou]->Avertical);
			if (cx[cou]->horizontal)
				oneStrplusDoubleq(hor, cx[cou]->horizontal);
			if (cx[cou]->AwrapText)
				oneStrplusDoubleq(wt, cx[cou]->AwrapText);
			oneStrwrite(et);
		}
		else {
			oneStrwrite(et);// />�^�O
		}

		oneStrwrite(etag);

		cou++;
	}

	oneStrwrite((UINT8*)Exfs);

}

//�t�H���g��������
inline void StyleWrite::fontswrite()
{
	UINT8 stag[] = "<font>";
	UINT8 etag[] = "</font>";

	UINT8 sztag[] = "<sz";
	UINT8 coltag[] = "<color";
	UINT8 namtag[] = "<name";
	UINT8 famtag[] = "<family";
	UINT8 chatag[] = "<charset";
	UINT8 schtag[] = "<scheme";

	UINT8 valtag[] = " val=\"";
	UINT8 thetag[] = " theme=\"";
	UINT8 rgbtag[] = " rgb=\"";
	UINT8 indextag[] = " indexed=\"";

	UINT8 kf[] = " x14ac:knownFonts=\"";

	UINT8 closetag[] = ">";
	UINT8 et[] = "/>";

	Fonts** f = fontRoot;

	size_t cou = 0;

	oneStrwrite((UINT8*)fonts);
	oneStrplusDoubleq((UINT8*)countstr, fontCount);
	if (kFonts)
		oneStrplusDoubleq(kf, kFonts);
	oneStrwrite(closetag);

	while (cou< fontNum)
	{
		oneStrwrite(stag);

		// <b/> <i/>
		if (f[cou]->biu) {
			oneStrwrite(f[cou]->biu);
		}
		if (f[cou]->sz) {
			oneStrwrite(sztag);
			oneStrplusDoubleq(valtag, f[cou]->sz);
			oneStrwrite(et);
		}
		if (f[cou]->color) {
			oneStrwrite(coltag);
			oneStrplusDoubleq(thetag, f[cou]->color);
			oneStrwrite(et);
		}
		if (f[cou]->rgb) {
			oneStrwrite(coltag);
			oneStrplusDoubleq(rgbtag, f[cou]->rgb);
			oneStrwrite(et);
		}
		if (f[cou]->indexed) {
			oneStrwrite(coltag);
			oneStrplusDoubleq(indextag, f[cou]->indexed);
			oneStrwrite(et);
		}
		if (f[cou]->name) {
			oneStrwrite(namtag);
			oneStrplusDoubleq(valtag, f[cou]->name);
			oneStrwrite(et);
		}
		if (f[cou]->charset) {
			oneStrwrite(chatag);
			oneStrplusDoubleq(valtag, f[cou]->charset);
			oneStrwrite(et);
		}
		if (f[cou]->family) {
			oneStrwrite(famtag);
			oneStrplusDoubleq(valtag, f[cou]->family);
			oneStrwrite(et);
		}
		if (f[cou]->scheme) {
			oneStrwrite(schtag);
			oneStrplusDoubleq(valtag, f[cou]->scheme);
			oneStrwrite(et);
		}

		oneStrwrite(etag);

		cou++;
	}

	oneStrwrite((UINT8*)Efonts);

}

inline void StyleWrite::fillwrite()
{
	UINT8 cou[] = " count=\"";

	UINT8 et[] = "/>";
	UINT8 e[] = ">";

	UINT8 sz[] = "<sz";
	UINT8 col[] = "<color";
	UINT8 na[] = "<name";
	UINT8 fa[] = "<family";

	UINT8 stag[] = "<fill>";
	UINT8 etag[] = "</fill>";

	UINT8 pat[] = "<patternFill";
	UINT8 epat[] = "</patternFill>";
	UINT8 patT[] = " patternType=\"";

	UINT8 fg[] = "<fgColor";
	UINT8 rgb[] = " rgb=\"";
	UINT8 bg[] = "<bgColor";
	UINT8 theme[] = " theme=\"";
	UINT8 tint[] = " tint=\"";
	UINT8 ind[] = " indexed=\"";

	oneStrwrite((UINT8*)fills);
	oneStrplusDoubleq(cou, fillCount);
	oneStrwrite(e);

	Fills** fi = fillroot;
	size_t nu = 0;
	while (nu< fillNum)
	{
		oneStrwrite(stag);
		if (fi[nu]->patten) {
			oneStrwrite(pat);
			oneStrplusDoubleq(patT, fi[nu]->patten->patternType);
		}

		if (!fi[nu]->fg && !fi[nu]->bg) {
			oneStrwrite(et);
		}
		else {
			oneStrwrite(e);
		}

		if (fi[nu]->fg) {
			oneStrwrite(fg);
			if (fi[nu]->fg->rgb)
				oneStrplusDoubleq(rgb, fi[nu]->fg->rgb);
			if (fi[nu]->fg->theme)
				oneStrplusDoubleq(theme, fi[nu]->fg->theme);
			if (fi[nu]->fg->tint)
				oneStrplusDoubleq(tint, fi[nu]->fg->tint);
			oneStrwrite(et);
		}
		if (fi[nu]->bg) {
			oneStrwrite(bg);
			if (fi[nu]->bg->indexed)
				oneStrplusDoubleq(ind, fi[nu]->bg->indexed);
			oneStrwrite(et);
		}

		if (fi[nu]->fg || fi[nu]->bg) {
			oneStrwrite(epat);
		}

		oneStrwrite(etag);
		nu++;
	}

	oneStrwrite((UINT8*)endfil);

}

inline void StyleWrite::borderwrite()
{
	borders** bor = BorderRoot;

	UINT8 stag[] = "<border>";
	UINT8 etag[] = "</border>";

	UINT8 l[] = "<left";
	UINT8 r[] = "<right";
	UINT8 t[] = "<top";
	UINT8 b[] = "<bottom";
	UINT8 d[] = "<diagonal";

	UINT8 el[] = "</left>";
	UINT8 er[] = "</right>";
	UINT8 eto[] = "</top>";
	UINT8 eb[] = "</bottom>";
	UINT8 ed[] = "</diagonal>";

	UINT8 au[] = " auto=\"";
	UINT8 sty[] = " style=\"";
	UINT8 colo[] = "<color";
	UINT8 ind[] = " indexed=\"";
	UINT8 tin[] = " tint=\"";

	UINT8 et[] = "/>";
	UINT8 e[] = ">";

	oneStrwrite((UINT8*)border);
	oneStrplusDoubleq((UINT8*)countstr, borderCount);
	oneStrwrite(e);

	size_t co = 0;
	while (co< borderNum)
	{
		oneStrwrite(stag);
		if (bor[co]->left) {
			oneStrwrite(l);
			if (bor[co]->left->style) {
				oneStrplusDoubleq(sty, bor[co]->left->style);
				oneStrwrite(e);// >�^�O
			}

			if (bor[co]->left->Auto) {
				oneStrwrite(colo);
				oneStrplusDoubleq(au, bor[co]->left->Auto);
			}
			if (bor[co]->left->index) {
				oneStrwrite(colo);
				oneStrplusDoubleq(ind, bor[co]->left->index);
			}
			if (bor[co]->left->tint) {
				oneStrwrite(colo);
				oneStrplusDoubleq(tin, bor[co]->left->tint);
			}
			oneStrwrite(et);

			oneStrwrite(el);
		}
		else {
			oneStrwrite(l);
			oneStrwrite(et);
		}

		if (bor[co]->right) {
			oneStrwrite(r);
			if (bor[co]->right->style) {
				oneStrplusDoubleq(sty, bor[co]->right->style);
				oneStrwrite(e);// >�^�O
			}
			if (bor[co]->right->Auto) {
				oneStrwrite(colo);
				oneStrplusDoubleq(au, bor[co]->right->Auto);
			}
			if (bor[co]->right->index) {
				oneStrwrite(colo);
				oneStrplusDoubleq(ind, bor[co]->right->index);
			}
			if (bor[co]->right->tint) {
				oneStrwrite(colo);
				oneStrplusDoubleq(tin, bor[co]->right->tint);
			}
			oneStrwrite(et);

			oneStrwrite(er);
		}
		else {
			oneStrwrite(r);
			oneStrwrite(et);
		}

		if (bor[co]->bottom) {
			oneStrwrite(b);
			if (bor[co]->bottom->style) {
				oneStrplusDoubleq(sty, bor[co]->bottom->style);
				oneStrwrite(e);// >�^�O
			}

			if (bor[co]->bottom->Auto) {
				oneStrwrite(colo);
				oneStrplusDoubleq(au, bor[co]->bottom->Auto);
			}
			if (bor[co]->bottom->index) {
				oneStrwrite(colo);
				oneStrplusDoubleq(ind, bor[co]->bottom->index);
			}
			if (bor[co]->bottom->tint) {
				oneStrwrite(colo);
				oneStrplusDoubleq(tin, bor[co]->bottom->tint);
			}
			oneStrwrite(et);
			oneStrwrite(eb);
		}
		else {
			oneStrwrite(b);
			oneStrwrite(et);
		}

		if (bor[co]->top) {
			oneStrwrite(t);
			if (bor[co]->top->style) {
				oneStrplusDoubleq(sty, bor[co]->top->style);
				oneStrwrite(e);// >�^�O
			}
			if (bor[co]->top->Auto) {
				oneStrwrite(colo);
				oneStrplusDoubleq(au, bor[co]->top->Auto);
			}
			if (bor[co]->top->index) {
				oneStrwrite(colo);
				oneStrplusDoubleq(ind, bor[co]->top->index);
			}
			if (bor[co]->top->tint) {
				oneStrwrite(colo);
				oneStrplusDoubleq(tin, bor[co]->top->tint);
			}
			oneStrwrite(et);

			oneStrwrite(eto);
		}
		else {
			oneStrwrite(t);
			oneStrwrite(et);
		}

		if (bor[co]->diagonal) {
			oneStrwrite(d);
			if (bor[co]->diagonal->style) {
				oneStrplusDoubleq(sty, bor[co]->diagonal->style);
				oneStrwrite(e);// >�^�O
			}
			if (bor[co]->diagonal->Auto) {
				oneStrwrite(colo);
				oneStrplusDoubleq(au, bor[co]->diagonal->Auto);
			}
			if (bor[co]->diagonal->index) {
				oneStrwrite(colo);
				oneStrplusDoubleq(ind, bor[co]->diagonal->index);
			}
			if (bor[co]->diagonal->tint) {
				oneStrwrite(colo);
				oneStrplusDoubleq(tin, bor[co]->diagonal->tint);
			}
			oneStrwrite(et);

			oneStrwrite(ed);
		}
		else {
			oneStrwrite(d);
			oneStrwrite(et);
		}
		oneStrwrite(etag);
		co++;
	}

	oneStrwrite((UINT8*)enbor);

}

inline void StyleWrite::cellstyleXfs()
{
	UINT8 et[] = "/>";
	UINT8 e[] = ">";

	UINT8 stag[] = "<xf";
	UINT8 etag[] = "</xf>";

	UINT8 num[] = " numFmtId=\"";
	UINT8 fon[] = " fontId=\"";
	UINT8 fil[] = " fillId=\"";
	UINT8 bor[] = " borderId=\"";
	UINT8 an[] = " applyNumberFormat=\"";
	UINT8 ab[] = " applyBorder=\"";
	UINT8 aa[] = " applyAlignment=\"";
	UINT8 ap[] = " applyProtection=\"";
	UINT8 af[] = " applyFill=\"";
	UINT8 align[] = "<alignment";
	UINT8 ver[] = " vertical=\"";
	UINT8 afo[] = " applyFont=\"";

	stylexf** csx = cellstyleXfsRoot;
	size_t co = 0;

	oneStrwrite((UINT8*)cellstXfs);
	oneStrplusDoubleq((UINT8*)countstr, cellStyleXfsCount);
	oneStrwrite(e);

	while (csx)
	{
		oneStrwrite(stag);

		if (csx[co]->numFmtId)
			oneStrplusDoubleq(num, csx[co]->numFmtId);
		if (csx[co]->fontId)
			oneStrplusDoubleq(fon, csx[co]->fontId);
		if (csx[co]->fillId)
			oneStrplusDoubleq(fil, csx[co]->fillId);
		if (csx[co]->borderId)
			oneStrplusDoubleq(bor, csx[co]->borderId);
		if (csx[co]->applyNumberFormat)
			oneStrplusDoubleq(an, csx[co]->applyNumberFormat);
		if (csx[co]->applyFont)
			oneStrplusDoubleq(afo, csx[co]->applyFont);
		if (csx[co]->applyFill)
			oneStrplusDoubleq(af, csx[co]->applyFill);
		if (csx[co]->applyBorder)
			oneStrplusDoubleq(ab, csx[co]->applyBorder);
		if (csx[co]->applyAlignment)
			oneStrplusDoubleq(aa, csx[co]->applyAlignment);
		if (csx[co]->applyProtection)
			oneStrplusDoubleq(ap, csx[co]->applyProtection);

		if (csx[co]->Avertical) {
			oneStrwrite(e);// >�^�O
			oneStrwrite(align);
			oneStrplusDoubleq(ver, csx[co]->Avertical);
			oneStrwrite(et);
		}
		else {
			oneStrwrite(et);// />�^�O
		}

		oneStrwrite(etag);

		co++;
	}

	oneStrwrite((UINT8*)EncsXfs);

}

//numFmt ��������
inline void StyleWrite::numFmidwrite()
{
	UINT8 st[] = "<numFmt";
	UINT8 id[] = " numFmtId=\"";
	UINT8 code[] = " formatCode=\"";
	UINT8 cou[] = " count=\"";
	UINT8 et[] = "/>";
	UINT8 e[] = ">";

	oneStrwrite((UINT8*)Fmts);
	oneStrplusDoubleq(cou, numFmtsCount);
	oneStrwrite((UINT8*)e);

	numFMts** n = numFmtsRoot;
	size_t co = 0;

	while (co< numFmtsNum) {
		if (n[co])
			std::cout << "numfmts : " << std::endl;
		oneStrwrite(st);
		if (n[co]->Id) {
			std::cout << n[co]->Id;
			oneStrplusDoubleq(id, n[co]->Id);
		}
		if (n[co]->Code)
			oneStrplusDoubleq(code, n[co]->Code);
		oneStrwrite(et);

		co++;
	}

	oneStrwrite((UINT8*)Efmts);
}

//xx="~" ��������
inline void StyleWrite::oneStrplusDoubleq(UINT8* str, UINT8* v) {
	int i = 0;
	char d = '"';

	while (str[i] != '\0') {
		wd[wdlen] = str[i];
		i++;
		wdlen++;
	}
	i = 0;
	while (v[i] != '\0') {
		wd[wdlen] = v[i];
		i++;
		wdlen++;
	}

	wd[wdlen] = d;
	wdlen++;
}

//tag��������
inline void StyleWrite::oneStrwrite(UINT8* str) {
	int i = 0;

	while (str[i] != '\0') {
		wd[wdlen] = str[i];
		i++;
		wdlen++;
	}
}

inline void StyleWrite::writedata(UINT8* stag, UINT8* v, UINT8* etag)
{
	int i = 0;

	while (stag[i] != '\0') {
		wd[wdlen] = stag[i];
		i++;
		wdlen++;
	}
	i = 0;
	while (v[i] != '\0') {
		wd[wdlen] = v[i];
		i++;
		wdlen++;
	}
	i = 0;
	while (etag[i] != '\0') {
		wd[wdlen] = etag[i];
		i++;
		wdlen++;
	}
}


