#pragma once

#include "excel_style.h"



class StyleWrite:public styleread
{
public:

	StyleWrite();
	~StyleWrite();

	void styledatawrite();

	void Dxfswrite();

	void writeColors();

	void writefirst();

	void writeextLst();

	void CELLStyle();

	FILE* fr;

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
	
}

StyleWrite::~StyleWrite()
{
}

inline void StyleWrite::styledatawrite() {
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

	Dxf* dx = dxfRoot;

	oneStrwrite((UINT8*)dxf);
	oneStrplusDoubleq((UINT8*)countstr, dxfsCount);
	oneStrwrite(e);

	while (dx) {
		oneStrwrite(stag);

		if (dx->fon) {
			oneStrwrite(fstag);

			if (dx->fon->b)
				oneStrwrite(dx->fon->b);
			if (dx->fon->ival) {
				oneStrwrite(istr);
				oneStrplusDoubleq(val, dx->fon->ival);
				oneStrwrite(et);
			}

			if (dx->fon->rgb)
			{
				oneStrwrite(col);
				oneStrplusDoubleq(rgb, dx->fon->rgb);
				oneStrwrite(et);
			}

			oneStrwrite(fetag);
		}
		if (dx->fil) {
			oneStrwrite(sfil);

			if (dx->fil->rgb) {
				oneStrwrite(spat);
				oneStrwrite(bgc);
				oneStrplusDoubleq(rgb, dx->fil->rgb);
				oneStrwrite(et);
				oneStrwrite(epat);
			}

			oneStrwrite(efil);
		}

		oneStrwrite(etag);

		dx = dx->next;
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
//write 最初の文字列
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

	cellstyle* cs = cellStyleRoot;

	oneStrwrite((UINT8*)CellStyl);
	oneStrplusDoubleq((UINT8*)countstr, cellstyleCount);
	oneStrwrite(e);

	while (cs) {
		oneStrwrite(stag);

		if (cs->name)
			oneStrplusDoubleq(nam, cs->name);
		if (cs->xfId)
			oneStrplusDoubleq(xfid, cs->xfId);
		if (cs->xruid)
			oneStrplusDoubleq(xruid, cs->xruid);
		if (cs->builtinId)
			oneStrplusDoubleq(bui, cs->builtinId);
		if (cs->customBuilt)
			oneStrplusDoubleq(cb, cs->customBuilt);
		
		oneStrwrite(et);// />タグ

		cs = cs->next;
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

	cellxfs* cx = cellXfsRoot;

	oneStrwrite((UINT8*)Xfs);
	oneStrplusDoubleq((UINT8*)countstr, cellXfsCount);
	oneStrwrite(e);

	while (cx) {
		oneStrwrite(stag);

		if (cx->numFmtId)
			oneStrplusDoubleq(num, cx->numFmtId);
		if (cx->fontId)
			oneStrplusDoubleq(fon, cx->fontId);
		if (cx->fillId)
			oneStrplusDoubleq(fil, cx->fillId);
		if (cx->borderId)
			oneStrplusDoubleq(bor, cx->borderId);
		if (cx->applyNumberFormat)
			oneStrplusDoubleq(an, cx->applyNumberFormat);
		if (cx->applyFont)
			oneStrplusDoubleq(afo, cx->applyFont);
		if (cx->applyFill)
			oneStrplusDoubleq(af, cx->applyFill);
		if (cx->applyBorder)
			oneStrplusDoubleq(ab, cx->applyBorder);
		if (cx->applyAlignment)
			oneStrplusDoubleq(aa, cx->applyAlignment);
		if (cx->xfId)
			oneStrplusDoubleq(xi, cx->xfId);

		if (cx->Avertical || cx->horizontal||cx->AwrapText) {
			oneStrwrite(e);// >タグ
			oneStrwrite(align);
			if (cx->Avertical)
				oneStrplusDoubleq(ver, cx->Avertical);
			if(cx->horizontal)
				oneStrplusDoubleq(hor, cx->horizontal);
			if (cx->AwrapText)
				oneStrplusDoubleq(wt, cx->AwrapText);
			oneStrwrite(et);
		}
		else {
			oneStrwrite(et);// />タグ
		}

		oneStrwrite(etag);

		cx = cx->next;
	}

	oneStrwrite((UINT8*)Exfs);

}

//フォント書き込み
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

	Fonts* f = fontRoot;

	oneStrwrite((UINT8*)fonts);
	oneStrplusDoubleq((UINT8*)countstr, fontCount);
	if (kFonts)
		oneStrplusDoubleq(kf, kFonts);
	oneStrwrite(closetag);

	while (f)
	{
		oneStrwrite(stag);

		// <b/> <i/>
		if (f->biu) {
			oneStrwrite(f->biu);
		}
		if (f->sz) {
			oneStrwrite(sztag);
			oneStrplusDoubleq(valtag, f->sz);
			oneStrwrite(et);
		}
		if (f->color) {
			oneStrwrite(coltag);
			oneStrplusDoubleq(thetag, f->color);
			oneStrwrite(et);
		}
		if (f->rgb) {
			oneStrwrite(coltag);
			oneStrplusDoubleq(rgbtag, f->rgb);
			oneStrwrite(et);
		}
		if (f->indexed) {
			oneStrwrite(coltag);
			oneStrplusDoubleq(indextag, f->indexed);
			oneStrwrite(et);
		}
		if (f->name) {
			oneStrwrite(namtag);
			oneStrplusDoubleq(valtag, f->name);
			oneStrwrite(et);
		}
		if (f->charset) {
			oneStrwrite(chatag);
			oneStrplusDoubleq(valtag, f->charset);
			oneStrwrite(et);
		}
		if (f->family) {
			oneStrwrite(famtag);
			oneStrplusDoubleq(valtag, f->family);
			oneStrwrite(et);
		}
		if (f->scheme) {
			oneStrwrite(schtag);
			oneStrplusDoubleq(valtag, f->scheme);
			oneStrwrite(et);
		}

		oneStrwrite(etag);

		f = f->next;
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
	UINT8 patT[]=" patternType=\"";

	UINT8 fg[] = "<fgColor";
	UINT8 rgb[]=" rgb=\"";
	UINT8 bg[] = "<bgColor";
	UINT8 theme[] = " theme=\"";
	UINT8 tint[] = " tint=\"";
	UINT8 ind[] = " indexed=\"";

	oneStrwrite((UINT8*)fills);
	oneStrplusDoubleq(cou, fillCount);
	oneStrwrite(e);

	Fills* fi = fillroot;
	while (fi)
	{
		oneStrwrite(stag);
		if (fi->patten) {
			oneStrwrite(pat);
			oneStrplusDoubleq(patT, fi->patten->patternType);
		}
		
		if(!fi->fg && !fi->bg) {
			oneStrwrite(et);
		}
		else {
			oneStrwrite(e);
		}

		if (fi->fg) {
			oneStrwrite(fg);
			if (fi->fg->rgb)
				oneStrplusDoubleq(rgb, fi->fg->rgb);
			if (fi->fg->theme)
				oneStrplusDoubleq(theme, fi->fg->theme);
			if (fi->fg->tint)
				oneStrplusDoubleq(tint, fi->fg->tint);
			oneStrwrite(et);
		}
		if (fi->bg) {
			oneStrwrite(bg);
			if (fi->bg->indexed)
				oneStrplusDoubleq(ind, fi->bg->indexed);
			oneStrwrite(et);
		}
		
		if (fi->fg || fi->bg) {
			oneStrwrite(epat);
		}

		oneStrwrite(etag);
		fi = fi->next;
	}

	oneStrwrite((UINT8*)endfil);

}

inline void StyleWrite::borderwrite()
{
	borders* bor = BorderRoot;

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

	while (bor)
	{
		oneStrwrite(stag);
		if (bor->left) {
			oneStrwrite(l);
			if (bor->left->style) {
				oneStrplusDoubleq(sty, bor->left->style);
				oneStrwrite(e);// >タグ
			}

			if (bor->left->Auto) {
				oneStrwrite(colo);
				oneStrplusDoubleq(au, bor->left->Auto);
			}
			if (bor->left->index) {
				oneStrwrite(colo);
				oneStrplusDoubleq(ind, bor->left->index);
			}
			if (bor->left->tint) {
				oneStrwrite(colo);
				oneStrplusDoubleq(tin, bor->left->tint);
			}
			oneStrwrite(et);

			oneStrwrite(el);
		}
		else {
			oneStrwrite(l);
			oneStrwrite(et);
		}

		if (bor->right) {
			oneStrwrite(r);
			if (bor->right->style) {
				oneStrplusDoubleq(sty, bor->right->style);
				oneStrwrite(e);// >タグ
			}
			if (bor->right->Auto) {
				oneStrwrite(colo);
				oneStrplusDoubleq(au, bor->right->Auto);
			}
			if (bor->right->index) {
				oneStrwrite(colo);
				oneStrplusDoubleq(ind, bor->right->index);
			}
			if (bor->right->tint) {
				oneStrwrite(colo);
				oneStrplusDoubleq(tin, bor->right->tint);
			}
			oneStrwrite(et);

			oneStrwrite(er);
		}
		else {
			oneStrwrite(r);
			oneStrwrite(et);
		}

		if (bor->bottom) {
			oneStrwrite(b);
			if (bor->bottom->style) {
				oneStrplusDoubleq(sty, bor->bottom->style);
				oneStrwrite(e);// >タグ
			}

			if (bor->bottom->Auto) {
				oneStrwrite(colo);
				oneStrplusDoubleq(au, bor->bottom->Auto);
			}
			if (bor->bottom->index) {
				oneStrwrite(colo);
				oneStrplusDoubleq(ind, bor->bottom->index);
			}
			if (bor->bottom->tint) {
				oneStrwrite(colo);
				oneStrplusDoubleq(tin, bor->bottom->tint);
			}
			oneStrwrite(et);
			oneStrwrite(eb);
		}
		else {
			oneStrwrite(b);
			oneStrwrite(et);
		}

		if (bor->top) {
			oneStrwrite(t);
			if (bor->top->style) {
				oneStrplusDoubleq(sty, bor->top->style);
				oneStrwrite(e);// >タグ
			}
			if (bor->top->Auto) {
				oneStrwrite(colo);
				oneStrplusDoubleq(au, bor->top->Auto);
			}
			if (bor->top->index) {
				oneStrwrite(colo);
				oneStrplusDoubleq(ind, bor->top->index);
			}
			if (bor->top->tint) {
				oneStrwrite(colo);
				oneStrplusDoubleq(tin, bor->top->tint);
			}
			oneStrwrite(et);

			oneStrwrite(eto);
		}
		else {
			oneStrwrite(t);
			oneStrwrite(et);
		}

		if (bor->diagonal) {
			oneStrwrite(d);
			if (bor->diagonal->style) {
				oneStrplusDoubleq(sty, bor->diagonal->style);
				oneStrwrite(e);// >タグ
			}
			if (bor->diagonal->Auto) {
				oneStrwrite(colo);
				oneStrplusDoubleq(au, bor->diagonal->Auto);
			}
			if (bor->diagonal->index) {
				oneStrwrite(colo);
				oneStrplusDoubleq(ind, bor->diagonal->index);
			}
			if (bor->diagonal->tint) {
				oneStrwrite(colo);
				oneStrplusDoubleq(tin, bor->diagonal->tint);
			}
			oneStrwrite(et);

			oneStrwrite(ed);
		}
		else {
			oneStrwrite(d);
			oneStrwrite(et);
		}
		oneStrwrite(etag);
		bor = bor->next;
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

	stylexf* csx = cellstyleXfsRoot;

	oneStrwrite((UINT8*)cellstXfs);
	oneStrplusDoubleq((UINT8*)countstr, cellStyleXfsCount);
	oneStrwrite(e);

	while (csx)
	{
		oneStrwrite(stag);

		if (csx->numFmtId)
			oneStrplusDoubleq(num, csx->numFmtId);
		if (csx->fontId)
			oneStrplusDoubleq(fon, csx->fontId);
		if (csx->fillId)
			oneStrplusDoubleq(fil, csx->fillId);
		if (csx->borderId)
			oneStrplusDoubleq(bor, csx->borderId);
		if(csx->applyNumberFormat)
			oneStrplusDoubleq(an, csx->applyNumberFormat);
		if (csx->applyFont)
			oneStrplusDoubleq(afo, csx->applyFont);
		if (csx->applyFill)
			oneStrplusDoubleq(af, csx->applyFill);
		if (csx->applyBorder)
			oneStrplusDoubleq(ab, csx->applyBorder);
		if (csx->applyAlignment)
			oneStrplusDoubleq(aa, csx->applyAlignment);
		if (csx->applyProtection)
			oneStrplusDoubleq(ap, csx->applyProtection);

		if (csx->Avertical) {
			oneStrwrite(e);// >タグ
			oneStrwrite(align);
			oneStrplusDoubleq(ver, csx->Avertical);
			oneStrwrite(et);
		}
		else {
			oneStrwrite(et);// />タグ
		}

		oneStrwrite(etag);

		csx = csx->next;
	}

	oneStrwrite((UINT8*)EncsXfs);

}

//numFmt 書き込み
inline void StyleWrite::numFmidwrite()
{
	UINT8 st[] = "<numFmt";
	UINT8 id[]=" numFmtId=\"";
	UINT8 code[] = " formatCode=\"";
	UINT8 cou[] = " count=\"";
	UINT8 et[] = "/>";
	UINT8 e[] = ">";

	oneStrwrite((UINT8*)Fmts);
	oneStrplusDoubleq(cou, numFmtsCount);
	oneStrwrite((UINT8*)e);

	numFMts* n = numFmtsRoot;

	while (n) {
		if(n)
			std::cout << "numfmts : " << std::endl;
		oneStrwrite(st);
		if (n->Id) {
			std::cout << n->Id;
			oneStrplusDoubleq(id, n->Id);
		}
		if (n->Code)
			oneStrplusDoubleq(code, n->Code);
		oneStrwrite(et);

		n = n->next;
	}

	oneStrwrite((UINT8*)Efmts);
}

//xx="~" 書き込み
inline void StyleWrite::oneStrplusDoubleq(UINT8* str,UINT8* v) {
	int i = 0;
	char d = '"';

	while (str[i] != '\0') {
		fwrite(&str[i], 1, 1, fr);
		i++;
	}
	i = 0;
	while (v[i] != '\0') {
		fwrite(&v[i], 1, 1, fr);
		i++;
	}

	fwrite(&d, 1, 1, fr);
}

//tag書き込み
inline void StyleWrite::oneStrwrite(UINT8* str) {
	int i = 0;

	while (str[i] != '\0') {
		fwrite(&str[i], 1, 1, fr);
		i++;
	}
}

inline void StyleWrite::writedata(UINT8* stag, UINT8* v, UINT8* etag)
{
	int i = 0;

	while (stag[i] != '\0') {
		fwrite(&stag[i], 1, 1, fr);
		i++;
	}
	i = 0;
	while (v[i] != '\0') {
		fwrite(&v[i], 1, 1, fr);
		i++;
	}
	i = 0;
	while (etag[i] != '\0') {
		fwrite(&etag[i], 1, 1, fr);
		i++;
	}
}


