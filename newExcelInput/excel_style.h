#pragma once
#include "typechange.h"
#include <stdlib.h>
#include <string.h>


struct Fonts {
	UINT8* sz;
	UINT8* color;
	UINT8* name;
	UINT8* family;
	UINT8* charset;
	UINT8* scheme;
	Fonts* next;
};

struct Fills
{
	UINT8* type;
	UINT8* fg;
	UINT8* bg;
	Fills* next;
};
//ボーダー内容
struct borderstyle
{
	UINT8* style;
	UINT8* theme;
	UINT8* tint;
	UINT8* rgb;
};

struct borders
{
	borderstyle* left;
	borderstyle* right;
	borderstyle* top;
	borderstyle* bottom;
	borderstyle* diagonal;
	borders* next;
};

//cellStyleXfs 内容
struct stylexf
{
	UINT8* numFmtId;
	UINT8* fontId;
	UINT8* fillId;
	UINT8* borderId;
	UINT8* applyNumberFormat;
	UINT8* applyBorder;
	UINT8* applyAlignment;
	UINT8* applyProtection;
	UINT8* Avertical;
	stylexf* next;
};

struct cellxfs
{
	UINT8* numFmtId;
	UINT8* fontId;
	UINT8* fillId;
	UINT8* borderId;
	UINT8* xfId;
	UINT8* applyNumberFormat;
	UINT8* applyFont;
	UINT8* applyFill;
	UINT8* applyBorder;
	UINT8* applyAlignment;
	UINT8* Avertical;
	UINT8* AwrapText;
	cellxfs* next;
};

struct cellstyle
{
	UINT8* name;
	UINT8* xfId;
	UINT8* builtinId;
	UINT8* customBuilt;
	UINT8* xruid;
	cellstyle* next;
};

struct Dxf {
	UINT8* color;
	UINT8* bgcolor;
	Dxf* next;
};

struct colors
{
	UINT8* color;
	colors* next;
};

struct numFMts
{
	UINT8* Id;
	UINT8* Code;
	numFMts* next;
};

class styleread {
public:
	const char fonts[7] = "<fonts";
	const char Efont[9] = "</fonts>";
	const char fills[7] = "<fills";
	const char endfil[9] = "</fills>";
	const char border[9] = "<borders";
	const char enbor[11] = "</borders>";
	const char cellstXfs[14] = "<cellStyleXfs";
	const char EncsXfs[16] = "</cellStyleXfs>";
	const char color[9] = "<colors>";
	const char Ecolor[10] = "</colors>";
	const char Xfs[9] = "<cellXfs";
	const char Exfs[11] = "</cellXfs>";
	const char CellStyl[12] = "<cellStyles";
	const char Ecellsty[14] = "</cellStyles>";
	const char dxf[6] = "<dxfs";
	const char Edxf[8] = "</dxfs>";
	const char Endstyle[14] = "</styleSheet>";
	const char Fmts[9] = "<numFmts";
	const char Efmts[11] = "</numFmts>";

	int fontcou;

	UINT64 readp;

	styleread();
	~styleread();

	Fonts* addfonts(Fonts* f, UINT8* sz, UINT8* co, UINT8* na, UINT8* fa, UINT8* cha, UINT8* sch);

	Fills* addfill(Fills* fi, UINT8* t, UINT8* f, UINT8* b);

	borders* addborder(borders* b, borderstyle* l, borderstyle* r, borderstyle* t, borderstyle* bo, borderstyle* d);

	stylexf* addstylexf(stylexf* sx, UINT8* n, UINT8* fo, UINT8* fi, UINT8* bi, UINT8* an, UINT8* ab, UINT8* aa, UINT8* ap, UINT8* av);

	cellxfs* addcellxfs(cellxfs* c, UINT8* n, UINT8* fo, UINT8* fi, UINT8* bi, UINT8* an, UINT8* ab, UINT8* aa, UINT8* av, UINT8* at);

	cellstyle* addcellstyle(cellstyle* c, UINT8* n, UINT8* x, UINT8* b, UINT8* cu, UINT8* xr);

	Dxf* adddxf(Dxf* d, UINT8* c, UINT8* bc);

	colors* addcolor(colors* c, UINT8* co);

	numFMts* addnumFmts(numFMts* n, UINT8* i, UINT8* c);

	UINT8* readfonts(UINT8* sd, UINT64 len, UINT64 rp);

	void readstyle(UINT8* sdata,UINT64 sdatalen);
};