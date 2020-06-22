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

/*fill���e--------------*/

struct FillPattern
{
	UINT8* patternType;
};

struct fgcolor
{
	UINT8* theme;
	UINT8* tint;
};

struct bgcolor
{
	UINT8* indexed;
};

struct Fills
{
	FillPattern* patten;
	fgcolor* fg;
	bgcolor* bg;
	Fills* next;
};
/*-----------------------*/

//�{�[�_�[���e
struct borderstyle
{
	UINT8* style;
	UINT8* theme;
	UINT8* tint;
	UINT8* rgb;
	UINT8* index;
	UINT8* Auto;
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

//cellStyleXfs ���e
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
	UINT8* horizontal;
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
	const char Efonts[9] = "</fonts>";
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

	Fonts* fontRoot;
	Fills* fillroot;
	borders* BorderRoot;
	stylexf* cellstyleXfsRoot;
	cellxfs* cellXfsRoot;

	UINT8* fontCount;
	UINT8* fillCount;
	UINT8* borderCount;
	UINT8* cellStyleXfsCount;
	UINT8* cellXfsCount;

	styleread();
	~styleread();

	Fonts* addfonts(Fonts* f, UINT8* sz, UINT8* co, UINT8* na, UINT8* fa, UINT8* cha, UINT8* sch);

	Fills* addfill(Fills* fi, FillPattern* p, fgcolor* f, bgcolor* b);

	borders* addborder(borders* b, borderstyle* l, borderstyle* r, borderstyle* t, borderstyle* bo, borderstyle* d);

	stylexf* addstylexf(stylexf* sx, UINT8* n, UINT8* fo, UINT8* fi, UINT8* bi, UINT8* an, UINT8* ab, UINT8* aa, UINT8* ap, UINT8* av);

	cellxfs* addcellxfs(cellxfs* c, UINT8* n, UINT8* fo, UINT8* fi, UINT8* bi, UINT8* an, UINT8* ab, UINT8* aa, UINT8* av, UINT8* at, UINT8* ho, UINT8* afo, UINT8* afi, UINT8* xid);

	cellstyle* addcellstyle(cellstyle* c, UINT8* n, UINT8* x, UINT8* b, UINT8* cu, UINT8* xr);

	Dxf* adddxf(Dxf* d, UINT8* c, UINT8* bc);

	colors* addcolor(colors* c, UINT8* co);

	numFMts* addnumFmts(numFMts* n, UINT8* i, UINT8* c);

	void readfontV(UINT8* dat);

	UINT8* readfonts(UINT8* sd);

	void getfillV(UINT8* d);

	UINT8* readfill(UINT8* sd);

	fgcolor* readfillFg(UINT8* dat);

	bgcolor* readbgcolor(UINT8* dat);

	borderstyle* getborV(UINT8* dat, UINT8* endT, size_t endTlen);

	void getBorder(UINT8* d);

	UINT8* readboerder(UINT8* d);

	UINT8* getValue(UINT8* d);

	void getCellStyleXfsV(UINT8* d);

	void getCellXfsV(UINT8* d);

	void readCellStyleXfs(UINT8* d);

	void readCellXfs(UINT8* d);

	void readstyle(UINT8* sdata,UINT64 sdatalen);
};