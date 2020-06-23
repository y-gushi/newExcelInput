#pragma once

#include "excel_style.h"

class StyleWrite:public styleread
{
public:

	StyleWrite(FILE* f);
	~StyleWrite();

	FILE* fr;

	void stylewrite();
	void writedata(UINT8* stag, UINT8* v, UINT8* etag);

	void checkstyle(UINT8* num);

private:
	
};

StyleWrite::StyleWrite(FILE* f)
{
	//**fonts = { "11", "FF006100","‚l‚r ‚oƒSƒVƒbƒN", "2","128","minor" };
	
	fr = f;

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

	styleSheetStr = nullptr;
	extLstStr = nullptr;
}

StyleWrite::~StyleWrite()
{
	freefonts(fontRoot);
	freefill(fillroot);
	freeborder(BorderRoot);
	freestylexf(cellstyleXfsRoot);
	freecellxfs(cellXfsRoot);
	freenumFMts(numFmtsRoot);
	freecellstyle(cellStyleRoot);
	freeDxf(dxfRoot);

	free(styleSheetStr);
	free(extLstStr);
}

inline void StyleWrite::stylewrite()
{
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

inline void StyleWrite::checkstyle(UINT8* num)
{
	/*-----------------------------------
	bee styleİ’è
	-----------------------------------*/
	//font sz rgb name family charset scheme
	const char* fonts[] = { "11", "FF006100","‚l‚r ‚oƒSƒVƒbƒN", "2","128","minor" };
	//patternType fgrgb
	const char* fill[] = { "solid","FFC6EFCE" };
	//border all null
	//numFmtId
	//code
	const char numFmCode[] = "[$$-409]#,##0.00;[$$-409]#,##0.00";

	//xfId
	//applyBorder applyAlignment
	const char* xfids[] = { "0","0","center" };

	/*-----------------------------------
	shoplist styleİ’è
	-----------------------------------*/
	//font sz theme name family charset scheme
	const char* fonts[] = { "11", "0","‚l‚r ‚oƒSƒVƒbƒN", "2","128","minor" };
	//patternType fgtheme
	const char* fill[] = { "solid","4" };
	//border all null
	//numFmtId
	//code
	const char numFmId[] = "0";

	//xfId
	//applyBorder applyAlignment
	const char* xfids[] = { "0","0","center" };


	ArrayNumber* changeStr = new ArrayNumber;

	UINT32 stylenum = changeStr->RowArraytoNum(num, strlen((const char*)num));//style ”Ô†@”š‚É•ÏŠ·

	int stylec = 0;
	cellxfs* cxf = cellXfsRoot;

	while (stylec < stylenum && cxf) {
		cxf = cxf->next;
		stylec++;
	}
}
