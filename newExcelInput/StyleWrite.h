#pragma once

#include "excel_style.h"



class StyleWrite:public styleread
{
public:

	StyleWrite();
	~StyleWrite();

	FILE* fr;

	void cellXfswrite();
	void fontswrite();
	void fillwrite();
	void borderwrite();
	void cellstyleXfs();
	void numFmidwrite();

	void writedata(UINT8* stag, UINT8* v, UINT8* etag);

private:
	
};

StyleWrite::StyleWrite()
{
	
}

StyleWrite::~StyleWrite()
{
}

//write cellXfs
inline void StyleWrite::cellXfswrite()
{
	while (cellXfsCount) {
		cellXfsRoot = cellXfsRoot->next;
	}

	if (cellXfsRoot->numFmtId) {
		std::cout << "numFmtId : " <<cellXfsRoot->numFmtId << std::endl;
	}
	if (cellXfsRoot->fontId) {
		std::cout << "fontId : " <<cellXfsRoot->fontId << std::endl;
	}
	if (cellXfsRoot->fillId)
		std::cout << "fillId : " <<cellXfsRoot->fillId << std::endl;
	if (cellXfsRoot->borderId)
		std::cout << "borderId : " <<cellXfsRoot->borderId << std::endl;
	if (cellXfsRoot->xfId)
		std::cout << "xfId : " <<cellXfsRoot->xfId << std::endl;
	if (cellXfsRoot->applyNumberFormat)
		std::cout << "applyNumberFormat : " <<cellXfsRoot->applyNumberFormat << std::endl;
	if (cellXfsRoot->applyFont)
		std::cout << "applyFont : " <<cellXfsRoot->applyFont << std::endl;
	if (cellXfsRoot->applyFill)
		std::cout << "applyFill : " <<cellXfsRoot->applyFill << std::endl;
	if (cellXfsRoot->applyBorder)
		std::cout << "applyBorder : " <<cellXfsRoot->applyBorder << std::endl;
	if (cellXfsRoot->applyAlignment)
		std::cout << "applyAlignment : " <<cellXfsRoot->applyAlignment << std::endl;
	if (cellXfsRoot->Avertical)
		std::cout << "vertical : " <<cellXfsRoot->Avertical << std::endl;
	if (cellXfsRoot->horizontal)
		std::cout << "horizontal : " <<cellXfsRoot->horizontal << std::endl;
	if (cellXfsRoot->AwrapText)
		std::cout << "wrapText : " <<cellXfsRoot->AwrapText << std::endl;
}

inline void StyleWrite::fontswrite()
{
	while (fontRoot)
	{
		fontRoot = fontRoot->next;
	}
	if (fontRoot->sz)
		std::cout << "sz : " << fontRoot->sz << std::endl;
	if (fontRoot->color)
		std::cout << "theme : " << fontRoot->color << std::endl;
	if (fontRoot->rgb)
		std::cout << "rgb : " << fontRoot->rgb << std::endl;
	if (fontRoot->name)
		std::cout << "name : " << fontRoot->name << std::endl;
	if (fontRoot->family)
		std::cout << "family : " << fontRoot->family << std::endl;
	if (fontRoot->charset)
		std::cout << "charset : " << fontRoot->charset << std::endl;
	if (fontRoot->scheme)
		std::cout << "scheme : " << fontRoot->scheme << std::endl;
	if (fontRoot->rgb)
		std::cout << "rgb : " << fontRoot->rgb << std::endl;
}

inline void StyleWrite::fillwrite()
{
	while (fillroot)
	{
		fillroot = fillroot->next;
	}
	if (fillroot->patten)
		std::cout << "patternType : " << fillroot->patten->patternType << std::endl;
	if (fillroot->fg) {
		if (fillroot->fg->rgb)
			std::cout << "fgColor rgb : " << fillroot->fg->rgb << std::endl;
		if (fillroot->fg->theme)
			std::cout << "fgColor theme : " << fillroot->fg->theme << std::endl;
		if (fillroot->fg->tint)
			std::cout << "fgColor tint : " << fillroot->fg->tint << std::endl;
	}
	if (fillroot->bg) {
		if (fillroot->bg->indexed)
			std::cout << "bgColor indexed : " << fillroot->bg->indexed << std::endl;
	}
}

inline void StyleWrite::borderwrite()
{
	while (BorderRoot)
	{
		BorderRoot = BorderRoot->next;
	}
	if (BorderRoot->left)
		std::cout << "left style : " << BorderRoot->left->style << std::endl;
	else
		std::cout << "left 何もなし " << std::endl;
	if (BorderRoot->right)
		std::cout << "right style : " << BorderRoot->right->style << std::endl;
	else
		std::cout << "right 何もなし " << std::endl;
	if (BorderRoot->bottom)
		std::cout << "bottom style : " << BorderRoot->top->style << std::endl;
	else
		std::cout << "bottom 何もなし " << std::endl;
	if (BorderRoot->top)
		std::cout << "top style : " << BorderRoot->top->style << std::endl;
	else
		std::cout << "top 何もなし " << std::endl;
}

inline void StyleWrite::cellstyleXfs()
{
	while (cellstyleXfsRoot)
	{
		cellstyleXfsRoot = cellstyleXfsRoot->next;
	}
	if (cellstyleXfsRoot->numFmtId)
		std::cout << "numFmtId : " << cellstyleXfsRoot->numFmtId << std::endl;
	if (cellstyleXfsRoot->fontId)
		std::cout << "fontId : " << cellstyleXfsRoot->fontId << std::endl;
	if (cellstyleXfsRoot->fillId)
		std::cout << "fillId : " << cellstyleXfsRoot->fillId << std::endl;
	if (cellstyleXfsRoot->borderId)
		std::cout << "borderId : " << cellstyleXfsRoot->borderId << std::endl;
	if (cellstyleXfsRoot->applyNumberFormat)
		std::cout << "applyNumberFormat : " << cellstyleXfsRoot->applyNumberFormat << std::endl;
	if (cellstyleXfsRoot->applyFont)
		std::cout << "applyFont : " << cellstyleXfsRoot->applyFont << std::endl;
	if (cellstyleXfsRoot->applyBorder)
		std::cout << "applyBorder : " << cellstyleXfsRoot->applyBorder << std::endl;
	if (cellstyleXfsRoot->applyAlignment)
		std::cout << "applyAlignment : " << cellstyleXfsRoot->applyAlignment << std::endl;
	if (cellstyleXfsRoot->applyProtection)
		std::cout << "applyProtection : " << cellstyleXfsRoot->applyProtection << std::endl;
	if (cellstyleXfsRoot->Avertical)
		std::cout << "vertical : " << cellstyleXfsRoot->Avertical << std::endl;
}

inline void StyleWrite::numFmidwrite()
{
	while (numFmtsRoot) {
		numFmtsRoot = numFmtsRoot->next;
	}
	if (numFmtsRoot->Id)
		std::cout << "numFmtId : " << numFmtsRoot->Id << std::endl;
	if (numFmtsRoot->Code)
		std::cout << "code : " << numFmtsRoot->Code << std::endl;
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


