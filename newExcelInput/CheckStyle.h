#pragma once

#include "excel_style.h"

class checkstyle:public StyleWrite
{
public:
	checkstyle(FILE* f);
	~checkstyle();

	int searchfonts(Fonts* fs);

	int searchfills(Fills* fs);

	int strcompare(UINT8* sl, UINT8* sr);

	void configstyle(UINT8* num);

	char* SJIStoUTF8(char* szSJIS, char* bufUTF8, int size);

	ArrayNumber strchange;

private:
	char zo[5] = "zozo";
	char sm[7] = "smarby";
	char sh[9] = "shoplist";
	char be[4] = "bee";
};

checkstyle::checkstyle(FILE* f)
{
	fr = f;
}

checkstyle::~checkstyle()
{
}

inline int checkstyle::searchfonts(Fonts* fs)
{
	Fonts* f = fontRoot;
	int result = 0;
	int count = 0;
	bool flag = false;

	std::cout << "fs sz : " << fs->sz << std::endl;
	while (f) {
		result = strcompare(f->sz, fs->sz);
		if (result == 0) {
			result = strcompare(f->rgb, fs->rgb);
			if (result == 0) {
				result = strcompare(f->charset, fs->charset);
				if (result == 0) {
					result = strcompare(f->scheme, fs->scheme);
					if (result == 0) {
						result = strcompare(f->color, fs->color);
						if (result == 0) {
							result = strcompare(f->name, fs->name);
							if (result == 0) {
								result = strcompare(f->family, fs->family);
								if (result == 0) {
									flag = true;
									break;
								}
							}
						}
					}
				}
			}
		}

		f = f->next;
		count++;
	}

	if (!flag)
		count = -1;

	return count;
}

inline int checkstyle::searchfills(Fills* fs)
{
	Fills* f = fillroot;
	int result = 0;
	int count = 0;
	bool flag = false;

	bool fgflag = false;
	bool bgflag = false;

	while (f) {
		if (f->patten && fs->patten) {
			result = strcompare(f->patten->patternType, fs->patten->patternType);
			if (result == 0) {
				flag = true;
			}
		}
		else if (!f->patten && !fs->patten) {
			flag = true;
		}
		else {
			flag = false;
		}

		if (f->fg && fs->fg) {
			result = strcompare(f->fg->rgb, fs->fg->rgb);
			if (result == 0) {
				result = strcompare(f->fg->theme, fs->fg->theme);
				if (result == 0) {
					result = strcompare(f->fg->tint, fs->fg->tint);
					if (result == 0) {
						fgflag = true;
					}
				}
			}
		}
		else if (!f->fg && !fs->fg)
		{
			fgflag = true;
		}
		else {
			fgflag = false;
		}
		
		if (f->bg && fs->bg) {
			result = strcompare(f->bg->indexed, fs->bg->indexed);
			if (result == 0) {
				bgflag = true;
			}
		}
		else if (!f->bg && !fs->bg)
		{
			bgflag = true;
		}
		else {
			bgflag = false;
		}

		if (flag && fgflag && bgflag)
			break;

		f = f->next;
		count++;
	}

	if (!flag || !fgflag || !bgflag) {
		count = -1;
		std::cout << "not match fills : " << count << std::endl;
	}
	else
		std::cout << "match fills : " << count << std::endl;

	return count;
}

inline int checkstyle::strcompare(UINT8* sl, UINT8* sr)
{
	int result = 1;

	if (sl && sr)
		result = strcmp((const char*)sl, (const char*)sr);
	else if (!sl && !sr)
		result = 0;

	return result;
}

inline void checkstyle::configstyle(UINT8* num)
{
	Fonts* fo = (Fonts*)malloc(sizeof(Fonts));
	Fills* fi = (Fills*)malloc(sizeof(Fills));
	cellxfs* cx = (cellxfs*)malloc(sizeof(cellxfs));
	stylexf* csx = (stylexf*)malloc(sizeof(stylexf));

	UINT8* nullbox = (UINT8*)malloc(1); nullbox = nullptr;

	int fontnumb = 0;//検索番号
	int place = 0;//検索番号　桁数

	//入力文字　ショップ別
	int resultshop = 0;
	resultshop = strcmp((const char*)num, be);

	if (resultshop == 0) {
		//shop split bee
		//applyNumberFormat applyAlignment vertical wraptext
		const char* bewraptext[] = { "1","1","center","1" };

		cx->applyNumberFormat = (UINT8*)malloc(2);
		memcpy(cx->applyNumberFormat, (UINT8*)bewraptext[0], 2);

		cx->applyAlignment = (UINT8*)bewraptext[1];
		cx->Avertical = (UINT8*)bewraptext[2];
		cx->AwrapText = (UINT8*)bewraptext[3];

		//他検索後入力

		//fonts 設定
		//font sz rgb name family charset scheme
		const char* befonts[] = { "11", "FF006100","ＭＳ Ｐゴシック", "2","128","minor" };
		char foname[255] = { 0 };
		char* fon = nullptr;
		fon = SJIStoUTF8((char*)befonts[2], foname, 255);//UTF-8に変換

		fo->sz = (UINT8*)malloc(3);
		memcpy(fo->sz, (UINT8*)befonts[0], 3);

		fo->rgb = (UINT8*)malloc(9);
		memcpy(fo->rgb, (UINT8*)befonts[1], 9);

		fo->name = (UINT8*)fon;// (UINT8*)befonts[2];

		fo->family = (UINT8*)malloc(2);
		memcpy(fo->family, (UINT8*)befonts[3], 2);

		fo->charset = (UINT8*)malloc(4);
		memcpy(fo->charset, (UINT8*)befonts[4], 4);

		fo->scheme = (UINT8*)malloc(6);
		memcpy(fo->scheme, (UINT8*)befonts[5], 6);

		fo->color = (UINT8*)malloc(1);
		fo->color = nullptr;

		fontnumb = searchfonts(fo);
		if (fontnumb != -1)
		{
			cx->fontId = strchange.InttoChar(fontnumb, &place);//一致番号入力
			csx->fontId = strchange.InttoChar(fontnumb, &place);
			free(fo->sz); free(fo->rgb); free(fo->family); free(fo->charset); free(fo->scheme); free(fo->color); free(fo);
		}
		else {
			cx->fontId = (UINT8*)malloc(1); cx->fontId=nullptr;//cellxfs font null
			csx->fontId = (UINT8*)malloc(1); csx->fontId = nullptr;
			//フォント設定追加　必要
		}
		
		//fill 設定
		//patternType fgrgb
		const char* befill[] = { "solid","FFC6EFCE" };//他null
		fi->patten = (FillPattern*)malloc(sizeof(FillPattern));

		fi->patten->patternType = (UINT8*)malloc(6);
		memcpy(fi->patten->patternType, (UINT8*)befill[0], 6);

		fi->fg = (fgcolor*)malloc(sizeof(fgcolor));

		fi->fg->rgb = (UINT8*)malloc(9);
		memcpy(fi->fg->rgb, (UINT8*)befill[1], 9);

		fi->fg->theme = (UINT8*)malloc(1); fi->fg->theme = nullptr;
		fi->fg->tint = (UINT8*)malloc(1); fi->fg->tint = nullptr;
		fi->bg = (bgcolor*)malloc(sizeof(bgcolor));//なければストラクトをヌルに
		fi->bg = nullptr;

		fontnumb = searchfills(fi);
		if (fontnumb != -1)
		{
			cx->fillId = strchange.InttoChar(fontnumb, &place);//一致番号入力
			csx->fillId = strchange.InttoChar(fontnumb, &place);
			free(fi->patten->patternType); free(fi->fg->rgb); free(fi->fg->theme); free(fi->fg->tint); free(fi->bg); free(fi->fg);
			free(fi->patten); free(fi);
		}
		else {
			cx->fillId = (UINT8*)malloc(1); cx->fillId = nullptr;//cellxfs font null
			csx->fillId = (UINT8*)malloc(1); csx->fillId = nullptr;
			//フィル設定追加　必要
		}

		//xfIdの設定
		//applyBorder applyAlignment
		const char* bexfids[] = { "0","0","center" };
		cx->applyAlignment = (UINT8*)malloc(2);
		memcpy(cx->applyAlignment, (UINT8*)bexfids[1], 2);

		free(cx->fontId); free(csx->fontId);
		free(cx->fillId); free(csx->fillId);
	}


	/*-----------------------------------
	bee style設定
	-----------------------------------*/

	//border all null

	//numFmtId
	//code
	const char benumFmCode[] = "[$$-409]#,##0.00;[$$-409]#,##0.00"; //id 177(template)

	//xfId
	

	//23
	/*-----------------------------------
	shoplist style設定
	-----------------------------------*/
	//applyNumberFormat vertical horizontal wraptext
	const char* shwraptext[] = { "1","center","left","1" };
	//font sz theme name family charset scheme
	const char* shfonts[] = { "11", "0","ＭＳ Ｐゴシック", "2","128","minor" };
	//patternType fgtheme
	const char* shfill[] = { "solid","4" };//他 null

										   //border all null

	//numFmtId
	//code
	const char shnumFmId[] = "\"\"¥\"#, ##0; [Red] \"¥\"\ - #, ##0\"";//id 0

	//xfId
	//applyBorder applyAlignment
	const char* shxfids[] = { "0","0","center" };

	//26 zozo
	/*-----------------------------------
	zozo style設定
	-----------------------------------*/
	//applyNumberFormat applyAlignment vertical wraptext
	const char* zowraptext[] = { "1","1","center","1" };

	//font sz theme name family charset scheme
	const char* zofonts[] = { "11", "1","ＭＳ Ｐゴシック", "2","128","minor" };

	//patternType fgtheme fgtint bgindexed
	const char* zofill[] = { "solid","4","0.79998168889431442","65" };

	//border all null

	//numFmtId //code
	const char zonumFmId[] = "\"\"¥\"#, ##0; [Red] \"¥\"\ - #, ##0\"";//id=0 

	//xfId
	//applyBorder applyAlignment
	const char* zoxfids[] = { "0","0","center" };

	//28 smarby
	/*-----------------------------------
	smarby style設定
	-----------------------------------*/
	//applyNumberFormat applyAlignment vertical wraptext
	const char* smwraptext[] = { "1","1","center","1" };

	//font sz theme name family charset scheme
	const char* smfonts[] = { "11", "1","ＭＳ Ｐゴシック", "2","128","minor" };

	//patternType fgtheme
	const char* smfill[] = { "solid","5","0.79998168889431442","65" };

	//border all null

	//numFmtId //code
	const char smnumFmId[] = "\"\"¥\"#, ##0; [Red] \"¥\"\ - #, ##0\"";//id=0

	//xfId
	//applyBorder applyAlignment
	const char* smxfids[] = { "0","0","center" };

	//30 magaseek
	/*-----------------------------------
	magaseek style設定
	-----------------------------------*/
	//applyNumberFormat applyAlignment vertical wraptext
	const char* mawraptext[] = { "1","1","center","1" };

	//font sz theme name family charset scheme
	const char* mafonts[] = { "11", "1","ＭＳ Ｐゴシック", "2","128","minor" };
	//patternType fgtheme
	const char* mafill[] = { "solid","6","0.79998168889431442","65" };

	//border all null

	//numFmtId //code
	const char manumFmId[] = "\"\"¥\"#, ##0; [Red] \"¥\"\ - #, ##0\"";//id=0

	//xfId
	//applyBorder applyAlignment
	const char* maxfids[] = { "0","0","center" };


	//UINT32 stylenum = strchange.RowArraytoNum(num, strlen((const char*)num));//style 番号　数字に変換
}

inline char* checkstyle::SJIStoUTF8(char* szSJIS, char* bufUTF8, int size)
{
	wchar_t bufUnicode[100];
	int lenUnicode = MultiByteToWideChar(CP_ACP, 0, szSJIS, strlen(szSJIS) + 1, bufUnicode, 100);
	WideCharToMultiByte(CP_UTF8, 0, bufUnicode, lenUnicode, bufUTF8, size, NULL, NULL);
	return bufUTF8;
}

