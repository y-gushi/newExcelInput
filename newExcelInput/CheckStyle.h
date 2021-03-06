#pragma once

#include "stylewrite.h"
#include <stringapiset.h>

class checkstyle:public StyleWrite
{
public:
	UINT8* style = nullptr;//s 決定

	checkstyle();
	~checkstyle();

	int searchfonts(Fonts* fs);

	int searchcellstylexfs(stylexf* fs);

	int searchcellxfs(cellxfs* fs);

	int searchcellstyle(cellstyle* fs);

	int searchfills(Fills* fs);

	bool searchborderstyle(borderstyle* f, borderstyle* fs);

	int searchborders(borders* fs);

	void CountincIiment(UINT8* cstr);

	UINT8* searchnumFmts(numFMts* fs);

	int strcompare(UINT8* sl, UINT8* sr);

	UINT8* configstyle(UINT8* num);

	char* SJIStoUTF8(char* szSJIS, char* bufUTF8, int size);

private:
	char zo[5] = "zozo";
	char sm[7] = "smarby";
	char sh[9] = "shoplist";
	char be[4] = "bee";
	char ms[9] = "magaseek";
};

checkstyle::checkstyle()
{
}

checkstyle::~checkstyle()
{
}

inline int checkstyle::searchfonts(Fonts* fs)
{
	Fonts** f = fontRoot;
	int result = 0;
	int count = 0;
	bool flag = false;

	//std::cout << "fs theme : " << fs->color << std::endl;
	while (count<frcount) {
		result = strcompare(f[count]->sz, fs->sz);
		if (result == 0) {
			result = strcompare(f[count]->rgb, fs->rgb);
			if (result == 0) {
				result = strcompare(f[count]->charset, fs->charset);
				if (result == 0) {
					result = strcompare(f[count]->scheme, fs->scheme);
					if (result == 0) {
						result = strcompare(f[count]->color, fs->color);
						if (result == 0) {
							result = strcompare(f[count]->name, fs->name);
							if (result == 0) {
								result = strcompare(f[count]->family, fs->family);
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

		count++;
	}

	if (!flag)
		count = -1;
	else
		std::cout << "match font : " << count << std::endl;

	return count;
}

inline int checkstyle::searchcellstylexfs(stylexf* fs)
{
	stylexf** f = cellstyleXfsRoot;
	int result = 0;
	int count = 0;
	bool flag = false;

	//std::cout << "fs aA : " << fs->fillId << " " << f->fillId << std::endl;
	while (count<csXcount) {
		result = strcompare(f[count]->applyAlignment, fs->applyAlignment);
		if (result == 0) {
			result = strcompare(f[count]->applyBorder, fs->applyBorder);
			if (result == 0) {
				result = strcompare(f[count]->applyFont, fs->applyFont);
				if (result == 0) {
					result = strcompare(f[count]->applyNumberFormat, fs->applyNumberFormat);
					if (result == 0) {
						result = strcompare(f[count]->applyProtection, fs->applyProtection);
						if (result == 0) {
							result = strcompare(f[count]->Avertical, fs->Avertical);
							if (result == 0) {
								result = strcompare(f[count]->borderId, fs->borderId);
								if (result == 0) {
									result = strcompare(f[count]->fillId, fs->fillId);
									if (result == 0) {
										result = strcompare(f[count]->fontId, fs->fontId);
										if (result == 0) {
											result = strcompare(f[count]->numFmtId, fs->numFmtId);
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
				}
			}
		}
		count++;
	}

	if (!flag)
		count = -1;
	else
		std::cout << "match cellstylexfms : " << count << std::endl;

	return count;
}

inline int checkstyle::searchcellxfs(cellxfs* fs)
{
	cellxfs** f = cellXfsRoot;
	int result = 0;
	int count = 0;
	bool flag = false;

	//std::cout << "fs aA : " << fs->fontId << " " << f->fontId << std::endl;
	while (count<cXcount) {
		result = strcompare(f[count]->applyAlignment, fs->applyAlignment);
		if (result == 0) {
			result = strcompare(f[count]->applyBorder, fs->applyBorder);
			if (result == 0) {
				result = strcompare(f[count]->applyFont, fs->applyFont);
				if (result == 0) {
					result = strcompare(f[count]->applyNumberFormat, fs->applyNumberFormat);
					if (result == 0) {
						result = strcompare(f[count]->applyFill, fs->applyFill);
						if (result == 0) {
							result = strcompare(f[count]->Avertical, fs->Avertical);
							if (result == 0) {
								result = strcompare(f[count]->borderId, fs->borderId);
								if (result == 0) {
									result = strcompare(f[count]->fillId, fs->fillId);
									if (result == 0) {
										result = strcompare(f[count]->fontId, fs->fontId);
										if (result == 0) {
											result = strcompare(f[count]->numFmtId, fs->numFmtId);
											if (result == 0) {
												result = strcompare(f[count]->AwrapText, fs->AwrapText);
												if (result == 0) {
													result = strcompare(f[count]->xfId, fs->xfId);
													if (result == 0) {
														result = strcompare(f[count]->horizontal, fs->horizontal);
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
							}
						}
					}
				}
			}
		}

		count++;
	}

	if (!flag)
		count = -1;
	else
		std::cout << "match cellxfms : " << count << std::endl;

	return count;
}
//cellstyle 検索　xfid戻り
inline int checkstyle::searchcellstyle(cellstyle* fs)
{
	cellstyle** f = cellStyleRoot;
	int result = 0;
	int count = 0;
	bool flag = false;

	//std::cout << "fs name : " << fs->name << " " << f->name << std::endl;
	while (cscount) {
		if (f[count]->name)
			result = strcompare(f[count]->xfId, fs->xfId);
		if (result == 0) {
			flag = true;
			break;
		}

		count++;
	}

	if (!flag) {
		count = -1;
	}
	else {
		std::cout << "match xfid : " << f[count]->xfId << std::endl;
	}

	return count;
}

inline int checkstyle::searchfills(Fills* fs)
{
	Fills** f = fillroot;
	int result = 0;
	int count = 0;
	bool flag = false;

	bool fgflag = false;
	bool bgflag = false;

	while (count<ficount) {
		if (f[count]->patten && fs->patten) {
			result = strcompare(f[count]->patten->patternType, fs->patten->patternType);
			if (result == 0) {
				flag = true;
			}
		}
		else if (!f[count]->patten && !fs->patten) {
			flag = true;
		}
		else {
			flag = false;
		}

		if (f[count]->fg && fs->fg) {
			result = strcompare(f[count]->fg->rgb, fs->fg->rgb);
			if (result == 0) {
				result = strcompare(f[count]->fg->theme, fs->fg->theme);
				if (result == 0) {
					result = strcompare(f[count]->fg->tint, fs->fg->tint);
					if (result == 0) {
						fgflag = true;
					}
				}
			}
		}
		else if (!f[count]->fg && !fs->fg)
		{
			fgflag = true;
		}
		else {
			fgflag = false;
		}

		if (f[count]->bg && fs->bg) {
			result = strcompare(f[count]->bg->indexed, fs->bg->indexed);
			if (result == 0) {
				bgflag = true;
			}
		}
		else if (!f[count]->bg && !fs->bg)
		{
			bgflag = true;
		}
		else {
			bgflag = false;
		}

		if (flag && fgflag && bgflag)
			break;

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

inline bool checkstyle::searchborderstyle(borderstyle* f, borderstyle* fs) {
	bool flag = false;
	int result = 0;

	if (f && fs) {
		result = strcompare(f->rgb, fs->rgb);
		if (result == 0) {
			result = strcompare(f->Auto, fs->Auto);
			if (result == 0) {
				result = strcompare(f->index, fs->index);
				if (result == 0) {
					result = strcompare(f->style, fs->style);
					if (result == 0) {
						result = strcompare(f->theme, fs->theme);
						if (result == 0) {
							result = strcompare(f->tint, fs->tint);
							if (result == 0) {
								flag = true;
							}
						}
					}
				}
			}
		}
	}
	else if (!f && !fs) {
		flag = true;
	}

	return flag;
}

inline int checkstyle::searchborders(borders* fs)
{
	borders** f = BorderRoot;
	int result = 0;
	int count = 0;
	bool lflag = false;

	bool rflag = false;
	bool bflag = false;
	bool tflag = false;
	bool dflag = false;

	while (count<bocount) {

		lflag = searchborderstyle(f[count]->left, fs->left);

		rflag = searchborderstyle(f[count]->right, fs->right);

		bflag = searchborderstyle(f[count]->bottom, fs->bottom);

		tflag = searchborderstyle(f[count]->top, fs->top);

		dflag = searchborderstyle(f[count]->diagonal, fs->diagonal);

		if (lflag && rflag && bflag && tflag && dflag)
			break;

		count++;
	}

	if (!lflag || !rflag || !tflag || !bflag || !dflag) {
		count = -1;
		std::cout << "not match border : " << count << std::endl;
	}
	else
		std::cout << "match border : " << count << std::endl;

	return count;
}

inline void checkstyle::CountincIiment(UINT8* cstr)
{
	int l = 0;
	UINT32 count = 0;
	while (cstr[l] != '\0')
		l++;

	std::cout << "count ;" << cstr << std::endl;

	count = strchange.RowArraytoNum(cstr, l);
	count++;

	cstr = strchange.InttoChar(count, &l);

	std::cout << "count plus ;" << cstr << std::endl;
}

inline UINT8* checkstyle::searchnumFmts(numFMts* fs)
{
	//numFMts* f = numFmtsRoot;
	int result = 0;
	int count = 0;
	bool flag = false;
	UINT8* nu = nullptr;

	//std::cout << "fs sz : " << fs->sz << std::endl;
	while (count<nucount) {
		result = strcompare(numFmtsRoot[count]->Code, fs->Code);
		if (result == 0) {
			flag = true;
			break;
		}
		count++;
	}

	if (!flag) {
		nu = (UINT8*)malloc(1); nu = nullptr;
		return nu;
	}
	else
		std::cout << "match numFmts : " << count << std::endl;

	size_t leng = strlen((const char*)numFmtsRoot[count]->Id) + 1;
	nu = (UINT8*)malloc(leng);
	memcpy(nu, numFmtsRoot[count]->Id, leng);

	return nu;
}

inline int checkstyle::strcompare(UINT8* sl, UINT8* sr)
{
	int result = 1;

	if (sl && sr)
		result = strcmp((const char*)sl, (const char*)sr);
	else if (!sl && !sr)
		result = 0;
	else
		result = 1;

	return result;
}

inline UINT8* checkstyle::configstyle(UINT8* num)
{
	Fonts* fo = (Fonts*)malloc(sizeof(Fonts));
	Fills* fi = (Fills*)malloc(sizeof(Fills));
	cellxfs* cx = (cellxfs*)malloc(sizeof(cellxfs));
	stylexf* csx = (stylexf*)malloc(sizeof(stylexf));
	numFMts* nf = (numFMts*)malloc(sizeof(numFMts));
	borders* bd = (borders*)malloc(sizeof(borders));
	cellstyle* cs = nullptr;

	UINT8* nullbox = (UINT8*)malloc(1); nullbox = nullptr;

	int fontnumb = 0;//検索番号
	int place = 0;//検索番号　桁数

	/************共通設定　cellXFs**************/
	//applyNumberFormat applyAlignment vertical wraptext numFmId
	const char* bewraptext[] = { "1","1","center","1","0" };

	//ID検索以外の値入れる
	cx->applyNumberFormat = (UINT8*)malloc(2);
	memcpy(cx->applyNumberFormat, (UINT8*)bewraptext[0], 2);
	cx->applyAlignment = (UINT8*)malloc(2);
	memcpy(cx->applyAlignment, (UINT8*)bewraptext[1], 2);
	cx->Avertical = (UINT8*)malloc(7);
	memcpy(cx->Avertical, (UINT8*)bewraptext[2], 7);
	cx->AwrapText = (UINT8*)malloc(2);
	memcpy(cx->AwrapText, (UINT8*)bewraptext[3], 2);
	cx->numFmtId = (UINT8*)malloc(2);
	memcpy(cx->numFmtId, (UINT8*)bewraptext[4], 2);
	cx->applyBorder = (UINT8*)malloc(1); cx->applyBorder = nullptr;
	cx->horizontal = (UINT8*)malloc(1); cx->horizontal = nullptr;
	cx->applyFill = (UINT8*)malloc(1); cx->applyFill = nullptr;
	cx->applyFont = (UINT8*)malloc(1); cx->applyFont = nullptr;

	/************共通設定　フォント**************/
	const char fname[] = "ＭＳ Ｐゴシック";
	char foname[255] = { 0 };
	char* fon = nullptr;
	fon = SJIStoUTF8((char*)fname, foname, 255);//フォントname UTF-8に変換

	const char* sharefonts[] = { "11", "2","128","minor" };//sz family charset scheme
	fo->sz = (UINT8*)malloc(3);
	memcpy(fo->sz, (UINT8*)sharefonts[0], 3);
	fo->name = (UINT8*)malloc(strlen(fon) + 1);
	memcpy(fo->name, (UINT8*)fon, strlen(fon) + 1);
	fo->family = (UINT8*)malloc(2);
	memcpy(fo->family, (UINT8*)sharefonts[1], 2);
	fo->charset = (UINT8*)malloc(4);
	memcpy(fo->charset, (UINT8*)sharefonts[2], 4);
	fo->scheme = (UINT8*)malloc(6);
	memcpy(fo->scheme, (UINT8*)sharefonts[3], 6);

	/************共通設定　フィル**************/
	//patternType fgrgb
	const char patternFill[] = "solid";

	fi->patten = (FillPattern*)malloc(sizeof(FillPattern));
	fi->patten->patternType = (UINT8*)malloc(6);
	memcpy(fi->patten->patternType, (UINT8*)patternFill, 6);

	/************ 共通設定　ボーダー **************/
	bd->left = (borderstyle*)malloc(sizeof(borderstyle)); bd->left = nullptr;
	bd->right = (borderstyle*)malloc(sizeof(borderstyle)); bd->right = nullptr;
	bd->top = (borderstyle*)malloc(sizeof(borderstyle)); bd->top = nullptr;
	bd->bottom = (borderstyle*)malloc(sizeof(borderstyle)); bd->bottom = nullptr;
	bd->diagonal = (borderstyle*)malloc(sizeof(borderstyle)); bd->diagonal = nullptr;

	/************共通設定　cellstyleXFs**************/

	const char* bexfids[] = { "0","0","center" };
	csx->applyAlignment = (UINT8*)malloc(2);
	memcpy(csx->applyAlignment, (UINT8*)bexfids[1], 2);
	csx->applyBorder = (UINT8*)malloc(2);
	memcpy(csx->applyBorder, (UINT8*)bexfids[0], 2);
	csx->Avertical = (UINT8*)malloc(7);
	memcpy(csx->Avertical, (UINT8*)bexfids[2], 7);

	csx->applyNumberFormat = (UINT8*)malloc(1); csx->applyNumberFormat = nullptr;
	csx->applyFont = (UINT8*)malloc(1); csx->applyFont = nullptr;
	csx->applyProtection = (UINT8*)malloc(1); csx->applyProtection = nullptr;
	csx->applyFill = (UINT8*)malloc(1); csx->applyFill = nullptr;

	//入力文字　ショップ別
	int resultshop = 0;
	resultshop = strcmp((const char*)num, be);
	/*-----------------------------------
	bee style設定
	-----------------------------------*/
	if (resultshop == 0) {
		//bee
				//fonts 設定
		//------------- ショップ別 フォント色 -----------------//
		//font sz rgb name family charset scheme
		const char befonts[] = "FF006100";
		fo->rgb = (UINT8*)malloc(9);
		memcpy(fo->rgb, (UINT8*)befonts, 9);
		fo->color = (UINT8*)malloc(1); fo->color = nullptr;

		//fill 設定
		//------------- ショップ別 フォント色 -----------------//
		const char befill[] = "FFC6EFCE";//他null
		fi->fg = (fgcolor*)malloc(sizeof(fgcolor));
		fi->fg->rgb = (UINT8*)malloc(9);
		memcpy(fi->fg->rgb, (UINT8*)befill, 9);
		fi->fg->theme = (UINT8*)malloc(1); fi->fg->theme = nullptr;
		fi->fg->tint = (UINT8*)malloc(1); fi->fg->tint = nullptr;
		fi->bg = (bgcolor*)malloc(sizeof(bgcolor));//なければストラクトをヌルに
		fi->bg = nullptr;

		//numFmid 検索
		//numFmtId //code
		const char benumFmCode[] = "[$$-409]#,##0.00;[$$-409]#,##0.00"; //id 177(template) beeだけ
		nf->Code = (UINT8*)malloc(34);
		memcpy(nf->Code, (UINT8*)benumFmCode, 34);

		nf->Id = (UINT8*)malloc(1); nf->Id = nullptr;

		csx->numFmtId = searchnumFmts(nf);//numFmt 加える
		if (csx->numFmtId) {
			std::cout << "match numfmts : " << csx->numFmtId << std::endl;
			free(nf->Code); free(nf->Id); free(nf);
		}
		else {
			//numFmts 設定 0にする
			UINT8 ze[] = "0";
			free(csx->numFmtId);
			csx->numFmtId = (UINT8*)malloc(2);
			memcpy(csx->numFmtId, ze, 2);
		}
	}

	//23
	/*-----------------------------------
	shoplist style設定 他　fillが違うだけ
	-----------------------------------*/
	resultshop = strcmp((const char*)num, sh);
	if (resultshop == 0) {
		std::cout << "shoplist" << std::endl;
		//applyNumberFormat vertical horizontal wraptext
	//const char* shwraptext[] = { "1","center","left","1" };
	//------------- ショップ別 cellXFs -----------------//
	//shoplist
		const char sholhorizen[] = "left";//vartival
		//vertical変更
		free(cx->horizontal);
		cx->horizontal = (UINT8*)malloc(5);//color theme
		memcpy(cx->horizontal, (UINT8*)sholhorizen, 5);

		//------------- ショップ別 フォント色 -----------------//
		//shoplist
		const char sholtheme[] = "0";//theme
		fo->color = (UINT8*)malloc(2);//color theme
		memcpy(fo->color, (UINT8*)sholtheme, 2);
		fo->rgb = (UINT8*)malloc(1); fo->rgb = nullptr;//color rgb どちらか

		//fill 設定
		//------------- ショップ別 フィル色 -----------------//
		//shoplist
		const char sholfill[] = "4";//theme 他null
		fi->fg = (fgcolor*)malloc(sizeof(fgcolor));
		fi->fg->rgb = (UINT8*)malloc(1); fi->fg->rgb = nullptr;//rgb
		fi->fg->theme = (UINT8*)malloc(2);
		memcpy(fi->fg->theme, (UINT8*)sholfill, 2);//theme
		fi->fg->tint = (UINT8*)malloc(1); fi->fg->tint = nullptr;
		fi->bg = (bgcolor*)malloc(sizeof(bgcolor));//なければストラクトをヌルに
		fi->bg = nullptr;

		//cellstyleXF 設定
		UINT8 zero[] = "0";
		csx->numFmtId = (UINT8*)malloc(2);//numFmId 0
		memcpy(csx->numFmtId, zero, 2);

		//cellstyle なし
	}

	resultshop = strcmp((const char*)num, zo);
	if (resultshop == 0) {
		std::cout << "zozo" << std::endl;
		//------------- ショップ別 フォント色 -----------------//
		//zozo
		const char zozotheme[] = "1";//theme
		fo->color = (UINT8*)malloc(2);//color theme
		memcpy(fo->color, (UINT8*)zozotheme, 2);
		fo->rgb = (UINT8*)malloc(1); fo->rgb = nullptr;//color rgb どちらか

		//------------- ショップ別 フィル色 -----------------//
		//zozo
		// fgtheme fgtint bgindexed
		const char* zofill[] = { "4","0.79998168889431442","65" };
		fi->fg = (fgcolor*)malloc(sizeof(fgcolor));
		fi->fg->theme = (UINT8*)malloc(2);
		memcpy(fi->fg->theme, (UINT8*)zofill[0], 2);//theme
		fi->fg->tint = (UINT8*)malloc(20);
		memcpy(fi->fg->tint, (UINT8*)zofill[1], 20);//tint
		fi->fg->rgb = (UINT8*)malloc(1); fi->fg->rgb = nullptr;//rgb

		fi->bg = (bgcolor*)malloc(sizeof(bgcolor));//なければストラクトをヌルに
		fi->bg->indexed = (UINT8*)malloc(3);
		memcpy(fi->bg->indexed, (UINT8*)zofill[2], 3);//tint

		//cellstyleXF 設定
		UINT8 zero[] = "0";
		csx->numFmtId = (UINT8*)malloc(2);//numFmId 0
		memcpy(csx->numFmtId, zero, 2);

		//cellstyle なし
	}

	resultshop = strcmp((const char*)num, sm);
	if (resultshop == 0) {
		std::cout << "smarby" << std::endl;
		//------------- ショップ別 フォント色 -----------------//
		//smarby
		const char smtheme[] = "1";//theme
		fo->color = (UINT8*)malloc(2);//color theme
		memcpy(fo->color, (UINT8*)smtheme, 2);
		fo->rgb = (UINT8*)malloc(1); fo->rgb = nullptr;//color rgb どちらか

		//------------- ショップ別 フィル色 -----------------//
		//smarby
		const char* smfill[] = { "5","0.79998168889431442","65" };
		fi->fg = (fgcolor*)malloc(sizeof(fgcolor));
		fi->fg->theme = (UINT8*)malloc(2);
		memcpy(fi->fg->theme, (UINT8*)smfill[0], 2);//theme
		fi->fg->tint = (UINT8*)malloc(20);
		memcpy(fi->fg->tint, (UINT8*)smfill[1], 20);//tint
		fi->fg->rgb = (UINT8*)malloc(1); fi->fg->rgb = nullptr;//rgb

		fi->bg = (bgcolor*)malloc(sizeof(bgcolor));//なければストラクトをヌルに
		fi->bg->indexed = (UINT8*)malloc(3);
		memcpy(fi->bg->indexed, (UINT8*)smfill[2], 3);//tint

		//cellstyleXF 設定
		UINT8 zero[] = "0";
		csx->numFmtId = (UINT8*)malloc(2);//numFmId 0
		memcpy(csx->numFmtId, zero, 2);

		//cellstyle なし
	}

	resultshop = strcmp((const char*)num, ms);
	if (resultshop == 0) {
		std::cout << "magaseek : " << std::endl;
		//------------- ショップ別 フォント色 -----------------//
		//magaseek
		const char mstheme[] = "1";//theme
		fo->color = (UINT8*)malloc(2);//color theme
		memcpy(fo->color, (UINT8*)mstheme, 2);
		fo->rgb = (UINT8*)malloc(1); fo->rgb = nullptr;//color rgb どちらか

		//------------- ショップ別 フィル色 -----------------//
		//magaseek
		const char* mafill[] = { "6","0.79998168889431442","65" };
		fi->fg = (fgcolor*)malloc(sizeof(fgcolor));
		fi->fg->theme = (UINT8*)malloc(2);
		memcpy(fi->fg->theme, (UINT8*)mafill[0], 2);//theme
		fi->fg->tint = (UINT8*)malloc(20);
		memcpy(fi->fg->tint, (UINT8*)mafill[1], 20);//tint
		fi->fg->rgb = (UINT8*)malloc(1); fi->fg->rgb = nullptr;

		fi->bg = (bgcolor*)malloc(sizeof(bgcolor));//なければストラクトをヌルに
		fi->bg->indexed = (UINT8*)malloc(3);
		memcpy(fi->bg->indexed, (UINT8*)mafill[2], 3);//tint

		//cellstyleXF 設定
		UINT8 zero[] = "0";
		csx->numFmtId = (UINT8*)malloc(2);//numFmId 0
		memcpy(csx->numFmtId, zero, 2);

		//cellstyle なし
	}

	/*フォントidの検索*/
	fontnumb = searchfonts(fo);
	if (fontnumb != -1)
	{
		cx->fontId = strchange.InttoChar(fontnumb, &place);//一致番号入力　fontID 加える
		csx->fontId = strchange.InttoChar(fontnumb, &place);
		free(fo->sz); free(fo->rgb); free(fo->family); free(fo->charset); free(fo->scheme); free(fo->color); free(fo);
	}
	else {
		//フォント設定追加　必要

		size_t resiz = fontNum + 1;
		Fonts** rexfs = nullptr;

		rexfs = (Fonts**)realloc(fillroot, resiz);

		for (int i = 0; i < fontNum; i++)
			freefonts(fontRoot[i]);

		fontRoot = rexfs;

		if (!rexfs) {
			std::cout << "not keep memory" << std::endl;
			UINT8 er[] = "memory error";
			return er;
		}
		else {
			for (int i = 0; i < fontNum; i++)
				freefonts(fontRoot[i]);
		}

		std::cout << "need add cellstyleXfs" << std::endl;

		rexfs[fillNum] = (Fonts*)malloc(sizeof(Fonts));
		rexfs[fillNum]->sz = fo->sz;
		rexfs[fillNum]->rgb = fo->rgb;
		rexfs[fillNum]->family = fo->family;
		rexfs[fillNum]->charset = fo->charset;
		rexfs[fillNum]->scheme = fo->scheme;
		rexfs[fillNum]->color = fo->color;

		frcount++; fontNum++;//ボーダー数プラス

		free(fontCount);
		fontCount = strchange.InttoChar(fontNum, &place);

		cx->fontId = strchange.InttoChar(fontNum, &place);
		csx->fontId = strchange.InttoChar(fontNum, &place);
	}

	/*フィルidの検索*/
	fontnumb = searchfills(fi);
	if (fontnumb != -1)
	{
		cx->fillId = strchange.InttoChar(fontnumb, &place);//一致番号入力 フィルID加える
		csx->fillId = strchange.InttoChar(fontnumb, &place);
		free(fi->patten->patternType); free(fi->fg->rgb); free(fi->fg->theme); free(fi->fg->tint); free(fi->bg); free(fi->bg->indexed); free(fi->fg);
		free(fi->patten); free(fi);
	}
	else {
		//フィル設定追加　必要
		size_t resiz = fillNum + 1;
		Fills** rexfs = nullptr;

		rexfs = (Fills**)realloc(fillroot, resiz);
		if (!rexfs) {
			std::cout << "not keep memory" << std::endl;
			UINT8 er[] = "memory error";
			return er;
		}

		std::cout << "need add cellstyleXfs" << std::endl;

		rexfs[fillNum] = (Fills*)malloc(sizeof(Fills));
		rexfs[fillNum]->patten = (FillPattern*)malloc(sizeof(FillPattern));
		rexfs[fillNum]->patten->patternType = fi->patten->patternType;
		rexfs[fillNum]->fg = (fgcolor*)malloc(sizeof(fgcolor));
		rexfs[fillNum]->fg->rgb = fi->fg->rgb;
		rexfs[fillNum]->fg->theme = fi->fg->theme;
		rexfs[fillNum]->fg->tint = fi->fg->tint;
		rexfs[fillNum]->bg = (bgcolor*)malloc(sizeof(bgcolor));
		rexfs[fillNum]->bg->indexed = fi->bg->indexed;

		ficount++; fillNum++;//ボーダー数プラス

		free(fillCount);
		fillCount = strchange.InttoChar(fillNum, &place);

		cx->fillId = strchange.InttoChar(fillNum, &place);
		csx->fillId = strchange.InttoChar(fillNum, &place);
	}

	/*ボーダーidの検索*/
	fontnumb = searchborders(bd);
	if (fontnumb != -1)
	{
		cx->borderId = strchange.InttoChar(fontnumb, &place);//一致番号入力 ボーダー設定加える
		csx->borderId = strchange.InttoChar(fontnumb, &place);
		free(bd->left); free(bd->right); free(bd->top); free(bd->bottom); free(bd->diagonal); free(bd);
	}
	else {

		//ボーダー設定追加　必要
		size_t resiz = borderNum + 1;
		borders** rexfs = nullptr;

		rexfs = (borders**)realloc(BorderRoot, resiz);
		if (!rexfs) {
			std::cout << "not keep memory" << std::endl;
			UINT8 er[] = "memory error";
			return er;
		}

		std::cout << "need add cellstyleXfs" << std::endl;

		rexfs[borderNum] = (borders*)malloc(sizeof(borders));
		rexfs[borderNum]->left = bd->left;
		rexfs[borderNum]->right = bd->right;
		rexfs[borderNum]->top = bd->top;
		rexfs[borderNum]->left = bd->left;
		rexfs[borderNum]->bottom = bd->bottom;
		rexfs[borderNum]->diagonal = bd->diagonal;

		bocount++; borderNum++;//ボーダー数プラス

		free(borderCount);
		borderCount = strchange.InttoChar(borderNum, &place);

		cx->borderId = strchange.InttoChar(borderNum, &place);
		csx->borderId = strchange.InttoChar(borderNum, &place);

	}

	//cellstyle xfs 番号取得
	resultshop = strcmp((const char*)num, be);
	if (resultshop != 0) {
		//bee 以外　xfId取得
		fontnumb = searchcellstylexfs(csx);
		if (fontnumb != -1) {
			cx->xfId = strchange.InttoChar(fontnumb, &place);//一致番号入力 bee以外はこっち
			free(csx->applyAlignment); free(csx->applyBorder); free(csx->applyFont); free(csx->applyNumberFormat); free(csx->applyProtection);
			free(csx->Avertical); free(csx);
		}
		else {
			//cellstyle xfs 設定追加　必要

			size_t resiz = cellstyleXfsNum + 1;
			stylexf** rexfs = nullptr;

			rexfs = (stylexf**)realloc(cellstyleXfsRoot, resiz);
			if (!rexfs) {
				std::cout << "not keep memory" << std::endl;
				UINT8 er[] = "memory error";
				return er;
			}

			std::cout << "need add cellstyleXfs" << std::endl;
			
			rexfs[csXcount] = (stylexf*)malloc(sizeof(stylexf));
			rexfs[csXcount]->numFmtId = csx->numFmtId;
			rexfs[csXcount]->fontId = csx->fontId;
			rexfs[csXcount]->fillId = csx->fillId;
			rexfs[csXcount]->borderId = csx->borderId;
			rexfs[csXcount]->applyNumberFormat = csx->applyNumberFormat;
			rexfs[csXcount]->applyBorder = csx->applyBorder;
			rexfs[csXcount]->applyAlignment = csx->applyAlignment;
			rexfs[csXcount]->applyProtection = csx->applyProtection;
			rexfs[csXcount]->Avertical = csx->Avertical;//alignment vertical
			rexfs[csXcount]->applyFont = csx->applyFont;
			rexfs[csXcount]->applyFill = csx->applyFill;
			
			csXcount++; cellstyleXfsNum++;

			free(cellStyleXfsCount);
			cellStyleXfsCount = strchange.InttoChar(dxfsNm, &place);

			cx->xfId= strchange.InttoChar(dxfsNm, &place);

		}
	}
	else {
		//------------- セルスタイル設定-----------------//
		//bee だけ　スタイルセット検索
		cs = (cellstyle*)malloc(sizeof(cellstyle));
		//name xr:uid xfid
		const char* celsty[] = { "良い 11 10 2 3 2 2 2 4 2 2" ,"{00000000-0005-0000-0000-000039000000}" ,"58" };
		char styname[255] = { 0 };
		char* cstyn = nullptr;
		cstyn = SJIStoUTF8((char*)celsty[0], styname, 255);//フォントname UTF-8に変換

		cs->name = (UINT8*)cstyn;
		cs->xruid = (UINT8*)malloc(39);
		cs->xfId = (UINT8*)malloc(3);
		memcpy(cs->xruid, (UINT8*)celsty[1], 39);
		memcpy(cs->xfId, (UINT8*)celsty[2], 3);
		cs->customBuilt = (UINT8*)malloc(1); cs->customBuilt = nullptr;
		cs->builtinId = (UINT8*)malloc(1); cs->builtinId = nullptr;

		fontnumb = searchcellstyle(cs);
		if (fontnumb != -1) {
			//style set 設定あり 
			cx->xfId = (UINT8*)malloc(3);
			memcpy(cx->xfId, (UINT8*)celsty[2], 3);
		}
		//free(cs->name); 
		free(cs->xfId); free(cs->builtinId); free(cs->xruid); free(cs);
	}

	//style 決定
	fontnumb = searchcellxfs(cx);
	if (fontnumb != -1)
	{
		style = strchange.InttoChar(fontnumb, &place);//一致番号入力　style 決定
		std::cout << "style : " << style << std::endl;
		free(cx->fontId);
		free(cx->fillId); free(cx->applyAlignment); free(cx->applyBorder); free(cx->applyFill); free(cx->applyFont);
		free(cx->applyNumberFormat);  free(cx->borderId);
		free(cx->numFmtId); free(cx->xfId); free(cx->AwrapText);
		free(cx->Avertical);
	}
	else {
		
		//xfID 設定追加　必要
		size_t resiz = dxfsNm + 1;
		cellxfs** rexfs = nullptr;

		rexfs = (cellxfs**)realloc(cellXfsRoot, resiz);
		if (!rexfs) {
			std::cout << "not keep memory" << std::endl;
			UINT8 er[] = "memory error";
			return er;
		}
		rexfs[dxfsNm] = (cellxfs*)malloc(sizeof(cellxfs));
		rexfs[dxfsNm]->fontId = cx->fontId;
		rexfs[dxfsNm]->fillId = cx->fillId;
		rexfs[dxfsNm]->applyAlignment = cx->applyAlignment;
		rexfs[dxfsNm]->applyBorder = cx->applyBorder;
		rexfs[dxfsNm]->applyFill = cx->applyFill;
		rexfs[dxfsNm]->applyFont = cx->applyFont;
		rexfs[dxfsNm]->applyNumberFormat = cx->applyNumberFormat;
		rexfs[dxfsNm]->borderId = cx->borderId;
		rexfs[dxfsNm]->numFmtId = cx->numFmtId;
		rexfs[dxfsNm]->xfId = cx->xfId;
		rexfs[dxfsNm]->AwrapText = cx->AwrapText;
		rexfs[dxfsNm]->Avertical = cx->Avertical;

		dxfsNm++; dxcount++;
		free(dxfsCount);
		dxfsCount = strchange.InttoChar(dxfsNm, &place);

		style = (UINT8*)malloc(place+1); 
		memcpy(style,dxfsCount,place+1);//cellxfs font null
	}

	//UINT32 stylenum = strchange.RowArraytoNum(num, strlen((const char*)num));//style 番号　数字に変換
	return nullptr;
}

inline char* checkstyle::SJIStoUTF8(char* szSJIS, char* bufUTF8, int size)
{
	wchar_t bufUnicode[100];
	int lenUnicode = MultiByteToWideChar(CP_ACP, 0, szSJIS, strlen(szSJIS) + 1, bufUnicode, 100);
	WideCharToMultiByte(CP_UTF8, 0, bufUnicode, lenUnicode, bufUTF8, size, NULL, NULL);
	return bufUTF8;
}

