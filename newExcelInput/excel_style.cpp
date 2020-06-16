#include "excel_style.h"

styleread::styleread()
{
	readp = 0;
	fontcou = 0;
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

Fills* styleread::addfill(Fills* fi, UINT8* t, UINT8* f, UINT8* b) {
	if (!fi) {
		fi = (Fills*)malloc(sizeof(Fills));
		fi->type = t;
		fi->fg = f;
		fi->bg = b;
		fi->next = nullptr;
	}
	else {
		fi->next = addfill(fi->next, t, f, b);
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
		sx->Avertical = av;
		sx->next = nullptr;
	}
	else {
		sx->next = addstylexf(sx->next, n, fo, fi, bi, an, ab, aa, ap, av);
	}

	return sx;
}

cellxfs* styleread::addcellxfs(cellxfs* c, UINT8* n, UINT8* fo, UINT8* fi, UINT8* bi, UINT8* an, UINT8* ab, UINT8* aa, UINT8* av, UINT8* at) {
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
		c->next = nullptr;
	}
	else {
		c->next = addcellxfs(c->next, n, fo, fi, bi, an, ab, aa, av,at);
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

UINT8* styleread::readfonts(UINT8* sd, UINT64 len,UINT64 rp) {

	const char sz[9] = "sz val=\"";//8
	const char colo[14] = "color theme=\"";//13
	const char name[11] = "name val=\"";//10
	const char fami[13] = "family val=\"";//12
	const char chars[14] = "charset val=\"";//13
	const char sche[13] = "scheme val=\"";//12

	UINT8 endtag[9] = { 0 };//8文字
	UINT8 strsear[14] = { 0 };//13文字
	UINT8 seafam[13] = { 0 };//12文字
	UINT8 seanam[11] = { 0 };//10文字

	size_t msize = 0;//malloc用

	size_t vallen = 0;//tag 値
	size_t taglen = 0;//タグ　読み込み開始位置

	int eresult = 1;
	int txtresult = 0;
	while (eresult != 0) {
		//1つずつずらす
		for (int i = 0; i < 12; i++) {
			strsear[i] = strsear[i + 1];
			if (i < 7)
				endtag[i] = endtag[i + 1];
			if (i < 12 - 1)
				seafam[i] = seafam[i + 1];
			if (i < 10 - 1)
				seanam[i] = seanam[i + 1];
		}
		strsear[12] = endtag[7]=seafam[11]=seanam[9] = sd[rp];
		strsear[13] = endtag[8] = seafam[12] = seanam[10] = '\0';
		rp++;

		txtresult = strncmp((const char*)endtag, sz, 8);
		if (txtresult == 0) {
			//sz読み込み
			vallen = 0;
			taglen = size_t(rp);
			while (sd[rp] != '>') {
				if (sd[rp] != '"' && sd[rp] != '/')//read data
					vallen++;
				rp++;
			}
			msize = vallen + 1;
			UINT8* sz = (UINT8*)malloc(msize);
			for (size_t i = 0; i < vallen; i++)
				sz[i] = sd[taglen + i];
			sz[vallen] = '\0';
		}
		
		eresult = strcmp((const char*)endtag, Efont);
	}

	return nullptr;
}

void styleread::readstyle(UINT8* sdata, UINT64 sdatalen)
{
	UINT8 sEs[14] = { 0 };//13文字
	UINT8 sExfs[16] = { 0 };//15文字
	UINT8 sfon[7] = { 0 };//6文字
	UINT8 sdxf[6] = { 0 };//5文字
	UINT8 sbor[9] = { 0 };//8文字
	UINT8 sstyle[12] = { 0 };//11文字
	int result = 1;
	int otherresult = 0;

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
			
		}

		otherresult = strncmp((const char*)sfon, fills, 6);
		if (otherresult == 0) {
			//フィル一致 read font

		}

		otherresult = strncmp((const char*)sbor, border, 8);
		if (otherresult == 0) {
			//ボーダー一致 read border

		}

		otherresult = strncmp((const char*)sEs, cellstXfs, 13);
		if (otherresult == 0) {
			//セルスタイル一致 read cellstyleXfs

		}

		otherresult = strncmp((const char*)sbor, Xfs, 8);
		if (otherresult == 0) {
			//cellXfs一致 read cellxfs

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
