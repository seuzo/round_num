/**
 * 	appliedFontで使用されてる合成フォントの
 *	半角数字部分に使われているフォントをappliedFontとして適用する
 */

if (app.documents.length === 0 || app.selection.length === 1) {
	main();
};

function main () {
	var doc = app.documents[0];
	var sel = app.selection[0];
	if(sel.hasOwnProperty('baseline')){
		// if composite font?
		if(sel.appliedFont.fontType === FontTypes.ATC) {
			var target_font_name = sel.appliedFont.fontFamily;
			var comp_ent = doc.compositeFonts.item(target_font_name).compositeFontEntries;
			// number (composite font entry-5)
			var fname = comp_ent[5].appliedFont + "\t" + comp_ent[5].fontStyle;
			try {
				sel.appliedFont = app.fonts.item(fname);
			}
			catch(e){}
		}
	}
}

#include "round_num.jsx"
