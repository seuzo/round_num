/*
round_num.jsx
(c)2008-2009 www.seuzo.jp
選択したテキスト中の数字を丸数字などに変換します。

2008-09-19	ver.0.1	とりあえず
2008-09-20	ver.0.2	2桁以上の選択文字がストーリーの最後にある時、処理が失敗するのを修正した。文字列「111」の最初の「11」を選択しているとき、「⑪⑪」と変換されてしまうのを修正した。
2009-05-26	ver.0.3	InDesign CS4対応版。選択しているテキストが数字だけでなく、他の文字列を選択していても、選択文字列中の数字列を変換できるようにした。
2011-06-24  ver.0.4 合成フォントの対応について、milligrammeさんに修正していただきました。 https://github.com/milligramme/round_num/commit/92dc8fdfe7662d386fc928ab8e4ef0316ef2af4d ありがとうございます。合成フォントを選択している時は、フォントを漢字フォントに変更します。InDesign CS5にて動作確認。表組みやセルを選択していても置換できるようにした。
*/

////////////////////////////////////////////設定
#target "InDesign"
var my_report = true;//処理の最後にレポートダイアログを表示するかどうか


////////////////////////////////////////////エラー処理 
function myerror(mess) { 
  if (arguments.length > 0) { alert(mess); }
  exit();
}


////////////////////////////////////////////ラジオダイアログ
/*
myTitle	ダイアログ（バー）のタイトル
myPrompt	メッセージ
myList	ラジオボタンに展開するリスト

result	選択したリスト番号
*/
function radioDialog(my_title, my_prompt, my_list){
	var my_dialog = app.dialogs.add({name:my_title, canCancel:true});
	with(my_dialog) {
		with(dialogColumns.add()) {
			// プロンプト
			staticTexts.add({staticLabel:my_prompt});
			with (borderPanels.add()) {
				var my_radio_group = radiobuttonGroups.add();
				with (my_radio_group) {
					for (var i = 0; i < my_list.length; i++){
						if (i === 0) {
							radiobuttonControls.add({staticLabel:my_list[i], checkedState:true});
						} else {
						radiobuttonControls.add({staticLabel:my_list[i]});
						}
					}
				}
			}
		}
	}


	if (my_dialog.show() === true) {
		var ans = my_radio_group.selectedButton;
		//正常にダイアログを片付ける
		my_dialog.destroy();
		//選択したアイテムの番号を返す
		return ans;
	} else {
		// ユーザが「キャンセル」をクリックしたので、メモリからダイアログボックスを削除
		my_dialog.destroy();
	}
}

////////////////////////////////////////////正規表現検索
//正規表現で検索して、ヒットオブジェクトを（お尻から）返すだけ
function my_regex(my_range_obj, my_find_str) {
        //検索の初期化
        app.findGrepPreferences = NothingEnum.nothing;
        app.changeGrepPreferences = NothingEnum.nothing;
        //検索オプション
        app.findChangeGrepOptions.includeLockedLayersForFind = false;//ロックされたレイヤーをふくめるかどうか
        app.findChangeGrepOptions.includeLockedStoriesForFind = false;//ロックされたストーリーを含めるかどうか
        app.findChangeGrepOptions.includeHiddenLayers = false;//非表示レイヤーを含めるかどうか
        app.findChangeGrepOptions.includeMasterPages = false;//マスターページを含めるかどうか
        app.findChangeGrepOptions.includeFootnotes = false;//脚注を含めるかどうか
        app.findChangeGrepOptions.kanaSensitive = true;//カナを区別するかどうか
        app.findChangeGrepOptions.widthSensitive = true;//全角半角を区別するかどうか

        app.findGrepPreferences.findWhat = my_find_str;//検索文字の設定
        //app.changeGrepPreferences.changeTo = my_change_str;//置換文字の設定
        return my_range_obj.findGrep(true);//検索の実行（reverse）
}

////////////////////////////////////////////字形検索置換
function change_glyph(my_range_obj, find_font, find_gid, change_font, change_gid) {
	var my_doc = app.activeDocument;
	app.findGlyphPreferences = NothingEnum.nothing;
	app.changeGlyphPreferences = NothingEnum.nothing;
	app.findGlyphPreferences.appliedFont = find_font;
	app.changeGlyphPreferences.appliedFont = change_font;
	app.findGlyphPreferences.glyphID = find_gid;
	app.changeGlyphPreferences.glyphID = change_gid;
	var my_result = my_range_obj.changeGlyph ();
	return my_result;
}



////////////////////////////////////////////以下メイン実行
////////////////まずは選択しているもののチェック
if (app.documents.length === 0) {myerror("ドキュメントが開かれていません")}
var my_doc = app.documents[0];
if (my_doc.selection.length === 0) {myerror("テキストを選択してください")}
var my_selection = my_doc.selection[0];
var my_class =my_selection.reflect.name;
my_class = "Text, TextColumn, Story, Paragraph, Line, Word, Character, TextStyleRange, Table, Cell".match(my_class);
if (my_class === null) {myerror("テキストを選択してください")}

var hit_obj = my_regex(my_selection, "[0-9,.]*[0-9]+");//数字列の検索
var target_obj = new Array();//ターゲットとなるオブジェクトの配列
for (var i = 0; i< hit_obj.length; i++) {
	var tmp_str = hit_obj[i].contents;
	if ((tmp_str.match(/[,.]/) === null) && (tmp_str.match(/^(100|\d\d?)$/) !== null)) {//数字列にカンマやピリオドが含まれておらず、かつ、0〜100の数字列ならば
		target_obj.push(hit_obj[i]);//ターゲットとする
	}
}
if (target_obj.length === 0) {myerror("変換可能な数字はありませんでした")}


////////////////検索前処理
//処理の選択ダイアログ
var myList = ["丸数字にする　①", 
"白抜き丸数字にする　❷", 
"四角内数字にする", 
"四角（ラウンド）内数字にする", 
"黒四角内白抜き数字にする", 
"黒四角（ラウンド）内白抜き数字にする", 
"括弧内数字にする"];
app.scriptPreferences.userInteractionLevel = UserInteractionLevels.interactWithAll;
var ans_int = radioDialog("round_num", "選択した数字テキストを丸数字などに変換します。処理を選択してください\n", myList);

//CIDテーブルのセット
var my_tbl = new Array();//my_tblはCID番号の配列。0番地から100番地までが「0」〜「100」の字形の対応し、101番地から111番地までが「00」〜「09」の字形の対応している。
if (ans_int === 0) {//丸数字にする　①
	my_tbl = [8224, 7555, 7556, 7557, 7558, 7559, 7560, 7561, 7562, 7563, 7564, 7565, 7566, 7567, 7568, 7569, 7570, 7571, 7572, 7573, 7574, 8091, 8102, 8103, 8104, 8105, 8106, 8107, 8108, 8109, 8110, 8111, 10244, 10245, 10246, 10247, 10248, 10249, 10250, 10251, 10252, 10253, 10254, 10255, 10256, 10257, 10258, 10259, 10260, 10261, 10262, 10263, 10264, 10265, 10266, 10267, 10268, 10269, 10270, 10271, 10272, 10273, 10274, 10275, 10276, 10277, 10278, 10279, 10280, 10281, 10282, 10283, 10284, 10285, 10286, 10287, 10288, 10289, 10290, 10291, 10292, 10293, 10294, 10295, 10296, 10297, 10298, 10299, 10300, 10301, 10302, 10303, 10304, 10305, 10306, 10307, 10308, 10309, 10310, 10311, 10312, 10234, 10235, 10236, 10237, 10238, 10239, 10240, 10241, 10242, 10243];
} else if (ans_int === 1) {//白抜き丸数字にする　❷
	my_tbl = [10503, 8286, 8287, 8288, 8289, 8290, 8291, 8292, 8293, 8294, 10514, 10515, 10516, 10517, 10518, 10519, 10520, 10521, 10522, 10523, 10524, 10525, 10526, 10527, 10528, 10529, 10530, 10531, 10532, 10533, 10534, 10535, 10536, 10537, 10538, 10539, 10540, 10541, 10542, 10543, 10544, 10545, 10546, 10547, 10548, 10549, 10550, 10551, 10552, 10553, 10554, 10555, 10556, 10557, 10558, 10559, 10560, 10561, 10562, 10563, 10564, 10565, 10566, 10567, 10568, 10569, 10570, 10571, 10572, 10573, 10574, 10575, 10576, 10577, 10578, 10579, 10580, 10581, 10582, 10583, 10584, 10585, 10586, 10587, 10588, 10589, 10590, 10591, 10592, 10593, 10594, 10595, 10596, 10597, 10598, 10599, 10600, 10601, 10602, 10603, 10604, 10504, 10505, 10506, 10507, 10508, 10509, 10510, 10511, 10512, 10513];
} else if (ans_int === 2) {//四角内数字にする
	my_tbl = [10764, 10766, 10768, 10770, 10772, 10774, 10776, 10778, 10780, 10782, 10784, 10785, 10786, 10787, 10788, 10789, 10790, 10791, 10792, 10793, 10794, 10795, 10796, 10797, 10798, 10799, 10800, 10801, 10802, 10803, 10804, 10805, 10806, 10807, 10808, 10809, 10810, 10811, 10812, 10813, 10814, 10815, 10816, 10817, 10818, 10819, 10820, 10821, 10822, 10823, 10824, 10825, 10826, 10827, 10828, 10829, 10830, 10831, 10832, 10833, 10834, 10835, 10836, 10837, 10838, 10839, 10840, 10841, 10842, 10843, 10844, 10845, 10846, 10847, 10848, 10849, 10850, 10851, 10852, 10853, 10854, 10855, 10856, 10857, 10858, 10859, 10860, 10861, 10862, 10863, 10864, 10865, 10866, 10867, 10868, 10869, 10870, 10871, 10872, 10873, 10874, 10765, 10767, 10769, 10771, 10773, 10775, 10777, 10779, 10781, 10783];
} else if (ans_int === 3) {//四角（ラウンド）内数字にする
	my_tbl = [11307, 11309, 11311, 11313, 11315, 11317, 11319, 11321, 11323, 11325, 11327, 11328, 11329, 11330, 11331, 11332, 11333, 11334, 11335, 11336, 11337, 11338, 11339, 11340, 11341, 11342, 11343, 11344, 11345, 11346, 11347, 11348, 11349, 11350, 11351, 11352, 11353, 11354, 11355, 11356, 11357, 11358, 11359, 11360, 11361, 11362, 11363, 11364, 11365, 11366, 11367, 11368, 11369, 11370, 11371, 11372, 11373, 11374, 11375, 11376, 11377, 11378, 11379, 11380, 11381, 11382, 11383, 11384, 11385, 11386, 11387, 11388, 11389, 11390, 11391, 11392, 11393, 11394, 11395, 11396, 11397, 11398, 11399, 11400, 11401, 11402, 11403, 11404, 11405, 11406, 11407, 11408, 11409, 11410, 11411, 11412, 11413, 11414, 11415, 11416, 11417, 11308, 11310, 11312, 11314, 11316, 11318, 11320, 11322, 11324, 11326];
} else if (ans_int === 4) {//黒四角内白抜き数字にする
	my_tbl =[11037, 11039, 11041, 11043, 11045, 11047, 11049, 11051, 11053, 11055, 11057, 11058, 11059, 11060, 11061, 11062, 11063, 11064, 11065, 11066, 11067, 11068, 11069, 11070, 11071, 11072, 11073, 11074, 11075, 11076, 11077, 11078, 11079, 11080, 11081, 11082, 11083, 11084, 11085, 11086, 11087, 11088, 11089, 11090, 11091, 11092, 11093, 11094, 11095, 11096, 11097, 11098, 11099, 11100, 11101, 11102, 11103, 11104, 11105, 11106, 11107, 11108, 11109, 11110, 11111, 11112, 11113, 11114, 11115, 11116, 11117, 11118, 11119, 11120, 11121, 11122, 11123, 11124, 11125, 11126, 11127, 11128, 11129, 11130, 11131, 11132, 11133, 11134, 11135, 11136, 11137, 11138, 11139, 11140, 11141, 11142, 11143, 11144, 11145, 11146, 11147, 11038, 11040, 11042, 11044, 11046, 11048, 11050, 11052, 11054, 11056];
} else if (ans_int === 5) {//黒四角（ラウンド）内白抜き数字にする
	my_tbl =[11576, 11578, 11580, 11582, 11584, 11586, 11588, 11590, 11592, 11594, 11596, 11597, 11598, 11599, 11600, 11601, 11602, 11603, 11604, 11605, 11606, 11607, 11608, 11609, 11610, 11611, 11612, 11613, 11614, 11615, 11616, 11617, 11618, 11619, 11620, 11621, 11622, 11623, 11624, 11625, 11626, 11627, 11628, 11629, 11630, 11631, 11632, 11633, 11634, 11635, 11636, 11637, 11638, 11639, 11640, 11641, 11642, 11643, 11644, 11645, 11646, 11647, 11648, 11649, 11650, 11651, 11652, 11653, 11654, 11655, 11656, 11657, 11658, 11659, 11660, 11661, 11662, 11663, 11664, 11665, 11666, 11667, 11668, 11669, 11670, 11671, 11672, 11673, 11674, 11675, 11676, 11677, 11678, 11679, 11680, 11681, 11682, 11683, 11684, 11685, 11686, 11577, 11579, 11581, 11583, 11585, 11587, 11589, 11591, 11593, 11595];
} else if (ans_int === 6) {//括弧内数字にする
	my_tbl =[8227, 8071, 8072, 8073, 8074, 8075, 8076, 8077, 8078, 8079, 8080, 8081, 8082, 8083, 8084, 8085, 8086, 8087, 8088, 8089, 8090, 9894, 9895, 9896, 9897, 9898, 9899, 9900, 9901, 9902, 9903, 9904, 9905, 9906, 9907, 9908, 9909, 9910, 9911, 9912, 9913, 9914, 9915, 9916, 9917, 9918, 9919, 9920, 9921, 9922, 9923, 9924, 9925, 9926, 9927, 9928, 9929, 9930, 9931, 9932, 9933, 9934, 9935, 9936, 9937, 9938, 9939, 9940, 9941, 9942, 9943, 9944, 9945, 9946, 9947, 9948, 9949, 9950, 9951, 9952, 9953, 9954, 9955, 9956, 9957, 9958, 9959, 9960, 9961, 9962, 9963, 9964, 9965, 9966, 9967, 9968, 9969, 9970, 9971, 9972, 9973, 9884, 9885, 9886, 9887, 9888, 9889, 9890, 9891, 9892, 9893];
} else {
	myerror("処理をキャンセルしました");
}

//ターゲットをループ
var my_count = 0;
var tmp_result = "";
for (var i = 0; i < target_obj.length; i++) {
	//target_obj[i].select();
	var my_contents = target_obj[i].contents;//ターゲットの数字（テキスト）
	var my_font = target_obj[i].appliedFont;//ターゲットテキストのフォント
    
    //合成フォントへの対応
    var f_atc = false; //合成フォントかどうかのフラグ
    var my_font_org = my_font;//復帰する時のフォント
    if (my_font.fontType === FontTypes.ATC) {
        f_atc = true;
        var target_font_name = my_font.fontFamily;//合成フォント名
        var comp_ent = my_doc.compositeFonts.item(target_font_name).compositeFontEntries;//合成フォントリスト
        var fname = comp_ent[0].appliedFont + "\t" + comp_ent[0].fontStyle;//合成フォント中の「漢字」フォントを使用
        my_font = app.fonts.item(fname);
        target_obj[i].appliedFont = my_font;
    }
	
	//添字のスライド
	var my_suffix = 0;
	if (my_contents.match(/^0\d$/)) {//00〜09の時
		my_suffix = 101 + parseInt (my_contents);
	} else {
		my_suffix = parseInt (my_contents);
	}

	//選択文字列の変換
	target_obj[i].contents = "1";//強制的に選択文字を「1」にする

	//字形検索置換
	try {
		tmp_result = change_glyph(target_obj[i], my_font, 18, my_font, my_tbl[my_suffix]);
		if (tmp_result === "") {
			target_obj[i].contents = my_contents;
		} else {
			my_count++;
		}
        //if ( f_atc ) {target_obj[i].appliedFont = my_font_org; }//合成フォントのフラグが立っていたら、フォントを元の合成フォントに戻す。★overwriteしないとエラーになる（保留）
	} catch(e) {
        $.writeln(e);
		target_obj[i].contents = my_contents;
	}
}

//結果レポート
if (my_report) {
	if (my_count === 0) {
		alert ("変更箇所はありませんでした。字形をもたないフォントである可能性があります");
	} else if (target_obj.length > my_count) {
		var failure_count = target_obj.length - my_count;
		alert (target_obj.length + "箇所中、" + my_count + "箇所の数字を変換しました\n失敗した" + failure_count + "箇所については字形をもたないフォントである可能性があります");
	} else {
		alert ( my_count + "箇所の数字を変換しました");
	}
}
