Attribute VB_Name = "ShiminCorrectionModule"
'
' 市民大会のプログラム補正
'
Sub 市民プログラム補正()

    Call EventChange(False)

    Sheets("エントリー一覧").Activate
    ' 1 中学女子4×50Mメドレーリレー
    Call ModCell("F2", "8")
    Call ModCell("G2", "石井美優弥．矢口優依奈．栃原美菜．大内彩乃")
    Call ModCell("F3", "9")
    Call ModCell("G3", "高橋絵梨奈．山田彩夏．慶長彩奈．熊谷柚奈")
    Call ModCell("F4", "6")
    Call ModCell("G4", "野村美友．今岡杏瞳．高橋かいり．水口果音")
    Call ModCell("F5", "7")
    Call ModCell("G5", "荒川美南．福田恋生．緒形優海．奥山碧")
    Call ModCell("F6", "5")
    Call ModCell("G6", "萩野谷和奏．小林美蘭．竹村夏芽．小林紗菜")
    Call ModCell("F7", "4")
    Call ModCell("G8", "鈴木日和子．三浦真帆．高橋くるみ．高野慶")
    Call ModCell("F9", "5")
    Call ModCell("G9", "鈴木梨緒．間瀨小桜．勝山りお．村田美咲")
    Call ModCell("C10", "2")
    Call ModCell("F10", "8")
    Call ModCell("G10", "斎藤陽子．宮崎明子．宮浦めぐみ．鈴木慶子")
    Call ModCell("F11", "8")
    Call ModCell("G11", "横山孔明．小澤優峨．清水乃弥．森出晃弘")
    Call ModCell("F12", "7")
    Call ModCell("G12", "脇凛太朗．河野光太．谷口皇真．新本修吾")
    Call ModCell("F13", "6")
    Call ModCell("G13", "皆越継太．後藤馨諒．畑中秀介．岩瀬瑛")
    Call ModCell("F14", "9")
    Call ModCell("G14", "佐藤瞬．相川晴城．加我洋太朗．山根太郎")
    Call ModCell("F15", "5")
    Call ModCell("G15", "冨士川瑛太．長利幸多．田中優人．早﨑紀信")
    Call ModCell("F16", "4")
    Call ModCell("G16", "久野海瑠．半田啓人．髙岡蒼矢．長尾圭一郎")
    Call ModCell("F17", "3")
    Call ModCell("G17", "福岡空．對馬葵．高板俊太．水島快")
    Call ModCell("F18", "4")
    Call ModCell("G18", "藁谷祐人．佐野翔哉．二本木奨．田中優貴大")
    Call ModCell("F19", "5")
    Call ModCell("G19", "前田茅サニー．河上怜音．田原海靖．鈴木大貴")
    Call ModCell("F20", "6")
    Call ModCell("G20", "村瀬波輝．近江颯太．佐野翔大．菅野柊")
    Call ModCell("F21", "7")
    Call ModCell("G21", "鈴木公大．渡辺凌．神山渚於．松本昴大")
    Call ModCell("C22", "5")
    Call ModCell("F22", "8")
    Call ModCell("G22", "鈴木治輝．角田祐樹．高橋篤．宮川浩司")
    Call ModCell("C23", "5")
    Call ModCell("F23", "9")
    Call ModCell("G23", "鈴木修平．板越秀介．前徳直和．上原優士")
    
    ' 7－9　同一レース
    Call ModCell(SearchCell(7, "八巻　玲香", "レーン"), "3")
    Call ModCell(SearchCell(7, "谷本　夏子", "レーン"), "4")
    Call ModCell(SearchCell(7, "野村　美友", "レーン"), "5")
    Call ModCell(SearchCell(7, "竹村　夏芽", "レーン"), "6")
    Call ModCell(SearchCell(7, "福田　恋生", "レーン"), "7")
    Call ModCell(SearchCell(9, "菅谷　幸江", "レースNo"), "7")
    Call ModCell(SearchCell(9, "菅谷　幸江", "レーン"), "8")
    
    ' 13
    Call ModCell(SearchCell(13, "栃原　美菜", "レーン"), "8")
    Call ModCell(SearchCell(13, "笠原　吏桜", "レーン"), "5")
    Call ModCell(SearchCell(13, "大内　彩乃", "レーン"), "4")

    ' 17
    Call ModCell(SearchCell(17, "田中　優貴大", "レーン"), "3")
    Call ModCell(SearchCell(17, "小澤　航太", "レーン"), "4")
    Call ModCell(SearchCell(17, "山内　政虎", "レーン"), "5")
    Call ModCell(SearchCell(17, "平田　直大", "レーン"), "6")
    Call ModCell(SearchCell(17, "光田　侑吾", "レーン"), "7")

    ' 18
    Call ModCell(SearchCell(18, "岩井　太造", "レースNo"), "14")
    Call ModCell(SearchCell(18, "岩井　太造", "レーン"), "9")

    ' 19
    Call ModCell(SearchCell(17, "石井　美優弥", "レーン"), "4")
    Call ModCell(SearchCell(17, "降旗　千瑛", "レーン"), "5")
    Call ModCell(SearchCell(17, "庄司　晏樺", "レーン"), "6")

    ' 20
    Call ModCell(SearchCell(20, "鈴木　梨緒", "レースNo"), "16")
    Call ModCell(SearchCell(20, "鈴木　梨緒", "レーン"), "8")

    ' 34
    Call ModCell(SearchCell(34, "竹山　海成", "レーン"), "5")
    Call ModCell(SearchCell(34, "伊藤　吏琥", "レースNo"), "25")
    Call ModCell(SearchCell(34, "伊藤　吏琥", "組"), "1")
    Call ModCell(SearchCell(34, "伊藤　吏琥", "レーン"), "6")
    Call ModCell(SearchCell(34, "三田　駿斗", "レーン"), "7")
    Call ModCell(SearchCell(34, "竹林　侑弥", "レーン"), "8")

    ' 37
    Call ModCell(SearchCell(37, "野村　美友", "レースNo"), "30")
    Call ModCell(SearchCell(37, "野村　美友", "組"), "1")
    Call ModCell(SearchCell(37, "野村　美友", "レーン"), "6")
    Call ModCell(SearchCell(37, "奥山　　碧", "レースNo"), "30")
    Call ModCell(SearchCell(37, "奥山　　碧", "組"), "1")
    Call ModCell(SearchCell(37, "奥山　　碧", "レーン"), "8")
    Call ModCell(SearchCell(37, "前田　成美", "レーン"), "4")
    Call ModCell(SearchCell(37, "熊谷　柚奈", "レーン"), "5")
    Call ModCell(SearchCell(37, "緒方　優海", "レーン"), "6")
    Call ModCell(SearchCell(37, "小林　紗奈", "レーン"), "7")
    Call ModCell(SearchCell(37, "小林　紗菜", "レーン"), "8")
    Call ModCell(SearchCell(37, "竹見　真優", "レースNo"), "31")
    Call ModCell(SearchCell(37, "竹見　真優", "組"), "2")
    Call ModCell(SearchCell(37, "竹見　真優", "レーン"), "9")

    ' 38-39
    Call ModCell(SearchCell(38, "鈴木　日和子", "レーン"), "4")
    Call ModCell(SearchCell(38, "荻野　澪奈", "レーン"), "5")
    Call ModCell(SearchCell(38, "高橋　くるみ", "レーン"), "6")
    Call ModCell(SearchCell(38, "多田　かの", "レーン"), "7")
    Call ModCell(SearchCell(39, "鈴木　慶子", "レースNo"), "32")
    Call ModCell(SearchCell(39, "鈴木　慶子", "レーン"), "9")

    ' 40
    Call ModCell(SearchCell(40, "田中　雅貴", "レーン"), "5")
    Call ModCell(SearchCell(40, "新本　修吾", "レースNo"), "34")
    Call ModCell(SearchCell(40, "新本　修吾", "組"), "1")
    Call ModCell(SearchCell(40, "新本　修吾", "レーン"), "6")
    Call ModCell(SearchCell(40, "守谷　柊杜", "レーン"), "7")
    Call ModCell(SearchCell(40, "渡邊　不二道", "レースNo"), "34")
    Call ModCell(SearchCell(40, "渡邊　不二道", "組"), "1")
    Call ModCell(SearchCell(40, "渡邊　不二道", "レーン"), "8")
    Call ModCell(SearchCell(40, "鎌田　竜我", "レースNo"), "35")
    Call ModCell(SearchCell(40, "鎌田　竜我", "組"), "2")
    Call ModCell(SearchCell(40, "鎌田　竜我", "レーン"), "3")
    Call ModCell(SearchCell(40, "竹山　海成", "レーン"), "4")
    Call ModCell(SearchCell(40, "早﨑　紀信", "レーン"), "5")
    Call ModCell(SearchCell(40, "川本　秀斗", "レーン"), "6")
    Call ModCell(SearchCell(40, "田中　優人", "レーン"), "7")
    Call ModCell(SearchCell(40, "富士川　瑛太", "レーン"), "8")
    Call ModCell(SearchCell(40, "畠中　優之介", "レースNo"), "35")
    Call ModCell(SearchCell(40, "畠中　優之介", "組"), "2")
    Call ModCell(SearchCell(40, "畠中　優之介", "レーン"), "9")
    Call ModCell(SearchCell(40, "中野　叶大", "レースNo"), "36")
    Call ModCell(SearchCell(40, "中野　叶大", "組"), "3")
    Call ModCell(SearchCell(40, "中野　叶大", "レーン"), "3")
    Call ModCell(SearchCell(40, "長尾　圭一郎", "レーン"), "4")
    Call ModCell(SearchCell(40, "北川　優翔", "レーン"), "5")
    Call ModCell(SearchCell(40, "松浦　海音", "レーン"), "6")
    Call ModCell(SearchCell(40, "加我　洋太朗", "レーン"), "7")
    Call ModCell(SearchCell(40, "川嶋　康晟", "レーン"), "8")
    Call ModCell(SearchCell(40, "小澤　翔太", "レースNo"), "36")
    Call ModCell(SearchCell(40, "小澤　翔太", "組"), "3")
    Call ModCell(SearchCell(40, "小澤　翔太", "レーン"), "9")
    
    
    Call ModCell(SearchCell(41, "鈴木　大貴", "レースNo"), "37")
    Call ModCell(SearchCell(41, "鈴木　大貴", "組"), "1")
    Call ModCell(SearchCell(41, "鈴木　大貴", "レーン"), "5")
    Call ModCell(SearchCell(41, "田原　海靖", "レーン"), "6")
    Call ModCell(SearchCell(41, "野邊　悠河", "レースNo"), "37")
    Call ModCell(SearchCell(41, "野邊　悠河", "組"), "1")
    Call ModCell(SearchCell(41, "野邊　悠河", "レーン"), "7")
    Call ModCell(SearchCell(41, "入口　景仁", "レースNo"), "37")
    Call ModCell(SearchCell(41, "入口　景仁", "組"), "1")
    Call ModCell(SearchCell(41, "入口　景仁", "レーン"), "8")
    
    Call ModCell(SearchCell(41, "森川　真吾", "レースNo"), "38")
    Call ModCell(SearchCell(41, "森川　真吾", "組"), "2")
    Call ModCell(SearchCell(41, "森川　真吾", "レーン"), "4")
    Call ModCell(SearchCell(41, "佐藤　美志", "レーン"), "5")
    Call ModCell(SearchCell(41, "佐久間　慎", "レーン"), "6")
    Call ModCell(SearchCell(41, "菅野　　柊", "レーン"), "7")
    Call ModCell(SearchCell(41, "鯉渕　航希", "レースNo"), "38")
    Call ModCell(SearchCell(41, "鯉渕　航希", "組"), "2")
    Call ModCell(SearchCell(41, "鯉渕　航希", "レーン"), "8")
    
    Call ModCell(SearchCell(41, "枝広　一希", "レースNo"), "39")
    Call ModCell(SearchCell(41, "枝広　一希", "組"), "3")
    Call ModCell(SearchCell(41, "枝広　一希", "レーン"), "3")
    Call ModCell(SearchCell(41, "福岡　　空", "レースNo"), "39")
    Call ModCell(SearchCell(41, "福岡　　空", "組"), "3")
    Call ModCell(SearchCell(41, "福岡　　空", "レーン"), "4")
    Call ModCell(SearchCell(41, "川村　怜維", "レースNo"), "39")
    Call ModCell(SearchCell(41, "川村　怜維", "組"), "3")
    Call ModCell(SearchCell(41, "川村　怜維", "レーン"), "5")
    Call ModCell(SearchCell(41, "後藤　憲亮", "レーン"), "6")
    Call ModCell(SearchCell(41, "鈴木　隼矢", "レーン"), "7")
    Call ModCell(SearchCell(41, "三谷　大賀", "レーン"), "8")
    Call ModCell(SearchCell(41, "堀江　和希", "レーン"), "9")
    
    Call ModCell(SearchCell(41, "村瀬　波輝", "レースNo"), "40")
    Call ModCell(SearchCell(41, "村瀬　波輝", "組"), "4")
    Call ModCell(SearchCell(41, "村瀬　波輝", "レーン"), "3")
    Call ModCell(SearchCell(41, "田中　優貴大", "レーン"), "4")
    Call ModCell(SearchCell(41, "佐野　翔哉", "レーン"), "8")

    Call ModCell(SearchCell(43, "川口　夏海", "レースNo"), "41")
    Call ModCell(SearchCell(43, "川口　夏海", "組"), "1")
    Call ModCell(SearchCell(43, "川口　夏海", "レーン"), "6")
    Call ModCell(SearchCell(43, "高橋　絵梨奈", "レーン"), "7")
    Call ModCell(SearchCell(43, "鈴木　珠美", "レーン"), "8")

    Call ModCell(SearchCell(45, "三上　美佐枝", "レーン"), "3")
    Call ModCell(SearchCell(45, "川上　京子", "レーン"), "4")
    Call ModCell(SearchCell(45, "柴　　恵子", "レーン"), "5")
    Call ModCell(SearchCell(45, "高橋　和枝", "レーン"), "6")
    Call ModCell(SearchCell(45, "鈴木　愛子", "レーン"), "7")
    Call ModCell(SearchCell(45, "衛藤　由里子", "レーン"), "8")

    Call ModCell(SearchCell(48, "岡田　　彰", "レーン"), "4")
    Call ModCell(SearchCell(48, "", "レーン"), "5")
    Call ModCell(SearchCell(48, "", "レーン"), "6")
    Call ModCell(SearchCell(48, "", "レーン"), "7")
    Call ModCell(SearchCell(48, "", "レーン"), "8")
    Call ModCell(SearchCell(48, "", "レーン"), "9")


    Call ModCell(SearchCell(48, "", "レースNo"), "48")
    Call ModCell(SearchCell(48, "", "組"), "1")
    Call ModCell(SearchCell(48, "", "レーン"), "3")

    Call EventChange(True)
    
    ' ブックの保存
    ActiveWorkbook.Save
End Sub

'
' 令和元年度のプログラムの記録を入れる
'
Sub 市民記録入力()
    Sheets("記録画面").Select
    Call SetRace(1, 1)
    Call SetLean(1, 3, "34709")
    Call SetLean(2, 4, "25977")
    Call SetLean(3, 6, "")
    Call SetLean(4, 8, "30973")
    Call 登録
    Call 順位決定
    Call 初期化
    
    ' ブックの保存
    ActiveWorkbook.Save
End Sub
