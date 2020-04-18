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
    
    ' 2－3　同一レース
    Call ModCell("F9", "5")
    Call ModCell("G9", "鈴木梨緒．間瀨小桜．勝山りお．村田美咲")
    Call ModCell("C10", "20")
    Call ModCell("F10", "8")
    Call ModCell("G10", "斎藤陽子．宮崎明子．宮浦めぐみ．鈴木慶子")
    
    ' 4 中学 男子 4×50M メドレーリレー
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
    
    ' 5－6　同一レース
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
    ' 6
    Call ModCell("C22", "50")
    Call ModCell("F22", "8")
    Call ModCell("G22", "鈴木治輝．角田祐樹．高橋篤．宮川浩司")
    Call ModCell("C23", "50")
    Call ModCell("F23", "9")
    Call ModCell("G23", "鈴木修平．板越秀介．前徳直和．上原優士")
    
    ' 7－9　同一レース
    Call ModCell(SearchCell(7, "八巻　玲香", "レーン"), "3")
    Call ModCell(SearchCell(7, "谷本　夏子", "レーン"), "4")
    Call ModCell(SearchCell(7, "野村　美友", "レーン"), "5")
    Call ModCell(SearchCell(7, "竹村　夏芽", "レーン"), "6")
    Call ModCell(SearchCell(7, "福田　恋生", "レーン"), "7")
    ' 9
    Call ModCell(SearchCell(9, "菅谷　幸江", "レースNo"), "70")
    Call ModCell(SearchCell(9, "菅谷　幸江", "レーン"), "9")
    
    ' 17-18 同一レース
    Call ModCell(SearchCell(17, "田中　優貴大", "レーン"), "3")
    Call ModCell(SearchCell(17, "小澤　航太", "レーン"), "4")
    Call ModCell(SearchCell(17, "山内　政虎", "レーン"), "5")
    Call ModCell(SearchCell(17, "光田　侑吾", "レーン"), "6")
    Call ModCell(SearchCell(17, "平田　直大", "レーン"), "7")
    ' 18
    Call ModCell(SearchCell(18, "岩井　太造", "レースNo"), "140")
    Call ModCell(SearchCell(18, "岩井　太造", "レーン"), "9")

    ' 19-20 同一レース
    Call ModCell(SearchCell(19, "石井　美優弥", "レーン"), "4")
    Call ModCell(SearchCell(19, "降旗　千瑛", "レーン"), "5")
    Call ModCell(SearchCell(19, "庄司　晏樺", "レーン"), "6")
    ' 20
    Call ModCell(SearchCell(20, "鈴木　梨緒", "レースNo"), "160")
    Call ModCell(SearchCell(20, "鈴木　梨緒", "レーン"), "8")

    ' 38-39 同一レース
    Call ModCell(SearchCell(38, "鈴木　日和子", "レーン"), "4")
    Call ModCell(SearchCell(38, "荻野　澪奈", "レーン"), "5")
    Call ModCell(SearchCell(38, "高橋　くるみ", "レーン"), "6")
    Call ModCell(SearchCell(38, "多田　かの", "レーン"), "7")
    ' 39
    Call ModCell(SearchCell(39, "鈴木　慶子", "レースNo"), "320")
    Call ModCell(SearchCell(39, "鈴木　慶子", "レーン"), "9")

    ' 50-51 同一レース
    Call ModCell(SearchCell(50, "西田　陽香", "レーン"), "3")
    ' 51
    Call ModCell(SearchCell(51, "菅谷　幸江", "レースNo"), "510")
    Call ModCell(SearchCell(51, "菅谷　幸江", "レーン"), "6")
    Call ModCell(SearchCell(51, "宮浦　めぐみ", "レースNo"), "510")
    Call ModCell(SearchCell(51, "宮浦　めぐみ", "レーン"), "7")
    Call ModCell(SearchCell(51, "佐野　華子", "レースNo"), "510")
    Call ModCell(SearchCell(51, "佐野　華子", "レーン"), "8")
    Call ModCell(SearchCell(51, "斎藤　陽子", "レースNo"), "510")
    Call ModCell(SearchCell(51, "斎藤　陽子", "レーン"), "9")

   ' 52
    Call ModCell(SearchCell(52, "清水　乃弥", "レーン"), "4")
    Call ModCell(SearchCell(52, "岩瀬　　瑛", "レースNo"), "540")
    Call ModCell(SearchCell(52, "岩瀬　　瑛", "組"), "2")
    Call ModCell(SearchCell(52, "岩瀬　　瑛", "レーン"), "7")
    Call ModCell(SearchCell(52, "山岸　直人", "レーン"), "8")

   ' 54
    Call ModCell(SearchCell(54, "小林　慶雄", "レーン"), "4")
    Call ModCell(SearchCell(54, "丸茂　健二", "レーン"), "5")
    Call ModCell(SearchCell(54, "高橋　俊也", "レースNo"), "570")
    Call ModCell(SearchCell(54, "高橋　俊也", "組"), "1")
    Call ModCell(SearchCell(54, "高橋　俊也", "レーン"), "7")

   ' 60 1組(76)
    Call ModCell(SearchCell(60, "山本　眞人", "レーン"), "3")
    Call ModCell(SearchCell(60, "小林　慶雄", "レーン"), "4")
    Call ModCell(SearchCell(60, "三浦　達也", "レーン"), "5")
    Call ModCell(SearchCell(60, "小杉　邦洋", "レーン"), "7")
    Call ModCell(SearchCell(60, "青木　興一", "レーン"), "8")
    Call ModCell(SearchCell(60, "鶴田　三男", "レーン"), "9")
    Call ModCell(SearchCell(60, "渡辺　峰男", "レースNo"), "760")
    Call ModCell(SearchCell(60, "渡辺　峰男", "組"), "1")
    Call ModCell(SearchCell(60, "渡辺　峰男", "レーン"), "6")

   ' 60 2組(77)
    Call ModCell(SearchCell(60, "丸茂　健二", "レーン"), "5")
    Call ModCell(SearchCell(60, "田川　　宏", "レーン"), "6")
    Call ModCell(SearchCell(60, "堀川　憲雄", "レーン"), "7")

   ' 60 3組(78)
    Call ModCell(SearchCell(60, "安藤　英之", "レースNo"), "780")
    Call ModCell(SearchCell(60, "安藤　英之", "組"), "3")
    Call ModCell(SearchCell(60, "安藤　英之", "レーン"), "4")
    Call ModCell(SearchCell(60, "野村　雅浩", "レースNo"), "780")
    Call ModCell(SearchCell(60, "野村　雅浩", "組"), "3")
    Call ModCell(SearchCell(60, "野村　雅浩", "レーン"), "5")
    Call ModCell(SearchCell(60, "安部　良彦", "レースNo"), "780")
    Call ModCell(SearchCell(60, "安部　良彦", "組"), "3")
    Call ModCell(SearchCell(60, "安部　良彦", "レーン"), "8")
    
    Call ModCell(SearchCell(60, "高橋　俊也", "組"), "3")
    Call ModCell(SearchCell(60, "高橋　俊也", "レーン"), "6")
    Call ModCell(SearchCell(60, "宮川　浩司", "組"), "3")
    Call ModCell(SearchCell(60, "宮川　浩司", "レーン"), "7")
    Call ModCell(SearchCell(60, "角田　祐樹", "組"), "3")
    Call ModCell(SearchCell(60, "角田　祐樹", "レーン"), "9")

   ' 60 4組(79)
    Call ModCell(SearchCell(60, "草地　　哲", "レースNo"), "790")
    Call ModCell(SearchCell(60, "草地　　哲", "組"), "4")
    Call ModCell(SearchCell(60, "草地　　哲", "レーン"), "4")
    Call ModCell(SearchCell(60, "山口　雄太郎", "レースNo"), "790")
    Call ModCell(SearchCell(60, "山口　雄太郎", "組"), "4")
    Call ModCell(SearchCell(60, "山口　雄太郎", "レーン"), "5")
    Call ModCell(SearchCell(60, "前徳　直和", "レースNo"), "790")
    Call ModCell(SearchCell(60, "前徳　直和", "組"), "4")
    Call ModCell(SearchCell(60, "前徳　直和", "レーン"), "8")
    Call ModCell(SearchCell(60, "伴野　孝輔", "レースNo"), "790")
    Call ModCell(SearchCell(60, "伴野　孝輔", "組"), "4")
    Call ModCell(SearchCell(60, "伴野　孝輔", "レーン"), "9")
    
    Call ModCell(SearchCell(60, "落合　誉通", "レーン"), "6")
    Call ModCell(SearchCell(60, "宮浦　隆徳", "レーン"), "7")

   ' 60 5組(80)
    Call ModCell(SearchCell(60, "土屋　和生", "レースNo"), "791")
    Call ModCell(SearchCell(60, "土屋　和生", "組"), "5")
    Call ModCell(SearchCell(60, "土屋　和生", "レーン"), "4")
    Call ModCell(SearchCell(60, "守谷　智行", "レースNo"), "791")
    Call ModCell(SearchCell(60, "守谷　智行", "組"), "5")
    Call ModCell(SearchCell(60, "守谷　智行", "レーン"), "5")
    Call ModCell(SearchCell(60, "中山　嘉隆", "レースNo"), "791")
    Call ModCell(SearchCell(60, "中山　嘉隆", "組"), "5")
    Call ModCell(SearchCell(60, "中山　嘉隆", "レーン"), "6")
    Call ModCell(SearchCell(60, "橋爪　久司", "レースNo"), "791")
    Call ModCell(SearchCell(60, "橋爪　久司", "組"), "5")
    Call ModCell(SearchCell(60, "橋爪　久司", "レーン"), "7")
    Call ModCell(SearchCell(60, "上原　優士", "レースNo"), "791")
    Call ModCell(SearchCell(60, "上原　優士", "組"), "5")
    Call ModCell(SearchCell(60, "上原　優士", "レーン"), "8")

    ' 61
    Call ModCell(SearchCell(61, "今岡　杏瞳", "レーン"), "4")
    Call ModCell(SearchCell(61, "田中　千瑞", "レーン"), "5")
    Call ModCell(SearchCell(61, "前田　成美", "レースNo"), "800")
    Call ModCell(SearchCell(61, "前田　成美", "レーン"), "7")
    Call ModCell(SearchCell(61, "前田　成美", "組"), "1")

    ' 63
    Call ModCell(SearchCell(63, "竹村　昌子", "レーン"), "5")
    Call ModCell(SearchCell(63, "武石　淳子", "レーン"), "6")

    ' 66
    Call ModCell(SearchCell(66, "板越　秀介", "レーン"), "5")
    Call ModCell(SearchCell(66, "宮浦　隆徳", "レーン"), "6")
    Call ModCell(SearchCell(66, "山口　雄太郎", "レーン"), "7")
    Call ModCell(SearchCell(66, "鈴木　修平", "レーン"), "8")
    Call ModCell(SearchCell(66, "米澤　将克", "レーン"), "9")

    ' 67
    Call ModCell("G444", "高橋和枝．柴恵子．三上美佐枝．山口孝子")
    Call ModCell("G445", "蓑和真理子．川上京子．武石淳子．大村怜子")

    ' 68
    Call ModCell("F446", "5")
    Call ModCell("G446", "大内彩乃．矢口優依奈．栃原美菜．八巻玲香")
    Call ModCell("F447", "9")
    Call ModCell("G447", "慶長彩奈．高橋絵梨奈．熊谷柚奈．山田彩夏")
    Call ModCell("F448", "6")
    Call ModCell("G448", "緒形優海．奥山碧．荒川美南．福田恋生")
    Call ModCell("F449", "8")
    Call ModCell("G449", "野村美友．水口果音．高橋かいり．岩波もも")
    Call ModCell("F450", "7")
    Call ModCell("G450", "小林美蘭．小林紗菜．萩野谷和奏．竹村夏芽")
    Call ModCell("F451", "4")
    
    ' 69
    Call ModCell("G452", "勝山りお．間瀨小桜．村田美咲．福宮友里")
    Call ModCell("G453", "高橋くるみ．高野慶．三浦真帆．鈴木日和子")
    Call ModCell("G454", "鈴木菜緒．市村比南乃．青山友花．髙松紗弥")
    
    ' 70
    Call ModCell("G455", "宮川浩司．高橋篤．鈴木治輝．角田祐樹")
    Call ModCell("G456", "上原優士．前徳直和．伴野孝輔．鈴木修平")
    
    ' 71
    Call ModCell("F457", "8")
    Call ModCell("G457", "小澤優峨．森出晃弘．定森陽人．関大翔")
    Call ModCell("F458", "7")
    Call ModCell("G458", "谷口皇真．渡邉不二道．脇凛太朗．新本修吾")
    Call ModCell("F459", "9")
    Call ModCell("G459", "竹林侑弥．山岸直人．髙橋圭大．松浦海音")
    Call ModCell("F460", "6")
    Call ModCell("G460", "後藤馨諒．皆越継太．畑中秀介．岩瀬瑛")
    Call ModCell("F461", "5")
    Call ModCell("G461", "早﨑紀信．唐澤健太郎．田中優人．長利幸多")
    Call ModCell("F462", "4")
    Call ModCell("G462", "半田啓人．川本秀斗．中野叶大．長尾圭一郎")
    
    ' 72
    Call ModCell("G463", "高坂俊太．堀江和希．水島快．福岡空")
    Call ModCell("G464", "縄田樹．清水翔太．鈴木大翔．河井湧太郎")
    Call ModCell("G465", "早川春翔．前田茅サニー．田原海靖．野邊悠雅")
    Call ModCell("G466", "佐野翔哉．鈴木隼矢．二本木奨．田中優貴大")
    Call ModCell("G467", "松本昴大．神山渚於．山内政虎．鈴木公大")
    Call ModCell("G468", "村瀬波輝．川村怜．菅野柊．佐野翔大")
    Call ModCell("G469", "後藤憲亮．荒木隆弥．渡辺海音．佐久間慎")

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
