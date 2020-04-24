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
    Call SetLean(2, 4, "")
    Call SetLean(3, 5, "24938")
    Call SetLean(4, 6, "22309")
    Call SetLean(5, 7, "22423")
    Call SetLean(6, 8, "25897")
    Call SetLean(7, 9, "23297")
    Call 登録
    Call 順位決定
    Call 初期化
    
    Call SetRace(2, 1)
    Call SetLean(3, 5, "22599")
    Call SetLean(4, 6, "23540")
    Call SetLean(6, 8, "25072")
    Call 登録
    Call 順位決定
    Call 初期化
    
    Call SetRace(4, 1)
    Call SetLean(2, 4, "22424")
    Call SetLean(3, 5, "23780")
    Call SetLean(4, 6, "21010")
    Call SetLean(5, 7, "21811")
    Call SetLean(6, 8, "22916")
    Call SetLean(7, 9, "")
    Call 登録
    Call 順位決定
    Call 初期化
    
    Call SetRace(5, 1)
    Call SetLean(1, 3, "22365")
    Call SetLean(2, 4, "21235")
    Call SetLean(3, 5, "20719")
    Call SetLean(4, 6, "21631")
    Call SetLean(5, 7, "20928")
    Call SetLean(6, 8, "21169")
    Call SetLean(7, 9, "24807")
    Call 登録
    Call 順位決定
    Call 初期化
    
    Call SetRace(7, 1)
    Call SetLean(1, 3, "34435")
    Call SetLean(2, 4, "")
    Call SetLean(3, 5, "24062")
    Call SetLean(4, 6, "30724")
    Call SetLean(5, 7, "30434")
    Call SetLean(7, 9, "35667")
    Call 登録
    Call 順位決定
    Call 初期化
    
    Call SetRace(10, 1)
    Call SetLean(3, 5, "31305")
    Call SetLean(4, 6, "25735")
    Call SetLean(5, 7, "30160")
    Call SetLean(6, 8, "32499")
    Call 登録
    Call 順位決定
    Call 初期化
    
    Call SetRace(11, 1)
    Call SetLean(2, 4, "25004")
    Call SetLean(3, 5, "24581")
    Call SetLean(4, 6, "22740")
    Call SetLean(5, 7, "24042")
    Call SetLean(6, 8, "24881")
    Call 登録
    Call 順位決定
    Call 初期化
    
    Call SetRace(13, 1)
    Call SetLean(1, 3, "23918")
    Call SetLean(2, 4, "30064")
    Call SetLean(3, 5, "24587")
    Call SetLean(4, 6, "22224")
    Call SetLean(5, 7, "23323")
    Call SetLean(6, 8, "31235")
    Call SetLean(7, 9, "31706")
    Call 登録
    Call 順位決定
    Call 初期化
    
    Call SetRace(15, 1)
    Call SetLean(3, 5, "35575")
    Call 登録
    Call 順位決定
    Call 初期化
    
    Call SetRace(16, 1)
    Call SetLean(3, 5, "25363")
    Call SetLean(4, 6, "22796")
    Call SetLean(5, 7, "30755")
    Call SetLean(6, 8, "30751")
    Call 登録
    Call 順位決定
    Call 初期化
    
    Call SetRace(17, 1)
    Call SetLean(1, 3, "22062")
    Call SetLean(2, 4, "")
    Call SetLean(3, 5, "22898")
    Call SetLean(4, 6, "22915")
    Call SetLean(5, 7, "24717")
    Call SetLean(7, 9, "30434")
    Call 登録
    Call 順位決定
    Call 初期化
    
    Call SetRace(19, 1)
    Call SetLean(2, 4, "20830")
    Call SetLean(3, 5, "11333")
    Call SetLean(4, 6, "13980")
    Call SetLean(6, 8, "11386")
    Call 登録
    Call 順位決定
    Call 初期化
    
    Call SetRace(22, 1)
    Call SetLean(2, 4, "20111")
    Call SetLean(3, 5, "12057")
    Call SetLean(4, 6, "11994")
    Call SetLean(5, 7, "11670")
    Call SetLean(6, 8, "12784")
    Call SetLean(7, 9, "14289")
    Call 登録
    Call 順位決定
    Call 初期化
    
    Call SetRace(23, 1)
    Call SetLean(2, 4, "12995")
    Call SetLean(3, 5, "11638")
    Call SetLean(4, 6, "11738")
    Call SetLean(5, 7, "12031")
    Call SetLean(6, 8, "11801")
    Call 登録
    Call 順位決定
    Call 初期化
    
    Call SetRace(25, 1)
    Call SetLean(2, 4, "12640")
    Call SetLean(3, 5, "12018")
    Call SetLean(4, 6, "11138")
    Call SetLean(5, 7, "11324")
    Call SetLean(6, 8, "13674")
    Call 登録
    Call 順位決定
    Call 初期化
    
    Call SetRace(28, 1)
    Call SetLean(2, 4, "14644")
    Call SetLean(3, 5, "10731")
    Call SetLean(4, 6, "10368")
    Call SetLean(5, 7, "10990")
    Call SetLean(6, 8, "11899")
    Call 登録
    Call 順位決定
    Call 初期化
    
    Call SetRace(29, 1)
    Call SetLean(3, 5, "12089")
    Call SetLean(4, 6, "10305")
    Call SetLean(5, 7, "11099")
    Call SetLean(6, 8, "14712")
    Call 登録
    Call 順位決定
    Call 初期化
    
    Call SetRace(31, 1)
    Call SetLean(1, 3, "13799")
    Call SetLean(2, 4, "13871")
    Call SetLean(3, 5, "13620")
    Call SetLean(4, 6, "12619")
    Call SetLean(5, 7, "")
    Call SetLean(6, 8, "14078")
    Call SetLean(7, 9, "14068")
    Call 登録
    Call 順位決定
    Call 初期化
    
    Call SetRace(32, 1)
    Call SetLean(3, 5, "13605")
    Call SetLean(4, 6, "12586")
    Call SetLean(5, 7, "13312")
    Call SetLean(6, 8, "13401")
    Call 登録
    Call 順位決定
    Call 初期化
    
    Call SetRace(34, 1)
    Call SetLean(3, 5, "14140")
    Call SetLean(4, 6, "13239")
    Call SetLean(5, 7, "14228")
    Call SetLean(6, 8, "13390")
    Call 登録
    Call 初期化
    
    Call SetRace(34, 2)
    Call SetLean(2, 4, "13036")
    Call SetLean(3, 5, "12716")
    Call SetLean(4, 6, "11561")
    Call SetLean(5, 7, "12083")
    Call SetLean(6, 8, "12306")
    Call SetLean(7, 9, "13349")
    Call 登録
    Call 順位決定
    Call 初期化
    
    Call SetRace(35, 1)
    Call SetLean(2, 4, "12848")
    Call SetLean(3, 5, "12981")
    Call SetLean(4, 6, "13040")
    Call SetLean(5, 7, "13181")
    Call SetLean(6, 8, "")
    Call 登録
    Call 初期化
    
    Call SetRace(35, 2)
    Call SetLean(1, 3, "12555")
    Call SetLean(2, 4, "12129")
    Call SetLean(3, 5, "12105")
    Call SetLean(4, 6, "11723")
    Call SetLean(5, 7, "12156")
    Call SetLean(6, 8, "12522")
    Call SetLean(7, 9, "13346")
    Call 登録
    Call 順位決定
    Call 初期化
    
    Call SetRace(36, 1)
    Call SetLean(4, 6, "14475")
    Call 登録
    Call 順位決定
    Call 初期化
    
    Call SetRace(37, 1)
    Call SetLean(3, 5, "14081")
    Call SetLean(4, 6, "10553")
    Call SetLean(5, 7, "12134")
    Call SetLean(6, 8, "11015")
    Call 登録
    Call 初期化
    
    Call SetRace(37, 2)
    Call SetLean(2, 4, "12470")
    Call SetLean(3, 5, "11403")
    Call SetLean(4, 6, "10581")
    Call SetLean(5, 7, "10690")
    Call SetLean(6, 8, "12389")
    Call SetLean(7, 9, "")
    Call 登録
    Call 順位決定
    Call 初期化
    
    Call SetRace(38, 1)
    Call SetLean(2, 4, "11752")
    Call SetLean(3, 5, "10969")
    Call SetLean(4, 6, "11287")
    Call SetLean(5, 7, "")
    Call SetLean(7, 9, "14828")
    Call 登録
    Call 順位決定
    Call 初期化
    
    Call SetRace(40, 1)
    Call SetLean(3, 5, "")
    Call SetLean(4, 6, "10507")
    Call SetLean(5, 7, "12790")
    Call SetLean(6, 8, "11032")
    Call 登録
    Call 初期化
    
    Call SetRace(40, 2)
    Call SetLean(1, 3, "12490")
    Call SetLean(2, 4, "12380")
    Call SetLean(3, 5, "12627")
    Call SetLean(4, 6, "11300")
    Call SetLean(5, 7, "11132")
    Call SetLean(6, 8, "12414")
    Call SetLean(7, 9, "11324")
    Call 登録
    Call 初期化
    
    Call SetRace(40, 3)
    Call SetLean(1, 3, "11645")
    Call SetLean(2, 4, "10817")
    Call SetLean(3, 5, "10908")
    Call SetLean(4, 6, "10205")
    Call SetLean(5, 7, "10788")
    Call SetLean(6, 8, "10806")
    Call SetLean(7, 9, "10843")
    Call 登録
    Call 順位決定
    Call 初期化
    
    Call SetRace(41, 1)
    Call SetLean(3, 5, "10357")
    Call SetLean(4, 6, "10609")
    Call SetLean(5, 7, "10417")
    Call SetLean(6, 8, "10665")
    Call 登録
    Call 初期化
    
    Call SetRace(41, 2)
    Call SetLean(2, 4, "12154")
    Call SetLean(3, 5, "11460")
    Call SetLean(4, 6, "11580")
    Call SetLean(5, 7, "11239")
    Call SetLean(6, 8, "12302")
    Call 登録
    Call 初期化
    
    Call SetRace(41, 3)
    Call SetLean(1, 3, "11034")
    Call SetLean(2, 4, "10749")
    Call SetLean(3, 5, "11444")
    Call SetLean(4, 6, "10455")
    Call SetLean(5, 7, "10823")
    Call SetLean(6, 8, "11687")
    Call SetLean(7, 9, "11689")
    Call 登録
    Call 初期化
    
    Call SetRace(41, 4)
    Call SetLean(1, 3, "10528")
    Call SetLean(2, 4, "10443")
    Call SetLean(3, 5, "10141")
    Call SetLean(4, 6, "5455")
    Call SetLean(5, 7, "5813")
    Call SetLean(6, 8, "10248")
    Call SetLean(7, 9, "10338")
    Call 登録
    Call 順位決定
    Call 初期化
    
    Call SetRace(43, 1)
    Call SetLean(3, 5, "4328")
    Call SetLean(4, 6, "")
    Call SetLean(5, 7, "4104")
    Call SetLean(6, 8, "4635")
    Call 登録
    Call 初期化
    
    Call SetRace(43, 2)
    Call SetLean(3, 5, "")
    Call SetLean(4, 6, "3636")
    Call SetLean(5, 7, "3939")
    Call SetLean(6, 8, "4443")
    Call 登録
    Call 順位決定
    Call 初期化
    
    Call SetRace(44, 1)
    Call SetLean(4, 6, "4286")
    Call 登録
    Call 順位決定
    Call 初期化
    
    Call SetRace(45, 1)
    Call SetLean(1, 3, "10063")
    Call SetLean(2, 4, "5022")
    Call SetLean(3, 5, "5786")
    Call SetLean(4, 6, "5627")
    Call SetLean(5, 7, "")
    Call SetLean(6, 8, "4475")
    Call 登録
    Call 順位決定
    Call 初期化
    
    Call SetRace(46, 1)
    Call SetLean(3, 5, "4404")
    Call SetLean(4, 6, "4090")
    Call SetLean(5, 7, "3959")
    Call SetLean(6, 8, "4141")
    Call 登録
    Call 初期化
    
    Call SetRace(46, 2)
    Call SetLean(1, 3, "3900")
    Call SetLean(2, 4, "4019")
    Call SetLean(3, 5, "3293")
    Call SetLean(4, 6, "3099")
    Call SetLean(5, 7, "3526")
    Call SetLean(6, 8, "3525")
    Call SetLean(7, 9, "4026")
    Call 登録
    Call 順位決定
    Call 初期化
    
    Call SetRace(47, 1)
    Call SetLean(3, 5, "3935")
    Call SetLean(4, 6, "3526")
    Call SetLean(5, 7, "3738")
    Call 登録
    Call 順位決定
    Call 初期化
    
    Call SetRace(48, 1)
    Call SetLean(2, 4, "10629")
    Call SetLean(3, 5, "5186")
    Call SetLean(4, 6, "5087")
    Call SetLean(5, 7, "5567")
    Call SetLean(6, 8, "11465")
    Call SetLean(7, 9, "")
    Call 登録
    Call 順位決定
    Call 初期化
    
    Call SetRace(49, 1)
    Call SetLean(3, 5, "4262")
    Call SetLean(4, 6, "3517")
    Call SetLean(5, 7, "3572")
    Call SetLean(6, 8, "4474")
    Call 登録
    Call 初期化
    
    Call SetRace(49, 2)
    Call SetLean(3, 5, "3250")
    Call SetLean(4, 6, "3250")
    Call SetLean(5, 7, "3329")
    Call SetLean(6, 8, "3303")
    Call 登録
    Call 順位決定
    Call 初期化
    
    Call SetRace(50, 1)
    Call SetLean(1, 3, "3665")
    Call SetLean(4, 6, "4607")
    Call SetLean(5, 7, "3348")
    Call SetLean(6, 8, "3724")
    Call SetLean(7, 9, "3393")
    Call 登録
    Call 順位決定
    Call 初期化
    
    Call SetRace(52, 1)
    Call SetLean(3, 5, "3459")
    Call SetLean(4, 6, "3689")
    Call SetLean(5, 7, "3740")
    Call SetLean(6, 8, "3728")
    Call 登録
    Call 初期化
    
    Call SetRace(52, 2)
    Call SetLean(2, 4, "3534")
    Call SetLean(3, 5, "3516")
    Call SetLean(4, 6, "3529")
    Call SetLean(5, 7, "3059")
    Call SetLean(6, 8, "3131")
    Call 登録
    Call 初期化
    
    Call SetRace(52, 3)
    Call SetLean(2, 4, "3125")
    Call SetLean(3, 5, "3028")
    Call SetLean(4, 6, "2839")
    Call SetLean(5, 7, "2894")
    Call SetLean(6, 8, "3353")
    Call SetLean(7, 9, "3252")
    Call 登録
    Call 順位決定
    Call 初期化
    
    Call SetRace(53, 1)
    Call SetLean(2, 4, "3335")
    Call SetLean(3, 5, "3069")
    Call SetLean(4, 6, "2671")
    Call SetLean(5, 7, "3122")
    Call SetLean(6, 8, "4283")
    Call SetLean(7, 9, "")
    Call 登録
    Call 順位決定
    Call 初期化
    
    Call SetRace(54, 1)
    Call SetLean(1, 4, "4987")
    Call SetLean(2, 5, "4188")
    Call SetLean(3, 6, "4022")
    Call SetLean(4, 7, "3161")
    Call SetLean(5, 8, "3213")
    Call 登録
    Call 初期化
    
    Call SetRace(54, 2)
    Call SetLean(3, 5, "3781")
    Call SetLean(4, 6, "3135")
    Call SetLean(5, 7, "3161")
    Call SetLean(6, 8, "3003")
    Call 登録
    Call 順位決定
    Call 初期化
    
    Call SetRace(55, 1)
    Call SetLean(3, 5, "4333")
    Call SetLean(4, 6, "3778")
    Call SetLean(5, 7, "3785")
    Call SetLean(6, 8, "")
    Call 登録
    Call 初期化
    
    Call SetRace(55, 2)
    Call SetLean(2, 4, "3965")
    Call SetLean(3, 5, "3417")
    Call SetLean(4, 6, "3406")
    Call SetLean(5, 7, "3467")
    Call SetLean(6, 8, "")
    Call 登録
    Call 初期化
    
    Call SetRace(55, 3)
    Call SetLean(1, 3, "3453")
    Call SetLean(2, 4, "3204")
    Call SetLean(3, 5, "3231")
    Call SetLean(4, 6, "2990")
    Call SetLean(5, 7, "3098")
    Call SetLean(6, 8, "3209")
    Call SetLean(7, 9, "3451")
    Call 登録
    Call 順位決定
    Call 初期化
    
    Call SetRace(56, 1)
    Call SetLean(1, 3, "4131")
    Call SetLean(2, 4, "4473")
    Call SetLean(3, 5, "")
    Call SetLean(4, 6, "3869")
    Call SetLean(5, 7, "")
    Call SetLean(6, 8, "4477")
    Call SetLean(7, 9, "4311")
    Call 登録
    Call 初期化
    
    Call SetRace(56, 2)
    Call SetLean(1, 3, "3989")
    Call SetLean(2, 4, "3440")
    Call SetLean(3, 5, "3030")
    Call SetLean(4, 6, "2959")
    Call SetLean(5, 7, "3085")
    Call SetLean(6, 8, "3234")
    Call SetLean(7, 9, "3766")
    Call 登録
    Call 順位決定
    Call 初期化
    
    Call SetRace(57, 1)
    Call SetLean(2, 4, "5687")
    Call SetLean(3, 5, "5718")
    Call SetLean(4, 6, "4581")
    Call SetLean(5, 7, "4517")
    Call 登録
    Call 初期化
    
    Call SetRace(57, 2)
    Call SetLean(2, 4, "")
    Call SetLean(3, 5, "3852")
    Call SetLean(4, 6, "3403")
    Call SetLean(5, 7, "4070")
    Call SetLean(6, 8, "4047")
    Call SetLean(7, 9, "4037")
    Call 登録
    Call 順位決定
    Call 初期化
    
    Call SetRace(58, 1)
    Call SetLean(3, 5, "3758")
    Call SetLean(4, 6, "3425")
    Call SetLean(5, 7, "3809")
    Call SetLean(6, 8, "4717")
    Call 登録
    Call 初期化
    
    Call SetRace(58, 2)
    Call SetLean(2, 4, "")
    Call SetLean(3, 5, "3452")
    Call SetLean(4, 6, "3669")
    Call SetLean(5, 7, "")
    Call SetLean(6, 8, "3723")
    Call SetLean(7, 9, "")
    Call 登録
    Call 初期化
    
    Call SetRace(58, 3)
    Call SetLean(1, 3, "3228")
    Call SetLean(2, 4, "3313")
    Call SetLean(3, 5, "3303")
    Call SetLean(4, 6, "3172")
    Call SetLean(5, 7, "3021")
    Call SetLean(6, 8, "3131")
    Call SetLean(7, 9, "3427")
    Call 登録
    Call 初期化
    
    Call SetRace(58, 4)
    Call SetLean(1, 3, "2927")
    Call SetLean(2, 4, "")
    Call SetLean(3, 5, "3090")
    Call SetLean(4, 6, "2998")
    Call SetLean(5, 7, "3000")
    Call SetLean(6, 8, "3163")
    Call SetLean(7, 9, "3108")
    Call 登録
    Call 初期化
    
    Call SetRace(58, 5)
    Call SetLean(1, 3, "2858")
    Call SetLean(2, 4, "3031")
    Call SetLean(3, 5, "2766")
    Call SetLean(4, 6, "")
    Call SetLean(5, 7, "2817")
    Call SetLean(6, 8, "2814")
    Call SetLean(7, 9, "")
    Call 登録
    Call 順位決定
    Call 初期化
    
    Call SetRace(59, 1)
    Call SetLean(1, 3, "3926")
    Call SetLean(2, 4, "3583")
    Call SetLean(3, 5, "3846")
    Call SetLean(4, 6, "3686")
    Call SetLean(5, 7, "")
    Call SetLean(6, 8, "3457")
    Call SetLean(7, 9, "3486")
    Call 登録
    Call 初期化
    
    Call SetRace(59, 2)
    Call SetLean(1, 3, "3277")
    Call SetLean(2, 4, "3182")
    Call SetLean(3, 5, "3237")
    Call SetLean(4, 6, "3079")
    Call SetLean(5, 7, "")
    Call SetLean(6, 8, "3227")
    Call SetLean(7, 9, "3162")
    Call 登録
    Call 初期化
    
    Call SetRace(59, 3)
    Call SetLean(1, 3, "3123")
    Call SetLean(2, 4, "3137")
    Call SetLean(3, 5, "3118")
    Call SetLean(4, 6, "")
    Call SetLean(5, 7, "3055")
    Call SetLean(6, 8, "2919")
    Call SetLean(7, 9, "3002")
    Call 登録
    Call 初期化
    
    Call SetRace(59, 4)
    Call SetLean(1, 3, "3017")
    Call SetLean(2, 4, "")
    Call SetLean(3, 5, "2856")
    Call SetLean(4, 6, "2930")
    Call SetLean(5, 7, "2909")
    Call SetLean(6, 8, "2952")
    Call SetLean(7, 9, "3014")
    Call 登録
    Call 初期化
    
    Call SetRace(59, 5)
    Call SetLean(1, 3, "2821")
    Call SetLean(2, 4, "2881")
    Call SetLean(3, 5, "")
    Call SetLean(4, 6, "2747")
    Call SetLean(5, 7, "2740")
    Call SetLean(6, 8, "2841")
    Call SetLean(7, 9, "2805")
    Call 登録
    Call 順位決定
    Call 初期化
    
    Call SetRace(60, 1)
    Call SetLean(1, 3, "4423")
    Call SetLean(2, 4, "3869")
    Call SetLean(3, 5, "4349")
    Call SetLean(4, 6, "3492")
    Call SetLean(5, 7, "3539")
    Call SetLean(6, 8, "3966")
    Call SetLean(7, 9, "4361")
    Call 登録
    Call 初期化
    
    Call SetRace(60, 2)
    Call SetLean(3, 5, "3939")
    Call SetLean(4, 6, "3694")
    Call SetLean(5, 7, "3716")
    Call 登録
    Call 初期化
    
    Call SetRace(60, 3)
    Call SetLean(2, 4, "3702")
    Call SetLean(3, 5, "3296")
    Call SetLean(4, 6, "2922")
    Call SetLean(5, 7, "3034")
    Call SetLean(6, 8, "3054")
    Call SetLean(7, 9, "2934")
    Call 登録
    Call 初期化
    
    Call SetRace(60, 4)
    Call SetLean(2, 4, "")
    Call SetLean(3, 5, "2795")
    Call SetLean(4, 6, "2594")
    Call SetLean(5, 7, "2808")
    Call SetLean(6, 8, "3382")
    Call SetLean(7, 9, "3759")
    Call 登録
    Call 初期化
    
    Call SetRace(60, 5)
    Call SetLean(2, 4, "3368")
    Call SetLean(3, 5, "3005")
    Call SetLean(4, 6, "2580")
    Call SetLean(5, 7, "2846")
    Call SetLean(6, 8, "3339")
    Call 登録
    Call 順位決定
    Call 初期化
    
    Call SetRace(61, 1)
    Call SetLean(2, 4, "4411")
    Call SetLean(3, 5, "4965")
    Call SetLean(4, 6, "4694")
    Call SetLean(5, 7, "4896")
    Call SetLean(6, 8, "5187")
    Call 登録
    Call 初期化
    
    Call SetRace(61, 2)
    Call SetLean(2, 4, "4574")
    Call SetLean(3, 5, "4710")
    Call SetLean(4, 6, "3936")
    Call SetLean(5, 7, "3837")
    Call SetLean(6, 8, "4242")
    Call 登録
    Call 順位決定
    Call 初期化
    
    Call SetRace(62, 1)
    Call SetLean(2, 4, "5251")
    Call SetLean(3, 5, "4368")
    Call SetLean(4, 6, "4025")
    Call SetLean(5, 7, "3905")
    Call SetLean(6, 8, "4276")
    Call 登録
    Call 順位決定
    Call 初期化
    
    Call SetRace(63, 1)
    Call SetLean(3, 5, "5149")
    Call SetLean(4, 6, "4705")
    Call SetLean(5, 7, "4370")
    Call SetLean(6, 8, "4291")
    Call 登録
    Call 順位決定
    Call 初期化
    
    Call SetRace(64, 1)
    Call SetLean(2, 4, "5332")
    Call SetLean(3, 5, "5407")
    Call SetLean(4, 6, "4121")
    Call SetLean(5, 7, "4121")
    Call SetLean(6, 8, "5233")
    Call 登録
    Call 初期化
    
    Call SetRace(64, 2)
    Call SetLean(1, 3, "4217")
    Call SetLean(2, 4, "3916")
    Call SetLean(3, 5, "3817")
    Call SetLean(4, 6, "3367")
    Call SetLean(5, 7, "3684")
    Call SetLean(6, 8, "3932")
    Call SetLean(7, 9, "")
    Call 登録
    Call 順位決定
    Call 初期化
    
    Call SetRace(65, 1)
    Call SetLean(3, 5, "4148")
    Call SetLean(4, 6, "5294")
    Call SetLean(5, 7, "")
    Call SetLean(6, 8, "12304")
    Call 登録
    Call 初期化
    
    Call SetRace(65, 2)
    Call SetLean(2, 4, "4222")
    Call SetLean(3, 5, "4239")
    Call SetLean(4, 6, "4525")
    Call SetLean(5, 7, "4232")
    Call SetLean(6, 8, "")
    Call 登録
    Call 初期化
    
    Call SetRace(65, 3)
    Call SetLean(1, 3, "3958")
    Call SetLean(2, 4, "")
    Call SetLean(3, 5, "3721")
    Call SetLean(4, 6, "3462")
    Call SetLean(5, 7, "3687")
    Call SetLean(6, 8, "3787")
    Call SetLean(7, 9, "")
    Call 登録
    Call 順位決定
    Call 初期化
    
    Call SetRace(66, 1)
    Call SetLean(3, 5, "4714")
    Call SetLean(4, 6, "4801")
    Call SetLean(5, 7, "5058")
    Call SetLean(6, 8, "10681")
    Call 登録
    Call 初期化
    
    Call SetRace(66, 2)
    Call SetLean(2, 4, "4454")
    Call SetLean(3, 5, "4010")
    Call SetLean(4, 6, "3580")
    Call SetLean(5, 7, "3501")
    Call SetLean(6, 8, "4772")
    Call 登録
    Call 順位決定
    Call 初期化
    
    Call SetRace(67, 1)
    Call SetLean(4, 6, "31220")
    Call SetLean(5, 7, "23717")
    Call 登録
    Call 順位決定
    Call 初期化
    
    Call SetRace(68, 1)
    Call SetLean(2, 4, "")
    Call SetLean(3, 5, "22941")
    Call SetLean(4, 6, "20615")
    Call SetLean(5, 7, "22545")
    Call SetLean(6, 8, "20436")
    Call SetLean(7, 9, "21796")
    Call 登録
    Call 順位決定
    Call 初期化
    
    Call SetRace(69, 1)
    Call SetLean(3, 5, "22397")
    Call SetLean(4, 6, "21915")
    Call SetLean(5, 7, "23395")
    Call 登録
    Call 順位決定
    Call 初期化
    
    Call SetRace(70, 1)
    Call SetLean(4, 6, "15782")
    Call SetLean(5, 7, "23187")
    Call 登録
    Call 順位決定
    Call 初期化
    
    Call SetRace(71, 1)
    Call SetLean(2, 4, "20869")
    Call SetLean(3, 5, "21902")
    Call SetLean(4, 6, "15415")
    Call SetLean(5, 7, "20069")
    Call SetLean(6, 8, "20970")
    Call SetLean(7, 9, "20218")
    Call 登録
    Call 順位決定
    Call 初期化
    
    Call SetRace(72, 1)
    Call SetLean(1, 3, "20111")
    Call SetLean(2, 4, "20657")
    Call SetLean(3, 5, "15166")
    Call SetLean(4, 6, "15471")
    Call SetLean(5, 7, "15622")
    Call SetLean(6, 8, "15981")
    Call SetLean(7, 9, "21366")
    Call 登録
    Call 順位決定
    Call 初期化
    
    ' ブックの保存
    ActiveWorkbook.Save
End Sub
