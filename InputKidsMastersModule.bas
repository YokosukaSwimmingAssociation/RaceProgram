Attribute VB_Name = "KidsMastersCorrectionModule"
'
' 学童マスターズのプログラム補正
'
Sub 学マプログラム補正()

    Call EventChange(False)

    Sheets("エントリー一覧").Activate
    ' 1-2-4 同一レース
    Call ModCell("F2", "3")
    Call ModCell("G2", "長谷部弘子・新原喜伊子・大串玲・桝田衣代")
    Call ModCell("F3", "4")
    Call ModCell("G3", "長和佐・鶴岡美佐・佐野華子・木田美津子")
    Call ModCell("C4", "10")
    Call ModCell("C5", "10")
    Call ModCell("F5", "8")
    Call ModCell("G5", "津田宏志朗・平山新・小林大虎・碇光太郎")
    Call ModCell("G6", "清島邦子・鈴木博・新原光男・栗原恵美")
    Call ModCell("G7", "高橋和枝・竹村昌子・丸茂健二・名倉健吾")
    Call ModCell("G8", "鈴木治輝・大橋恵子・岩崎優美・森川昌三")
    Call ModCell("G9", "岩根央学・市丸恵子・兼森伸児・野村洋子")
    ' 12-13 同一レース
    Call ModCell(SearchCell(12, "鈴木　とみ子", "レーン"), "4") ' "F33"
    Call ModCell(SearchCell(12, "川上　京子", "レーン"), "5") ' "F34"
    Call ModCell(SearchCell(13, "尾谷　庸三", "レースNo"), "110") ' "C35"
    Call ModCell(SearchCell(13, "尾谷　庸三", "レーン"), "7") ' "F35"

    ' 19 男子100M自由形
    Call ModCell(SearchCell(19, "中島　輝夫", "レーン"), "4") ' "F47"
    Call ModCell(SearchCell(19, "冨田　清雄", "レーン"), "6") '"F48"

    ' 26-29 同一レース
    Call ModCell(SearchCell(26, "鈴木　さくら", "レーン"), "3") ' "F78"
    Call ModCell(SearchCell(29, "緒形　輝海洋", "レースNo"), "240") ' "C79"
    Call ModCell(SearchCell(29, "緒形　輝海洋", "レーン"), "5")  ' "F79"
    Call ModCell(SearchCell(29, "グラハム　ニコラ", "レースNo"), "240") '"C80"
    Call ModCell(SearchCell(29, "グラハム　ニコラ", "レーン"), "6") '"F80"
    Call ModCell(SearchCell(29, "杉山　瑚太朗", "レースNo"), "240") '"C81"
    Call ModCell(SearchCell(29, "杉山　瑚太朗", "レーン"), "7") '"F81"
    Call ModCell(SearchCell(29, "蛭田　成海", "レースNo"), "240") '"C82"
    Call ModCell(SearchCell(29, "蛭田　成海", "レーン"), "8") '"F82"
    
    ' 33-34 同一レース
    Call ModCell(SearchCell(33, "堀　　圭佑", "レーン"), "4")  '"F93"
    Call ModCell(SearchCell(34, "根本　珠奈", "レースNo"), "280") '"C94"
    Call ModCell(SearchCell(34, "根本　珠奈", "レーン"), "7")  '"F94"
    Call ModCell(SearchCell(34, "石井　春花", "レースNo"), "280") '"C95"
    Call ModCell(SearchCell(34, "石井　春花", "レーン"), "6")  '"F95"
    
    ' 36 混合メドレーリレー
    Call ModCell("G99", "望月真澄・大串玲・武田深雪・国広真司")
    
    ' 37 女子50M背泳ぎ
    Call ModCell(SearchCell(37, "岩崎　摩耶子", "レーン"), "3")  ' "F100"
    Call ModCell(SearchCell(37, "三上　美佐枝", "レーン"), "4")  ' "F101"
    Call ModCell(SearchCell(37, "柴　　恵子", "レーン"), "5")  ' "F102"
    Call ModCell(SearchCell(37, "長谷部　弘子", "レースNo"), "320") ' "C103"
    Call ModCell(SearchCell(37, "長谷部　弘子", "組"), "1")  ' "E103"
    Call ModCell(SearchCell(37, "長谷部　弘子", "レーン"), "8")  ' "F103"
    Call ModCell(SearchCell(37, "高橋　和枝", "レースNo"), "320") ' "C104"
    Call ModCell(SearchCell(37, "高橋　和枝", "組"), "1")  ' "E104"
    Call ModCell(SearchCell(37, "高橋　和枝", "レーン"), "6")  ' "F104"
    Call ModCell(SearchCell(37, "座間　千鶴子", "レースNo"), "320") ' "C105"
    Call ModCell(SearchCell(37, "座間　千鶴子", "組"), "1")  ' "E105"
    Call ModCell(SearchCell(37, "座間　千鶴子", "レーン"), "7")  ' "F105"
    Call ModCell(SearchCell(37, "長　　和佐", "レーン"), "4")  ' "F106"
    Call ModCell(SearchCell(37, "鈴木　愛子", "レーン"), "5")  ' "F107"
    Call ModCell(SearchCell(37, "衛藤　由里子", "レーン"), "6")  ' "F108"
    Call ModCell(SearchCell(37, "田浦　一恵", "レーン"), "7")  ' "F109"
    
    ' 40 小学１・２年男子50M背泳ぎ
    Call ModCell(SearchCell(40, "末吉　貫汰", "レーン"), "4")  ' "F124"
    Call ModCell(SearchCell(40, "飯岡　旺士郎", "レーン"), "7")  ' "F125"
    Call ModCell(SearchCell(40, "吉井　陸人", "レーン"), "8")  ' "F126"
    Call ModCell(SearchCell(40, "速水　　陽", "レースNo"), "370") ' "C127"
    Call ModCell(SearchCell(40, "速水　　陽", "組"), "1")  ' "E127"
    Call ModCell(SearchCell(40, "速水　　陽", "レーン"), "6")  ' "F127"
    Call ModCell(SearchCell(40, "工藤　映空", "レースNo"), "370") ' "C133"
    Call ModCell(SearchCell(40, "工藤　映空", "組"), "1")  ' "E133"
    Call ModCell(SearchCell(40, "工藤　映空", "レーン"), "5")  ' "F133"
    
    ' 41 女子50M背泳ぎ
    Call ModCell(SearchCell(41, "鈴木　萌依彩", "レーン"), "6")
    Call ModCell(SearchCell(41, "高田　楓佳", "レースNo"), "390") ' "C137"
    Call ModCell(SearchCell(41, "高田　楓佳", "組"), "1")  ' "E137"
    Call ModCell(SearchCell(41, "高田　楓佳", "レーン"), "5")  ' "F137"
    
    ' 43 女子50M背泳ぎ
    Call ModCell(SearchCell(43, "岩澤　みずき", "レーン"), "8")  ' "F152"
    Call ModCell(SearchCell(43, "小澤　琉楓", "レーン"), "7")  ' "F153"
    Call ModCell(SearchCell(43, "鈴木　緋彩", "レーン"), "5")  ' "F154"
    Call ModCell(SearchCell(43, "栃原　一菜", "レーン"), "4")  ' "F155"
    Call ModCell(SearchCell(43, "桝永　優郁", "レーン"), "3")  ' "F156"
    Call ModCell(SearchCell(43, "高橋　凛花", "レースNo"), "430") ' "C163"
    Call ModCell(SearchCell(43, "高橋　凛花", "組"), "1")  ' "E163"
    Call ModCell(SearchCell(43, "高橋　凛花", "レーン"), "6")  ' "F163"
    
    ' 44 男子50M背泳ぎ
    Call ModCell(SearchCell(44, "林　　大真", "レーン"), "4")
    Call ModCell(SearchCell(44, "鈴木　輪友", "レーン"), "5")
    Call ModCell(SearchCell(44, "土佐　優大", "レースNo"), "450")
    Call ModCell(SearchCell(44, "土佐　優大", "組"), "1")
    Call ModCell(SearchCell(44, "土佐　優大", "レーン"), "6")
    Call ModCell(SearchCell(44, "高橋　波瑠", "レーン"), "7")
    Call ModCell(SearchCell(44, "林　　大智", "レーン"), "8")
    
    ' 45 女子50M自由形
    Call ModCell(SearchCell(45, "山本　栄子", "レーン"), "3")  ' "F175"
    Call ModCell(SearchCell(45, "瀬谷　とし子", "レーン"), "5")  ' "F176"
    Call ModCell(SearchCell(45, "清島　邦子", "レーン"), "4")  ' "F177"
    Call ModCell(SearchCell(45, "新原　喜伊子", "レースNo"), "470") ' "C178"
    Call ModCell(SearchCell(45, "新原　喜伊子", "組"), "1")  ' "E178"
    Call ModCell(SearchCell(45, "新原　喜伊子", "レーン"), "8")  ' "F178"
    Call ModCell(SearchCell(45, "栗原　恵美", "レースNo"), "470") ' "C179"
    Call ModCell(SearchCell(45, "栗原　恵美", "組"), "1")  ' "E179"
    Call ModCell(SearchCell(45, "栗原　恵美", "レーン"), "6")  ' "F179"
    Call ModCell(SearchCell(45, "高梨　加代子", "レースNo"), "470") ' "C180"
    Call ModCell(SearchCell(45, "高梨　加代子", "組"), "1")  ' "E180"
    Call ModCell(SearchCell(45, "高梨　加代子", "レーン"), "7")  ' "F180"
    Call ModCell(SearchCell(45, "薮崎　峰子", "レーン"), "4")  ' "F181"
    Call ModCell(SearchCell(45, "山口　孝子", "レーン"), "5")  ' "F182"
    Call ModCell(SearchCell(45, "古澤　さとみ", "レーン"), "6")  ' "F183"
    Call ModCell(SearchCell(45, "長谷部　弘子", "レーン"), "8")  ' "F184"
    Call ModCell(SearchCell(45, "武田　深雪", "レースNo"), "480") ' "C185"
    Call ModCell(SearchCell(45, "武田　深雪", "組"), "2")  ' "E185"
    Call ModCell(SearchCell(45, "武田　深雪", "レーン"), "7")  ' "F185"
    Call ModCell(SearchCell(45, "大串　　玲", "レーン"), "3")  ' "F186"
    Call ModCell(SearchCell(45, "鶴岡　美佐", "レーン"), "4")  ' "F187"
    Call ModCell(SearchCell(45, "長　　和佐", "レーン"), "5")  ' "F188"
    Call ModCell(SearchCell(45, "木田　美津子", "レーン"), "6")  ' "F189"
    Call ModCell(SearchCell(45, "佐野　華子", "レーン"), "7")  ' "F190"
    Call ModCell(SearchCell(45, "齊藤　香奈", "レーン"), "8")  ' "F191"
    
    ' 46 男子50M自由形
    Call ModCell(SearchCell(46, "小杉　邦洋", "レーン"), "5")  ' "F195"
    Call ModCell(SearchCell(46, "高間　秀泰", "レーン"), "6")  ' "F194"
    Call ModCell(SearchCell(46, "中田　光彦", "レーン"), "7")
    Call ModCell(SearchCell(46, "西川　博保", "レーン"), "8")
    Call ModCell(SearchCell(46, "山本　眞人", "レーン"), "9")
    Call ModCell(SearchCell(46, "岩根　央学", "レーン"), "3")
    Call ModCell(SearchCell(46, "岩井　太造", "レーン"), "5")
    
    ' 49 女子50M自由形
    Call ModCell(SearchCell(49, "永野　瑞季", "レーン"), "4")  ' "F240"
    Call ModCell(SearchCell(49, "齋藤　百々奈", "レーン"), "7")  ' "F241"
    Call ModCell(SearchCell(49, "小和田　陽彩", "レーン"), "8")  ' "F242"
    Call ModCell(SearchCell(49, "野間　愛莉", "レースNo"), "570") ' "C243"
    Call ModCell(SearchCell(49, "野間　愛莉", "組"), "1")  ' "E243"
    Call ModCell(SearchCell(49, "野間　愛莉", "レーン"), "6")  ' "F243"
    Call ModCell(SearchCell(49, "砂川　絢音", "レーン"), "8")  ' "F244"
    Call ModCell(SearchCell(49, "高田　楓佳", "レーン"), "7")  ' "F245"
    Call ModCell(SearchCell(49, "斎藤　由奈", "レーン"), "5")  ' "F246"
    Call ModCell(SearchCell(49, "齋藤　詩恵里", "レーン"), "4")  ' "F247"
    Call ModCell(SearchCell(49, "中村　亜優音", "レーン"), "3")  ' "F248"
    Call ModCell(SearchCell(49, "鈴木　萌依彩", "レースNo"), "570") ' "C249"
    Call ModCell(SearchCell(49, "鈴木　萌依彩", "組"), "1")  ' "E249"
    Call ModCell(SearchCell(49, "鈴木　萌依彩", "レーン"), "5")  ' "F249"
    Call ModCell(SearchCell(49, "笠原　成珠", "レースNo"), "580") ' "C256"
    Call ModCell(SearchCell(49, "笠原　成珠", "組"), "2")  ' "E256"
    Call ModCell(SearchCell(49, "笠原　成珠", "レーン"), "6")  ' "F256"
    ' 50 男子50M自由形
    Call ModCell(SearchCell(50, "小澤　鳳乙", "レーン"), "3")  ' "F257"
    Call ModCell(SearchCell(50, "鈴木　頼主", "レーン"), "4")  ' "F258"
    Call ModCell(SearchCell(50, "三冨　啓幸", "レーン"), "7")  ' "F259"
    Call ModCell(SearchCell(50, "岩澤　健太", "レーン"), "8")  ' "F260"
    Call ModCell(SearchCell(50, "水野　　庄", "レーン"), "3")  ' "F261"
    Call ModCell(SearchCell(50, "石川　遼大", "レーン"), "8")  ' "F262"
    Call ModCell(SearchCell(50, "高橋　雄大", "レーン"), "7")  ' "F263"
    Call ModCell(SearchCell(50, "松澤　優樹", "レーン"), "5")  ' "F264"
    Call ModCell(SearchCell(50, "蛭田　誠司", "レーン"), "4")  ' "F265"
    Call ModCell(SearchCell(50, "伊東　　楓", "レースNo"), "600") ' "C266"
    Call ModCell(SearchCell(50, "伊東　　楓", "組"), "1")  ' "E266"
    Call ModCell(SearchCell(50, "伊東　　楓", "レーン"), "6")  ' "F266"
    Call ModCell(SearchCell(50, "岩見　陽太朗", "レースNo"), "600") ' "C267"
    Call ModCell(SearchCell(50, "岩見　陽太朗", "組"), "1")  ' "E267"
    Call ModCell(SearchCell(50, "岩見　陽太朗", "レーン"), "5")  ' "F267"
    Call ModCell(SearchCell(50, "碇　　光太郎", "レースNo"), "610") ' "C274"
    Call ModCell(SearchCell(50, "碇　　光太郎", "組"), "1")  ' "E274"
    Call ModCell(SearchCell(50, "碇　　光太郎", "レーン"), "6")  ' "F274"
    
    ' 51 女子50M自由形
    Call ModCell(SearchCell(51, "岩澤　みずき", "レーン"), "4")  ' "F275"
    Call ModCell(SearchCell(51, "齋藤　百合音", "レーン"), "7")  ' "F276"
    Call ModCell(SearchCell(51, "桝永　優郁", "レーン"), "8")  ' "F277"
    Call ModCell(SearchCell(51, "神尾　美樹", "レースNo"), "630") ' "C278"
    Call ModCell(SearchCell(51, "神尾　美樹", "組"), "1")  ' "E278"
    Call ModCell(SearchCell(51, "神尾　美樹", "レーン"), "6")  ' "F278"
    Call ModCell(SearchCell(51, "栃原　一菜", "レーン"), "4")  ' "F279"
    Call ModCell(SearchCell(51, "小澤　琉楓", "レーン"), "7")  ' "F280"
    Call ModCell(SearchCell(51, "鈴木　緋彩", "レーン"), "8")  ' "F281"
    Call ModCell(SearchCell(51, "北田　侑那", "レースNo"), "630") ' "C282"
    Call ModCell(SearchCell(51, "北田　侑那", "組"), "1")  ' "E282"
    Call ModCell(SearchCell(51, "北田　侑那", "レーン"), "5")  ' "F282"
    Call ModCell(SearchCell(51, "南　　恵仁和", "レースNo"), "640") ' "C283"
    Call ModCell(SearchCell(51, "南　　恵仁和", "組"), "2")  ' "E283"
    Call ModCell(SearchCell(51, "南　　恵仁和", "レーン"), "6")  ' "F283"
    Call ModCell(SearchCell(51, "鳥井　果乃", "レースNo"), "640") ' "C289"
    Call ModCell(SearchCell(51, "鳥井　果乃", "組"), "2")  ' "E289"
    Call ModCell(SearchCell(51, "鳥井　果乃", "レーン"), "5")  ' "F289"
    
    ' 52 男子50M自由形
    Call ModCell(SearchCell(52, "藤本　翔太", "レーン"), "6")  ' "F292"
    Call ModCell(SearchCell(52, "久保田　快", "レースNo"), "670") ' "C293"
    Call ModCell(SearchCell(52, "久保田　快", "組"), "2")  ' "E293"
    Call ModCell(SearchCell(52, "久保田　快", "レーン"), "9")  ' "F293"
    Call ModCell(SearchCell(52, "高橋　波瑠", "レーン"), "5")  ' "F294"
    Call ModCell(SearchCell(52, "高橋　　煌", "レースNo"), "660") ' "C296"
    Call ModCell(SearchCell(52, "高橋　　煌", "組"), "1")  ' "E296"
    Call ModCell(SearchCell(52, "高橋　　煌", "レーン"), "7")  ' "F296"
    Call ModCell(SearchCell(52, "小泉　柊介", "レーン"), "3")  ' "F302"
    
    ' 53 女子50Mバタフライ
    Call ModCell(SearchCell(53, "桝田　衣代", "レーン"), "7")  ' "F311"
    Call ModCell(SearchCell(53, "菅谷　幸江", "レーン"), "6")  ' "F312"
    
    ' 54 男子50Mバタフライ
    Call ModCell(SearchCell(54, "越川　唯幸", "レースNo"), "700") ' "C296"
    Call ModCell(SearchCell(54, "越川　唯幸", "組"), "1")  ' "E296"
    Call ModCell(SearchCell(54, "越川　唯幸", "レーン"), "8")  ' "F296"
    
    
    ' 61 女子50M平泳ぎ
    Call ModCell(SearchCell(61, "古澤　さとみ", "レーン"), "8")  ' "F344"
    Call ModCell(SearchCell(61, "竹村　昌子", "レースNo"), "780") ' "C345"
    Call ModCell(SearchCell(61, "竹村　昌子", "組"), "1")  ' "E345"
    Call ModCell(SearchCell(61, "竹村　昌子", "レーン"), "6")  ' "F345"
    Call ModCell(SearchCell(61, "市丸　恵子", "レースNo"), "780") ' "C346"
    Call ModCell(SearchCell(61, "市丸　恵子", "組"), "1")  ' "E346"
    Call ModCell(SearchCell(61, "市丸　恵子", "レーン"), "7")  ' "F346"
    Call ModCell(SearchCell(61, "鶴岡　美佐", "レーン"), "4")  ' "F347"
    Call ModCell(SearchCell(61, "柳澤　亜矢子", "レーン"), "5")  ' "F349"
    Call ModCell(SearchCell(61, "木田　美津子", "レーン"), "6")  ' "F348"
    Call ModCell(SearchCell(61, "大村　怜子", "レーン"), "7")  ' "F350"
    ' 70-71 同一レース
    Call ModCell("G386", "山口孝子・山本栄子・三上美佐枝・柴恵子")
    Call ModCell("G387", "野村洋子・市丸恵子・高梨加代子・座間千鶴子")
    Call ModCell("G388", "武石淳子・柳澤亜矢子・川上京子・大村怜子")
    Call ModCell("G389", "佐野華子・長和佐・木田美津子・鶴岡美佐")
    Call ModCell("G390", "西昇・寺島圭一・西川博保・山田博康")
    Call ModCell("G391", "安田晙一・中田光彦・青木貢・青木悟史")
    Call ModCell("G392", "小林慶雄・杉山弘・南清志・池田薫")
    Call ModCell("F393", "3")
    Call ModCell("G393", "斉藤百合音・鈴木萌依彩・高田楓佳・砂川絢音")
    Call ModCell("F394", "4")
    Call ModCell("G394", "小澤琉楓・鈴木緋彩・桝永優郁・栃原一菜")
    Call ModCell("C395", "900")
    Call ModCell("F395", "6")
    Call ModCell("G395", "鈴木輪友・水野庄・伊東楓・三冨啓幸")
    Call ModCell("C396", "900")
    Call ModCell("F396", "7")
    Call ModCell("G396", "小林大虎・津田宏志朗・碇光太郎・平山新")
    Call ModCell("C397", "900")
    Call ModCell("F397", "8")
    Call ModCell("G397", "藤本翔太・小泉柊介・高橋波瑠・高橋煌")

    Call EventChange(True)
    
    ' ブックの保存
    ActiveWorkbook.Save
End Sub

'
' 令和元年度のプログラムの記録を入れる
'
Sub 学マ記録入力()
    Sheets("記録画面").Select
    Call SetRace(1, 1)
    Call SetLean(1, 3, "34709")
    Call SetLean(2, 4, "25977")
    Call SetLean(3, 6, "")
    Call SetLean(4, 8, "30973")
    Call 登録
    Call 順位決定
    Call 初期化

    Call SetRace(5, 1)
    Call SetLean(1, 4, "32844")
    Call SetLean(2, 5, "30994")
    Call SetLean(3, 6, "24303")
    Call SetLean(4, 7, "25060")
    Call 登録
    Call 順位決定
    Call 初期化

    Call SetRace(6, 1)
    Call SetLean(1, 5, "35775")
    Call SetLean(2, 6, "34009")
    Call SetLean(3, 7, "30945")
    Call 登録
    Call 順位決定
    Call 初期化

    Call SetRace(7, 1)
    Call SetLean(1, 5, "40596")
    Call SetLean(2, 6, "34035")
    Call SetLean(3, 7, "25870")
    Call 登録
    Call 順位決定
    Call 初期化

    Call SetRace(8, 1)
    Call SetLean(1, 3, "35996")
    Call SetLean(2, 4, "33734")
    Call SetLean(3, 5, "30026")
    Call SetLean(4, 6, "32161")
    Call SetLean(5, 7, "32447")
    Call SetLean(6, 8, "40131")
    Call SetLean(7, 9, "44509")
    Call 登録
    Call 順位決定
    Call 初期化

    Call SetRace(9, 1)
    Call SetLean(2, 4, "33965")
    Call SetLean(3, 5, "30134")
    Call SetLean(4, 6, "25908")
    Call SetLean(5, 7, "33560")
    Call 登録
    Call 順位決定
    Call 初期化

    Call SetRace(10, 1)
    Call SetLean(3, 5, "24066")
    Call SetLean(4, 6, "24885")
    Call SetLean(5, 7, "31742")
    Call 登録
    Call 順位決定
    Call 初期化

    Call SetRace(11, 1)
    Call SetLean(3, 5, "30128")
    Call SetLean(4, 6, "24081")
    Call SetLean(5, 7, "32400")
    Call 登録
    Call 順位決定
    Call 初期化

    Call SetRace(12, 1)
    Call SetLean(2, 4, "22203")
    Call SetLean(3, 5, "14942")
    Call SetLean(5, 7, "13107")
    Call 登録
    Call 順位決定
    Call 初期化

    Call SetRace(15, 1)
    Call SetLean(3, 5, "15144")
    Call SetLean(4, 6, "14670")
    Call SetLean(5, 7, "21162")
    Call 登録
    Call 順位決定
    Call 初期化

    Call SetRace(16, 1)
    Call SetLean(3, 5, "11638")
    Call SetLean(4, 6, "11011")
    Call SetLean(5, 7, "13879")
    Call 登録
    Call 順位決定
    Call 初期化

    Call SetRace(18, 1)
    Call SetLean(2, 4, "20698")
    Call SetLean(3, 5, "15065")
    Call SetLean(4, 6, "12312")
    Call SetLean(5, 7, "14285")
    Call SetLean(6, 8, "12609")
    Call 登録
    Call 順位決定
    Call 初期化

    Call SetRace(19, 1)
    Call SetLean(2, 4, "12738")
    Call SetLean(3, 5, "12811")
    Call SetLean(4, 6, "14208")
    Call SetLean(5, 7, "14290")
    Call 登録
    Call 初期化

    Call SetRace(19, 2)
    Call SetLean(2, 4, "12566")
    Call SetLean(3, 5, "11229")
    Call SetLean(4, 6, "11493")
    Call SetLean(5, 7, "12115")
    Call 登録
    Call 順位決定
    Call 初期化

    Call SetRace(20, 1)
    Call SetLean(2, 4, "14773")
    Call SetLean(3, 5, "")
    Call SetLean(4, 6, "11633")
    Call SetLean(5, 7, "13965")
    Call SetLean(6, 8, "")
    Call 登録
    Call 順位決定
    Call 初期化

    Call SetRace(21, 1)
    Call SetLean(3, 5, "13404")
    Call SetLean(4, 6, "12150")
    Call SetLean(5, 7, "")
    Call 登録
    Call 順位決定
    Call 初期化

    Call SetRace(22, 1)
    Call SetLean(2, 4, "11381")
    Call SetLean(3, 5, "11281")
    Call SetLean(4, 6, "")
    Call SetLean(5, 7, "11346")
    Call SetLean(6, 8, "13004")
    Call 登録
    Call 順位決定
    Call 初期化

    Call SetRace(23, 1)
    Call SetLean(1, 3, "20131")
    Call SetLean(2, 4, "11678")
    Call SetLean(3, 5, "11109")
    Call SetLean(4, 6, "10461")
    Call SetLean(5, 7, "11666")
    Call SetLean(6, 8, "12853")
    Call 登録
    Call 順位決定
    Call 初期化

    Call SetRace(24, 1)
    Call SetLean(3, 5, "20436")
    Call SetLean(4, 6, "14259")
    Call 登録
    Call 順位決定
    Call 初期化

    Call SetRace(25, 1)
    Call SetLean(3, 5, "21903")
    Call SetLean(4, 6, "11985")
    Call 登録
    Call 順位決定
    Call 初期化

    Call SetRace(26, 1)
    Call SetLean(1, 3, "12098")
    Call SetLean(3, 5, "12302")
    Call SetLean(4, 6, "11185")
    Call SetLean(5, 7, "11273")
    Call SetLean(6, 8, "12923")
    Call 登録
    Call 順位決定
    Call 初期化

    Call SetRace(30, 1)
    Call SetLean(2, 4, "14856")
    Call SetLean(3, 5, "13937")
    Call SetLean(4, 6, "14220")
    Call SetLean(5, 7, "13409")
    Call 登録
    Call 順位決定
    Call 初期化

    Call SetRace(31, 1)
    Call SetLean(1, 3, "21711")
    Call SetLean(2, 4, "15887")
    Call SetLean(3, 5, "15345")
    Call SetLean(4, 6, "15235")
    Call SetLean(5, 7, "14402")
    Call SetLean(6, 8, "13034")
    Call 登録
    Call 順位決定
    Call 初期化

    Call SetRace(33, 1)
    Call SetLean(2, 4, "14304")
    Call SetLean(4, 6, "13373")
    Call SetLean(5, 7, "13388")
    Call 登録
    Call 順位決定
    Call 初期化

    Call SetRace(35, 1)
    Call SetLean(3, 5, "13812")
    Call SetLean(4, 6, "13189")
    Call SetLean(5, 7, "14818")
    Call 登録
    Call 順位決定
    Call 初期化

    Call SetRace(36, 1)
    Call SetLean(4, 6, "22462")
    Call 登録
    Call 順位決定
    Call 初期化

    Call SetRace(37, 1)
    Call SetLean(1, 3, "11180")
    Call SetLean(2, 4, "5993")
    Call SetLean(3, 5, "5695")
    Call SetLean(4, 6, "5757")
    Call SetLean(5, 7, "4797")
    Call SetLean(6, 8, "5548")
    Call 登録
    Call 初期化

    Call SetRace(37, 2)
    Call SetLean(2, 4, "4905")
    Call SetLean(3, 5, "4497")
    Call SetLean(4, 6, "4634")
    Call SetLean(5, 7, "4034")
    Call 登録
    Call 順位決定
    Call 初期化

    Call SetRace(38, 1)
    Call SetLean(2, 4, "10412")
    Call SetLean(3, 5, "4200")
    Call SetLean(4, 6, "4949")
    Call SetLean(5, 7, "11151")
    Call 登録
    Call 初期化

    Call SetRace(38, 2)
    Call SetLean(2, 4, "4274")
    Call SetLean(3, 5, "")
    Call SetLean(4, 6, "3067")
    Call SetLean(5, 7, "2844")
    Call 登録
    Call 順位決定
    Call 初期化

    Call SetRace(39, 1)
    Call SetLean(1, 3, "10692")
    Call SetLean(2, 4, "5841")
    Call SetLean(3, 5, "5253")
    Call SetLean(4, 6, "")
    Call SetLean(5, 7, "5480")
    Call SetLean(6, 8, "5486")
    Call 登録
    Call 順位決定
    Call 初期化

    Call SetRace(40, 1)
    Call SetLean(2, 4, "5817")
    Call SetLean(3, 5, "10205")
    Call SetLean(4, 6, "10429")
    Call SetLean(5, 7, "10879")
    Call SetLean(6, 8, "10224")
    Call 登録
    Call 初期化

    Call SetRace(40, 2)
    Call SetLean(2, 4, "5227")
    Call SetLean(3, 5, "")
    Call SetLean(4, 6, "4270")
    Call SetLean(5, 7, "10319")
    Call SetLean(6, 8, "5011")
    Call 登録
    Call 順位決定
    Call 初期化

    Call SetRace(41, 1)
    Call SetLean(2, 4, "11365")
    Call SetLean(3, 5, "10335")
    Call SetLean(4, 6, "10911")
    Call SetLean(5, 7, "10980")
    Call 登録
    Call 初期化

    Call SetRace(41, 2)
    Call SetLean(2, 4, "5318")
    Call SetLean(3, 5, "4952")
    Call SetLean(4, 6, "4242")
    Call SetLean(5, 7, "4640")
    Call SetLean(6, 8, "5242")
    Call 登録
    Call 順位決定
    Call 初期化

    Call SetRace(42, 1)
    Call SetLean(2, 4, "11682")
    Call SetLean(3, 5, "10808")
    Call SetLean(4, 6, "")
    Call SetLean(5, 7, "11071")
    Call 登録
    Call 初期化

    Call SetRace(42, 2)
    Call SetLean(2, 4, "10134")
    Call SetLean(3, 5, "4025")
    Call SetLean(4, 6, "3967")
    Call SetLean(5, 7, "4625")
    Call SetLean(6, 8, "5993")
    Call 登録
    Call 順位決定
    Call 初期化

    Call SetRace(43, 1)
    Call SetLean(1, 3, "12698")
    Call SetLean(2, 4, "5893")
    Call SetLean(3, 5, "5669")
    Call SetLean(4, 6, "5149")
    Call SetLean(5, 7, "5519")
    Call SetLean(6, 8, "11451")
    Call 登録
    Call 初期化

    Call SetRace(43, 2)
    Call SetLean(1, 3, "5554")
    Call SetLean(2, 4, "4237")
    Call SetLean(3, 5, "3527")
    Call SetLean(4, 6, "3191")
    Call SetLean(5, 7, "")
    Call SetLean(6, 8, "4485")
    Call 登録
    Call 順位決定
    Call 初期化

    Call SetRace(44, 1)
    Call SetLean(2, 4, "12903")
    Call SetLean(3, 5, "10568")
    Call SetLean(4, 6, "10494")
    Call SetLean(5, 7, "5811")
    Call SetLean(6, 8, "12655")
    Call 登録
    Call 初期化

    Call SetRace(44, 2)
    Call SetLean(1, 3, "5530")
    Call SetLean(2, 4, "5019")
    Call SetLean(3, 5, "4340")
    Call SetLean(4, 6, "4377")
    Call SetLean(5, 7, "4793")
    Call SetLean(6, 8, "5455")
    Call 登録
    Call 順位決定
    Call 初期化

    Call SetRace(45, 1)
    Call SetLean(1, 3, "5597")
    Call SetLean(2, 4, "5079")
    Call SetLean(3, 5, "5992")
    Call SetLean(4, 6, "4862")
    Call SetLean(5, 7, "4438")
    Call SetLean(6, 8, "5421")
    Call 登録
    Call 初期化

    Call SetRace(45, 2)
    Call SetLean(1, 4, "5648")
    Call SetLean(2, 5, "4434")
    Call SetLean(3, 6, "4434")
    Call SetLean(4, 7, "3886")
    Call SetLean(5, 8, "4741")
    Call 登録
    Call 初期化

    Call SetRace(45, 3)
    Call SetLean(1, 3, "4250")
    Call SetLean(2, 4, "4029")
    Call SetLean(3, 5, "3897")
    Call SetLean(4, 6, "3947")
    Call SetLean(5, 7, "3427")
    Call SetLean(6, 8, "3565")
    Call 登録
    Call 順位決定
    Call 初期化

    Call SetRace(46, 1)
    Call SetLean(1, 3, "4472")
    Call SetLean(2, 4, "5380")
    Call SetLean(3, 5, "3590")
    Call SetLean(4, 6, "5410")
    Call SetLean(5, 7, "10543")
    Call SetLean(6, 8, "4431")
    Call SetLean(7, 9, "4593")
    Call 登録
    Call 初期化

    Call SetRace(46, 2)
    Call SetLean(1, 3, "4019")
    Call SetLean(2, 4, "3704")
    Call SetLean(3, 5, "3372")
    Call SetLean(4, 6, "3945")
    Call SetLean(5, 7, "3970")
    Call SetLean(6, 8, "3329")
    Call SetLean(7, 9, "3722")
    Call 登録
    Call 初期化

    Call SetRace(46, 3)
    Call SetLean(1, 3, "3399")
    Call SetLean(2, 4, "2927")
    Call SetLean(3, 5, "3089")
    Call SetLean(4, 6, "3083")
    Call SetLean(5, 7, "3185")
    Call SetLean(6, 8, "3200")
    Call SetLean(7, 9, "3143")
    Call 登録
    Call 初期化

    Call SetRace(46, 4)
    Call SetLean(1, 3, "3415")
    Call SetLean(2, 4, "3109")
    Call SetLean(3, 5, "3184")
    Call SetLean(4, 6, "3151")
    Call SetLean(5, 7, "2655")
    Call SetLean(6, 8, "2665")
    Call SetLean(7, 9, "2599")
    Call 登録
    Call 順位決定
    Call 初期化

    Call SetRace(47, 1)
    Call SetLean(1, 3, "5073")
    Call SetLean(2, 4, "4746")
    Call SetLean(3, 5, "4559")
    Call SetLean(4, 6, "4347")
    Call SetLean(5, 7, "4631")
    Call SetLean(6, 8, "5017")
    Call SetLean(7, 9, "5421")
    Call 登録
    Call 順位決定
    Call 初期化

    Call SetRace(48, 1)
    Call SetLean(1, 3, "14051")
    Call SetLean(2, 4, "4753")
    Call SetLean(3, 5, "4608")
    Call SetLean(4, 6, "5210")
    Call SetLean(5, 7, "5825")
    Call SetLean(6, 8, "10298")
    Call 登録
    Call 初期化

    Call SetRace(48, 2)
    Call SetLean(1, 3, "5616")
    Call SetLean(2, 4, "4663")
    Call SetLean(3, 5, "3714")
    Call SetLean(4, 6, "3510")
    Call SetLean(5, 7, "")
    Call SetLean(6, 8, "5373")
    Call SetLean(7, 9, "5920")
    Call 登録
    Call 順位決定
    Call 初期化

    Call SetRace(49, 1)
    Call SetLean(2, 4, "10380")
    Call SetLean(3, 5, "10323")
    Call SetLean(4, 6, "5590")
    Call SetLean(5, 7, "5562")
    Call SetLean(6, 8, "10423")
    Call 登録
    Call 初期化

    Call SetRace(49, 2)
    Call SetLean(1, 3, "5943")
    Call SetLean(2, 4, "4975")
    Call SetLean(3, 5, "4646")
    Call SetLean(4, 6, "4755")
    Call SetLean(5, 7, "4708")
    Call SetLean(6, 8, "4998")
    Call 登録
    Call 初期化

    Call SetRace(49, 3)
    Call SetLean(1, 3, "4825")
    Call SetLean(2, 4, "4334")
    Call SetLean(3, 5, "3936")
    Call SetLean(4, 6, "3412")
    Call SetLean(5, 7, "4267")
    Call SetLean(6, 8, "4464")
    Call 登録
    Call 順位決定
    Call 初期化

    Call SetRace(50, 1)
    Call SetLean(1, 3, "10313")
    Call SetLean(2, 4, "10799")
    Call SetLean(3, 5, "5809")
    Call SetLean(4, 6, "5885")
    Call SetLean(5, 7, "10165")
    Call SetLean(6, 8, "10802")
    Call 登録
    Call 初期化

    Call SetRace(50, 2)
    Call SetLean(1, 3, "12327")
    Call SetLean(2, 4, "4970")
    Call SetLean(3, 5, "4615")
    Call SetLean(4, 6, "4176")
    Call SetLean(5, 7, "4655")
    Call SetLean(6, 8, "5041")
    Call 登録
    Call 初期化

    Call SetRace(50, 3)
    Call SetLean(1, 3, "3851")
    Call SetLean(2, 4, "3920")
    Call SetLean(3, 5, "3565")
    Call SetLean(4, 6, "3642")
    Call SetLean(5, 7, "3610")
    Call SetLean(6, 8, "3814")
    Call 登録
    Call 順位決定
    Call 初期化

    Call SetRace(51, 1)
    Call SetLean(2, 4, "")
    Call SetLean(3, 5, "4031")
    Call SetLean(4, 6, "4901")
    Call SetLean(5, 7, "5324")
    Call SetLean(6, 8, "11015")
    Call 登録
    Call 初期化

    Call SetRace(51, 2)
    Call SetLean(2, 4, "4146")
    Call SetLean(3, 5, "4063")
    Call SetLean(4, 6, "3928")
    Call SetLean(5, 7, "4726")
    Call SetLean(6, 8, "4779")
    Call 登録
    Call 初期化

    Call SetRace(51, 3)
    Call SetLean(2, 4, "3720")
    Call SetLean(3, 5, "3404")
    Call SetLean(4, 6, "3343")
    Call SetLean(5, 7, "3427")
    Call SetLean(6, 8, "3587")
    Call 登録
    Call 順位決定
    Call 初期化

    Call SetRace(52, 1)
    Call SetLean(1, 3, "5502")
    Call SetLean(2, 4, "")
    Call SetLean(3, 5, "4846")
    Call SetLean(4, 6, "5145")
    Call SetLean(5, 7, "5834")
    Call SetLean(6, 8, "5675")
    Call 登録
    Call 初期化

    Call SetRace(52, 2)
    Call SetLean(1, 3, "4510")
    Call SetLean(2, 4, "4263")
    Call SetLean(3, 5, "3766")
    Call SetLean(4, 6, "3863")
    Call SetLean(5, 7, "4033")
    Call SetLean(6, 8, "")
    Call SetLean(7, 9, "4648")
    Call 登録
    Call 初期化

    Call SetRace(52, 3)
    Call SetLean(1, 3, "3341")
    Call SetLean(2, 4, "3562")
    Call SetLean(3, 5, "3365")
    Call SetLean(4, 6, "3217")
    Call SetLean(5, 7, "3354")
    Call SetLean(6, 8, "")
    Call SetLean(7, 9, "3336")
    Call 登録
    Call 順位決定
    Call 初期化

    Call SetRace(53, 1)
    Call SetLean(3, 5, "11244")
    Call SetLean(4, 6, "4733")
    Call SetLean(5, 7, "11311")
    Call 登録
    Call 順位決定
    Call 初期化

    Call SetRace(54, 1)
    Call SetLean(2, 4, "")
    Call SetLean(3, 5, "3956")
    Call SetLean(4, 6, "4566")
    Call SetLean(5, 7, "4251")
    Call SetLean(6, 8, "")
    Call 登録
    Call 初期化

    Call SetRace(54, 2)
    Call SetLean(2, 4, "3220")
    Call SetLean(3, 5, "3207")
    Call SetLean(4, 6, "3557")
    Call SetLean(5, 7, "3294")
    Call SetLean(6, 8, "2760")
    Call 登録
    Call 順位決定
    Call 初期化

    Call SetRace(55, 1)
    Call SetLean(3, 5, "5525")
    Call SetLean(4, 6, "4874")
    Call 登録
    Call 順位決定
    Call 初期化

    Call SetRace(56, 1)
    Call SetLean(3, 5, "4025")
    Call SetLean(4, 6, "3930")
    Call SetLean(5, 7, "5488")
    Call 登録
    Call 順位決定
    Call 初期化

    Call SetRace(57, 1)
    Call SetLean(3, 5, "4470")
    Call SetLean(4, 6, "5093")
    Call SetLean(5, 7, "10113")
    Call 登録
    Call 順位決定
    Call 初期化
    
    Call SetRace(58, 1)
    Call SetLean(3, 5, "4307")
    Call SetLean(4, 6, "4481")
    Call SetLean(5, 7, "4518")
    Call 登録
    Call 順位決定
    Call 初期化
    
    Call SetRace(59, 1)
    Call SetLean(1, 3, "4370")
    Call SetLean(2, 4, "3781")
    Call SetLean(3, 5, "3298")
    Call SetLean(4, 6, "3104")
    Call SetLean(5, 7, "3725")
    Call SetLean(6, 8, "4731")
    Call SetLean(7, 9, "5361")
    Call 登録
    Call 順位決定
    Call 初期化
    
    Call SetRace(60, 1)
    Call SetLean(4, 6, "10957")
    Call 登録
    Call 順位決定
    Call 初期化
    
    Call SetRace(61, 1)
    Call SetLean(2, 4, "10738")
    Call SetLean(3, 5, "5395")
    Call SetLean(4, 6, "5182")
    Call SetLean(5, 7, "5092")
    Call SetLean(6, 8, "5730")
    Call 登録
    Call 初期化
    
    Call SetRace(61, 2)
    Call SetLean(2, 4, "5160")
    Call SetLean(3, 5, "4732")
    Call SetLean(4, 6, "4920")
    Call SetLean(5, 7, "4264")
    Call 登録
    Call 順位決定
    Call 初期化
    
    Call SetRace(62, 1)
    Call SetLean(2, 4, "5005")
    Call SetLean(3, 5, "4947")
    Call SetLean(4, 6, "4361")
    Call SetLean(5, 7, "")
    Call 登録
    Call 初期化
    
    Call SetRace(62, 2)
    Call SetLean(2, 4, "4687")
    Call SetLean(3, 5, "5207")
    Call SetLean(4, 6, "4637")
    Call SetLean(5, 7, "")
    Call 登録
    Call 順位決定
    Call 初期化
    
    Call SetRace(63, 1)
    Call SetLean(3, 5, "10867")
    Call SetLean(4, 6, "5831")
    Call 登録
    Call 順位決定
    Call 初期化
    
    Call SetRace(64, 1)
    Call SetLean(2, 4, "11389")
    Call SetLean(3, 5, "")
    Call SetLean(4, 6, "4756")
    Call SetLean(5, 7, "")
    Call SetLean(6, 8, "10765")
    Call 登録
    Call 順位決定
    Call 初期化
    
    Call SetRace(65, 1)
    Call SetLean(1, 3, "11373")
    Call SetLean(2, 4, "10850")
    Call SetLean(3, 5, "5412")
    Call SetLean(4, 6, "5646")
    Call SetLean(5, 7, "10492")
    Call SetLean(6, 8, "11011")
    Call 登録
    Call 順位決定
    Call 初期化
    
    Call SetRace(66, 1)
    Call SetLean(2, 4, "")
    Call SetLean(3, 5, "5590")
    Call SetLean(4, 6, "5300")
    Call SetLean(5, 7, "10270")
    Call 登録
    Call 順位決定
    Call 初期化
    
    Call SetRace(67, 1)
    Call SetLean(3, 5, "10237")
    Call SetLean(4, 6, "4312")
    Call SetLean(5, 7, "5075")
    Call 登録
    Call 順位決定
    Call 初期化

    Call SetRace(68, 1)
    Call SetLean(1, 3, "4509")
    Call SetLean(2, 4, "5073")
    Call SetLean(3, 5, "4589")
    Call SetLean(4, 6, "4222")
    Call SetLean(5, 7, "")
    Call SetLean(6, 8, "4983")
    Call SetLean(7, 9, "5973")
    Call 登録
    Call 順位決定
    Call 初期化

    Call SetRace(69, 1)
    Call SetLean(2, 4, "32483")
    Call SetLean(3, 5, "24343")
    Call SetLean(4, 6, "")
    Call SetLean(5, 7, "23027")
    Call 登録
    Call 順位決定
    Call 初期化

    Call SetRace(70, 1)
    Call SetLean(3, 5, "23512")
    Call SetLean(4, 6, "25694")
    Call SetLean(5, 7, "21541")
    Call 登録
    Call 順位決定
    Call 初期化

    Call SetRace(71, 1)
    Call SetLean(1, 3, "35138")
    Call SetLean(2, 4, "33887")
    Call SetLean(4, 6, "42363")
    Call SetLean(5, 7, "25021")
    Call SetLean(6, 8, "32220")
    Call 登録
    Call 順位決定
    Call 初期化
    
    ' ブックの保存
    ActiveWorkbook.Save
End Sub

