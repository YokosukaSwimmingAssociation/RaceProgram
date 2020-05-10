Attribute VB_Name = "InputSenshukenModule"
'
' 市民大会の自動試験
'
Public Sub 試験選手権大会()

    GetRange("大会名").Value = 選手権大会
    GetRange("大会組最少人数").Value = 3
    GetRange("大会組合せ方式").Value = "単純方式"
    GetRange("大会年").Value = 2019
    GetRange("大会回数").Value = 32
    GetRange("大会元号年").Value = "令和元年"
    GetRange("大会月").Value = 9
    GetRange("大会日").Value = 29

    Call エントリー読込み("C:\Users\新\Documents\水泳協会\競技運営部\選手権申込み")
    Call 組み合わせ決定
    'Call 試験選手権補正
    Call レース番号修正
    Call プログラム作成
    Call 試験選手権予選記録
    Call 試験選手権リレー
    Call 試験選手権決勝記録
    Call 優勝者一覧作成
    Call 大会記録更新

End Sub

'
' 選手権大会のプログラム補正
'
'Public Sub 試験選手権補正()
'
'    Call EventChange(False)
'
'    Sheets("エントリー一覧").Activate
'
'    Call EventChange(True)
'
'    ' ブックの保存
'    ActiveWorkbook.Save
'End Sub

'
' 令和元年度のプログラムの記録を入れる
'
Public Sub 試験選手権予選記録()
    Sheets("記録画面").Select
    ' 2 男子 50M 背泳ぎ 予選
    Call SetRace(2, 1)
    Call SetLean(3, 5, "4112")
    Call SetLean(4, 6, "4070")
    Call SetLean(5, 7, "3800")
    Call SetLean(6, 8, "3961")
    Call 登録
    Call 初期化
    
    Call SetRace(2, 2)
    Call SetLean(1, 3, "3787")
    Call SetLean(2, 4, "3515")
    Call SetLean(3, 5, "")
    Call SetLean(4, 6, "3505")
    Call SetLean(5, 7, "3594")
    Call SetLean(6, 8, "3519")
    Call SetLean(7, 9, "3489")
    Call 登録
    Call 初期化
    
    Call SetRace(2, 3)
    Call SetLean(1, 3, "3599")
    Call SetLean(2, 4, "")
    Call SetLean(3, 5, "3052")
    Call SetLean(4, 6, "2986")
    Call SetLean(5, 7, "3075")
    Call SetLean(6, 8, "3036")
    Call SetLean(7, 9, "3315")
    Call 登録
    Call 順位決定
    Call 決勝登録
    Call 初期化
    
    ' 3 女子 50M バタフライ 予選
    Call SetRace(3, 1)
    Call SetLean(1, 3, "4114")
    Call SetLean(2, 4, "", "スタート失格")
    Call SetLean(3, 5, "3729")
    Call SetLean(4, 6, "3589")
    Call SetLean(5, 7, "3569")
    Call SetLean(6, 8, "3613")
    Call SetLean(7, 9, "3866")
    Call 登録
    Call 初期化
    
    Call SetRace(3, 2)
    Call SetLean(1, 3, "3331")
    Call SetLean(2, 4, "3375")
    Call SetLean(3, 5, "3361")
    Call SetLean(4, 6, "3308")
    Call SetLean(5, 7, "3449")
    Call SetLean(6, 8, "3313")
    Call SetLean(7, 9, "3314")
    Call 登録
    Call 順位決定
    Call 決勝登録
    Call 初期化
    
    ' 4 男子 50M バタフライ 予選
    Call SetRace(4, 1)
    Call SetLean(3, 5, "4018")
    Call SetLean(4, 6, "3508")
    Call SetLean(5, 7, "3491")
    Call 登録
    Call 初期化
    
    Call SetRace(4, 2)
    Call SetLean(1, 4, "3349")
    Call SetLean(2, 5, "", "スタート失格")
    Call SetLean(3, 6, "3195")
    Call SetLean(4, 7, "3224")
    Call SetLean(5, 8, "3276")
    Call SetLean(6, 9, "", "失格")
    Call 登録
    Call 初期化
    
    Call SetRace(4, 3)
    Call SetLean(1, 3, "3298")
    Call SetLean(2, 4, "")
    Call SetLean(3, 5, "3144")
    Call SetLean(4, 6, "3039")
    Call SetLean(5, 7, "3051")
    Call SetLean(6, 8, "3101")
    Call SetLean(7, 9, "")
    Call 登録
    Call 初期化
    
    Call SetRace(4, 4)
    Call SetLean(1, 3, "2977")
    Call SetLean(2, 4, "3009")
    Call SetLean(3, 5, "2682")
    Call SetLean(4, 6, "", "スタート失格")
    Call SetLean(5, 7, "2812")
    Call SetLean(6, 8, "2801")
    Call SetLean(7, 9, "3112")
    Call 登録
    Call 順位決定
    Call 決勝登録
    Call 初期化
    
    ' 5 女子 50M 平泳ぎ 予選
    Call SetRace(5, 1)
    Call SetLean(3, 5, "4715")
    Call SetLean(4, 6, "4390")
    Call SetLean(5, 7, "4366")
    Call 登録
    Call 初期化
    
    Call SetRace(5, 2)
    Call SetLean(1, 3, "4531")
    Call SetLean(2, 4, "4305")
    Call SetLean(3, 5, "4016")
    Call SetLean(4, 6, "3992")
    Call SetLean(5, 7, "4064")
    Call SetLean(6, 8, "4285")
    Call SetLean(7, 9, "4404")
    Call 登録
    Call 順位決定
    Call 決勝登録
    Call 初期化
    
    ' 6 男子 50M 平泳ぎ 予選
    Call SetRace(6, 1)
    Call SetLean(3, 5, "")
    Call SetLean(4, 6, "")
    Call SetLean(5, 7, "4060")
    Call 登録
    Call 初期化
    
    Call SetRace(6, 2)
    Call SetLean(1, 4, "4280")
    Call SetLean(2, 5, "4065")
    Call SetLean(3, 6, "", "スタート失格")
    Call SetLean(4, 7, "3991")
    Call SetLean(5, 8, "4207")
    Call SetLean(6, 9, "4062")
    Call 登録
    Call 初期化
    
    Call SetRace(6, 3)
    Call SetLean(1, 3, "3888")
    Call SetLean(2, 4, "3942")
    Call SetLean(3, 5, "3718")
    Call SetLean(4, 6, "3741")
    Call SetLean(5, 7, "3843")
    Call SetLean(6, 8, "3817")
    Call SetLean(7, 9, "3969")
    Call 登録
    Call 初期化
    
    Call SetRace(6, 4)
    Call SetLean(1, 3, "3759")
    Call SetLean(2, 4, "3710")
    Call SetLean(3, 5, "3464")
    Call SetLean(4, 6, "3208", "OP")
    Call SetLean(5, 7, "3489")
    Call SetLean(6, 8, "3475")
    Call SetLean(7, 9, "")
    Call 登録
    Call 順位決定
    Call 決勝登録
    Call 初期化
    
    ' 7 女子 50M 自由形 予選
    Call SetRace(7, 1)
    Call SetLean(2, 4, "")
    Call SetLean(3, 5, "3358")
    Call SetLean(4, 6, "3665")
    Call SetLean(5, 7, "")
    Call SetLean(6, 8, "3840")
    Call 登録
    Call 初期化
    
    Call SetRace(7, 2)
    Call SetLean(1, 3, "3644")
    Call SetLean(2, 4, "3028")
    Call SetLean(3, 5, "3374")
    Call SetLean(4, 6, "", "スタート失格")
    Call SetLean(5, 7, "3488")
    Call SetLean(6, 8, "3452")
    Call SetLean(7, 9, "3522")
    Call 登録
    Call 初期化
    
    Call SetRace(7, 3)
    Call SetLean(1, 3, "3439")
    Call SetLean(2, 4, "3407")
    Call SetLean(3, 5, "3314")
    Call SetLean(4, 6, "3168")
    Call SetLean(5, 7, "3138")
    Call SetLean(6, 8, "3301")
    Call SetLean(7, 9, "3467")
    Call 登録
    Call 初期化
    
    Call SetRace(7, 4)
    Call SetLean(1, 3, "3384")
    Call SetLean(2, 4, "3253")
    Call SetLean(3, 5, "3265")
    Call SetLean(4, 6, "2994")
    Call SetLean(5, 7, "3199")
    Call SetLean(6, 8, "3152")
    Call SetLean(7, 9, "3275")
    Call 登録
    Call 初期化
    
    Call SetRace(7, 5)
    Call SetLean(1, 3, "3197")
    Call SetLean(2, 4, "3050")
    Call SetLean(3, 5, "2961")
    Call SetLean(4, 6, "2971")
    Call SetLean(5, 7, "2883")
    Call SetLean(6, 8, "2998")
    Call SetLean(7, 9, "3072")
    Call 登録
    Call 順位決定
    Call 決勝登録
    Call 初期化
    
    ' 8 男子 50M 自由形 予選
    Call SetRace(8, 1)
    Call SetLean(1, 3, "3234")
    Call SetLean(2, 4, "3105")
    Call SetLean(3, 5, "3277")
    Call SetLean(4, 6, "2961")
    Call SetLean(5, 7, "3337")
    Call SetLean(6, 8, "3029")
    Call SetLean(7, 9, "3018")
    Call 登録
    Call 初期化
    
    Call SetRace(8, 2)
    Call SetLean(1, 3, "3288")
    Call SetLean(2, 4, "3239")
    Call SetLean(3, 5, "3242")
    Call SetLean(4, 6, "2911")
    Call SetLean(5, 7, "3101")
    Call SetLean(6, 8, "3104")
    Call SetLean(7, 9, "3492")
    Call 登録
    Call 初期化
    
    Call SetRace(8, 3)
    Call SetLean(1, 3, "3095")
    Call SetLean(2, 4, "3081")
    Call SetLean(3, 5, "3070")
    Call SetLean(4, 6, "3070")
    Call SetLean(5, 7, "2986")
    Call SetLean(6, 8, "3198")
    Call SetLean(7, 9, "3076")
    Call 登録
    Call 初期化
    
    Call SetRace(8, 4)
    Call SetLean(1, 3, "3013")
    Call SetLean(2, 4, "3036")
    Call SetLean(3, 5, "3001")
    Call SetLean(4, 6, "2975")
    Call SetLean(5, 7, "")
    Call SetLean(6, 8, "3058")
    Call SetLean(7, 9, "3015")
    Call 登録
    Call 初期化
    
    Call SetRace(8, 5)
    Call SetLean(1, 3, "3039")
    Call SetLean(2, 4, "", "スタート失格")
    Call SetLean(3, 5, "2885")
    Call SetLean(4, 6, "3037")
    Call SetLean(5, 7, "2967")
    Call SetLean(6, 8, "2941")
    Call SetLean(7, 9, "2948")
    Call 登録
    Call 順位決定
    Call 初期化
    
    Call SetRace(8, 6)
    Call SetLean(1, 3, "2977")
    Call SetLean(2, 4, "3058")
    Call SetLean(3, 5, "2768")
    Call SetLean(4, 6, "2921")
    Call SetLean(5, 7, "")
    Call SetLean(6, 8, "2893")
    Call SetLean(7, 9, "2952")
    Call 登録
    Call 初期化
    
    Call SetRace(8, 7)
    Call SetLean(1, 3, "2966")
    Call SetLean(2, 4, "2943")
    Call SetLean(3, 5, "2796")
    Call SetLean(4, 6, "2859")
    Call SetLean(5, 7, "")
    Call SetLean(6, 8, "2839")
    Call SetLean(7, 9, "2837")
    Call 登録
    Call 初期化
    
    Call SetRace(8, 8)
    Call SetLean(1, 3, "")
    Call SetLean(2, 4, "2862")
    Call SetLean(3, 5, "2798")
    Call SetLean(4, 6, "2832")
    Call SetLean(5, 7, "2836")
    Call SetLean(6, 8, "2812")
    Call SetLean(7, 9, "2826")
    Call 登録
    Call 初期化
    
    Call SetRace(8, 9)
    Call SetLean(1, 3, "2841")
    Call SetLean(2, 4, "2871")
    Call SetLean(3, 5, "2856")
    Call SetLean(4, 6, "2671")
    Call SetLean(5, 7, "2763")
    Call SetLean(6, 8, "2790")
    Call SetLean(7, 9, "2879")
    Call 登録
    Call 初期化
    
    Call SetRace(8, 10)
    Call SetLean(1, 3, "2751")
    Call SetLean(2, 4, "2746")
    Call SetLean(3, 5, "2718")
    Call SetLean(4, 6, "2685")
    Call SetLean(5, 7, "2749")
    Call SetLean(6, 8, "")
    Call SetLean(7, 9, "2765")
    Call 登録
    Call 初期化

    Call SetRace(8, 11)
    Call SetLean(1, 3, "2739")
    Call SetLean(2, 4, "2656")
    Call SetLean(3, 5, "2590")
    Call SetLean(4, 6, "2639")
    Call SetLean(5, 7, "2549")
    Call SetLean(6, 8, "2583")
    Call SetLean(7, 9, "2647")
    Call 登録
    Call 順位決定
    Call 決勝登録
    Call 初期化
    
    ' 10 男子 100M 背泳ぎ 予選
    Call SetRace(10, 1)
    Call SetLean(1, 3, "13211")
    Call SetLean(2, 4, "11747")
    Call SetLean(3, 5, "11781")
    Call SetLean(4, 6, "11801")
    Call SetLean(5, 7, "12082")
    Call SetLean(6, 8, "11768")
    Call SetLean(7, 9, "12105")
    Call 登録
    Call 初期化
    
    Call SetRace(10, 2)
    Call SetLean(1, 3, "")
    Call SetLean(2, 4, "10952")
    Call SetLean(3, 5, "10737")
    Call SetLean(4, 6, "10586")
    Call SetLean(5, 7, "10729")
    Call SetLean(6, 8, "10950")
    Call SetLean(7, 9, "11309")
    Call 登録
    Call 順位決定
    Call 決勝登録
    Call 初期化
     
    ' 12 男子 100M バタフライ 予選
    Call SetRace(12, 1)
    Call SetLean(2, 4, "12215")
    Call SetLean(3, 5, "11642")
    Call SetLean(4, 6, "10823")
    Call SetLean(5, 7, "11418")
    Call SetLean(6, 8, "12128")
    Call 登録
    Call 初期化
    
    Call SetRace(12, 2)
    Call SetLean(1, 3, "11008")
    Call SetLean(2, 4, "10395")
    Call SetLean(3, 5, "10060")
    Call SetLean(4, 6, "10020")
    Call SetLean(5, 7, "10120")
    Call SetLean(6, 8, "10504")
    Call SetLean(7, 9, "10817")
    Call 登録
    Call 順位決定
    Call 決勝登録
    Call 初期化
    
    ' 13 女子 100M 平泳ぎ 予選
    Call SetRace(13, 1)
    Call SetLean(3, 5, "13527")
    Call SetLean(4, 6, "13537")
    Call SetLean(5, 7, "13295")
    Call 登録
    Call 初期化
    
    Call SetRace(13, 2)
    Call SetLean(2, 4, "12982")
    Call SetLean(3, 5, "13174")
    Call SetLean(4, 6, "12403")
    Call SetLean(5, 7, "12850")
    Call SetLean(6, 8, "12898")
    Call SetLean(7, 9, "13994")
    Call 登録
    Call 順位決定
    Call 決勝登録
    Call 初期化

    ' 14 男子 100M 平泳ぎ 予選
    Call SetRace(14, 1)
    Call SetLean(2, 4, "13127")
    Call SetLean(3, 5, "12764")
    Call SetLean(4, 6, "")
    Call SetLean(5, 7, "12397")
    Call SetLean(6, 8, "", "スタート失格")
    Call SetLean(7, 9, "13544")
    Call 登録
    Call 初期化
    
    Call SetRace(14, 2)
    Call SetLean(1, 3, "12981")
    Call SetLean(2, 4, "12675")
    Call SetLean(3, 5, "11770")
    Call SetLean(4, 6, "12265")
    Call SetLean(5, 7, "12212")
    Call SetLean(6, 8, "12276")
    Call SetLean(7, 9, "12153")
    Call 登録
    Call 初期化
    
    Call SetRace(14, 3)
    Call SetLean(1, 3, "11739")
    Call SetLean(2, 4, "11628")
    Call SetLean(3, 5, "")
    Call SetLean(4, 6, "11160")
    Call SetLean(5, 7, "11112", "OP")
    Call SetLean(6, 8, "11172")
    Call SetLean(7, 9, "11287")
    Call 登録
    Call 順位決定
    Call 決勝登録
    Call 初期化

    ' 15 女子 100M 自由形 予選
    Call SetRace(15, 1)
    Call SetLean(3, 5, "")
    Call SetLean(4, 6, "11186")
    Call SetLean(5, 7, "11163")
    Call 登録
    Call 初期化
    
    Call SetRace(15, 2)
    Call SetLean(2, 4, "11726")
    Call SetLean(3, 5, "11583")
    Call SetLean(4, 6, "11145")
    Call SetLean(5, 7, "11297")
    Call SetLean(6, 8, "11561")
    Call SetLean(7, 9, "11756")
    Call 登録
    Call 初期化
    
    Call SetRace(15, 3)
    Call SetLean(1, 3, "10971")
    Call SetLean(2, 4, "10670")
    Call SetLean(3, 5, "10776")
    Call SetLean(4, 6, "")
    Call SetLean(5, 7, "10598")
    Call SetLean(6, 8, "10644")
    Call SetLean(7, 9, "11004")
    Call 登録
    Call 初期化
    
    Call SetRace(15, 4)
    Call SetLean(1, 3, "10538")
    Call SetLean(2, 4, "10384")
    Call SetLean(3, 5, "10490")
    Call SetLean(4, 6, "10227")
    Call SetLean(5, 7, "10611")
    Call SetLean(6, 8, "10688")
    Call SetLean(7, 9, "10728")
    Call 登録
    Call 順位決定
    Call 決勝登録
    Call 初期化

    ' 16 男子 100M 自由形 予選
    Call SetRace(16, 1)
    Call SetLean(3, 5, "10821")
    Call SetLean(4, 6, "10847")
    Call SetLean(5, 7, "11592")
    Call 登録
    Call 初期化
    
    Call SetRace(16, 2)
    Call SetLean(2, 4, "11621")
    Call SetLean(3, 5, "10633")
    Call SetLean(4, 6, "10894")
    Call SetLean(5, 7, "")
    Call SetLean(6, 8, "10756")
    Call 登録
    Call 初期化
    
    Call SetRace(16, 3)
    Call SetLean(1, 3, "")
    Call SetLean(2, 4, "10831")
    Call SetLean(3, 5, "10579")
    Call SetLean(4, 6, "", "スタート失格")
    Call SetLean(5, 7, "10466")
    Call SetLean(6, 8, "")
    Call SetLean(7, 9, "")
    Call 登録
    Call 初期化
    
    Call SetRace(16, 4)
    Call SetLean(1, 3, "10631")
    Call SetLean(2, 4, "10418")
    Call SetLean(3, 5, "10536")
    Call SetLean(4, 6, "10561")
    Call SetLean(5, 7, "10492")
    Call SetLean(6, 8, "10783")
    Call SetLean(7, 9, "10548")
    Call 登録
    Call 初期化
    
    Call SetRace(16, 5)
    Call SetLean(1, 3, "10466")
    Call SetLean(2, 4, "10584")
    Call SetLean(3, 5, "10331")
    Call SetLean(4, 6, "10417")
    Call SetLean(5, 7, "10387")
    Call SetLean(6, 8, "10506")
    Call SetLean(7, 9, "")
    Call 登録
    Call 初期化
    
    Call SetRace(16, 6)
    Call SetLean(1, 3, "10178")
    Call SetLean(2, 4, "10137")
    Call SetLean(3, 5, "10169")
    Call SetLean(4, 6, "10064")
    Call SetLean(5, 7, "10112")
    Call SetLean(6, 8, "10279")
    Call SetLean(7, 9, "10362")
    Call 登録
    Call 初期化
    
    Call SetRace(16, 7)
    Call SetLean(1, 3, "10075")
    Call SetLean(2, 4, "5965")
    Call SetLean(3, 5, "5783")
    Call SetLean(4, 6, "5743")
    Call SetLean(5, 7, "10042")
    Call SetLean(6, 8, "5941")
    Call SetLean(7, 9, "10126")
    Call 登録
    Call 初期化
    
    Call SetRace(16, 8)
    Call SetLean(1, 3, "5966")
    Call SetLean(2, 4, "5888")
    Call SetLean(3, 5, "5843")
    Call SetLean(4, 6, "5521")
    Call SetLean(5, 7, "5551")
    Call SetLean(6, 8, "5831")
    Call SetLean(7, 9, "")
    Call 登録
    Call 順位決定
    Call 決勝登録
    Call 初期化

End Sub

Public Sub 試験選手権リレー()

    Sheets(プログラムシート).Select

    ' 27 女子 4×50M メドレーリレー
    Call ModCell(SearchRelayCell(27, "追浜高校", "", "Prog氏名"), "鈴木．三浦．高橋．高野")
    Call ModCell(SearchRelayCell(27, "KSS横須賀", "", "Prog氏名"), "日向．緒形．佐野．吉田")
    Call ModCell(SearchRelayCell(27, "三浦学苑高校", "", "Prog氏名"), "杉山．増田．福島．二宮")
    Call ModCell(SearchRelayCell(27, "久里浜中学校", "", "Prog氏名"), "野村．今岡．高橋．杉山")
    Call ModCell(SearchRelayCell(27, "オアシスよこすか", "", "Prog氏名"), "岩田．小林．猪熊．増田")
    Call ModCell(SearchRelayCell(27, "逗子高校", "", "Prog氏名"), "鈴木．勝山．間瀬．村田")
    
    ' 28 男子 4×50M メドレーリレー
    Call ModCell(SearchRelayCell(28, "横須賀大津高校", "", "Prog氏名"), "鈴木．渡邊．神山．川畑")
    Call ModCell(SearchRelayCell(28, "県立横須賀高校", "", "Prog氏名"), "佐野．松永．二本木．田中")
    Call ModCell(SearchRelayCell(28, "追浜高校", "", "Prog氏名"), "村野．梶谷．小林．光田")
    
    Call ModCell(SearchRelayCell(28, "久里浜中学校", "", "Prog氏名"), "脇．小澤．谷口．新本T")
    Call ModCell(SearchRelayCell(28, "大津中学", "", "Prog氏名"), "矢部．清水．後藤．皆越")
    Call ModCell(SearchRelayCell(28, "野比中学", "", "Prog氏名"), "小菅．山田．小澤．松浦")
    Call ModCell(SearchRelayCell(28, "逗子開成高校", "", "Prog氏名"), "岡村．井上．清水．丸山")
    Call ModCell(SearchRelayCell(28, "三浦学苑高校", "", "Prog氏名"), "村上．佐枝．堀口．前場")
    Call ModCell(SearchRelayCell(28, "TeamOGS", "", "Prog氏名"), "平田．関．杉山．石田")
    Call ModCell(SearchRelayCell(28, "逗子開成中学", "", "Prog氏名"), "石田．小湊．石井．松本")
    
    ' 45 女子 4×50M フリーリレー
    Call ModCell(SearchRelayCell(45, "オアシスよこすか", "", "Prog氏名"), "猪熊．増田．岩田．小林")
    Call ModCell(SearchRelayCell(45, "KSS横須賀", "", "Prog氏名"), "吉田．日向．緒形．佐野")
    Call ModCell(SearchRelayCell(45, "久里浜中学校", "", "Prog氏名"), "野村．高橋．岩波．杉山")
    Call ModCell(SearchRelayCell(45, "三浦学苑高校", "", "Prog氏名"), "二宮．増田．三中．福島")
    Call ModCell(SearchRelayCell(45, "追浜高校", "", "Prog氏名"), "鈴木．高野．三浦．高橋")
    Call ModCell(SearchRelayCell(45, "逗子高校", "", "Prog氏名"), "鈴木．間瀬．勝山．村田")
    
    ' 46 男子 4×50M フリーリレー
    Call ModCell(SearchRelayCell(46, "逗子高校", "", "Prog氏名"), "福岡．堀江．水島．高坂")
    Call ModCell(SearchRelayCell(46, "追浜高校", "", "Prog氏名"), "小林．村野．光田．梶谷")
    Call ModCell(SearchRelayCell(46, "大津中学", "", "Prog氏名"), "後藤．皆越．小林．清水")
    Call ModCell(SearchRelayCell(46, "久里浜中学校", "", "Prog氏名"), "新本．高橋．谷口．小澤")
    Call ModCell(SearchRelayCell(46, "津久井浜高校", "", "Prog氏名"), "後藤．荒木．渡辺．佐久間")
    Call ModCell(SearchRelayCell(46, "逗子開成中学", "", "Prog氏名"), "石田．松本．西田．石井")
    Call ModCell(SearchRelayCell(46, "横須賀大津高校", "", "Prog氏名"), "神山．渡邊．川畑．山内")
    Call ModCell(SearchRelayCell(46, "TeamOGS", "", "Prog氏名"), "石田．杉山．熱田．平田")
    Call ModCell(SearchRelayCell(46, "三浦学苑高校", "", "Prog氏名"), "前場．堀口．佐枝．村上")
    Call ModCell(SearchRelayCell(46, "逗子開成高校", "", "Prog氏名"), "丸山．岡村．井上．白石")
    Call ModCell(SearchRelayCell(46, "野比中学", "", "Prog氏名"), "小澤．小菅．川島．山田")
    Call ModCell(SearchRelayCell(46, "県立横須賀高校", "", "Prog氏名"), "佐野．佐々木．二本木．田中")


End Sub

Public Sub 試験選手権決勝記録()
    Sheets("記録画面").Select

    ' 17 女子 200M 個人メドレー 決勝
    Call SetRace(17, 1)
    Call SetLean(3, 5, "32193")
    Call SetLean(4, 6, "31578")
    Call SetLean(5, 7, "32097")
    Call SetLean(6, 8, "32957")
    Call 登録
    Call 初期化
    
    Call SetRace(17, 2)
    Call SetLean(1, 3, "30927")
    Call SetLean(2, 4, "30384")
    Call SetLean(3, 5, "24751")
    Call SetLean(4, 6, "23682")
    Call SetLean(5, 7, "24920")
    Call SetLean(6, 8, "25636")
    Call SetLean(7, 9, "30448")
    Call 登録
    Call 順位決定
    Call 初期化
    
    ' 18 男子 200M 個人メドレー 決勝
    Call SetRace(18, 1)
    Call SetLean(3, 5, "23085")
    Call SetLean(4, 6, "25990")
    Call SetLean(5, 7, "30024")
    Call 登録
    Call 初期化
    
    Call SetRace(18, 2)
    Call SetLean(2, 4, "30396")
    Call SetLean(3, 5, "25733")
    Call SetLean(4, 6, "22217")
    Call SetLean(5, 7, "23127")
    Call SetLean(6, 8, "25535")
    Call SetLean(7, 9, "25736")
    Call 登録
    Call 初期化
    
    Call SetRace(18, 3)
    Call SetLean(1, 3, "22038")
    Call SetLean(2, 4, "22646")
    Call SetLean(3, 5, "24005")
    Call SetLean(4, 6, "24195")
    Call SetLean(5, 7, "24792")
    Call SetLean(6, 8, "25084")
    Call SetLean(7, 9, "24905")
    Call 登録
    Call 初期化
    
    Call SetRace(18, 4)
    Call SetLean(1, 3, "23627")
    Call SetLean(2, 4, "23006")
    Call SetLean(3, 5, "23722")
    Call SetLean(4, 6, "22564")
    Call SetLean(5, 7, "22767")
    Call SetLean(6, 8, "22968")
    Call SetLean(7, 9, "24176")
    Call 登録
    Call 初期化
    
    Call SetRace(18, 5)
    Call SetLean(1, 3, "22456")
    Call SetLean(2, 4, "22628")
    Call SetLean(3, 5, "22197")
    Call SetLean(4, 6, "21254")
    Call SetLean(5, 7, "21700")
    Call SetLean(6, 8, "21459")
    Call SetLean(7, 9, "22285")
    Call 登録
    Call 順位決定
    Call 初期化
    
    ' 19 女子 200M 背泳ぎ 決勝
    Call SetRace(19, 1)
    Call SetLean(4, 6, "23948")
    Call 登録
    Call 順位決定
    Call 初期化
    
    ' 20 男子 200M 背泳ぎ 決勝
    Call SetRace(20, 1)
    Call SetLean(2, 4, "24444")
    Call SetLean(3, 5, "")
    Call SetLean(4, 6, "22160")
    Call SetLean(5, 7, "22191")
    Call SetLean(6, 8, "24323")
    Call 登録
    Call 順位決定
    Call 初期化
    
    ' 22 男子 200M バタフライ 決勝
    Call SetRace(22, 1)
    Call SetLean(2, 4, "24085")
    Call SetLean(3, 5, "22586")
    Call SetLean(4, 6, "21589")
    Call SetLean(5, 7, "21574")
    Call SetLean(6, 8, "22825")
    Call SetLean(7, 9, "23472")
    Call 登録
    Call 順位決定
    Call 初期化
    
    ' 23 女子 200M 平泳ぎ 決勝
    Call SetRace(23, 1)
    Call SetLean(2, 4, "31738")
    Call SetLean(3, 5, "32097")
    Call SetLean(4, 6, "30712")
    Call SetLean(5, 7, "30568")
    Call SetLean(6, 8, "33634")
    Call SetLean(7, 9, "33684")
    Call 登録
    Call 順位決定
    Call 初期化
    
    ' 24 男子 200M 平泳ぎ 決勝
    Call SetRace(24, 1)
    Call SetLean(1, 3, "31244")
    Call SetLean(2, 4, "30046")
    Call SetLean(3, 5, "23874")
    Call SetLean(4, 6, "")
    Call SetLean(5, 7, "24575")
    Call SetLean(6, 8, "25526")
    Call SetLean(7, 9, "30580")
    Call 登録
    Call 順位決定
    Call 初期化
    
    ' 25 女子 200M 自由形 決勝
    Call SetRace(25, 1)
    Call SetLean(2, 4, "25877")
    Call SetLean(3, 5, "25123")
    Call SetLean(4, 6, "23285")
    Call SetLean(5, 7, "24311")
    Call SetLean(6, 8, "24433")
    Call SetLean(7, 9, "31132")
    Call 登録
    Call 初期化
    
    Call SetRace(25, 2)
    Call SetLean(1, 3, "22393")
    Call SetLean(2, 4, "21747")
    Call SetLean(3, 5, "21789")
    Call SetLean(4, 6, "21206")
    Call SetLean(5, 7, "22218")
    Call SetLean(6, 8, "21777")
    Call SetLean(7, 9, "23182")
    Call 登録
    Call 順位決定
    Call 初期化
    
    ' 26 男子 200M 自由形 決勝
    Call SetRace(26, 1)
    Call SetLean(3, 5, "25197")
    Call SetLean(4, 6, "24034")
    Call SetLean(5, 7, "22936")
    Call 登録
    Call 初期化
    
    Call SetRace(26, 2)
    Call SetLean(2, 4, "23384")
    Call SetLean(3, 5, "23123")
    Call SetLean(4, 6, "23138")
    Call SetLean(5, 7, "")
    Call SetLean(6, 8, "")
    Call SetLean(7, 9, "22773")
    Call 登録
    Call 初期化
    
    Call SetRace(26, 3)
    Call SetLean(1, 3, "22693")
    Call SetLean(2, 4, "22460")
    Call SetLean(3, 5, "21594")
    Call SetLean(4, 6, "22000")
    Call SetLean(5, 7, "21858")
    Call SetLean(6, 8, "22105")
    Call SetLean(7, 9, "22491")
    Call 登録
    Call 初期化
    
    Call SetRace(26, 4)
    Call SetLean(1, 3, "21169")
    Call SetLean(2, 4, "21003")
    Call SetLean(3, 5, "20905")
    Call SetLean(4, 6, "21363")
    Call SetLean(5, 7, "20435")
    Call SetLean(6, 8, "21306")
    Call SetLean(7, 9, "21292")
    Call 登録
    Call 初期化
    
    Call SetRace(26, 5)
    Call SetLean(1, 3, "20861")
    Call SetLean(2, 4, "20259")
    Call SetLean(3, 5, "20279")
    Call SetLean(4, 6, "15961")
    Call SetLean(5, 7, "21003")
    Call SetLean(6, 8, "21383")
    Call SetLean(7, 9, "20153")
    Call 登録
    Call 順位決定
    Call 初期化
    
    ' 27 女子 4×50M メドレーリレー 決勝
    Call SetRace(27, 1)
    Call SetLean(2, 4, "23406")
    Call SetLean(3, 5, "21899")
    Call SetLean(4, 6, "21147")
    Call SetLean(5, 7, "22430")
    Call SetLean(6, 8, "22744")
    Call SetLean(7, 9, "22599")
    Call 登録
    Call 順位決定
    Call 初期化
    
    ' 28 男子 4×50M メドレーリレー 決勝
    Call SetRace(28, 1)
    Call SetLean(3, 5, "20969")
    Call SetLean(4, 6, "21210")
    Call SetLean(5, 7, "21633")
    Call 登録
    Call 初期化
    
    Call SetRace(28, 2)
    Call SetLean(1, 3, "21786")
    Call SetLean(2, 4, "21063")
    Call SetLean(3, 5, "15908")
    Call SetLean(4, 6, "15174")
    Call SetLean(5, 7, "14823")
    Call SetLean(6, 8, "15949")
    Call SetLean(7, 9, "20758")
    Call 登録
    Call 順位決定
    Call 初期化
    
    ' 29 女子 50M 背泳ぎ 決勝
    Call SetRace(29, 1)
    Call SetLean(1, 3, "4258")
    Call SetLean(2, 4, "4596")
    Call SetLean(3, 5, "4250")
    Call SetLean(4, 6, "3187")
    Call SetLean(5, 7, "3489")
    Call SetLean(6, 8, "3780")
    Call SetLean(7, 9, "4112")
    Call 登録
    Call 順位決定
    Call 初期化
    
    ' 30 男子 50M 背泳ぎ 決勝
    Call SetRace(30, 1)
    Call SetLean(1, 3, "3496")
    Call SetLean(2, 4, "3217")
    Call SetLean(3, 5, "3060")
    Call SetLean(4, 6, "2700")
    Call SetLean(5, 7, "3047")
    Call SetLean(6, 8, "2888")
    Call SetLean(7, 9, "3471")
    Call 登録
    Call 順位決定
    Call 初期化
    
    ' 31 女子 50M バタフライ 決勝
    Call SetRace(31, 1)
    Call SetLean(1, 3, "3390")
    Call SetLean(2, 4, "3274")
    Call SetLean(3, 5, "3267")
    Call SetLean(4, 6, "3318")
    Call SetLean(5, 7, "3347")
    Call SetLean(6, 8, "3312")
    Call SetLean(7, 9, "3316")
    Call 登録
    Call 順位決定
    Call 初期化
    
    ' 32 男子 50M バタフライ 決勝
    Call SetRace(32, 1)
    Call SetLean(1, 3, "3012")
    Call SetLean(2, 4, "3076")
    Call SetLean(3, 5, "2806")
    Call SetLean(4, 6, "2720")
    Call SetLean(5, 7, "2746")
    Call SetLean(6, 8, "2969")
    Call SetLean(7, 9, "3071")
    Call 登録
    Call 順位決定
    Call 初期化
    
    ' 33 女子 50M 平泳ぎ 決勝
    Call SetRace(33, 1)
    Call SetLean(1, 3, "4510")
    Call SetLean(2, 4, "4340")
    Call SetLean(3, 5, "3894")
    Call SetLean(4, 6, "3944")
    Call SetLean(5, 7, "4067")
    Call SetLean(6, 8, "4214")
    Call SetLean(7, 9, "4473")
    Call 登録
    Call 順位決定
    Call 初期化
    
    ' 34 男子 50M 平泳ぎ 決勝
    Call SetRace(34, 1)
    Call SetLean(1, 3, "3747")
    Call SetLean(2, 4, "3710")
    Call SetLean(3, 5, "3351")
    Call SetLean(4, 6, "3498")
    Call SetLean(5, 7, "3507")
    Call SetLean(6, 8, "3712")
    Call SetLean(7, 9, "3762")
    Call 登録
    Call 順位決定
    Call 初期化
    
    ' 35 女子 50M 自由形 決勝
    Call SetRace(35, 1)
    Call SetLean(1, 3, "3049")
    Call SetLean(2, 4, "2971")
    Call SetLean(3, 5, "2961")
    Call SetLean(4, 6, "2910")
    Call SetLean(5, 7, "2950")
    Call SetLean(6, 8, "2992")
    Call SetLean(7, 9, "3019")
    Call 登録
    Call 順位決定
    Call 初期化
    
    ' 36 男子 50M 自由形 決勝
    Call SetRace(36, 1)
    Call SetLean(1, 3, "2670")
    Call SetLean(2, 4, "2606")
    Call SetLean(3, 5, "2758")
    Call SetLean(4, 6, "")
    Call SetLean(5, 7, "2557")
    Call SetLean(6, 8, "2462")
    Call SetLean(7, 9, "2634")
    Call 登録
    Call 順位決定
    Call 初期化
    
    ' 37 女子 100M 背泳ぎ 決勝
    Call SetRace(37, 1)
    Call SetLean(1, 4, "12675")
    Call SetLean(2, 5, "11329")
    Call SetLean(3, 6, "11321")
    Call SetLean(4, 7, "11418")
    Call SetLean(5, 8, "11870")
    Call SetLean(6, 9, "12552")
    Call 登録
    Call 順位決定
    Call 初期化
    
    ' 38 男子 100M 背泳ぎ 決勝
    Call SetRace(38, 1)
    Call SetLean(1, 3, "11640")
    Call SetLean(2, 4, "10781")
    Call SetLean(3, 5, "10537")
    Call SetLean(4, 6, "5797")
    Call SetLean(5, 7, "10352")
    Call SetLean(6, 8, "10581")
    Call SetLean(7, 9, "11229")
    Call 登録
    Call 順位決定
    Call 初期化
    
    ' 39 女子 100M バタフライ 決勝
    Call SetRace(39, 1)
    Call SetLean(3, 5, "12500")
    Call SetLean(4, 6, "11064")
    Call SetLean(5, 7, "")
    Call SetLean(6, 8, "12421")
    Call 登録
    Call 順位決定
    Call 初期化
    
    ' 40 男子 100M バタフライ 決勝
    Call SetRace(40, 1)
    Call SetLean(1, 3, "10735")
    Call SetLean(2, 4, "10292")
    Call SetLean(3, 5, "5837")
    Call SetLean(4, 6, "5762")
    Call SetLean(5, 7, "5951")
    Call SetLean(6, 8, "10513")
    Call SetLean(7, 9, "10717")
    Call 登録
    Call 順位決定
    Call 初期化
    
    ' 41 女子 100M 平泳ぎ 決勝
    Call SetRace(41, 1)
    Call SetLean(1, 3, "13416")
    Call SetLean(2, 4, "13173")
    Call SetLean(3, 5, "12786")
    Call SetLean(4, 6, "12398")
    Call SetLean(5, 7, "12492")
    Call SetLean(6, 8, "12797")
    Call SetLean(7, 9, "13259")
    Call 登録
    Call 順位決定
    Call 初期化
    
    ' 42 男子 100M 平泳ぎ 決勝
    Call SetRace(42, 1)
    Call SetLean(1, 3, "12017")
    Call SetLean(2, 4, "11618")
    Call SetLean(3, 5, "11193")
    Call SetLean(4, 6, "10842")
    Call SetLean(5, 7, "10968")
    Call SetLean(6, 8, "11559")
    Call SetLean(7, 9, "11587")
    Call 登録
    Call 順位決定
    Call 初期化
    
    ' 43 女子 100M 自由形 決勝
    Call SetRace(43, 1)
    Call SetLean(1, 3, "10638")
    Call SetLean(2, 4, "10623")
    Call SetLean(3, 5, "10494")
    Call SetLean(4, 6, "10164")
    Call SetLean(5, 7, "10393")
    Call SetLean(6, 8, "10684")
    Call SetLean(7, 9, "10502")
    Call 登録
    Call 順位決定
    Call 初期化
    
    ' 44 男子 100M 自由形 決勝
    Call SetRace(44, 1)
    Call SetLean(1, 3, "5804")
    Call SetLean(2, 4, "5610")
    Call SetLean(3, 5, "5660")
    Call SetLean(4, 6, "5407")
    Call SetLean(5, 7, "5703")
    Call SetLean(6, 8, "5817")
    Call SetLean(7, 9, "5762")
    Call 登録
    Call 順位決定
    Call 初期化
    
    ' 45 女子 4×50M フリーリレー 決勝
    Call SetRace(45, 1)
    Call SetLean(1, 4, "21036")
    Call SetLean(2, 5, "20276")
    Call SetLean(3, 6, "20348")
    Call SetLean(4, 7, "15704")
    Call SetLean(5, 8, "22045")
    Call SetLean(6, 9, "21219")
    Call 登録
    Call 順位決定
    Call 初期化
    
    ' 46 男子 4×50M フリーリレー 決勝
    Call SetRace(46, 1)
    Call SetLean(2, 4, "20186")
    Call SetLean(3, 5, "15683")
    Call SetLean(4, 6, "15668")
    Call SetLean(5, 7, "20125")
    Call SetLean(6, 8, "21219")
    Call 登録
    Call 初期化
    
    Call SetRace(46, 2)
    Call SetLean(1, 3, "15133")
    Call SetLean(2, 4, "15314")
    Call SetLean(3, 5, "14480")
    Call SetLean(4, 6, "13775")
    Call SetLean(5, 7, "13931")
    Call SetLean(6, 8, "15328")
    Call SetLean(7, 9, "15953")
    Call 登録
    Call 順位決定
    Call 初期化
    
    ' ブックの保存
    ActiveWorkbook.Save
End Sub

