Attribute VB_Name = "ShiminCorrectionModule"
'
' s¯åïÌvOâ³
'
Sub s¯vOâ³()

    Call EventChange(False)

    Sheets("Gg[ê").Activate
    ' 1 wq4~50Mh[[
    Call ModCell("F2", "8")
    Call ModCell("G2", "ÎäüDíDîûDËÞDÈ´üØDåàÊT")
    Call ModCell("F3", "9")
    Call ModCell("G3", "´GÞDRcÊÄDc·ÊÞDFJMÞ")
    Call ModCell("F4", "6")
    Call ModCell("G4", "ìºüFD¡ªÇµD´©¢èDûÊ¹")
    Call ModCell("F5", "7")
    Call ModCell("G5", "rìüìDcö¶D`DCDRÉ")
    Call ModCell("F6", "5")
    Call ModCell("G6", "ìJatD¬ÑüD|ºÄèD¬ÑÑØ")
    Call ModCell("F7", "4")
    Call ModCell("G8", "éØúaqDOY^¿D´­éÝDìc")
    
    ' 2|3@¯ê[X
    Call ModCell("F9", "5")
    Call ModCell("G9", "éØDÔûP¬÷DRè¨Dºcüç")
    Call ModCell("C10", "20")
    Call ModCell("F10", "8")
    Call ModCell("G10", "Ö¡zqD{è¾qD{Yß®ÝDéØcq")
    
    ' 4 w jq 4~50M h[[
    Call ModCell("F11", "8")
    Call ModCell("G11", "¡RE¾D¬àVDãD´TíDXoWO")
    Call ModCell("F12", "7")
    Call ModCell("G12", "ez¾NDÍìõ¾DJûc^DV{Cá")
    Call ModCell("F13", "6")
    Call ModCell("G13", "Fzp¾Dã¡]ÈD¨GîDâ£l")
    Call ModCell("F14", "9")
    Call ModCell("G14", "²¡uDì°éDÁäm¾NDRª¾Y")
    Call ModCell("F15", "5")
    Call ModCell("G15", "ymìl¾D·K½DcDlDú±IM")
    Call ModCell("F16", "4")
    Call ModCell("G16", "vìCÚD¼c[lDûüªîD·ö\êY")
    
    ' 5|6@¯ê[X
    Call ModCell("F17", "3")
    Call ModCell("G17", "ªóDn¨DÂr¾Dõ")
    Call ModCell("F18", "4")
    Call ModCell("G18", "mJSlD²ìãÄÆDñ{Ø§DcDMå")
    Call ModCell("F19", "5")
    Call ModCell("G19", "OcTj[DÍãå¹Dc´CõDéØåM")
    Call ModCell("F20", "6")
    Call ModCell("G20", "º£gPDß]éD¾D²ìãÄåDìA")
    Call ModCell("F21", "7")
    Call ModCell("G21", "éØöåDnÓ½D_RD¼{ãå")
    ' 6
    Call ModCell("C22", "50")
    Call ModCell("F22", "8")
    Call ModCell("G22", "éØ¡PDpcS÷D´ÄD{ì_i")
    Call ModCell("C23", "50")
    Call ModCell("F23", "9")
    Call ModCell("G23", "éØC½DÂzGîDO¿¼aDã´Dm")
    
    ' 7|9@¯ê[X
    Call ModCell(SearchCell(7, "ªª@æ", "["), "3")
    Call ModCell(SearchCell(7, "J{@Äq", "["), "4")
    Call ModCell(SearchCell(7, "ìº@üF", "["), "5")
    Call ModCell(SearchCell(7, "|º@Äè", "["), "6")
    Call ModCell(SearchCell(7, "c@ö¶", "["), "7")
    ' 9
    Call ModCell(SearchCell(9, "J@K]", "[XNo"), "70")
    Call ModCell(SearchCell(9, "J@K]", "["), "9")
    
    ' 17-18 ¯ê[X
    Call ModCell(SearchCell(17, "c@DMå", "["), "3")
    Call ModCell(SearchCell(17, "¬àV@q¾", "["), "4")
    Call ModCell(SearchCell(17, "Rà@­Õ", "["), "5")
    Call ModCell(SearchCell(17, "õc@Ðá", "["), "6")
    Call ModCell(SearchCell(17, "½c@¼å", "["), "7")
    ' 18
    Call ModCell(SearchCell(18, "âä@¾¢", "[XNo"), "140")
    Call ModCell(SearchCell(18, "âä@¾¢", "["), "9")

    ' 19-20 ¯ê[X
    Call ModCell(SearchCell(19, "Îä@üDí", "["), "4")
    Call ModCell(SearchCell(19, "~ø@çl", "["), "5")
    Call ModCell(SearchCell(19, "¯i@å", "["), "6")
    ' 20
    Call ModCell(SearchCell(20, "éØ@", "[XNo"), "160")
    Call ModCell(SearchCell(20, "éØ@", "["), "8")

    ' 38-39 ¯ê[X
    Call ModCell(SearchCell(38, "éØ@úaq", "["), "4")
    Call ModCell(SearchCell(38, "¬ì@àYÞ", "["), "5")
    Call ModCell(SearchCell(38, "´@­éÝ", "["), "6")
    Call ModCell(SearchCell(38, "½c@©Ì", "["), "7")
    ' 39
    Call ModCell(SearchCell(39, "éØ@cq", "[XNo"), "320")
    Call ModCell(SearchCell(39, "éØ@cq", "["), "9")

    ' 50-51 ¯ê[X
    Call ModCell(SearchCell(50, "¼c@z", "["), "3")
    ' 51
    Call ModCell(SearchCell(51, "J@K]", "[XNo"), "510")
    Call ModCell(SearchCell(51, "J@K]", "["), "6")
    Call ModCell(SearchCell(51, "{Y@ß®Ý", "[XNo"), "510")
    Call ModCell(SearchCell(51, "{Y@ß®Ý", "["), "7")
    Call ModCell(SearchCell(51, "²ì@Øq", "[XNo"), "510")
    Call ModCell(SearchCell(51, "²ì@Øq", "["), "8")
    Call ModCell(SearchCell(51, "Ö¡@zq", "[XNo"), "510")
    Call ModCell(SearchCell(51, "Ö¡@zq", "["), "9")

   ' 52
    Call ModCell(SearchCell(52, "´@Tí", "["), "4")
    Call ModCell(SearchCell(52, "â£@@l", "[XNo"), "540")
    Call ModCell(SearchCell(52, "â£@@l", "g"), "2")
    Call ModCell(SearchCell(52, "â£@@l", "["), "7")
    Call ModCell(SearchCell(52, "RÝ@¼l", "["), "8")

   ' 54
    Call ModCell(SearchCell(54, "¬Ñ@cY", "["), "4")
    Call ModCell(SearchCell(54, "ÛÎ@ñ", "["), "5")
    Call ModCell(SearchCell(54, "´@rç", "[XNo"), "570")
    Call ModCell(SearchCell(54, "´@rç", "g"), "1")
    Call ModCell(SearchCell(54, "´@rç", "["), "7")

   ' 60 1g(76)
    Call ModCell(SearchCell(60, "R{@áÁl", "["), "3")
    Call ModCell(SearchCell(60, "¬Ñ@cY", "["), "4")
    Call ModCell(SearchCell(60, "OY@Bç", "["), "5")
    Call ModCell(SearchCell(60, "¬@Mm", "["), "7")
    Call ModCell(SearchCell(60, "ÂØ@»ê", "["), "8")
    Call ModCell(SearchCell(60, "ßc@Oj", "["), "9")
    Call ModCell(SearchCell(60, "nÓ@ôj", "[XNo"), "760")
    Call ModCell(SearchCell(60, "nÓ@ôj", "g"), "1")
    Call ModCell(SearchCell(60, "nÓ@ôj", "["), "6")

   ' 60 2g(77)
    Call ModCell(SearchCell(60, "ÛÎ@ñ", "["), "5")
    Call ModCell(SearchCell(60, "cì@@G", "["), "6")
    Call ModCell(SearchCell(60, "xì@Y", "["), "7")

   ' 60 3g(78)
    Call ModCell(SearchCell(60, "À¡@pV", "[XNo"), "780")
    Call ModCell(SearchCell(60, "À¡@pV", "g"), "3")
    Call ModCell(SearchCell(60, "À¡@pV", "["), "4")
    Call ModCell(SearchCell(60, "ìº@ë_", "[XNo"), "780")
    Call ModCell(SearchCell(60, "ìº@ë_", "g"), "3")
    Call ModCell(SearchCell(60, "ìº@ë_", "["), "5")
    Call ModCell(SearchCell(60, "À@ÇF", "[XNo"), "780")
    Call ModCell(SearchCell(60, "À@ÇF", "g"), "3")
    Call ModCell(SearchCell(60, "À@ÇF", "["), "8")
    
    Call ModCell(SearchCell(60, "´@rç", "g"), "3")
    Call ModCell(SearchCell(60, "´@rç", "["), "6")
    Call ModCell(SearchCell(60, "{ì@_i", "g"), "3")
    Call ModCell(SearchCell(60, "{ì@_i", "["), "7")
    Call ModCell(SearchCell(60, "pc@S÷", "g"), "3")
    Call ModCell(SearchCell(60, "pc@S÷", "["), "9")

   ' 60 4g(79)
    Call ModCell(SearchCell(60, "n@@N", "[XNo"), "790")
    Call ModCell(SearchCell(60, "n@@N", "g"), "4")
    Call ModCell(SearchCell(60, "n@@N", "["), "4")
    Call ModCell(SearchCell(60, "Rû@Y¾Y", "[XNo"), "790")
    Call ModCell(SearchCell(60, "Rû@Y¾Y", "g"), "4")
    Call ModCell(SearchCell(60, "Rû@Y¾Y", "["), "5")
    Call ModCell(SearchCell(60, "O¿@¼a", "[XNo"), "790")
    Call ModCell(SearchCell(60, "O¿@¼a", "g"), "4")
    Call ModCell(SearchCell(60, "O¿@¼a", "["), "8")
    Call ModCell(SearchCell(60, "ºì@Fã", "[XNo"), "790")
    Call ModCell(SearchCell(60, "ºì@Fã", "g"), "4")
    Call ModCell(SearchCell(60, "ºì@Fã", "["), "9")
    
    Call ModCell(SearchCell(60, "@_Ê", "["), "6")
    Call ModCell(SearchCell(60, "{Y@²¿", "["), "7")

   ' 60 5g(80)
    Call ModCell(SearchCell(60, "y®@a¶", "[XNo"), "791")
    Call ModCell(SearchCell(60, "y®@a¶", "g"), "5")
    Call ModCell(SearchCell(60, "y®@a¶", "["), "4")
    Call ModCell(SearchCell(60, "çJ@qs", "[XNo"), "791")
    Call ModCell(SearchCell(60, "çJ@qs", "g"), "5")
    Call ModCell(SearchCell(60, "çJ@qs", "["), "5")
    Call ModCell(SearchCell(60, "R@Ã²", "[XNo"), "791")
    Call ModCell(SearchCell(60, "R@Ã²", "g"), "5")
    Call ModCell(SearchCell(60, "R@Ã²", "["), "6")
    Call ModCell(SearchCell(60, "´Ü@vi", "[XNo"), "791")
    Call ModCell(SearchCell(60, "´Ü@vi", "g"), "5")
    Call ModCell(SearchCell(60, "´Ü@vi", "["), "7")
    Call ModCell(SearchCell(60, "ã´@Dm", "[XNo"), "791")
    Call ModCell(SearchCell(60, "ã´@Dm", "g"), "5")
    Call ModCell(SearchCell(60, "ã´@Dm", "["), "8")

    ' 61
    Call ModCell(SearchCell(61, "¡ª@Çµ", "["), "4")
    Call ModCell(SearchCell(61, "c@ç", "["), "5")
    Call ModCell(SearchCell(61, "Oc@¬ü", "[XNo"), "800")
    Call ModCell(SearchCell(61, "Oc@¬ü", "["), "7")
    Call ModCell(SearchCell(61, "Oc@¬ü", "g"), "1")

    ' 63
    Call ModCell(SearchCell(63, "|º@¹q", "["), "5")
    Call ModCell(SearchCell(63, "Î@~q", "["), "6")

    ' 66
    Call ModCell(SearchCell(66, "Âz@Gî", "["), "5")
    Call ModCell(SearchCell(66, "{Y@²¿", "["), "6")
    Call ModCell(SearchCell(66, "Rû@Y¾Y", "["), "7")
    Call ModCell(SearchCell(66, "éØ@C½", "["), "8")
    Call ModCell(SearchCell(66, "ÄàV@«", "["), "9")

    ' 67
    Call ModCell("G444", "´a}DÄbqDOãü²}DRûFq")
    Call ModCell("G445", "ªa^qDìãqDÎ~qDåºåq")

    ' 68
    Call ModCell("F446", "5")
    Call ModCell("G446", "åàÊTDîûDËÞDÈ´üØDªªæ")
    Call ModCell("F447", "9")
    Call ModCell("G447", "c·ÊÞD´GÞDFJMÞDRcÊÄ")
    Call ModCell("F448", "6")
    Call ModCell("G448", "`DCDRÉDrìüìDcö¶")
    Call ModCell("F449", "8")
    Call ModCell("G449", "ìºüFDûÊ¹D´©¢èDâgàà")
    Call ModCell("F450", "7")
    Call ModCell("G450", "¬ÑüD¬ÑÑØDìJatD|ºÄè")
    Call ModCell("F451", "4")
    
    ' 69
    Call ModCell("G452", "Rè¨DÔûP¬÷DºcüçD{F¢")
    Call ModCell("G453", "´­éÝDìcDOY^¿DéØúaq")
    Call ModCell("G454", "éØØDsºäìTDÂRFÔDûü¼Ñí")
    
    ' 70
    Call ModCell("G455", "{ì_iD´ÄDéØ¡PDpcS÷")
    Call ModCell("G456", "ã´DmDO¿¼aDºìFãDéØC½")
    
    ' 71
    Call ModCell("F457", "8")
    Call ModCell("G457", "¬àVDãDXoWODèXzlDÖåãÄ")
    Call ModCell("F458", "7")
    Call ModCell("G458", "Jûc^Dnç³sñ¹Dez¾NDV{Cá")
    Call ModCell("F459", "9")
    Call ModCell("G459", "|ÑÐíDRÝ¼lDûü´\åD¼YC¹")
    Call ModCell("F460", "6")
    Call ModCell("G460", "ã¡]ÈDFzp¾D¨GîDâ£l")
    Call ModCell("F461", "5")
    Call ModCell("G461", "ú±IMDàV¾YDcDlD·K½")
    Call ModCell("F462", "4")
    Call ModCell("G462", "¼c[lDì{GlDìåD·ö\êY")
    
    ' 72
    Call ModCell("G463", "âr¾Dx]aóDõDªó")
    Call ModCell("G464", "êc÷D´ãÄ¾DéØåãÄDÍäN¾Y")
    Call ModCell("G465", "ìtãÄDOcTj[Dc´CõDìç²Ië")
    Call ModCell("G466", "²ìãÄÆDéØ¹îDñ{Ø§DcDMå")
    Call ModCell("G467", "¼{ãåD_RDRà­ÕDéØöå")
    Call ModCell("G468", "º£gPDìºåDìAD²ìãÄå")
    Call ModCell("G469", "ã¡ºDrØ²íDnÓC¹D²vÔT")

    Call EventChange(True)
    
    ' ubNÌÛ¶
    ActiveWorkbook.Save
End Sub

'
' ßa³NxÌvOÌL^ðüêé
'
Sub s¯L^üÍ()
    Sheets("L^æÊ").Select
    Call SetRace(1, 1)
    Call SetLean(1, 3, "34709")
    Call SetLean(2, 4, "25977")
    Call SetLean(3, 6, "")
    Call SetLean(4, 8, "30973")
    Call o^
    Call Êè
    Call ú»
    
    ' ubNÌÛ¶
    ActiveWorkbook.Save
End Sub
