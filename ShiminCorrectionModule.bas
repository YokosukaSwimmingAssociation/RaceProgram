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
    Call ModCell("F9", "5")
    Call ModCell("G9", "éØDÔûP¬÷DRè¨Dºcüç")
    Call ModCell("C10", "2")
    Call ModCell("F10", "8")
    Call ModCell("G10", "Ö¡zqD{è¾qD{Yß®ÝDéØcq")
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
    Call ModCell("C22", "5")
    Call ModCell("F22", "8")
    Call ModCell("G22", "éØ¡PDpcS÷D´ÄD{ì_i")
    Call ModCell("C23", "5")
    Call ModCell("F23", "9")
    Call ModCell("G23", "éØC½DÂzGîDO¿¼aDã´Dm")
    
    ' 7|9@¯ê[X
    Call ModCell(SearchCell(7, "ªª@æ", "["), "3")
    Call ModCell(SearchCell(7, "J{@Äq", "["), "4")
    Call ModCell(SearchCell(7, "ìº@üF", "["), "5")
    Call ModCell(SearchCell(7, "|º@Äè", "["), "6")
    Call ModCell(SearchCell(7, "c@ö¶", "["), "7")
    Call ModCell(SearchCell(9, "J@K]", "[XNo"), "7")
    Call ModCell(SearchCell(9, "J@K]", "["), "8")
    
    ' 13
    Call ModCell(SearchCell(13, "È´@üØ", "["), "8")
    Call ModCell(SearchCell(13, "}´@÷", "["), "5")
    Call ModCell(SearchCell(13, "åà@ÊT", "["), "4")

    ' 17
    Call ModCell(SearchCell(17, "c@DMå", "["), "3")
    Call ModCell(SearchCell(17, "¬àV@q¾", "["), "4")
    Call ModCell(SearchCell(17, "Rà@­Õ", "["), "5")
    Call ModCell(SearchCell(17, "½c@¼å", "["), "6")
    Call ModCell(SearchCell(17, "õc@Ðá", "["), "7")

    ' 18
    Call ModCell(SearchCell(18, "âä@¾¢", "[XNo"), "14")
    Call ModCell(SearchCell(18, "âä@¾¢", "["), "9")

    ' 19
    Call ModCell(SearchCell(17, "Îä@üDí", "["), "4")
    Call ModCell(SearchCell(17, "~ø@çl", "["), "5")
    Call ModCell(SearchCell(17, "¯i@å", "["), "6")

    ' 20
    Call ModCell(SearchCell(20, "éØ@", "[XNo"), "16")
    Call ModCell(SearchCell(20, "éØ@", "["), "8")

    ' 34
    Call ModCell(SearchCell(34, "|R@C¬", "["), "5")
    Call ModCell(SearchCell(34, "É¡@àæ", "[XNo"), "25")
    Call ModCell(SearchCell(34, "É¡@àæ", "g"), "1")
    Call ModCell(SearchCell(34, "É¡@àæ", "["), "6")
    Call ModCell(SearchCell(34, "Oc@xl", "["), "7")
    Call ModCell(SearchCell(34, "|Ñ@Ðí", "["), "8")

    ' 37
    Call ModCell(SearchCell(37, "ìº@üF", "[XNo"), "30")
    Call ModCell(SearchCell(37, "ìº@üF", "g"), "1")
    Call ModCell(SearchCell(37, "ìº@üF", "["), "6")
    Call ModCell(SearchCell(37, "R@@É", "[XNo"), "30")
    Call ModCell(SearchCell(37, "R@@É", "g"), "1")
    Call ModCell(SearchCell(37, "R@@É", "["), "8")
    Call ModCell(SearchCell(37, "Oc@¬ü", "["), "4")
    Call ModCell(SearchCell(37, "FJ@MÞ", "["), "5")
    Call ModCell(SearchCell(37, "û@DC", "["), "6")
    Call ModCell(SearchCell(37, "¬Ñ@ÑÞ", "["), "7")
    Call ModCell(SearchCell(37, "¬Ñ@ÑØ", "["), "8")
    Call ModCell(SearchCell(37, "|©@^D", "[XNo"), "31")
    Call ModCell(SearchCell(37, "|©@^D", "g"), "2")
    Call ModCell(SearchCell(37, "|©@^D", "["), "9")

    ' 38-39
    Call ModCell(SearchCell(38, "éØ@úaq", "["), "4")
    Call ModCell(SearchCell(38, "¬ì@àYÞ", "["), "5")
    Call ModCell(SearchCell(38, "´@­éÝ", "["), "6")
    Call ModCell(SearchCell(38, "½c@©Ì", "["), "7")
    Call ModCell(SearchCell(39, "éØ@cq", "[XNo"), "32")
    Call ModCell(SearchCell(39, "éØ@cq", "["), "9")

    ' 40
    Call ModCell(SearchCell(40, "c@ëM", "["), "5")
    Call ModCell(SearchCell(40, "V{@Cá", "[XNo"), "34")
    Call ModCell(SearchCell(40, "V{@Cá", "g"), "1")
    Call ModCell(SearchCell(40, "V{@Cá", "["), "6")
    Call ModCell(SearchCell(40, "çJ@Am", "["), "7")
    Call ModCell(SearchCell(40, "nç²@sñ¹", "[XNo"), "34")
    Call ModCell(SearchCell(40, "nç²@sñ¹", "g"), "1")
    Call ModCell(SearchCell(40, "nç²@sñ¹", "["), "8")
    Call ModCell(SearchCell(40, "c@³ä", "[XNo"), "35")
    Call ModCell(SearchCell(40, "c@³ä", "g"), "2")
    Call ModCell(SearchCell(40, "c@³ä", "["), "3")
    Call ModCell(SearchCell(40, "|R@C¬", "["), "4")
    Call ModCell(SearchCell(40, "ú±@IM", "["), "5")
    Call ModCell(SearchCell(40, "ì{@Gl", "["), "6")
    Call ModCell(SearchCell(40, "c@Dl", "["), "7")
    Call ModCell(SearchCell(40, "xmì@l¾", "["), "8")
    Call ModCell(SearchCell(40, "©@DVî", "[XNo"), "35")
    Call ModCell(SearchCell(40, "©@DVî", "g"), "2")
    Call ModCell(SearchCell(40, "©@DVî", "["), "9")
    Call ModCell(SearchCell(40, "ì@å", "[XNo"), "36")
    Call ModCell(SearchCell(40, "ì@å", "g"), "3")
    Call ModCell(SearchCell(40, "ì@å", "["), "3")
    Call ModCell(SearchCell(40, "·ö@\êY", "["), "4")
    Call ModCell(SearchCell(40, "kì@DãÄ", "["), "5")
    Call ModCell(SearchCell(40, "¼Y@C¹", "["), "6")
    Call ModCell(SearchCell(40, "Áä@m¾N", "["), "7")
    Call ModCell(SearchCell(40, "ì@Nî", "["), "8")
    Call ModCell(SearchCell(40, "¬àV@ãÄ¾", "[XNo"), "36")
    Call ModCell(SearchCell(40, "¬àV@ãÄ¾", "g"), "3")
    Call ModCell(SearchCell(40, "¬àV@ãÄ¾", "["), "9")
    
    
    Call ModCell(SearchCell(41, "éØ@åM", "[XNo"), "37")
    Call ModCell(SearchCell(41, "éØ@åM", "g"), "1")
    Call ModCell(SearchCell(41, "éØ@åM", "["), "5")
    Call ModCell(SearchCell(41, "c´@Cõ", "["), "6")
    Call ModCell(SearchCell(41, "ìç²@IÍ", "[XNo"), "37")
    Call ModCell(SearchCell(41, "ìç²@IÍ", "g"), "1")
    Call ModCell(SearchCell(41, "ìç²@IÍ", "["), "7")
    Call ModCell(SearchCell(41, "üû@im", "[XNo"), "37")
    Call ModCell(SearchCell(41, "üû@im", "g"), "1")
    Call ModCell(SearchCell(41, "üû@im", "["), "8")
    
    Call ModCell(SearchCell(41, "Xì@^á", "[XNo"), "38")
    Call ModCell(SearchCell(41, "Xì@^á", "g"), "2")
    Call ModCell(SearchCell(41, "Xì@^á", "["), "4")
    Call ModCell(SearchCell(41, "²¡@üu", "["), "5")
    Call ModCell(SearchCell(41, "²vÔ@T", "["), "6")
    Call ModCell(SearchCell(41, "ì@@A", "["), "7")
    Call ModCell(SearchCell(41, "ïº@qó", "[XNo"), "38")
    Call ModCell(SearchCell(41, "ïº@qó", "g"), "2")
    Call ModCell(SearchCell(41, "ïº@qó", "["), "8")
    
    Call ModCell(SearchCell(41, "}L@êó", "[XNo"), "39")
    Call ModCell(SearchCell(41, "}L@êó", "g"), "3")
    Call ModCell(SearchCell(41, "}L@êó", "["), "3")
    Call ModCell(SearchCell(41, "ª@@ó", "[XNo"), "39")
    Call ModCell(SearchCell(41, "ª@@ó", "g"), "3")
    Call ModCell(SearchCell(41, "ª@@ó", "["), "4")
    Call ModCell(SearchCell(41, "ìº@åÛ", "[XNo"), "39")
    Call ModCell(SearchCell(41, "ìº@åÛ", "g"), "3")
    Call ModCell(SearchCell(41, "ìº@åÛ", "["), "5")
    Call ModCell(SearchCell(41, "ã¡@º", "["), "6")
    Call ModCell(SearchCell(41, "éØ@¹î", "["), "7")
    Call ModCell(SearchCell(41, "OJ@åê", "["), "8")
    Call ModCell(SearchCell(41, "x]@aó", "["), "9")
    
    Call ModCell(SearchCell(41, "º£@gP", "[XNo"), "40")
    Call ModCell(SearchCell(41, "º£@gP", "g"), "4")
    Call ModCell(SearchCell(41, "º£@gP", "["), "3")
    Call ModCell(SearchCell(41, "c@DMå", "["), "4")
    Call ModCell(SearchCell(41, "²ì@ãÄÆ", "["), "8")

    Call ModCell(SearchCell(43, "ìû@ÄC", "[XNo"), "41")
    Call ModCell(SearchCell(43, "ìû@ÄC", "g"), "1")
    Call ModCell(SearchCell(43, "ìû@ÄC", "["), "6")
    Call ModCell(SearchCell(43, "´@GÞ", "["), "7")
    Call ModCell(SearchCell(43, "éØ@ìü", "["), "8")

    Call ModCell(SearchCell(45, "Oã@ü²}", "["), "3")
    Call ModCell(SearchCell(45, "ìã@q", "["), "4")
    Call ModCell(SearchCell(45, "Ä@@bq", "["), "5")
    Call ModCell(SearchCell(45, "´@a}", "["), "6")
    Call ModCell(SearchCell(45, "éØ@¤q", "["), "7")
    Call ModCell(SearchCell(45, "q¡@R¢q", "["), "8")

    Call ModCell(SearchCell(48, "ªc@@²", "["), "4")
    Call ModCell(SearchCell(48, "", "["), "5")
    Call ModCell(SearchCell(48, "", "["), "6")
    Call ModCell(SearchCell(48, "", "["), "7")
    Call ModCell(SearchCell(48, "", "["), "8")
    Call ModCell(SearchCell(48, "", "["), "9")


    Call ModCell(SearchCell(48, "", "[XNo"), "48")
    Call ModCell(SearchCell(48, "", "g"), "1")
    Call ModCell(SearchCell(48, "", "["), "3")

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
