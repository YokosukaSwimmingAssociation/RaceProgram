Attribute VB_Name = "ShiminCorrectionModule"
'
' �s�����̃v���O�����␳
'
Sub �s���v���O�����␳()

    Call EventChange(False)

    Sheets("�G���g���[�ꗗ").Activate
    ' 1 ���w���q4�~50M���h���[�����[
    Call ModCell("F2", "8")
    Call ModCell("G2", "�Έ���D��D����D�˓ށD�Ȍ����؁D����ʔT")
    Call ModCell("F3", "9")
    Call ModCell("G3", "�����G���ށD�R�c�ʉāD�c���ʓށD�F�J�M��")
    Call ModCell("F4", "6")
    Call ModCell("G4", "�쑺���F�D�����Ǔ��D����������D�����ʉ�")
    Call ModCell("F5", "7")
    Call ModCell("G5", "�r�����D���c�����D���`�D�C�D���R��")
    Call ModCell("F6", "5")
    Call ModCell("G6", "����J�a�t�D���є����D�|���ĉ�D���юэ�")
    Call ModCell("F7", "4")
    Call ModCell("G8", "��ؓ��a�q�D�O�Y�^���D��������݁D����c")
    
    ' 2�|3�@���ꃌ�[�X
    Call ModCell("F9", "5")
    Call ModCell("G9", "��ؗ����D���P�����D���R�肨�D���c����")
    Call ModCell("C10", "20")
    Call ModCell("F10", "8")
    Call ModCell("G10", "�֓��z�q�D�{�薾�q�D�{�Y�߂��݁D��،c�q")
    
    ' 4 ���w �j�q 4�~50M ���h���[�����[
    Call ModCell("F11", "8")
    Call ModCell("G11", "���R�E���D���V�D��D�����T��D�X�o�W�O")
    Call ModCell("F12", "7")
    Call ModCell("G12", "�e�z���N�D�͖�����D�J���c�^�D�V�{�C��")
    Call ModCell("F13", "6")
    Call ModCell("G13", "�F�z�p���D�㓡�]�ȁD�����G��D�␣�l")
    Call ModCell("F14", "9")
    Call ModCell("G14", "�����u�D���쐰��D����m���N�D�R�����Y")
    Call ModCell("F15", "5")
    Call ModCell("G15", "�y�m��l���D�����K���D�c���D�l�D�����I�M")
    Call ModCell("F16", "4")
    Call ModCell("G16", "�v��C�ځD���c�[�l�D��������D�����\��Y")
    
    ' 5�|6�@���ꃌ�[�X
    Call ModCell("F17", "3")
    Call ModCell("G17", "������D���n���D���r���D������")
    Call ModCell("F18", "4")
    Call ModCell("G18", "�m�J�S�l�D�����čƁD��{�؏��D�c���D�M��")
    Call ModCell("F19", "5")
    Call ModCell("G19", "�O�c���T�j�[�D�͏�剹�D�c���C���D��ؑ�M")
    Call ModCell("F20", "6")
    Call ModCell("G20", "�����g�P�D�ߍ]�D���D�����đ�D����A")
    Call ModCell("F21", "7")
    Call ModCell("G21", "��،���D�n�ӗ��D�_�R�����D���{���")
    ' 6
    Call ModCell("C22", "50")
    Call ModCell("F22", "8")
    Call ModCell("G22", "��؎��P�D�p�c�S���D�����āD�{��_�i")
    Call ModCell("C23", "50")
    Call ModCell("F23", "9")
    Call ModCell("G23", "��؏C���D�z�G��D�O�����a�D�㌴�D�m")
    
    ' 7�|9�@���ꃌ�[�X
    Call ModCell(SearchCell(7, "�����@�捁", "���[��"), "3")
    Call ModCell(SearchCell(7, "�J�{�@�Ďq", "���[��"), "4")
    Call ModCell(SearchCell(7, "�쑺�@���F", "���[��"), "5")
    Call ModCell(SearchCell(7, "�|���@�ĉ�", "���[��"), "6")
    Call ModCell(SearchCell(7, "���c�@����", "���[��"), "7")
    ' 9
    Call ModCell(SearchCell(9, "���J�@�K�]", "���[�XNo"), "70")
    Call ModCell(SearchCell(9, "���J�@�K�]", "���[��"), "9")
    
    ' 17-18 ���ꃌ�[�X
    Call ModCell(SearchCell(17, "�c���@�D�M��", "���[��"), "3")
    Call ModCell(SearchCell(17, "���V�@�q��", "���[��"), "4")
    Call ModCell(SearchCell(17, "�R���@����", "���[��"), "5")
    Call ModCell(SearchCell(17, "���c�@�Ќ�", "���[��"), "6")
    Call ModCell(SearchCell(17, "���c�@����", "���[��"), "7")
    ' 18
    Call ModCell(SearchCell(18, "���@����", "���[�XNo"), "140")
    Call ModCell(SearchCell(18, "���@����", "���[��"), "9")

    ' 19-20 ���ꃌ�[�X
    Call ModCell(SearchCell(19, "�Έ�@���D��", "���[��"), "4")
    Call ModCell(SearchCell(19, "�~���@��l", "���[��"), "5")
    Call ModCell(SearchCell(19, "���i�@�劒", "���[��"), "6")
    ' 20
    Call ModCell(SearchCell(20, "��؁@����", "���[�XNo"), "160")
    Call ModCell(SearchCell(20, "��؁@����", "���[��"), "8")

    ' 38-39 ���ꃌ�[�X
    Call ModCell(SearchCell(38, "��؁@���a�q", "���[��"), "4")
    Call ModCell(SearchCell(38, "����@�Y��", "���[��"), "5")
    Call ModCell(SearchCell(38, "�����@�����", "���[��"), "6")
    Call ModCell(SearchCell(38, "���c�@����", "���[��"), "7")
    ' 39
    Call ModCell(SearchCell(39, "��؁@�c�q", "���[�XNo"), "320")
    Call ModCell(SearchCell(39, "��؁@�c�q", "���[��"), "9")

    ' 50-51 ���ꃌ�[�X
    Call ModCell(SearchCell(50, "���c�@�z��", "���[��"), "3")
    ' 51
    Call ModCell(SearchCell(51, "���J�@�K�]", "���[�XNo"), "510")
    Call ModCell(SearchCell(51, "���J�@�K�]", "���[��"), "6")
    Call ModCell(SearchCell(51, "�{�Y�@�߂���", "���[�XNo"), "510")
    Call ModCell(SearchCell(51, "�{�Y�@�߂���", "���[��"), "7")
    Call ModCell(SearchCell(51, "����@�؎q", "���[�XNo"), "510")
    Call ModCell(SearchCell(51, "����@�؎q", "���[��"), "8")
    Call ModCell(SearchCell(51, "�֓��@�z�q", "���[�XNo"), "510")
    Call ModCell(SearchCell(51, "�֓��@�z�q", "���[��"), "9")

   ' 52
    Call ModCell(SearchCell(52, "�����@�T��", "���[��"), "4")
    Call ModCell(SearchCell(52, "�␣�@�@�l", "���[�XNo"), "540")
    Call ModCell(SearchCell(52, "�␣�@�@�l", "�g"), "2")
    Call ModCell(SearchCell(52, "�␣�@�@�l", "���[��"), "7")
    Call ModCell(SearchCell(52, "�R�݁@���l", "���[��"), "8")

   ' 54
    Call ModCell(SearchCell(54, "���с@�c�Y", "���[��"), "4")
    Call ModCell(SearchCell(54, "�ۖ΁@����", "���[��"), "5")
    Call ModCell(SearchCell(54, "�����@�r��", "���[�XNo"), "570")
    Call ModCell(SearchCell(54, "�����@�r��", "�g"), "1")
    Call ModCell(SearchCell(54, "�����@�r��", "���[��"), "7")

   ' 60 1�g(76)
    Call ModCell(SearchCell(60, "�R�{�@���l", "���[��"), "3")
    Call ModCell(SearchCell(60, "���с@�c�Y", "���[��"), "4")
    Call ModCell(SearchCell(60, "�O�Y�@�B��", "���[��"), "5")
    Call ModCell(SearchCell(60, "�����@�M�m", "���[��"), "7")
    Call ModCell(SearchCell(60, "�؁@����", "���[��"), "8")
    Call ModCell(SearchCell(60, "�ߓc�@�O�j", "���[��"), "9")
    Call ModCell(SearchCell(60, "�n�Ӂ@���j", "���[�XNo"), "760")
    Call ModCell(SearchCell(60, "�n�Ӂ@���j", "�g"), "1")
    Call ModCell(SearchCell(60, "�n�Ӂ@���j", "���[��"), "6")

   ' 60 2�g(77)
    Call ModCell(SearchCell(60, "�ۖ΁@����", "���[��"), "5")
    Call ModCell(SearchCell(60, "�c��@�@�G", "���[��"), "6")
    Call ModCell(SearchCell(60, "�x��@���Y", "���[��"), "7")

   ' 60 3�g(78)
    Call ModCell(SearchCell(60, "�����@�p�V", "���[�XNo"), "780")
    Call ModCell(SearchCell(60, "�����@�p�V", "�g"), "3")
    Call ModCell(SearchCell(60, "�����@�p�V", "���[��"), "4")
    Call ModCell(SearchCell(60, "�쑺�@��_", "���[�XNo"), "780")
    Call ModCell(SearchCell(60, "�쑺�@��_", "�g"), "3")
    Call ModCell(SearchCell(60, "�쑺�@��_", "���[��"), "5")
    Call ModCell(SearchCell(60, "�����@�ǕF", "���[�XNo"), "780")
    Call ModCell(SearchCell(60, "�����@�ǕF", "�g"), "3")
    Call ModCell(SearchCell(60, "�����@�ǕF", "���[��"), "8")
    
    Call ModCell(SearchCell(60, "�����@�r��", "�g"), "3")
    Call ModCell(SearchCell(60, "�����@�r��", "���[��"), "6")
    Call ModCell(SearchCell(60, "�{��@�_�i", "�g"), "3")
    Call ModCell(SearchCell(60, "�{��@�_�i", "���[��"), "7")
    Call ModCell(SearchCell(60, "�p�c�@�S��", "�g"), "3")
    Call ModCell(SearchCell(60, "�p�c�@�S��", "���[��"), "9")

   ' 60 4�g(79)
    Call ModCell(SearchCell(60, "���n�@�@�N", "���[�XNo"), "790")
    Call ModCell(SearchCell(60, "���n�@�@�N", "�g"), "4")
    Call ModCell(SearchCell(60, "���n�@�@�N", "���[��"), "4")
    Call ModCell(SearchCell(60, "�R���@�Y���Y", "���[�XNo"), "790")
    Call ModCell(SearchCell(60, "�R���@�Y���Y", "�g"), "4")
    Call ModCell(SearchCell(60, "�R���@�Y���Y", "���[��"), "5")
    Call ModCell(SearchCell(60, "�O���@���a", "���[�XNo"), "790")
    Call ModCell(SearchCell(60, "�O���@���a", "�g"), "4")
    Call ModCell(SearchCell(60, "�O���@���a", "���[��"), "8")
    Call ModCell(SearchCell(60, "����@�F��", "���[�XNo"), "790")
    Call ModCell(SearchCell(60, "����@�F��", "�g"), "4")
    Call ModCell(SearchCell(60, "����@�F��", "���[��"), "9")
    
    Call ModCell(SearchCell(60, "�����@�_��", "���[��"), "6")
    Call ModCell(SearchCell(60, "�{�Y�@����", "���[��"), "7")

   ' 60 5�g(80)
    Call ModCell(SearchCell(60, "�y���@�a��", "���[�XNo"), "791")
    Call ModCell(SearchCell(60, "�y���@�a��", "�g"), "5")
    Call ModCell(SearchCell(60, "�y���@�a��", "���[��"), "4")
    Call ModCell(SearchCell(60, "��J�@�q�s", "���[�XNo"), "791")
    Call ModCell(SearchCell(60, "��J�@�q�s", "�g"), "5")
    Call ModCell(SearchCell(60, "��J�@�q�s", "���[��"), "5")
    Call ModCell(SearchCell(60, "���R�@�×�", "���[�XNo"), "791")
    Call ModCell(SearchCell(60, "���R�@�×�", "�g"), "5")
    Call ModCell(SearchCell(60, "���R�@�×�", "���[��"), "6")
    Call ModCell(SearchCell(60, "���܁@�v�i", "���[�XNo"), "791")
    Call ModCell(SearchCell(60, "���܁@�v�i", "�g"), "5")
    Call ModCell(SearchCell(60, "���܁@�v�i", "���[��"), "7")
    Call ModCell(SearchCell(60, "�㌴�@�D�m", "���[�XNo"), "791")
    Call ModCell(SearchCell(60, "�㌴�@�D�m", "�g"), "5")
    Call ModCell(SearchCell(60, "�㌴�@�D�m", "���[��"), "8")

    ' 61
    Call ModCell(SearchCell(61, "�����@�Ǔ�", "���[��"), "4")
    Call ModCell(SearchCell(61, "�c���@�琐", "���[��"), "5")
    Call ModCell(SearchCell(61, "�O�c�@����", "���[�XNo"), "800")
    Call ModCell(SearchCell(61, "�O�c�@����", "���[��"), "7")
    Call ModCell(SearchCell(61, "�O�c�@����", "�g"), "1")

    ' 63
    Call ModCell(SearchCell(63, "�|���@���q", "���[��"), "5")
    Call ModCell(SearchCell(63, "���΁@�~�q", "���[��"), "6")

    ' 66
    Call ModCell(SearchCell(66, "�z�@�G��", "���[��"), "5")
    Call ModCell(SearchCell(66, "�{�Y�@����", "���[��"), "6")
    Call ModCell(SearchCell(66, "�R���@�Y���Y", "���[��"), "7")
    Call ModCell(SearchCell(66, "��؁@�C��", "���[��"), "8")
    Call ModCell(SearchCell(66, "���V�@����", "���[��"), "9")

    ' 67
    Call ModCell("G444", "�����a�}�D�Čb�q�D�O������}�D�R���F�q")
    Call ModCell("G445", "���a�^���q�D��㋞�q�D���Ώ~�q�D�呺��q")

    ' 68
    Call ModCell("F446", "5")
    Call ModCell("G446", "����ʔT�D����D�˓ށD�Ȍ����؁D�����捁")
    Call ModCell("F447", "9")
    Call ModCell("G447", "�c���ʓށD�����G���ށD�F�J�M�ށD�R�c�ʉ�")
    Call ModCell("F448", "6")
    Call ModCell("G448", "���`�D�C�D���R�ɁD�r�����D���c����")
    Call ModCell("F449", "8")
    Call ModCell("G449", "�쑺���F�D�����ʉ��D����������D��g����")
    Call ModCell("F450", "7")
    Call ModCell("G450", "���є����D���юэ؁D����J�a�t�D�|���ĉ�")
    Call ModCell("F451", "4")
    
    ' 69
    Call ModCell("G452", "���R�肨�D���P�����D���c����D���{�F��")
    Call ModCell("G453", "��������݁D����c�D�O�Y�^���D��ؓ��a�q")
    Call ModCell("G454", "��؍؏��D�s�����T�D�R�F�ԁD�����і�")
    
    ' 70
    Call ModCell("G455", "�{��_�i�D�����āD��؎��P�D�p�c�S��")
    Call ModCell("G456", "�㌴�D�m�D�O�����a�D����F��D��؏C��")
    
    ' 71
    Call ModCell("F457", "8")
    Call ModCell("G457", "���V�D��D�X�o�W�O�D��X�z�l�D�֑���")
    Call ModCell("F458", "7")
    Call ModCell("G458", "�J���c�^�D�n糕s�񓹁D�e�z���N�D�V�{�C��")
    Call ModCell("F459", "9")
    Call ModCell("G459", "�|�јЖ�D�R�ݒ��l�D�����\��D���Y�C��")
    Call ModCell("F460", "6")
    Call ModCell("G460", "�㓡�]�ȁD�F�z�p���D�����G��D�␣�l")
    Call ModCell("F461", "5")
    Call ModCell("G461", "�����I�M�D���V�����Y�D�c���D�l�D�����K��")
    Call ModCell("F462", "4")
    Call ModCell("G462", "���c�[�l�D��{�G�l�D���슐��D�����\��Y")
    
    ' 72
    Call ModCell("G463", "����r���D�x�]�a��D�������D������")
    Call ModCell("G464", "��c���D�����đ��D��ؑ��āD�͈�N���Y")
    Call ModCell("G465", "����t�āD�O�c���T�j�[�D�c���C���D��粗I��")
    Call ModCell("G466", "�����čƁD��ؔ���D��{�؏��D�c���D�M��")
    Call ModCell("G467", "���{���D�_�R�����D�R�����ՁD��،���")
    Call ModCell("G468", "�����g�P�D�쑺��D����A�D�����đ�")
    Call ModCell("G469", "�㓡�����D�r�ؗ���D�n�ӊC���D���v�ԐT")

    Call EventChange(True)
    
    ' �u�b�N�̕ۑ�
    ActiveWorkbook.Save
End Sub

'
' �ߘa���N�x�̃v���O�����̋L�^������
'
Sub �s���L�^����()
    Sheets("�L�^���").Select
    Call SetRace(1, 1)
    Call SetLean(2, 4, "")
    Call SetLean(3, 5, "24938")
    Call SetLean(4, 6, "22309")
    Call SetLean(5, 7, "22423")
    Call SetLean(6, 8, "25897")
    Call SetLean(7, 9, "23297")
    Call �o�^
    Call ���ʌ���
    Call ������
    
    Call SetRace(2, 1)
    Call SetLean(3, 5, "22599")
    Call SetLean(4, 6, "23540")
    Call SetLean(6, 8, "25072")
    Call �o�^
    Call ���ʌ���
    Call ������
    
    Call SetRace(4, 1)
    Call SetLean(2, 4, "22424")
    Call SetLean(3, 5, "23780")
    Call SetLean(4, 6, "21010")
    Call SetLean(5, 7, "21811")
    Call SetLean(6, 8, "22916")
    Call SetLean(7, 9, "")
    Call �o�^
    Call ���ʌ���
    Call ������
    
    Call SetRace(5, 1)
    Call SetLean(1, 3, "22365")
    Call SetLean(2, 4, "21235")
    Call SetLean(3, 5, "20719")
    Call SetLean(4, 6, "21631")
    Call SetLean(5, 7, "20928")
    Call SetLean(6, 8, "21169")
    Call SetLean(7, 9, "24807")
    Call �o�^
    Call ���ʌ���
    Call ������
    
    Call SetRace(7, 1)
    Call SetLean(1, 3, "34435")
    Call SetLean(2, 4, "")
    Call SetLean(3, 5, "24062")
    Call SetLean(4, 6, "30724")
    Call SetLean(5, 7, "30434")
    Call SetLean(7, 9, "35667")
    Call �o�^
    Call ���ʌ���
    Call ������
    
    Call SetRace(10, 1)
    Call SetLean(3, 5, "31305")
    Call SetLean(4, 6, "25735")
    Call SetLean(5, 7, "30160")
    Call SetLean(6, 8, "32499")
    Call �o�^
    Call ���ʌ���
    Call ������
    
    Call SetRace(11, 1)
    Call SetLean(2, 4, "25004")
    Call SetLean(3, 5, "24581")
    Call SetLean(4, 6, "22740")
    Call SetLean(5, 7, "24042")
    Call SetLean(6, 8, "24881")
    Call �o�^
    Call ���ʌ���
    Call ������
    
    Call SetRace(13, 1)
    Call SetLean(1, 3, "23918")
    Call SetLean(2, 4, "30064")
    Call SetLean(3, 5, "24587")
    Call SetLean(4, 6, "22224")
    Call SetLean(5, 7, "23323")
    Call SetLean(6, 8, "31235")
    Call SetLean(7, 9, "31706")
    Call �o�^
    Call ���ʌ���
    Call ������
    
    Call SetRace(15, 1)
    Call SetLean(3, 5, "35575")
    Call �o�^
    Call ���ʌ���
    Call ������
    
    Call SetRace(16, 1)
    Call SetLean(3, 5, "25363")
    Call SetLean(4, 6, "22796")
    Call SetLean(5, 7, "30755")
    Call SetLean(6, 8, "30751")
    Call �o�^
    Call ���ʌ���
    Call ������
    
    Call SetRace(17, 1)
    Call SetLean(1, 3, "22062")
    Call SetLean(2, 4, "")
    Call SetLean(3, 5, "22898")
    Call SetLean(4, 6, "22915")
    Call SetLean(5, 7, "24717")
    Call SetLean(7, 9, "30434")
    Call �o�^
    Call ���ʌ���
    Call ������
    
    Call SetRace(19, 1)
    Call SetLean(2, 4, "20830")
    Call SetLean(3, 5, "11333")
    Call SetLean(4, 6, "13980")
    Call SetLean(6, 8, "11386")
    Call �o�^
    Call ���ʌ���
    Call ������
    
    Call SetRace(22, 1)
    Call SetLean(2, 4, "20111")
    Call SetLean(3, 5, "12057")
    Call SetLean(4, 6, "11994")
    Call SetLean(5, 7, "11670")
    Call SetLean(6, 8, "12784")
    Call SetLean(7, 9, "14289")
    Call �o�^
    Call ���ʌ���
    Call ������
    
    Call SetRace(23, 1)
    Call SetLean(2, 4, "12995")
    Call SetLean(3, 5, "11638")
    Call SetLean(4, 6, "11738")
    Call SetLean(5, 7, "12031")
    Call SetLean(6, 8, "11801")
    Call �o�^
    Call ���ʌ���
    Call ������
    
    Call SetRace(25, 1)
    Call SetLean(2, 4, "12640")
    Call SetLean(3, 5, "12018")
    Call SetLean(4, 6, "11138")
    Call SetLean(5, 7, "11324")
    Call SetLean(6, 8, "13674")
    Call �o�^
    Call ���ʌ���
    Call ������
    
    Call SetRace(28, 1)
    Call SetLean(2, 4, "14644")
    Call SetLean(3, 5, "10731")
    Call SetLean(4, 6, "10368")
    Call SetLean(5, 7, "10990")
    Call SetLean(6, 8, "11899")
    Call �o�^
    Call ���ʌ���
    Call ������
    
    Call SetRace(29, 1)
    Call SetLean(3, 5, "12089")
    Call SetLean(4, 6, "10305")
    Call SetLean(5, 7, "11099")
    Call SetLean(6, 8, "14712")
    Call �o�^
    Call ���ʌ���
    Call ������
    
    Call SetRace(31, 1)
    Call SetLean(1, 3, "13799")
    Call SetLean(2, 4, "13871")
    Call SetLean(3, 5, "13620")
    Call SetLean(4, 6, "12619")
    Call SetLean(5, 7, "")
    Call SetLean(6, 8, "14078")
    Call SetLean(7, 9, "14068")
    Call �o�^
    Call ���ʌ���
    Call ������
    
    Call SetRace(32, 1)
    Call SetLean(3, 5, "13605")
    Call SetLean(4, 6, "12586")
    Call SetLean(5, 7, "13312")
    Call SetLean(6, 8, "13401")
    Call �o�^
    Call ���ʌ���
    Call ������
    
    Call SetRace(34, 1)
    Call SetLean(3, 5, "14140")
    Call SetLean(4, 6, "13239")
    Call SetLean(5, 7, "14228")
    Call SetLean(6, 8, "13390")
    Call �o�^
    Call ������
    
    Call SetRace(34, 2)
    Call SetLean(2, 4, "13036")
    Call SetLean(3, 5, "12716")
    Call SetLean(4, 6, "11561")
    Call SetLean(5, 7, "12083")
    Call SetLean(6, 8, "12306")
    Call SetLean(7, 9, "13349")
    Call �o�^
    Call ���ʌ���
    Call ������
    
    Call SetRace(35, 1)
    Call SetLean(2, 4, "12848")
    Call SetLean(3, 5, "12981")
    Call SetLean(4, 6, "13040")
    Call SetLean(5, 7, "13181")
    Call SetLean(6, 8, "")
    Call �o�^
    Call ������
    
    Call SetRace(35, 2)
    Call SetLean(1, 3, "12555")
    Call SetLean(2, 4, "12129")
    Call SetLean(3, 5, "12105")
    Call SetLean(4, 6, "11723")
    Call SetLean(5, 7, "12156")
    Call SetLean(6, 8, "12522")
    Call SetLean(7, 9, "13346")
    Call �o�^
    Call ���ʌ���
    Call ������
    
    Call SetRace(36, 1)
    Call SetLean(4, 6, "14475")
    Call �o�^
    Call ���ʌ���
    Call ������
    
    Call SetRace(37, 1)
    Call SetLean(3, 5, "14081")
    Call SetLean(4, 6, "10553")
    Call SetLean(5, 7, "12134")
    Call SetLean(6, 8, "11015")
    Call �o�^
    Call ������
    
    Call SetRace(37, 2)
    Call SetLean(2, 4, "12470")
    Call SetLean(3, 5, "11403")
    Call SetLean(4, 6, "10581")
    Call SetLean(5, 7, "10690")
    Call SetLean(6, 8, "12389")
    Call SetLean(7, 9, "")
    Call �o�^
    Call ���ʌ���
    Call ������
    
    Call SetRace(38, 1)
    Call SetLean(2, 4, "11752")
    Call SetLean(3, 5, "10969")
    Call SetLean(4, 6, "11287")
    Call SetLean(5, 7, "")
    Call SetLean(7, 9, "14828")
    Call �o�^
    Call ���ʌ���
    Call ������
    
    Call SetRace(40, 1)
    Call SetLean(3, 5, "")
    Call SetLean(4, 6, "10507")
    Call SetLean(5, 7, "12790")
    Call SetLean(6, 8, "11032")
    Call �o�^
    Call ������
    
    Call SetRace(40, 2)
    Call SetLean(1, 3, "12490")
    Call SetLean(2, 4, "12380")
    Call SetLean(3, 5, "12627")
    Call SetLean(4, 6, "11300")
    Call SetLean(5, 7, "11132")
    Call SetLean(6, 8, "12414")
    Call SetLean(7, 9, "11324")
    Call �o�^
    Call ������
    
    Call SetRace(40, 3)
    Call SetLean(1, 3, "11645")
    Call SetLean(2, 4, "10817")
    Call SetLean(3, 5, "10908")
    Call SetLean(4, 6, "10205")
    Call SetLean(5, 7, "10788")
    Call SetLean(6, 8, "10806")
    Call SetLean(7, 9, "10843")
    Call �o�^
    Call ���ʌ���
    Call ������
    
    Call SetRace(41, 1)
    Call SetLean(3, 5, "10357")
    Call SetLean(4, 6, "10609")
    Call SetLean(5, 7, "10417")
    Call SetLean(6, 8, "10665")
    Call �o�^
    Call ������
    
    Call SetRace(41, 2)
    Call SetLean(2, 4, "12154")
    Call SetLean(3, 5, "11460")
    Call SetLean(4, 6, "11580")
    Call SetLean(5, 7, "11239")
    Call SetLean(6, 8, "12302")
    Call �o�^
    Call ������
    
    Call SetRace(41, 3)
    Call SetLean(1, 3, "11034")
    Call SetLean(2, 4, "10749")
    Call SetLean(3, 5, "11444")
    Call SetLean(4, 6, "10455")
    Call SetLean(5, 7, "10823")
    Call SetLean(6, 8, "11687")
    Call SetLean(7, 9, "11689")
    Call �o�^
    Call ������
    
    Call SetRace(41, 4)
    Call SetLean(1, 3, "10528")
    Call SetLean(2, 4, "10443")
    Call SetLean(3, 5, "10141")
    Call SetLean(4, 6, "5455")
    Call SetLean(5, 7, "5813")
    Call SetLean(6, 8, "10248")
    Call SetLean(7, 9, "10338")
    Call �o�^
    Call ���ʌ���
    Call ������
    
    Call SetRace(43, 1)
    Call SetLean(3, 5, "4328")
    Call SetLean(4, 6, "")
    Call SetLean(5, 7, "4104")
    Call SetLean(6, 8, "4635")
    Call �o�^
    Call ������
    
    Call SetRace(43, 2)
    Call SetLean(3, 5, "")
    Call SetLean(4, 6, "3636")
    Call SetLean(5, 7, "3939")
    Call SetLean(6, 8, "4443")
    Call �o�^
    Call ���ʌ���
    Call ������
    
    Call SetRace(44, 1)
    Call SetLean(4, 6, "4286")
    Call �o�^
    Call ���ʌ���
    Call ������
    
    Call SetRace(45, 1)
    Call SetLean(1, 3, "10063")
    Call SetLean(2, 4, "5022")
    Call SetLean(3, 5, "5786")
    Call SetLean(4, 6, "5627")
    Call SetLean(5, 7, "")
    Call SetLean(6, 8, "4475")
    Call �o�^
    Call ���ʌ���
    Call ������
    
    Call SetRace(46, 1)
    Call SetLean(3, 5, "4404")
    Call SetLean(4, 6, "4090")
    Call SetLean(5, 7, "3959")
    Call SetLean(6, 8, "4141")
    Call �o�^
    Call ������
    
    Call SetRace(46, 2)
    Call SetLean(1, 3, "3900")
    Call SetLean(2, 4, "4019")
    Call SetLean(3, 5, "3293")
    Call SetLean(4, 6, "3099")
    Call SetLean(5, 7, "3526")
    Call SetLean(6, 8, "3525")
    Call SetLean(7, 9, "4026")
    Call �o�^
    Call ���ʌ���
    Call ������
    
    Call SetRace(47, 1)
    Call SetLean(3, 5, "3935")
    Call SetLean(4, 6, "3526")
    Call SetLean(5, 7, "3738")
    Call �o�^
    Call ���ʌ���
    Call ������
    
    Call SetRace(48, 1)
    Call SetLean(2, 4, "10629")
    Call SetLean(3, 5, "5186")
    Call SetLean(4, 6, "5087")
    Call SetLean(5, 7, "5567")
    Call SetLean(6, 8, "11465")
    Call SetLean(7, 9, "")
    Call �o�^
    Call ���ʌ���
    Call ������
    
    Call SetRace(49, 1)
    Call SetLean(3, 5, "4262")
    Call SetLean(4, 6, "3517")
    Call SetLean(5, 7, "3572")
    Call SetLean(6, 8, "4474")
    Call �o�^
    Call ������
    
    Call SetRace(49, 2)
    Call SetLean(3, 5, "3250")
    Call SetLean(4, 6, "3250")
    Call SetLean(5, 7, "3329")
    Call SetLean(6, 8, "3303")
    Call �o�^
    Call ���ʌ���
    Call ������
    
    Call SetRace(50, 1)
    Call SetLean(1, 3, "3665")
    Call SetLean(4, 6, "4607")
    Call SetLean(5, 7, "3348")
    Call SetLean(6, 8, "3724")
    Call SetLean(7, 9, "3393")
    Call �o�^
    Call ���ʌ���
    Call ������
    
    Call SetRace(52, 1)
    Call SetLean(3, 5, "3459")
    Call SetLean(4, 6, "3689")
    Call SetLean(5, 7, "3740")
    Call SetLean(6, 8, "3728")
    Call �o�^
    Call ������
    
    Call SetRace(52, 2)
    Call SetLean(2, 4, "3534")
    Call SetLean(3, 5, "3516")
    Call SetLean(4, 6, "3529")
    Call SetLean(5, 7, "3059")
    Call SetLean(6, 8, "3131")
    Call �o�^
    Call ������
    
    Call SetRace(52, 3)
    Call SetLean(2, 4, "3125")
    Call SetLean(3, 5, "3028")
    Call SetLean(4, 6, "2839")
    Call SetLean(5, 7, "2894")
    Call SetLean(6, 8, "3353")
    Call SetLean(7, 9, "3252")
    Call �o�^
    Call ���ʌ���
    Call ������
    
    Call SetRace(53, 1)
    Call SetLean(2, 4, "3335")
    Call SetLean(3, 5, "3069")
    Call SetLean(4, 6, "2671")
    Call SetLean(5, 7, "3122")
    Call SetLean(6, 8, "4283")
    Call SetLean(7, 9, "")
    Call �o�^
    Call ���ʌ���
    Call ������
    
    Call SetRace(54, 1)
    Call SetLean(1, 4, "4987")
    Call SetLean(2, 5, "4188")
    Call SetLean(3, 6, "4022")
    Call SetLean(4, 7, "3161")
    Call SetLean(5, 8, "3213")
    Call �o�^
    Call ������
    
    Call SetRace(54, 2)
    Call SetLean(3, 5, "3781")
    Call SetLean(4, 6, "3135")
    Call SetLean(5, 7, "3161")
    Call SetLean(6, 8, "3003")
    Call �o�^
    Call ���ʌ���
    Call ������
    
    Call SetRace(55, 1)
    Call SetLean(3, 5, "4333")
    Call SetLean(4, 6, "3778")
    Call SetLean(5, 7, "3785")
    Call SetLean(6, 8, "")
    Call �o�^
    Call ������
    
    Call SetRace(55, 2)
    Call SetLean(2, 4, "3965")
    Call SetLean(3, 5, "3417")
    Call SetLean(4, 6, "3406")
    Call SetLean(5, 7, "3467")
    Call SetLean(6, 8, "")
    Call �o�^
    Call ������
    
    Call SetRace(55, 3)
    Call SetLean(1, 3, "3453")
    Call SetLean(2, 4, "3204")
    Call SetLean(3, 5, "3231")
    Call SetLean(4, 6, "2990")
    Call SetLean(5, 7, "3098")
    Call SetLean(6, 8, "3209")
    Call SetLean(7, 9, "3451")
    Call �o�^
    Call ���ʌ���
    Call ������
    
    Call SetRace(56, 1)
    Call SetLean(1, 3, "4131")
    Call SetLean(2, 4, "4473")
    Call SetLean(3, 5, "")
    Call SetLean(4, 6, "3869")
    Call SetLean(5, 7, "")
    Call SetLean(6, 8, "4477")
    Call SetLean(7, 9, "4311")
    Call �o�^
    Call ������
    
    Call SetRace(56, 2)
    Call SetLean(1, 3, "3989")
    Call SetLean(2, 4, "3440")
    Call SetLean(3, 5, "3030")
    Call SetLean(4, 6, "2959")
    Call SetLean(5, 7, "3085")
    Call SetLean(6, 8, "3234")
    Call SetLean(7, 9, "3766")
    Call �o�^
    Call ���ʌ���
    Call ������
    
    Call SetRace(57, 1)
    Call SetLean(2, 4, "5687")
    Call SetLean(3, 5, "5718")
    Call SetLean(4, 6, "4581")
    Call SetLean(5, 7, "4517")
    Call �o�^
    Call ������
    
    Call SetRace(57, 2)
    Call SetLean(2, 4, "")
    Call SetLean(3, 5, "3852")
    Call SetLean(4, 6, "3403")
    Call SetLean(5, 7, "4070")
    Call SetLean(6, 8, "4047")
    Call SetLean(7, 9, "4037")
    Call �o�^
    Call ���ʌ���
    Call ������
    
    Call SetRace(58, 1)
    Call SetLean(3, 5, "3758")
    Call SetLean(4, 6, "3425")
    Call SetLean(5, 7, "3809")
    Call SetLean(6, 8, "4717")
    Call �o�^
    Call ������
    
    Call SetRace(58, 2)
    Call SetLean(2, 4, "")
    Call SetLean(3, 5, "3452")
    Call SetLean(4, 6, "3669")
    Call SetLean(5, 7, "")
    Call SetLean(6, 8, "3723")
    Call SetLean(7, 9, "")
    Call �o�^
    Call ������
    
    Call SetRace(58, 3)
    Call SetLean(1, 3, "3228")
    Call SetLean(2, 4, "3313")
    Call SetLean(3, 5, "3303")
    Call SetLean(4, 6, "3172")
    Call SetLean(5, 7, "3021")
    Call SetLean(6, 8, "3131")
    Call SetLean(7, 9, "3427")
    Call �o�^
    Call ������
    
    Call SetRace(58, 4)
    Call SetLean(1, 3, "2927")
    Call SetLean(2, 4, "")
    Call SetLean(3, 5, "3090")
    Call SetLean(4, 6, "2998")
    Call SetLean(5, 7, "3000")
    Call SetLean(6, 8, "3163")
    Call SetLean(7, 9, "3108")
    Call �o�^
    Call ������
    
    Call SetRace(58, 5)
    Call SetLean(1, 3, "2858")
    Call SetLean(2, 4, "3031")
    Call SetLean(3, 5, "2766")
    Call SetLean(4, 6, "")
    Call SetLean(5, 7, "2817")
    Call SetLean(6, 8, "2814")
    Call SetLean(7, 9, "")
    Call �o�^
    Call ���ʌ���
    Call ������
    
    Call SetRace(59, 1)
    Call SetLean(1, 3, "3926")
    Call SetLean(2, 4, "3583")
    Call SetLean(3, 5, "3846")
    Call SetLean(4, 6, "3686")
    Call SetLean(5, 7, "")
    Call SetLean(6, 8, "3457")
    Call SetLean(7, 9, "3486")
    Call �o�^
    Call ������
    
    Call SetRace(59, 2)
    Call SetLean(1, 3, "3277")
    Call SetLean(2, 4, "3182")
    Call SetLean(3, 5, "3237")
    Call SetLean(4, 6, "3079")
    Call SetLean(5, 7, "")
    Call SetLean(6, 8, "3227")
    Call SetLean(7, 9, "3162")
    Call �o�^
    Call ������
    
    Call SetRace(59, 3)
    Call SetLean(1, 3, "3123")
    Call SetLean(2, 4, "3137")
    Call SetLean(3, 5, "3118")
    Call SetLean(4, 6, "")
    Call SetLean(5, 7, "3055")
    Call SetLean(6, 8, "2919")
    Call SetLean(7, 9, "3002")
    Call �o�^
    Call ������
    
    Call SetRace(59, 4)
    Call SetLean(1, 3, "3017")
    Call SetLean(2, 4, "")
    Call SetLean(3, 5, "2856")
    Call SetLean(4, 6, "2930")
    Call SetLean(5, 7, "2909")
    Call SetLean(6, 8, "2952")
    Call SetLean(7, 9, "3014")
    Call �o�^
    Call ������
    
    Call SetRace(59, 5)
    Call SetLean(1, 3, "2821")
    Call SetLean(2, 4, "2881")
    Call SetLean(3, 5, "")
    Call SetLean(4, 6, "2747")
    Call SetLean(5, 7, "2740")
    Call SetLean(6, 8, "2841")
    Call SetLean(7, 9, "2805")
    Call �o�^
    Call ���ʌ���
    Call ������
    
    Call SetRace(60, 1)
    Call SetLean(1, 3, "4423")
    Call SetLean(2, 4, "3869")
    Call SetLean(3, 5, "4349")
    Call SetLean(4, 6, "3492")
    Call SetLean(5, 7, "3539")
    Call SetLean(6, 8, "3966")
    Call SetLean(7, 9, "4361")
    Call �o�^
    Call ������
    
    Call SetRace(60, 2)
    Call SetLean(3, 5, "3939")
    Call SetLean(4, 6, "3694")
    Call SetLean(5, 7, "3716")
    Call �o�^
    Call ������
    
    Call SetRace(60, 3)
    Call SetLean(2, 4, "3702")
    Call SetLean(3, 5, "3296")
    Call SetLean(4, 6, "2922")
    Call SetLean(5, 7, "3034")
    Call SetLean(6, 8, "3054")
    Call SetLean(7, 9, "2934")
    Call �o�^
    Call ������
    
    Call SetRace(60, 4)
    Call SetLean(2, 4, "")
    Call SetLean(3, 5, "2795")
    Call SetLean(4, 6, "2594")
    Call SetLean(5, 7, "2808")
    Call SetLean(6, 8, "3382")
    Call SetLean(7, 9, "3759")
    Call �o�^
    Call ������
    
    Call SetRace(60, 5)
    Call SetLean(2, 4, "3368")
    Call SetLean(3, 5, "3005")
    Call SetLean(4, 6, "2580")
    Call SetLean(5, 7, "2846")
    Call SetLean(6, 8, "3339")
    Call �o�^
    Call ���ʌ���
    Call ������
    
    Call SetRace(61, 1)
    Call SetLean(2, 4, "4411")
    Call SetLean(3, 5, "4965")
    Call SetLean(4, 6, "4694")
    Call SetLean(5, 7, "4896")
    Call SetLean(6, 8, "5187")
    Call �o�^
    Call ������
    
    Call SetRace(61, 2)
    Call SetLean(2, 4, "4574")
    Call SetLean(3, 5, "4710")
    Call SetLean(4, 6, "3936")
    Call SetLean(5, 7, "3837")
    Call SetLean(6, 8, "4242")
    Call �o�^
    Call ���ʌ���
    Call ������
    
    Call SetRace(62, 1)
    Call SetLean(2, 4, "5251")
    Call SetLean(3, 5, "4368")
    Call SetLean(4, 6, "4025")
    Call SetLean(5, 7, "3905")
    Call SetLean(6, 8, "4276")
    Call �o�^
    Call ���ʌ���
    Call ������
    
    Call SetRace(63, 1)
    Call SetLean(3, 5, "5149")
    Call SetLean(4, 6, "4705")
    Call SetLean(5, 7, "4370")
    Call SetLean(6, 8, "4291")
    Call �o�^
    Call ���ʌ���
    Call ������
    
    Call SetRace(64, 1)
    Call SetLean(2, 4, "5332")
    Call SetLean(3, 5, "5407")
    Call SetLean(4, 6, "4121")
    Call SetLean(5, 7, "4121")
    Call SetLean(6, 8, "5233")
    Call �o�^
    Call ������
    
    Call SetRace(64, 2)
    Call SetLean(1, 3, "4217")
    Call SetLean(2, 4, "3916")
    Call SetLean(3, 5, "3817")
    Call SetLean(4, 6, "3367")
    Call SetLean(5, 7, "3684")
    Call SetLean(6, 8, "3932")
    Call SetLean(7, 9, "")
    Call �o�^
    Call ���ʌ���
    Call ������
    
    Call SetRace(65, 1)
    Call SetLean(3, 5, "4148")
    Call SetLean(4, 6, "5294")
    Call SetLean(5, 7, "")
    Call SetLean(6, 8, "12304")
    Call �o�^
    Call ������
    
    Call SetRace(65, 2)
    Call SetLean(2, 4, "4222")
    Call SetLean(3, 5, "4239")
    Call SetLean(4, 6, "4525")
    Call SetLean(5, 7, "4232")
    Call SetLean(6, 8, "")
    Call �o�^
    Call ������
    
    Call SetRace(65, 3)
    Call SetLean(1, 3, "3958")
    Call SetLean(2, 4, "")
    Call SetLean(3, 5, "3721")
    Call SetLean(4, 6, "3462")
    Call SetLean(5, 7, "3687")
    Call SetLean(6, 8, "3787")
    Call SetLean(7, 9, "")
    Call �o�^
    Call ���ʌ���
    Call ������
    
    Call SetRace(66, 1)
    Call SetLean(3, 5, "4714")
    Call SetLean(4, 6, "4801")
    Call SetLean(5, 7, "5058")
    Call SetLean(6, 8, "10681")
    Call �o�^
    Call ������
    
    Call SetRace(66, 2)
    Call SetLean(2, 4, "4454")
    Call SetLean(3, 5, "4010")
    Call SetLean(4, 6, "3580")
    Call SetLean(5, 7, "3501")
    Call SetLean(6, 8, "4772")
    Call �o�^
    Call ���ʌ���
    Call ������
    
    Call SetRace(67, 1)
    Call SetLean(4, 6, "31220")
    Call SetLean(5, 7, "23717")
    Call �o�^
    Call ���ʌ���
    Call ������
    
    Call SetRace(68, 1)
    Call SetLean(2, 4, "")
    Call SetLean(3, 5, "22941")
    Call SetLean(4, 6, "20615")
    Call SetLean(5, 7, "22545")
    Call SetLean(6, 8, "20436")
    Call SetLean(7, 9, "21796")
    Call �o�^
    Call ���ʌ���
    Call ������
    
    Call SetRace(69, 1)
    Call SetLean(3, 5, "22397")
    Call SetLean(4, 6, "21915")
    Call SetLean(5, 7, "23395")
    Call �o�^
    Call ���ʌ���
    Call ������
    
    Call SetRace(70, 1)
    Call SetLean(4, 6, "15782")
    Call SetLean(5, 7, "23187")
    Call �o�^
    Call ���ʌ���
    Call ������
    
    Call SetRace(71, 1)
    Call SetLean(2, 4, "20869")
    Call SetLean(3, 5, "21902")
    Call SetLean(4, 6, "15415")
    Call SetLean(5, 7, "20069")
    Call SetLean(6, 8, "20970")
    Call SetLean(7, 9, "20218")
    Call �o�^
    Call ���ʌ���
    Call ������
    
    Call SetRace(72, 1)
    Call SetLean(1, 3, "20111")
    Call SetLean(2, 4, "20657")
    Call SetLean(3, 5, "15166")
    Call SetLean(4, 6, "15471")
    Call SetLean(5, 7, "15622")
    Call SetLean(6, 8, "15981")
    Call SetLean(7, 9, "21366")
    Call �o�^
    Call ���ʌ���
    Call ������
    
    ' �u�b�N�̕ۑ�
    ActiveWorkbook.Save
End Sub
