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
    Call SetLean(1, 3, "34709")
    Call SetLean(2, 4, "25977")
    Call SetLean(3, 6, "")
    Call SetLean(4, 8, "30973")
    Call �o�^
    Call ���ʌ���
    Call ������
    
    ' �u�b�N�̕ۑ�
    ActiveWorkbook.Save
End Sub
