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
    Call ModCell("F9", "5")
    Call ModCell("G9", "��ؗ����D���P�����D���R�肨�D���c����")
    Call ModCell("C10", "2")
    Call ModCell("F10", "8")
    Call ModCell("G10", "�֓��z�q�D�{�薾�q�D�{�Y�߂��݁D��،c�q")
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
    Call ModCell("C22", "5")
    Call ModCell("F22", "8")
    Call ModCell("G22", "��؎��P�D�p�c�S���D�����āD�{��_�i")
    Call ModCell("C23", "5")
    Call ModCell("F23", "9")
    Call ModCell("G23", "��؏C���D�z�G��D�O�����a�D�㌴�D�m")
    
    ' 7�|9�@���ꃌ�[�X
    Call ModCell(SearchCell(7, "�����@�捁", "���[��"), "3")
    Call ModCell(SearchCell(7, "�J�{�@�Ďq", "���[��"), "4")
    Call ModCell(SearchCell(7, "�쑺�@���F", "���[��"), "5")
    Call ModCell(SearchCell(7, "�|���@�ĉ�", "���[��"), "6")
    Call ModCell(SearchCell(7, "���c�@����", "���[��"), "7")
    Call ModCell(SearchCell(9, "���J�@�K�]", "���[�XNo"), "7")
    Call ModCell(SearchCell(9, "���J�@�K�]", "���[��"), "8")
    
    ' 13
    Call ModCell(SearchCell(13, "�Ȍ��@����", "���[��"), "8")
    Call ModCell(SearchCell(13, "�}���@����", "���[��"), "5")
    Call ModCell(SearchCell(13, "����@�ʔT", "���[��"), "4")

    ' 17
    Call ModCell(SearchCell(17, "�c���@�D�M��", "���[��"), "3")
    Call ModCell(SearchCell(17, "���V�@�q��", "���[��"), "4")
    Call ModCell(SearchCell(17, "�R���@����", "���[��"), "5")
    Call ModCell(SearchCell(17, "���c�@����", "���[��"), "6")
    Call ModCell(SearchCell(17, "���c�@�Ќ�", "���[��"), "7")

    ' 18
    Call ModCell(SearchCell(18, "���@����", "���[�XNo"), "14")
    Call ModCell(SearchCell(18, "���@����", "���[��"), "9")

    ' 19
    Call ModCell(SearchCell(17, "�Έ�@���D��", "���[��"), "4")
    Call ModCell(SearchCell(17, "�~���@��l", "���[��"), "5")
    Call ModCell(SearchCell(17, "���i�@�劒", "���[��"), "6")

    ' 20
    Call ModCell(SearchCell(20, "��؁@����", "���[�XNo"), "16")
    Call ModCell(SearchCell(20, "��؁@����", "���[��"), "8")

    ' 34
    Call ModCell(SearchCell(34, "�|�R�@�C��", "���[��"), "5")
    Call ModCell(SearchCell(34, "�ɓ��@����", "���[�XNo"), "25")
    Call ModCell(SearchCell(34, "�ɓ��@����", "�g"), "1")
    Call ModCell(SearchCell(34, "�ɓ��@����", "���[��"), "6")
    Call ModCell(SearchCell(34, "�O�c�@�x�l", "���[��"), "7")
    Call ModCell(SearchCell(34, "�|�с@�Ж�", "���[��"), "8")

    ' 37
    Call ModCell(SearchCell(37, "�쑺�@���F", "���[�XNo"), "30")
    Call ModCell(SearchCell(37, "�쑺�@���F", "�g"), "1")
    Call ModCell(SearchCell(37, "�쑺�@���F", "���[��"), "6")
    Call ModCell(SearchCell(37, "���R�@�@��", "���[�XNo"), "30")
    Call ModCell(SearchCell(37, "���R�@�@��", "�g"), "1")
    Call ModCell(SearchCell(37, "���R�@�@��", "���[��"), "8")
    Call ModCell(SearchCell(37, "�O�c�@����", "���[��"), "4")
    Call ModCell(SearchCell(37, "�F�J�@�M��", "���[��"), "5")
    Call ModCell(SearchCell(37, "�����@�D�C", "���[��"), "6")
    Call ModCell(SearchCell(37, "���с@�ѓ�", "���[��"), "7")
    Call ModCell(SearchCell(37, "���с@�э�", "���[��"), "8")
    Call ModCell(SearchCell(37, "�|���@�^�D", "���[�XNo"), "31")
    Call ModCell(SearchCell(37, "�|���@�^�D", "�g"), "2")
    Call ModCell(SearchCell(37, "�|���@�^�D", "���[��"), "9")

    ' 38-39
    Call ModCell(SearchCell(38, "��؁@���a�q", "���[��"), "4")
    Call ModCell(SearchCell(38, "����@�Y��", "���[��"), "5")
    Call ModCell(SearchCell(38, "�����@�����", "���[��"), "6")
    Call ModCell(SearchCell(38, "���c�@����", "���[��"), "7")
    Call ModCell(SearchCell(39, "��؁@�c�q", "���[�XNo"), "32")
    Call ModCell(SearchCell(39, "��؁@�c�q", "���[��"), "9")

    ' 40
    Call ModCell(SearchCell(40, "�c���@��M", "���[��"), "5")
    Call ModCell(SearchCell(40, "�V�{�@�C��", "���[�XNo"), "34")
    Call ModCell(SearchCell(40, "�V�{�@�C��", "�g"), "1")
    Call ModCell(SearchCell(40, "�V�{�@�C��", "���[��"), "6")
    Call ModCell(SearchCell(40, "��J�@�A�m", "���[��"), "7")
    Call ModCell(SearchCell(40, "�n粁@�s��", "���[�XNo"), "34")
    Call ModCell(SearchCell(40, "�n粁@�s��", "�g"), "1")
    Call ModCell(SearchCell(40, "�n粁@�s��", "���[��"), "8")
    Call ModCell(SearchCell(40, "���c�@����", "���[�XNo"), "35")
    Call ModCell(SearchCell(40, "���c�@����", "�g"), "2")
    Call ModCell(SearchCell(40, "���c�@����", "���[��"), "3")
    Call ModCell(SearchCell(40, "�|�R�@�C��", "���[��"), "4")
    Call ModCell(SearchCell(40, "�����@�I�M", "���[��"), "5")
    Call ModCell(SearchCell(40, "��{�@�G�l", "���[��"), "6")
    Call ModCell(SearchCell(40, "�c���@�D�l", "���[��"), "7")
    Call ModCell(SearchCell(40, "�x�m��@�l��", "���[��"), "8")
    Call ModCell(SearchCell(40, "�����@�D�V��", "���[�XNo"), "35")
    Call ModCell(SearchCell(40, "�����@�D�V��", "�g"), "2")
    Call ModCell(SearchCell(40, "�����@�D�V��", "���[��"), "9")
    Call ModCell(SearchCell(40, "����@����", "���[�XNo"), "36")
    Call ModCell(SearchCell(40, "����@����", "�g"), "3")
    Call ModCell(SearchCell(40, "����@����", "���[��"), "3")
    Call ModCell(SearchCell(40, "�����@�\��Y", "���[��"), "4")
    Call ModCell(SearchCell(40, "�k��@�D��", "���[��"), "5")
    Call ModCell(SearchCell(40, "���Y�@�C��", "���[��"), "6")
    Call ModCell(SearchCell(40, "����@�m���N", "���[��"), "7")
    Call ModCell(SearchCell(40, "�쓈�@�N��", "���[��"), "8")
    Call ModCell(SearchCell(40, "���V�@�đ�", "���[�XNo"), "36")
    Call ModCell(SearchCell(40, "���V�@�đ�", "�g"), "3")
    Call ModCell(SearchCell(40, "���V�@�đ�", "���[��"), "9")
    
    
    Call ModCell(SearchCell(41, "��؁@��M", "���[�XNo"), "37")
    Call ModCell(SearchCell(41, "��؁@��M", "�g"), "1")
    Call ModCell(SearchCell(41, "��؁@��M", "���[��"), "5")
    Call ModCell(SearchCell(41, "�c���@�C��", "���[��"), "6")
    Call ModCell(SearchCell(41, "��粁@�I��", "���[�XNo"), "37")
    Call ModCell(SearchCell(41, "��粁@�I��", "�g"), "1")
    Call ModCell(SearchCell(41, "��粁@�I��", "���[��"), "7")
    Call ModCell(SearchCell(41, "�����@�i�m", "���[�XNo"), "37")
    Call ModCell(SearchCell(41, "�����@�i�m", "�g"), "1")
    Call ModCell(SearchCell(41, "�����@�i�m", "���[��"), "8")
    
    Call ModCell(SearchCell(41, "�X��@�^��", "���[�XNo"), "38")
    Call ModCell(SearchCell(41, "�X��@�^��", "�g"), "2")
    Call ModCell(SearchCell(41, "�X��@�^��", "���[��"), "4")
    Call ModCell(SearchCell(41, "�����@���u", "���[��"), "5")
    Call ModCell(SearchCell(41, "���v�ԁ@�T", "���[��"), "6")
    Call ModCell(SearchCell(41, "����@�@�A", "���[��"), "7")
    Call ModCell(SearchCell(41, "��@�q��", "���[�XNo"), "38")
    Call ModCell(SearchCell(41, "��@�q��", "�g"), "2")
    Call ModCell(SearchCell(41, "��@�q��", "���[��"), "8")
    
    Call ModCell(SearchCell(41, "�}�L�@���", "���[�XNo"), "39")
    Call ModCell(SearchCell(41, "�}�L�@���", "�g"), "3")
    Call ModCell(SearchCell(41, "�}�L�@���", "���[��"), "3")
    Call ModCell(SearchCell(41, "�����@�@��", "���[�XNo"), "39")
    Call ModCell(SearchCell(41, "�����@�@��", "�g"), "3")
    Call ModCell(SearchCell(41, "�����@�@��", "���[��"), "4")
    Call ModCell(SearchCell(41, "�쑺�@���", "���[�XNo"), "39")
    Call ModCell(SearchCell(41, "�쑺�@���", "�g"), "3")
    Call ModCell(SearchCell(41, "�쑺�@���", "���[��"), "5")
    Call ModCell(SearchCell(41, "�㓡�@����", "���[��"), "6")
    Call ModCell(SearchCell(41, "��؁@����", "���[��"), "7")
    Call ModCell(SearchCell(41, "�O�J�@���", "���[��"), "8")
    Call ModCell(SearchCell(41, "�x�]�@�a��", "���[��"), "9")
    
    Call ModCell(SearchCell(41, "�����@�g�P", "���[�XNo"), "40")
    Call ModCell(SearchCell(41, "�����@�g�P", "�g"), "4")
    Call ModCell(SearchCell(41, "�����@�g�P", "���[��"), "3")
    Call ModCell(SearchCell(41, "�c���@�D�M��", "���[��"), "4")
    Call ModCell(SearchCell(41, "����@�č�", "���[��"), "8")

    Call ModCell(SearchCell(43, "����@�ĊC", "���[�XNo"), "41")
    Call ModCell(SearchCell(43, "����@�ĊC", "�g"), "1")
    Call ModCell(SearchCell(43, "����@�ĊC", "���[��"), "6")
    Call ModCell(SearchCell(43, "�����@�G����", "���[��"), "7")
    Call ModCell(SearchCell(43, "��؁@���", "���[��"), "8")

    Call ModCell(SearchCell(45, "�O��@�����}", "���[��"), "3")
    Call ModCell(SearchCell(45, "���@���q", "���[��"), "4")
    Call ModCell(SearchCell(45, "�ā@�@�b�q", "���[��"), "5")
    Call ModCell(SearchCell(45, "�����@�a�}", "���[��"), "6")
    Call ModCell(SearchCell(45, "��؁@���q", "���[��"), "7")
    Call ModCell(SearchCell(45, "�q���@�R���q", "���[��"), "8")

    Call ModCell(SearchCell(48, "���c�@�@��", "���[��"), "4")
    Call ModCell(SearchCell(48, "", "���[��"), "5")
    Call ModCell(SearchCell(48, "", "���[��"), "6")
    Call ModCell(SearchCell(48, "", "���[��"), "7")
    Call ModCell(SearchCell(48, "", "���[��"), "8")
    Call ModCell(SearchCell(48, "", "���[��"), "9")


    Call ModCell(SearchCell(48, "", "���[�XNo"), "48")
    Call ModCell(SearchCell(48, "", "�g"), "1")
    Call ModCell(SearchCell(48, "", "���[��"), "3")

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
