Attribute VB_Name = "KidsMastersCorrectionModule"
'
' �w���}�X�^�[�Y�̃v���O�����␳
'
Sub �w�}�v���O�����␳()

    Call EventChange(False)

    Sheets("�G���g���[�ꗗ").Activate
    ' 1-2-4 ���ꃌ�[�X
    Call ModCell("F2", "3")
    Call ModCell("G2", "���J���O�q�E�V����Ɏq�E�����E���c�ߑ�")
    Call ModCell("F3", "4")
    Call ModCell("G3", "���a���E�߉������E����؎q�E�ؓc���Îq")
    Call ModCell("C4", "10")
    Call ModCell("C5", "10")
    Call ModCell("F5", "8")
    Call ModCell("G5", "�Óc�G�u�N�E���R�V�E���ё�ՁE������Y")
    Call ModCell("G6", "�����M�q�E��ؔ��E�V�����j�E�I���b��")
    Call ModCell("G7", "�����a�}�E�|�����q�E�ۖΌ���E���q����")
    Call ModCell("G8", "��؎��P�E�勴�b�q�E���D���E�X�쏹�O")
    Call ModCell("G9", "�⍪���w�E�s�یb�q�E���X�L���E�쑺�m�q")
    ' 12-13 ���ꃌ�[�X
    Call ModCell(SearchCell(12, "��؁@�Ƃݎq", "���[��"), "4") ' "F33"
    Call ModCell(SearchCell(12, "���@���q", "���[��"), "5") ' "F34"
    Call ModCell(SearchCell(13, "���J�@�f�O", "���[�XNo"), "110") ' "C35"
    Call ModCell(SearchCell(13, "���J�@�f�O", "���[��"), "7") ' "F35"

    ' 19 �j�q100M���R�`
    Call ModCell(SearchCell(19, "�����@�P�v", "���[��"), "4") ' "F47"
    Call ModCell(SearchCell(19, "�y�c�@���Y", "���[��"), "6") '"F48"

    ' 26-29 ���ꃌ�[�X
    Call ModCell(SearchCell(26, "��؁@������", "���[��"), "3") ' "F78"
    Call ModCell(SearchCell(29, "���`�@�P�C�m", "���[�XNo"), "240") ' "C79"
    Call ModCell(SearchCell(29, "���`�@�P�C�m", "���[��"), "5")  ' "F79"
    Call ModCell(SearchCell(29, "�O���n���@�j�R��", "���[�XNo"), "240") '"C80"
    Call ModCell(SearchCell(29, "�O���n���@�j�R��", "���[��"), "6") '"F80"
    Call ModCell(SearchCell(29, "���R�@�葾�N", "���[�XNo"), "240") '"C81"
    Call ModCell(SearchCell(29, "���R�@�葾�N", "���[��"), "7") '"F81"
    Call ModCell(SearchCell(29, "�g�c�@���C", "���[�XNo"), "240") '"C82"
    Call ModCell(SearchCell(29, "�g�c�@���C", "���[��"), "8") '"F82"
    
    ' 33-34 ���ꃌ�[�X
    Call ModCell(SearchCell(33, "�x�@�@�\�C", "���[��"), "4")  '"F93"
    Call ModCell(SearchCell(34, "���{�@���", "���[�XNo"), "280") '"C94"
    Call ModCell(SearchCell(34, "���{�@���", "���[��"), "7")  '"F94"
    Call ModCell(SearchCell(34, "�Έ�@�t��", "���[�XNo"), "280") '"C95"
    Call ModCell(SearchCell(34, "�Έ�@�t��", "���[��"), "6")  '"F95"
    
    ' 36 �������h���[�����[
    Call ModCell("G99", "�]���^���E�����E���c�[��E���L�^�i")
    
    ' 37 ���q50M�w�j��
    Call ModCell(SearchCell(37, "���@����q", "���[��"), "3")  ' "F100"
    Call ModCell(SearchCell(37, "�O��@�����}", "���[��"), "4")  ' "F101"
    Call ModCell(SearchCell(37, "�ā@�@�b�q", "���[��"), "5")  ' "F102"
    Call ModCell(SearchCell(37, "���J���@�O�q", "���[�XNo"), "320") ' "C103"
    Call ModCell(SearchCell(37, "���J���@�O�q", "�g"), "1")  ' "E103"
    Call ModCell(SearchCell(37, "���J���@�O�q", "���[��"), "8")  ' "F103"
    Call ModCell(SearchCell(37, "�����@�a�}", "���[�XNo"), "320") ' "C104"
    Call ModCell(SearchCell(37, "�����@�a�}", "�g"), "1")  ' "E104"
    Call ModCell(SearchCell(37, "�����@�a�}", "���[��"), "6")  ' "F104"
    Call ModCell(SearchCell(37, "���ԁ@��ߎq", "���[�XNo"), "320") ' "C105"
    Call ModCell(SearchCell(37, "���ԁ@��ߎq", "�g"), "1")  ' "E105"
    Call ModCell(SearchCell(37, "���ԁ@��ߎq", "���[��"), "7")  ' "F105"
    Call ModCell(SearchCell(37, "���@�@�a��", "���[��"), "4")  ' "F106"
    Call ModCell(SearchCell(37, "��؁@���q", "���[��"), "5")  ' "F107"
    Call ModCell(SearchCell(37, "�q���@�R���q", "���[��"), "6")  ' "F108"
    Call ModCell(SearchCell(37, "�c�Y�@��b", "���[��"), "7")  ' "F109"
    
    ' 40 ���w�P�E�Q�N�j�q50M�w�j��
    Call ModCell(SearchCell(40, "���g�@�ё�", "���[��"), "4")  ' "F124"
    Call ModCell(SearchCell(40, "�щ��@���m�Y", "���[��"), "7")  ' "F125"
    Call ModCell(SearchCell(40, "�g��@���l", "���[��"), "8")  ' "F126"
    Call ModCell(SearchCell(40, "�����@�@�z", "���[�XNo"), "370") ' "C127"
    Call ModCell(SearchCell(40, "�����@�@�z", "�g"), "1")  ' "E127"
    Call ModCell(SearchCell(40, "�����@�@�z", "���[��"), "6")  ' "F127"
    Call ModCell(SearchCell(40, "�H���@�f��", "���[�XNo"), "370") ' "C133"
    Call ModCell(SearchCell(40, "�H���@�f��", "�g"), "1")  ' "E133"
    Call ModCell(SearchCell(40, "�H���@�f��", "���[��"), "5")  ' "F133"
    
    ' 41 ���q50M�w�j��
    Call ModCell(SearchCell(41, "��؁@�G�ˍ�", "���[��"), "6")
    Call ModCell(SearchCell(41, "���c�@����", "���[�XNo"), "390") ' "C137"
    Call ModCell(SearchCell(41, "���c�@����", "�g"), "1")  ' "E137"
    Call ModCell(SearchCell(41, "���c�@����", "���[��"), "5")  ' "F137"
    
    ' 43 ���q50M�w�j��
    Call ModCell(SearchCell(43, "���V�@�݂���", "���[��"), "8")  ' "F152"
    Call ModCell(SearchCell(43, "���V�@����", "���[��"), "7")  ' "F153"
    Call ModCell(SearchCell(43, "��؁@���", "���[��"), "5")  ' "F154"
    Call ModCell(SearchCell(43, "�Ȍ��@���", "���[��"), "4")  ' "F155"
    Call ModCell(SearchCell(43, "���i�@�D��", "���[��"), "3")  ' "F156"
    Call ModCell(SearchCell(43, "�����@�z��", "���[�XNo"), "430") ' "C163"
    Call ModCell(SearchCell(43, "�����@�z��", "�g"), "1")  ' "E163"
    Call ModCell(SearchCell(43, "�����@�z��", "���[��"), "6")  ' "F163"
    
    ' 44 �j�q50M�w�j��
    Call ModCell(SearchCell(44, "�с@�@��^", "���[��"), "4")
    Call ModCell(SearchCell(44, "��؁@�֗F", "���[��"), "5")
    Call ModCell(SearchCell(44, "�y���@�D��", "���[�XNo"), "450")
    Call ModCell(SearchCell(44, "�y���@�D��", "�g"), "1")
    Call ModCell(SearchCell(44, "�y���@�D��", "���[��"), "6")
    Call ModCell(SearchCell(44, "�����@�g��", "���[��"), "7")
    Call ModCell(SearchCell(44, "�с@�@��q", "���[��"), "8")
    
    ' 45 ���q50M���R�`
    Call ModCell(SearchCell(45, "�R�{�@�h�q", "���[��"), "3")  ' "F175"
    Call ModCell(SearchCell(45, "���J�@�Ƃ��q", "���[��"), "5")  ' "F176"
    Call ModCell(SearchCell(45, "�����@�M�q", "���[��"), "4")  ' "F177"
    Call ModCell(SearchCell(45, "�V���@��Ɏq", "���[�XNo"), "470") ' "C178"
    Call ModCell(SearchCell(45, "�V���@��Ɏq", "�g"), "1")  ' "E178"
    Call ModCell(SearchCell(45, "�V���@��Ɏq", "���[��"), "8")  ' "F178"
    Call ModCell(SearchCell(45, "�I���@�b��", "���[�XNo"), "470") ' "C179"
    Call ModCell(SearchCell(45, "�I���@�b��", "�g"), "1")  ' "E179"
    Call ModCell(SearchCell(45, "�I���@�b��", "���[��"), "6")  ' "F179"
    Call ModCell(SearchCell(45, "�����@����q", "���[�XNo"), "470") ' "C180"
    Call ModCell(SearchCell(45, "�����@����q", "�g"), "1")  ' "E180"
    Call ModCell(SearchCell(45, "�����@����q", "���[��"), "7")  ' "F180"
    Call ModCell(SearchCell(45, "����@��q", "���[��"), "4")  ' "F181"
    Call ModCell(SearchCell(45, "�R���@�F�q", "���[��"), "5")  ' "F182"
    Call ModCell(SearchCell(45, "���V�@���Ƃ�", "���[��"), "6")  ' "F183"
    Call ModCell(SearchCell(45, "���J���@�O�q", "���[��"), "8")  ' "F184"
    Call ModCell(SearchCell(45, "���c�@�[��", "���[�XNo"), "480") ' "C185"
    Call ModCell(SearchCell(45, "���c�@�[��", "�g"), "2")  ' "E185"
    Call ModCell(SearchCell(45, "���c�@�[��", "���[��"), "7")  ' "F185"
    Call ModCell(SearchCell(45, "����@�@��", "���[��"), "3")  ' "F186"
    Call ModCell(SearchCell(45, "�߉��@����", "���[��"), "4")  ' "F187"
    Call ModCell(SearchCell(45, "���@�@�a��", "���[��"), "5")  ' "F188"
    Call ModCell(SearchCell(45, "�ؓc�@���Îq", "���[��"), "6")  ' "F189"
    Call ModCell(SearchCell(45, "����@�؎q", "���[��"), "7")  ' "F190"
    Call ModCell(SearchCell(45, "ꎓ��@����", "���[��"), "8")  ' "F191"
    
    ' 46 �j�q50M���R�`
    Call ModCell(SearchCell(46, "�����@�M�m", "���[��"), "5")  ' "F195"
    Call ModCell(SearchCell(46, "���ԁ@�G��", "���[��"), "6")  ' "F194"
    Call ModCell(SearchCell(46, "���c�@���F", "���[��"), "7")
    Call ModCell(SearchCell(46, "����@����", "���[��"), "8")
    Call ModCell(SearchCell(46, "�R�{�@���l", "���[��"), "9")
    Call ModCell(SearchCell(46, "�⍪�@���w", "���[��"), "3")
    Call ModCell(SearchCell(46, "���@����", "���[��"), "5")
    
    ' 49 ���q50M���R�`
    Call ModCell(SearchCell(49, "�i��@���G", "���[��"), "4")  ' "F240"
    Call ModCell(SearchCell(49, "�V���@�S�X��", "���[��"), "7")  ' "F241"
    Call ModCell(SearchCell(49, "���a�c�@�z��", "���[��"), "8")  ' "F242"
    Call ModCell(SearchCell(49, "��ԁ@���", "���[�XNo"), "570") ' "C243"
    Call ModCell(SearchCell(49, "��ԁ@���", "�g"), "1")  ' "E243"
    Call ModCell(SearchCell(49, "��ԁ@���", "���[��"), "6")  ' "F243"
    Call ModCell(SearchCell(49, "����@����", "���[��"), "8")  ' "F244"
    Call ModCell(SearchCell(49, "���c�@����", "���[��"), "7")  ' "F245"
    Call ModCell(SearchCell(49, "�֓��@�R��", "���[��"), "5")  ' "F246"
    Call ModCell(SearchCell(49, "�V���@���b��", "���[��"), "4")  ' "F247"
    Call ModCell(SearchCell(49, "�����@���D��", "���[��"), "3")  ' "F248"
    Call ModCell(SearchCell(49, "��؁@�G�ˍ�", "���[�XNo"), "570") ' "C249"
    Call ModCell(SearchCell(49, "��؁@�G�ˍ�", "�g"), "1")  ' "E249"
    Call ModCell(SearchCell(49, "��؁@�G�ˍ�", "���[��"), "5")  ' "F249"
    Call ModCell(SearchCell(49, "�}���@����", "���[�XNo"), "580") ' "C256"
    Call ModCell(SearchCell(49, "�}���@����", "�g"), "2")  ' "E256"
    Call ModCell(SearchCell(49, "�}���@����", "���[��"), "6")  ' "F256"
    ' 50 �j�q50M���R�`
    Call ModCell(SearchCell(50, "���V�@�P��", "���[��"), "3")  ' "F257"
    Call ModCell(SearchCell(50, "��؁@����", "���[��"), "4")  ' "F258"
    Call ModCell(SearchCell(50, "�O�y�@�[�K", "���[��"), "7")  ' "F259"
    Call ModCell(SearchCell(50, "���V�@����", "���[��"), "8")  ' "F260"
    Call ModCell(SearchCell(50, "����@�@��", "���[��"), "3")  ' "F261"
    Call ModCell(SearchCell(50, "�ΐ�@�ɑ�", "���[��"), "8")  ' "F262"
    Call ModCell(SearchCell(50, "�����@�Y��", "���[��"), "7")  ' "F263"
    Call ModCell(SearchCell(50, "���V�@�D��", "���[��"), "5")  ' "F264"
    Call ModCell(SearchCell(50, "�g�c�@���i", "���[��"), "4")  ' "F265"
    Call ModCell(SearchCell(50, "�ɓ��@�@��", "���[�XNo"), "600") ' "C266"
    Call ModCell(SearchCell(50, "�ɓ��@�@��", "�g"), "1")  ' "E266"
    Call ModCell(SearchCell(50, "�ɓ��@�@��", "���[��"), "6")  ' "F266"
    Call ModCell(SearchCell(50, "�〈�@�z���N", "���[�XNo"), "600") ' "C267"
    Call ModCell(SearchCell(50, "�〈�@�z���N", "�g"), "1")  ' "E267"
    Call ModCell(SearchCell(50, "�〈�@�z���N", "���[��"), "5")  ' "F267"
    Call ModCell(SearchCell(50, "��@�@�����Y", "���[�XNo"), "610") ' "C274"
    Call ModCell(SearchCell(50, "��@�@�����Y", "�g"), "1")  ' "E274"
    Call ModCell(SearchCell(50, "��@�@�����Y", "���[��"), "6")  ' "F274"
    
    ' 51 ���q50M���R�`
    Call ModCell(SearchCell(51, "���V�@�݂���", "���[��"), "4")  ' "F275"
    Call ModCell(SearchCell(51, "�V���@�S����", "���[��"), "7")  ' "F276"
    Call ModCell(SearchCell(51, "���i�@�D��", "���[��"), "8")  ' "F277"
    Call ModCell(SearchCell(51, "�_���@����", "���[�XNo"), "630") ' "C278"
    Call ModCell(SearchCell(51, "�_���@����", "�g"), "1")  ' "E278"
    Call ModCell(SearchCell(51, "�_���@����", "���[��"), "6")  ' "F278"
    Call ModCell(SearchCell(51, "�Ȍ��@���", "���[��"), "4")  ' "F279"
    Call ModCell(SearchCell(51, "���V�@����", "���[��"), "7")  ' "F280"
    Call ModCell(SearchCell(51, "��؁@���", "���[��"), "8")  ' "F281"
    Call ModCell(SearchCell(51, "�k�c�@�Г�", "���[�XNo"), "630") ' "C282"
    Call ModCell(SearchCell(51, "�k�c�@�Г�", "�g"), "1")  ' "E282"
    Call ModCell(SearchCell(51, "�k�c�@�Г�", "���[��"), "5")  ' "F282"
    Call ModCell(SearchCell(51, "��@�@�b�m�a", "���[�XNo"), "640") ' "C283"
    Call ModCell(SearchCell(51, "��@�@�b�m�a", "�g"), "2")  ' "E283"
    Call ModCell(SearchCell(51, "��@�@�b�m�a", "���[��"), "6")  ' "F283"
    Call ModCell(SearchCell(51, "����@�ʔT", "���[�XNo"), "640") ' "C289"
    Call ModCell(SearchCell(51, "����@�ʔT", "�g"), "2")  ' "E289"
    Call ModCell(SearchCell(51, "����@�ʔT", "���[��"), "5")  ' "F289"
    
    ' 52 �j�q50M���R�`
    Call ModCell(SearchCell(52, "���{�@�đ�", "���[��"), "6")  ' "F292"
    Call ModCell(SearchCell(52, "�v�ۓc�@��", "���[�XNo"), "670") ' "C293"
    Call ModCell(SearchCell(52, "�v�ۓc�@��", "�g"), "2")  ' "E293"
    Call ModCell(SearchCell(52, "�v�ۓc�@��", "���[��"), "9")  ' "F293"
    Call ModCell(SearchCell(52, "�����@�g��", "���[��"), "5")  ' "F294"
    Call ModCell(SearchCell(52, "�����@�@��", "���[�XNo"), "660") ' "C296"
    Call ModCell(SearchCell(52, "�����@�@��", "�g"), "1")  ' "E296"
    Call ModCell(SearchCell(52, "�����@�@��", "���[��"), "7")  ' "F296"
    Call ModCell(SearchCell(52, "����@�A��", "���[��"), "3")  ' "F302"
    
    ' 53 ���q50M�o�^�t���C
    Call ModCell(SearchCell(53, "���c�@�ߑ�", "���[��"), "7")  ' "F311"
    Call ModCell(SearchCell(53, "���J�@�K�]", "���[��"), "6")  ' "F312"
    
    ' 54 �j�q50M�o�^�t���C
    Call ModCell(SearchCell(54, "�z��@�B�K", "���[�XNo"), "700") ' "C296"
    Call ModCell(SearchCell(54, "�z��@�B�K", "�g"), "1")  ' "E296"
    Call ModCell(SearchCell(54, "�z��@�B�K", "���[��"), "8")  ' "F296"
    
    
    ' 61 ���q50M���j��
    Call ModCell(SearchCell(61, "���V�@���Ƃ�", "���[��"), "8")  ' "F344"
    Call ModCell(SearchCell(61, "�|���@���q", "���[�XNo"), "780") ' "C345"
    Call ModCell(SearchCell(61, "�|���@���q", "�g"), "1")  ' "E345"
    Call ModCell(SearchCell(61, "�|���@���q", "���[��"), "6")  ' "F345"
    Call ModCell(SearchCell(61, "�s�ہ@�b�q", "���[�XNo"), "780") ' "C346"
    Call ModCell(SearchCell(61, "�s�ہ@�b�q", "�g"), "1")  ' "E346"
    Call ModCell(SearchCell(61, "�s�ہ@�b�q", "���[��"), "7")  ' "F346"
    Call ModCell(SearchCell(61, "�߉��@����", "���[��"), "4")  ' "F347"
    Call ModCell(SearchCell(61, "���V�@����q", "���[��"), "5")  ' "F349"
    Call ModCell(SearchCell(61, "�ؓc�@���Îq", "���[��"), "6")  ' "F348"
    Call ModCell(SearchCell(61, "�呺�@��q", "���[��"), "7")  ' "F350"
    ' 70-71 ���ꃌ�[�X
    Call ModCell("G386", "�R���F�q�E�R�{�h�q�E�O������}�E�Čb�q")
    Call ModCell("G387", "�쑺�m�q�E�s�یb�q�E��������q�E���Ԑ�ߎq")
    Call ModCell("G388", "���Ώ~�q�E���V����q�E��㋞�q�E�呺��q")
    Call ModCell("G389", "����؎q�E���a���E�ؓc���Îq�E�߉�����")
    Call ModCell("G390", "�����E�����\��E���씎�ہE�R�c���N")
    Call ModCell("G391", "���c�׈�E���c���F�E�؍v�E�،�j")
    Call ModCell("G392", "���ьc�Y�E���R�O�E�쐴�u�E�r�c�O")
    Call ModCell("F393", "3")
    Call ModCell("G393", "�ē��S�����E��ؖG�ˍʁE���c�����E���숺��")
    Call ModCell("F394", "4")
    Call ModCell("G394", "���V�����E��ؔ�ʁE���i�D��E�Ȍ����")
    Call ModCell("C395", "900")
    Call ModCell("F395", "6")
    Call ModCell("G395", "��ؗ֗F�E���쏯�E�ɓ����E�O�y�[�K")
    Call ModCell("C396", "900")
    Call ModCell("F396", "7")
    Call ModCell("G396", "���ё�ՁE�Óc�G�u�N�E������Y�E���R�V")
    Call ModCell("C397", "900")
    Call ModCell("F397", "8")
    Call ModCell("G397", "���{�đ��E����A��E�����g�ځE������")

    Call EventChange(True)
    
    ' �u�b�N�̕ۑ�
    ActiveWorkbook.Save
End Sub

'
' �ߘa���N�x�̃v���O�����̋L�^������
'
Sub �w�}�L�^����()
    Sheets("�L�^���").Select
    Call SetRace(1, 1)
    Call SetLean(1, 3, "34709")
    Call SetLean(2, 4, "25977")
    Call SetLean(3, 6, "")
    Call SetLean(4, 8, "30973")
    Call �o�^
    Call ���ʌ���
    Call ������

    Call SetRace(5, 1)
    Call SetLean(1, 4, "32844")
    Call SetLean(2, 5, "30994")
    Call SetLean(3, 6, "24303")
    Call SetLean(4, 7, "25060")
    Call �o�^
    Call ���ʌ���
    Call ������

    Call SetRace(6, 1)
    Call SetLean(1, 5, "35775")
    Call SetLean(2, 6, "34009")
    Call SetLean(3, 7, "30945")
    Call �o�^
    Call ���ʌ���
    Call ������

    Call SetRace(7, 1)
    Call SetLean(1, 5, "40596")
    Call SetLean(2, 6, "34035")
    Call SetLean(3, 7, "25870")
    Call �o�^
    Call ���ʌ���
    Call ������

    Call SetRace(8, 1)
    Call SetLean(1, 3, "35996")
    Call SetLean(2, 4, "33734")
    Call SetLean(3, 5, "30026")
    Call SetLean(4, 6, "32161")
    Call SetLean(5, 7, "32447")
    Call SetLean(6, 8, "40131")
    Call SetLean(7, 9, "44509")
    Call �o�^
    Call ���ʌ���
    Call ������

    Call SetRace(9, 1)
    Call SetLean(2, 4, "33965")
    Call SetLean(3, 5, "30134")
    Call SetLean(4, 6, "25908")
    Call SetLean(5, 7, "33560")
    Call �o�^
    Call ���ʌ���
    Call ������

    Call SetRace(10, 1)
    Call SetLean(3, 5, "24066")
    Call SetLean(4, 6, "24885")
    Call SetLean(5, 7, "31742")
    Call �o�^
    Call ���ʌ���
    Call ������

    Call SetRace(11, 1)
    Call SetLean(3, 5, "30128")
    Call SetLean(4, 6, "24081")
    Call SetLean(5, 7, "32400")
    Call �o�^
    Call ���ʌ���
    Call ������

    Call SetRace(12, 1)
    Call SetLean(2, 4, "22203")
    Call SetLean(3, 5, "14942")
    Call SetLean(5, 7, "13107")
    Call �o�^
    Call ���ʌ���
    Call ������

    Call SetRace(15, 1)
    Call SetLean(3, 5, "15144")
    Call SetLean(4, 6, "14670")
    Call SetLean(5, 7, "21162")
    Call �o�^
    Call ���ʌ���
    Call ������

    Call SetRace(16, 1)
    Call SetLean(3, 5, "11638")
    Call SetLean(4, 6, "11011")
    Call SetLean(5, 7, "13879")
    Call �o�^
    Call ���ʌ���
    Call ������

    Call SetRace(18, 1)
    Call SetLean(2, 4, "20698")
    Call SetLean(3, 5, "15065")
    Call SetLean(4, 6, "12312")
    Call SetLean(5, 7, "14285")
    Call SetLean(6, 8, "12609")
    Call �o�^
    Call ���ʌ���
    Call ������

    Call SetRace(19, 1)
    Call SetLean(2, 4, "12738")
    Call SetLean(3, 5, "12811")
    Call SetLean(4, 6, "14208")
    Call SetLean(5, 7, "14290")
    Call �o�^
    Call ������

    Call SetRace(19, 2)
    Call SetLean(2, 4, "12566")
    Call SetLean(3, 5, "11229")
    Call SetLean(4, 6, "11493")
    Call SetLean(5, 7, "12115")
    Call �o�^
    Call ���ʌ���
    Call ������

    Call SetRace(20, 1)
    Call SetLean(2, 4, "14773")
    Call SetLean(3, 5, "")
    Call SetLean(4, 6, "11633")
    Call SetLean(5, 7, "13965")
    Call SetLean(6, 8, "")
    Call �o�^
    Call ���ʌ���
    Call ������

    Call SetRace(21, 1)
    Call SetLean(3, 5, "13404")
    Call SetLean(4, 6, "12150")
    Call SetLean(5, 7, "")
    Call �o�^
    Call ���ʌ���
    Call ������

    Call SetRace(22, 1)
    Call SetLean(2, 4, "11381")
    Call SetLean(3, 5, "11281")
    Call SetLean(4, 6, "")
    Call SetLean(5, 7, "11346")
    Call SetLean(6, 8, "13004")
    Call �o�^
    Call ���ʌ���
    Call ������

    Call SetRace(23, 1)
    Call SetLean(1, 3, "20131")
    Call SetLean(2, 4, "11678")
    Call SetLean(3, 5, "11109")
    Call SetLean(4, 6, "10461")
    Call SetLean(5, 7, "11666")
    Call SetLean(6, 8, "12853")
    Call �o�^
    Call ���ʌ���
    Call ������

    Call SetRace(24, 1)
    Call SetLean(3, 5, "20436")
    Call SetLean(4, 6, "14259")
    Call �o�^
    Call ���ʌ���
    Call ������

    Call SetRace(25, 1)
    Call SetLean(3, 5, "21903")
    Call SetLean(4, 6, "11985")
    Call �o�^
    Call ���ʌ���
    Call ������

    Call SetRace(26, 1)
    Call SetLean(1, 3, "12098")
    Call SetLean(3, 5, "12302")
    Call SetLean(4, 6, "11185")
    Call SetLean(5, 7, "11273")
    Call SetLean(6, 8, "12923")
    Call �o�^
    Call ���ʌ���
    Call ������

    Call SetRace(30, 1)
    Call SetLean(2, 4, "14856")
    Call SetLean(3, 5, "13937")
    Call SetLean(4, 6, "14220")
    Call SetLean(5, 7, "13409")
    Call �o�^
    Call ���ʌ���
    Call ������

    Call SetRace(31, 1)
    Call SetLean(1, 3, "21711")
    Call SetLean(2, 4, "15887")
    Call SetLean(3, 5, "15345")
    Call SetLean(4, 6, "15235")
    Call SetLean(5, 7, "14402")
    Call SetLean(6, 8, "13034")
    Call �o�^
    Call ���ʌ���
    Call ������

    Call SetRace(33, 1)
    Call SetLean(2, 4, "14304")
    Call SetLean(4, 6, "13373")
    Call SetLean(5, 7, "13388")
    Call �o�^
    Call ���ʌ���
    Call ������

    Call SetRace(35, 1)
    Call SetLean(3, 5, "13812")
    Call SetLean(4, 6, "13189")
    Call SetLean(5, 7, "14818")
    Call �o�^
    Call ���ʌ���
    Call ������

    Call SetRace(36, 1)
    Call SetLean(4, 6, "22462")
    Call �o�^
    Call ���ʌ���
    Call ������

    Call SetRace(37, 1)
    Call SetLean(1, 3, "11180")
    Call SetLean(2, 4, "5993")
    Call SetLean(3, 5, "5695")
    Call SetLean(4, 6, "5757")
    Call SetLean(5, 7, "4797")
    Call SetLean(6, 8, "5548")
    Call �o�^
    Call ������

    Call SetRace(37, 2)
    Call SetLean(2, 4, "4905")
    Call SetLean(3, 5, "4497")
    Call SetLean(4, 6, "4634")
    Call SetLean(5, 7, "4034")
    Call �o�^
    Call ���ʌ���
    Call ������

    Call SetRace(38, 1)
    Call SetLean(2, 4, "10412")
    Call SetLean(3, 5, "4200")
    Call SetLean(4, 6, "4949")
    Call SetLean(5, 7, "11151")
    Call �o�^
    Call ������

    Call SetRace(38, 2)
    Call SetLean(2, 4, "4274")
    Call SetLean(3, 5, "")
    Call SetLean(4, 6, "3067")
    Call SetLean(5, 7, "2844")
    Call �o�^
    Call ���ʌ���
    Call ������

    Call SetRace(39, 1)
    Call SetLean(1, 3, "10692")
    Call SetLean(2, 4, "5841")
    Call SetLean(3, 5, "5253")
    Call SetLean(4, 6, "")
    Call SetLean(5, 7, "5480")
    Call SetLean(6, 8, "5486")
    Call �o�^
    Call ���ʌ���
    Call ������

    Call SetRace(40, 1)
    Call SetLean(2, 4, "5817")
    Call SetLean(3, 5, "10205")
    Call SetLean(4, 6, "10429")
    Call SetLean(5, 7, "10879")
    Call SetLean(6, 8, "10224")
    Call �o�^
    Call ������

    Call SetRace(40, 2)
    Call SetLean(2, 4, "5227")
    Call SetLean(3, 5, "")
    Call SetLean(4, 6, "4270")
    Call SetLean(5, 7, "10319")
    Call SetLean(6, 8, "5011")
    Call �o�^
    Call ���ʌ���
    Call ������

    Call SetRace(41, 1)
    Call SetLean(2, 4, "11365")
    Call SetLean(3, 5, "10335")
    Call SetLean(4, 6, "10911")
    Call SetLean(5, 7, "10980")
    Call �o�^
    Call ������

    Call SetRace(41, 2)
    Call SetLean(2, 4, "5318")
    Call SetLean(3, 5, "4952")
    Call SetLean(4, 6, "4242")
    Call SetLean(5, 7, "4640")
    Call SetLean(6, 8, "5242")
    Call �o�^
    Call ���ʌ���
    Call ������

    Call SetRace(42, 1)
    Call SetLean(2, 4, "11682")
    Call SetLean(3, 5, "10808")
    Call SetLean(4, 6, "")
    Call SetLean(5, 7, "11071")
    Call �o�^
    Call ������

    Call SetRace(42, 2)
    Call SetLean(2, 4, "10134")
    Call SetLean(3, 5, "4025")
    Call SetLean(4, 6, "3967")
    Call SetLean(5, 7, "4625")
    Call SetLean(6, 8, "5993")
    Call �o�^
    Call ���ʌ���
    Call ������

    Call SetRace(43, 1)
    Call SetLean(1, 3, "12698")
    Call SetLean(2, 4, "5893")
    Call SetLean(3, 5, "5669")
    Call SetLean(4, 6, "5149")
    Call SetLean(5, 7, "5519")
    Call SetLean(6, 8, "11451")
    Call �o�^
    Call ������

    Call SetRace(43, 2)
    Call SetLean(1, 3, "5554")
    Call SetLean(2, 4, "4237")
    Call SetLean(3, 5, "3527")
    Call SetLean(4, 6, "3191")
    Call SetLean(5, 7, "")
    Call SetLean(6, 8, "4485")
    Call �o�^
    Call ���ʌ���
    Call ������

    Call SetRace(44, 1)
    Call SetLean(2, 4, "12903")
    Call SetLean(3, 5, "10568")
    Call SetLean(4, 6, "10494")
    Call SetLean(5, 7, "5811")
    Call SetLean(6, 8, "12655")
    Call �o�^
    Call ������

    Call SetRace(44, 2)
    Call SetLean(1, 3, "5530")
    Call SetLean(2, 4, "5019")
    Call SetLean(3, 5, "4340")
    Call SetLean(4, 6, "4377")
    Call SetLean(5, 7, "4793")
    Call SetLean(6, 8, "5455")
    Call �o�^
    Call ���ʌ���
    Call ������

    Call SetRace(45, 1)
    Call SetLean(1, 3, "5597")
    Call SetLean(2, 4, "5079")
    Call SetLean(3, 5, "5992")
    Call SetLean(4, 6, "4862")
    Call SetLean(5, 7, "4438")
    Call SetLean(6, 8, "5421")
    Call �o�^
    Call ������

    Call SetRace(45, 2)
    Call SetLean(1, 4, "5648")
    Call SetLean(2, 5, "4434")
    Call SetLean(3, 6, "4434")
    Call SetLean(4, 7, "3886")
    Call SetLean(5, 8, "4741")
    Call �o�^
    Call ������

    Call SetRace(45, 3)
    Call SetLean(1, 3, "4250")
    Call SetLean(2, 4, "4029")
    Call SetLean(3, 5, "3897")
    Call SetLean(4, 6, "3947")
    Call SetLean(5, 7, "3427")
    Call SetLean(6, 8, "3565")
    Call �o�^
    Call ���ʌ���
    Call ������

    Call SetRace(46, 1)
    Call SetLean(1, 3, "4472")
    Call SetLean(2, 4, "5380")
    Call SetLean(3, 5, "3590")
    Call SetLean(4, 6, "5410")
    Call SetLean(5, 7, "10543")
    Call SetLean(6, 8, "4431")
    Call SetLean(7, 9, "4593")
    Call �o�^
    Call ������

    Call SetRace(46, 2)
    Call SetLean(1, 3, "4019")
    Call SetLean(2, 4, "3704")
    Call SetLean(3, 5, "3372")
    Call SetLean(4, 6, "3945")
    Call SetLean(5, 7, "3970")
    Call SetLean(6, 8, "3329")
    Call SetLean(7, 9, "3722")
    Call �o�^
    Call ������

    Call SetRace(46, 3)
    Call SetLean(1, 3, "3399")
    Call SetLean(2, 4, "2927")
    Call SetLean(3, 5, "3089")
    Call SetLean(4, 6, "3083")
    Call SetLean(5, 7, "3185")
    Call SetLean(6, 8, "3200")
    Call SetLean(7, 9, "3143")
    Call �o�^
    Call ������

    Call SetRace(46, 4)
    Call SetLean(1, 3, "3415")
    Call SetLean(2, 4, "3109")
    Call SetLean(3, 5, "3184")
    Call SetLean(4, 6, "3151")
    Call SetLean(5, 7, "2655")
    Call SetLean(6, 8, "2665")
    Call SetLean(7, 9, "2599")
    Call �o�^
    Call ���ʌ���
    Call ������

    Call SetRace(47, 1)
    Call SetLean(1, 3, "5073")
    Call SetLean(2, 4, "4746")
    Call SetLean(3, 5, "4559")
    Call SetLean(4, 6, "4347")
    Call SetLean(5, 7, "4631")
    Call SetLean(6, 8, "5017")
    Call SetLean(7, 9, "5421")
    Call �o�^
    Call ���ʌ���
    Call ������

    Call SetRace(48, 1)
    Call SetLean(1, 3, "14051")
    Call SetLean(2, 4, "4753")
    Call SetLean(3, 5, "4608")
    Call SetLean(4, 6, "5210")
    Call SetLean(5, 7, "5825")
    Call SetLean(6, 8, "10298")
    Call �o�^
    Call ������

    Call SetRace(48, 2)
    Call SetLean(1, 3, "5616")
    Call SetLean(2, 4, "4663")
    Call SetLean(3, 5, "3714")
    Call SetLean(4, 6, "3510")
    Call SetLean(5, 7, "")
    Call SetLean(6, 8, "5373")
    Call SetLean(7, 9, "5920")
    Call �o�^
    Call ���ʌ���
    Call ������

    Call SetRace(49, 1)
    Call SetLean(2, 4, "10380")
    Call SetLean(3, 5, "10323")
    Call SetLean(4, 6, "5590")
    Call SetLean(5, 7, "5562")
    Call SetLean(6, 8, "10423")
    Call �o�^
    Call ������

    Call SetRace(49, 2)
    Call SetLean(1, 3, "5943")
    Call SetLean(2, 4, "4975")
    Call SetLean(3, 5, "4646")
    Call SetLean(4, 6, "4755")
    Call SetLean(5, 7, "4708")
    Call SetLean(6, 8, "4998")
    Call �o�^
    Call ������

    Call SetRace(49, 3)
    Call SetLean(1, 3, "4825")
    Call SetLean(2, 4, "4334")
    Call SetLean(3, 5, "3936")
    Call SetLean(4, 6, "3412")
    Call SetLean(5, 7, "4267")
    Call SetLean(6, 8, "4464")
    Call �o�^
    Call ���ʌ���
    Call ������

    Call SetRace(50, 1)
    Call SetLean(1, 3, "10313")
    Call SetLean(2, 4, "10799")
    Call SetLean(3, 5, "5809")
    Call SetLean(4, 6, "5885")
    Call SetLean(5, 7, "10165")
    Call SetLean(6, 8, "10802")
    Call �o�^
    Call ������

    Call SetRace(50, 2)
    Call SetLean(1, 3, "12327")
    Call SetLean(2, 4, "4970")
    Call SetLean(3, 5, "4615")
    Call SetLean(4, 6, "4176")
    Call SetLean(5, 7, "4655")
    Call SetLean(6, 8, "5041")
    Call �o�^
    Call ������

    Call SetRace(50, 3)
    Call SetLean(1, 3, "3851")
    Call SetLean(2, 4, "3920")
    Call SetLean(3, 5, "3565")
    Call SetLean(4, 6, "3642")
    Call SetLean(5, 7, "3610")
    Call SetLean(6, 8, "3814")
    Call �o�^
    Call ���ʌ���
    Call ������

    Call SetRace(51, 1)
    Call SetLean(2, 4, "")
    Call SetLean(3, 5, "4031")
    Call SetLean(4, 6, "4901")
    Call SetLean(5, 7, "5324")
    Call SetLean(6, 8, "11015")
    Call �o�^
    Call ������

    Call SetRace(51, 2)
    Call SetLean(2, 4, "4146")
    Call SetLean(3, 5, "4063")
    Call SetLean(4, 6, "3928")
    Call SetLean(5, 7, "4726")
    Call SetLean(6, 8, "4779")
    Call �o�^
    Call ������

    Call SetRace(51, 3)
    Call SetLean(2, 4, "3720")
    Call SetLean(3, 5, "3404")
    Call SetLean(4, 6, "3343")
    Call SetLean(5, 7, "3427")
    Call SetLean(6, 8, "3587")
    Call �o�^
    Call ���ʌ���
    Call ������

    Call SetRace(52, 1)
    Call SetLean(1, 3, "5502")
    Call SetLean(2, 4, "")
    Call SetLean(3, 5, "4846")
    Call SetLean(4, 6, "5145")
    Call SetLean(5, 7, "5834")
    Call SetLean(6, 8, "5675")
    Call �o�^
    Call ������

    Call SetRace(52, 2)
    Call SetLean(1, 3, "4510")
    Call SetLean(2, 4, "4263")
    Call SetLean(3, 5, "3766")
    Call SetLean(4, 6, "3863")
    Call SetLean(5, 7, "4033")
    Call SetLean(6, 8, "")
    Call SetLean(7, 9, "4648")
    Call �o�^
    Call ������

    Call SetRace(52, 3)
    Call SetLean(1, 3, "3341")
    Call SetLean(2, 4, "3562")
    Call SetLean(3, 5, "3365")
    Call SetLean(4, 6, "3217")
    Call SetLean(5, 7, "3354")
    Call SetLean(6, 8, "")
    Call SetLean(7, 9, "3336")
    Call �o�^
    Call ���ʌ���
    Call ������

    Call SetRace(53, 1)
    Call SetLean(3, 5, "11244")
    Call SetLean(4, 6, "4733")
    Call SetLean(5, 7, "11311")
    Call �o�^
    Call ���ʌ���
    Call ������

    Call SetRace(54, 1)
    Call SetLean(2, 4, "")
    Call SetLean(3, 5, "3956")
    Call SetLean(4, 6, "4566")
    Call SetLean(5, 7, "4251")
    Call SetLean(6, 8, "")
    Call �o�^
    Call ������

    Call SetRace(54, 2)
    Call SetLean(2, 4, "3220")
    Call SetLean(3, 5, "3207")
    Call SetLean(4, 6, "3557")
    Call SetLean(5, 7, "3294")
    Call SetLean(6, 8, "2760")
    Call �o�^
    Call ���ʌ���
    Call ������

    Call SetRace(55, 1)
    Call SetLean(3, 5, "5525")
    Call SetLean(4, 6, "4874")
    Call �o�^
    Call ���ʌ���
    Call ������

    Call SetRace(56, 1)
    Call SetLean(3, 5, "4025")
    Call SetLean(4, 6, "3930")
    Call SetLean(5, 7, "5488")
    Call �o�^
    Call ���ʌ���
    Call ������

    Call SetRace(57, 1)
    Call SetLean(3, 5, "4470")
    Call SetLean(4, 6, "5093")
    Call SetLean(5, 7, "10113")
    Call �o�^
    Call ���ʌ���
    Call ������
    
    Call SetRace(58, 1)
    Call SetLean(3, 5, "4307")
    Call SetLean(4, 6, "4481")
    Call SetLean(5, 7, "4518")
    Call �o�^
    Call ���ʌ���
    Call ������
    
    Call SetRace(59, 1)
    Call SetLean(1, 3, "4370")
    Call SetLean(2, 4, "3781")
    Call SetLean(3, 5, "3298")
    Call SetLean(4, 6, "3104")
    Call SetLean(5, 7, "3725")
    Call SetLean(6, 8, "4731")
    Call SetLean(7, 9, "5361")
    Call �o�^
    Call ���ʌ���
    Call ������
    
    Call SetRace(60, 1)
    Call SetLean(4, 6, "10957")
    Call �o�^
    Call ���ʌ���
    Call ������
    
    Call SetRace(61, 1)
    Call SetLean(2, 4, "10738")
    Call SetLean(3, 5, "5395")
    Call SetLean(4, 6, "5182")
    Call SetLean(5, 7, "5092")
    Call SetLean(6, 8, "5730")
    Call �o�^
    Call ������
    
    Call SetRace(61, 2)
    Call SetLean(2, 4, "5160")
    Call SetLean(3, 5, "4732")
    Call SetLean(4, 6, "4920")
    Call SetLean(5, 7, "4264")
    Call �o�^
    Call ���ʌ���
    Call ������
    
    Call SetRace(62, 1)
    Call SetLean(2, 4, "5005")
    Call SetLean(3, 5, "4947")
    Call SetLean(4, 6, "4361")
    Call SetLean(5, 7, "")
    Call �o�^
    Call ������
    
    Call SetRace(62, 2)
    Call SetLean(2, 4, "4687")
    Call SetLean(3, 5, "5207")
    Call SetLean(4, 6, "4637")
    Call SetLean(5, 7, "")
    Call �o�^
    Call ���ʌ���
    Call ������
    
    Call SetRace(63, 1)
    Call SetLean(3, 5, "10867")
    Call SetLean(4, 6, "5831")
    Call �o�^
    Call ���ʌ���
    Call ������
    
    Call SetRace(64, 1)
    Call SetLean(2, 4, "11389")
    Call SetLean(3, 5, "")
    Call SetLean(4, 6, "4756")
    Call SetLean(5, 7, "")
    Call SetLean(6, 8, "10765")
    Call �o�^
    Call ���ʌ���
    Call ������
    
    Call SetRace(65, 1)
    Call SetLean(1, 3, "11373")
    Call SetLean(2, 4, "10850")
    Call SetLean(3, 5, "5412")
    Call SetLean(4, 6, "5646")
    Call SetLean(5, 7, "10492")
    Call SetLean(6, 8, "11011")
    Call �o�^
    Call ���ʌ���
    Call ������
    
    Call SetRace(66, 1)
    Call SetLean(2, 4, "")
    Call SetLean(3, 5, "5590")
    Call SetLean(4, 6, "5300")
    Call SetLean(5, 7, "10270")
    Call �o�^
    Call ���ʌ���
    Call ������
    
    Call SetRace(67, 1)
    Call SetLean(3, 5, "10237")
    Call SetLean(4, 6, "4312")
    Call SetLean(5, 7, "5075")
    Call �o�^
    Call ���ʌ���
    Call ������

    Call SetRace(68, 1)
    Call SetLean(1, 3, "4509")
    Call SetLean(2, 4, "5073")
    Call SetLean(3, 5, "4589")
    Call SetLean(4, 6, "4222")
    Call SetLean(5, 7, "")
    Call SetLean(6, 8, "4983")
    Call SetLean(7, 9, "5973")
    Call �o�^
    Call ���ʌ���
    Call ������

    Call SetRace(69, 1)
    Call SetLean(2, 4, "32483")
    Call SetLean(3, 5, "24343")
    Call SetLean(4, 6, "")
    Call SetLean(5, 7, "23027")
    Call �o�^
    Call ���ʌ���
    Call ������

    Call SetRace(70, 1)
    Call SetLean(3, 5, "23512")
    Call SetLean(4, 6, "25694")
    Call SetLean(5, 7, "21541")
    Call �o�^
    Call ���ʌ���
    Call ������

    Call SetRace(71, 1)
    Call SetLean(1, 3, "35138")
    Call SetLean(2, 4, "33887")
    Call SetLean(4, 6, "42363")
    Call SetLean(5, 7, "25021")
    Call SetLean(6, 8, "32220")
    Call �o�^
    Call ���ʌ���
    Call ������
    
    ' �u�b�N�̕ۑ�
    ActiveWorkbook.Save
End Sub

