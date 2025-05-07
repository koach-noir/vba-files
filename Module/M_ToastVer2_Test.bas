Attribute VB_Name = "M_ToastVer2_Test"
Option Explicit

' ===================================================
' �g�[�X�g�ʒmV2 - �e�X�g���W���[��
' ===================================================
' �쐬��: 2025/04/25
' �쐬��: Claude
' �T�v: ���ǔŃg�[�X�g�ʒm���e�X�g���邽�߂̃��W���[��
' �@�\:
'   - �e��g�[�X�g�ʒm�̕\���e�X�g
'   - �A���\���̃e�X�g
'   - �J�X�^�}�C�Y�̃f�����X�g���[�V����
' ===================================================

' ��{�I�ȃg�[�X�g�ʒm�̃e�X�g
Public Sub TestBasicToast()
    ' ��{�I�ȏ��g�[�X�g
    ShowToastVer2 "��{�I�ȃg�[�X�g�ʒm�ł�", "info", 2000
End Sub

' �e��g�[�X�g��ނ̃e�X�g
Public Sub TestAllToastTypes()
    ' ���g�[�X�g
    ShowInfoToast "��񃁃b�Z�[�W�ł�", 2000
    
    ' �����ҋ@
    Application.Wait Now + TimeSerial(0, 0, 3)
    
    ' �����g�[�X�g
    ShowSuccessToast "�������������܂���", 2000
    
    ' �����ҋ@
    Application.Wait Now + TimeSerial(0, 0, 3)
    
    ' �x���g�[�X�g
    ShowWarningToast "���ӂ��K�v�ł�", 2000
    
    ' �����ҋ@
    Application.Wait Now + TimeSerial(0, 0, 3)
    
    ' �G���[�g�[�X�g
    ShowErrorToast "�G���[���������܂���", 2000
End Sub

' �A���g�[�X�g�\���̃e�X�g�i�L���[�@�\�j
Public Sub TestToastQueue()
    ' �����̃g�[�X�g��A�����ĕ\��
    ShowInfoToast "�������J�n���܂�"
    ShowInfoToast "�f�[�^��ǂݍ���ł��܂�..."
    ShowInfoToast "�v�Z��..."
    ShowSuccessToast "�������������܂���"
End Sub

' �������b�Z�[�W�̃e�X�g
Public Sub TestLongMessage()
    ' �������b�Z�[�W�̃g�[�X�g
    ShowInfoToast "����͒������b�Z�[�W�̃e�X�g�ł��B�g�[�X�g�ʒm�͒������b�Z�[�W�ł��K�؂ɕ\�������悤�ɐ݌v����Ă��܂��B�K�v�ɉ����ă��b�Z�[�W�̒����ɍ��킹�ăt�H�[���T�C�Y����������܂��B", 4000
End Sub

' �\�����Ԃ̃J�X�^�}�C�Y�e�X�g
Public Sub TestCustomDuration()
    ' �Z���\������
    ShowInfoToast "�Z���\���i1�b�j", 1000
    
    ' �����ҋ@
    Application.Wait Now + TimeSerial(0, 0, 2)
    
    ' �W���̕\������
    ShowInfoToast "�W���̕\���i2�b�j", 2000
    
    ' �����ҋ@
    Application.Wait Now + TimeSerial(0, 0, 3)
    
    ' �����\������
    ShowInfoToast "�����\���i5�b�j", 5000
End Sub

' ���ׂẴe�X�g�����s
Public Sub RunAllTests()
    ' �X�̃e�X�g�����ԂɎ��s
    TestBasicToast
    
    ' �����ҋ@
    Application.Wait Now + TimeSerial(0, 0, 3)
    
    TestAllToastTypes
    
    ' �����ҋ@
    Application.Wait Now + TimeSerial(0, 0, 12)
    
    TestToastQueue
    
    ' �����ҋ@
    Application.Wait Now + TimeSerial(0, 0, 8)
    
    TestLongMessage
    
    ' �����ҋ@
    Application.Wait Now + TimeSerial(0, 0, 5)
    
    TestCustomDuration
    
    ' �e�X�g�������b�Z�[�W
    MsgBox "�g�[�X�g�ʒmV2�̂��ׂẴe�X�g���������܂����B", vbInformation, "�e�X�g����"
End Sub

' ���ۂ̎g�p��̃f�����X�g���[�V����
Public Sub DemonstrateRealUsage()
    ' �����J�n�̒ʒm
    ShowInfoToast "�f�[�^���͂��J�n���܂�..."
    
    ' ���ۂ̏�����͋[�i5�b�ԑҋ@�j
    Application.Wait Now + TimeSerial(0, 0, 2)
    
    ' ���ԕ�
    ShowInfoToast "50���̃f�[�^��������..."
    
    ' ����ɏ�����͋[�i3�b�ԑҋ@�j
    Application.Wait Now + TimeSerial(0, 0, 2)
    
    ' ���������̒ʒm
    ShowSuccessToast "�f�[�^���͂��������܂����I", 3000
End Sub

' �g�[�X�g�L���[�N���A�̃e�X�g
Public Sub TestClearQueue()
    ' �������̃g�[�X�g���L���[�ɒǉ�
    ShowInfoToast "�g�[�X�g1"
    ShowInfoToast "�g�[�X�g2"
    ShowInfoToast "�g�[�X�g3"
    
    ' 1�b�ҋ@���Ă���L���[���N���A
    Application.Wait Now + TimeSerial(0, 0, 1)
    
    ' �L���[���N���A
    ClearToastQueue
    
    ' �N���A��̒ʒm
    ShowWarningToast "�L���[���N���A����܂���", 3000
End Sub

' �J�X�^���ݒ�Ńg�[�X�g��\������e�X�g
Public Sub TestCustomToast()
    ' �J�X�^���ݒ�Ńg�[�X�g��\��
    ShowToastVer2 "�J�X�^���ݒ�̃g�[�X�g", "success", 3000, "info"
End Sub
