Attribute VB_Name = "M_ToastVer2"
Option Explicit

' ===================================================
' �g�[�X�g�ʒmV2 - �w���p�[���W���[��
' ===================================================
' �쐬��: 2025/04/25
' �쐬��: Claude
' �T�v: ���ǔŃg�[�X�g�ʒm�̃w���p�[�֐���񋟂��郂�W���[��
' �@�\:
'   - �ȈՌĂяo���p�̃w���p�[�֐�
'   - API�֘A�̐錾
'   - �g���C�[�W���O�֐�
'   - ���[�e�B���e�B�֐�
' ===================================================

' Windows API�錾
#If VBA7 Then
    ' 64�r�b�g Office�p
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    ' 32�r�b�g Office�p
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

' ===================================================
' ���J�w���p�[�֐�
' ===================================================

' ��{�I�ȃg�[�X�g�ʒm��\���i�ȈՃo�[�W�����j
Public Sub ShowToastVer2(Message As String, Optional ToastType As String = "info", _
                       Optional Duration As Integer = 0, Optional IconType As String = "")
    ToastManagerVer2.ShowToast Message, ToastType, Duration, IconType
End Sub

' ���g�[�X�g�ʒm�i�F�j
Public Sub ShowInfoToast(Message As String, Optional Duration As Integer = 0)
    ToastManagerVer2.ShowToast Message, "info", Duration, "info"
End Sub

' �����g�[�X�g�ʒm�i�ΐF�j
Public Sub ShowSuccessToast(Message As String, Optional Duration As Integer = 0)
    ToastManagerVer2.ShowToast Message, "success", Duration, "success"
End Sub

' �x���g�[�X�g�ʒm�i�I�����W�F�j
Public Sub ShowWarningToast(Message As String, Optional Duration As Integer = 0)
    ToastManagerVer2.ShowToast Message, "warning", Duration, "warning"
End Sub

' �G���[�g�[�X�g�ʒm�i�ԐF�j
Public Sub ShowErrorToast(Message As String, Optional Duration As Integer = 0)
    ToastManagerVer2.ShowToast Message, "error", Duration, "error"
End Sub

' ===================================================
' �g���C�[�W���O�֐�
' ===================================================

' �l�X�ȃC�[�W���O�֐���񋟁i�A�j���[�V�����̓������J�X�^�}�C�Y���邽�߁j
' t: 0.0�`1.0�̒l�i�i�s�x�j

' ���`�i��葬�x�j
Public Function LinearEasing(t As Double) As Double
    LinearEasing = t
End Function

' �񎟊֐��C�[�W���O�C���i�����j
Public Function EaseInQuad(t As Double) As Double
    EaseInQuad = t * t
End Function

' �񎟊֐��C�[�W���O�A�E�g�i�����j
Public Function EaseOutQuad(t As Double) As Double
    EaseOutQuad = -t * (t - 2)
End Function

' �񎟊֐��C�[�W���O�C���A�E�g�i����-�����j
Public Function EaseInOutQuad(t As Double) As Double
    t = t * 2
    If t < 1 Then
        EaseInOutQuad = 0.5 * t * t
    Else
        t = t - 1
        EaseInOutQuad = -0.5 * (t * (t - 2) - 1)
    End If
End Function

' �O���֐��C�[�W���O�C���i�����j
Public Function EaseInCubic(t As Double) As Double
    EaseInCubic = t * t * t
End Function

' �O���֐��C�[�W���O�A�E�g�i�����j
Public Function EaseOutCubic(t As Double) As Double
    t = t - 1
    EaseOutCubic = t * t * t + 1
End Function

' �O���֐��C�[�W���O�C���A�E�g�i����-�����j
Public Function EaseInOutCubic(t As Double) As Double
    t = t * 2
    If t < 1 Then
        EaseInOutCubic = 0.5 * t * t * t
    Else
        t = t - 2
        EaseInOutCubic = 0.5 * (t * t * t + 2)
    End If
End Function

' �e���C�[�W���O�A�E�g�i�o�E���h���ʁj
Public Function EaseOutElastic(t As Double) As Double
    Dim p As Double
    p = 0.3
    
    If t = 0 Then
        EaseOutElastic = 0
        Exit Function
    End If
    
    If t = 1 Then
        EaseOutElastic = 1
        Exit Function
    End If
    
    Dim s As Double
    s = p / 4
    
    EaseOutElastic = 2 ^ (-10 * t) * Sin((t - s) * (2 * WorksheetFunction.Pi) / p) + 1
End Function

' �o�E���X�C�[�W���O�A�E�g�i���˕Ԃ���ʁj
Public Function EaseOutBounce(t As Double) As Double
    If t < (1 / 2.75) Then
        EaseOutBounce = 7.5625 * t * t
    ElseIf t < (2 / 2.75) Then
        t = t - (1.5 / 2.75)
        EaseOutBounce = 7.5625 * t * t + 0.75
    ElseIf t < (2.5 / 2.75) Then
        t = t - (2.25 / 2.75)
        EaseOutBounce = 7.5625 * t * t + 0.9375
    Else
        t = t - (2.625 / 2.75)
        EaseOutBounce = 7.5625 * t * t + 0.984375
    End If
End Function

' ===================================================
' ���[�e�B���e�B�֐�
' ===================================================

' ��ʂ̍X�V���ꎞ��~/�ĊJ����֐�
Public Sub FreezeDraw(Optional freeze As Boolean = True)
    If freeze Then
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.Calculation = xlCalculationManual
    Else
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.Calculation = xlCalculationAutomatic
    End If
End Sub

' ���݂̃L���[���g�[�X�g�����擾
Public Function GetToastQueueSize() As Integer
    GetToastQueueSize = ToastManagerVer2.QueueSize
End Function

' ���ݏ��������ǂ������擾
Public Function IsProcessingToasts() As Boolean
    IsProcessingToasts = ToastManagerVer2.IsProcessing
End Function

' �g�[�X�g�L���[���N���A
Public Sub ClearToastQueue()
    ToastManagerVer2.ClearQueue
End Sub

' ===================================================
' �g�p��
' ===================================================
' ��{�I�Ȏg�p���@:
'  ShowToastVer2 "�������������܂���", "success", 3000
'
' ��p�֐��̎g�p��:
'  ShowSuccessToast "�ۑ����܂���", 3000
'  ShowErrorToast "�G���[���������܂���", 5000
'  ShowInfoToast "��񃁃b�Z�[�W"
'  ShowWarningToast "���ӂ��K�v�ł�"
'
' �����̒ʒm��A�����ĕ\��:
'  ShowInfoToast "�������J�n���܂�"
'  ' ���炩�̏���
'  ShowSuccessToast "�������������܂���"
'
' �L���[�̃N���A:
'  ClearToastQueue
' ===================================================
