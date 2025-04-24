Attribute VB_Name = "M_MacrosHelper"
Option Explicit
#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

' �O���[�o���ϐ�
Dim GlobalToast As Toast
'******************************
'Toast
'******************************

Sub ShowToast(msg As String)
    Dim i As Double
    Dim startPos As Double
    Dim midPos As Double
    Dim endPos As Double
    Dim duration As Double
    Dim halfway As Double
    Dim currentTime As Double
    Dim changeInPosition As Double

    ' ������Toast������΃A�����[�h����
    CloseToast
    
    Set GlobalToast = New Toast
    With GlobalToast
        .messageText = msg
        .StartUpPosition = 0
        .Left = Application.Left + (Application.Width / 2) - (.Width / 2)
        ' �g�[�X�g�̏����ʒu�͉�ʂ̉������炳���10�s�N�Z�����ɐݒ�
        startPos = Application.Top + Application.Height - .Height - 50
        ' �g�[�X�g���ŏ��Ɉړ����钆�Ԉʒu�́A�����ʒu���10�s�N�Z����ɐݒ�
        midPos = startPos - 10
        ' �g�[�X�g�̏I���ʒu�́A�����ʒu���炳���100�s�N�Z�����ɐݒ�
        endPos = startPos + 100
        duration = 100 ' �A�j���[�V�����̑����Ԃ�ݒ�
        halfway = duration / 2 ' �A�j���[�V�����̒��ԓ_��ݒ�

        ' UserForm��\��
        .Show vbModeless

        On Error Resume Next ' �G���[�n���h�����O��L����

        ' �A�j���[�V�������[�v
        For i = 0 To duration
            If Not .Visible Then Exit For ' UserForm�������Ă����烋�[�v���I��
            If Err.Number <> 0 Then Exit For ' ���炩�̃G���[�����������烋�[�v���I��
            currentTime = i / duration
            If i <= halfway Then
                ' ���ԓ_�܂ł̏㏸�A�j���[�V����
                .Top = startPos - (startPos - midPos) * EaseOut(currentTime * 2)
            Else
                ' ���ԓ_����̉��~�A�j���[�V����
                .Top = midPos + (endPos - midPos) * EaseIn((currentTime - 0.5) * 2)
            End If
            DoEvents
            Sleep 15
        Next i
        On Error GoTo 0 ' �G���[�n���h�����O��ʏ탂�[�h�ɖ߂�
    End With
End Sub

Function EaseOut(t As Double) As Double
    ' �C�[�W���O�֐��iEaseOut Cubic�j
    t = t - 1
    EaseOut = (t * t * t + 1)
End Function

Function EaseIn(t As Double) As Double
    ' �C�[�W���O�֐��iEaseIn Cubic�j
    EaseIn = t * t * t
End Function

Sub CloseToast()
    ' ���ׂĂ̊J���Ă���Toast�����
    Dim frm As Object
    For Each frm In VBA.UserForms
        If TypeName(frm) = "Toast" Then
            Unload frm
        End If
    Next frm
    Set GlobalToast = Nothing
End Sub

