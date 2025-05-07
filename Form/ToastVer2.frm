VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ToastVer2 
   Caption         =   "ToastVer2"
   ClientHeight    =   1440
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4800
   OleObjectBlob   =   "ToastVer2.frx":0000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   BorderStyle     =   0  'None
End
Attribute VB_Name = "ToastVer2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ===================================================
' �g�[�X�g�ʒmV2 - �t�H�[��
' ===================================================
' �쐬��: 2025/04/25
' �쐬��: Claude
' �T�v: ���ǔŃg�[�X�g�ʒm�p�̃��[�U�[�t�H�[��
' �@�\:
'   - ���炩�ȃA�j���[�V�����\��/��\��
'   - ��ޕʂ̐F�����i���/����/�x��/�G���[�j
'   - �A�C�R���\���Ή�
'   - �N���b�N�܂��̓L�[���͂ŕ���@�\
' ===================================================

' �v���C�x�[�g�ϐ�
Private m_Message As String     ' �\�����郁�b�Z�[�W
Private m_ToastType As String   ' �ʒm�̎�ށi"info", "success", "warning", "error"�j
Private m_Duration As Integer   ' �\�����ԁi�~���b�j
Private m_IconType As String    ' �A�C�R���^�C�v
Private m_AnimationActive As Boolean ' �A�j���[�V�����i�s���t���O

' �A�j���[�V�����֘A�萔
Private Const ANIM_STEP_COUNT As Integer = 20   ' �A�j���[�V�����̃X�e�b�v��
Private Const ANIM_STEP_DELAY As Integer = 15   ' �X�e�b�v�Ԃ̒x���i�~���b�j

' ===================================================
' �v���p�e�B
' ===================================================

' ���b�Z�[�W�v���p�e�B
Public Property Let Message(value As String)
    m_Message = value
    If Me.Visible Then lblMessage.Caption = m_Message
End Property

Public Property Get Message() As String
    Message = m_Message
End Property

' �g�[�X�g��ރv���p�e�B
Public Property Let ToastType(value As String)
    m_ToastType = value
    ApplyToastStyle
End Property

Public Property Get ToastType() As String
    ToastType = m_ToastType
End Property

' �\�����ԃv���p�e�B
Public Property Let Duration(value As Integer)
    m_Duration = value
End Property

Public Property Get Duration() As Integer
    Duration = m_Duration
End Property

' �A�C�R���^�C�v�v���p�e�B
Public Property Let IconType(value As String)
    m_IconType = value
    ApplyIconStyle
End Property

Public Property Get IconType() As String
    IconType = m_IconType
End Property

' ===================================================
' �t�H�[���C�x���g
' ===================================================

' �t�H�[����������
Private Sub UserForm_Initialize()
    ' �f�t�H���g�l�̐ݒ�
    m_ToastType = "info"
    m_Duration = 2000
    m_AnimationActive = False
    
    ' �t�H�[���̏����ݒ�
    Me.BackColor = RGB(68, 68, 68)  ' �f�t�H���g�̔w�i�F
    Me.Width = 250
    Me.Height = 60
    
    ' �p�ە\���̂��߂̃t���[���ݒ�
    Frame1.BackColor = Me.BackColor
    Frame1.BorderStyle = 0
    
    ' ���x���̏����ݒ�
    lblMessage.BackColor = Me.BackColor
    lblMessage.ForeColor = RGB(255, 255, 255)
    lblIcon.BackColor = Me.BackColor
    lblIcon.ForeColor = RGB(255, 255, 255)
End Sub

' �t�H�[���\����
Private Sub UserForm_Activate()
    ' ���b�Z�[�W�̐ݒ�
    lblMessage.Caption = m_Message
    
    ' �X�^�C���̓K�p
    ApplyToastStyle
    ApplyIconStyle
    
    ' �t�H�[���̈ʒu�ݒ�
    PositionForm
    
    ' �X���C�h�C���A�j���[�V�����J�n
    SlideInAnimation
    
    ' �\���^�C�}�[�̐ݒ�
    Application.OnTime Now + TimeSerial(0, 0, m_Duration / 1000), "ToastManagerVer2.ContinueToastQueue"
End Sub

' �N���b�N�ŕ���
Private Sub UserForm_Click()
    CloseToast
End Sub

' ���x���N���b�N�ł�����
Private Sub lblMessage_Click()
    CloseToast
End Sub

' �A�C�R�����x���N���b�N�ł�����
Private Sub lblIcon_Click()
    CloseToast
End Sub

' �L�[���͂ł�����
Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    CloseToast
End Sub

' ===================================================
' ���\�b�h
' ===================================================

' �g�[�X�g��ނɉ������X�^�C����K�p
Private Sub ApplyToastStyle()
    Dim backColor As Long
    
    ' �g�[�X�g��ނɉ������w�i�F�̐ݒ�
    Select Case LCase(m_ToastType)
        Case "info"
            backColor = RGB(0, 120, 215)  ' �n
        Case "success"
            backColor = RGB(16, 124, 16)  ' �Όn
        Case "warning"
            backColor = RGB(234, 157, 35) ' �I�����W�n
        Case "error"
            backColor = RGB(232, 17, 35)  ' �Ԍn
        Case Else
            backColor = RGB(68, 68, 68)   ' �O���[�i�f�t�H���g�j
    End Select
    
    ' �w�i�F�̓K�p
    Me.BackColor = backColor
    Frame1.BackColor = backColor
    lblMessage.BackColor = backColor
    lblIcon.BackColor = backColor
End Sub

' �A�C�R����ނɉ������X�^�C����K�p
Private Sub ApplyIconStyle()
    ' �A�C�R�������̐ݒ�
    ' �����ł�Unicode�̋L�����g�p (PowerPoint���̃A�C�R��)
    Select Case LCase(m_IconType)
        Case "info", ""
            lblIcon.Caption = "i"
            lblIcon.Font.Bold = True
        Case "success"
            lblIcon.Caption = "?"
        Case "warning"
            lblIcon.Caption = "!"
        Case "error"
            lblIcon.Caption = "�~"
        Case Else
            lblIcon.Caption = ""
    End Select
End Sub

' �t�H�[���̈ʒu��ݒ�
Private Sub PositionForm()
    ' �A�N�e�B�u�E�B���h�E�̒��������ɕ\��
    ' Excel�̃A�v���P�[�V�����E�B���h�E�ł͂Ȃ��A�N�e�B�u��Excel�E�B���h�E�ɑ΂��Ĉʒu����
    Dim activeWin As Window
    Set activeWin = Application.ActiveWindow
    
    Dim winLeft As Double
    Dim winTop As Double
    Dim winWidth As Double
    Dim winHeight As Double
    
    ' �A�N�e�B�u�E�B���h�E�̈ʒu�ƃT�C�Y���擾
    winLeft = activeWin.WindowLeft
    winTop = activeWin.WindowTop
    winWidth = activeWin.Width
    winHeight = activeWin.Height
    
    ' �t�H�[���̈ʒu��ݒ�i���������j
    Me.Left = winLeft + (winWidth / 2) - (Me.Width / 2)
    Me.Top = winTop + winHeight  ' ��ʉ��i�X���C�h�C���̊J�n�ʒu�j
End Sub

' �X���C�h�C���A�j���[�V����
Private Sub SlideInAnimation()
    m_AnimationActive = True
    
    Dim i As Integer
    Dim startTop As Double
    Dim endTop As Double
    Dim currentTop As Double
    
    ' �J�n�ʒu�ƏI���ʒu�̐ݒ�
    startTop = Me.Top
    endTop = startTop - Me.Height - 10  ' 10�s�N�Z���]�T����������
    
    ' �A�j���[�V�������[�v
    For i = 0 To ANIM_STEP_COUNT
        ' �C�[�W���O�֐����g�p�������炩�ȓ���
        currentTop = startTop - (startTop - endTop) * EaseOutCubic(i / ANIM_STEP_COUNT)
        Me.Top = currentTop
        
        ' ��ʂ̍X�V�ƒx��
        DoEvents
        Sleep ANIM_STEP_DELAY
    Next i
    
    m_AnimationActive = False
End Sub

' �X���C�h�A�E�g�A�j���[�V����
Private Sub SlideOutAnimation()
    If m_AnimationActive Then Exit Sub
    m_AnimationActive = True
    
    Dim i As Integer
    Dim startTop As Double
    Dim endTop As Double
    Dim currentTop As Double
    
    ' �J�n�ʒu�ƏI���ʒu�̐ݒ�
    startTop = Me.Top
    endTop = startTop + Me.Height + 10  ' 10�s�N�Z���]�T����������
    
    ' �A�j���[�V�������[�v
    For i = 0 To ANIM_STEP_COUNT
        ' �C�[�W���O�֐����g�p�������炩�ȓ���
        currentTop = startTop + (endTop - startTop) * EaseInCubic(i / ANIM_STEP_COUNT)
        Me.Top = currentTop
        
        ' ��ʂ̍X�V�ƒx��
        DoEvents
        Sleep ANIM_STEP_DELAY
    Next i
    
    m_AnimationActive = False
    Me.Hide
End Sub

' �g�[�X�g�����
Private Sub CloseToast()
    ' �A�j���[�V�������Ȃ牽�����Ȃ�
    If m_AnimationActive Then Exit Sub
    
    ' �^�C�}�[���L�����Z��
    On Error Resume Next
    Application.OnTime Now + TimeSerial(0, 0, m_Duration / 1000), "ToastManagerVer2.ContinueToastQueue", , False
    On Error GoTo 0
    
    ' �X���C�h�A�E�g�A�j���[�V����
    SlideOutAnimation
    
    ' �}�l�[�W���[�ɒʒm
    ToastManagerVer2.CloseCurrentToast
End Sub

' ===================================================
' �C�[�W���O�֐�
' ===================================================

' �C�[�W���O�C���i�����j
Private Function EaseInCubic(t As Double) As Double
    EaseInCubic = t * t * t
End Function

' �C�[�W���O�A�E�g�i�����j
Private Function EaseOutCubic(t As Double) As Double
    t = t - 1
    EaseOutCubic = t * t * t + 1
End Function

' ===================================================
' ���[�e�B���e�B�֐�
' ===================================================

' �X���[�v�֐�
Private Sub Sleep(milliseconds As Long)
    Dim startTime As Double
    startTime = Timer
    
    Do While Timer < startTime + (milliseconds / 1000)
        DoEvents
    Loop
End Sub
