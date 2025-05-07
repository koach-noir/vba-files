Attribute VB_Name = "M_MenuBuilder"
Option Explicit

' ���j���[���̒萔
Const CUSTOM_MENU_NAME As String = "EUMControlsMenu"
Const SETTINGS_FILE_PATH As String = "vba-files\Module\EUMMenuSettings.txt"

' �Z�N�V�������̒萔
Const SECTION_INDIVIDUAL_BUTTONS As String = "[IndividualButtons]"
Const SECTION_DROPDOWN_1 As String = "[DropDownList&Buttons1]"
Const SECTION_DROPDOWN_2 As String = "[DropDownList&Buttons2]"
Const SECTION_DROPDOWN_3 As String = "[DropDownList&Buttons3]"

' EUM�T�t�B�b�N�X
Const EUM_SUFFIX As String = "_EUM"

' �^�[�Q�b�g�Ƃ��郂�W���[�����̃��X�g
Private TargetModules As Variant

' ��ӂ�ID�����p�̕ϐ�
Private controlIdCounter As Long

' ���j���[�\���p�̃R���N�V����
Private individualButtons As Collection
Private dropdownList1 As Collection
Private dropdownList2 As Collection
Private dropdownList3 As Collection

' �}�N������ۑ����邽�߂̃R���N�V�����i�L�[=�\����, �l=�}�N�����j
Private displayToMacroMap1 As Object
Private displayToMacroMap2 As Object
Private displayToMacroMap3 As Object
Private displayToMacroMap As Object

' �V���[�g�J�b�g�L�[�̃��X�g
Private shortcutKeysList() As String
Private currentShortcutKeyIndex As Integer

Sub InitializeModule()
    TargetModules = Array("M_Macros") ' �K�v�ɉ����đΏۃ��W���[����ǉ�
    
    ' �e�R���N�V�����̏�����
    Set individualButtons = New Collection
    Set dropdownList1 = New Collection
    Set dropdownList2 = New Collection
    Set dropdownList3 = New Collection
    
    ' �}�N�����}�b�s���O�̏�����
    Set displayToMacroMap1 = CreateObject("Scripting.Dictionary")
    Set displayToMacroMap2 = CreateObject("Scripting.Dictionary")
    Set displayToMacroMap3 = CreateObject("Scripting.Dictionary")
    Set displayToMacroMap = CreateObject("Scripting.Dictionary") ' ����݊����̂���
    
    ' �V���[�g�J�b�g�L�[�̏�����
    shortcutKeysList = Split("Q,W,E,R,T,Y,U,I,O,A,S,D,F,G,H,J,K,L", ",")
    currentShortcutKeyIndex = 0
    
    ' �ݒ�t�@�C�����烁�j���[���ڂ�ǂݍ���
    LoadMenuSettingsFromFile
End Sub

' �V���[�g�J�b�g�L�[���L���v�V�����ɒǉ�����֐�
Function AssignShortcutKey(caption As String) As String
    ' �V���[�g�J�b�g�L�[�����蓖�Ă�
    If currentShortcutKeyIndex <= UBound(shortcutKeysList) Then
        ' �V���[�g�J�b�g�L�[��ǉ�
        AssignShortcutKey = caption & "(&" & shortcutKeysList(currentShortcutKeyIndex) & ")"
        ' �C���f�b�N�X�𑝂₷
        currentShortcutKeyIndex = currentShortcutKeyIndex + 1
    Else
        ' �V���[�g�J�b�g�L�[������Ȃ��ꍇ�͂��̂܂ܕԂ�
        AssignShortcutKey = caption
    End If
End Function

' �}�N��������\���p�̖��O���擾����֐�
Function GetDisplayName(macroName As String) As String
    ' �}�N�����̖�����_EUM������ꍇ�͏Ȃ�
    If Right(macroName, Len(EUM_SUFFIX)) = EUM_SUFFIX Then
        GetDisplayName = Left(macroName, Len(macroName) - Len(EUM_SUFFIX))
    Else
        GetDisplayName = macroName
    End If
End Function

' �ݒ�t�@�C�����烁�j���[�ݒ��ǂݍ���
Sub LoadMenuSettingsFromFile()
    Dim fso As Object
    Dim textFile As Object
    Dim filePath As String
    Dim textLine As String
    Dim currentSection As String
    
    ' �t�@�C���V�X�e���I�u�W�F�N�g�̍쐬
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' �t�@�C���p�X�̐ݒ�
    filePath = ThisWorkbook.Path & "\" & SETTINGS_FILE_PATH
    
    ' �t�@�C�������݂��Ȃ��ꍇ�̓G���[���b�Z�[�W��\�����ďI��
    If Not fso.FileExists(filePath) Then
        MsgBox "�ݒ�t�@�C�� " & SETTINGS_FILE_PATH & " ��������܂���B", vbExclamation
        Exit Sub
    End If
    
    ' �e�L�X�g�t�@�C�����J��
    Set textFile = fso.OpenTextFile(filePath, 1)
    
    ' �t�@�C���̏I���܂�1�s���ǂݍ���
    currentSection = ""
    
    Do Until textFile.AtEndOfStream
        textLine = Trim(textFile.ReadLine)
        
        ' ��s�̓X�L�b�v
        If textLine = "" Then
            ' �������Ȃ�
        ' �Z�N�V�������̍s�̏ꍇ
        ElseIf Left(textLine, 1) = "[" And Right(textLine, 1) = "]" Then
            currentSection = textLine
        ' �}�N�����̍s�̏ꍇ
        Else
            ' ���݂̃Z�N�V�����Ɋ�Â��ăR���N�V�����ɒǉ�
            Select Case currentSection
                Case SECTION_INDIVIDUAL_BUTTONS
                    individualButtons.Add textLine
                Case SECTION_DROPDOWN_1
                    dropdownList1.Add textLine
                Case SECTION_DROPDOWN_2
                    dropdownList2.Add textLine
                Case SECTION_DROPDOWN_3
                    dropdownList3.Add textLine
            End Select
        End If
    Loop
    
    ' �t�@�C�������
    textFile.Close
    
    ' �I�u�W�F�N�g�̉��
    Set textFile = Nothing
    Set fso = Nothing
End Sub

' �J�X�^�����j���[�̍폜
Sub RemoveCustomControlsMenu()
    On Error Resume Next
    Application.CommandBars(CUSTOM_MENU_NAME).Delete
    On Error GoTo 0
    MsgBox "�J�X�^�����j���[���폜���܂����B", vbInformation
End Sub

' ���I���j���[�̐���
Sub GenerateDynamicMenu()
    InitializeModule
    
    ' �����̃��j���[������΍폜
    On Error Resume Next
    Application.CommandBars(CUSTOM_MENU_NAME).Delete
    On Error GoTo 0
    
    ' �V�����R�}���h�o�[�̍쐬
    Dim customBar As CommandBar
    Set customBar = Application.CommandBars.Add(Name:=CUSTOM_MENU_NAME, Position:=msoBarTop, Temporary:=True)
    
    ' �ʃ{�^���̒ǉ�
    AddIndividualButtons customBar
    
    ' �h���b�v�_�E�����X�g�̒ǉ�
    AddDropdownList customBar, "1", dropdownList1
    AddDropdownList customBar, "2", dropdownList2
    AddDropdownList customBar, "3", dropdownList3
    
    ' �R�}���h�o�[��\��
    customBar.Visible = True
End Sub

' �ʃ{�^����ǉ�����
Private Sub AddIndividualButtons(bar As CommandBar)
    Dim i As Integer
    Dim btn As CommandBarButton
    Dim macroName As String
    
    ' �ʃ{�^���̒ǉ�
    For i = 1 To individualButtons.Count
        Set btn = bar.Controls.Add(Type:=msoControlButton)
        
        macroName = individualButtons(i)
        
        With btn
            .Style = msoButtonIconAndCaption
            ' �V���[�g�J�b�g�L�[�����蓖�Ă�i�\������_EUM���Ȃ��j
            .Caption = AssignShortcutKey(GetDisplayName(macroName))
            ' .Picture = LoadPicture(ThisWorkbook.Path & "\vba-files\Module\Icons\" & macroName & ".ico")
            .OnAction = macroName
            ' �傫�߂̃{�^���ɂ���
            .Height = 40
            .Width = 100
            ' �W���A�C�R���̐ݒ�i�K�v�ɉ����Ē����j
            ' .FaceId = 100 + i ' �A�ԂŃA�C�R����ݒ�i�K�X�����j
            .FaceId = 1
            .BeginGroup = (i = 1) ' �ŏ��̃{�^���̑O�ɋ�؂����ǉ�
        End With
    Next i
End Sub

' �h���b�v�_�E�����X�g�ƃ{�^����ǉ�����
Private Sub AddDropdownList(bar As CommandBar, caption As String, menuItems As Collection)
    ' �h���b�v�_�E���R���g���[���̍쐬
    Dim ctrl As CommandBarComboBox
    Set ctrl = bar.Controls.Add(Type:=msoControlDropdown)
    
    ' ��ӂ̃R���g���[��ID�𐶐�
    Dim controlId As String
    controlId = GetUniqueControlId()
    
    ' �Ή�����}�b�s���O�f�B�N�V���i����I��
    Dim currentMap As Object
    Select Case caption
        Case "1"
            Set currentMap = displayToMacroMap1
        Case "2"
            Set currentMap = displayToMacroMap2
        Case "3"
            Set currentMap = displayToMacroMap3
        Case Else
            Set currentMap = displayToMacroMap
    End Select
    
    ' �L���v�V�����ƃ}�b�v�����^�O�Ɋi�[�iJSON�̂悤�Ȍ`���Łj
    Dim mapTag As String
    mapTag = caption & ":" & controlId
    
    With ctrl
        ' �V���[�g�J�b�g�L�[�����蓖�Ă�
        .Caption = AssignShortcutKey(caption)
        
        ' �R���N�V�������烁�j���[���ڂ�ǉ�
        Dim i As Integer
        Dim macroName As String
        Dim displayName As String
        
        ' �}�b�v�̃N���A�i�e�h���b�v�_�E���p�̃}�b�v���N���A�j
        currentMap.RemoveAll
        
        For i = 1 To menuItems.Count
            macroName = menuItems(i)
            displayName = GetDisplayName(macroName)
            
            ' �\�����Ǝ��ۂ̃}�N�����̃}�b�s���O��ۑ�
            currentMap(displayName) = macroName
            
            ' �O���[�o���}�b�v�ɂ��ǉ��i����݊����̂��߁j
            displayToMacroMap(displayName) = macroName
            
            ' �\�����݂̂�ǉ�
            .AddItem displayName
        Next i
        
        .Width = 200  ' �h���b�v�_�E���̕���ݒ�
        .Tag = mapTag
        .BeginGroup = True  ' �O�̃R���g���[���Ƃ̊Ԃɋ�؂����ǉ�
        
        ' �����I����ݒ�
        If .ListCount > 0 Then
            .ListIndex = 1  ' �f�t�H���g�ōŏ��̍��ڂ�I��
        End If
    End With
    
    ' ���s�{�^���̍쐬
    Dim btn As CommandBarButton
    Set btn = bar.Controls.Add(Type:=msoControlButton)

    With btn
        .Style = msoButtonIconAndCaption
        ' �V���[�g�J�b�g�L�[�����蓖�Ă�
        .Caption = AssignShortcutKey(" ")
        .OnAction = "ExecuteSelectedMacro"
        .FaceId = 156
        .Tag = mapTag
    End With
End Sub

' �I�����ꂽ�}�N�������s����
Sub ExecuteSelectedMacro()
    Dim btn As CommandBarControl
    Set btn = Application.CommandBars.ActionControl
    
    Dim ctrl As CommandBarComboBox
    Set ctrl = GetControlFromTag(btn.Parent, btn.Tag)
    
    If Not ctrl Is Nothing Then
        If ctrl.Text <> "" Then
            ' �h���b�v�_�E�����X�g�̑I�����ڂ̃e�L�X�g�i�\�����j���擾
            Dim displayName As String
            displayName = ctrl.Text
            
            ' �^�O����h���b�v�_�E���ԍ����擾
            Dim dropdownNumber As String
            Dim tagParts As Variant
            tagParts = Split(ctrl.Tag, ":")
            dropdownNumber = tagParts(0)
            
            ' �Ή�����}�N�������擾�i�h���b�v�_�E���ԍ��ɉ������}�b�v���g�p�j
            Dim macroName As String
            Dim currentMap As Object
            
            Select Case dropdownNumber
                Case "1"
                    Set currentMap = displayToMacroMap1
                Case "2"
                    Set currentMap = displayToMacroMap2
                Case "3"
                    Set currentMap = displayToMacroMap3
                Case Else
                    Set currentMap = displayToMacroMap
            End Select
            
            If currentMap.Exists(displayName) Then
                macroName = currentMap(displayName)
                Application.Run macroName
            Else
                ' �}�b�s���O���Ȃ��ꍇ�͕\���������̂܂܎g�p�i����݊����j
                Application.Run displayName
            End If
        Else
            MsgBox "�}�N�����I������Ă��܂���B", vbExclamation
        End If
    Else
        MsgBox "�Ή�����R���g���[����������܂���B", vbExclamation
    End If
End Sub

' �^�O����R���g���[�����擾����
Function GetControlFromTag(bar As CommandBar, tagValue As String) As CommandBarComboBox
    Dim ctrl As CommandBarControl
    
    ' �^�O����R���g���[��ID�𒊏o�i�t�H�[�}�b�g: "�ԍ�:�R���g���[��ID"�j
    Dim controlId As String
    Dim tagParts As Variant
    
    ' �^�O�𕪉�
    tagParts = Split(tagValue, ":")
    If UBound(tagParts) < 1 Then
        ' ���`���̃^�O�̏ꍇ�͌݊����̂��߂ɏ���
        controlId = tagValue
    Else
        controlId = tagParts(1)
    End If
    
    ' �R���g���[��ID����v����R���g���[��������
    For Each ctrl In bar.Controls
        If ctrl.Tag <> "" Then
            Dim ctrlTagParts As Variant
            ctrlTagParts = Split(ctrl.Tag, ":")
            
            ' �V�`���̃^�O�̏ꍇ
            If UBound(ctrlTagParts) >= 1 Then
                If ctrlTagParts(1) = controlId And TypeOf ctrl Is CommandBarComboBox Then
                    Set GetControlFromTag = ctrl
                    Exit Function
                End If
            ' ���`���̃^�O�̏ꍇ�i�݊����̂��߁j
            ElseIf ctrl.Tag = controlId And TypeOf ctrl Is CommandBarComboBox Then
                Set GetControlFromTag = ctrl
                Exit Function
            End If
        End If
    Next ctrl
    
    Set GetControlFromTag = Nothing
End Function

' ��ӂ̃R���g���[��ID�𐶐�����
Private Function GetUniqueControlId() As String
    controlIdCounter = controlIdCounter + 1
    GetUniqueControlId = "Ctrl_" & controlIdCounter
End Function

' �Ώۃ��W���[�����ǂ����𔻒肷��
Function IsTargetModule(moduleName As String) As Boolean
    Dim i As Integer
    For i = LBound(TargetModules) To UBound(TargetModules)
        If TargetModules(i) = moduleName Then
            IsTargetModule = True
            Exit Function
        End If
    Next i
    IsTargetModule = False
End Function
