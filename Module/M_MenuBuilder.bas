Attribute VB_Name = "M_MenuBuilder"
Option Explicit

' ���j���[���̒萔
Const CUSTOM_MENU_NAME As String = "CustomControlsMenu"

' �^�[�Q�b�g�Ƃ��郂�W���[�����̃��X�g
Private TargetModules As Variant

' ��ӂ�ID�����p�̕ϐ�
Private controlIdCounter As Long

' �D�揇�ʂ̒�`
Private PriorityMacros As Variant

Sub InitializeModule()
    TargetModules = Array("M_Macros") ' �K�v�ɉ����đΏۃ��W���[����ǉ�
    
    ' �D�揇�ʂ̍����}�N�����`
    PriorityMacros = Array( _
        "�I���Z���s����_EUM", _
        "�I���Z���s����_EUM", _
        "�s�����񕝎�������_EUM", _
        "�I���Z���s�R�s�[�}��_EUM", _
        "������v�Z���ꊇ�I��_EUM", _
        "�u�b�N���R�s�[_EUM", _
        "�V�[�g���R�s�[Selected_EUM", _
        "�u�b�N�p�X���G�N�X�v���[���[�ŕ\��_EUM", _
        "�S�ẴV�[�g��A1�I��_EUM" _
    )
End Sub

Sub RemoveCustomControlsMenu()
    On Error Resume Next
    Application.CommandBars(CUSTOM_MENU_NAME).Delete
    On Error GoTo 0
    MsgBox "Custom menu has been removed.", vbInformation
End Sub

Sub GenerateDynamicMenu()
    InitializeModule
    
    On Error Resume Next
    Application.CommandBars(CUSTOM_MENU_NAME).Delete
    On Error GoTo 0
    
    Dim customBar As CommandBar
    Set customBar = Application.CommandBars.Add(Name:=CUSTOM_MENU_NAME, Position:=msoBarTop, Temporary:=True)
    
    ' ���j���[���ڂ����W�p��Collection���쐬
    Dim menuItems As Collection
    Set menuItems = New Collection
    
    ' �܂��D�揇�ʂ̍����}�N����ǉ�
    Dim priorityMacro As Variant
    For Each priorityMacro In PriorityMacros
        menuItems.Add priorityMacro
    Next priorityMacro
        
    ' ���j���[�A�C�e�����쐬
    Dim i As Integer
    Dim shortcutKeysList() As String
    Dim shortcutKeysBtn() As String
    shortcutKeysList = Split("Q,W,E,R,T,Y,U,I,O", ",")
    shortcutKeysBtn = Split("A,S,D,F,G,H,J,K,L", ",")
    
    ' Collection ��z��ɕϊ�
    Dim menuItemsArray() As String
    ReDim menuItemsArray(1 To menuItems.Count)
    For i = 1 To menuItems.Count
        menuItemsArray(i) = menuItems(i)
    Next i
    
    For i = 0 To menuItems.Count - 1
        If i <= UBound(shortcutKeysList) Then
            AddControlDropdownWithButton customBar, " ", shortcutKeysList(i), shortcutKeysBtn(i), menuItemsArray, i + 1
        End If
    Next i
    
    customBar.Visible = True
End Sub

Private Function IsInPriorityList(macroName As String) As Boolean
    Dim item As Variant
    For Each item In PriorityMacros
        If item = macroName Then
            IsInPriorityList = True
            Exit Function
        End If
    Next item
    IsInPriorityList = False
End Function

Private Sub AddControlDropdownWithButton(bar As CommandBar, caption As String, shortcutKeyList As String, shortcutKeyBtn As String, menuItems() As String, initialSelection As Integer)
    Dim ctrl As CommandBarComboBox
    Set ctrl = bar.Controls.Add(Type:=msoControlDropdown)
    
    Dim controlId As String
    controlId = GetUniqueControlId()
    
    With ctrl
        .caption = caption & "(&" & shortcutKeyList & ")"
        
        ' ���I�Ƀ��j���[���ڂ�ǉ�
        Dim item As Variant
        For Each item In menuItems
            .AddItem item
        Next item
        
        .Width = 200  ' �h���b�v�_�E���̕���ݒ�
        .Tag = controlId
        .BeginGroup = True  ' �O�̃R���g���[���Ƃ̊Ԃɋ�؂����ǉ�
        
        ' �����I����ݒ�i�C���Łj
        If .ListCount > 0 Then
            If initialSelection > 0 And initialSelection <= .ListCount Then
                .ListIndex = initialSelection
            Else
                .ListIndex = 1  ' �f�t�H���g�ōŏ��̍��ڂ�I��
            End If
        End If
    End With
    
    Dim btn As CommandBarButton
    Set btn = bar.Controls.Add(Type:=msoControlButton)
    
    With btn
        .Style = msoButtonIconAndCaption
        .caption = " " & "(&" & shortcutKeyBtn & ")"
        .OnAction = "ExecuteSelectedMacro"
        .FaceId = 44
        .Tag = controlId
    End With
End Sub

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

Sub ExecuteSelectedMacro()
    Dim btn As CommandBarControl
    Set btn = Application.CommandBars.ActionControl
    
    Dim ctrl As CommandBarComboBox
    Set ctrl = GetControlFromTag(btn.Parent, btn.Tag)
    
    If Not ctrl Is Nothing Then
        If ctrl.Text <> "" Then
            Application.Run ctrl.Text
        Else
            MsgBox "�}�N�����I������Ă��܂���B", vbExclamation
        End If
    Else
        MsgBox "�Ή�����R���g���[����������܂���B", vbExclamation
    End If
End Sub

Function GetControlFromTag(bar As CommandBar, tagValue As String) As CommandBarComboBox
    Dim ctrl As CommandBarControl
    For Each ctrl In bar.Controls
        If ctrl.Tag = tagValue Then
            If TypeOf ctrl Is CommandBarComboBox Then
                Set GetControlFromTag = ctrl
                Exit Function
            End If
        End If
    Next ctrl
    Set GetControlFromTag = Nothing
End Function

Private Function GetUniqueControlId() As String
    controlIdCounter = controlIdCounter + 1
    GetUniqueControlId = "Ctrl_" & controlIdCounter
End Function
