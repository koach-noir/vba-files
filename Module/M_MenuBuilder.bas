Attribute VB_Name = "M_MenuBuilder"
Option Explicit

' メニュー名の定数
Const CUSTOM_MENU_NAME As String = "CustomControlsMenu"

' ターゲットとするモジュール名のリスト
Private TargetModules As Variant

' 一意のID生成用の変数
Private controlIdCounter As Long

' 優先順位の定義
Private PriorityMacros As Variant

Sub InitializeModule()
    TargetModules = Array("M_Macros") ' 必要に応じて対象モジュールを追加
    
    ' 優先順位の高いマクロを定義
    PriorityMacros = Array( _
        "選択セル行高小_EUM", _
        "選択セル行高大_EUM", _
        "行高さ列幅自動調整_EUM", _
        "選択セル行コピー挿入_EUM", _
        "部分一致セル一括選択_EUM", _
        "ブック名コピー_EUM", _
        "シート名コピーSelected_EUM", _
        "ブックパスをエクスプローラーで表示_EUM", _
        "全てのシートでA1選択_EUM" _
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
    
    ' メニュー項目を収集用のCollectionを作成
    Dim menuItems As Collection
    Set menuItems = New Collection
    
    ' まず優先順位の高いマクロを追加
    Dim priorityMacro As Variant
    For Each priorityMacro In PriorityMacros
        menuItems.Add priorityMacro
    Next priorityMacro
        
    ' メニューアイテムを作成
    Dim i As Integer
    Dim shortcutKeysList() As String
    Dim shortcutKeysBtn() As String
    shortcutKeysList = Split("Q,W,E,R,T,Y,U,I,O", ",")
    shortcutKeysBtn = Split("A,S,D,F,G,H,J,K,L", ",")
    
    ' Collection を配列に変換
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
        
        ' 動的にメニュー項目を追加
        Dim item As Variant
        For Each item In menuItems
            .AddItem item
        Next item
        
        .Width = 200  ' ドロップダウンの幅を設定
        .Tag = controlId
        .BeginGroup = True  ' 前のコントロールとの間に区切り線を追加
        
        ' 初期選択を設定（修正版）
        If .ListCount > 0 Then
            If initialSelection > 0 And initialSelection <= .ListCount Then
                .ListIndex = initialSelection
            Else
                .ListIndex = 1  ' デフォルトで最初の項目を選択
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
            MsgBox "マクロが選択されていません。", vbExclamation
        End If
    Else
        MsgBox "対応するコントロールが見つかりません。", vbExclamation
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
