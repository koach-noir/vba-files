Attribute VB_Name = "M_MenuBuilder"
Option Explicit

' メニュー名の定数
Const CUSTOM_MENU_NAME As String = "EUMControlsMenu"
Const SETTINGS_FILE_PATH As String = "vba-files\Module\EUMMenuSettings.txt"

' セクション名の定数
Const SECTION_INDIVIDUAL_BUTTONS As String = "[IndividualButtons]"
Const SECTION_DROPDOWN_1 As String = "[DropDownList&Buttons1]"
Const SECTION_DROPDOWN_2 As String = "[DropDownList&Buttons2]"
Const SECTION_DROPDOWN_3 As String = "[DropDownList&Buttons3]"

' EUMサフィックス
Const EUM_SUFFIX As String = "_EUM"

' ターゲットとするモジュール名のリスト
Private TargetModules As Variant

' 一意のID生成用の変数
Private controlIdCounter As Long

' メニュー構成用のコレクション
Private individualButtons As Collection
Private dropdownList1 As Collection
Private dropdownList2 As Collection
Private dropdownList3 As Collection

' マクロ名を保存するためのコレクション（キー=表示名, 値=マクロ名）
Private displayToMacroMap1 As Object
Private displayToMacroMap2 As Object
Private displayToMacroMap3 As Object
Private displayToMacroMap As Object

' ショートカットキーのリスト
Private shortcutKeysList() As String
Private currentShortcutKeyIndex As Integer

Sub InitializeModule()
    TargetModules = Array("M_Macros") ' 必要に応じて対象モジュールを追加
    
    ' 各コレクションの初期化
    Set individualButtons = New Collection
    Set dropdownList1 = New Collection
    Set dropdownList2 = New Collection
    Set dropdownList3 = New Collection
    
    ' マクロ名マッピングの初期化
    Set displayToMacroMap1 = CreateObject("Scripting.Dictionary")
    Set displayToMacroMap2 = CreateObject("Scripting.Dictionary")
    Set displayToMacroMap3 = CreateObject("Scripting.Dictionary")
    Set displayToMacroMap = CreateObject("Scripting.Dictionary") ' 後方互換性のため
    
    ' ショートカットキーの初期化
    shortcutKeysList = Split("Q,W,E,R,T,Y,U,I,O,A,S,D,F,G,H,J,K,L", ",")
    currentShortcutKeyIndex = 0
    
    ' 設定ファイルからメニュー項目を読み込む
    LoadMenuSettingsFromFile
End Sub

' ショートカットキーをキャプションに追加する関数
Function AssignShortcutKey(caption As String) As String
    ' ショートカットキーを割り当てる
    If currentShortcutKeyIndex <= UBound(shortcutKeysList) Then
        ' ショートカットキーを追加
        AssignShortcutKey = caption & "(&" & shortcutKeysList(currentShortcutKeyIndex) & ")"
        ' インデックスを増やす
        currentShortcutKeyIndex = currentShortcutKeyIndex + 1
    Else
        ' ショートカットキーが足りない場合はそのまま返す
        AssignShortcutKey = caption
    End If
End Function

' マクロ名から表示用の名前を取得する関数
Function GetDisplayName(macroName As String) As String
    ' マクロ名の末尾に_EUMがある場合は省く
    If Right(macroName, Len(EUM_SUFFIX)) = EUM_SUFFIX Then
        GetDisplayName = Left(macroName, Len(macroName) - Len(EUM_SUFFIX))
    Else
        GetDisplayName = macroName
    End If
End Function

' 設定ファイルからメニュー設定を読み込む
Sub LoadMenuSettingsFromFile()
    Dim fso As Object
    Dim textFile As Object
    Dim filePath As String
    Dim textLine As String
    Dim currentSection As String
    
    ' ファイルシステムオブジェクトの作成
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' ファイルパスの設定
    filePath = ThisWorkbook.Path & "\" & SETTINGS_FILE_PATH
    
    ' ファイルが存在しない場合はエラーメッセージを表示して終了
    If Not fso.FileExists(filePath) Then
        MsgBox "設定ファイル " & SETTINGS_FILE_PATH & " が見つかりません。", vbExclamation
        Exit Sub
    End If
    
    ' テキストファイルを開く
    Set textFile = fso.OpenTextFile(filePath, 1)
    
    ' ファイルの終わりまで1行ずつ読み込む
    currentSection = ""
    
    Do Until textFile.AtEndOfStream
        textLine = Trim(textFile.ReadLine)
        
        ' 空行はスキップ
        If textLine = "" Then
            ' 何もしない
        ' セクション名の行の場合
        ElseIf Left(textLine, 1) = "[" And Right(textLine, 1) = "]" Then
            currentSection = textLine
        ' マクロ名の行の場合
        Else
            ' 現在のセクションに基づいてコレクションに追加
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
    
    ' ファイルを閉じる
    textFile.Close
    
    ' オブジェクトの解放
    Set textFile = Nothing
    Set fso = Nothing
End Sub

' カスタムメニューの削除
Sub RemoveCustomControlsMenu()
    On Error Resume Next
    Application.CommandBars(CUSTOM_MENU_NAME).Delete
    On Error GoTo 0
    MsgBox "カスタムメニューを削除しました。", vbInformation
End Sub

' 動的メニューの生成
Sub GenerateDynamicMenu()
    InitializeModule
    
    ' 既存のメニューがあれば削除
    On Error Resume Next
    Application.CommandBars(CUSTOM_MENU_NAME).Delete
    On Error GoTo 0
    
    ' 新しいコマンドバーの作成
    Dim customBar As CommandBar
    Set customBar = Application.CommandBars.Add(Name:=CUSTOM_MENU_NAME, Position:=msoBarTop, Temporary:=True)
    
    ' 個別ボタンの追加
    AddIndividualButtons customBar
    
    ' ドロップダウンリストの追加
    AddDropdownList customBar, "1", dropdownList1
    AddDropdownList customBar, "2", dropdownList2
    AddDropdownList customBar, "3", dropdownList3
    
    ' コマンドバーを表示
    customBar.Visible = True
End Sub

' 個別ボタンを追加する
Private Sub AddIndividualButtons(bar As CommandBar)
    Dim i As Integer
    Dim btn As CommandBarButton
    Dim macroName As String
    
    ' 個別ボタンの追加
    For i = 1 To individualButtons.Count
        Set btn = bar.Controls.Add(Type:=msoControlButton)
        
        macroName = individualButtons(i)
        
        With btn
            .Style = msoButtonIconAndCaption
            ' ショートカットキーを割り当てる（表示名は_EUMを省く）
            .Caption = AssignShortcutKey(GetDisplayName(macroName))
            ' .Picture = LoadPicture(ThisWorkbook.Path & "\vba-files\Module\Icons\" & macroName & ".ico")
            .OnAction = macroName
            ' 大きめのボタンにする
            .Height = 40
            .Width = 100
            ' 標準アイコンの設定（必要に応じて調整）
            ' .FaceId = 100 + i ' 連番でアイコンを設定（適宜調整）
            .FaceId = 1
            .BeginGroup = (i = 1) ' 最初のボタンの前に区切り線を追加
        End With
    Next i
End Sub

' ドロップダウンリストとボタンを追加する
Private Sub AddDropdownList(bar As CommandBar, caption As String, menuItems As Collection)
    ' ドロップダウンコントロールの作成
    Dim ctrl As CommandBarComboBox
    Set ctrl = bar.Controls.Add(Type:=msoControlDropdown)
    
    ' 一意のコントロールIDを生成
    Dim controlId As String
    controlId = GetUniqueControlId()
    
    ' 対応するマッピングディクショナリを選択
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
    
    ' キャプションとマップ情報をタグに格納（JSONのような形式で）
    Dim mapTag As String
    mapTag = caption & ":" & controlId
    
    With ctrl
        ' ショートカットキーを割り当てる
        .Caption = AssignShortcutKey(caption)
        
        ' コレクションからメニュー項目を追加
        Dim i As Integer
        Dim macroName As String
        Dim displayName As String
        
        ' マップのクリア（各ドロップダウン用のマップをクリア）
        currentMap.RemoveAll
        
        For i = 1 To menuItems.Count
            macroName = menuItems(i)
            displayName = GetDisplayName(macroName)
            
            ' 表示名と実際のマクロ名のマッピングを保存
            currentMap(displayName) = macroName
            
            ' グローバルマップにも追加（後方互換性のため）
            displayToMacroMap(displayName) = macroName
            
            ' 表示名のみを追加
            .AddItem displayName
        Next i
        
        .Width = 200  ' ドロップダウンの幅を設定
        .Tag = mapTag
        .BeginGroup = True  ' 前のコントロールとの間に区切り線を追加
        
        ' 初期選択を設定
        If .ListCount > 0 Then
            .ListIndex = 1  ' デフォルトで最初の項目を選択
        End If
    End With
    
    ' 実行ボタンの作成
    Dim btn As CommandBarButton
    Set btn = bar.Controls.Add(Type:=msoControlButton)

    With btn
        .Style = msoButtonIconAndCaption
        ' ショートカットキーを割り当てる
        .Caption = AssignShortcutKey(" ")
        .OnAction = "ExecuteSelectedMacro"
        .FaceId = 156
        .Tag = mapTag
    End With
End Sub

' 選択されたマクロを実行する
Sub ExecuteSelectedMacro()
    Dim btn As CommandBarControl
    Set btn = Application.CommandBars.ActionControl
    
    Dim ctrl As CommandBarComboBox
    Set ctrl = GetControlFromTag(btn.Parent, btn.Tag)
    
    If Not ctrl Is Nothing Then
        If ctrl.Text <> "" Then
            ' ドロップダウンリストの選択項目のテキスト（表示名）を取得
            Dim displayName As String
            displayName = ctrl.Text
            
            ' タグからドロップダウン番号を取得
            Dim dropdownNumber As String
            Dim tagParts As Variant
            tagParts = Split(ctrl.Tag, ":")
            dropdownNumber = tagParts(0)
            
            ' 対応するマクロ名を取得（ドロップダウン番号に応じたマップを使用）
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
                ' マッピングがない場合は表示名をそのまま使用（後方互換性）
                Application.Run displayName
            End If
        Else
            MsgBox "マクロが選択されていません。", vbExclamation
        End If
    Else
        MsgBox "対応するコントロールが見つかりません。", vbExclamation
    End If
End Sub

' タグからコントロールを取得する
Function GetControlFromTag(bar As CommandBar, tagValue As String) As CommandBarComboBox
    Dim ctrl As CommandBarControl
    
    ' タグからコントロールIDを抽出（フォーマット: "番号:コントロールID"）
    Dim controlId As String
    Dim tagParts As Variant
    
    ' タグを分解
    tagParts = Split(tagValue, ":")
    If UBound(tagParts) < 1 Then
        ' 旧形式のタグの場合は互換性のために処理
        controlId = tagValue
    Else
        controlId = tagParts(1)
    End If
    
    ' コントロールIDが一致するコントロールを検索
    For Each ctrl In bar.Controls
        If ctrl.Tag <> "" Then
            Dim ctrlTagParts As Variant
            ctrlTagParts = Split(ctrl.Tag, ":")
            
            ' 新形式のタグの場合
            If UBound(ctrlTagParts) >= 1 Then
                If ctrlTagParts(1) = controlId And TypeOf ctrl Is CommandBarComboBox Then
                    Set GetControlFromTag = ctrl
                    Exit Function
                End If
            ' 旧形式のタグの場合（互換性のため）
            ElseIf ctrl.Tag = controlId And TypeOf ctrl Is CommandBarComboBox Then
                Set GetControlFromTag = ctrl
                Exit Function
            End If
        End If
    Next ctrl
    
    Set GetControlFromTag = Nothing
End Function

' 一意のコントロールIDを生成する
Private Function GetUniqueControlId() As String
    controlIdCounter = controlIdCounter + 1
    GetUniqueControlId = "Ctrl_" & controlIdCounter
End Function

' 対象モジュールかどうかを判定する
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
