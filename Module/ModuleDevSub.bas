Attribute VB_Name = "ModuleDevSub"

' グローバル変数としてCuidGeneratorインスタンスを保持
Global cuidGenerator As Object

Function EnsureGeneratorInstance() As Object
    If cuidGenerator Is Nothing Then
        ' インスタンスが未初期化の場合のみ、新しいインスタンスを作成
        Set cuidGenerator = CreateObject("DevSupLibrary.CuidGenerator")
    End If
    Set EnsureGeneratorInstance = cuidGenerator
End Function

'Function GenerateCUID() As String
'    ' CuidGeneratorインスタンスを確実に取得
'    Dim gen As Object
'    Set gen = EnsureGeneratorInstance()
'    ' 生成されたCUIDを返す
'    GenerateCUID = gen.GenerateCUID()
'End Function

Sub TestGenerateCUID()
    ' CuidGeneratorクラスのインスタンスを作成
    Dim generator As Object
    Set generator = CreateObject("DevSupLibrary.CuidGenerator")

    ' CUIDを生成
    Dim cuid As String
    cuid = generator.GenerateCUID()
    Debug.Print cuid  ' イミディエイトウィンドウにCUIDを出力
End Sub


Sub DeleteSheetsExceptSheet1()
    Dim ws As Worksheet
    Dim wsToDelete As Worksheet

    ' シート1以外のすべてのシートを削除
    Application.DisplayAlerts = False ' 削除の確認ダイアログを表示しないように設定
    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> "Sheet1" Then
            'ws.Delete ' ●事故防止CO
        End If
    Next ws
    Application.DisplayAlerts = True ' デフォルトの設定に戻す
End Sub
