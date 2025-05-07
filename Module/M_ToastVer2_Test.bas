Attribute VB_Name = "M_ToastVer2_Test"
Option Explicit

' ===================================================
' トースト通知V2 - テストモジュール
' ===================================================
' 作成日: 2025/04/25
' 作成者: Claude
' 概要: 改良版トースト通知をテストするためのモジュール
' 機能:
'   - 各種トースト通知の表示テスト
'   - 連続表示のテスト
'   - カスタマイズのデモンストレーション
' ===================================================

' 基本的なトースト通知のテスト
Public Sub TestBasicToast()
    ' 基本的な情報トースト
    ShowToastVer2 "基本的なトースト通知です", "info", 2000
End Sub

' 各種トースト種類のテスト
Public Sub TestAllToastTypes()
    ' 情報トースト
    ShowInfoToast "情報メッセージです", 2000
    
    ' 少し待機
    Application.Wait Now + TimeSerial(0, 0, 3)
    
    ' 成功トースト
    ShowSuccessToast "処理が完了しました", 2000
    
    ' 少し待機
    Application.Wait Now + TimeSerial(0, 0, 3)
    
    ' 警告トースト
    ShowWarningToast "注意が必要です", 2000
    
    ' 少し待機
    Application.Wait Now + TimeSerial(0, 0, 3)
    
    ' エラートースト
    ShowErrorToast "エラーが発生しました", 2000
End Sub

' 連続トースト表示のテスト（キュー機能）
Public Sub TestToastQueue()
    ' 複数のトーストを連続して表示
    ShowInfoToast "処理を開始します"
    ShowInfoToast "データを読み込んでいます..."
    ShowInfoToast "計算中..."
    ShowSuccessToast "処理が完了しました"
End Sub

' 長いメッセージのテスト
Public Sub TestLongMessage()
    ' 長いメッセージのトースト
    ShowInfoToast "これは長いメッセージのテストです。トースト通知は長いメッセージでも適切に表示されるように設計されています。必要に応じてメッセージの長さに合わせてフォームサイズが調整されます。", 4000
End Sub

' 表示時間のカスタマイズテスト
Public Sub TestCustomDuration()
    ' 短い表示時間
    ShowInfoToast "短い表示（1秒）", 1000
    
    ' 少し待機
    Application.Wait Now + TimeSerial(0, 0, 2)
    
    ' 標準の表示時間
    ShowInfoToast "標準の表示（2秒）", 2000
    
    ' 少し待機
    Application.Wait Now + TimeSerial(0, 0, 3)
    
    ' 長い表示時間
    ShowInfoToast "長い表示（5秒）", 5000
End Sub

' すべてのテストを実行
Public Sub RunAllTests()
    ' 個々のテストを順番に実行
    TestBasicToast
    
    ' 少し待機
    Application.Wait Now + TimeSerial(0, 0, 3)
    
    TestAllToastTypes
    
    ' 少し待機
    Application.Wait Now + TimeSerial(0, 0, 12)
    
    TestToastQueue
    
    ' 少し待機
    Application.Wait Now + TimeSerial(0, 0, 8)
    
    TestLongMessage
    
    ' 少し待機
    Application.Wait Now + TimeSerial(0, 0, 5)
    
    TestCustomDuration
    
    ' テスト完了メッセージ
    MsgBox "トースト通知V2のすべてのテストが完了しました。", vbInformation, "テスト完了"
End Sub

' 実際の使用例のデモンストレーション
Public Sub DemonstrateRealUsage()
    ' 処理開始の通知
    ShowInfoToast "データ分析を開始します..."
    
    ' 実際の処理を模擬（5秒間待機）
    Application.Wait Now + TimeSerial(0, 0, 2)
    
    ' 中間報告
    ShowInfoToast "50件のデータを処理中..."
    
    ' さらに処理を模擬（3秒間待機）
    Application.Wait Now + TimeSerial(0, 0, 2)
    
    ' 処理完了の通知
    ShowSuccessToast "データ分析が完了しました！", 3000
End Sub

' トーストキュークリアのテスト
Public Sub TestClearQueue()
    ' いくつかのトーストをキューに追加
    ShowInfoToast "トースト1"
    ShowInfoToast "トースト2"
    ShowInfoToast "トースト3"
    
    ' 1秒待機してからキューをクリア
    Application.Wait Now + TimeSerial(0, 0, 1)
    
    ' キューをクリア
    ClearToastQueue
    
    ' クリア後の通知
    ShowWarningToast "キューがクリアされました", 3000
End Sub

' カスタム設定でトーストを表示するテスト
Public Sub TestCustomToast()
    ' カスタム設定でトーストを表示
    ShowToastVer2 "カスタム設定のトースト", "success", 3000, "info"
End Sub
