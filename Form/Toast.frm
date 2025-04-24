VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Toast 
   Caption         =   "EUM"
   ClientHeight    =   960
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4152
   OleObjectBlob   =   "Toast.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "Toast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


' Toast コード
Private msgText As String

' メッセージテキストを設定するためのプロパティ
Public Property Let messageText(value As String)
    msgText = value
End Property

Private Sub UserForm_Activate()
    ' ラベルにテキストを設定
    Label1.caption = msgText

    ' タイマーを設定（ここではn秒後に設定）
    Application.OnTime Now + TimeValue("00:00:02"), "CloseToast"
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Unload Me
End Sub

