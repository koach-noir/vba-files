Attribute VB_Name = "Module4"
Sub ConnectToPostgreSQLUsingODBC()
    Dim conn As Object
    Dim sqlQuery As String
    
    ' ADODB.Connectionオブジェクトを初期化
    Set conn = CreateObject("ADODB.Connection")
       
    ' ODBCデータソース名(DSN)を指定した接続文字列
    ' "YourDSN"はODBCデータソースアドミニストレーターで設定した名前に置き換えてください
    ' ユーザー名とパスワードが必要な場合は、接続文字列に追加してください
    Dim connectionString As String
    connectionString = "DSN=PostgreSQL-local-mrojapan;UID=postgres;PWD=noir;"
    
    ' データベース接続を開く
    conn.Open connectionString
    
    ' SQLクエリ
    sqlQuery = "SELECT current_database() AS col1;"
    'sqlQuery = "SELECT tablename AS col1 FROM pg_catalog.pg_tables WHERE schemaname = 'public';"
    'sqlQuery = "SELECT ""employeeId"" AS col1 FROM ""MasterUser"";"
    'sqlQuery = "SELECT 1 AS col1;"
    
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    
    ' SQLクエリを実行し、結果をレコードセットに格納
    rs.Open sqlQuery, conn, 1, 3
    
    ' 結果セットをループ出力（イミディエイトウィンドウ）
    While Not rs.EOF
        Debug.Print rs.Fields("col1").value
        rs.MoveNext
    Wend
    
    ' レコードセットと接続をクリーンアップ
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
End Sub
