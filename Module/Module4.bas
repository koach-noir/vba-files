Attribute VB_Name = "Module4"
Sub ConnectToPostgreSQLUsingODBC()
    Dim conn As Object
    Dim sqlQuery As String
    
    ' ADODB.Connection�I�u�W�F�N�g��������
    Set conn = CreateObject("ADODB.Connection")
       
    ' ODBC�f�[�^�\�[�X��(DSN)���w�肵���ڑ�������
    ' "YourDSN"��ODBC�f�[�^�\�[�X�A�h�~�j�X�g���[�^�[�Őݒ肵�����O�ɒu�������Ă�������
    ' ���[�U�[���ƃp�X���[�h���K�v�ȏꍇ�́A�ڑ�������ɒǉ����Ă�������
    Dim connectionString As String
    connectionString = "DSN=PostgreSQL-local-mrojapan;UID=postgres;PWD=noir;"
    
    ' �f�[�^�x�[�X�ڑ����J��
    conn.Open connectionString
    
    ' SQL�N�G��
    sqlQuery = "SELECT current_database() AS col1;"
    'sqlQuery = "SELECT tablename AS col1 FROM pg_catalog.pg_tables WHERE schemaname = 'public';"
    'sqlQuery = "SELECT ""employeeId"" AS col1 FROM ""MasterUser"";"
    'sqlQuery = "SELECT 1 AS col1;"
    
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    
    ' SQL�N�G�������s���A���ʂ����R�[�h�Z�b�g�Ɋi�[
    rs.Open sqlQuery, conn, 1, 3
    
    ' ���ʃZ�b�g�����[�v�o�́i�C�~�f�B�G�C�g�E�B���h�E�j
    While Not rs.EOF
        Debug.Print rs.Fields("col1").value
        rs.MoveNext
    Wend
    
    ' ���R�[�h�Z�b�g�Ɛڑ����N���[���A�b�v
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
End Sub
