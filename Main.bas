Attribute VB_Name = "Main"
Option Explicit

Private Declare PtrSafe Function sqlite3_open Lib "winsqlite3.dll" (ByVal filename As LongPtr, ByRef ppDb As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_exec Lib "winsqlite3.dll" (ByVal db As LongPtr, ByVal sql As LongPtr, ByVal callback As LongPtr, _
                                                                    ByVal param As LongPtr, ByRef errmsg As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_close Lib "winsqlite3.dll" (ByVal db As LongPtr) As Long

Const DBPATH = "z:\test.db"

Sub ExecuteQuery()
    Dim db As LongPtr
    Dim Retval As Long
    Dim strSQL As String
    Dim errmsg As LongPtr
    
    ' データベースを開く
    Retval = sqlite3_open(StrPtr(StringToUTF8(DBPATH)), db)
    If Retval <> 0 Then
        Call MsgBox("データベースのオープンにに失敗しました")
        Exit Sub
    End If
    
    ' テーブルを作成するSQL文
    strSQL = "CREATE TABLE IF NOT EXISTS テスト ( " & _
          "ID INTEGER PRIMARY KEY, " & _
          "名前 TEXT NOT NULL, " & _
          "数学 INTEGER, " & _
          "英語 INTEGER, " & _
          "理科 INTEGER);"

    Retval = sqlite3_exec(db, StrPtr(StringToUTF8(strSQL)), 0, 0, errmsg)
    If Retval <> 0 Then
        Call MsgBox("テーブルの作成に失敗しました: " & UTF8ToString(errmsg))
        Call sqlite3_close(db)
        Exit Sub
    End If

    ' データを挿入するSQL文
    strSQL = "INSERT INTO テスト (名前, 数学, 英語, 理科) VALUES " & _
          "('鈴木', 85, 90, 88), " & _
          "('佐藤', 78, 82, 85), " & _
          "('田中', 92, 88, 100);"

    ' SQLを実行してデータを挿入
    Retval = sqlite3_exec(db, StrPtr(StringToUTF8(strSQL)), 0, 0, errmsg)
    If Retval <> 0 Then
        Call MsgBox("データの挿入に失敗しました: " & UTF8ToString(errmsg))
        Call sqlite3_close(db)
        Exit Sub
    End If
    
    Call sqlite3_close(db)
End Sub

