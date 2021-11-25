Option Explicit

' select文の実行
Sub ExecuteSelect()

    ' 接続設定
    On Error GoTo ErrorHandler
    Dim con As Object: Set con = DbConnection
    con.Open
    
    ' SQL
    Dim sql As String: sql = Worksheets("Main").Cells(14, 3)
    
    ' SQL実行結果取得
    Dim res() As String: res = getResult(sql, con)
    Dim rows As Long: rows = UBound(res, 1)
    Dim columns As Long: columns = UBound(res, 2)
    
    ' 出力用シート
    createOutputSheet ("Result")
    
    '結果をシートに貼り付け
    Worksheets("Result").Range(Cells(1, 1), Cells(rows, columns)) = res
    
    GoTo Finally
    
ErrorHandler:
    MsgBox "データ取得に失敗しました。"

Finally:
    
End Sub

' 結果出力用シートの作成
Function createOutputSheet(sheetName As String)
    
    ' Resultシートの存在判定
    Dim flg As Boolean
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = sheetName Then
            flg = True
            ws.Cells.Clear
            ws.Select
            Exit For
        End If
    Next
    
    ' Resultシートが存在しない場合、シートを作成する
    If flg = False Then
        Worksheets.Add
        ActiveSheet.Name = sheetName
    End If
End Function

' DBへの接続を取得
Function DbConnection() As Object
    ' DB設定値を取得
    Dim confsheet As Worksheet: Set confsheet = Worksheets("Main")
    Dim driver As String: driver = confsheet.Cells(4, 5)
    Dim server As String: server = confsheet.Cells(5, 5)
    Dim port As String: port = confsheet.Cells(6, 5)
    Dim database As String: database = confsheet.Cells(7, 5)
    Dim user As String: user = confsheet.Cells(8, 5)
    Dim password As String: password = confsheet.Cells(9, 5)
    
    ' 接続設定
    Dim con As Object
    Set con = CreateObject("ADODB.Connection")
    con.ConnectionString = "Driver={" & driver & "};" & _
                          "Server=" & server & ";" & _
                          "Port=" & port & ";" & _
                          "Database=" & database & ";" & _
                          "User=" & user & ";" & _
                          "Password=" & password & ";"
    Set DbConnection = con
End Function

' SQLからselect結果を取得する
Function getResult(sql As String, con As Object) As String()
    Dim result() As String
    Dim record As Collection
    Dim recordList As Collection: Set recordList = New Collection
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")
    
    rs.Open sql, con
    
    ' カラム名を取得
    Dim col As Long
    Set record = New Collection
    For col = 1 To rs.Fields.Count
        record.Add rs.Fields(col - 1).Name
    Next
    recordList.Add record
    
    ' データレコードを取得
    Dim row As Long: row = 1
    Do Until rs.EOF
        Set record = New Collection
        For col = 1 To rs.Fields.Count
            record.Add rs(col - 1).Value
        Next
        recordList.Add record
        rs.MoveNext
        row = row + 1
    Loop
    
    ' 二次元配列に入れ替え
    ReDim result(row, rs.Fields.Count)
    Dim x As Long
    Dim y As Long
    For x = 1 To row
        For y = 1 To rs.Fields.Count
            If IsNull(recordList(x)(y)) Then
                result(x - 1, y - 1) = WorksheetFunction.Unichar(171) & "NULL" & WorksheetFunction.Unichar(187)
            Else
                result(x - 1, y - 1) = recordList(x)(y)
            End If
        Next
    Next
    
    getResult = result
    
End Function
