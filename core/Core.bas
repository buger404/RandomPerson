Attribute VB_Name = "Core"
Public MusicList As GMusicList
Function ReadExcel(ByVal path As String, Optional ByVal SheetIndex As Integer = 1) As String()
    Dim App As Object, WBook As Object, WB As Object, Sheet As Object

    Set App = CreateObject("Excel.Application")
    Set WBook = App.WorkBooks
    Set WB = App.WorkBooks.Open(path)
    Set Sheet = WB.WorkSheets(SheetIndex)
    
    Dim Row As Integer, Line As Integer
    Dim data(), Ret() As String
    
    data = Sheet.UsedRange.value
    ReDim Ret(UBound(data, 2) - 1, UBound(data, 1) - 1)
    
    For Row = 1 To UBound(data, 1)
        For Line = 1 To UBound(data, 2)
            Ret(Line - 1, Row - 1) = data(Row, Line)
        Next
    Next
    
    App.Quit
    Set Sheet = Nothing
    Set WBook = Nothing
    Set App = Nothing
    
    ReadExcel = Ret
End Function
