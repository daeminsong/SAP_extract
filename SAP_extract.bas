Attribute VB_Name = "Module1"
Sub SAPExtract()

    Dim Monthenddate As String
    Dim Monthdigit As String
    Dim Yeardigit As String
    Dim wkb As Excel.Workbook
    Dim SheetToFind As String
    Dim worksh As Integer
    Dim worksheetexists As Boolean
    Dim cell As Range
    Dim rng As Range

    On Error GoTo 0
    Set wkb = ThisWorkbook

    Monthenddate = InputBox("Please enter month end period in 'MMYYYY' form (i.e 012018)")
    If Monthenddate = vbNullString Then
    MsgBox "User canceled."
        Exit Sub
    End If
    If Not IsNumeric(Monthenddate) Then
        MsgBox ("You should put only numbers (i.e. 012018), no other characters allowed")
        Exit Sub
    End If
    
    Monthdigit = Left(Monthenddate, 2)
    Yeardigit = Right(Monthenddate, 4)
    SheetToFind = "Worksheet name" ''''''''' enter the name of the worksheet that you would like to work on ''''''''''''''
    
    wkb.Activate
    wkb.Worksheets(2).Select
    worksh = Application.Sheets.Count
    worksheetexists = False
    For x = 1 To worksh
        If Worksheets(x).Name = SheetToFind Then
            worksheetexists = True
            Exit For
        End If
    Next x
        If worksheetexists = False Then
            Debug.Print "Transformed exists"
            wkb.Worksheets(2).Copy Before:=Sheets(2)
            wkb.Worksheets(2).Name = SheetToFind
        End If
    
    'SAP Extract
    Set SapGuiAuto = GetObject("SAPGUI")  'Get the SAP GUI Scripting object
    Set SAPApp = SapGuiAuto.GetScriptingEngine 'Get the currently running SAP GUI
    Set SAPCon = SAPApp.Children(0) 'Get the first system that is currently connected
    Set session = SAPCon.Children(0) 'Get the first session (window) on that connection

    ''''''''' you need to specify what variables you would like to use''''''''''''''''''''''
    
    session.StartTransaction "Trnasaction name" 'Transaction
    session.findById("wnd[0]/usr/ctxtSD_KTOPL-LOW").Text = "Strings"
    
    ''''''''' you need to specify what variables you would like to use''''''''''''''''''''''
    
    'SAP Extract part done

    With wkb.Worksheets(SheetToFind)
    .Activate
    .Range("A:R").Clear
    .Range("A1").Select
    .Range("A1").PasteSpecial
    .Range("A:A").TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="|", FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
    .Range("F:F").Select
    Selection.Insert Shift:=xlToRight
    .Range("E:E").TextToColumns Destination:=Range("E1"), DataType:=xlFixedWidth, _
        OtherChar:="|", FieldInfo:=Array(Array(0, 1), Array(8, 1)), _
        TrailingMinusNumbers:=True
    Columns("E:I").EntireColumn.AutoFit
    End With


    'move minus sign from right to left on entire worksheet
    On Error Resume Next
    Set rng = ActiveSheet.Range("G:J"). _
            SpecialCells(xlCellTypeConstants, xlTextValues)
    On Error GoTo 0
    For Each cell In rng
        If IsNumeric(cell.Value) Then
            cell.Value = CDbl(cell.Value)
        End If
    Next cell
    'Number format
    With wkb.Worksheets(SheetToFind).Range("G:I")
    .NumberFormat = "0.00"
    .Style = "Comma"
    End With
    
    MsgBox "Done"
    
End Sub

