' SheetMind VBA Add-in
' This code can be used to create a traditional .xlam Excel add-in
' 
' Instructions:
' 1. Open Excel > Developer > Visual Basic
' 2. Insert > Module
' 3. Paste this code
' 4. File > Save As > Excel Add-in (.xlam)

Option Explicit

' Configuration
Const SHEETMIND_API_URL As String = "http://localhost:8000"

' Main SheetMind function
Public Sub ShowSheetMindDialog()
    Dim userInput As String
    Dim response As String
    
    ' Get user input
    userInput = InputBox("Ask SheetMind to do something with your data:", "🧠 SheetMind", "Calculate the average of selected cells")
    
    If userInput = "" Then Exit Sub
    
    ' Get current selection info
    Dim excelContext As String
    excelContext = GetExcelContext()
    
    ' Call SheetMind API
    response = CallSheetMindAPI(userInput, excelContext)
    
    ' Show response
    MsgBox response, vbInformation, "🧠 SheetMind Response"
End Sub

' Get current Excel context
Private Function GetExcelContext() As String
    Dim context As String
    Dim sel As Range
    
    Set sel = Selection
    
    context = "{"
    context = context & """worksheet"": {""name"": """ & ActiveSheet.Name & """},"
    context = context & """selection"": {"
    context = context & """address"": """ & sel.Address & ""","
    context = context & """rowCount"": " & sel.Rows.Count & ","
    context = context & """columnCount"": " & sel.Columns.Count & ","
    context = context & """values"": " & GetRangeValues(sel)
    context = context & "}}"
    
    GetExcelContext = context
End Function

' Convert range values to JSON array
Private Function GetRangeValues(rng As Range) As String
    Dim values As String
    Dim r As Integer, c As Integer
    Dim cellValue As Variant
    
    values = "["
    
    For r = 1 To rng.Rows.Count
        If r > 1 Then values = values & ","
        values = values & "["
        
        For c = 1 To rng.Columns.Count
            If c > 1 Then values = values & ","
            
            cellValue = rng.Cells(r, c).Value
            If IsNumeric(cellValue) Then
                values = values & cellValue
            ElseIf IsNull(cellValue) Or cellValue = "" Then
                values = values & "null"
            Else
                values = values & """" & Replace(CStr(cellValue), """", "\""") & """"
            End If
        Next c
        
        values = values & "]"
    Next r
    
    values = values & "]"
    GetRangeValues = values
End Function

' Call SheetMind API
Private Function CallSheetMindAPI(message As String, context As String) As String
    On Error GoTo ErrorHandler
    
    Dim http As Object
    Dim requestBody As String
    Dim response As String
    
    ' Create HTTP request
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    ' Prepare request body
    requestBody = "{"
    requestBody = requestBody & """message"": """ & Replace(message, """", "\""") & ""","
    requestBody = requestBody & """context"": " & context
    requestBody = requestBody & "}"
    
    ' Make request
    http.Open "POST", SHEETMIND_API_URL & "/chat-excel", False
    http.setRequestHeader "Content-Type", "application/json"
    http.send requestBody
    
    ' Parse response
    If http.Status = 200 Then
        Dim jsonResponse As String
        jsonResponse = http.responseText
        
        ' Simple JSON parsing to extract response field
        Dim startPos As Integer, endPos As Integer
        startPos = InStr(jsonResponse, """response"":""") + 12
        endPos = InStr(startPos, jsonResponse, """,""")
        If endPos = 0 Then endPos = InStr(startPos, jsonResponse, """}")
        
        If startPos > 12 And endPos > startPos Then
            response = Mid(jsonResponse, startPos, endPos - startPos)
            response = Replace(response, "\""", """")
            response = Replace(response, "\\", "\")
        Else
            response = "✅ SheetMind processed your request!"
        End If
    Else
        response = "❌ Error: Could not connect to SheetMind API (Status: " & http.Status & ")"
    End If
    
    CallSheetMindAPI = response
    Exit Function
    
ErrorHandler:
    CallSheetMindAPI = "❌ Error: " & Err.Description & vbCrLf & vbCrLf & "Make sure SheetMind backend is running at " & SHEETMIND_API_URL
End Function

' Add ribbon button (this would go in ThisWorkbook module)
Private Sub Workbook_AddinInstall()
    ' Add SheetMind button to ribbon
    AddSheetMindButton
End Sub

Private Sub AddSheetMindButton()
    ' This is a simplified version - real ribbon customization requires XML
    Application.OnKey "^+M", "ShowSheetMindDialog" ' Ctrl+Shift+M shortcut
End Sub

' Quick test function
Public Sub TestSheetMind()
    ShowSheetMindDialog
End Sub 