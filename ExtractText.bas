Attribute VB_Name = "ExtractText"
Option Explicit
'Open script file and pull all lines into array
Public Sub ExtractDataFromFile()
    Dim srcPath As String
    Dim srcFile As String
    Dim text As String
    Dim textline As String
    Dim srcText() As String
    Dim indentGroup() As Long
    Dim iLineCount As Long
    
    srcPath = GetScriptsFolder
    
    If srcPath = "" Then Exit Sub
    srcFile = Dir(srcPath & "*.rpy")
    
    Application.ScreenUpdating = False
    
    Call ClearData
    
    Do While srcFile <> ""
    
        Open srcPath & srcFile For Input As #1
        iLineCount = 0
        
        Do While Not EOF(1)
            'Line Input #1, textline 'Line input not stopping at carriage return/line feed
            text = Input(1, #1)
            'Check if ASCII code is a line feed and if the line has anything other than spaces
            If Asc(text) = 10 And Len(Replace(textline, Chr(32), vbNullString)) > 0 Then
                ReDim Preserve srcText(0 To iLineCount)
                srcText(iLineCount) = textline
                textline = vbNullString
                iLineCount = iLineCount + 1
            ElseIf Asc(text) = 10 And Len(Trim(textline)) = 0 Then 'Remove white lines
                textline = vbNullString
            'ASCII code for carriage return, ignore
            ElseIf Asc(text) = 13 Then
                'Do nothing
            'ASCII for space, doesn't get recognized by trim without subbing in " "
            ElseIf Asc(text) = 32 Then
                textline = textline & " "
            Else
                textline = textline & text
            End If
        Loop
        
        Close #1
        textline = vbNullString
        text = vbNullString
        
        
        indentGroup = LeftTrim(srcText)
        Call FindResponseMaximums(srcFile, srcText, indentGroup)
        'Call FindLabelsAndConditionals(srcText)
        
        srcFile = Dir
        
        Erase srcText
        Erase indentGroup
    Loop
    
    Application.ScreenUpdating = True
    
End Sub
'Get Scripts folder from sheet or prompt user
Private Function GetScriptsFolder() As String
    Dim srcFolder As FileDialog
    Dim srcPath As String
    
    srcPath = Sheet1.Range("I1").Value2
    
    If srcPath = "" Then
        Set srcFolder = Application.FileDialog(msoFileDialogFolderPicker)
        
        With srcFolder
            .Title = "Select folder with scripts"
            .AllowMultiSelect = False
            If .Show <> -1 Then Exit Function
            srcPath = .SelectedItems(1) & "\"
        End With
        Sheet1.Range("I1").Value2 = srcPath
    End If
    
    GetScriptsFolder = srcPath
    
End Function
'Loop through conditionals and log highest response value
Private Sub FindResponseMaximums( _
    ByVal srcFile As String, _
    ByRef srcText() As String, _
    ByRef indentGroup() As Long)
    Dim arrResponseData() As String
    Dim i As Long
    Dim iLineCount As Long
    Dim iGroup As Long                              'Indent group tracking for if blocks
    Dim iMenuGroup As Long                          'Indent group for menu blocks
    Dim iStrSearch As Long                          'Variable for position within string
    Dim iStrSearchEnd As Long
    Dim lblGroup As String
    Dim menuGroup As String
    Dim varGroup As String
    Dim val As String
    Dim isInMenu As Boolean
    Dim hasUndefinedValue As Boolean
    
    srcFile = Left(srcFile, Len(srcFile) - 4)
    iLineCount = 0
    isInMenu = False
    hasUndefinedValue = False
    ReDim arrResponseData(1 To 5, 1 To 1)
    
    For i = 0 To UBound(srcText)
        If Left(srcText(i), 5) = "label" Then
            lblGroup = Mid(srcText(i), 7, InStr(srcText(i), ":"))
        ElseIf Left(LTrim(srcText(i)), 4) = "menu" Then
                    'Save indent level and toggle isInMenu
            iMenuGroup = indentGroup(i)
            isInMenu = True
                    'If first non-Space character is a quote, ends in a colon, and we are in a menu, log menu choice
        ElseIf Left(LTrim(srcText(i)), 1) = Chr(34) And Right(RTrim(srcText(i)), 1) = Chr(58) And isInMenu Then
            iStrSearch = 1 + InStr(srcText(i), Chr(34))
            menuGroup = Mid(srcText(i), iStrSearch, InStr(iStrSearch, srcText(i), Chr(34)) - iStrSearch)
                    'If line starts with a conditional, log the first variable and equality check/value
        ElseIf Left(LTrim(srcText(i)), 2) = "if" Then
            iStrSearch = 1 + indentGroup(i) * 4 + 3  'Indent group spacing, "if " spacing, plus one
            iStrSearchEnd = MultiInStr(iStrSearch, srcText(i))
            varGroup = Mid(srcText(i), iStrSearch, iStrSearchEnd - iStrSearch)
            iStrSearch = 1 + iStrSearchEnd
            val = Mid(srcText(i), iStrSearch, InStr(iStrSearch, srcText(i), ":") - iStrSearch)
            If Left(val, 1) = "=" Then val = "'" & val
                        'If description following conditional is not defined, log the parameters
        ElseIf InStr(srcText(i), "not defined") > 0 Then
            'Check for duplicate record
            If hasUndefinedValue Then
                If Not lblGroup = arrResponseData(2, iLineCount) Or _
                    (isInMenu And Not menuGroup = arrResponseData(3, iLineCount)) Or _
                    Not varGroup = arrResponseData(4, iLineCount) Then
                        hasUndefinedValue = False
                End If
            End If
            
            If Not hasUndefinedValue Then
                iLineCount = iLineCount + 1
                ReDim Preserve arrResponseData(1 To 5, 1 To iLineCount)
                arrResponseData(1, iLineCount) = srcFile
                arrResponseData(2, iLineCount) = lblGroup
                If isInMenu Then arrResponseData(3, iLineCount) = menuGroup
                arrResponseData(4, iLineCount) = varGroup
                arrResponseData(5, iLineCount) = val
                
                hasUndefinedValue = True
            End If
                            'Check if we're out of the menu indent level, currently doesn't handle nested menus
        ElseIf isInMenu Then
            If indentGroup(i) <= iMenuGroup Then
                isInMenu = False
            End If
        End If
    Next i
    
    If iLineCount > 0 Then
        Sheet1.Range("A" & Sheet1.Cells(Sheet1.Rows.Count, 1).End(xlUp).Row + 1).Resize(iLineCount, UBound(arrResponseData, 1)).Value2 = WorksheetFunction.Transpose(arrResponseData)
    End If
    
    Erase arrResponseData
    
End Sub
'Find several delimiters for string search
Private Function MultiInStr(ByVal iStartChar As Long, ByVal searchText As String) As Long
    Dim iResult As Long
    Dim iTemp(0 To 2) As Long
    Dim i As Long
    
    iResult = InStr(iStartChar, searchText, " ")
    
    If iResult > 0 Then
        MultiInStr = iResult
        Exit Function
    End If
    
    iTemp(0) = InStr(iStartChar, searchText, "=")
    iTemp(1) = InStr(iStartChar, searchText, "<")
    iTemp(2) = InStr(iStartChar, searchText, ">")
    
    For i = 0 To UBound(iTemp)
        If iTemp(i) > 0 And (iTemp(i) < iResult Or iResult = 0) Then
            iResult = iTemp(i)
        End If
    Next i
    
    If iResult = 0 Then
        MultiInStr = iStartChar + 1
        Exit Function
    End If
    
    MultiInStr = iResult
    
End Function
'Check indent to determine nesting group
Private Function LeftTrim(ByRef srcText() As String) As Long()
    Dim iTabs() As Long
    Dim i As Long
    
    i = UBound(srcText)
    ReDim iTabs(0 To i)
    
    For i = 0 To UBound(srcText)
        iTabs(i) = (Len(srcText(i)) - Len(LTrim(srcText(i)))) / 4
    Next i
    
    LeftTrim = iTabs
    
End Function
'Loop through array and keep only label and conditional lines
Private Sub FindLabelsAndConditionals(ByRef srcText() As String)
    Dim outText() As String
    Dim i As Long
    Dim iLineCount As Long
    
    iLineCount = 0
    
    For i = 0 To UBound(srcText)
        If Left(srcText(i), 5) = "label" Then
            ReDim Preserve outText(0 To iLineCount)
            outText(iLineCount) = srcText(i)
            iLineCount = iLineCount + 1
        ElseIf Left(srcText(i), 2) = "if" Then
            ReDim Preserve outText(0 To iLineCount)
            outText(iLineCount) = srcText(i)
            iLineCount = iLineCount + 1
        End If
    Next i
    
    srcText = outText
    
    Erase outText
    
End Sub
'Clear previous entries
Private Sub ClearData()
    
    If Sheet1.Cells(Sheet1.Rows.Count, 1).End(xlUp).Row > 1 Then
        Sheet1.Range("A2:E" & Sheet1.Cells(Sheet1.Rows.Count, 1).End(xlUp).Row).ClearContents
    End If
    
End Sub
