Option Explicit
Private CollectedNumbers As String
Private Declare PtrSafe Function OpenClipboard Lib "user32" _
    (ByVal hwnd As LongPtr) As Long

Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long

Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long

Private Declare PtrSafe Function SetClipboardData Lib "user32" _
    (ByVal uFormat As Long, ByVal hMem As LongPtr) As LongPtr

Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" _
    (ByVal uFlags As Long, ByVal dwBytes As LongPtr) As LongPtr

Private Declare PtrSafe Function GlobalLock Lib "kernel32" _
    (ByVal hMem As LongPtr) As LongPtr

Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" _
    (ByVal hMem As LongPtr) As Long

Private Declare PtrSafe Function GlobalFree Lib "kernel32" _
    (ByVal hMem As LongPtr) As LongPtr

Private Declare PtrSafe Sub CopyMemory Lib "kernel32" _
    Alias "RtlMoveMemory" _
    (ByVal Destination As LongPtr, ByVal Source As LongPtr, _
     ByVal Length As LongPtr)
     
Private Declare PtrSafe Function GetClipboardData Lib "user32" _
    (ByVal uFormat As Long) As LongPtr

Private Declare PtrSafe Function lstrlenW Lib "kernel32" _
    (ByVal lpString As LongPtr) As Long

Sub DecimalRight()
    AdjustScale -1
End Sub

Sub DecimalLeft()
    AdjustScale 1
End Sub

Private Sub AdjustScale(ByVal Change As Long)

    Dim c As Range
    Dim rx As Object
    Dim matches As Object
    Dim BaseFormula As String
    Dim ScalePower As Long
    Dim CurrentFormula As String

    Set rx = CreateObject("VBScript.RegExp")

    'Recognizes formulas created by this macro:
    '=(original formula)*10^number
    rx.Pattern = "^=\((.*)\)\*10\^(-?\d+)$"
    rx.IgnoreCase = True
    rx.Global = False

    For Each c In Selection.Cells

        If c.HasFormula Then

            CurrentFormula = c.Formula

            If rx.Test(CurrentFormula) Then
                Set matches = rx.Execute(CurrentFormula)

                BaseFormula = matches(0).SubMatches(0)
                ScalePower = CLng(matches(0).SubMatches(1))
            Else
                BaseFormula = Mid$(CurrentFormula, 2)
                ScalePower = 0
            End If

            ScalePower = ScalePower + Change

            If ScalePower = 0 Then
                'Restore the original formula
                c.Formula = "=" & BaseFormula
            Else
                c.Formula = "=(" & BaseFormula & ")*10^" & ScalePower
            End If

        ElseIf IsNumeric(c.value) And Not IsEmpty(c.value) Then

            c.Formula = "=(" & Trim$(Str$(c.Value2)) & ")*10^" & Change

        End If

    Next c

End Sub
Sub AutoColor()

    Dim c As Range
    Dim FirstColor As Long
    Dim MixedColors As Boolean

    'Get the color of the first cell
    FirstColor = Selection.Cells(1, 1).Font.Color

    'Check whether selected cells have different colors
    For Each c In Selection
        If c.Font.Color <> FirstColor Then
            MixedColors = True
            Exit For
        End If
    Next c

    'If colors are mixed, reset everything to black
    If MixedColors Then
        Selection.Font.Color = RGB(0, 0, 0)
        Exit Sub
    End If

    'Otherwise cycle the entire selection together
    Select Case FirstColor

        Case RGB(0, 0, 0)          'Black
            Selection.Font.Color = RGB(0, 0, 255)      'Blue

        Case RGB(0, 0, 255)        'Blue
            Selection.Font.Color = RGB(0, 176, 80)     'Green

        Case Else                  'Green or anything else
            Selection.Font.Color = RGB(0, 0, 0)        'Black

    End Select

End Sub

Sub ToggleSigns()

    Dim c As Range
    Dim expr As String
    Dim wasNegative As Boolean

    For Each c In Selection.Cells

        If c.HasFormula Then

            expr = Mid$(c.Formula, 2)
            wasNegative = False

            'Remove existing outer negations, tracking their net effect.
            Do While HasOuterNegation(expr)
                expr = Mid$(expr, 3, Len(expr) - 3)
                wasNegative = Not wasNegative
            Loop

            'Toggle the sign using at most one outer negation.
            If wasNegative Then
                c.Formula = "=" & expr
            Else
                c.Formula = "=-(" & expr & ")"
            End If

        ElseIf Not IsError(c.Value2) Then
            If Not IsEmpty(c.Value2) And IsNumeric(c.Value2) Then
                c.Value2 = -c.Value2
            End If
        End If

    Next c

End Sub

Private Function HasOuterNegation(ByVal expr As String) As Boolean

    Dim i As Long
    Dim depth As Long
    Dim ch As String
    Dim quote As String

    If Left$(expr, 2) <> "-(" Then Exit Function
    If Right$(expr, 1) <> ")" Then Exit Function

    For i = 2 To Len(expr)
        ch = Mid$(expr, i, 1)

        'Ignore parentheses inside text or quoted sheet names.
        If quote <> "" Then
            If ch = quote Then quote = ""
        Else
            Select Case ch
                Case """", "'"
                    quote = ch
                Case "("
                    depth = depth + 1
                Case ")"
                    depth = depth - 1

                    'The outer parentheses must enclose everything.
                    If depth = 0 Then
                        HasOuterNegation = (i = Len(expr))
                        Exit Function
                    End If
            End Select
        End If
    Next i

End Function

Sub FormatPercent()
    Selection.NumberFormat = "0.0%_);(0.0%);@_)"
End Sub

Sub FormatMultiples()
    Selection.NumberFormat = "0.0x_);(0.0x);@_)"
End Sub

Sub FormatCurrency()
    Selection.NumberFormat = "$#,##0.00_);($#,##0.00);@_)"
End Sub

Sub FormatNumber()
    Selection.NumberFormat = "#,##0_);(#,##0);@_)"
End Sub

Sub PasteSkipBlanks()

    Dim Text As String
    Dim Lines As Variant
    Dim Fields As Variant
    Dim Line As Variant
    Dim RowsToPaste As Collection
    Dim Output() As Variant
    Dim RowCount As Long
    Dim ColumnCount As Long
    Dim r As Long
    Dim j As Long
    Dim CheckText As String
    Dim target As Range
    Dim PossibleError() As Boolean
    Dim LooksNumeric As Object

    If TypeName(Selection) <> "Range" Then Exit Sub

    On Error GoTo PasteFailed

    Text = ReadClipboardText()

    If Len(Text) = 0 Then
        MsgBox "Couldn't read clipboard text. Copy it again.", vbExclamation
        Exit Sub
    End If

    Text = Replace(Text, vbCrLf, vbLf)
    Text = Replace(Text, vbCr, vbLf)
    Lines = Split(Text, vbLf)

    Set RowsToPaste = New Collection

    For Each Line In Lines

        'Skip lines containing only whitespace.
        CheckText = Replace(CStr(Line), vbTab, "")
        CheckText = Replace(CheckText, ChrW(160), " ")

        If Len(Trim$(CheckText)) > 0 Then
            RowsToPaste.Add CStr(Line)

            Fields = Split(CStr(Line), vbTab)

            If UBound(Fields) + 1 > ColumnCount Then
                ColumnCount = UBound(Fields) + 1
            End If
        End If

    Next Line

    RowCount = RowsToPaste.Count
    If RowCount = 0 Then Exit Sub

    If RowCount > ActiveSheet.rows.Count - ActiveCell.Row + 1 _
       Or ColumnCount > ActiveSheet.Columns.Count - ActiveCell.Column + 1 Then
        MsgBox "There isn't enough room to paste here.", vbExclamation
        Exit Sub
    End If

    ReDim Output(1 To RowCount, 1 To ColumnCount)
    ReDim PossibleError(1 To RowCount, 1 To ColumnCount)

    Set LooksNumeric = CreateObject("VBScript.RegExp")
    LooksNumeric.Pattern = "[0-9]"

    For r = 1 To RowCount
        Fields = Split(RowsToPaste(r), vbTab)

        For j = 0 To UBound(Fields)
            Dim Entry As String
            Dim NumberCheck As Object
            
            Entry = Trim$(Fields(j))
            
            Set NumberCheck = CreateObject("VBScript.RegExp")
            NumberCheck.Pattern = "^[+-]?([0-9]+|[0-9]{1,3}(,[0-9]{3})+)(\.[0-9]+)?$"
            
            If NumberCheck.Test(Entry) Then
                Output(r, j + 1) = Val(Replace(Entry, ",", ""))
            Else
                'Flag entries containing digits but no letters.
                PossibleError(r, j + 1) = LooksNumeric.Test(Entry) _
                                         And Not (LCase$(Entry) Like "*[a-z]*")
            
                Select Case Left$(Entry, 1)
                    Case "=", "+", "-", "@"
                        Output(r, j + 1) = "'" & Entry
                    Case Else
                        Output(r, j + 1) = Entry
                End Select
            End If
        Next j
    Next r

    Set target = ActiveCell.Resize(RowCount, ColumnCount)
    target.Value2 = Output
    For r = 1 To RowCount
        For j = 1 To ColumnCount
            If PossibleError(r, j) Then
                target.Cells(r, j).Interior.Color = RGB(255, 199, 206)
            End If
        Next j
    Next r
    Exit Sub

PasteFailed:
    MsgBox "Couldn't paste: " & Err.Description, vbExclamation

End Sub

Sub IncreaseDecimals()

    Application.CommandBars.ExecuteMso "DecimalsIncrease"

End Sub

Sub DecreaseDecimals()

    Application.CommandBars.ExecuteMso "DecimalsDecrease"

End Sub

Sub AddDashForZero()

    Dim c As Range
    Dim parts As Variant
    Dim fmt As String

    For Each c In Selection.Cells

        fmt = c.NumberFormat
        parts = Split(fmt, ";")

        Select Case UBound(parts)

            Case 3
                'Positive; Negative; Zero; Text
                parts(2) = "-_)"
                c.NumberFormat = Join(parts, ";")

            Case 2
                'If third section is text, insert a zero section before it
                If InStr(parts(2), "@") > 0 Then
                    c.NumberFormat = parts(0) & ";" & _
                                     parts(1) & ";-_);" & _
                                     parts(2)
                Else
                    'Third section is already the zero section
                    parts(2) = "-_)"
                    c.NumberFormat = Join(parts, ";")
                End If

            Case 1
                'Positive; Negative
                c.NumberFormat = parts(0) & ";" & _
                                 parts(1) & ";-_);@_)"

        End Select

    Next c

End Sub
Sub CopyCellValueOnly()

    Dim area As Range
    Dim rw As Range
    Dim c As Range
    Dim v As Variant
    Dim ClipboardText As String
    Dim RowText As String
    Dim NewNumbers As String
    Dim FirstRow As Boolean
    Dim FirstCell As Boolean

    If TypeName(Selection) <> "Range" Then Exit Sub

    FirstRow = True

    For Each area In Selection.Areas
        For Each rw In area.rows

            RowText = ""
            FirstCell = True

            For Each c In rw.Cells

                v = c.Value2

                If Not FirstCell Then RowText = RowText & vbTab
                FirstCell = False

                'Copy values, not formulas, to the clipboard.
                If IsError(v) Then
                    RowText = RowText & c.Text
                ElseIf Not IsEmpty(v) Then
                    RowText = RowText & CStr(v)

                    'Only numbers go into the accumulated SUM list.
                    If IsNumeric(v) Then
                        If Len(NewNumbers) > 0 Then
                            NewNumbers = NewNumbers & ","
                        End If

                        NewNumbers = NewNumbers & Trim$(Str$(CDbl(v)))
                    End If
                End If

            Next c

            If Not FirstRow Then
                ClipboardText = ClipboardText & vbCrLf
            End If

            ClipboardText = ClipboardText & RowText
            FirstRow = False

        Next rw
    Next area

    If Not PutTextOnClipboard(ClipboardText) Then
        MsgBox "Couldn't copy to the clipboard. Try again." & _
               vbCrLf & "No numbers were added to the list.", vbExclamation
        Exit Sub
    End If

    If Len(NewNumbers) > 0 Then
        If Len(CollectedNumbers) > 0 Then
            CollectedNumbers = CollectedNumbers & ","
        End If

        CollectedNumbers = CollectedNumbers & NewNumbers
    End If

End Sub
Private Function PutTextOnClipboard(ByVal Text As String) As Boolean

    Dim hMem As LongPtr
    Dim pMem As LongPtr
    Dim ByteCount As LongPtr
    Dim ClipboardIsOpen As Boolean

    On Error GoTo CleanUp

    'Allocate zero-initialized memory, including a Unicode terminator.
    ByteCount = LenB(Text) + 2
    hMem = GlobalAlloc(&H42, ByteCount)
    If hMem = 0 Then GoTo CleanUp

    pMem = GlobalLock(hMem)
    If pMem = 0 Then GoTo CleanUp

    If LenB(Text) > 0 Then
        CopyMemory pMem, StrPtr(Text), LenB(Text)
    End If

    GlobalUnlock hMem

    If OpenClipboard(Application.hwnd) = 0 Then GoTo CleanUp
    ClipboardIsOpen = True

    If EmptyClipboard() = 0 Then GoTo CleanUp

    '13 = Unicode text.
    If SetClipboardData(13, hMem) = 0 Then GoTo CleanUp

    'Windows now owns this memory.
    hMem = 0
    PutTextOnClipboard = True

CleanUp:
    If ClipboardIsOpen Then CloseClipboard
    If hMem <> 0 Then GlobalFree hMem

End Function

Sub SumClipboardNumbers()

    If Len(CollectedNumbers) = 0 Then
        MsgBox "No numbers collected yet."
        Exit Sub
    End If

    ActiveCell.Formula = "=SUM(" & CollectedNumbers & ")"

    'Clear the list after inserting the formula
    CollectedNumbers = ""

End Sub

Sub ClearCollectedNumbers()

    CollectedNumbers = ""

    MsgBox "Collected numbers cleared."

End Sub

Private Function ReadClipboardText() As String

    Dim hMem As LongPtr
    Dim pMem As LongPtr
    Dim CharCount As Long
    Dim Text As String
    Dim ClipboardIsOpen As Boolean

    On Error GoTo CleanUp

    If OpenClipboard(Application.hwnd) = 0 Then GoTo CleanUp
    ClipboardIsOpen = True

    '13 = Unicode text.
    hMem = GetClipboardData(13)
    If hMem = 0 Then GoTo CleanUp

    pMem = GlobalLock(hMem)
    If pMem = 0 Then GoTo CleanUp

    CharCount = lstrlenW(pMem)

    If CharCount > 0 Then
        Text = String$(CharCount, vbNullChar)
        CopyMemory StrPtr(Text), pMem, LenB(Text)
        ReadClipboardText = Text
    End If

CleanUp:
    If pMem <> 0 Then GlobalUnlock hMem
    If ClipboardIsOpen Then CloseClipboard

End Function

Sub PasteTextAcrossColumns()

    Dim Text As String
    Dim Lines As Variant
    Dim Item As Variant
    Dim Entry As String
    Dim Items As Collection
    Dim Output() As Variant
    Dim i As Long
    Dim target As Range
    Dim rx As Object
    Dim PossibleError() As Boolean
    Dim LooksNumeric As Object

    If TypeName(Selection) <> "Range" Then Exit Sub

    On Error GoTo PasteFailed

    Text = ReadClipboardText()

    If Len(Text) = 0 Then
        MsgBox "Couldn't read text from the clipboard. Copy it again.", _
               vbExclamation
        Exit Sub
    End If

    'Normalize Windows and other line endings.
    Text = Replace(Text, vbCrLf, vbLf)
    Text = Replace(Text, vbCr, vbLf)

    Lines = Split(Text, vbLf)
    Set Items = New Collection

    For Each Item In Lines
        Entry = CStr(Item)
    
        'Remove dollar signs and normalize spaces.
        Entry = Replace(Entry, "$", "")
        Entry = Replace(Entry, ChrW(160), " ")
        Entry = Trim$(Entry)
    
        'Remove any existing leading apostrophes.
        Do While Len(Entry) > 0
            If Left$(Entry, 1) <> "'" Then Exit Do
            Entry = Trim$(Mid$(Entry, 2))
        Loop
    
        'Skip blank lines, including lines that contained only $.
        If Len(Entry) > 0 Then Items.Add Entry
    Next Item

    If Items.Count = 0 Then Exit Sub

    If Items.Count > ActiveSheet.Columns.Count - ActiveCell.Column + 1 Then
        MsgBox "There aren't enough columns to paste these entries.", _
               vbExclamation
        Exit Sub
    End If

    'Recognize numbers using US-style commas and decimal points.
    Set rx = CreateObject("VBScript.RegExp")
    rx.Pattern = "^[+-]?([0-9]+|[0-9]{1,3}(,[0-9]{3})+)(\.[0-9]+)?$"

    ReDim Output(1 To 1, 1 To Items.Count)
    ReDim PossibleError(1 To Items.Count)

    Set LooksNumeric = CreateObject("VBScript.RegExp")
    LooksNumeric.Pattern = "[0-9]"

    For i = 1 To Items.Count
        Entry = Items(i)
    
        If rx.Test(Entry) Then
            Output(1, i) = Val(Replace(Entry, ",", ""))
        Else
            'Keep entries that fail the number check as text.
            Output(1, i) = "'" & Entry
    
            'Flag entries containing digits but no letters.
            PossibleError(i) = LooksNumeric.Test(Entry) _
                               And Not (LCase$(Entry) Like "*[a-z]*")
        End If
    Next i
    
    Set target = ActiveCell.Resize(1, Items.Count)
    target.Value2 = Output
    
    For i = 1 To Items.Count
        If PossibleError(i) Then
            With target.Cells(1, i)
                .Interior.Color = RGB(255, 199, 206) 'Light red
            End With
        End If
    Next i

    Exit Sub

PasteFailed:
    MsgBox "Couldn't paste: " & Err.Description, vbExclamation

End Sub

Sub SetCustomShortcuts()

    Application.OnKey "%+{.}", "DecimalRight"  'Alt + Shift + .
    Application.OnKey "%+{,}", "DecimalLeft"   'Alt + Shift + ,
    Application.OnKey "^%a", "AutoColor"
    Application.OnKey "^%s", "ToggleSigns"
    Application.OnKey "^%4", "FormatCurrency"    '$
    Application.OnKey "^%5", "FormatPercent"     '%'
    Application.OnKey "^%6", "FormatMultiples"   'x
    Application.OnKey "^%7", "FormatNumber"      '#
    Application.OnKey "^%q", "PasteSkipBlanks"   'Ctrl + Alt + Q
    Application.OnKey "^%{[}", "DecreaseDecimals"
    Application.OnKey "^%{]}", "IncreaseDecimals"
    Application.OnKey "^%8", "AddDashForZero"
    Application.OnKey "%+c", "CopyCellValueOnly"  'Alt + Shift + C
    Application.OnKey "%+v", "SumClipboardNumbers"
    Application.OnKey "%+x", "ClearCollectedNumbers" 'Clear list
    Application.OnKey "%+z", "PasteTextAcrossColumns" 'Alt + Shift + Z

    MsgBox "Shortcuts activated"

End Sub
