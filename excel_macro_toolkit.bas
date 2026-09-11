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

        ElseIf IsNumeric(c.Value) And Not IsEmpty(c.Value) Then

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

    For Each c In Selection.Cells

        If c.HasFormula Then
            c.Formula = "=-(" & Mid(c.Formula, 2) & ")"
        ElseIf IsNumeric(c.Value) And Not IsEmpty(c.Value) Then
            c.Value = -c.Value
        End If

    Next c

End Sub

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

    ActiveSheet.Paste

    On Error Resume Next
    Selection.SpecialCells(xlCellTypeBlanks).Delete Shift:=xlUp
    On Error GoTo 0

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
        For Each rw In area.Rows

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

    MsgBox "Shortcuts activated"

End Sub
