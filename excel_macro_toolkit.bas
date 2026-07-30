Option Explicit

Sub DecimalRight()

    Dim c As Range

    For Each c In Selection.Cells
        If Not c.HasFormula Then
            If IsNumeric(c.Value) And Not IsEmpty(c.Value) Then
                c.Value = c.Value / 10
            End If
        End If
    Next c

End Sub

Sub DecimalLeft()

    Dim c As Range

    For Each c In Selection.Cells
        If Not c.HasFormula Then
            If IsNumeric(c.Value) And Not IsEmpty(c.Value) Then
                c.Value = c.Value * 10
            End If
        End If
    Next c

End Sub

Sub AutoColor()

    Dim c As Range

    For Each c In Selection

        Select Case c.Font.Color

            Case RGB(0, 0, 0)          'Black
                c.Font.Color = RGB(0, 0, 255)      'Blue

            Case RGB(0, 0, 255)        'Blue
                c.Font.Color = RGB(0, 176, 80)     'Green

            Case Else                  'Green or anything else
                c.Font.Color = RGB(0, 0, 0)        'Black

        End Select

    Next c

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

Sub SetCustomShortcuts()

    Application.OnKey "%+{.}", "DecimalRight"
    Application.OnKey "%+{,}", "DecimalLeft"
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

    MsgBox "Shortcuts activated"

End Sub
