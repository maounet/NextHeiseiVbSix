Attribute VB_Name = "Module1"

'VB6のFormatをオーバーライドして無理やり平成の次の元号に対応する。
'平成の次の次には対応していません。
Public Function Format(ByVal Expr As Variant, Optional dF As String, Optional fDw As VbDayOfWeek = vbSunday, Optional fWy As VbFirstWeekOfYear = vbFirstJan1) As Variant

    Const G3 As String = "新号"
    Const G2 As String = "新"
    Const G1 As String = "\N"

    If InStr(1, dF, "g", vbTextCompare) = 0 And InStr(1, dF, "e", vbTextCompare) = 0 Then
        '和暦に関係するFormatが存在しない
        Format = Strings.Format(Expr, dF, fDw, fWy)
    ElseIf Val(Strings.Format(Expr, "YYYYMMDD")) < 20190501 Then
        '日付変換した場合の日付が平成以前/そもそも日付変換できない
        Format = Strings.Format(Expr, dF, fDw, fWy)
    Else
        dF = Replace(dF, "ggg", G3, , , vbTextCompare)
        dF = Replace(dF, "gg", G2, , , vbTextCompare)
        dF = Replace(dF, "g", G1, , , vbTextCompare)
        
        Dim iE As Long
        Dim hE As String
        Dim hEE As String
        iE = Val(Strings.Format(Expr, "e")) - 30
        hE = Replace(Strings.Format(iE, "0"), "0", "\0")
        hEE = Replace(Strings.Format(iE, "00"), "0", "\0")
        
        dF = Replace(dF, "ee", hEE, , , vbTextCompare)
        dF = Replace(dF, "e", hE, , , vbTextCompare)
        
        '.がある場合にうまくいかない暫定対応
        dF = Replace(dF, ".", "\,")
        
        Format = Strings.Format(Expr, dF, fDw, fWy)
    End If
    
End Function

