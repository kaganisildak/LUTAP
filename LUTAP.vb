Class LUTAP
 Dim list As String = "ABCÇDEFGĞHIİJKLMNOÖPRSŞTUÜVYZ0123456789.,:"
    Function LUTAP()
        Dim a1, b1, c1, d1, e1, f1
        Dim a2, b2, c2, d2, e2, f2
        Dim a3, b3, c3, d3, e3, f3
        Dim a4, b4, c4, d4, e4, f4
        Dim a5, b5, c5, d5, e5, f5
        Dim a6, b6, c6, d6, e6, f6
        Dim a7, b7, c7, d7, e7, f7
        a1 = LUTAPRAN()
        b1 = LUTAPRAN()
        c1 = LUTAPRAN()
        d1 = LUTAPRAN()
        e1 = LUTAPRAN()
        f1 = LUTAPRAN()
        a2 = LUTAPRAN()
        b2 = LUTAPRAN()
        c2 = LUTAPRAN()
        d2 = LUTAPRAN()
        e2 = LUTAPRAN()
        f2 = LUTAPRAN()
        a3 = LUTAPRAN()
        b3 = LUTAPRAN()
        c3 = LUTAPRAN()
        d3 = LUTAPRAN()
        e3 = LUTAPRAN()
        f3 = LUTAPRAN()
        a4 = LUTAPRAN()
        b4 = LUTAPRAN()
        c4 = LUTAPRAN()
        d4 = LUTAPRAN()
        e4 = LUTAPRAN()
        f4 = LUTAPRAN()
        a5 = LUTAPRAN()
        b5 = LUTAPRAN()
        c5 = LUTAPRAN()
        d5 = LUTAPRAN()
        e5 = LUTAPRAN()
        f5 = LUTAPRAN()
        a6 = LUTAPRAN()
        b6 = LUTAPRAN()
        c6 = LUTAPRAN()
        d6 = LUTAPRAN()
        e6 = LUTAPRAN()
        f6 = LUTAPRAN()
        a7 = LUTAPRAN()
        b7 = LUTAPRAN()
        c7 = LUTAPRAN()
        d7 = LUTAPRAN()
        e7 = LUTAPRAN()
        f7 = LUTAPRAN()
        Return "0 1 2 3 4 5 6" & vbCrLf & "1 " & a1 & " " & b1 & " " & c1 & " " & d1 & " " & e1 & " " & f1 & " " & vbCrLf & "2 " & a2 & " " & b2 & " " & c2 & " " & d2 & " " & e2 & " " & f2 & " " & vbCrLf & "3 " & a3 & " " & b3 & " " & c3 & " " & d3 & " " & e3 & " " & f3 & " " & vbCrLf & "4 " & a4 & " " & b4 & " " & c4 & " " & d4 & " " & e4 & " " & f4 & " " & vbCrLf & "5 " & a5 & " " & b5 & " " & c5 & " " & d5 & " " & e5 & " " & f5 & " " & vbCrLf & "6 " & a6 & " " & b6 & " " & c6 & " " & d6 & " " & e6 & " " & f6 & " " & vbCrLf & "7 " & a7 & " " & b7 & " " & c7 & " " & d7 & " " & e7 & " " & f7 & " "
    End Function
    Function LUTAPRAN() As String
        Dim rand As New Random()
        Dim cikis = list.Substring(rand.Next(list.Length - 1), 1)
        list = list.Replace(cikis, "")
        Return cikis
    End Function
End Class
