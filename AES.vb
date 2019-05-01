Module AES
    Function BtoC_2(val As String)

        Dim ArrString(val.Length) As String
        Dim StackInteger As Integer = 0
        Dim splitbit(8) As Integer
        Dim DtoC As String = ""
        Dim counter As Integer = 0
        For i = 1 To val.Length Step 8
            counter += 1
            StackInteger = 0
            ArrString(counter) = Mid(val, i, 8)
            For j = 1 To 8
                splitbit(j) = Convert.ToInt32(Mid(ArrString(counter), j, 1))
            Next


            For j = 1 To 8
                StackInteger += Math.Pow(2, 8 - j) * splitbit(j)
            Next
            DtoC += Chr(StackInteger)
        Next


        Return DtoC
    End Function

    Public Function tobin8(ByVal a As Long) As String
        Dim b As String = ""
        For i = 1 To 8
            b = a Mod 2 & b
            a = a \ 2
        Next
        Return b
    End Function
    Public Function tobin4(ByVal a As Long) As String
        Dim b As String = ""
        For i = 1 To 4
            b = a Mod 2 & b
            a = a \ 2
        Next
        Return b
    End Function
    Public Function Xor128(ByVal a As String, ByVal b As String) As String
        Dim x As String = ""
        For i = 1 To 128
            x = x & (Val(Mid(a, i, 1)) Xor Val(Mid(b, i, 1))).ToString
        Next
        Return x
    End Function
    Public Function int(ByVal a As String) As Long
        Dim x As Long
        Dim z As Integer : z = 0
        For i = Len(a) To 1 Step -1
            x = (Val(Mid(a, i, 1))) * (2 ^ z) + x
            z = z + 1
        Next
        Return (x)
    End Function
    Public Function Aes_shift(ByVal s, ByVal b, ByVal m)
        Dim l = Len(s)
        If m = "l" Then
            For j = 0 To b - 1
                Dim t = Mid(s, 1, 1)
                For i = 1 To l - 1
                    Mid(s, i, 1) = Mid(s, i + 1, 1)
                Next
                Mid(s, l, 1) = t
            Next

        ElseIf m = "r" Then
            For j = 0 To b - 1
                Dim t = Mid(s, l, 1)
                For i = l To 2 Step -1
                    Mid(s, i, 1) = Mid(s, i - 1, 1)
                Next
                Mid(s, 1, 1) = t
            Next
        End If
        Return (s)
    End Function
    Public Function binaryhex(ByVal hexnumber)
        Dim BinaryNumber As String = ""
        Dim Number As Integer = CInt("&H" & hexnumber)
        For i = 1 To 8
            BinaryNumber = (Number Mod 2).ToString & BinaryNumber
            Number = Number \ 2
        Next
        Return BinaryNumber
    End Function
    Public Function xorhex(ByVal a As String, ByVal b As String, ByVal c As String, ByVal d As String) As String
        Dim x As String = ""
        For i = 1 To 8
            x = x & (Val(Mid(a, i, 1)) Xor Val(Mid(b, i, 1)) Xor Val(Mid(c, i, 1)) Xor Val(Mid(d, i, 1))).ToString
        Next
        Return x
    End Function
    Public Function mixcolumns(ByVal ab As String)
        Dim ee(,) As String = {{"01", "03", "05", "0F", "11", "33", "55", "FF", "1A", "2E", "72", "96", "A1", "F8", "13", "35"}, {"5F", "E1", "38", "48", "D8", "73", "95", "A4", "f7", "02", "06", "0A", "1E", "22", "66", "AA"}, {"E5", "34", "5C", "E4", "37", "59", "EB", "26", "6A", "BE", "D9", "70", "90", "AB", "E6", "31"}, {"53", "F5", "04", "0C", "14", "3C", "44", "CC", "4F", "D1", "68", "B8", "D3", "6E", "B2", "CD"}, {"4C", "D4", "67", "A9", "E0", "3B", "4D", "D7", "62", "A6", "F1", "08", "18", "28", "78", "88"}, {"83", "9E", "B9", "D0", "6B", "BD", "DC", "7F", "81", "98", "B3", "CE", "49", "DB", "76", "9A"}, {"B5", "C4", "57", "F9", "10", "30", "50", "F0", "0B", "1D", "27", "69", "BB", "D6", "61", "A3"}, {"FE", "19", "2B", "7D", "87", "92", "AD", "EC", "2F", "71", "93", "AE", "E9", "20", "60", "A0"}, {"FB", "16", "3A", "4E", "D2", "6D", "B7", "C2", "5D", "E7", "32", "56", "FA", "15", "3F", "41"}, {"C3", "5E", "E2", "3D", "47", "C9", "40", "C0", "5B", "ED", "2C", "74", "9C", "BF", "DA", "75"}, {"9F", "BA", "D5", "64", "AC", "EF", "2A", "7E", "82", "9D", "BC", "DF", "7A", "8E", "89", "80"}, {"9B", "B6", "C1", "58", "E8", "23", "65", "AF", "EA", "25", "6F", "B1", "C8", "43", "C5", "54"}, {"FC", "1F", "21", "63", "A5", "F4", "07", "09", "1B", "2D", "77", "99", "B0", "CB", "46", "CA"}, {"45", "CF", "4A", "DE", "79", "8B", "86", "91", "A8", "E3", "3E", "42", "C6", "51", "F3", "0E"}, {"12", "36", "5A", "EE", "29", "7B", "8D", "8C", "8F", "8A", "85", "94", "A7", "F2", "0D", "17"}, {"39", "4B", "DD", "7C", "84", "97", "A2", "FD", "1C", "24", "6C", "B4", "C7", "52", "F6", "01"}}
        Dim l(,) As String = {{"", "00", "19", "01", "32", "02", "1A", "C6", "4B", "C7", "1B", "68", "33", "EE", "DF", "03"}, {"64", "04", "E0", "0E", "34", "8D", "81", "EF", "4C", "71", "08", "C8", "F8", "69", "1C", "C1"}, {"7D", "C2", "1D", "B5", "F9", "B9", "27", "6A", "4D", "E4", "A6", "72", "9A", "C9", "09", "78"}, {"65", "2F", "8A", "05", "21", "0F", "E1", "24", "12", "F0", "82", "45", "35", "93", "DA", "8E"}, {"96", "8F", "DB", "BD", "36", "D0", "CE", "94", "13", "5C", "D2", "F1", "40", "46", "83", "38"}, {"66", "DD", "FD", "30", "BF", "06", "8B", "62", "B3", "25", "E2", "98", "22", "88", "91", "10"}, {"7E", "6E", "48", "C3", "A3", "B6", "1E", "42", "3A", "6B", "28", "54", "FA", "85", "3D", "BA"}, {"2B", "79", "0A", "15", "9B", "9F", "5E", "CA", "4E", "D4", "AC", "E5", "F3", "73", "A7", "57"}, {"AF", "58", "A8", "50", "F4", "EA", "D6", "74", "4F", "AE", "E9", "D5", "E7", "E6", "AD", "E8"}, {"2C", "D7", "75", "7A", "EB", "16", "0B", "F5", "59", "CB", "5F", "B0", "9C", "A9", "51", "A0"}, {"7F", "0C", "F6", "6F", "17", "C4", "49", "EC", "D8", "43", "1F", "2D", "A4", "76", "7B", "B7"}, {"CC", "BB", "3E", "5A", "FB", "60", "B1", "86", "3B", "52", "A1", "6C", "AA", "55", "29", "9D"}, {"97", "B2", "87", "90", "61", "BE", "DC", "FC", "BC", "95", "CF", "CD", "37", "3F", "5B", "D1"}, {"53", "39", "84", "3C", "41", "A2", "6D", "47", "14", "2A", "9E", "5D", "56", "F2", "D3", "AB"}, {"44", "11", "92", "D9", "23", "20", "2E", "89", "B4", "7C", "B8", "26", "77", "99", "E3", "A5"}, {"67", "4A", "ED", "DE", "C5", "31", "FE", "18", "0D", "63", "8C", "80", "C0", "F7", "70", "07"}}
        Dim a(,) As String = {{"02", "03", "01", "01"}, {"01", "02", "03", "01"}, {"01", "01", "02", "03"}, {"03", "01", "01", "02"}}
        Dim x, n, r, v, d As String : d = ""
        Dim m, m1, f, f1, k, kk, j, jj, mn As Integer : j = 0 : jj = 0
        Dim c, w, hex1, hex2, p, u, u1, y, t, q As String : y = "" : t = ""
        For mn = 0 To 3
            For i = 1 To Len(ab) Step +8
                c = Mid(ab, i, 8)
                If c = "00000000" Then
                    t += c
                    GoTo 1
                End If
                m = int(Mid(c, 1, 4))
                m1 = int(Mid(c, 5, 4))
                hex1 = l(m, m1)
                w = a(jj, j)
                If w = "01" Then
                    t += c
                    GoTo 1
                End If
                f = (Mid(w, 1, 1))
                f1 = (Mid(w, 2, 1))
                hex2 = l(f, f1)
                Dim int1 As Integer = Convert.ToInt32(hex1, 16)
                Dim int2 As Integer = Convert.ToInt32(hex2, 16)
                Dim result As Integer = (int1 + int2) Mod &HFF
                Dim hexResult As String = Convert.ToString(result, 16)
                q = Len(hexResult)
                If (q = 1) Then
                    hexResult = "0" + hexResult
                End If
                u = Mid(hexResult, 1, 1)
                u1 = Mid(hexResult, 2, 1)
                If u = "a" Then
                    kk = "10"
                ElseIf u = "b" Then
                    kk = "11"
                ElseIf u = "c" Then
                    kk = "12"
                ElseIf u = "d" Then
                    kk = "13"
                ElseIf u = "e" Then
                    kk = "14"
                ElseIf u = "f" Then
                    kk = "15"
                Else
                    kk = u
                End If
                If u1 = "a" Then
                    k = "10"
                ElseIf u1 = "b" Then
                    k = "11"
                ElseIf u1 = "c" Then
                    k = "12"
                ElseIf u1 = "d" Then
                    k = "13"
                ElseIf u1 = "e" Then
                    k = "14"
                ElseIf u1 = "f" Then
                    k = "15"
                Else
                    k = Val(u1)
                End If
                p = ee(Val(kk), Val(k))
                y = binaryhex(p)
                t += y
1:              j = j + 1
            Next
            x = Mid(t, 1, 8)
            n = Mid(t, 9, 8)
            r = Mid(t, 17, 8)
            v = Mid(t, 25, 8)
            d += xorhex(x, n, r, v)
            jj = jj + 1
            j = 0
            t = ""
        Next
        Return d
    End Function
    Public Function mixcolumnsinverse(ByVal s As String)
        Dim ee(,) As String = {{"01", "03", "05", "0F", "11", "33", "55", "FF", "1A", "2E", "72", "96", "A1", "F8", "13", "35"}, {"5F", "E1", "38", "48", "D8", "73", "95", "A4", "f7", "02", "06", "0A", "1E", "22", "66", "AA"}, {"E5", "34", "5C", "E4", "37", "59", "EB", "26", "6A", "BE", "D9", "70", "90", "AB", "E6", "31"}, {"53", "F5", "04", "0C", "14", "3C", "44", "CC", "4F", "D1", "68", "B8", "D3", "6E", "B2", "CD"}, {"4C", "D4", "67", "A9", "E0", "3B", "4D", "D7", "62", "A6", "F1", "08", "18", "28", "78", "88"}, {"83", "9E", "B9", "D0", "6B", "BD", "DC", "7F", "81", "98", "B3", "CE", "49", "DB", "76", "9A"}, {"B5", "C4", "57", "F9", "10", "30", "50", "F0", "0B", "1D", "27", "69", "BB", "D6", "61", "A3"}, {"FE", "19", "2B", "7D", "87", "92", "AD", "EC", "2F", "71", "93", "AE", "E9", "20", "60", "A0"}, {"FB", "16", "3A", "4E", "D2", "6D", "B7", "C2", "5D", "E7", "32", "56", "FA", "15", "3F", "41"}, {"C3", "5E", "E2", "3D", "47", "C9", "40", "C0", "5B", "ED", "2C", "74", "9C", "BF", "DA", "75"}, {"9F", "BA", "D5", "64", "AC", "EF", "2A", "7E", "82", "9D", "BC", "DF", "7A", "8E", "89", "80"}, {"9B", "B6", "C1", "58", "E8", "23", "65", "AF", "EA", "25", "6F", "B1", "C8", "43", "C5", "54"}, {"FC", "1F", "21", "63", "A5", "F4", "07", "09", "1B", "2D", "77", "99", "B0", "CB", "46", "CA"}, {"45", "CF", "4A", "DE", "79", "8B", "86", "91", "A8", "E3", "3E", "42", "C6", "51", "F3", "0E"}, {"12", "36", "5A", "EE", "29", "7B", "8D", "8C", "8F", "8A", "85", "94", "A7", "F2", "0D", "17"}, {"39", "4B", "DD", "7C", "84", "97", "A2", "FD", "1C", "24", "6C", "B4", "C7", "52", "F6", "01"}}
        Dim l(,) As String = {{"", "00", "19", "01", "32", "02", "1A", "C6", "4B", "C7", "1B", "68", "33", "EE", "DF", "03"}, {"64", "04", "E0", "0E", "34", "8D", "81", "EF", "4C", "71", "08", "C8", "F8", "69", "1C", "C1"}, {"7D", "C2", "1D", "B5", "F9", "B9", "27", "6A", "4D", "E4", "A6", "72", "9A", "C9", "09", "78"}, {"65", "2F", "8A", "05", "21", "0F", "E1", "24", "12", "F0", "82", "45", "35", "93", "DA", "8E"}, {"96", "8F", "DB", "BD", "36", "D0", "CE", "94", "13", "5C", "D2", "F1", "40", "46", "83", "38"}, {"66", "DD", "FD", "30", "BF", "06", "8B", "62", "B3", "25", "E2", "98", "22", "88", "91", "10"}, {"7E", "6E", "48", "C3", "A3", "B6", "1E", "42", "3A", "6B", "28", "54", "FA", "85", "3D", "BA"}, {"2B", "79", "0A", "15", "9B", "9F", "5E", "CA", "4E", "D4", "AC", "E5", "F3", "73", "A7", "57"}, {"AF", "58", "A8", "50", "F4", "EA", "D6", "74", "4F", "AE", "E9", "D5", "E7", "E6", "AD", "E8"}, {"2C", "D7", "75", "7A", "EB", "16", "0B", "F5", "59", "CB", "5F", "B0", "9C", "A9", "51", "A0"}, {"7F", "0C", "F6", "6F", "17", "C4", "49", "EC", "D8", "43", "1F", "2D", "A4", "76", "7B", "B7"}, {"CC", "BB", "3E", "5A", "FB", "60", "B1", "86", "3B", "52", "A1", "6C", "AA", "55", "29", "9D"}, {"97", "B2", "87", "90", "61", "BE", "DC", "FC", "BC", "95", "CF", "CD", "37", "3F", "5B", "D1"}, {"53", "39", "84", "3C", "41", "A2", "6D", "47", "14", "2A", "9E", "5D", "56", "F2", "D3", "AB"}, {"44", "11", "92", "D9", "23", "20", "2E", "89", "B4", "7C", "B8", "26", "77", "99", "E3", "A5"}, {"67", "4A", "ED", "DE", "C5", "31", "FE", "18", "0D", "63", "8C", "80", "C0", "F7", "70", "07"}}
        Dim a1(,) As String = {{"0e", "0b", "0d", "09"}, {"09", "0e", "0b", "0d"}, {"0d", "09", "0e", "0b"}, {"0b", "0d", "09", "0e"}}
        Dim x, n, r, v, d, c, w, hex1, hex2, p, u, u1, y, t, f, f1, z, z1, q As String : d = "" : y = "" : t = ""
        Dim m, m1, k, kk, j, jj, mn As Integer : j = 0 : jj = 0
        For mn = 0 To 3
            For i = 1 To Len(s) Step +8
                c = Mid(s, i, 8)
                If c = "00000000" Then
                    y = c
                    GoTo 2
                End If
                m = int(Mid(c, 1, 4))
                m1 = int(Mid(c, 5, 4))
                hex1 = l(m, m1)
                w = a1(jj, j)
                f = (Mid(w, 1, 1))
                f1 = (Mid(w, 2, 1))
                If f = "a" Then
                    z = "10"
                ElseIf f = "b" Then
                    z = "11"
                ElseIf f = "c" Then
                    z = "12"
                ElseIf f = "d" Then
                    z = "13"
                ElseIf f = "e" Then
                    z = "14"
                ElseIf f = "f" Then
                    z = "15"
                Else
                    z = f
                End If
                If f1 = "a" Then
                    z1 = "10"
                ElseIf f1 = "b" Then
                    z1 = "11"
                ElseIf f1 = "c" Then
                    z1 = "12"
                ElseIf f1 = "d" Then
                    z1 = "13"
                ElseIf f1 = "e" Then
                    z1 = "14"
                ElseIf f1 = f Then
                    z1 = "15"
                Else
                    z1 = f1
                End If
                hex2 = l(z, z1)
                Dim int1 As Integer = Convert.ToInt32(hex1, 16)
                Dim int2 As Integer = Convert.ToInt32(hex2, 16)
                Dim result As Integer = (int1 + int2) Mod &HFF
                Dim hexResult As String = Convert.ToString(result, 16)
                q = Len(hexResult)
                If (q = 1) Then
                    hexResult = "0" + hexResult
                End If
                u = Mid(hexResult, 1, 1)
                u1 = Mid(hexResult, 2, 1)
                If u = "a" Then
                    kk = "10"
                ElseIf u = "b" Then
                    kk = "11"
                ElseIf u = "c" Then
                    kk = "12"
                ElseIf u = "d" Then
                    kk = "13"
                ElseIf u = "e" Then
                    kk = "14"
                ElseIf u = "f" Then
                    kk = "15"
                Else
                    kk = u
                End If
                If u1 = "a" Then
                    k = "10"
                ElseIf u1 = "b" Then
                    k = "11"
                ElseIf u1 = "c" Then
                    k = "12"
                ElseIf u1 = "d" Then
                    k = "13"
                ElseIf u1 = "e" Then
                    k = "14"
                ElseIf u1 = "f" Then
                    k = "15"
                ElseIf u1 = "0" Then
                    p = ee(Val(kk), 0)
                    GoTo 1
                Else
                    k = Val(u1)
                End If
                p = ee(Val(kk), Val(k))
1:              y = binaryhex(p)
2:              t += y
                j = j + 1
            Next
            x = Mid(t, 1, 8)
            n = Mid(t, 9, 8)
            r = Mid(t, 17, 8)
            v = Mid(t, 25, 8)
            d += xorhex(x, n, r, v)
            jj = jj + 1
            j = 0
            t = ""
        Next
        Return d
    End Function
    Public Function subbyte(ByVal a As String) As String
        Dim u(,) As String = {{"63", "7c", "77", "7B", "F2", "6B", "6F", "C5", "30", "01", "67", "2B", "FE", "D7", "AB", "76"}, {"CA", "82", "C9", "7D", "FA", "59", "47", "F0", "Ad", "D4", "A2", "AF", "9C", "A4", "72", "C0"}, {"B7", "FD", "93", "26", "36", "3F", "F7", "CC", "34", "A5", "E5", "F1", "71", "D8", "31", "15"}, {"04", "C7", "23", "C3", "18", "96", "05", "9A", "07", "12", "80", "E2", "EB", "27", "B2", "75"}, {"09", "83", "2C", "1A", "1B", "6E", "5A", "A0", "52", "3B", "D6", "B3", "29", "E3", "2F", "84"}, {"53", "D1", "00", "ED", "20", "FC", "B1", "5B", "6A", "CB", "BE", "39", "4A", "4C", "58", "CF"}, {"D0", "EF", "AA", "FB", "43", "4D", "33", "85", "45", "F9", "02", "7F", "50", "3C", "9F", "A8"}, {"51", "A3", "40", "8F", "92", "9D", "38", "F5", "BC", "B6", "DA", "21", "10", "FF", "F3", "D2"}, {"CD", "0C", "13", "EC", "5F", "97", "44", "17", "C4", "A7", "7E", "3D", "64", "5D", "19", "73"}, {"60", "81", "4F", "DC", "22", "2A", "90", "88", "46", "EE", "B8", "14", "DE", "5E", "0B", "DB"}, {"E0", "32", "3A", "0A", "49", "06", "24", "5C", "C2", "D3", "AC", "62", "91", "95", "E4", "79"}, {"E7", "C8", "37", "6D", "8D", "D5", "4E", "A9", "6C", "56", "F4", "EA", "65", "7A", "AE", "08"}, {"BA", "78", "25", "2E", "1C", "A6", "B4", "C6", "E8", "DD", "74", "1F", "4B", "BD", "8B", "8A"}, {"70", "3E", "B5", "66", "48", "03", "F6", "0E", "61", "35", "57", "B9", "86", "C1", "1D", "9e"}, {"E1", "F8", "98", "11", "69", "D9", "8E", "94", "9B", "1E", "87", "E9", "CE", "55", "28", "DF"}, {"8C", "A1", "89", "0D", "BF", "E6", "42", "68", "41", "99", "2D", "0F", "B0", "54", "BB", "16"}}
        Dim s, ss, aa, a1, f, r As String : s = "" : f = ""
        Dim k, l, tt, t1, v, z As Integer : k = 0 : l = 4
        For v = 1 To Len(a) Step +8
            aa = Mid(a, v, 8)
            s = Mid(aa, 1, 4)
            ss = Mid(aa, 5, 4)
            tt = int(s)
            t1 = int(ss)
            r = u(tt, t1)
            z = Len(r)
            If (z = 1) Then
                r = "0" + r
            End If
            a1 = binaryhex(r)
            f += a1
        Next
        Return f
    End Function
    Public Function inverssubbyte(ByVal a As String) As String
        Dim uu(,) As String = {{"52", "09", "6A", "D5", "30", "36", "A5", "38", "BF", "40", "A3", "9E", "81", "F3", "D7", "FB"}, {"7C", "E3", "39", "82", "9B", "2F", "FF", "87", "34", "8E", "43", "44", "C4", "DE", "E9", "CB"}, {"54", "7B", "94", "32", "A6", "C2", "23", "3D", "EE", "4C", "95", "0B", "42", "FA", "C3", "4E"}, {"08", "2E", "A1", "66", "28", "D9", "24", "B2", "76", "5B", "A2", "49", "6D", "8B", "D1", "25"}, {"72", "F8", "F6", "64", "86", "68", "98", "16", "D4", "A4", "5C", "CC", "5D", "65", "B6", "92"}, {"6C", "70", "48", "50", "FD", "ED", "B9", "DA", "5E", "15", "46", "57", "A7", "8D", "9D", "84"}, {"90", "D8", "AB", "00", "8C", "BC", "D3", "0A", "F7", "E4", "58", "05", "B8", "B3", "45", "06"}, {"D0", "2C", "1E", "8F", "CA", "3F", "0F", "02", "C1", "AF", "BD", "03", "01", "13", "8A", "6B"}, {"3A", "91", "11", "41", "4F", "67", "DC", "EA", "97", "F2", "CF", "CE", "F0", "B4", "E6", "73"}, {"96", "AC", "74", "22", "E7", "AD", "35", "85", "E2", "F9", "37", "E8", "1C", "75", "DF", "6E"}, {"47", "F1", "1A", "71", "1D", "29", "C5", "89", "6F", "B7", "62", "0E", "AA", "18", "BE", "1B"}, {"FC", "56", "3E", "4B", "C6", "D2", "79", "20", "9A", "DB", "C0", "FE", "78", "CD", "5A", "F4"}, {"1F", "DD", "A8", "33", "88", "07", "C7", "31", "B1", "12", "10", "59", "27", "80", "EC", "5F"}, {"60", "51", "7F", "A9", "19", "B5", "4A", "0D", "2D", "E5", "7A", "9F", "93", "C9", "9C", "EF"}, {"A0", "E0", "3B", "4D", "AE", "2A", "F5", "B0", "C8", "EB", "BB", "3c", "83", "53", "99", "61"}, {"17", "2B", "04", "7E", "BA", "77", "D6", "26", "E1", "69", "14", "63", "55", "21", "0C", "7D"}}
        Dim s1, ss1, f, a1, aa1, r1 As String : s1 = "" : f = ""
        Dim k1, l1, tt1, t11, j, z As Integer : k1 = 0 : l1 = 4
        For j = 1 To Len(a) Step +8
            aa1 = Mid(a, j, 8)
            s1 = Mid(aa1, 1, 4)
            ss1 = Mid(aa1, 5, 4)
            tt1 = int(s1)
            t11 = int(ss1)
            r1 = uu(tt1, t11)
            z = Len(r1)
            If (z = 1) Then
                r1 = "0" + r1
            End If
            a1 = binaryhex(r1)
            f += a1
        Next
        Return f
    End Function
    Public Function shiftrowtransformation(ByVal s As String) As String
        Dim a, b, c, d As String
        a = Mid(s, 1, 32)
        b = Mid(s, 33, 32)
        c = Mid(s, 65, 32)
        d = Mid(s, 97, 32)
        Dim b1, c1, d1, m As String
        b1 = Aes_shift(b, 8, "r")
        c1 = Aes_shift(c, 16, "r")
        d1 = Aes_shift(d, 24, "r")
        m = a + b1 + c1 + d1
        Return m
    End Function
    Public Function inversshiftrowtransformation(ByVal s As String) As String
        Dim a, b, c, d As String
        a = Mid(s, 1, 32)
        b = Mid(s, 33, 32)
        c = Mid(s, 65, 32)
        d = Mid(s, 97, 32)
        Dim b2, c2, d2, m As String
        b2 = Aes_shift(b, 8, "l")
        c2 = Aes_shift(c, 16, "l")
        d2 = Aes_shift(d, 24, "l")
        m = a + b2 + c2 + d2
        Return m
    End Function
    Public Function add(ByVal a As String, ByVal key As String)
        Dim z, f, k As String : z = "" : f = ""
        Dim m As Integer
        z = Xor128(a, key)
        For i = 1 To Len(z) Step +4
            m = int(Mid(z, i, 4))
            If m = 10 Then
                k = "A"
            ElseIf m = 11 Then
                k = "B"
            ElseIf m = 12 Then
                k = "C"
            ElseIf m = 13 Then
                k = "D"
            ElseIf m = 14 Then
                k = "E"
            ElseIf m = 15 Then
                k = "F"
            Else
                k = m
            End If
            f += k
        Next
        Dim f1, k1 As String : f1 = ""
        Dim m1 As String
        For i = 1 To Len(f)
            m1 = (Mid(f, i, 1)).ToString
            If m1 = "A" Then
                k1 = "1010"
            ElseIf m1 = "B" Then
                k1 = "1011"
            ElseIf m1 = "C" Then
                k1 = "1100"
            ElseIf m1 = "D" Then
                k1 = "1101"
            ElseIf m1 = "E" Then
                k1 = "1110"
            ElseIf m1 = "F" Then
                k1 = "1111"
            Else
                k1 = tobin4(m1)
            End If
            f1 += k1
        Next
        Return f1
    End Function
    Public Function addinvers(ByVal b As String, ByVal key As String)
        Dim z As String
        z = Xor128(b, key)
        Return z
    End Function
    Public Function mix(ByVal a As String)
        Dim i As Integer
        Dim r As String : r = ""
        For i = 1 To 128 Step +32
            r += mixcolumns(Mid(a, i, 32))
        Next
        Return r
    End Function
    Public Function mixinvers(ByVal a As String)
        Dim i As Integer
        Dim r As String : r = ""
        For i = 1 To 128 Step +32
            r += mixcolumnsinverse(Mid(a, i, 32))
        Next
        Return r
    End Function
    Public Function binary(ByVal a As String)
        Dim m As String : m = ""
        Dim z, z1, n As Integer
        For i = Len(a) To 1 Step -1
            z = Asc(Mid(a, i, 1))
            For n = 1 To 8
                z1 = z Mod 2
                m = z1.ToString + m
                z = z \ 2
            Next
        Next
        Return (m)
    End Function
    Public Function aesencryption(ByVal a As String, ByVal key As String)
        Dim d = binary(a)
        a = d
        Dim n, v, s, m As String : m = ""
        Dim round As Integer
        For round = 1 To 14
            n = subbyte(a)
            v = shiftrowtransformation(n)
            s = mix(v)
            m = add(s, key)
            a = m
        Next
        Return m
    End Function
    Public Function aesdecryption(ByVal a As String, ByVal key As String)
        Dim m, s, n, v As String : v = ""
        Dim round As Integer
        For round = 1 To 14
            m = addinvers(a, key)
            s = mixinvers(m)
            n = inversshiftrowtransformation(s)
            v = inverssubbyte(n)
            a = v
        Next
        Dim mm, z As String : z = ""
        Dim n1, uu As Integer
        For n1 = 1 To Len(v) Step +8
            mm = Mid(v, n1, 8)
            uu = int(mm)
            z += Chr(uu)
        Next
        Return z
    End Function
    Public Function completeencryption(ByVal a As String, ByVal key As String)
        Dim vv, newvv, complete, encryption As String : encryption = "" : complete = ""
        Dim length, kk As Integer
        For ff = 1 To Len(a) Step +16
            vv = Mid(a, ff, 16)
            If Len(vv) = 16 Then
                encryption += aesencryption(vv, key)
            ElseIf Len(vv) < 16 Then
                length = Len(vv)
                kk = 16 - length
                For i = 1 To kk
                    complete += " "
                Next
                newvv = vv + complete
                encryption += aesencryption(newvv, key)
            End If
        Next
        Return encryption
    End Function
    Public Function completedecryption(ByVal a As String, ByVal key As String)
        Dim decryption, vv, newvv, complete As String : decryption = "" : complete = ""
        Dim length, kk As Integer
        For ff = 1 To Len(a) Step +128
            vv = Mid(a, ff, 128)
            If Len(vv) = 128 Then
                decryption += aesdecryption(vv, key)
            ElseIf Len(vv) < 128 Then
                length = Len(vv)
                kk = 128 - length
                For i = 1 To kk
                    complete += "0"
                Next
                newvv = vv + complete
                decryption += aesdecryption(newvv, key)
            End If
        Next
        Return decryption
    End Function
    Public Function bin_word(ByVal t As String)
        Dim mm, z, zz As String : zz = ""
        Dim uu, n1 As Integer
        For n1 = 1 To Len(t) Step +8
            mm = Mid(t, n1, 8)
            uu = int(mm)
            If uu = 0 Then
                z = ""
                GoTo 1
            End If
            z = Chr((uu))
1:          zz &= z
        Next
        Return zz
    End Function



    Function Encrypt(r As String)
        Dim f As String
        'OpenFileDialog1.InitialDirectory = "e:\"
        'OpenFileDialog1.FileName = "open a file..."
        'OpenFileDialog1.Filter = "only text files (*.txt)| *.txt"
        'OpenFileDialog1.ShowDialog()
        'Dim r As New IO.StreamReader(OpenFileDialog1.FileName)
        'TextBox1.Text = r.ReadToEnd
        'r.Close()
        '  MsgBox(Len(TextBox1.Text))
        Dim key As String = "11111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111"
        'MsgBox(Len(key))
        'Dim watch As Stopwatch = Stopwatch.StartNew()
        'watch.Start()
        f = completeencryption(r, key)
        'watch.Stop()
        'TextBox4.Text = (watch.Elapsed.TotalSeconds)
        'Dim objWriter As New System.IO.StreamWriter("e:\Decrypt_Aes.txt", False)
        'objWriter.WriteLine(f)
        'objWriter.Close()
        Return bin_word(f)
    End Function

    Function Decrypt(mm As String)
        'Dim kk, ee, mm As String
        'OpenFileDialog1.InitialDirectory = "e:\"
        'OpenFileDialog1.FileName = "open a file..."
        'OpenFileDialog1.Filter = "only text files (*.txt)| *.txt"
        'OpenFileDialog1.ShowDialog()
        'Dim rr As New IO.StreamReader(OpenFileDialog1.FileName)
        'mm = rr.ReadToEnd
        'rr.Close()
        'TextBox2.Text = bin_word(mm)
        Dim kk = "11111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111"
        'Dim watch As Stopwatch = Stopwatch.StartNew()
        'watch.Start()
        Dim ee = completedecryption(mm, kk)
        Dim aa = Mid(ee, 1, Len(ee))
        'watch.Stop()
        'TextBox5.Text = (watch.Elapsed.TotalSeconds)
        Return aa.TrimEnd()
    End Function
End Module
