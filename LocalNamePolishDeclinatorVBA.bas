Attribute VB_Name = "Declinator"


Private Function CORRECTENDING(nominative As String, ending As String) As String
    
    Dim result As String: result = ending
   
    If InStr(1, nominative, "œciniec") > 0 And InStr(1, ending, "ñca") > 0 Then 'e.g. Goœciniec, -ñca -> -œciñca
        result = Replace(ending, "ñca", "œciñca")
    ElseIf InStr(1, nominative + "#", "sieniec") > 0 And InStr(1, ending, "ñca") > 0 Then 'e.g. Lesieniec, -ñca -> -sieñca
        result = Replace(ending, "ñca", "sieñca")
    ElseIf InStr(1, nominative + "#", "oniec") > 0 And InStr(1, ending, "ñca") > 0 Then 'e.g. Koniec, -ñca -> -oñca
        result = Replace(ending, "ñca", "oñca")
    ElseIf InStr(1, nominative + "#", "aniec") > 0 And InStr(1, ending, "ñca") > 0 Then 'e.g. Koniec, -ñca -> -añca
        result = Replace(ending, "ñca", "añca")
    ElseIf InStr(1, nominative + "#", "yniec") > 0 And InStr(1, ending, "ñca") > 0 Then 'e.g. Tyniec, -ñca -> -yñca
        result = Replace(ending, "ñca", "yñca")
    ElseIf InStr(1, nominative + "#", "mieniec") > 0 Then 'e.g. Kamieniec, -ñca -> -mieñca
        result = Replace(ending, "ñca", "mieñca")
    ElseIf InStr(1, nominative + "#", "liniec") > 0 Then 'e.g. Czapliniec, -ñca -> -liñca
        result = Replace(ending, "ñca", "liñca")
    ElseIf InStr(1, nominative + "#", "yñ#") > 0 And InStr(1, ending, "nia") > 0 Then 'e.g. Ci¹¿yn, -nia -> -ynia
        result = Replace(ending, "nia", "ynia")
    ElseIf InStr(1, nominative, "Rynek") > 0 And InStr(1, ending, "ku") > 0 Then 'e.g. Rynek, -ku -> -nku
        result = Replace(ending, "ku", "nku")
    ElseIf InStr(1, nominative, "Grobla") > 0 And InStr(1, ending, "ni") > 0 Then 'e.g. Grobla, -ni -> -li
        result = "-li"
    End If
    
    CORRECTENDING = result

End Function







Function GETGENITIVE(nominative As String, genitive_ending As String) As String
    
    Dim result As String: result = ""
    
    genitive_ending = Replace(genitive_ending, "-", "")

    For i = Len(nominative) To 1 Step -1
           
        If StrComp(Mid(nominative, i, 1), Left(genitive_ending, 1)) = 0 Then
            result = Mid(nominative, 1, i - 1) + genitive_ending
            Exit For
        End If
    Next i
       
    GETGENITIVE = result

End Function





Function GetParadigm(nominative As String, genitive As String) As String
    
    Dim result As String: result = ""

    Dim nominative_last_letter As String: nominative_last_letter = Right(nominative, 1)
    Dim genitive_last_letter As String: genitive_last_letter = Right(genitive, 1)
        
    Select Case nominative_last_letter
        Case "a" 'e.g. Polska, Moskwa
            Select Case genitive_last_letter
                Case "y": result = "sf" 'e.g. Moskwa, Moskwy
                Case "i": result = "sf" 'e.g. Polska, Polski
                Case "j": result = "af" 'e.g. Nowa, Nowej
                Case Else: result = "sp"
            End Select
        Case "e" 'e.g. Nowe, Brzeszcze
            Select Case genitive_last_letter
                Case "a": result = "sn" 'e.g. Pole, Pola
                Case "o": result = "an" 'e.g. Nowe, Nowego
                Case "h": result = "ap" 'e.g. Nowe, Nowych
                Case Else: result = "sp"
            End Select
        Case "i" 'e.g. Krótki, Tani
            Select Case genitive_last_letter
                Case "o": result = "am" 'e.g. Krótki, Krótkiego
                Case "h": result = "ap" 'e.g. Krótki, Krótkich
                Case Else: result = "sp"
            End Select
        Case "o" 'e.g. Gniezno
            Select Case genitive_last_letter
                Case "a": result = "sn" 'e.g. Gniezno, Gniezna
            End Select
        Case "y" 'e.g. Nowy
            Select Case genitive_last_letter
                Case "o": result = "am" 'e.g. Nowy, Nowego
                Case "h": result = "ap" 'e.g.
                Case Else: result = "sp"
            End Select
        Case Else
            Select Case genitive_last_letter
                Case "a": result = "sm" 'e.g. Kraków, Krakowa
                Case "u": result = "sm"
                Case "i": result = "sf"
                Case "o": result = "am" 'e.g. Jacków, Jackowego
            End Select
    End Select
    
    GetParadigm = result

End Function





Function GetPos(nominative As String, genitive As String) As String
    Dim paradigm As String: paradigm = GetParadigm(nominative, genitive)
    GetPos = Mid(paradigm, 1, 1)
End Function





Function GetGender(nominative As String, genitive As String) As String
    Dim paradigm As String: paradigm = GetParadigm(nominative, genitive)
    GetGender = Mid(paradigm, 2, 1)
End Function




Function AddEnding(form As String, oldEnding As String, newEnding As String) As String

    Dim endMarker As String: endMarker = "#"
    
    form = form + endMarker
    oldEnding = oldEnding + endMarker
    
    If InStr(1, form, oldEnding, vbTextCompare) > 0 Then
        AddEnding = Replace(form, oldEnding, newEnding, 1, -1, vbTextCompare)
    Else
        AddEnding = Replace(form, "#", "")
    End If

End Function




Function GETDATIVE(nominative As String, genitive As String) As String
    
    Dim result As String: result = ""
    
    Dim pos As String: pos = GetPos(nominative, genitive)
    Dim gender As String: gender = GetGender(nominative, genitive)
    
    genitive = GETGENITIVE(nominative, genitive)
    
    Select Case pos
        Case "a"
            Select Case gender
                Case "f"
                    result = AddEnding(nominative, "ga", "giej")
                    result = AddEnding(result, "ka", "kiej")
                    result = AddEnding(result, "ia", "iej")
                    result = AddEnding(result, "a", "ej")
                Case "m"
                    result = AddEnding(nominative, "i", "iemu")
                    result = AddEnding(result, "y", "emu")
                    result = AddEnding(result, "ów", "owemu")
                Case "n"
                    result = AddEnding(nominative, "ie", "iemu")
                    result = AddEnding(result, "e", "emu")
                Case "p"
                    result = AddEnding(nominative, "ie", "im")
                    result = AddEnding(result, "e", "ym")
                    result = AddEnding(result, "y", "em")
            End Select
        Case "s"
            Select Case gender
                Case "f"
                    result = AddEnding(genitive, "cy", "cy!")
                    result = AddEnding(result, "dzy", "dzy!")
                    result = AddEnding(result, "szy", "szy!")
                    result = AddEnding(result, "czy", "czy!")
                    result = AddEnding(result, "¿y", "¿y!")
                    result = AddEnding(result, "rzy", "rzy!")
                    
                    result = AddEnding(result, "ki", "ce")
                    result = AddEnding(result, "gi", "dze")
                    result = AddEnding(result, "chy", "sze")
                    result = AddEnding(result, "hy", "¿e")
                    result = AddEnding(result, "dy", "dzie")
                    result = AddEnding(result, "ty", "cie")
                    result = AddEnding(result, "ry", "rze")
                    
                    result = AddEnding(result, "zny", "Ÿnie")
                    result = AddEnding(result, "sny", "œnie")
                    
                    result = AddEnding(result, "y", "ie")

                    result = AddEnding(result, "!", "")
                Case "m"
                    result = AddEnding(genitive, "a", "owi")
                    result = AddEnding(result, "u", "owi")
                Case "n"
                    result = AddEnding(genitive, "a", "u")
                Case "p"
                    result = AddEnding(nominative, "a", "om")
                    result = AddEnding(result, "e", "om")
                    result = AddEnding(result, "i", "om")
                    result = AddEnding(result, "y", "om")
            End Select
        Case Else
            result = nominative
    End Select
    
    GETDATIVE = result
    
End Function



Function GETACCUSATIVE(nominative As String, genitive As String) As String
    
    Dim result As String: result = ""
    
    Dim pos As String: pos = GetPos(nominative, genitive)
    Dim gender As String: gender = GetGender(nominative, genitive)
       
    Select Case pos
        Case "a"
            Select Case gender
                Case "f"
                    result = AddEnding(nominative, "a", "¹")
                Case Else
                    result = nominative
            End Select
        Case "s"
            Select Case gender
                Case "f"
                    result = AddEnding(nominative, "a", "ê")
                Case Else
                    result = nominative
            End Select
        Case Else
            result = nominative
    End Select
    
    GETACCUSATIVE = result
End Function





Function GETINSTRUMENTAL(nominative As String, genitive As String) As String
    
    Dim result As String: result = ""
    
    Dim pos As String: pos = GetPos(nominative, genitive)
    Dim gender As String: gender = GetGender(nominative, genitive)
    
    genitive = GETGENITIVE(nominative, genitive)
    
    Select Case pos
        Case "a"
            Select Case gender
                Case "f"
                    result = AddEnding(nominative, "a", "¹")
                Case "m"
                    result = AddEnding(nominative, "i", "im")
                    result = AddEnding(result, "y", "ym")
                    result = AddEnding(result, "ów", "owym")
                Case "n"
                    result = AddEnding(nominative, "ie", "im")
                    result = AddEnding(result, "e", "ym")
                Case "p"
                    result = AddEnding(nominative, "ie", "imi")
                    result = AddEnding(result, "e", "ymi")
                    result = AddEnding(result, "y", "ema")
            End Select
        Case "s"
            Select Case gender
                Case "f"
                    result = AddEnding(genitive, "ii", "i¹")
                    result = AddEnding(result, "ki", "k¹")
                    result = AddEnding(result, "gi", "g¹")
                    result = AddEnding(result, "hy", "h¹")
                    result = AddEnding(result, "ji", "j¹")
                    result = AddEnding(result, "i", "i¹")
                    result = AddEnding(result, "y", "¹")
                Case "m"
                    result = AddEnding(genitive, "ka", "kiem")
                    result = AddEnding(result, "ku", "kiem")
                    result = AddEnding(result, "ga", "giem")
                    result = AddEnding(result, "gu", "giem")
                    result = AddEnding(result, "u", "em")
                    result = AddEnding(result, "a", "em")
                Case "n"
                    result = AddEnding(genitive, "ga", "giem")
                    result = AddEnding(result, "ka", "kiem")
                    result = AddEnding(result, "a", "em")
                Case "p"
                    result = AddEnding(nominative, "i", "ami")
                    result = AddEnding(result, "e", "ami")
                    result = AddEnding(result, "a", "ami")
                    result = AddEnding(result, "y", "ami")
            End Select
        Case Else
            result = nominative
    End Select
    
    GETINSTRUMENTAL = result
End Function



Function GETLOCATIVE(nominative As String, genitive As String) As String
    
    Dim result As String: result = ""
    
    Dim pos As String: pos = GetPos(nominative, genitive)
    Dim gender As String: gender = GetGender(nominative, genitive)
    
    genitive = GETGENITIVE(nominative, genitive)
    
    Select Case pos
        Case "a"
            Select Case gender
                Case "f"
                    result = AddEnding(nominative, "ga", "giej")
                    result = AddEnding(result, "ka", "kiej")
                    result = AddEnding(result, "ia", "iej")
                    result = AddEnding(result, "a", "ej")
                Case "m"
                    result = GETINSTRUMENTAL(nominative, genitive)
                Case "n"
                    result = GETINSTRUMENTAL(nominative, genitive)
                Case "p"
                    result = AddEnding(nominative, "ie", "ich")
                    result = AddEnding(result, "e", "ych")
                    result = AddEnding(result, "y", "ech")
            End Select
        Case "s"
            Select Case gender
                Case "f"
                    result = GETDATIVE(nominative, genitive)
                Case "m"
                    result = AddEnding(genitive, "ia", "iu")
                    result = AddEnding(result, "la", "lu")
                    result = AddEnding(result, "ca", "cu")
                    result = AddEnding(result, "za", "zu")
                    result = AddEnding(result, "¿a", "¿u")
                    result = AddEnding(result, "ka", "ku")
                    result = AddEnding(result, "ga", "gu")
                    result = AddEnding(result, "ha", "hu")
                    result = AddEnding(result, "ra", "rze")
                    result = AddEnding(result, "ru", "rze")
                    result = AddEnding(result, "ta", "cie")
                    result = AddEnding(result, "tu", "cie")
                    result = AddEnding(result, "da", "dzie")
                    result = AddEnding(result, "du", "dzie")
                    result = AddEnding(result, "a", "ie")
                Case "n"
                    result = AddEnding(genitive, "za", "zu")
                    result = AddEnding(result, "ia", "iu")
                    result = AddEnding(result, "¿a", "¿u")
                    result = AddEnding(result, "la", "lu")
                    result = AddEnding(result, "ka", "ku")
                    result = AddEnding(result, "ga", "gu")
                    result = AddEnding(result, "ha", "hu")
                    result = AddEnding(result, "ra", "rze")
                    result = AddEnding(result, "ru", "rze")
                    result = AddEnding(result, "ta", "cie")
                    result = AddEnding(result, "tu", "cie")
                    result = AddEnding(result, "da", "dzie")
                    result = AddEnding(result, "du", "dzie")
                    result = AddEnding(result, "³a", "le")
                    result = AddEnding(result, "a", "ie")
                Case "p"
                    result = AddEnding(nominative, "a", "ach")
                    result = AddEnding(result, "e", "ach")
                    result = AddEnding(result, "i", "ach")
                    result = AddEnding(result, "y", "ach")
            End Select
        Case Else
            result = nominative
    End Select
    
    GETLOCATIVE = result
End Function



Function GETVOCATIVE(nominative As String, genitive As String) As String
    
    Dim result As String: result = ""
    
    Dim pos As String: pos = GetPos(nominative, genitive)
    Dim gender As String: gender = GetGender(nominative, genitive)
    
    genitive = GETGENITIVE(nominative, genitive)
    
    Select Case pos
        Case "a"
            result = nominative
        Case "s"
            Select Case gender
                Case "f"
                    Dim nominative_end As String: nominative_end = Right(nominative, 1)
                    
                    If nominative_end = "a" Then 'e.g. kobieta, niania
                        result = AddEnding(nominative, "a", "o")
                    Else 'e.g. moc, koœæ, gospodyni, pani
                        result = gen
                    End If
                Case "m"
                    result = AddEnding(genitive, "u", "u")
                    result = AddEnding(result, "a", "u")
                Case Else
                    result = nominative
            End Select
        Case Else
            result = nominative
    End Select
    
    GETVOCATIVE = result
End Function



Function GETFORM(nominative As String, genitive As String, caseNo As Integer) As String
        
    Dim result As String: result = ""
    
    genitive = GETGENITIVE(nominative, genitive)

    Select Case caseNo
        Case 1: result = genitive
        Case 2: result = GETDATIVE(nominative, genitive)
        Case 3: result = GETACCUSATIVE(nominative, genitive)
        Case 4: result = GETINSTRUMENTAL(nominative, genitive)
        Case 5: result = GETLOCATIVE(nominative, genitive)
        Case 6: result = GETVOCATIVE(nominative, genitive)
        Case Else: result = nominative
    End Select
        
    GETFORM = result

End Function




Private Function GETELEMENTS(text As String) As Collection
    
    Dim result As Collection: Set result = New Collection
    
    Dim buffer As String: buffer = ""
    
    For i = 1 To Len(text)
    
        Dim letter As String: letter = Mid(text, i, 1)
        
        If StrComp(letter, "-", vbTextCompare) = 0 Or StrComp(letter, " ", vbTextCompare) = 0 Then
            result.Add (buffer)
            result.Add (letter)
            buffer = ""
        Else
            buffer = buffer + letter
        End If
    Next i
    
    result.Add (buffer)

    Set GETELEMENTS = result

End Function




Private Function CheckIfDeclineable(nominative As String, genitive As String) As Boolean
    
    CheckIfDeclineable = StrComp(genitive, nominative) <> 0

End Function





Private Function CorrectGenitive(genitive As String) As String

    Dim result As String: result = Replace(genitive, " -", "-")
    
    If StrComp(Mid(genitive, 1, 1), "-") = 0 Then
        result = Mid(result, 2)
    End If
    
    CorrectGenitive = result

End Function





Function DECLINE(nominative As String, genitive As String, caseNo As Integer) As String
    
    Dim result As String: result = ""
    
    genitive = CorrectGenitive(genitive)

    Dim nominative_elements As Collection: Set nominative_elements = GETELEMENTS(nominative)
    Dim genitive_elements As Collection: Set genitive_elements = GETELEMENTS(genitive)
    
    Dim declineable As Boolean: declineable = True
    
    For i = 1 To nominative_elements.Count
    
        nominative = CStr(nominative_elements(i))
    
        If nominative <> "-" And nominative <> " " Then
        
            genitive = CStr(genitive_elements(i))
            genitive = GETGENITIVE(nominative, CORRECTENDING(nominative, genitive))
                        
            declineable = CheckIfDeclineable(nominative, genitive)
                                                
            If declineable Then
                result = result + GETFORM(nominative, genitive, caseNo)
            Else
                result = result + nominative
            End If
        Else
            result = result + nominative
            
        End If
    Next i
        
    DECLINE = result

End Function






