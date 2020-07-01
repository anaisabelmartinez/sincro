Attribute VB_Name = "Traductor"
Sub Tradueix(IdiomaOrigen, IdiomaDesti)    'http://ajax.googleapis.com/ajax/services/language/translate?v=1.0&q=Hola&langpair=es|en
    Dim URL As String, Paraula As String, Rs As rdoResultset, Original, Traduit, P, Pp, Q As rdoQuery, Orig, i
        
'CA    Catal
'ES    Castell
'DE    Alemany
'EN"   Angl
'FR    Franc
'IT    Itali
'PT    Portugu
                
        
    FesElConnect
    ExecutaComandaSql "CREATE INDEX diccionari_IdStr ON diccionari (IdStr)"
    ExecutaComandaSql "CREATE INDEX diccionari_Idioma ON diccionari (Idioma)"
    ExecutaComandaSql "CREATE INDEX diccionari_App ON diccionari (App)"
    ExecutaComandaSql "CREATE INDEX diccionari_TexteOriginal ON diccionari (TexteOriginal)"
    
    Set Rs = Db.OpenResultset("select d1.* from diccionari d1  left join diccionari d2 on d1.idstr = d2.idstr and d2.idioma = '" & IdiomaDesti & "'  where d2.id is null and d1.idioma = '" & IdiomaOrigen & "' ")
    Set Q = Db.CreateQuery("", "Insert Into diccionari (id,IdStr,App,Pagina,Idioma,TexteOriginal,Texte ) Values (newid(),?,?,?,?,?,?)")
    
    While Not Rs.EOF
        Original = Rs("Texte")
        
'        Orig = Split(Original, " ")
        
        Traduit = ""
        URL = "http://ajax.googleapis.com/ajax/services/language/translate?v=1.0&q=" & URLEncode(Original) & "&langpair=" & IdiomaOrigen & "|" & IdiomaDesti
        resposta = llegeigHtml(URL)
        If Len(resposta) > 0 And Not IsNull(Rs("IdStr")) Then
            P = InStr(resposta, "translatedText")
            Pp = InStr(P + 16, resposta, "}")
            Traduit = DecodeUTF8(Mid(resposta, P + 17, Pp - P - 18))
    '        Db.Execute "Insert Into diccionari (id,IdStr,App,Pagina,Idioma,TexteOriginal,Texte ) Values (newid(),'" & Rs("IdStr") & "','" & Rs("App") & "','" & Rs("Pagina") & "','" & IdiomaDesti & "','" & Rs("TexteOriginal") & "','" & Traduit & "')"
            
            Q.rdoParameters(0) = Rs("IdStr")
            Q.rdoParameters(1) = Rs("App")
            Q.rdoParameters(2) = Rs("Pagina")
            Q.rdoParameters(3) = IdiomaDesti
            Q.rdoParameters(4) = Rs("TexteOriginal")
            Q.rdoParameters(5) = Traduit
            Q.Execute
        
            Debug.Print Original & "-->" & Traduit
        End If
        Rs.MoveNext
    Wend


    

End Sub


Sub TradueixTot()
   Tradueix "CA", "ES"
   Tradueix "ES", "EN"
   Tradueix "ES", "DE"
   Tradueix "ES", "FR"
   Tradueix "ES", "IT"
   Tradueix "ES", "PT"
End Sub


Function URLEncode(ByVal Text As String) As String
    Dim i As Integer
    Dim acode As Integer
    Dim char As String
    
    URLEncode = Text
    
    For i = Len(URLEncode) To 1 Step -1
        acode = Asc(Mid$(URLEncode, i, 1))
        Select Case acode
            Case 48 To 57, 65 To 90, 97 To 122
                ' don't touch alphanumeric chars
            Case 32
                ' replace space with "+"
                Mid$(URLEncode, i, 1) = "+"
            Case Else
                ' replace punctuation chars with "%hex"
                URLEncode = Left$(URLEncode, i - 1) & "%" & Hex$(acode) & Mid$ _
                    (URLEncode, i + 1)
        End Select
    Next
    
End Function
Function DecodeUTF8(s)
  Dim i
  Dim c
  Dim n

  i = 1
  Do While i <= Len(s)
    c = Asc(Mid(s, i, 1))
    If c And &H80 Then
      n = 1
      Do While i + n < Len(s)
        If (Asc(Mid(s, i + n, 1)) And &HC0) <> &H80 Then
          Exit Do
        End If
        n = n + 1
      Loop
      If n = 2 And ((c And &HE0) = &HC0) Then
        c = Asc(Mid(s, i + 1, 1)) + &H40 * (c And &H1)
      Else
        c = 191
      End If
      s = Left(s, i - 1) + Chr(c) + Mid(s, i + n)
    End If
    i = i + 1
  Loop
  DecodeUTF8 = s
End Function

