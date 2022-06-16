Attribute VB_Name = "Module1"
Global tel As Integer


Sub besteMelder()
    Dim oMail As MailItem
    Dim oInspector As Inspector
    Dim melder As String
    Dim myHTMLText As String
    Dim myPlainText As String
    
    Set oMail = GetCurrentItem()
    
    melder = grijpMelder(oMail)
        
    myPlainText = "Beste " & melder & ","
    myHTMLText = "Beste " & blauw(melder) & ","
    myHTMLText = "<span style=" & Chr(34) & "font-family:verdana;font-size:9pt" & Chr(34) & ">" & myHTMLText & "</span>"
    
    '' Als eerste regel al gezet is dan niks doen anders Beste melder neerzetten
    If Mid(oMail.Body, 1, Len(myPlainText)) <> myPlainText Then plakTextInBody myHTMLText, oMail, oInspector

End Sub

Sub fixToCC()
    
    Call corrigeerToCC(GetCurrentItem())

End Sub

Sub aanhefEnFixCC()

    fixToCC
    besteMelder
    SendKeys "{DOWN}", True
    SendKeys "{Enter}", True
    
End Sub



Sub zeven24()

    Dim oMail As MailItem
    Set oMail = GetCurrentItem()
    Dim gevondenRegel As String
    Dim openBrack As Integer
    Dim closeBrack As Integer
    Dim opdracht() As String
    Dim element(0 To 9) As Vraagstuk
    Dim oldSubject As String
    Dim newSubject As String
    Dim opdrachtLen As Integer
    Dim offset As Integer
    Dim cnt As Integer
    Dim startCode As String
    Dim eindCode As String
    Dim HTMLBlok As String
    Dim eersteRegelNaCode As String
    Dim i As Integer
    Dim j As Integer
    Dim a As Integer
    Dim b As Integer
    
    startCode = "[["
    eindCode = "]]"
    cnt = 0
    offset = 1
    oldSubject = oMail.Subject
    
    While InStr(offset, oMail.Body, startCode) > 0
    
        Set element(cnt) = New Vraagstuk
        
        With element(cnt)
            .startPos = InStr(offset, oMail.Body, startCode)
            If .startPos = 0 Then
                MsgBox "Fout in opdrachtcode: Geen begincode gevonden '" & startCode & "'" '' zal nooit voor moeten komen
                Exit Sub
            End If
            .eindPos = InStr(.startPos, oMail.Body, eindCode)
            If .eindPos = 0 Then
                MsgBox "Fout in opdrachtcode: Geen eindcode gevonden '" & eindCode & "'"
                Exit Sub
            End If
            .index = cnt '' nog niet gebruikt
            
            .gevondenRegel = Mid(oMail.Body, .startPos + Len(startCode), .eindPos - .startPos - Len(eindCode))
            
            Erase opdracht
            opdracht() = Split(.gevondenRegel, ",")
            opdrachtLen = UBound(opdracht) - LBound(opdracht)
                        
            If opdrachtLen >= 0 Then .zoekWoord = Trim(opdracht(0))
            If opdrachtLen >= 1 Then .vervangMet = Trim(opdracht(1))
            If opdrachtLen >= 2 Then .stelVraag = Trim(opdracht(2))
            If opdrachtLen >= 3 Then .vraagTitel = Trim(opdracht(3))
            If opdrachtLen >= 4 Then .defaultInput = Trim(opdracht(4))
            
            If .zoekWoord = "subject" Then
                .defaultInput = mooi(Replace(oMail.Subject, "RE: ", ""))
            End If
              
            If .stelVraag <> "" Then
                .antwoord = InputBox(.stelVraag, .vraagTitel, .defaultInput)
                If StrPtr(.antwoord) = 0 Then Exit Sub
                .vervangMet = Replace(.vervangMet, .zoekWoord, .antwoord)
            End If
            
            
            offset = .eindPos + 2
            cnt = cnt + 1
        End With
    Wend
    cnt = cnt - 1
    If cnt < 0 Then Exit Sub
    
    eersteRegelNaCode = Trim(Mid(oMail.Body, element(cnt).eindPos + Len(eindCode), 50))
    eersteRegelNaCode = Replace(eersteRegelNaCode, Chr(13), "")
    eersteRegelNaCode = Trim(Replace(eersteRegelNaCode, Chr(10), ""))

    
    HTMLBlok = oMail.HTMLBody
    oldSubject = oMail.Subject
    
    '' Haal eerst de afgewerkte code weg uit de body
    For i = 0 To cnt
        With element(i)
            HTMLBlok = Replace(HTMLBlok, startCode & .gevondenRegel & eindCode, "")
        End With
    Next i
    
   
    
    '' eerst body fixen maar alles van subject overslaan
    For i = 0 To cnt
        For j = 0 To cnt
            If i <> j And element(i).zoekWoord <> "subject" And element(j).zoekWoord <> "subject" Then
                element(i).vervangMet = Replace(element(i).vervangMet, element(j).zoekWoord, element(j).vervangMet)
            End If
        Next j
    Next i
    
    For i = 0 To cnt
        With element(i)
            If .zoekWoord = "subject" Then
                HTMLBlok = Replace(HTMLBlok, "subject", .defaultInput)
                oldSubject = .defaultInput
            Else
                HTMLBlok = Replace(HTMLBlok, .zoekWoord, .vervangMet)
            End If
        End With
    Next i
    
    
    For i = 0 To cnt
        With element(i)
            If .zoekWoord = "subject" Then
                newSubject = .vervangMet
                newSubject = Replace(newSubject, "subject", oldSubject)
            End If
        End With
    Next i
    
    For i = 0 To cnt
        With element(i)
            If .zoekWoord <> "subject" Then
                If .antwoord = "" Then
                    newSubject = Replace(newSubject, .zoekWoord, .vervangMet)
                Else
                    newSubject = Replace(newSubject, .zoekWoord, .antwoord)
                End If
            End If
        End With
    Next i

    oMail.HTMLBody = HTMLBlok
    oMail.Subject = newSubject
    besteMelder
    
    Dim aa As Long
    Dim bb As Long
    Dim nieuwHap As String
    Dim hap As String
    Dim zoek As String
    Dim z As String
    Dim aantalWeghalen As Integer
    

    HTMLBlok = oMail.HTMLBody
    zoek = "</span>,</span>"
    aa = InStr(HTMLBlok, zoek) + Len(zoek)

    If aa <= Len(zoek) Then
    
        MsgBox (zoek & " niet gevonden")
        
    Else

        z = "<o:p>&nbsp;</o:p>"
        bb = InStr(aa, HTMLBlok, eersteRegelNaCode)
        hap = Mid(HTMLBlok, aa, bb - aa)
        aantalWeghalen = gevondenAantal(hap, z) - 2 '' twee regels laten staan

        If aantalWeghalen > 0 Then
        
            nieuwHap = Replace(hap, z, "", , aantalWeghalen)
            HTMLBlok = Replace(HTMLBlok, hap, nieuwHap)
            
        End If
        
        oMail.HTMLBody = HTMLBlok
        
    End If
    
    
    fixToCC
    
End Sub
Function week(txt As String) As String

    On Error Resume Next
    
    Dim weekdag() As Variant
    weekdag = Array("", "zondag", "maandag", "dinsdag", "woensdag", "donderdag", "vrijdag", "zaterdag")
    
    week = weekdag(Weekday(txt)) & " " & txt
    
End Function

Sub InsertText()
    
    '' init
    Dim myHTMLText As String
    Dim deAanhef As String
    Dim vanTot As String
    Dim opDatum As String
    Dim heenterug As String
    Dim ziekGemeld As String
    Dim beterGemeld As String
    Dim nogVragen As String
    Dim newMail    As MailItem
    Dim oInspector As Inspector
    Dim mySubject As String
    Dim newSubject As String
    Dim myNamespace As Outlook.NameSpace
    Dim zelf() As String

 
    '' Pak achter- en voornaam van huidige gebruiker  achternaam = zelf(0)  en  voornaam = zelf(1)
    Set myNamespace = Application.GetNamespace("MAPI")
    zelf = Split(myNamespace.CurrentUser, ",")
    
    '' Check of mail inline of in eigen window getoond wordt
    Set oInspector = Application.ActiveInspector
    If oInspector Is Nothing Then
        Set newMail = GetCurrentItem()
    Else
        Set newMail = oInspector.CurrentItem
    End If
    
    
    '' text declaraties
    deAanhef = "Beste [naamMelder],"
    vanTot = "Het vervoer (heen en retour) van  [naamClient]  is van  [vanDatum]  tot en met  [totDatum]  geannuleerd."
    opDatum = "Het vervoer (heen en retour) van  [naamClient]  is op  [opDatum]  geannuleerd."
    heenterug = "De [heenTerug]  van  [naamClient]  is voor  [opDatum]  geannuleerd."
    nogVragen = "Mocht u nog vragen hebben, neem dan gerust contact met ons op."
    ziekGemeld = "Het vervoer (heen en retour) van  [naamClient]  is [per] afgemeld tot nader order."
    beterGemeld = "Het vervoer (heen en retour) van  [naamClient]  is [per] weer aangemeld."
    
    
    '' *******************************************
    '' ** Subject veld uitpluizen en mooi maken **
    '' *******************************************
    Dim aanOfAfmelding As String
    Dim larry() As String
    Dim larryLength As Integer
    Dim doel As String
    Dim clientvolnaam As String
    Dim dat As String
    Dim heenOfterug  As String
    Dim doelen() As Variant
    Dim dl As Variant
    dl = "leeg"
    doelen = Array("apeldoorn", "bedum", "druten", "ermelo", "julianadorp", "monster", "noorderbrug", "noordwijk", "regio zuid", "wekerom", "zeeland")
    
    '' Splits alles door komma gescheiden en stop in array genaamd larry
    larry() = Split(newMail.Subject, ",")
    larryLength = UBound(larry) - LBound(larry)
     
    '' Plak eerste sectie voor de komma in doel
    doel = Trim(LCase(larry(0)))
    '' Kijk of perceelnaam geheel of gedeeltelijk ingevuld is
    If (IsInArray(doel, doelen)) Then
        dl = Filter(doelen, doel)
        doel = dl(0)
    End If
    If doel = "nw" Then doel = "noordwijk"
    If doel = "nb" Then doel = "noorderbrug"
    If doel = "z" Then doel = "zeeland"
    If doel = "zuid" Then doel = "regio zuid"
    doel = mooi(doel)
    
    '' Plak tweede woord in clientnaam
    If larryLength >= 1 Then
        clientvolnaam = mooi(Trim(larry(1)))
        clientvolnaam = Replace(clientvolnaam, " En ", " en ")
        clientvolnaam = Replace(clientvolnaam, " Van ", " van ")
        clientvolnaam = Replace(clientvolnaam, " De ", " de ")
        clientvolnaam = Replace(clientvolnaam, " Der ", " der ")
        clientvolnaam = Replace(clientvolnaam, " Den ", " den ")
        clientvolnaam = Replace(clientvolnaam, " Op ", " op ")
        clientvolnaam = Replace(clientvolnaam, "Begeleider", "begeleider")
        clientvolnaam = Replace(clientvolnaam, "Begeleiding", "begeleiding")
        clientvolnaam = Replace(clientvolnaam, "Personen", " personen")
        clientvolnaam = Replace(clientvolnaam, "Genoemde", " genoemde")
        clientvolnaam = Replace(clientvolnaam, "Onderstaande", "onderstaande")
        clientvolnaam = Replace(clientvolnaam, "Plus ", "plus ")
        clientvolnaam = Replace(clientvolnaam, ".", ",")
    End If
    
    '' Plak derde zin in dat
    If larryLength >= 2 Then
        dat = Trim(larry(2))
        If dat = "z" Then dat = "ziek"
        If dat = "b" Then dat = "beter"
        dat = Replace(dat, ".", ",")
        dat = Replace(dat, "mar", "maart")
        dat = Replace(dat, "mrt", "maart")
        dat = Replace(dat, "apr", "april")
        dat = Replace(dat, "juni", "jun")
        dat = Replace(dat, "jun", "juni")
        dat = Replace(dat, "juli", "jul")
        dat = Replace(dat, "jul", "juli")
        dat = Replace(dat, "aug", "augustus")
        dat = Replace(dat, "september", "sep")
        dat = Replace(dat, "sept", "sep")
        dat = Replace(dat, "sep", "september")
        dat = Replace(dat, "okt", "oktober")
        dat = Replace(dat, "nov", "november")
        dat = Replace(dat, "dec", "december")
        
        dat = Replace(dat, "ma ", "")
        dat = Replace(dat, "di ", "")
        dat = Replace(dat, "wo ", "")
        dat = Replace(dat, "do ", "")
        dat = Replace(dat, "vr ", "")
        dat = Replace(dat, "za ", "")
        dat = Replace(dat, "zo ", "")
        dat = Replace(dat, "maandag ", "")
        dat = Replace(dat, "dinsdag ", "")
        dat = Replace(dat, "woensdag ", "")
        dat = Replace(dat, "donderdag ", "")
        dat = Replace(dat, "vrijdag ", "")
        dat = Replace(dat, "zaterdag ", "")
        dat = Replace(dat, "zondag ", "")
    End If
    
    
    '' Plak vierde sectie in heenOfterug
    If larryLength >= 3 Then
        heenOfterug = Trim(larry(3))
        If heenOfterug = "heen" Then heenOfterug = "heenrit"
        If heenOfterug = "h" Then heenOfterug = "heenrit"
        If heenOfterug = "terug" Then heenOfterug = "terugrit"
        If heenOfterug = "t" Then heenOfterug = "terugrit"
        If heenOfterug = "retour" Then heenOfterug = "terugrit"
        If heenOfterug = "r" Then heenOfterug = "terugrit"
    End If
    


    '' Melder pakken voor aanhef
    Dim melder As String
    melder = grijpMelder(newMail)
   
    
    '' *********************************************
    '' ** Body text kiezen en variabelen invullen **
    '' *********************************************
    aanOfAfmelding = "Afmelding"
    myHTMLText = deAanhef & "<BR><BR><BR>"
    myHTMLText = Replace(myHTMLText, "[naamMelder]", blauw(melder))
    
    '' Als we niet genoeg argumenten krijgen dan alle drie de tekstopties printen om door gebruiker zelf te editten
    If larryLength <= 1 Then
        
        myHTMLText = myHTMLText & vanTot & "<BR><BR>"
        myHTMLText = myHTMLText & opDatum & "<BR><BR>"
        myHTMLText = myHTMLText & heenterug & "<BR>"
        
        If larryLength = 1 Then
            myHTMLText = Replace(myHTMLText, "[naamClient]", blauw(clientvolnaam))
        End If
        
        myHTMLText = Replace(myHTMLText, "[naamClient]", blauw("naamClient"))
        myHTMLText = Replace(myHTMLText, "[vanDatum]", blauw("vanDatum"))
        myHTMLText = Replace(myHTMLText, "[totDatum]", blauw("totDatum"))
        myHTMLText = Replace(myHTMLText, "[opDatum]", blauw("opDatum"))
        myHTMLText = Replace(myHTMLText, "[heenTerug]", blauw("heenOfTerug"))
    Else
        If (heenOfterug = "heenrit" Or heenOfterug = "terugrit") Then
            dat = week(dat)
            myHTMLText = myHTMLText & heenterug
            myHTMLText = Replace(myHTMLText, "[naamClient]", blauw(clientvolnaam))
            myHTMLText = Replace(myHTMLText, "[heenTerug]", blauw(heenOfterug))
            myHTMLText = Replace(myHTMLText, "[opDatum]", blauw(dat))
        Else
            dat = Replace(dat, " tm ", " t/m ")
            If (InStr(dat, "t/m")) Then
                Dim dats() As String
                dats = Split(dat, "t/m")
                dats(0) = week(dats(0))
                dats(1) = week(dats(1))
                dat = dats(0) & " t/m " & dats(1)
                myHTMLText = myHTMLText & vanTot
                myHTMLText = Replace(myHTMLText, "[naamClient]", blauw(clientvolnaam))
                myHTMLText = Replace(myHTMLText, "[vanDatum]", blauw(dats(0)))
                myHTMLText = Replace(myHTMLText, "[totDatum]", blauw(dats(1)))
            Else
                If Mid(dat, 1, 4) = "ziek" Then
                    dat = Trim(Replace(dat, "ziek", ""))
                    dat = week(dat)
                    myHTMLText = myHTMLText & ziekGemeld
                    myHTMLText = Replace(myHTMLText, "[naamClient]", blauw(clientvolnaam))
                    myHTMLText = Replace(myHTMLText, "[per]", blauw(dat))
                    aanOfAfmelding = "Ziekmelding"
                Else
                    If Mid(dat, 1, 5) = "beter" Then
                        dat = Trim(Replace(dat, "beter", ""))
                        dat = week(dat)
                        myHTMLText = myHTMLText & beterGemeld
                        myHTMLText = Replace(myHTMLText, "[naamClient]", blauw(clientvolnaam))
                        myHTMLText = Replace(myHTMLText, "[per]", blauw(dat))
                        aanOfAfmelding = "Betermelding"
                    Else
                        dat = week(dat)
                        myHTMLText = myHTMLText & opDatum
                        myHTMLText = Replace(myHTMLText, "[naamClient]", blauw(clientvolnaam))
                        myHTMLText = Replace(myHTMLText, "[opDatum]", blauw(dat))
                    End If
                End If
            End If
        End If
    End If
    
    myHTMLText = myHTMLText & "<BR><BR>" & nogVragen & "<BR>"
    myHTMLText = "<span style=" & Chr(34) & "font-family:verdana;font-size:9pt" & Chr(34) & ">" & myHTMLText & "</span>"
    
   
    '' Als we een lijstje hebben dan de hoofdletters weer terugzetten voor de subject
    clientvolnaam = Replace(clientvolnaam, " genoemde ", " Genoemde ")
    clientvolnaam = Replace(clientvolnaam, " onderstaande ", " Onderstaande ")
    
    '' We korten de hele boel weer af voor het subject
    dat = Replace(dat, "maart", "mrt")
    dat = Replace(dat, "april", "apr")
    ''dat = Replace(dat, "juni", "jun")
    ''dat = Replace(dat, "juli", "jul")
    dat = Replace(dat, "augustus", "aug")
    dat = Replace(dat, "september", "sep")
    dat = Replace(dat, "oktober", "okt")
    dat = Replace(dat, "november", "nov")
    dat = Replace(dat, "december", "dec")
    
    dat = Replace(dat, "maandag", "ma")
    dat = Replace(dat, "dinsdag", "di")
    dat = Replace(dat, "woensdag", "wo")
    dat = Replace(dat, "donderdag", "do")
    dat = Replace(dat, "vrijdag", "vr")
    dat = Replace(dat, "zaterdag", "za")
    dat = Replace(dat, "zondag", "zo")
    
    '' Gegevens verwerkt, controleer met gebruiker en vraag of we door mogen gaan
    mySubject = newMail.Subject
    newSubject = doel & " - " & clientvolnaam & " - " & aanOfAfmelding & " " & dat & " " & heenOfterug
    Dim door As Integer
    door = vbYes

    
    '' Om te controleren of er een komma-gescheiden subject is getypt kijken we naar de clientvolnaam
    '' MsgBox tel & Chr(13) & clientvolnaam & Chr(13) & "--"
    
    If clientvolnaam = "" Then
    
        If tel > 0 Then
            door = MsgBox("Doorgaan met standaard antwoord opties?", vbQuestion + vbYesNoCancel)
            If door = vbYes Then newSubject = mySubject
        Else
            '' Omdat het goed mogelijk is dat er nog geen Enter is gegeven in het subjectveld doen we het eerst even zelf
            SendKeys "{Enter}", True
            tel = tel + 1
            Call InsertText
            Exit Sub
        End If
    End If
    
    If door = vbCancel Then
        tel = 0
        Exit Sub
    End If
    
    If door = vbYes Then
    
        '' Geef onderwerpveld nieuwe subject
        newMail.Subject = newSubject
    
        '' Check op welke manier de reply is opengezet en probeer de nieuwe body text erin te proppen
        plakTextInBody myHTMLText, newMail, oInspector
    
        '' Corrigeer To en CC velden en verwijder onszelf
        corrigeerToCC newMail
       
        '' Wij zijn nu eenmaal feilloos dus we verzenden ook maar meteen de mail
        '' newMail.Send
    End If
    
    If door = vbNo Then
        If tel <= 1 Then
                
            '' Omdat het goed mogelijk is dat er nog geen Enter is gegeven in het subjectveld doen we het eerst even zelf
            SendKeys "{Enter}", True
            tel = tel + 1
            Call InsertText
            Exit Sub
            
        End If
    End If
    
    '' Reset teller
    tel = 0
    
    
    
End Sub

Sub corrigeerToCC(newMail As MailItem)

Dim zelf As String
Dim tos() As String

    On Error Resume Next
    
    '' Stop eigen naam (Facilitair) in zelf
    zelf = newMail.Sender.Name
    
    '' stop alle ontvangers uit To veld in array
    tos = Split(newMail.To, ";")
        
    '' Loop alle ontvangers door en kijk of het de eerste melder is en anders naar CC schoppen
    For Each Recipient In newMail.Recipients
      If Recipient.Name <> tos(0) Then Recipient.Type = olCC
    Next Recipient
    
    '' Loop ze nog allemaal eens door en verwijder de ontvanger als we het zelf zijn (lukt niet om in 1 loop te doen blijkbaar)
    For Each Recipient In newMail.Recipients
      If Recipient.Name = zelf Then Recipient.Delete
    Next Recipient
    
    '' Even netjes alles checken
    newMail.Recipients.ResolveAll

End Sub

Sub plakTextInBody(myHTMLText As String, newMail As MailItem, oInspector As Inspector)
    '' Check op welke manier de reply is opengezet en probeer de nieuwe body text erin te proppen
    If oInspector Is Nothing Then

        Select Case newMail.BodyFormat
            Case olFormatPlain, olFormatRichText, olFormatUnspecified
                newMail.Body = RemoveHTML(myHTMLText) & newMail.Body
            Case olFormatHTML
                newMail.HTMLBody = myHTMLText & newMail.HTMLBody
        End Select

    Else
        If oInspector.IsWordMail Then
        MsgBox "Dit is experimenteel. Beter is gewoon niet een los window openen om deze macro te gebruiken"
            ' Hurray. We can use the rich Word object model, with access
            ' the caret and everything.
            Dim oDoc As Object, oWrdApp As Object, oSelection As Object
            Set oDoc = oInspector.WordEditor
            Set oWrdApp = oDoc.Application
            Set oSelection = oWrdApp.Selection
            oSelection.InsertAfter RemoveHTML(myHTMLText)
            oSelection.Collapse 0
            Set oSelection = Nothing
            Set oWrdApp = Nothing
            Set oDoc = Nothing
        Else
        
        MsgBox "Dit is experimenteel. Beter is gewoon niet een los window openen om deze macro te gebruiken"
            ' No object model to work with. Must manipulate raw text.
            Select Case newMail.BodyFormat
                Case olFormatPlain, olFormatRichText, olFormatUnspecified
                    newMail.Body = newMail.Body & RemoveHTML(myHTMLText)
                Case olFormatHTML
                    newMail.HTMLBody = newMail.HTMLBody & "<p>" & myHTMLText & "</p>"
            End Select
        End If
    End If
End Sub

Function grijpMelder(newMail As MailItem) As String

    Dim txt As String
    Dim onderwerpPos1 As String
    Dim onderwerpPos2 As String
    Dim groetLen As Integer
    Dim groetenPos As Integer
    Dim gokNaam As String
    Dim melder As String
    Dim melders() As String
    Dim melderVoornaam As Integer
    Dim groetArray() As Variant
    Dim tussenvoegsel() As Variant
    Dim voegsel As Variant
    Dim senderHadGetal As Boolean
    
   
    '' alle ontvangers uitsplitsen
    melders = Split(newMail.To, ";")
    '' eerste ontvanger selecteren uit rij
    melder = melders(0)
    '' als er ergens een getal voorkomt is het waarschijnlijk een adres en niet een persoonsnaam
    senderHadGetal = getalGevonden(melder)
    '' splitsen op komma om voornaam te proberen te pakken
    melders = Split(melder, ",")
    '' kijken of we 1 of 2 elementen hebben nu zodat we mogelijk een voornaam hebben
    melderVoornaam = UBound(melders) - LBound(melders)
    '' laatste element pakken, als er een komma was hebben we waarschijnlijk wel de voornaam te pakken
    '' als het een enkele naam betreft kan straks nog gekeken worden of er een getal in zit of dat
    '' het aantal elementen dus 0 is en daardoor waarschijnlijk geen voornaam is
    melder = melders(melderVoornaam)
    '' voornaam kan bestaan uit voornaam plus tussenvoegsels
    tussenvoegsels = Array("van het", "van der", "van den", "van de", " aan", " bij", " in", " onder", " van", " den", " ten", " 't", " het", " de")
    For Each voegsel In tussenvoegsels
        melder = Replace(melder, voegsel, "")
    Next
    '' eventuele spaties eraf trimmen
    melder = Trim(melder)
    '' alvast gevonden voornaam optie in variabel zetten
    grijpMelder = melder
     
    '' Kijk of in gevonden meldernaam een getal zit of dat er geen sprake was van een komma-gescheiden meldernaam
    If (senderHadGetal Or melderVoornaam = 0) Then
        txt = newMail.Body
        onderwerpPos1 = InStr(1, txt, "Onderwerp: ")
        onderwerpPos2 = InStr(onderwerpPos1 + 10, txt, "Onderwerp: ")
        If onderwerpPos2 = 0 Then onderwerpPos2 = Len(txt)
        
        txt = LCase(Mid(txt, onderwerpPos1, onderwerpPos2 - onderwerpPos1 + 1))
        
        groetArray = Array("voorbaat dank,", "melder:", "mvgr", "mvrgr", "groet,", "groeten,", "groetjes van", "groetjes,", "groet:", "groeten:", "groetjes:", "groet van", "groet;", "groeten;", "groetjes;", "groetjes", "groeten", "groet", "dank!", "gr.", "mvg", "m.v.g.", "thanks,", "gr ", " gr ")
        
        For Each groet In groetArray
            groetLen = Len(groet)
            groetenPos = InStr(txt, groet)
            If groetenPos > 0 Then Exit For
        Next
        
        gokNaam = Mid(txt, groetenPos + groetLen, 25)
        gokNaam = Replace(gokNaam, Chr(13), " ")
        gokNaam = Replace(gokNaam, Chr(10), " ")
        gokNaam = Replace(gokNaam, Chr(12), " ")
        gokNaam = smartTrim(gokNaam)
        gokNaam = Trim(gokNaam)
        
        grijpMelder = mooi(gokNaam)
        
    End If
    
    '' Soms gaat het toch mis dus even paar fixjes
    If grijpMelder = "Derwerp:" Or grijpMelder = "Derwerp" Or grijpMelder = "Erwerp" Or grijpMelder = "Van" Or grijpMelder = "" Then grijpMelder = "melder"
    
End Function



Function GetCurrentItem() As Object
    Dim objApp As Outlook.Application

    Set objApp = Application

    On Error Resume Next
    Select Case TypeName(objApp.ActiveWindow)
        Case "Explorer"
            ' Set GetCurrentItem = objApp.ActiveExplorer.Selection.Item(1)
            Set GetCurrentItem = objApp.ActiveExplorer.ActiveInlineResponse
        Case "Inspector"
            Set GetCurrentItem = objApp.ActiveInspector.CurrentItem
    End Select

    Set objApp = Nothing
    
End Function



Function mooi(Sent As String) As String

    ''Dim multi() As String
    ''multi = Split(Sent, " ")
    
    ''For Each txt In multi
    ''    mooi = mooi & " " & UCase(Left(txt, 1)) & Mid(txt, 2)
    ''Next

    ''mooi = Trim(mooi)
    
    mooi = Trim(StrConv(Sent, vbProperCase))

End Function



Function gevondenAantal(txt As String, vind As String) As Integer

    Dim plek As Integer
    plek = InStr(txt, vind)

    While plek > 0
    
        gevondenAantal = gevondenAantal + 1
        plek = InStr(plek + Len(vind), txt, vind)

    Wend

End Function



Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean

    IsInArray = (UBound(Filter(arr, stringToBeFound)) > -1)
  
End Function



Function blauw(txt As String) As String
    blauw = "<span style=" & Chr(34) & "color:blue" & Chr(34) & ">" & txt & "</span>"
End Function



Function getalGevonden(txt As String) As Boolean
    Dim control As Boolean
    Dim i As Integer
    control = False
    For i = 1 To Len(txt)
        If IsNumeric(Mid(txt, i, 1)) Then control = True
    Next i
    getalGevonden = control
End Function



Function isLetter(txt As String) As Boolean
''MsgBox letter & Chr(13) & Asc(letter)
''    If letter = Null Then
''        isLetter = False
''        Return
''    End If
    
    isLetter = ((Asc(txt) >= 65 And Asc(txt) <= 90) Or (Asc(txt) >= 97 And Asc(txt) <= 122) Or (Asc(txt) >= 128 And Asc(txt) <= 165) Or (Asc(txt) >= 198 And Asc(txt) <= 237))
    
    
   '' isLetter = UCase(letter) = Not LCase(letter)

End Function



Function smartTrim(txt As String) As String
'' deze functie loopt de string door en neemt alleen letters over en haakt af bij een spatie
    Dim i As Integer
    Dim letter As String
    For i = 1 To Len(txt)
        letter = Mid(txt, i, 1)
        If isLetter(letter) Then smartTrim = smartTrim & letter
        If Len(smartTrim) > 1 And letter = " " Then Exit For
    Next i
    
End Function



Function RemoveHTML(text As String) As String
    Dim regexObject As Object
    Set regexObject = CreateObject("vbscript.regexp")

    With regexObject
        .Pattern = "<!*[^<>]*>"    'html tags and comments
        .Global = True
        .IgnoreCase = True
        .MultiLine = True
    End With

    text = Replace(text, "<BR>", Chr(13))
    text = text & Chr(13)
    RemoveHTML = regexObject.Replace(text, "")
End Function
