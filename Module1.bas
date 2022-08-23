Attribute VB_Name = "Module1"
Global tel As Integer
Public Const mijnFontSize As String = "10pt"
Public Const mijnFontFamily As String = "verdana"
Public Const stopTEXT As String = "Dit hulpmiddel werkt alleen goed als er geen andere mailtjes geopened zijn." & vbNewLine & vbNewLine & "Sluit eerst andere open mail windows en probeer het nog eens."
Public Const doelenLijst As String = _
"apeldoorn,bedum,druten,ermelo,julianadorp,monster,noorderbrug,noordwijk,regio zuid,wekerom,zeeland"

Function grijpEigenInitialen() As String
    ''''
    '''' Maakt initialen van eigen naam
    ''''
    Dim myNamespace As Outlook.NameSpace
    Dim naam() As String
    Dim i As Integer
    
    Set myNamespace = Application.GetNamespace("MAPI")
    
    naam() = Split(myNamespace.CurrentUser, ",")
    
    For i = UBound(naam) To LBound(naam) Step -1
    
       grijpEigenInitialen = grijpEigenInitialen & Mid(Trim(naam(i)), 1, 1)
    
    Next i
 
End Function


Sub AfmeldingVervoer(newmail As Outlook.MailItem, oInspector As Outlook.Inspector, onderwerp As String)
    
    '' init
    Dim myHTMLText     As String
    Dim deAanhef       As String
    Dim vanTot         As String
    Dim opDatum        As String
    Dim heenTerug      As String
    Dim ziekGemeld     As String
    Dim beterGemeld    As String
    Dim nogVragen      As String
    Dim mySubject      As String
    Dim newSubject     As String
    Dim aanOfAfmelding As String
    Dim larry()        As String
    Dim larryLength    As Integer
    Dim doel           As String
    Dim doelen()       As String
    Dim clientvolnaam  As String
    Dim dat            As String
    Dim heenOfterug    As String
    Dim dl             As Variant
    Dim melder         As String
    Dim perVanaf       As String
    Dim door           As Integer


''newmail = GetCurrentItem()
''onderwerp = newmail.Subject
Debug.Print "tel = " & tel & "  mailID = " & newmail.CreationTime & "  AfmeldingVervoer()1  newmail.Subject = " & newmail.Subject


If newmail.CreationTime <> "1-1-4501" Then
        Debug.Print "tel = " & tel & "  AfmeldingVervoer  newmail.CreationTime <> 1-1-4501 "
        Debug.Print " *** END ALL"
        MsgBox stopTEXT, vbCritical, "*** Foutje gevonden ***"
        End
End If

     
    ''''
    '''' tekst declaraties
    ''''
    deAanhef = "Beste [naamMelder],"
    '' vanTot triggert als 't/m' wordt gezien in datums
    vanTot = "Het vervoer (heen en retour) van  [naamClient]  is van  [vanDatum]  tot en met  [totDatum]  geannuleerd."
    '' opDatum triggert als er geen 't/m' of ' en ' in datum gezien wordt
    opDatum = "Het vervoer (heen en retour) van  [naamClient]  is op  [opDatum]  geannuleerd."
    '' heenTerug triggert wanneer er een extra komma met heenrit of terugrit gezien wordt
    heenTerug = "De [heenTerug]  van  [naamClient]  is voor  [opDatum]  geannuleerd."
    '' ziekGemeld triggert als er 'ziek' wordt gezien in datum
    ziekGemeld = "Het vervoer (heen en retour) van  [naamClient]  is [per] afgemeld tot nader order."
    '' beterGemeld triggert als er 'beter' wordt gezien in datum
    beterGemeld = "Het vervoer (heen en retour) van  [naamClient]  is [per] weer aangemeld."
    nogVragen = "Mocht u nog vragen hebben, neem dan gerust contact met ons op."
    
    
    '' *******************************************
    '' ** Subject veld uitpluizen en mooi maken **
    '' *******************************************
    
    dl = "leeg"
        
    '' Splits onderwerp string door slash / gescheiden en stop in array genaamd larry
    '' De RE: mag er ook af
    onderwerp = Replace(onderwerp, "RE: ", "")
    onderwerp = Replace(onderwerp, " t/m ", " tm ")
    larry() = Split(onderwerp, "/")
    larryLength = UBound(larry) - LBound(larry)
     
    
    '' Als onderwerp leeg is krijgen we errors dus afbreken
    If Trim(onderwerp) = "" Or Trim(onderwerp) = "RE:" Then Exit Sub
     
    '' Plak eerste sectie voor de komma in doel
    doel = Trim(LCase(larry(0)))
    
    '' Kijk of perceelnaam geheel of gedeeltelijk ingevuld is
    If doel = "nw" Then doel = "noordwijk"
    If doel = "nb" Then doel = "noorderbrug"
    If doel = "z" Then doel = "zeeland"
    If doel = "zuid" Then doel = "regio zuid"
    doelen = Split(doelenLijst, ",")
    If (UBound(Filter(doelen, doel)) > -1) Then
        dl = Filter(doelen, doel)
        doel = dl(0)
    End If
    doel = mooi(doel)
    
    
    If larryLength >= 1 Then
        '' Plak tweede woord in clientnaam
        clientvolnaam = mooi(Trim(larry(1)))
        
        '' Deze woorden hoeven niet met hoofdletter
        Dim verkleinLijst() As Variant
        verkleinLijst = Array(" En ", " Van ", " De ", " Der ", " Den ", " Op ", " Te ", "Begeleider", "Begeleiding", "Personen", "Genoemde", "Onderstaande", "Plus ")
        For Each woord In verkleinLijst
            clientvolnaam = Replace(clientvolnaam, woord, LCase(woord))
        Next
        
        '' Om ook een lijstje namen met komma te kunnen scheiden gebruiken we gewoon de punt
        clientvolnaam = Replace(clientvolnaam, ".", ",")
    End If
    
    
    If larryLength >= 2 Then
        '' Plak derde zin in dat en voeg een enkele spatie toe aan einde, dit om goed naar maanden en dagen te kunnen zoeken straks
        dat = LCase(Trim(larry(2))) & " "
        
        '' datums met komma scheiden kan als je de punt gebruikt maar weekdagen worden dan niet toegevoegd
        ''dat = Replace(dat, ".", ",")  uitgezet omdat nu slash gebruikt wordt ipv komma
        
        '' We zoeken straks zelf de dag op bij de datum dus hier strippen we alle dagaanduidingen
        dat = verwijderWeekdagen(dat)
        
        '' Alvast voor de bodytekst afgekorte datums voluit schrijven
        dat = maandenVoluit(dat)
        
        '' ziek- en betermelding opvangen met korte notatie z en b
        If dat = "z " Then dat = "ziek"
        If dat = "b " Then dat = "beter"
 
        '' Trim
        dat = Trim(dat)

    End If
    
    
    '' Plak eventuele vierde sectie in heenOfterug
    If larryLength >= 3 Then
        heenOfterug = Trim(LCase(larry(3)))
        If heenOfterug = "heen" _
        Or heenOfterug = "h" Then heenOfterug = "heenrit"
        If heenOfterug = "terug" _
        Or heenOfterug = "t" _
        Or heenOfterug = "retour" _
        Or heenOfterug = "r" Then heenOfterug = "terugrit"
    End If
    


    '' Melder pakken voor aanhef
    melder = grijpMelder(newmail)
   
    
    '' *********************************************
    '' ** Body text kiezen en variabelen invullen **
    '' *********************************************
    aanOfAfmelding = "Afmelding"
    myHTMLText = deAanhef & "<BR><BR><BR>"
    ''myHTMLText = Replace(myHTMLText, "[naamMelder]", blauw(melder))
    myHTMLText = Replace(myHTMLText, "[naamMelder]", melder)
    
    '' Als we niet genoeg argumenten krijgen dan alle drie de tekstopties klaarzetten om door gebruiker zelf te editten
    If larryLength <= 1 Then
        
        myHTMLText = myHTMLText & vanTot & "<BR><BR>"
        myHTMLText = myHTMLText & opDatum & "<BR><BR>"
        myHTMLText = myHTMLText & heenTerug & "<BR>"
        
        If larryLength = 1 Then myHTMLText = Replace(myHTMLText, "[naamClient]", blauw(clientvolnaam))
        
        myHTMLText = Replace(myHTMLText, "[naamClient]", blauw("naamClient"))
        myHTMLText = Replace(myHTMLText, "[vanDatum]", blauw("vanDatum"))
        myHTMLText = Replace(myHTMLText, "[totDatum]", blauw("totDatum"))
        myHTMLText = Replace(myHTMLText, "[opDatum]", blauw("opDatum"))
        myHTMLText = Replace(myHTMLText, "[heenTerug]", blauw("heenOfTerug"))
        
        
        
    Else
        '' We hebben minimaal 3 argumenten
        
        '' Maak gebruik van vandaag, morgen en overmorgen mogelijk
        dat = Replace(dat, "vandaag", Day(Now()) & " " & maandNaam(Month(Now())))
        dat = Replace(dat, "overmorgen", Day(Now() + 2) & " " & maandNaam(Month(Now() + 2)))
        dat = Replace(dat, "morgen", Day(Now() + 1) & " " & maandNaam(Month(Now() + 1)))
        
        If (heenOfterug = "heenrit" Or heenOfterug = "terugrit") Then
            '' Een getal zonder maand wordt aangezien als dag in huidige maand
            On Error Resume Next
            If Year(CDate(dat)) = "1900" Or Year(CDate(dat)) = "1899" Then dat = dat & " " & maandNaam(Month(Now()))
            '' kijk of datum voor vandaag ligt dan wordt aangenomen dat volgend jaar bedoeld wordt, tenzij er al een jaartal wordt meegegeven
            If Not jaarGenoemd(dat) _
            And CDate(Format(Now, "d-m-yyyy")) > CDate(dat) Then dat = dat & " " & Year(Now) + 1
            On Error GoTo 0
            dat = week(dat)
            dat = Trim(Replace(dat, " " & Year(Now) + 1, ""))
            myHTMLText = myHTMLText & heenTerug
            myHTMLText = Replace(myHTMLText, "[naamClient]", blauw(clientvolnaam))
            myHTMLText = Replace(myHTMLText, "[heenTerug]", blauw(heenOfterug))
            myHTMLText = Replace(myHTMLText, "[opDatum]", blauw(dat))
        Else
            dat = Replace(dat, " tm ", " t/m ")
            dat = Replace(dat, " tot en met ", " t/m ")
            dat = Replace(dat, " t.e.m. ", " t/m ")
            dat = Replace(dat, " t.e.m ", " t/m ")
            dat = Replace(dat, " tem ", " t/m ")
            If (InStr(dat, "t/m")) Then
                Dim dats() As String
                dats = Split(dat, "t/m")
                        
                '' Een getal zonder maand wordt aangezien als dag in huidige maand
                On Error Resume Next
                If Year(CDate(dats(0))) = "1900" Or Year(CDate(dats(0))) = "1899" Then dats(0) = dats(0) & " " & maandNaam(Month(Now()))
                If Year(CDate(dats(1))) = "1900" Or Year(CDate(dats(1))) = "1899" Then dats(1) = dats(1) & " " & maandNaam(Month(Now()))
                '' kijk of eerste datum voor vandaag ligt dan wordt aangenomen dat volgend jaar bedoeld wordt, tenzij er al een jaartal wordt meegegeven
                If Not jaarGenoemd(dats(0)) _
                And CDate(Format(Now, "d-m-yyyy")) > CDate(dats(0)) Then dats(0) = dats(0) & " " & Year(Now) + 1
                '' kijk of tweede datum voor de eerste ligt dan wordt aangenomen dat volgend jaar bedoeld wordt, tenzij er al een jaartal wordt meegegeven
                If Not jaarGenoemd(dats(1)) _
                And CDate(dats(0)) > CDate(dats(1)) Then dats(1) = dats(1) & " " & Year(Now) + 1
                On Error GoTo 0
                
                '' Nu de weekdag erbij zoeken
                dats(0) = week(Trim(dats(0)))
                dats(1) = week(Trim(dats(1)))
                
                '' eventuele jaartal mag er weer af
                dats(0) = Replace(dats(0), " " & Year(Now) - 1, "")
                dats(1) = Replace(dats(1), " " & Year(Now) - 1, "")
                dats(0) = Replace(dats(0), " " & Year(Now), "")
                dats(1) = Replace(dats(1), " " & Year(Now), "")
                dats(0) = Replace(dats(0), " " & Year(Now) + 1, "")
                dats(1) = Replace(dats(1), " " & Year(Now) + 1, "")
                
                dat = dats(0) & " t/m " & dats(1)

                myHTMLText = myHTMLText & vanTot
                myHTMLText = Replace(myHTMLText, "[naamClient]", blauw(clientvolnaam))
                myHTMLText = Replace(myHTMLText, "[vanDatum]", blauw(dats(0)))
                myHTMLText = Replace(myHTMLText, "[totDatum]", blauw(dats(1)))
            Else
                If Mid(dat, 1, 4) = "ziek" Then
                    dat = Trim(Replace(dat, "ziek", ""))
                    If Len(Trim(dat)) > 0 Then
                        perVanaf = ""
                        If InStr(dat, "per ") Then perVanaf = "per "
                        dat = Replace(dat, perVanaf, "")
                        If InStr(dat, "vanaf ") Then perVanaf = "vanaf "
                        dat = Replace(dat, perVanaf, "")
                        '' Een getal zonder maand wordt aangezien als dag in huidige maand
                        On Error Resume Next
                        If Year(CDate(dat)) = "1900" Or Year(CDate(dat)) = "1899" Then dat = dat & " " & maandNaam(Month(Now()))
                        '' kijk of datum voor vandaag ligt dan wordt aangenomen dat volgend jaar bedoeld wordt, tenzij er al een jaartal wordt meegegeven
                        If Not jaarGenoemd(dat) _
                        And CDate(Format(Now, "d-m-yyyy")) > CDate(dat) Then dat = dat & " " & Year(Now) + 1
                        On Error GoTo 0
                        dat = week(dat)
                        dat = Trim(Replace(dat, " " & Year(Now) + 1, ""))
                        dat = perVanaf & dat
                    End If
                    myHTMLText = myHTMLText & ziekGemeld
                    myHTMLText = Replace(myHTMLText, "[naamClient]", blauw(clientvolnaam))
                    myHTMLText = Replace(myHTMLText, "[per]", blauw(dat))
                    aanOfAfmelding = "Ziekmelding"
                Else
                    If Mid(dat, 1, 5) = "beter" Then
                        dat = Trim(Replace(dat, "beter", ""))
                        If Len(Trim(dat)) > 0 Then
                            perVanaf = ""
                            If InStr(dat, "per ") Then perVanaf = "per "
                            dat = Replace(dat, perVanaf, "")
                            If InStr(dat, "vanaf ") Then perVanaf = "vanaf "
                            dat = Replace(dat, perVanaf, "")
                            '' Een getal zonder maand wordt aangezien als dag in huidige maand
                            On Error Resume Next
                            If Year(CDate(dat)) = "1900" Or Year(CDate(dat)) = "1899" Then dat = dat & " " & maandNaam(Month(Now()))
                            '' kijk of datum voor vandaag ligt dan wordt aangenomen dat volgend jaar bedoeld wordt, tenzij er al een jaartal wordt meegegeven
                            If Not jaarGenoemd(dat) _
                            And CDate(Format(Now, "d-m-yyyy")) > CDate(dat) Then dat = dat & " " & Year(Now) + 1
                            On Error GoTo 0
                            dat = week(dat)
                            dat = Trim(Replace(dat, " " & Year(Now) + 1, ""))
                            dat = perVanaf & dat
                        End If
                        
                        myHTMLText = myHTMLText & beterGemeld
                        myHTMLText = Replace(myHTMLText, "[naamClient]", blauw(clientvolnaam))
                        myHTMLText = Replace(myHTMLText, "[per]", blauw(dat))
                        aanOfAfmelding = "Betermelding"
                    Else
                        If (InStr(dat, " en ")) Or (InStr(dat, ",")) Then
                            Dim datt() As String
                            dat = Replace(dat, ",", " en ")
                            datt = Split(dat, " en ")
                            '' Een getal zonder maand wordt aangezien als dag in huidige maand
                            On Error Resume Next
                            For i = 0 To UBound(datt)
                                If Year(CDate(datt(i))) = "1900" Or Year(CDate(datt(i))) = "1899" Then datt(i) = datt(i) & " " & maandNaam(Month(Now()))
                            Next i
                            '' kijk of eerste datum voor vandaag ligt dan wordt aangenomen dat volgend jaar bedoeld wordt, tenzij er al een jaartal wordt meegegeven
                            If Not jaarGenoemd(datt(0)) _
                            And CDate(Format(Now, "d-m-yyyy")) > CDate(datt(0)) Then datt(0) = datt(0) & " " & Year(Now) + 1
                            '' kijk of tweede datum voor de eerste ligt dan wordt aangenomen dat volgend jaar bedoeld wordt, tenzij er al een jaartal wordt meegegeven
                            For i = 1 To UBound(datt)
                                If Not jaarGenoemd(datt(i)) _
                                And CDate(datt(i - 1)) > CDate(datt(i)) Then datt(i) = datt(i) & " " & Year(Now) + 1
                            Next i
                            On Error GoTo 0
                            
                            '' Nu de weekdag erbij zoeken
                            '' eventuele jaartal mag er weer af
                            For i = 0 To UBound(datt)
                                datt(i) = week(Trim(datt(i)))
                                datt(i) = Replace(datt(i), " " & Year(Now) - 1, "")
                                datt(i) = Replace(datt(i), " " & Year(Now), "")
                                datt(i) = Replace(datt(i), " " & Year(Now) + 1, "")
                            Next i
                            
                            '' Lijstje datums netjes klaarmaken
                            dat = datt(0)
                            For i = 1 To UBound(datt)
                                If i = UBound(datt) Then
                                    dat = dat & zwart(" en ") & datt(i)
                                Else
                                    dat = dat & zwart(", ") & datt(i)
                                End If
                            Next i
                        Else
                            If getalGevonden(dat) Then
                                '' Een getal zonder maand wordt aangezien als dag in huidige maand
                                On Error Resume Next
                                If Year(CDate(dat)) = "1900" Or Year(CDate(dat)) = "1899" Then dat = dat & " " & maandNaam(Month(Now()))
                                '' kijk of datum voor vandaag ligt dan wordt aangenomen dat volgend jaar bedoeld wordt, tenzij er al een jaartal wordt meegegeven
                                If Not jaarGenoemd(dat) _
                                And CDate(Format(Now, "d-m-yyyy")) > CDate(dat) Then dat = dat & " " & Year(Now) + 1
                                On Error GoTo 0
                                dat = week(dat)
                                dat = Replace(dat, " " & Year(Now) - 1, "")
                                dat = Replace(dat, " " & Year(Now), "")
                                dat = Replace(dat, " " & Year(Now) + 1, "")
                            End If
                        End If
                        
                        myHTMLText = myHTMLText & opDatum
                        myHTMLText = Replace(myHTMLText, "[naamClient]", blauw(clientvolnaam))
                        myHTMLText = Replace(myHTMLText, "[opDatum]", blauw(dat))
                    End If
                End If
            End If
        End If
    End If
    
    myHTMLText = myHTMLText & "<BR><BR>" & nogVragen ''& "<BR>"
    myHTMLText = "<span style=" & Chr(34) & "font-family:" & mijnFontFamily & ";font-size:" & mijnFontSize & Chr(34) & ">" & myHTMLText & "</span>"
    
    
    ''''
    '''' Body is klaar, nu het onderwerpveld vorm geven
    ''''
    
   
    '' Als we een groep namen kregen dan deze woorden met hoofdletters weer terugzetten in het subject omdat dat mooier staat
    clientvolnaam = Replace(clientvolnaam, "genoemde ", "Genoemde ")
    clientvolnaam = Replace(clientvolnaam, "onderstaande ", "Onderstaande ")
    
    '' nu dat opmaken voor gebruik in onderwerpveld dus HTML eraf halen
    dat = RemoveHTML(dat)
    '' Er kruipt soms een Enter aan het eind waardoor heenrit of terugrit niet wordt vermeld
    dat = Replace(dat, Chr(13), "")
    
    '' We korten de maand- en weeknamen weer af voor in het subject
    dat = maandenVerkleind(dat)
    dat = wekenVerkleind(dat)
    
    '' Gegevens verwerkt, nieuw onderwerp kan opgemaakt worden
    mySubject = newmail.Subject
    newSubject = doel & " - " & clientvolnaam & " - " & aanOfAfmelding & " " & dat & " " & heenOfterug
    
    

        Debug.Print "tel = " & tel & "  mailID = " & newmail.CreationTime & "  AfmeldingVervoer()2  newmail.Subject = " & newmail.Subject
    

    ''
    '' Om te controleren of er een komma-gescheiden subject is getypt kijken we hier naar de clientvolnaam
    '' Een lege clientvolnaam betekent dat er helemaal geen komma's zijn gebruikt in het subject
    ''
    
    door = vbYes
    
    If clientvolnaam = "" Then
    
        If tel > 0 Then
            door = MsgBox("Onderwerpveld als volgt opmaken:" & vbNewLine & "Regio / Clientnaam ( / Datum afmelding ( / Heenrit/terugrit ))" & _
            vbNewLine & vbNewLine & "Doorgaan met de standaardzinnen?", vbQuestion + vbYesNo, _
            "Geen gegevens gevonden in onderwerpveld!")
            If door = vbYes Then newSubject = mySubject '' Subject mag zometeen blijven staan zoals hij was
        Else
            '' Omdat het goed mogelijk is dat er nog geen Enter is gegeven in het subjectveld doen we het eerst even zelf
            SendKeys "{Enter}", True
            '' Nu gaan we deze routine nog 1 keer runnen dus houden we een global teller bij
            tel = tel + 1
            Call InsertText
            '' Verder kunnen we deze instantie opruimen en afsluiten
            Set newmail = Nothing
            Set oInspector = Nothing
            Exit Sub
        End If
        
    End If
    
    '' Als op Cancel wordt gedrukt dan doen we helemaal niks meer
    If door = vbNo Then
        tel = 0
        Set newmail = Nothing
        Set oInspector = Nothing
        Exit Sub
    End If
    
            
    If door = vbYes Then
    
        '' Geef onderwerpveld nieuwe subject
        newmail.Subject = newSubject
        Debug.Print "tel = " & tel & "  mailID = " & newmail.CreationTime & "  AfmeldingVervoer()3  newmail.Subject = " & newmail.Subject
        
    
        '' Check op welke manier de reply is opengezet en probeer de nieuwe body text erin te proppen
        plakTextInBody myHTMLText, newmail, oInspector
    
        '' Corrigeer To en CC velden en verwijder onszelf
        corrigeerToCC newmail
       
        '' Wij zijn nu eenmaal feilloos dus we verzenden ook maar meteen de mail
        '' newMail.Send
    End If

        
    '' Reset global teller
    tel = 0

    '' Klaar
    Set newmail = Nothing
    Set oInspector = Nothing


    '' Alvast stempeltje klaarzetten
    ''Call clipboardStempel
        
    
End Sub



Sub clipboardStempel()
    ''''
    '''' Zet in paste geheugen wie en door welke mail de vervoerswijziging is gedaan
    ''''
    Dim oMail As MailItem
    Dim melder As String
    Dim verzonden As String
    
    On Error Resume Next
    
    Set oMail = GetCurrentItem()
    melder = grijpMelder(oMail)
    verzonden = grijpVerzonden(oMail)
        
    Clipboard (grijpEigenInitialen & " iov " & melder & " mail van " & Format(verzonden, "d mmmm"))
    
    Set oMail = Nothing

End Sub

Sub zetBesteMelder(s As Boolean)
    ''''
    '''' Stop bovenin body de aanhef 'Beste melder,'
    '''' Boolean switch s om wel of geen blauwe meldernaam neer te zetten
    ''''
    Dim oMail As MailItem
    Dim oInspector As Inspector
    Dim melder As String
    Dim myHTMLText As String
    Dim myPlainText As String
    
    Set oMail = GetCurrentItem()
    
    melder = grijpMelder(oMail)
        
    myPlainText = "Beste " & melder & ","
    
    If s Then melder = blauw(melder)
    
    myHTMLText = "Beste " & melder & ","
    
    myHTMLText = "<span style=" & Chr(34) & "font-family:" & mijnFontFamily & ";font-size:" & mijnFontSize & Chr(34) & ">" & myHTMLText & "</span>"
    
    '' Check of Beste melder al te vinden is op eerste regel, anders neerzetten
    If Mid(oMail.Body, 1, Len(myPlainText)) <> myPlainText Then plakTextInBody myHTMLText, oMail, oInspector
    
    Set oMail = Nothing
    Set oInspector = Nothing
End Sub
Sub besteMelder()
    ''''
    '''' Stop bovenin body de aanhef 'Beste melder,'
    '''' Boolean switch om geen blauwe meldernaam neer te zetten true = blauw  false = zwart
    ''''

    Call zetBesteMelder(False)
    
End Sub

Sub fixToCC()
    ''''
    '''' Loop 'To' veld door en verhuis alle geaddresseerden behalve originele zender naar 'CC' veld
    '''' Ook eigen email verwijderen uit 'To' veld
    ''''
    
    Call corrigeerToCC(GetCurrentItem())

End Sub

Sub aanhefEnFixCC()
    ''''
    '''' Loop 'To' veld door en verhuis alle geaddresseerden behalve originele zender naar 'CC' veld
    '''' Ook eigen email verwijderen uit 'To' veld
    ''''
    '''' Stop bovenin body de aanhef 'Beste melder,'
    '''' Boolean switch om geen blauwe meldernaam neer te zetten true = blauw  false = zwart
    ''''

    fixToCC
    Call zetBesteMelder(False)
    '' quick and dirty cursor plaatsen om meteen te typen
    SendKeys "{DOWN}", True
    SendKeys "{Enter}", True
    
End Sub



Sub zeven24()
    ''''
    '''' Vraagt gebruiker om input en vervangt binnen body tekst de betreffende woorden met die input
    '''' Zoekt in body naar string binnen 'startcode' en 'eindcode'
    '''' Kan ook het onderwerpveld aanpassen
    ''''
    '''' Syntax code:
    ''''           [[  zoekwoord , vervangMet , stelVraag , vraagTitel , defaultInput  ]]
    ''''
    '''' zoekwoord      Zoekt naar dit woord binnen body
    '''' vervangMet     Vervangt zoekwoord met dit woord, mag combinatie van tekst en opgegeven zoekwoord zijn
    '''' stelVraag      Opent tekst input window met deze string als vraag
    '''' vraagTitel     Het geopende window kun je hiermee ook een titel geven in de titelbalk
    '''' defaultInput   Hiermee kun je eventueel een standaardwaarde alvast in de inputbalk zetten
    ''''
    '''' Voor het onderwerpveld is een speciaal zoekwoord beschikbaar: 'subject'
    '''' Hierin is het tweede argument (vervangMet) ook de andere zoekwoord resultaten in te verwerken
    ''''
    '''' Plaats het onderstaande op de eerste regels voor de sjabloontekst in de handtekening
    '''' Voorbeeld:
    ''''
    ''''     [[planOnNummer, planOnNummer - subject, PlanOn nummer :, Zoef zoef]]
    ''''     [[24uOf7dagen, 24uOf7dagen, 24 uur of 7 dagen?, Sjongejonge, 7 dagen]]
    ''''     [[subject, Melding planOnNummer - subject]]
    ''''
    '''' Resultaat:
    ''''
    ''''  Toont eerste een dialoogvenster met de vraag 'Planon nummer :' en titel 'Zoef zoef'
    ''''  De inputbox is leeg
    ''''
    ''''  Toont daarna nog een dialoogvenster met de vraag '24 uur of 7 dagen?' en titel 'Sjongejonge'
    ''''  In de inputbox is alvast '7 dagen' ingevuld en geselecteerd zodat je meteen kan typen en het de standaardwaarde vervangt
    ''''
    ''''  Als laatste gaat hij zonder vraag te stellen het subject vervangen door wat de tekst:
    ''''  'Melding ' plus wat je bij de eerste vraag hebt geantwoord, dan een streepje en dan de originele onderwerpregel
    ''''  (Zal automatisch elk woord met hoofdletter plaatsen voor de mooiheid)
    ''''

    

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
                If StrPtr(.antwoord) = 0 Then Exit Sub '' Cancel gedrukt
                .antwoord = Replace(.antwoord, ".00", "")
                .antwoord = Replace(.antwoord, ".0", "")
                .vervangMet = Replace(.vervangMet, .zoekWoord, .antwoord)
            End If
            
            
            offset = .eindPos + 2
            cnt = cnt + 1
        End With
    Wend
    cnt = cnt - 1
    '' Als er geen code gevonden is dan stilletjes eindigen
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
    zetBesteMelder True
    
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
    
        MsgBox (zoek & " niet gevonden" & Chr(13) & "Hierdoor kan ik de ontstane lege ruimte onder de aanhef niet weghalen.")
        
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
    
    Set oMail = Nothing
    
End Sub

Function maandNaam(i As Integer) As String
    ''''
    '''' Geeft naam van de maand aan cijfer
    ''''

    On Error Resume Next
    
    Dim maanden() As Variant
    
    maanden = Array("", "januari ", "februari ", "maart ", "april ", "mei ", "juni ", "juli ", "augustus ", "september ", "oktober ", "november ", "december ")
    
    maandNaam = maanden(i)
    
End Function



Function week(dat As String) As String
    ''''
    '''' Zoekt bij de opgegeven datum de naam van de dag op
    '''' Geeft anders zichzelf ongewijzigd weer terug
    ''''

    On Error Resume Next
    
    Dim weekdag() As Variant
    Dim resultaat As Variant
    
    weekdag = Array("", "zondag ", "maandag ", "dinsdag ", "woensdag ", "donderdag ", "vrijdag ", "zaterdag ")
    
    resultaat = Weekday(dat)
    week = dat
    If resultaat > 0 And resultaat < 8 Then week = weekdag(resultaat) & dat
    
End Function

Sub InsertText()
    ''''
    '''' Schiet AfmeldingVervoer af
    ''''

    Dim newmail        As Outlook.MailItem
    Dim oInspector     As Outlook.Inspector

    '' Check of mail inline of in eigen window getoond wordt
    Set oInspector = Application.ActiveInspector
    If oInspector Is Nothing Then
        Set newmail = GetCurrentItem()
    Else
        Set newmail = oInspector.CurrentItem
    End If
    
    Call AfmeldingVervoer(newmail, oInspector, newmail.Subject)
    
    Set newmail = Nothing
    Set oInspector = Nothing
    

End Sub


Sub test()

    Dim newmail        As Outlook.MailItem
    Dim oInspector     As Outlook.Inspector

    '' Check of mail inline of in eigen window getoond wordt
    Set oInspector = Application.ActiveInspector
    If oInspector Is Nothing Then
        Set newmail = GetCurrentItem()
    Else
        Set newmail = oInspector.CurrentItem
    End If
    
    For Each ond In Array("nb,Bas kerkhof,23 dec tm 12 jan", "nb,bas kerkhgof,23 dec en 12 jan")
        Call AfmeldingVervoer(newmail, oInspector, ond)
    Next
    
    Set newmail = Nothing
    Set oInspector = Nothing
    
End Sub

Function jaarGenoemd(dat As String) As Boolean
    ''''
    '''' True wanneer vorig jaar, dit jaar of volgend jaar genoemd wordt in string
    ''''
    
    jaarGenoemd = (InStr(dat, Year(Now()) - 1) _
    Or InStr(dat, Year(Now())) _
    Or InStr(dat, Year(Now()) + 1))
    
End Function

Function verwijderWeekdagen(dat As String) As String
    ''''
    '''' Verwijder de genoemde weekdagen uit de string
    ''''
    Dim weekdagen() As Variant
    
    weekdagen = Array( _
    "maa ", "din ", "woe ", "don ", "vrij ", "zat ", "zon ", _
    "ma ", "di ", "wo ", "do ", "vr ", "za ", "zo ", _
    "maandag ", "dinsdag ", "woensdag ", "donderdag ", "vrijdag ", "zaterdag ", "zondag ")
    
    For Each weekdag In weekdagen
        dat = Replace(dat, weekdag, "")
    Next
    
    verwijderWeekdagen = dat
        
End Function


Function maandenVoluit(dat As String) As String
    ''''
    '''' Schrijft maandnamen voluit als ze verkort aangetroffen worden
    ''''

    dat = Replace(dat, "jan ", "januari ")
    dat = Replace(dat, "feb ", "februari ")
    dat = Replace(dat, "mar ", "maart ")
    dat = Replace(dat, "mrt ", "maart ")
    dat = Replace(dat, "apr ", "april ")
    dat = Replace(dat, "jun ", "juni ")
    dat = Replace(dat, "jul ", "juli ")
    dat = Replace(dat, "aug ", "augustus ")
    dat = Replace(dat, "sept ", "sep ")
    dat = Replace(dat, "sep ", "september ")
    dat = Replace(dat, "okt ", "oktober ")
    dat = Replace(dat, "oct ", "oktober ")
    dat = Replace(dat, "nov ", "november ")
    dat = Replace(dat, "dec ", "december ")
    
    maandenVoluit = dat
    
    
End Function


Function maandenVerkleind(dat As String) As String
    ''''
    '''' Schrijft maandnamen verkleind als ze voluit geschreven aangetroffen worden
    ''''
    
    For Each maandWoord In Array("januari", "februari", "april", "juni", "juli", "augustus", "september", "oktober", "november", "december")
        dat = Replace(dat, maandWoord, Mid(maandWoord, 1, 3))
    Next
    
    dat = Replace(dat, "maart", "mrt")  '' omdat anders maart maa wordt en dan later mis kan gaan als maandag wordt gevonden via maa en dan mrtndag werd
    
    maandenVerkleind = dat
    
End Function


Function wekenVerkleind(dat As String) As String
    ''''
    '''' Schrijft weekdagen in verkleinde vorm als ze voluit geschreven aangetroffen worden
    ''''
        
    For Each weekdag In Array("maandag", "dinsdag", "woensdag", "donderdag", "vrijdag", "zaterdag", "zondag")
        dat = Replace(dat, weekdag, Mid(weekdag, 1, 2))
    Next
    
    wekenVerkleind = dat
    
End Function





Sub gokSDRegio()
''''
'''' Niet gebruikt
''''

 Dim tos() As String
 Dim weZoeken As String
 Dim weVonden As String
 Dim doelen() As Variant
 Dim apeldoorn() As Variant
 Dim bedum() As Variant
 Dim druten() As Variant
 Dim ermelo() As Variant
 Dim julianadorp() As Variant
 Dim monster() As Variant
 Dim noordwijk() As Variant
 Dim wekerom() As Variant
 Dim zeeland() As Variant
 
 '' https://stackoverflow.com/questions/8849357/add-quotation-at-the-start-and-end-of-each-line-in-notepad
 doelen = Array("apeldoorn", "bedum", "druten", "ermelo", "julianadorp", "monster", "noorderbrug", "noordwijk", "regio zuid", "wekerom", "zeeland")
 apeldoorn = Array("Beekbergen", "Vaassen", "Voorst")
 bedum = Array("Appingedam", "Delfzijl", "Harlingen", "Leeuwarden", "Zuidwolde")
 druten = Array("Geldermalsen", "Zaltbommel", "Woudrichem", "Beneden Leeuwen", "Boven Leeuwen", "Grave", "Kerk Avezaath", "Nijmegen", "Nuenen", "Tiel", "Wijchen")
 ermelo = Array("Almere", "Amersfoort", "Baarn", "Biddinghuizen", "Blaricum", "Bunschoten", "Spakenburg", "De Meern", "Dronten", "Elburg", "Elspeet", "Emmeloord", "Ens", "Epe", "Epe", "t Harde", "Harderwijk", "Hoogland", "Laren", "Leersum", "Lelystad", "Nijkerk", "Nijkerkerveen", "Nunspeet", "Oldebroek", "Oosterwolde", "Soest", "Swifterbant", "Urk", "Utrecht", "Vleuten", "Wezep", "Zeewolde")
 julianadorp = Array("Alkmaar", "Amsterdam", "Anna Paulowna", "Bovenkarpsel", "Breezand", "Den Burg", "Den Helder", "Driehuis", "Koedijk", "Nieuw Niedorp", "Schagen", "Sint Pancras", "t Zand", "Wieringwerf", "Zwaag")
 monster = Array("Den Haag", "Gravenzande", "De Lier", "Den Haag", "Poeldijk", "Schipluiden", "Wateringen")
 noordwijk = Array("Hillegom", "Hoofddorp", "Hoogmade", "Katwijk", "Leiden", "Leiderdorp", "Lisse", "Lisserbroek", "Nieuw Vennep", "Oegstgeest", "Rijnsburg", "Sassenheim", "Voorhout", "Vijfhuizen")
 wekerom = Array("Arnhem", "Barneveld", "Bennekom", "Ede", "De Glind", "Lunteren", "Renkum", "Veenendaal", "Wageningen")
 zeeland = Array("Aagtekerke", "Arnemuiden", "Borssele", "Gapinge", "Goes", "Heinkeszand", "Hulst", "Koudekerke", "Meliskerke", "Middelburg", "Nieuw- en Sint Joosland", "Nieuwdorp", "Nisse", "Oostkapelle", "Ritthem", "Serooskerke", "Veere", "Vlissingen")
 
 
 Dim newmail    As MailItem
 Dim oInspector As Inspector
 
 On Error Resume Next
 
     '' Check of mail inline of in eigen window getoond wordt
    Set oInspector = Application.ActiveInspector
    If oInspector Is Nothing Then
        Set newmail = GetCurrentItem()
    Else
        Set newmail = oInspector.CurrentItem
    End If
    
 
    '' stop alle ontvangers uit To veld in array
    tos = Split(newmail.To, ";")
    
    weZoeken = tos(0)
    weVonden = GetOutlookAddressBookProperty(weZoeken, "City")
    Debug.Print (weZoeken)
    Debug.Print (weVonden)
    Debug.Print ("--->")
    
    If IsInArray(weVonden, doelen) Then
        Debug.Print ("Joepie")
    Else
        If IsInArray(weVonden, apeldoorn) Then weVonden = "Apeldoorn"
            
        If IsInArray(weVonden, bedum) Then weVonden = "Bedum"
        If IsInArray(weVonden, druten) Then weVonden = "Druten"
        If IsInArray(weVonden, ermelo) Then weVonden = "Ermelo"
        If IsInArray(weVonden, julianadorp) Then weVonden = "Julianadorp"
        If IsInArray(weVonden, monster) Then weVonden = "Monster"
        If IsInArray(weVonden, noordwijk) Then weVonden = "Noordwijk"
        If IsInArray(weVonden, wekerom) Then weVonden = "Wekerom"
        If IsInArray(weVonden, zeeland) Then weVonden = "Zeeland"
    End If
    
    Debug.Print (weVonden)
    
    
    
    SendKeys weVonden, True
    
    newmail = Nothing
    
 
End Sub

Public Function GetOutlookAddressBookProperty(alias As String, propertyName As String) As Variant

 
 
 '' https://www.dalesandro.net/retrieve-outlook-address-book-data-using-custom-excel-vba-function/
 
 
  On Error GoTo errorHandler
  Dim olApp As Outlook.Application
  Dim olNameSpace As NameSpace
  Dim olRecipient As Outlook.Recipient
  Dim olExchUser As Outlook.ExchangeUser
  Dim olContact As Outlook.AddressEntry
  Set olApp = CreateObject("Outlook.Application")
  Set olNameSpace = olApp.GetNamespace("MAPI")
  Set olRecipient = olNameSpace.CreateRecipient(LCase(Trim(alias)))
  olRecipient.Resolve
  If olRecipient.Resolved Then
    Set olExchUser = olRecipient.AddressEntry.GetExchangeUser
    If Not olExchUser Is Nothing Then
      'Attempt to extract information from Exchange
      GetOutlookAddressBookProperty = Switch(propertyName = "Job Title", olExchUser.JobTitle, _
                                             propertyName = "Company Name", olExchUser.CompanyName, _
                                             propertyName = "Department", olExchUser.Department, _
                                             propertyName = "Name", olExchUser.Name, _
                                             propertyName = "First Name", olExchUser.FirstName, _
                                             propertyName = "City", olExchUser.City, _
                                             propertyName = "Last Name", olExchUser.LastName)
    Else
      'If Exchange not available, then attempt to extract information from local Contacts
      Set olContact = olRecipient.AddressEntry
      
      If Not olContact Is Nothing Then
        GetOutlookAddressBookProperty = Switch(propertyName = "Job Title", olContact.GetContact.JobTitle, _
                                               propertyName = "Company Name", olContact.GetContact.CompanyName, _
                                               propertyName = "Department", olContact.GetContact.Department, _
                                               propertyName = "Name", olContact.GetContact.FullName, _
                                               propertyName = "First Name", olContact.GetContact.FirstName, _
                                               propertyName = "City", olContact.GetContact.BusinessAddressCity, _
                                               propertyName = "Last Name", olContact.GetContact.LastName)
      Else
        GetOutlookAddressBookProperty = CVErr(xlErrNA)
      End If
    End If
  Else
    GetOutlookAddressBookProperty = CVErr(xlErrNA)
  End If
errorHandler:
  If Err.Number <> 0 Then
    GetOutlookAddressBookProperty = CVErr(xlErrNA)
  End If
End Function

Sub corrigeerToCC(newmail As MailItem)

Dim zelf As String
Dim tos() As String

    On Error Resume Next
    
    '' Stop eigen naam (Facilitair) in zelf
    zelf = newmail.Sender.Name
    
    '' stop alle ontvangers uit To veld in array
    tos = Split(newmail.To, ";")
        
    '' Loop alle ontvangers door en kijk of het de eerste melder is en anders naar CC schoppen
    For Each Recipient In newmail.Recipients
      If Recipient.Name <> tos(0) Then Recipient.Type = olCC
    Next Recipient
    
    '' Loop ze nog allemaal eens door en verwijder de ontvanger als we het zelf zijn (lukt niet om in 1 loop te doen blijkbaar)
    For Each Recipient In newmail.Recipients
      If Recipient.Name = zelf Then Recipient.Delete
    Next Recipient
    
    '' Even netjes alles checken
    newmail.Recipients.ResolveAll

End Sub



Sub TESTplakTextInBody(myHTMLText As String, newmail As MailItem, oInspector As Inspector)
    '' Check op welke manier de reply is opengezet en probeer de nieuwe body text erin te proppen
    ''If oInspector Is Nothing Then
        Debug.Print "tel = " & tel & "  plakTextInBody() newmail.BodyFormat = " & newmail.BodyFormat

        Select Case newmail.BodyFormat
            Case olFormatPlain, olFormatRichText, olFormatUnspecified
                newmail.Body = RemoveHTML(myHTMLText) & newmail.Body
            Case olFormatHTML
                newmail.HTMLBody = myHTMLText & newmail.HTMLBody
        End Select

    ''Else
    ''    Debug.Print "tel = " & tel & "Target is niet inline"
    ''End If
End Sub



Sub plakTextInBody(myHTMLText As String, newmail As MailItem, oInspector As Inspector)
    '' Check op welke manier de reply is opengezet en probeer de nieuwe body text erin te proppen
    If oInspector Is Nothing Then
        Debug.Print "plakTextInBody 1  oInspector Is Nothing"

        Select Case newmail.BodyFormat
            Case olFormatPlain, olFormatRichText, olFormatUnspecified
                newmail.Body = RemoveHTML(myHTMLText) & newmail.Body
            Case olFormatHTML
                newmail.HTMLBody = myHTMLText & newmail.HTMLBody
        End Select

    Else
        If oInspector.IsWordMail Then
        
        
        Debug.Print "plakTextInBody 2  oInspector.IsWordMail"
        Debug.Print " *** END ALL"
        MsgBox stopTEXT, vbCritical, "Foutje gevonden"
        End
        
        
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
        Debug.Print "plakTextInBody 3   not oInspector.IsWordMail"
        MsgBox "Dit is experimenteel. Beter is gewoon niet een los window openen om deze macro te gebruiken"
            ' No object model to work with. Must manipulate raw text.
            Select Case newmail.BodyFormat
                Case olFormatPlain, olFormatRichText, olFormatUnspecified
                    newmail.Body = newmail.Body & RemoveHTML(myHTMLText)
                Case olFormatHTML
                    newmail.HTMLBody = newmail.HTMLBody & "<p>" & myHTMLText & "</p>"
            End Select
        End If
    End If
End Sub

Function grijpMelder(newmail As MailItem) As String

    Dim txt            As String
    Dim onderwerpPos1  As String
    Dim onderwerpPos2  As String
    Dim groetLen       As Integer
    Dim groetenPos     As Integer
    Dim gokNaam        As String
    Dim melder         As String
    Dim melders()      As String
    Dim melderVoornaam As Integer
    Dim senderHadGetal As Boolean
    Dim defaultAanhef  As String
    
    '' Als we geen goede naam vinden dan deze aanhef gebruiken
    defaultAanhef = "melder"
    
    '' eerst alles resolven
    newmail.Recipients.ResolveAll
    '' alle ontvangers uitsplitsen
    melders = Split(newmail.To, ";")
    '' controle of er iets in het To veld staat dan eerste ontvanger selecteren uit rij
    If newmail.To = "" Then melder = "123" Else melder = melders(0)
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
    '' voornaam kan bestaan uit voornaam plus tussenvoegsels, hier de meest voorkomende tussenvoegsels vervangen door niks
    For Each voegsel In Array("van het", "van der", "van den", "van de", " aan", " bij", " in", " onder", " van", " den", " ten", " 't", " het", " de")
        melder = Replace(melder, voegsel, "")
    Next
    '' eventuele spaties eraf trimmen
    melder = Trim(melder)
    '' alvast gevonden voornaam optie in variabel zetten
    grijpMelder = melder
     
    '' Hier een aantal ontdekte woorden die als voornaam worden aangezien maar niet kloppen
    For Each miswoord In Array("Groep", "De", "Het", "T", "'T", "Middelburg")
        If LCase(grijpMelder) = LCase(miswoord) Then melderVoornaam = 0
    Next
     
    '' Kijk of in gevonden meldernaam een getal zit of dat er geen sprake was van een komma-gescheiden meldernaam
    If (senderHadGetal Or melderVoornaam = 0) Then
        grijpMelder = defaultAanhef
        txt = newmail.Body
        onderwerpPos1 = InStr(1, txt, "Onderwerp: ")
        onderwerpPos2 = InStr(onderwerpPos1 + 10, txt, "Onderwerp: ")
        If onderwerpPos2 = 0 Then onderwerpPos2 = Len(txt)
        
        txt = LCase(Mid(txt, onderwerpPos1 + 11, onderwerpPos2 - onderwerpPos1 + 1))
             
        For Each groet In Array("voorbaat dank,", "melder:", "mvg,", "mvgr", "mvrgr", "groeten van", "groetjes van", "gr.van", "groetjes,", "groet:", "groeten:", "groetjes:", "groet van", "groet;", "groeten;", "groetjes;", "groetjes", "groeten", "groet", "dank!", "gr.", "mvg", "m.v.g.", "thanks,", "gr ", "gr" & Chr(13), " gr ", "groet,", "groeten,")
            groetLen = Len(groet)
            groetenPos = InStr(txt, groet)
            If groetenPos > 0 Then
                gokNaam = Mid(txt, groetenPos + groetLen, 25)
                '' Hier wordt wat gekunsteld om eventuele gekke handtekeningen uit te filteren zodat alsnog de naam naar voren komt, werkt in sommige gevallen prima
                gokNaam = Replace(gokNaam, Chr(13), " ")
                gokNaam = Replace(gokNaam, Chr(10), " ")
                gokNaam = Replace(gokNaam, Chr(12), " ")
                gokNaam = smartTrim(gokNaam)
                gokNaam = Trim(gokNaam)
                grijpMelder = mooi(gokNaam)
                Exit For
            End If
        Next
          
    End If
    
    '' Aantal misgrijpingen verbeteren
    For Each miswoord In Array("Team", "De", "Het", "Namens", "Jac", "DB", "Leer")
        If LCase(grijpMelder) = LCase(miswoord) Then grijpMelder = defaultAanhef
    Next
    
End Function



Function grijpVerzonden(newmail As MailItem) As String

    Dim txt As String
    Dim Pos1 As Integer
    Dim Pos2 As Integer
    Dim pakDatum As String
    

    txt = newmail.Body
    Pos1 = InStr(1, txt, "Verzonden: ") + 11
    Pos2 = InStr(1, txt, "Aan: ")
    If Pos2 = 0 Then Pos2 = Pos1 + 25
           
    pakDatum = Mid(txt, Pos1, Pos2 - Pos1)
    
    pakDatum = Replace(pakDatum, "maandag", "ma")
    pakDatum = Replace(pakDatum, "dinsdag", "di")
    pakDatum = Replace(pakDatum, "woensdag", "wo")
    pakDatum = Replace(pakDatum, "donderdag", "do")
    pakDatum = Replace(pakDatum, "vrijdag", "vr")
    pakDatum = Replace(pakDatum, "zaterdag", "za")
    pakDatum = Replace(pakDatum, "zondag", "zo")
    pakDatum = Replace(pakDatum, Format(Now(), "yyyy"), "")
    
    grijpVerzonden = pakDatum

End Function



Function GetCurrentItem() As Object
    Dim objApp As Outlook.Application

    Set objApp = Application

    On Error Resume Next
    Debug.Print "tel = " & tel & "  GetCurrentItem() typename = " & TypeName(objApp.ActiveWindow)
    
    Select Case TypeName(objApp.ActiveWindow)
        Case "Explorer"
            ' Set GetCurrentItem = objApp.ActiveExplorer.Selection.Item(1)
            Set GetCurrentItem = objApp.ActiveExplorer.ActiveInlineResponse
            Debug.Print "tel = " & tel & "  GetCurrentItem()  = objApp.ActiveExplorer.ActiveInlineResponse"
            
        Case "Inspector"
            Set GetCurrentItem = objApp.ActiveInspector.CurrentItem
            Debug.Print "tel = " & tel & "  GetCurrentItem()  = objApp.ActiveInspector.CurrentItem"
            
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
    
    '' Vond deze geniale oplossing op stackoverflow
    ''gevondenAantal = Len(txt) - Len(Replace(txt, vind, ""))

End Function



Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean

    ''IsInArray = (UBound(Filter(arr, stringToBeFound)) > -1)
    
    
    Dim i
    For i = LBound(arr) To UBound(arr)
        If LCase(arr(i)) = LCase(stringToBeFound) Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False
  
End Function



Function blauw(txt As String) As String
    blauw = "<span style=" & Chr(34) & "color:blue" & Chr(34) & ">" & txt & "</span>"
End Function

Function zwart(txt As String) As String
    zwart = "<span style=" & Chr(34) & "color:black" & Chr(34) & ">" & txt & "</span>"
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
