Module modMain
    Public entete As String

    Sub Main()
        Dim fichierINI As String, fichierECO As String, fichierEXP As String, moteurEXP As String
        Dim lecture As IO.FileStream, tabTampon() As Byte, pas As Integer, pos As Long
        Dim chaine As String, tabChaine() As String, tabTmp() As String
        Dim tabProf(256) As Integer, prof As Integer, totProf As Long, nbCoups As Integer, minProf As Integer, maxProf As Integer
        Dim tabECO(0) As String, i As Integer, j As Integer, compteur As Integer
        Dim depart As Integer, moteur_court As String
        
                fichierEXP = Replace(Command(), """", "")
        If fichierEXP = "" Or Not My.Computer.FileSystem.FileExists(fichierEXP) Then
            End
        End If

        fichierINI = My.Computer.Name & ".ini"
        fichierECO = "ouvertures.txt"
        moteurEXP = "Eman.exe"

        If My.Computer.FileSystem.FileExists(fichierINI) Then
            chaine = My.Computer.FileSystem.ReadAllText(fichierINI)
            If chaine <> "" And InStr(chaine, vbCrLf) > 0 Then
                tabChaine = Split(chaine, vbCrLf)
                For i = 0 To UBound(tabChaine)
                    If tabChaine(i) <> "" And InStr(tabChaine(i), " = ") > 0 Then
                        tabTmp = Split(tabChaine(i), " = ")
                        If tabTmp(0) <> "" And tabTmp(1) <> "" Then
                            If InStr(tabTmp(1), "//") > 0 Then
                                tabTmp(1) = Trim(gauche(tabTmp(1), tabTmp(1).IndexOf("//") - 1))
                            End If
                            Select Case tabTmp(0)
                                Case "moteurEXP"
                                    moteurEXP = Replace(tabTmp(1), """", "")
                                Case "fichierECO"
                                    fichierECO = Replace(tabTmp(1), """", "")
                                Case Else

                            End Select
                        End If
                    End If
                Next
            End If
        End If
        My.Computer.FileSystem.WriteAllText(fichierINI, "moteurEXP = " & moteurEXP & vbCrLf _
                                                      & "fichierECO = " & fichierECO & vbCrLf, False)
        moteur_court = nomFichier(moteurEXP)

        'ETAPE 1/2 : PROFONDEURS

        lecture = New IO.FileStream(fichierEXP, IO.FileMode.Open)

        'version 1 ou 2 ?
        pas = 32
        ReDim tabTampon(25)
        lecture.Read(tabTampon, 0, 26)
        If System.Text.Encoding.UTF8.GetString(tabTampon) = "SugaR Experience version 2" Then
            pas = 24
            'on part de là on avance de 16 octets
            lecture.Position = lecture.Position + 16

            'fichierEXP v2 et eman v6 ?
            If InStr(moteur_court, "eman 6", CompareMethod.Text) > 0 Then
                MsgBox("Please, select an another engine (eman v7.00/v8.00, hypnos, stockfishmz, aurora) !", MsgBoxStyle.Exclamation)
                lecture.Close()
                End
            End If
        Else
            'on revient au départ et on avance de 12 octets
            lecture.Position = 12

            'fichierEXP v1 et eman v7 ?
            If InStr(moteur_court, "eman 6", CompareMethod.Text) = 0 Then
                MsgBox("Please, select an another engine (eman v6.00) !", MsgBoxStyle.Exclamation)
                lecture.Close()
                End
            End If
        End If
        tabTampon = Nothing

        totProf = 0
        nbCoups = 0
        minProf = 1000
        maxProf = 0
        depart = Environment.TickCount
        For pos = lecture.Position To lecture.Length Step pas
            lecture.Position = pos
            prof = lecture.ReadByte()

            totProf = totProf + prof
            nbCoups = nbCoups + 1

            If prof < minProf Then
                minProf = prof
            ElseIf prof > maxProf Then
                maxProf = prof
            End If

            Select Case prof
                Case Is >= 100
                    tabProf(100) = tabProf(100) + 1

                Case Is >= 90
                    tabProf(90) = tabProf(90) + 1

                Case Is >= 80
                    tabProf(80) = tabProf(80) + 1

                Case Is >= 70
                    tabProf(70) = tabProf(70) + 1

                Case Is >= 60
                    tabProf(60) = tabProf(60) + 1

                Case Is >= 50
                    tabProf(50) = tabProf(50) + 1

                Case Is >= 40
                    tabProf(40) = tabProf(40) + 1

                Case Is >= 30
                    tabProf(30) = tabProf(30) + 1

                Case Is >= 20
                    tabProf(20) = tabProf(20) + 1

                Case Is >= 10
                    tabProf(10) = tabProf(10) + 1

                Case Else
                    tabProf(0) = tabProf(0) + 1
            End Select

            If nbCoups Mod 500000 = 0 Then
                Console.Clear()
                Console.Title = My.Computer.Name & " : " & nomFichier(fichierEXP) & " @ " & Format(lecture.Position / lecture.Length, "0.00%") & " (" & heureFin(depart, lecture.Position, lecture.Length, , True) & ")"

                Console.WriteLine("Moves : " & Trim(Format(nbCoups, "# ### ### ##0")) & vbCrLf)

                Console.WriteLine("Stats : min D" & minProf & " < avg D" & Format(totProf / nbCoups, "0") & " < max D" & maxProf & vbCrLf)

                Console.WriteLine("Depth :  < D10 | >= D10 | >= D20 | >= D30 | >= D40 | >= D50 | >= D60 | >= D70 | >= D80 | >= D90 | >=D100")
                Console.WriteLine("..... : " & Format(tabProf(0) / nbCoups, "00.00%") & " | " & Format(tabProf(10) / nbCoups, "00.00%") & " | " & Format(tabProf(20) / nbCoups, "00.00%") & " | " & Format(tabProf(30) / nbCoups, "00.00%") & " | " & Format(tabProf(40) / nbCoups, "00.00%") & " | " & Format(tabProf(50) / nbCoups, "00.00%") & " | " & Format(tabProf(60) / nbCoups, "00.00%") & " | " & Format(tabProf(70) / nbCoups, "00.00%") & " | " & Format(tabProf(8) / nbCoups, "00.00%") & " | " & Format(tabProf(90) / nbCoups, "00.00%") & " | " & Format(tabProf(100) / nbCoups, "00.00%") & vbCrLf)
            End If
        Next
        Console.Clear()
        Console.Title = My.Computer.Name & " : " & nomFichier(fichierEXP) & " @ " & Format(lecture.Position / lecture.Length, "0.00%")

        Console.WriteLine("Moves : " & Trim(Format(nbCoups, "# ### ### ##0")) & vbCrLf)

        Console.WriteLine("Stats : min D" & minProf & " < avg D" & Format(totProf / nbCoups, "0") & " < max D" & maxProf & vbCrLf)

        Console.WriteLine("Depth :  < D10 | >= D10 | >= D20 | >= D30 | >= D40 | >= D50 | >= D60 | >= D70 | >= D80 | >= D90 | >=D100")
        Console.WriteLine("..... : " & Format(tabProf(0) / nbCoups, "00.00%") & " | " & Format(tabProf(10) / nbCoups, "00.00%") & " | " & Format(tabProf(20) / nbCoups, "00.00%") & " | " & Format(tabProf(30) / nbCoups, "00.00%") & " | " & Format(tabProf(40) / nbCoups, "00.00%") & " | " & Format(tabProf(50) / nbCoups, "00.00%") & " | " & Format(tabProf(60) / nbCoups, "00.00%") & " | " & Format(tabProf(70) / nbCoups, "00.00%") & " | " & Format(tabProf(8) / nbCoups, "00.00%") & " | " & Format(tabProf(90) / nbCoups, "00.00%") & " | " & Format(tabProf(100) / nbCoups, "00.00%") & vbCrLf)

        lecture.Close()

        'ETAPE 2/2 : OUVERTURES

        If My.Computer.FileSystem.FileExists(fichierECO) Then
            Console.Write("Loading " & moteur_court & "... ")
            chargerMoteur(moteurEXP, fichierEXP)
            Console.WriteLine("OK" & vbCrLf)

            'entete = ""
            'Console.WriteLine("Defragging " & nomFichier(fichierEXP) & " :")
            'If InStr(moteur_court, "eman 6", CompareMethod.Text) > 0 And InStr(entete, "Collisions: 0", CompareMethod.Text) = 0 Then
            'entete = defragEXP(fichierEXP, 1, True, entree, sortie)
            'ElseIf InStr(moteur_court, "eman 7", CompareMethod.Text) > 0 And InStr(entete, "Duplicate moves: 0", CompareMethod.Text) = 0 Then
            'entete = defragEXP(fichierEXP, 4, True, entree, sortie)
            'End If
            Console.WriteLine(entete)

            chaine = My.Computer.FileSystem.ReadAllText(fichierECO)
            If chaine <> "" And InStr(chaine, vbCrLf) > 0 Then
                tabECO = Split(chaine, vbCrLf)
            End If

            For i = 0 To UBound(tabECO)
                If tabECO(i) <> "" Then
                    tabTmp = Split(tabECO(i), ":")
                    compteur = 0
                    While tabTmp(1) <> ""
                        If InStr(tabTmp(1), "|") = 0 Then
                            chaine = expListe(Trim(tabTmp(1)))
                        Else
                            chaine = expListe(Trim(tabTmp(1).Substring(0, tabTmp(1).IndexOf("|"))))
                        End If
                        If chaine <> "" Then
                            tabChaine = Split(chaine, vbCrLf)
                            For j = 0 To UBound(tabChaine)
                                If tabChaine(j) <> "" Then
                                    chaine = tabChaine(j).Substring(tabChaine(j).IndexOf("count:") + 7)
                                    chaine = chaine.Substring(0, chaine.IndexOf(","))
                                    compteur = compteur + CInt(chaine)
                                End If
                            Next
                        End If
                        If InStr(tabTmp(1), "|") = 0 Then
                            tabTmp(1) = ""
                        Else
                            tabTmp(1) = tabTmp(1).Substring(tabTmp(1).IndexOf("|") + 1)
                        End If
                    End While
                    If compteur = 0 Then
                        Console.WriteLine(tabTmp(0) & " : never played")
                    Else
                        Console.WriteLine(tabTmp(0) & " : played " & Format(compteur, "0 000") & " times")
                    End If
                End If
            Next
            Console.WriteLine("")

            dechargerMoteur()
        End If

        Console.WriteLine("Press ENTER to close the window.")
        Console.ReadLine()

    End Sub

End Module
