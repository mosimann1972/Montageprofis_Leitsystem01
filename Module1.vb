
Imports System.Data
Imports System.Console
Imports System.Data.SqlClient
Imports System.Data.Common
Imports Pervasive.Data.SqlClient

Imports System
Imports System.IO
Imports System.Collections

Module Module1

    Public SQLSERVER As String
    Public PSQLSERVER As String

    Public MitarbeiterNummer As String
    Public MitarbeiterName As String
    Public MitarbeiterCode As String
    Public MitarbeiterEintritt As String
    Public MitarbeiterAustritt As String
    Public MitarbeiterStandort As String
    Public MitarbeiterZuteilung As String
    Public MitarbeiterZuteilungCode As String

    Public MitarbeiterAdresse As String
    Public MitarbeiterPLZ As String
    Public MitarbeiterOrt As String
    Public MitarbeiterTelefon1 As String
    Public MitarbeiterTelefon2 As String
    Public MitarbeiterMobilnummer As String
    Public MitarbeiterBild As String


	Public FahrzeugADRESSNUMMER As String
	Public FahrzeugFAHRZEUGNUMMER As String
	Public FahrzeugMARKE As String
	Public FahrzeugMODELL As String
	Public FahrzeugANZAHLPLAETZE As String
	Public FahrzeugKILOMETERSTAND As String
	Public FahrzeugFARBE As String
	Public FahrzeugPOLIZEIKENNZEICHEN As String
	Public FahrzeugBILD As String
	Public FahrzeugSUCHSCHLUESSEL As String
	Public FahrzeugPARKPLATZNR As String
	Public FahrzeugZUGETMONTEUR As String
	Public FahrzeugDOKUMENT As String
	Public FahrzeugBENZINKARTENNR As String
	Public FahrzeugTREIBSTOFFART As String


    Public EinsatzAdressnummer As String
    Public EinsatzAuftragsnummer As String
    Public EinsatzTagesDatum As String

    Public EinsatzSTATUSMACHEFMON As String
    Public EinsatzABREISEZEIT As String
    Public EinsatzBAUSTELLENZEIT As String
    Public EinsatzENDZEIT As String
    Public EinsatzFAHRZEUGNUMMER As String
    Public EinsatzPARKPLATZ As String
    Public EinsatzHOTELADRESSNR As String
    Public EinsatzHOTELDETAIL As String
    Public EinsatzHOTELBESTELLT As String
    Public EinsatzBESTELLNUMMER As String


    Public TerminAdressnummer As String
    Public TerminDatum As String
    Public TerminArt As String
    Public TerminKW As String
    Public TerminMontagFlag As Boolean

    Public AuftragNummer As String
    Public AuftragZuteilung As String
    Public AuftragZuteilungCode As String
    Public AuftragName As String

    Public AuftragOBJEKTBESCHREIB1 As String
    Public AuftragOBJEKTBESCHREIB2 As String
    Public AuftragTERMIN1VON As String
    Public AuftragTERMIN1BIS As String
    Public AuftragANZAHLMA1 As String
    Public AuftragANKUNFTSZEIT1 As String
    Public AuftragTERMIN2VON As String
    Public AuftragTERMIN2BIS As String
    Public AuftragANZAHLMA2 As String
    Public AuftragANKUNFTSZEIT2 As String
    Public AuftragTERMIN3VON As String
    Public AuftragTERMIN3BIS As String
    Public AuftragANZAHLMA3 As String
    Public AuftragANKUNFTSZEIT3 As String
    Public AuftragTERMIN4VON As String
    Public AuftragTEMIN4BIS As String
    Public AuftragANZAHLMA4 As String
    Public AuftragANKUNFTSZEIT4 As String
    Public AuftragTERMIN5VON As String
    Public AuftragTERMIN5BIS As String
    Public AuftragANZAHLMA5 As String
    Public AuftragANKUNFTSZEIT5 As String
    Public AuftragFARBCODEAUFTRAG As String
    Public AuftragURSPRUNGSMANDANT As String
    Public AuftragSATERMIN1 As String
    Public AuftragSOTERMIN1 As String
    Public AuftragTERMIN1FEIERTAG1 As String
    Public AuftragSATERMIN2 As String
    Public AuftragSOTERMIN2 As String
    Public AuftragTERMIN2FEIERTAG2 As String
    Public AuftragSATERMIN3 As String
    Public AuftragSOTERMIN3 As String
    Public AuftragTERMIN3FEIERTAG3 As String
    Public AuftragSATERMIN4 As String
    Public AuftragSOTERMIN4 As String
    Public AuftragTERMIN4FEIERTAG4 As String
    Public AuftragSATERMIN5 As String
    Public AuftragSOTERMIN5 As String
    Public AuftragTERMIN5FEIERTAG5 As String
    Public AuftragMONTAGEART As String
    Public AuftragAUFTRAGSBETREUER As String
    Public AuftragBEMERKUNGEN As String
    Public AuftragUMRECHNUNGSFAKT As String
    Public AuftragAUFTRAGSEMAIL As String
    Public AuftragAUFTRAGSFAX As String
    Public AuftragAUFTRAGBESTAETIG As String
    Public AuftragCHMONTEUREMAIL As String
    Public AuftragCHMONTEURFAX As String

    Public RueckgabeDatum As DatumWert

    Public RDat1 As DatumWert
    Public RDat2 As DatumWert


    Sub Main()

		Readtxt()

		Console.WriteLine("START   LESEN Auftrag       --------------------------------------------------")
        AuftragLesen()
        Console.WriteLine("STOP    LESEN Auftrag       --------------------------------------------------")

		MitarbeiterLoeschenAlle()

        RueckgabeDatum = KalenderLesen(0)
        Console.WriteLine("0START   LESEN MITARBEITER   --------------------------------------------------")
        MitarbeiterLesenNeu()
        Console.WriteLine("0ENDE    LESEN MITARBEITER   --------------------------------------------------")

        RueckgabeDatum = KalenderLesen(1)
        Console.WriteLine("1START   LESEN MITARBEITER   --------------------------------------------------")
        MitarbeiterLesenNeu()
        Console.WriteLine("1ENDE    LESEN MITARBEITER   --------------------------------------------------")

        RueckgabeDatum = KalenderLesen(2)
        Console.WriteLine("2START   LESEN MITARBEITER   --------------------------------------------------")
        MitarbeiterLesenNeu()
        Console.WriteLine("2ENDE    LESEN MITARBEITER   --------------------------------------------------")

        RueckgabeDatum = KalenderLesen(3)
        Console.WriteLine("3START   LESEN MITARBEITER   --------------------------------------------------")
        MitarbeiterLesenNeu()
        Console.WriteLine("3ENDE    LESEN MITARBEITER   --------------------------------------------------")



        'RueckgabeDatum = KalenderLesen(0)
        'Console.WriteLine("START   LESEN MITARBEITER   --------------------------------------------------")
        'MitarbeiterLesen()
        'Console.WriteLine("ENDE    LESEN MITARBEITER   --------------------------------------------------")

        'RueckgabeDatum = KalenderLesen(1)
        'Console.WriteLine("START   LESEN MITARBEITER   --------------------------------------------------")
        'MitarbeiterLesen()
        'Console.WriteLine("ENDE    LESEN MITARBEITER   --------------------------------------------------")


        'RueckgabeDatum = KalenderLesen(2)
        'Console.WriteLine("START   LESEN MITARBEITER   --------------------------------------------------")
        'MitarbeiterLesen()
        'Console.WriteLine("ENDE    LESEN MITARBEITER   --------------------------------------------------")

        'RueckgabeDatum = KalenderLesen(3)
        'Console.WriteLine("START   LESEN MITARBEITER   --------------------------------------------------")
        'MitarbeiterLesen()
        'Console.WriteLine("ENDE    LESEN MITARBEITER   --------------------------------------------------")


        RueckgabeDatum = KalenderLesen(0)
        Console.WriteLine("START 0   LESEN TERMIN      --------------------------------------------------")
        TerminLesen()
        Console.WriteLine("ENDE  0   LESEN TERMIN      --------------------------------------------------")

        Console.WriteLine("START 0   LESEN EINSATZ     --------------------------------------------------")
        EinsatzLesen()
        Console.WriteLine("ENDE  0   LESEN EINSATZ     --------------------------------------------------")

        RueckgabeDatum = KalenderLesen(1)
        Console.WriteLine("START 1   LESEN TERMIN      --------------------------------------------------")
        TerminLesen()
        Console.WriteLine("ENDE  1   LESEN TERMIN      --------------------------------------------------")

        Console.WriteLine("START 1   LESEN EINSATZ     --------------------------------------------------")
        EinsatzLesen()
        Console.WriteLine("ENDE  1   LESEN EINSATZ     --------------------------------------------------")

        RueckgabeDatum = KalenderLesen(2)
        Console.WriteLine("START 2   LESEN TERMIN      --------------------------------------------------")
        TerminLesen()
        Console.WriteLine("ENDE  2   LESEN TERMIN      --------------------------------------------------")

        Console.WriteLine("START 2   LESEN EINSATZ     --------------------------------------------------")
        EinsatzLesen()
        Console.WriteLine("ENDE  2   LESEN EINSATZ     --------------------------------------------------")

        RueckgabeDatum = KalenderLesen(3)
        Console.WriteLine("START 3   LESEN TERMIN      --------------------------------------------------")
        TerminLesen()
        Console.WriteLine("ENDE  3   LESEN TERMIN      --------------------------------------------------")

        Console.WriteLine("START 3   LESEN EINSATZ     --------------------------------------------------")
        EinsatzLesen()
        Console.WriteLine("ENDE  3   LESEN EINSATZ     --------------------------------------------------")


        'Es wird jeweils das Montagsdatum der aktuellen Woche ermittelt.
		Console.WriteLine("Start LESEN EINSATZ WP      --------------------------------------------------")
        EinsatzLesenWP()
		Console.WriteLine("Ende  LESEN EINSATZ WP      --------------------------------------------------")


		Console.WriteLine("Start  LESEN Termin WP      --------------------------------------------------")
		TerminLesenWP()
		Console.WriteLine("Ende   LESEN Termin WP      --------------------------------------------------")

		Console.WriteLine("Start  Fahrzeug WP      --------------------------------------------------")
		FahrzeugLesen()
		Console.WriteLine("Ende   Fahrzeug WP      --------------------------------------------------")




    End Sub


    Private Function Readtxt()

        Dim objReader As New StreamReader("c:\Leitsystem\sql.txt")
        Dim sLine As String = ""
        Dim arrText As New ArrayList()

        Do
            sLine = objReader.ReadLine()
            If Not sLine Is Nothing Then
                arrText.Add(sLine)
            End If
        Loop Until sLine Is Nothing
        objReader.Close()

        SQLSERVER = arrText(0)
        PSQLSERVER = arrText(1)

        'For Each sLine In arrText
        'Console.WriteLine(sLine)
        'Next
        'Console.ReadLine()


    End Function

    Private Sub AuftragLesen()

        Dim pCon As PsqlConnection
        Dim pCmd As PsqlCommand
        Dim pReader As PsqlDataReader

        pCon = New PsqlConnection(PSQLSERVER)

        AufragLoeschen()

        Try
            pCon.Open()
            pCmd = New PsqlCommand("select * from T229Auftragsdetail", pCon)

            pReader = pCmd.ExecuteReader()

            While pReader.Read()

                AuftragNummer = pReader(2).ToString()
                AuftragName = Replace(pReader(3).ToString(), "'", "''")

				'Neue Nummerierung wegen Roland Update europa, 09.05.2016
				'AuftragZuteilung = Trim(pReader(47).ToString())
				AuftragZuteilung = Trim(pReader(50).ToString())


				'Neue Nummerierung wegen Roland Update europa, 09.05.2016
				'AuftragZuteilungCode = UCase(Left(Trim(pReader(47).ToString()), 3))
				AuftragZuteilungCode = UCase(Left(Trim(pReader(50).ToString()), 3))

                If AuftragZuteilungCode = "TÜR" Or AuftragZuteilungCode = "BET" Then
                    AuftragZuteilungCode = "T-B"
                End If


                'Gemäss Alli (15.07.2015) gehört Labor zu Innenausbau
                If AuftragZuteilungCode = "LAB" Then
                    AuftragZuteilungCode = "INN"
                End If


                AuftragOBJEKTBESCHREIB1 = Replace(Trim(pReader(4).ToString()), "'", "''")
                AuftragOBJEKTBESCHREIB2 = Replace(Trim(pReader(5).ToString()), "'", "''")


				'Neue Nummerierung wegen Roland Update europa, 09.05.2016
				'AuftragTERMIN1VON = Trim(pReader(6).ToString())
				'AuftragTERMIN1BIS = Trim(pReader(7).ToString())
				'AuftragANZAHLMA1 = Trim(pReader(8).ToString())
				'AuftragANZAHLMA1 = Trim(pReader(11).ToString())

				'AuftragANKUNFTSZEIT1 = Trim(pReader(9).ToString())

				'AuftragTERMIN2VON = Trim(pReader(10).ToString())
				'AuftragTERMIN2BIS = Trim(pReader(11).ToString())
				'AuftragANZAHLMA2 = Trim(pReader(12).ToString())
				'AuftragANKUNFTSZEIT2 = Trim(pReader(13).ToString())

				'AuftragTERMIN3VON = Trim(pReader(14).ToString())
				'AuftragTERMIN3BIS = Trim(pReader(15).ToString())
				'AuftragANZAHLMA3 = Trim(pReader(16).ToString())
				'AuftragANKUNFTSZEIT3 = Trim(pReader(17).ToString())

				'AuftragTERMIN4VON = Trim(pReader(18).ToString())
				'AuftragTEMIN4BIS = Trim(pReader(19).ToString())
				'AuftragANZAHLMA4 = Trim(pReader(20).ToString())
				'AuftragANKUNFTSZEIT4 = Trim(pReader(21).ToString())

				'AuftragTERMIN5VON = Trim(pReader(43).ToString())
				'AuftragTERMIN5BIS = Trim(pReader(44).ToString())
				'AuftragANZAHLMA5 = Trim(pReader(45).ToString())
				'AuftragANKUNFTSZEIT5 = Trim(pReader(46).ToString())

				'AuftragFARBCODEAUFTRAG = Trim(pReader(25).ToString())
				'AuftragURSPRUNGSMANDANT = Trim(pReader(26).ToString())

				'AuftragSATERMIN1 = Trim(pReader(29).ToString())
				'AuftragSOTERMIN1 = Trim(pReader(30).ToString())
				'AuftragTERMIN1FEIERTAG1 = Trim(pReader(31).ToString())

				'AuftragSATERMIN2 = Trim(pReader(32).ToString())
				'AuftragSOTERMIN2 = Trim(pReader(33).ToString())
				'AuftragTERMIN2FEIERTAG2 = Trim(pReader(34).ToString())

				'AuftragSATERMIN3 = Trim(pReader(35).ToString())
				'AuftragSOTERMIN3 = Trim(pReader(36).ToString())
				'AuftragTERMIN3FEIERTAG3 = Trim(pReader(37).ToString())

				'AuftragSATERMIN4 = Trim(pReader(38).ToString())
				'AuftragSOTERMIN4 = Trim(pReader(39).ToString())
				'AuftragTERMIN4FEIERTAG4 = Trim(pReader(40).ToString())

				'AuftragSATERMIN5 = Trim(pReader(48).ToString())
				'AuftragSOTERMIN5 = Trim(pReader(49).ToString())
				'AuftragTERMIN5FEIERTAG5 = Trim(pReader(50).ToString())


				'AuftragMONTAGEART = Trim(pReader(47).ToString())


				AuftragTERMIN1VON = Trim(pReader(9).ToString())
				AuftragTERMIN1BIS = Trim(pReader(10).ToString())
				AuftragANZAHLMA1 = Trim(pReader(11).ToString())
				AuftragANZAHLMA1 = Trim(pReader(15).ToString())

				AuftragANKUNFTSZEIT1 = Trim(pReader(12).ToString())

				AuftragTERMIN2VON = Trim(pReader(13).ToString())
				AuftragTERMIN2BIS = Trim(pReader(14).ToString())
				AuftragANZAHLMA2 = Trim(pReader(15).ToString())
				AuftragANKUNFTSZEIT2 = Trim(pReader(16).ToString())

				AuftragTERMIN3VON = Trim(pReader(17).ToString())
				AuftragTERMIN3BIS = Trim(pReader(18).ToString())
				AuftragANZAHLMA3 = Trim(pReader(19).ToString())
				AuftragANKUNFTSZEIT3 = Trim(pReader(20).ToString())

				AuftragTERMIN4VON = Trim(pReader(21).ToString())
				AuftragTEMIN4BIS = Trim(pReader(22).ToString())
				AuftragANZAHLMA4 = Trim(pReader(23).ToString())
				AuftragANKUNFTSZEIT4 = Trim(pReader(24).ToString())

				AuftragTERMIN5VON = Trim(pReader(46).ToString())
				AuftragTERMIN5BIS = Trim(pReader(47).ToString())
				AuftragANZAHLMA5 = Trim(pReader(48).ToString())
				AuftragANKUNFTSZEIT5 = Trim(pReader(49).ToString())


				AuftragFARBCODEAUFTRAG = Trim(pReader(28).ToString())
				AuftragURSPRUNGSMANDANT = Trim(pReader(29).ToString())

				AuftragSATERMIN1 = Trim(pReader(32).ToString())
				AuftragSOTERMIN1 = Trim(pReader(33).ToString())
				AuftragTERMIN1FEIERTAG1 = Trim(pReader(34).ToString())

				AuftragSATERMIN2 = Trim(pReader(35).ToString())
				AuftragSOTERMIN2 = Trim(pReader(36).ToString())
				AuftragTERMIN2FEIERTAG2 = Trim(pReader(37).ToString())

				AuftragSATERMIN3 = Trim(pReader(38).ToString())
				AuftragSOTERMIN3 = Trim(pReader(39).ToString())
				AuftragTERMIN3FEIERTAG3 = Trim(pReader(40).ToString())

				AuftragSATERMIN4 = Trim(pReader(41).ToString())
				AuftragSOTERMIN4 = Trim(pReader(42).ToString())
				AuftragTERMIN4FEIERTAG4 = Trim(pReader(43).ToString())

				AuftragSATERMIN5 = Trim(pReader(51).ToString())
				AuftragSOTERMIN5 = Trim(pReader(52).ToString())
				AuftragTERMIN5FEIERTAG5 = Trim(pReader(53).ToString())


				AuftragMONTAGEART = Trim(pReader(50).ToString())



                'AUFTRAGSBETREUER
                'BEMERKUNGEN
                'UMRECHNUNGSFAKT
                'AUFTRAGSEMAIL
                'AUFTRAGSFAX
                'AUFTRAGBESTAETIG
                'CHMONTEUREMAIL
                'CHMONTEURFAX


                AuftragErfassen()

                reset()

            End While

        Catch ex As PsqlException

            Console.WriteLine("Invalid" & ex.Message)

        Finally
            pCon.Close()
        End Try

    End Sub

Private Sub AufragLoeschen()

    Dim con As New SqlConnection
    Dim cmd As New SqlCommand
    Dim reader As DbDataReader

    con.ConnectionString = SQLSERVER

    Try

        con.Open()
        cmd.Connection = con

        cmd = New SqlCommand("delete from tbAuftrag", con)
        reader = cmd.ExecuteReader

        reader.Close()

    Catch ex As Exception

        Console.WriteLine("Invalid" & ex.Message)

    Finally
        con.Close()
    End Try


End Sub

Private Sub AuftragErfassen()

    Dim con As New SqlConnection
    Dim cmd As New SqlCommand

    Try

        con.ConnectionString = SQLSERVER
        con.Open()
        cmd.Connection = con

            cmd = New SqlCommand("insert into tbAuftrag ([Auftragsnummer]," _
                                & "[Auftragsname]," _
                                & "[Zuteilung]," _
                                & "[ZuteilungCode]," _
                                & "[OBJEKTBESCHREIB1]," _
                                & "[OBJEKTBESCHREIB2]," _
                                & "[TERMIN1VON]," _
                                & "[TERMIN1BIS]," _
                                & "[ANZAHLMA1]," _
                                & "[ANKUNFTSZEIT1]," _
                                & "[TERMIN2VON]," _
                                & "[TERMIN2BIS]," _
                                & "[ANZAHLMA2]," _
                                & "[ANKUNFTSZEIT2]," _
                                & "[TERMIN3VON]," _
                                & "[TERMIN3BIS]," _
                                & "[ANZAHLMA3]," _
                                & "[ANKUNFTSZEIT3]," _
                                & "[TERMIN4VON]," _
                                & "[TEMIN4BIS]," _
                                & "[ANZAHLMA4]," _
                                & "[ANKUNFTSZEIT4]," _
                                & "[TERMIN5VON]," _
                                & "[TERMIN5BIS]," _
                                & "[ANZAHLMA5]," _
                                & "[ANKUNFTSZEIT5]," _
                                & "[FARBCODEAUFTRAG]," _
                                & "[URSPRUNGSMANDANT]," _
                                & "[SATERMIN1]," _
                                & "[SOTERMIN1]," _
                                & "[TERMIN1FEIERTAG1]," _
                                & "[SATERMIN2]," _
                                & "[SOTERMIN2]," _
                                & "[TERMIN2FEIERTAG2]," _
                                & "[SATERMIN3]," _
                                & "[SOTERMIN3]," _
                                & "[TERMIN3FEIERTAG3]," _
                                & "[SATERMIN4]," _
                                & "[SOTERMIN4]," _
                                & "[TERMIN4FEIERTAG4]," _
                                & "[SATERMIN5]," _
                                & "[SOTERMIN5]," _
                                & "[TERMIN5FEIERTAG5]," _
                                & "[MONTAGEART]," _
                                & "[AUFTRAGSBETREUER]," _
                                & "[BEMERKUNGEN]," _
                                & "[UMRECHNUNGSFAKT]," _
                                & "[AUFTRAGSEMAIL]," _
                                & "[AUFTRAGSFAX]," _
                                & "[AUFTRAGBESTAETIG]," _
                                & "[CHMONTEUREMAIL]," _
                                & "[CHMONTEURFAX]) " _
                                & " values ('" & AuftragNummer & "'," _
                                & "'" & AuftragName & "'," _
                                & "'" & AuftragZuteilung & "'," _
                                & "'" & AuftragZuteilungCode & "'," _
                                & "'" & AuftragOBJEKTBESCHREIB1 & "'," _
                                & "'" & AuftragOBJEKTBESCHREIB2 & "'," _
                                & "'" & AuftragTERMIN1VON & "'," _
                                & "'" & AuftragTERMIN1BIS & "'," _
                                & "'" & AuftragANZAHLMA1 & "'," _
                                & "'" & AuftragANKUNFTSZEIT1 & "'," _
                                & "'" & AuftragTERMIN2VON & "'," _
                                & "'" & AuftragTERMIN2BIS & "'," _
                                & "'" & AuftragANZAHLMA2 & "'," _
                                & "'" & AuftragANKUNFTSZEIT2 & "'," _
                                & "'" & AuftragTERMIN3VON & "'," _
                                & "'" & AuftragTERMIN3BIS & "'," _
                                & "'" & AuftragANZAHLMA3 & "'," _
                                & "'" & AuftragANKUNFTSZEIT3 & "'," _
                                & "'" & AuftragTERMIN4VON & "'," _
                                & "'" & AuftragTEMIN4BIS & "'," _
                                & "'" & AuftragANZAHLMA4 & "'," _
                                & "'" & AuftragANKUNFTSZEIT4 & "'," _
                                & "'" & AuftragTERMIN5VON & "'," _
                                & "'" & AuftragTERMIN5BIS & "'," _
                                & "'" & AuftragANZAHLMA5 & "'," _
                                & "'" & AuftragANKUNFTSZEIT5 & "'," _
                                & "'" & AuftragFARBCODEAUFTRAG & "'," _
                                & "'" & AuftragURSPRUNGSMANDANT & "'," _
                                & "'" & AuftragSATERMIN1 & "'," _
                                & "'" & AuftragSOTERMIN1 & "'," _
                                & "'" & AuftragTERMIN1FEIERTAG1 & "'," _
                                & "'" & AuftragSATERMIN2 & "'," _
                                & "'" & AuftragSOTERMIN2 & "'," _
                                & "'" & AuftragTERMIN2FEIERTAG2 & "'," _
                                & "'" & AuftragSATERMIN3 & "'," _
                                & "'" & AuftragSOTERMIN3 & "'," _
                                & "'" & AuftragTERMIN3FEIERTAG3 & "'," _
                                & "'" & AuftragSATERMIN4 & "'," _
                                & "'" & AuftragSOTERMIN4 & "'," _
                                & "'" & AuftragTERMIN4FEIERTAG4 & "'," _
                                & "'" & AuftragSATERMIN5 & "'," _
                                & "'" & AuftragSOTERMIN5 & "'," _
                                & "'" & AuftragTERMIN5FEIERTAG5 & "'," _
                                & "'" & AuftragMONTAGEART & "'," _
                                & "'" & AuftragAUFTRAGSBETREUER & "'," _
                                & "'" & AuftragBEMERKUNGEN & "'," _
                                & "'" & AuftragUMRECHNUNGSFAKT & "'," _
                                & "'" & AuftragAUFTRAGSEMAIL & "'," _
                                & "'" & AuftragAUFTRAGSFAX & "'," _
                                & "'" & AuftragAUFTRAGBESTAETIG & "'," _
                                & "'" & AuftragCHMONTEUREMAIL & "'," _
                                & "'" & AuftragCHMONTEURFAX & "')", con)

        cmd.ExecuteNonQuery()

    Catch ex As Exception
        Console.WriteLine("Invalid" & ex.Message)
    Finally
        con.Close()
    End Try

End Sub


Private Sub MitarbeiterLesenNeu()


    Dim pCon As PsqlConnection
    Dim pCmd As PsqlCommand
    Dim pReader As PsqlDataReader

    pCon = New PsqlConnection(PSQLSERVER)

    Try
        pCon.Open()


		pCmd = New PsqlCommand("select * from T203Mitarbeiter where Faehigkeit = 'TL' or Faehigkeit = 'KO' or Faehigkeit = 'ML' or Faehigkeit = 'M' or Faehigkeit = 'TP' or Faehigkeit = 'FL' Order by MitarbeiterNr", pCon)

		'pCmd = New PsqlCommand("select * from T203Mitarbeiter where Eintritt <= '" & RDat1.D2 & "' and Austritt >= '" & RDat2.D2 & "' and Faehigkeit = 'TL'" _
        '                       & "or Eintritt <= '" & RDat1.D2 & "' and Austritt >= '" & RDat2.D2 & "' and Faehigkeit = 'KO'" _
        '                       & "or Eintritt <= '" & RDat1.D2 & "' and Austritt >= '" & RDat2.D2 & "' and Faehigkeit = 'ML'" _
        '                       & "or Eintritt <= '" & RDat1.D2 & "' and Austritt >= '" & RDat2.D2 & "' and Faehigkeit = 'M'" _
        '                       & "or Eintritt <= '" & RDat1.D2 & "' and Austritt >= '" & RDat2.D2 & "' and Faehigkeit = 'TP'" _
        '                       & "or Eintritt <= '" & RDat1.D2 & "' and Austritt >= '" & RDat2.D2 & "' and Faehigkeit = 'FL' Order by MitarbeiterNr", pCon)

        'pCmd = New PsqlCommand("select * from T203Mitarbeiter Order by MitarbeiterNr", pCon)

        pReader = pCmd.ExecuteReader()

        While pReader.Read()


            MitarbeiterNummer = pReader(1).ToString()
            MitarbeiterName = Replace(pReader(4).ToString(), "'", "''")

            MitarbeiterCode = Trim(pReader(42).ToString())

                MitarbeiterEintritt = Trim(pReader(35).ToString())
                MitarbeiterAustritt = Trim(pReader(36).ToString())
                MitarbeiterStandort = pReader(41).ToString()
                MitarbeiterZuteilung = Trim(pReader(52).ToString())

                MitarbeiterAdresse = Replace(Trim(pReader(6).ToString()), "'", "''")
                MitarbeiterPLZ = Trim(pReader(8).ToString())
                MitarbeiterOrt = Replace(Trim(pReader(9).ToString()), "'", "''")
                MitarbeiterTelefon1 = Trim(pReader(10).ToString())
                MitarbeiterTelefon2 = Trim(pReader(11).ToString())
                MitarbeiterMobilnummer = Trim(pReader(14).ToString())
                MitarbeiterBild = Replace(Trim(pReader(50).ToString()), "'", "''")


            Select Case MitarbeiterZuteilung

                Case "Innenausbau"
                    MitarbeiterZuteilungCode = "INN"
                Case "InnenausbauHQ"
                    MitarbeiterZuteilungCode = "INN"
                Case "InnenausbauKü"
                    MitarbeiterZuteilungCode = "INN"
                Case "Labor"
                    MitarbeiterZuteilungCode = "INN"



                Case "Ladenbau"
                    MitarbeiterZuteilungCode = "LAD"



                Case "Küchen"
                    MitarbeiterZuteilungCode = "KÜC"
                Case "KüchenIKEA"
                    MitarbeiterZuteilungCode = "KÜC"



                Case "Türen"
                    MitarbeiterZuteilungCode = "T-B"
                Case "Betrieb"
                    MitarbeiterZuteilungCode = "T-B"



                Case "Glastrennwand"
                    MitarbeiterZuteilungCode = "GLA"



                Case "Fenster"
                    MitarbeiterZuteilungCode = "FEN"



                Case "Messebau"
                    MitarbeiterZuteilungCode = "MES"



                Case "Maler"
                    MitarbeiterZuteilungCode = "MAL"



                Case "Stadler"
                    MitarbeiterZuteilungCode = "STA"



                Case Else
                    MitarbeiterZuteilungCode = "INN"
            End Select

			If MitarbeiterEintritt <= RueckgabeDatum.D2 And MitarbeiterAustritt = "" Or MitarbeiterEintritt <= RueckgabeDatum.D2 And MitarbeiterAustritt >= RueckgabeDatum.D2 Then
				MitarbeiterErfassen()
			End If

			reset()

		End While

    Catch ex As PsqlException
        Console.WriteLine("Invalid" & ex.Message)
    Finally
        pCon.Close()
    End Try
End Sub


Private Sub MitarbeiterLesen()

    Dim pCon As PsqlConnection
    Dim pCmd As PsqlCommand
    Dim pReader As PsqlDataReader

    pCon = New PsqlConnection(PSQLSERVER)

    Try
        pCon.Open()

        'pCmd = New PsqlCommand("select * from T203Mitarbeiter where Faehigkeit = 'TL' or Faehigkeit = 'KO' or Faehigkeit = 'ML' or Faehigkeit = 'M' or Faehigkeit = 'TP' or Faehigkeit = 'FL' Order by MitarbeiterNr", pCon)

        pCmd = New PsqlCommand("select * from T203Mitarbeiter where Eintritt <= '" & RueckgabeDatum.D2 & "' and Faehigkeit = 'TL'" _
                               & "or Eintritt <= '" & RueckgabeDatum.D2 & "' and Faehigkeit = 'KO'" _
                               & "or Eintritt <= '" & RueckgabeDatum.D2 & "' and Faehigkeit = 'ML'" _
                               & "or Eintritt <= '" & RueckgabeDatum.D2 & "' and Faehigkeit = 'M'" _
                               & "or Eintritt <= '" & RueckgabeDatum.D2 & "' and Faehigkeit = 'TP'" _
                               & "or Eintritt <= '" & RueckgabeDatum.D2 & "' and Faehigkeit = 'FL' Order by MitarbeiterNr", pCon)

        pReader = pCmd.ExecuteReader()

        While pReader.Read()


            MitarbeiterNummer = pReader(1).ToString()
            MitarbeiterName = Replace(pReader(4).ToString(), "'", "''")

            'MitarbeiterCode = Trim(pReader(15).ToString())
            MitarbeiterCode = Trim(pReader(42).ToString())

            MitarbeiterEintritt = Trim(pReader(35).ToString())
            MitarbeiterAustritt = Trim(pReader(36).ToString())
            MitarbeiterStandort = pReader(41).ToString()
            MitarbeiterZuteilung = Trim(pReader(52).ToString())

            Select Case MitarbeiterZuteilung

                Case "Innenausbau"
                    MitarbeiterZuteilungCode = "INN"
                Case "InnenausbauHQ"
                    MitarbeiterZuteilungCode = "INN"
                Case "InnenausbauKü"
                    MitarbeiterZuteilungCode = "INN"
                Case "Labor"
                    MitarbeiterZuteilungCode = "INN"



                Case "Ladenbau"
                    MitarbeiterZuteilungCode = "LAD"



                Case "Küchen"
                    MitarbeiterZuteilungCode = "KÜC"
                Case "KüchenIKEA"
                    MitarbeiterZuteilungCode = "KÜC"



                Case "Türen"
                    MitarbeiterZuteilungCode = "T-B"
                Case "Betrieb"
                    MitarbeiterZuteilungCode = "T-B"



                Case "Glastrennwand"
                    MitarbeiterZuteilungCode = "GLA"



                Case "Fenster"
                    MitarbeiterZuteilungCode = "FEN"



                Case "Messebau"
                    MitarbeiterZuteilungCode = "MES"



                Case "Maler"
                    MitarbeiterZuteilungCode = "MAL"



                Case "Stadler"
                    MitarbeiterZuteilungCode = "STA"



                Case Else
                    MitarbeiterZuteilungCode = "INN"
            End Select


            Dim DatumCheck1 As String
            DatumCheck1 = RueckgabeDatum.D2

            If MitarbeiterVorhanden(MitarbeiterNummer) <> 0 Then
                MitarbeiterLoeschen(MitarbeiterNummer)
            End If

            If MitarbeiterAustritt = "" Then
                MitarbeiterErfassen()
            Else

                If DatumCheck1 <= MitarbeiterAustritt Then
                    MitarbeiterLoeschen(MitarbeiterNummer)
                    MitarbeiterErfassen()
                Else
                    MitarbeiterLoeschen(MitarbeiterNummer)
                End If

            End If

            reset()

        End While

    Catch ex As PsqlException
        Console.WriteLine("Invalid" & ex.Message)
    Finally
        pCon.Close()
    End Try
End Sub


Private Function MitarbeiterVorhanden(strMaNummer As String)

    Dim con As New SqlConnection
    Dim cmd As New SqlCommand
    Dim reader As DbDataReader
    Dim count As Integer

    con.ConnectionString = SQLSERVER

    Try

        con.Open()
        cmd.Connection = con

        cmd = New SqlCommand("select * from tbMitarbeiter where AdressNummer = '" & strMaNummer & "'", con)
        reader = cmd.ExecuteReader

        While reader.Read()
            count = count + 1
        End While

        reader.Close()

    Catch ex As Exception

        Console.WriteLine("Invalid" & ex.Message)

    Finally
        con.Close()
    End Try

    Return count

End Function

Private Sub MitarbeiterErfassen()

    Dim con As New SqlConnection
    Dim cmd As New SqlCommand

    Try

        con.ConnectionString = SQLSERVER
        con.Open()
        cmd.Connection = con

            cmd = New SqlCommand("insert into tbMitarbeiter ([CheckDatum], [AdressNummer], [NameVorname], [Code], [Eintritt], [Austritt], [Standort], [Zuteilung], [ZuteilungCode], [Adresse], [PLZ], [Ort], [Telefon1], [Telefon2], [Mobilnummer], [Bild]) " _
                             & " values ('" & RueckgabeDatum.D1 & "'," _
                             & "'" & MitarbeiterNummer & "'," _
                             & "'" & MitarbeiterName & "'," _
                             & "'" & MitarbeiterCode & "'," _
                             & "'" & MitarbeiterEintritt & "'," _
                             & "'" & MitarbeiterAustritt & "'," _
                             & "'" & MitarbeiterStandort & "'," _
                             & "'" & MitarbeiterZuteilung & "'," _
                             & "'" & MitarbeiterZuteilungCode & "'," _
                             & "'" & MitarbeiterAdresse & "'," _
                             & "'" & MitarbeiterPLZ & "'," _
                             & "'" & MitarbeiterOrt & "'," _
                             & "'" & MitarbeiterTelefon1 & "'," _
                             & "'" & MitarbeiterTelefon2 & "'," _
                             & "'" & MitarbeiterMobilnummer & "'," _
                             & "'" & MitarbeiterBild & "')", con)

        cmd.ExecuteNonQuery()

    Catch ex As Exception

        Console.WriteLine("Invalid" & ex.Message)
    Finally
        con.Close()
    End Try

End Sub

Private Sub MitarbeiterLoeschen(strMaNummer As String)


    Dim con As New SqlConnection
    Dim cmd As New SqlCommand
    Dim reader As DbDataReader

    con.ConnectionString = SQLSERVER

    Try

        con.Open()
        cmd.Connection = con

        cmd = New SqlCommand("delete from tbMitarbeiter where AdressNummer = '" & strMaNummer & "'", con)
        reader = cmd.ExecuteReader

        reader.Close()

    Catch ex As Exception

        Console.WriteLine("Invalid" & ex.Message)

    Finally
        con.Close()
    End Try

End Sub

Private Sub MitarbeiterLoeschenAlle()


    Dim con As New SqlConnection
    Dim cmd As New SqlCommand
    Dim reader As DbDataReader

    con.ConnectionString = SQLSERVER

    Try

        con.Open()
        cmd.Connection = con

        cmd = New SqlCommand("delete from tbMitarbeiter", con)
        reader = cmd.ExecuteReader

        reader.Close()

    Catch ex As Exception

        Console.WriteLine("Invalid" & ex.Message)

    Finally
        con.Close()
    End Try

End Sub

Private Sub TerminLesen()

    Dim pCon As PsqlConnection
    Dim pCmd As PsqlCommand
    Dim pReader As PsqlDataReader

    Dim Datum As String = RueckgabeDatum.D2

    pCon = New PsqlConnection(PSQLSERVER)

    Try
        pCon.Open()
        pCmd = New PsqlCommand("select * from T208Termin where Datum = '" & Datum & "'", pCon)

        pReader = pCmd.ExecuteReader()

        TerminLoeschen(RueckgabeDatum.D1)

		While pReader.Read()

			TerminAdressnummer = pReader(4).ToString()
			TerminDatum = pReader(1).ToString()
			TerminArt = pReader(16).ToString()

			TerminErfassen()

			reset()

		End While

    Catch ex As PsqlException

        Console.WriteLine("Invalid" & ex.Message)

    Finally
        pCon.Close()
    End Try
End Sub

Private Sub TerminLesenWP()

	Dim pCon As PsqlConnection
	Dim pCmd As PsqlCommand
	Dim pReader As PsqlDataReader

	pCon = New PsqlConnection(PSQLSERVER)

	TerminLoeschenWP()

	Dim Datum As String = MontagFestlegen(Now)

	Try
		pCon.Open()
		pCmd = New PsqlCommand("select * from T208Termin where Datum >= '" & Datum & "'", pCon)

		pReader = pCmd.ExecuteReader()

		While pReader.Read()

			TerminAdressnummer = pReader(4).ToString()
			TerminDatum = pReader(1).ToString()
			TerminArt = pReader(16).ToString()

			TerminErfassenWP()

			reset()

		End While

	Catch ex As PsqlException

		Console.WriteLine("Invalid" & ex.Message)

	Finally
		pCon.Close()
	End Try
End Sub

Private Sub TerminErfassen()

	Dim con As New SqlConnection
	Dim cmd As New SqlCommand
	Dim Moddate As DateTime = Now()

	Try

		con.ConnectionString = SQLSERVER
		con.Open()
		cmd.Connection = con

		cmd = New SqlCommand("insert into tbTermin ([AdressNummer], [Datum], [Art], [Moddate]) " _
							 & " values ('" & TerminAdressnummer & "','" & TerminDatum & "','" & TerminArt & "','" & Moddate & "')", con)

		cmd.ExecuteNonQuery()

	Catch ex As Exception

		Console.WriteLine("Invalid" & ex.Message)

	Finally

		con.Close()

	End Try

End Sub

Private Sub TerminErfassenWP()

	Dim con As New SqlConnection
	Dim cmd As New SqlCommand
	Dim Moddate As DateTime = Now()

	Try

		con.ConnectionString = SQLSERVER
		con.Open()
		cmd.Connection = con

		cmd = New SqlCommand("insert into tbTerminWP ([AdressNummer], [Datum], [Art], [Moddate]) " _
							 & " values ('" & TerminAdressnummer & "','" & TerminDatum & "','" & TerminArt & "','" & Moddate & "')", con)

		cmd.ExecuteNonQuery()

	Catch ex As Exception

		Console.WriteLine("Invalid" & ex.Message)

	Finally

		con.Close()

	End Try

End Sub

Private Sub TerminLoeschen(Datum As Date)

	Dim con As New SqlConnection
	Dim cmd As New SqlCommand
	Dim reader As DbDataReader

	con.ConnectionString = SQLSERVER

	Try
		con.Open()
		cmd.Connection = con

		cmd = New SqlCommand("delete from tbTermin Where Datum = '" & Datum & "'", con)
		reader = cmd.ExecuteReader

		reader.Close()

	Catch ex As Exception

		Console.WriteLine("Invalid" & ex.Message)

	Finally
		con.Close()
	End Try

End Sub


Private Sub TerminLoeschenWP()

	Dim con As New SqlConnection
	Dim cmd As New SqlCommand
	Dim reader As DbDataReader

	con.ConnectionString = SQLSERVER

	Try
		con.Open()
		cmd.Connection = con

		cmd = New SqlCommand("delete from tbTerminWP", con)
		reader = cmd.ExecuteReader

		reader.Close()

	Catch ex As Exception

		Console.WriteLine("Invalid" & ex.Message)

	Finally
		con.Close()
	End Try

End Sub


Private Sub FahrzeugLesen()

	Dim pCon As PsqlConnection
	Dim pCmd As PsqlCommand
	Dim pReader As PsqlDataReader

	pCon = New PsqlConnection(PSQLSERVER)

	FahrzeugLoeschen()

	Try
		pCon.Open()
		pCmd = New PsqlCommand("select * from T085Fahrzeuge", pCon)

		pReader = pCmd.ExecuteReader()

		While pReader.Read()

			FahrzeugADRESSNUMMER = pReader(1).ToString()
			FahrzeugFAHRZEUGNUMMER = pReader(2).ToString()
			FahrzeugMARKE = pReader(3).ToString()
			FahrzeugMODELL = pReader(4).ToString()
			FahrzeugANZAHLPLAETZE = pReader(7).ToString()
			FahrzeugKILOMETERSTAND = pReader(17).ToString()
			FahrzeugFARBE = pReader(18).ToString()
			FahrzeugPOLIZEIKENNZEICHEN = pReader(29).ToString()
			FahrzeugBILD = pReader(34).ToString()
			FahrzeugSUCHSCHLUESSEL = pReader(48).ToString()
			FahrzeugPARKPLATZNR = pReader(49).ToString()
			FahrzeugZUGETMONTEUR = pReader(50).ToString()
			FahrzeugDOKUMENT = pReader(51).ToString()
			FahrzeugBENZINKARTENNR = pReader(53).ToString()
			FahrzeugTREIBSTOFFART = pReader(54).ToString()

			FahrzeugErfassen()

			reset()

		End While

	Catch ex As PsqlException

		Console.WriteLine("Invalid" & ex.Message)

	Finally
		pCon.Close()
	End Try
End Sub

Private Sub FahrzeugErfassen()

	Dim con As New SqlConnection
	Dim cmd As New SqlCommand
	Dim Moddate As DateTime = Now()

	Try

		con.ConnectionString = SQLSERVER
		con.Open()
		cmd.Connection = con

		cmd = New SqlCommand("insert into tbFahrzeug ([AdressNummer]," _
								& "	[FAHRZEUGNUMMER]," _
								& "	[MARKE]," _
								& "	[MODELL]," _
								& "	[ANZAHLPLAETZE]," _
								& "	[KILOMETERSTAND]," _
								& "	[FARBE]," _
								& "	[POLIZEIKENNZEICHEN]," _
								& "	[BILD]," _
								& "	[SUCHSCHLUESSEL]," _
								& "	[PARKPLATZ-NR]," _
								& "	[ZUGETMONTEUR]," _
								& "	[DOKUMENT]," _
								& "	[BENZINKARTEN-NR]," _
								& "	[TREIBSTOFFART]) " _
								& " values ('" & FahrzeugADRESSNUMMER & "'" _
								& ",'" & FahrzeugFAHRZEUGNUMMER & "'" _
								& ",'" & FahrzeugMARKE & "'" _
								& ",'" & FahrzeugMODELL & "'" _
								& ",'" & FahrzeugANZAHLPLAETZE & "'" _
								& ",'" & FahrzeugKILOMETERSTAND & "'" _
								& ",'" & FahrzeugFARBE & "'" _
								& ",'" & FahrzeugPOLIZEIKENNZEICHEN & "'" _
								& ",'" & FahrzeugBILD & "'" _
								& ",'" & FahrzeugSUCHSCHLUESSEL & "'" _
								& ",'" & FahrzeugPARKPLATZNR & "'" _
								& ",'" & FahrzeugZUGETMONTEUR & "'" _
								& ",'" & FahrzeugDOKUMENT & "'" _
								& ",'" & FahrzeugBENZINKARTENNR & "'" _
								& ",'" & FahrzeugTREIBSTOFFART & "')", con)

		cmd.ExecuteNonQuery()

	Catch ex As Exception

		Console.WriteLine("Invalid" & ex.Message)

	Finally

		con.Close()

	End Try

End Sub

Private Sub FahrzeugLoeschen()

	Dim con As New SqlConnection
	Dim cmd As New SqlCommand
	Dim reader As DbDataReader

	con.ConnectionString = SQLSERVER

	Try
		con.Open()
		cmd.Connection = con

		cmd = New SqlCommand("delete from tbFahrzeug", con)
		reader = cmd.ExecuteReader

		reader.Close()

	Catch ex As Exception

		Console.WriteLine("Invalid" & ex.Message)

	Finally
		con.Close()
	End Try

End Sub

Private Sub EinsatzLesen()

    Dim pCon As PsqlConnection
    Dim pCmd As PsqlCommand
    Dim pReader As PsqlDataReader

    Dim Datum As String = RueckgabeDatum.D2

    pCon = New PsqlConnection(PSQLSERVER)

    Try
        pCon.Open()
        pCmd = New PsqlCommand("select * from T230Einsatz where Datum = '" & Datum & "'", pCon)

        pReader = pCmd.ExecuteReader()

        EinsatzLoeschen(RueckgabeDatum.D1)

        While pReader.Read()


            Dim CheckAdressnummer As String = Trim(pReader(1).ToString())

            'If CheckAdressnummer <> "" Then

            EinsatzAdressnummer = pReader(1).ToString()
            EinsatzAuftragsnummer = pReader(2).ToString()
            EinsatzTagesDatum = pReader(3).ToString()

			EinsatzErfassen()

            'End If

            reset()

        End While

    Catch ex As PsqlException

        Console.WriteLine("Invalid" & ex.Message)

    Finally
        pCon.Close()
    End Try
End Sub

    Private Sub EinsatzLesenWP()

        Dim pCon As PsqlConnection
        Dim pCmd As PsqlCommand
        Dim pReader As PsqlDataReader


		Dim Datum As String = MontagFestlegen(Now)

        pCon = New PsqlConnection(PSQLSERVER)

        Try
            pCon.Open()
            pCmd = New PsqlCommand("select * from T230Einsatz where Datum >= '" & Datum & "'", pCon)

            pReader = pCmd.ExecuteReader()

            EinsatzLoeschenWP()

            While pReader.Read()

                EinsatzAdressnummer = pReader(1).ToString()
                EinsatzAuftragsnummer = pReader(2).ToString()
                EinsatzTagesDatum = pReader(3).ToString()

                EinsatzSTATUSMACHEFMON = pReader(7).ToString()
                EinsatzABREISEZEIT = pReader(9).ToString()
                EinsatzBAUSTELLENZEIT = pReader(10).ToString()
                EinsatzENDZEIT = pReader(11).ToString()
                EinsatzFAHRZEUGNUMMER = pReader(12).ToString()
                EinsatzPARKPLATZ = "Park"
                EinsatzHOTELADRESSNR = pReader(13).ToString()
                EinsatzHOTELDETAIL = Replace(pReader(14).ToString(), "'", "''")
                EinsatzHOTELBESTELLT = pReader(25).ToString()
                EinsatzBESTELLNUMMER = pReader(26).ToString()

				EinsatzErfassenWP()

                'End If

                reset()

            End While

        Catch ex As PsqlException

            Console.WriteLine("Invalid" & ex.Message)

        Finally
            pCon.Close()
        End Try
    End Sub

    Private Sub EinsatzErfassen()

        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        Dim Moddate As DateTime = Now()

        Try

            con.ConnectionString = SQLSERVER
            con.Open()
            cmd.Connection = con

            cmd = New SqlCommand("insert into tbEinsatz ([AdressNummer]," _
                                 & "[Auftragsnummer]," _
                                 & "[Datum]," _
                                 & "[Moddate]) " _
                                 & " values ('" & EinsatzAdressnummer & "'," _
                                 & "'" & EinsatzAuftragsnummer & "'," _
                                 & "'" & EinsatzTagesDatum & "'," _
                                 & "'" & Moddate & "')", con)

            cmd.ExecuteNonQuery()

        Catch ex As Exception

            Console.WriteLine("Invalid" & ex.Message)

        Finally

            con.Close()

        End Try

	End Sub

Private Sub EinsatzErfassenWP()

		Dim con As New SqlConnection
		Dim cmd As New SqlCommand
		Dim Moddate As DateTime = Now()

		Try

			con.ConnectionString = SQLSERVER
			con.Open()
			cmd.Connection = con

			cmd = New SqlCommand("insert into tbEinsatzWP ([AdressNummer]," _
								 & "[Auftragsnummer]," _
								 & "[Datum]," _
								 & "[Moddate]," _
								 & "[STATUSMACHEFMON]," _
								 & "[ABREISEZEIT]," _
								 & "[BAUSTELLENZEIT]," _
								 & "[ENDZEIT]," _
								 & "[FAHRZEUGNUMMER]," _
								 & "[PARKPLATZ]," _
								 & "[HOTELADRESSNR]," _
								 & "[HOTELDETAIL]," _
								 & "[HOTELBESTELLT]," _
								 & "[BESTELLNUMMER]) " _
								 & " values ('" & EinsatzAdressnummer & "'," _
								 & "'" & EinsatzAuftragsnummer & "'," _
								 & "'" & EinsatzTagesDatum & "'," _
								 & "'" & Moddate & "'," _
								 & "'" & EinsatzSTATUSMACHEFMON & "'," _
								 & "'" & EinsatzABREISEZEIT & "'," _
								 & "'" & EinsatzBAUSTELLENZEIT & "'," _
								 & "'" & EinsatzENDZEIT & "'," _
								 & "'" & EinsatzFAHRZEUGNUMMER & "'," _
								 & "'" & EinsatzPARKPLATZ & "'," _
								 & "'" & EinsatzHOTELADRESSNR & "'," _
								 & "'" & EinsatzHOTELDETAIL & "'," _
								 & "'" & EinsatzHOTELBESTELLT & "'," _
								 & "'" & EinsatzBESTELLNUMMER & "')", con)

			cmd.ExecuteNonQuery()

		Catch ex As Exception

			Console.WriteLine("Invalid" & ex.Message)

		Finally

			con.Close()

		End Try

End Sub

Private Sub EinsatzLoeschen(Datum As Date)

    Dim con As New SqlConnection
    Dim cmd As New SqlCommand
    Dim reader As DbDataReader

    con.ConnectionString = SQLSERVER

    Try
        con.Open()
        cmd.Connection = con

        cmd = New SqlCommand("delete from tbEinsatz Where Datum = '" & Datum & "'", con)
        reader = cmd.ExecuteReader

        reader.Close()

    Catch ex As Exception

        Console.WriteLine("Invalid" & ex.Message)

    Finally
        con.Close()
    End Try

End Sub
    Private Sub EinsatzLoeschenWP()

        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        Dim reader As DbDataReader

        con.ConnectionString = SQLSERVER

        Try
            con.Open()
            cmd.Connection = con

			cmd = New SqlCommand("delete from tbEinsatzWP", con)
            reader = cmd.ExecuteReader

            reader.Close()

        Catch ex As Exception

            Console.WriteLine("Invalid" & ex.Message)

        Finally
            con.Close()
        End Try

    End Sub

Public Structure DatumWert

    Public D1 As Date
    Public D2 As String

End Structure


Function KalenderLesen(Montag As Integer)

    Dim con As New SqlConnection
    Dim cmd As New SqlCommand
    Dim reader As DbDataReader

    Dim Suchdatum As System.DateTime

    If DateTime.Now.ToString("dddd") = "Freitag" Then
        Suchdatum = DateTime.Now.Date
    Else
        Suchdatum = Today.AddDays(1)
	End If

		con.ConnectionString = SQLSERVER

    Try
        con.Open()
        cmd.Connection = con

        If Montag = 0 Then
            cmd = New SqlCommand("Select * from tbKalender where Datum = '" & Suchdatum & "'", con)
        Else
            cmd = New SqlCommand("Select * from tbKalender where Datum > '" & Suchdatum & "'", con)
        End If


        reader = cmd.ExecuteReader

        Dim CountMontag As Integer = 0

        While reader.Read

            If (reader(3).ToString) = "MO" Then
                CountMontag = CountMontag + 1
            End If

            If Montag = 0 Then
                RueckgabeDatum.D1 = (reader(1).ToString)
                RueckgabeDatum.D2 = (reader(4).ToString)
                Exit While
            End If

            If Montag = 1 And CountMontag = 1 Then
                RueckgabeDatum.D1 = (reader(1).ToString)
                RueckgabeDatum.D2 = (reader(4).ToString)
                Exit While
            End If

            If Montag = 2 And CountMontag = 2 Then
                RueckgabeDatum.D1 = (reader(1).ToString)
                RueckgabeDatum.D2 = (reader(4).ToString)
                Exit While
            End If

            If Montag = 3 And CountMontag = 3 Then
                RueckgabeDatum.D1 = (reader(1).ToString)
                RueckgabeDatum.D2 = (reader(4).ToString)
				Exit While
			End If

        End While

        reader.Close()

    Catch ex As Exception

        Console.WriteLine("Invalid" & ex.Message)

    Finally
        con.Close()
    End Try

    KalenderLesen = RueckgabeDatum


End Function

Function MontagFestlegen(D As String)

	Dim SuchDat As System.DateTime

	If DateTime.Now.ToString("dddd") = "Montag" Then
		SuchDat = Today
	End If

	If DateTime.Now.ToString("dddd") = "Dienstag" Then
		SuchDat = Today.AddDays(-1)
	End If

	If DateTime.Now.ToString("dddd") = "Mittwoch" Then
		SuchDat = Today.AddDays(-2)
	End If

	If DateTime.Now.ToString("dddd") = "Donnerstag" Then
		SuchDat = Today.AddDays(-3)
	End If

	If DateTime.Now.ToString("dddd") = "Freitag" Then
		SuchDat = Today.AddDays(-4)
	End If

	If DateTime.Now.ToString("dddd") = "Samstag" Then
		SuchDat = Today.AddDays(-5)
	End If

	If DateTime.Now.ToString("dddd") = "Sonntag" Then
		SuchDat = Today.AddDays(-6)
	End If

	Dim con As New SqlConnection
	Dim cmd As New SqlCommand
	Dim reader As DbDataReader

	con.ConnectionString = SQLSERVER

	Try
		con.Open()
		cmd.Connection = con

		cmd = New SqlCommand("Select * from tbKalender where Datum = '" & SuchDat & "'", con)

		reader = cmd.ExecuteReader

		While reader.Read
			MontagFestlegen = reader(4).ToString
		End While

		reader.Close()

	Catch ex As Exception

		Console.WriteLine("Invalid" & ex.Message)

	Finally
		con.Close()
	End Try

End Function


Private Sub reset()

        MitarbeiterNummer = ""
        MitarbeiterName = ""
        MitarbeiterCode = ""
        MitarbeiterEintritt = ""
        MitarbeiterAustritt = ""
        MitarbeiterStandort = ""
        MitarbeiterZuteilung = ""
        MitarbeiterZuteilungCode = ""
        MitarbeiterAdresse = ""
        MitarbeiterPLZ = ""
        MitarbeiterOrt = ""
        MitarbeiterTelefon1 = ""
        MitarbeiterTelefon2 = ""
        MitarbeiterMobilnummer = ""
        MitarbeiterBild = ""


        EinsatzAdressnummer = ""
        EinsatzAuftragsnummer = ""
        EinsatzTagesDatum = ""

        EinsatzSTATUSMACHEFMON = ""
        EinsatzABREISEZEIT = ""
        EinsatzBAUSTELLENZEIT = ""
        EinsatzENDZEIT = ""
        EinsatzFAHRZEUGNUMMER = ""
        EinsatzPARKPLATZ = ""
        EinsatzHOTELADRESSNR = ""
        EinsatzHOTELDETAIL = ""
        EinsatzHOTELBESTELLT = ""
        EinsatzBESTELLNUMMER = ""

		FahrzeugADRESSNUMMER = ""
		FahrzeugFAHRZEUGNUMMER = ""
		FahrzeugMARKE = ""
		FahrzeugMODELL = ""
		FahrzeugANZAHLPLAETZE = ""
		FahrzeugKILOMETERSTAND = ""
		FahrzeugFARBE = ""
		FahrzeugPOLIZEIKENNZEICHEN = ""
		FahrzeugBILD = ""
		FahrzeugSUCHSCHLUESSEL = ""
		FahrzeugPARKPLATZNR = ""
		FahrzeugZUGETMONTEUR = ""
		FahrzeugDOKUMENT = ""
		FahrzeugBENZINKARTENNR = ""
		FahrzeugTREIBSTOFFART = ""

        TerminAdressnummer = ""
        TerminDatum = ""
        TerminArt = ""
        TerminKW = ""
        TerminMontagFlag = False

        AuftragNummer = ""
        AuftragZuteilung = ""
        AuftragZuteilungCode = ""
        AuftragName = ""

        AuftragOBJEKTBESCHREIB1 = ""
        AuftragOBJEKTBESCHREIB2 = ""
        AuftragTERMIN1VON = ""
        AuftragTERMIN1BIS = ""
        AuftragANZAHLMA1 = ""
        AuftragANKUNFTSZEIT1 = ""
        AuftragTERMIN2VON = ""
        AuftragTERMIN2BIS = ""
        AuftragANZAHLMA2 = ""
        AuftragANKUNFTSZEIT2 = ""
        AuftragTERMIN3VON = ""
        AuftragTERMIN3BIS = ""
        AuftragANZAHLMA3 = ""
        AuftragANKUNFTSZEIT3 = ""
        AuftragTERMIN4VON = ""
        AuftragTEMIN4BIS = ""
        AuftragANZAHLMA4 = ""
        AuftragANKUNFTSZEIT4 = ""
        AuftragTERMIN5VON = ""
        AuftragTERMIN5BIS = ""
        AuftragANZAHLMA5 = ""
        AuftragANKUNFTSZEIT5 = ""
        AuftragFARBCODEAUFTRAG = ""
        AuftragURSPRUNGSMANDANT = ""
        AuftragSATERMIN1 = ""
        AuftragSOTERMIN1 = ""
        AuftragTERMIN1FEIERTAG1 = ""
        AuftragSATERMIN2 = ""
        AuftragSOTERMIN2 = ""
        AuftragTERMIN2FEIERTAG2 = ""
        AuftragSATERMIN3 = ""
        AuftragSOTERMIN3 = ""
        AuftragTERMIN3FEIERTAG3 = ""
        AuftragSATERMIN4 = ""
        AuftragSOTERMIN4 = ""
        AuftragTERMIN4FEIERTAG4 = ""
        AuftragSATERMIN5 = ""
        AuftragSOTERMIN5 = ""
        AuftragTERMIN5FEIERTAG5 = ""
        AuftragMONTAGEART = ""
        AuftragAUFTRAGSBETREUER = ""
        AuftragBEMERKUNGEN = ""
        AuftragUMRECHNUNGSFAKT = ""
        AuftragAUFTRAGSEMAIL = ""
        AuftragAUFTRAGSFAX = ""
        AuftragAUFTRAGBESTAETIG = ""
        AuftragCHMONTEUREMAIL = ""
        AuftragCHMONTEURFAX = ""


End Sub

End Module
