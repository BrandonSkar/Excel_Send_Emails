Attribute VB_Name = "Module1"
Sub Send_EoD()
    Call email
End Sub

Private Sub email()
    '_____________________________________________________________________________________________________________________________________________________________________________________________
    '           EXPEDITORS CONTACTS TO SEND EMPTIES EMAILS TO, CHANGE ANY OF THESE TO UPDATE ALL EMAILS
    Const contacts As String = "Steve Tiffany <Steve.Tiffany@expeditors.com>; Danyel Clair <Danyel.Clair@expeditors.com>; Mallory Hill <Mallory.Hill@expeditors.com>;" _
        & "Rose Asu <Rose.Asu@expeditors.com>; Ishtiaq Aftab <Ishtiaq.Aftab@expeditors.com>; Dante Suarez <Dante.Suarez@expeditors.com>; Marissa Bateman <Marissa.Bateman@expeditors.com>;" _
        & "Jessica Wright <Jessica.Wright2@expeditors.com>;Kayla.Barbour@expeditors.com>;nick.dyer@expeditors.com;"
    '_____________________________________________________________________________________________________________________________________________________________________________________________
    
    Const path As String = "F:\EI\SEA\Distribution\S167\"
    
    'DEBUG
    'Const path As String = "C:\Users\sea-ishtiaqa\Desktop\y.xlsm"
    'Const path As String = "C:\Users\sea-brandons.EXPEDITORS\Desktop\YARD CHECK 3.6.18 .xlsm"
    
    Dim iLastRow As Integer
    Dim sFound As String
    sFound = Dir(path & "YARD CHECK*.xlsm")
    Dim bFileOpen As Boolean
    bFileOpen = IsWorkBookOpen(path & sFound)
    
    'DEBUG
    'bFileOpen = IsWorkBookOpen(path)
    
    If bFileOpen Then
        MsgBox sFound & " is open. If you really want to send out emails then close the yard first."
        Exit Sub
    End If
    Dim rng As Range
    Dim header As Range
    Dim finalRng As String
    Dim wkb As Workbook
    Dim OutApp As Object, OutMail As Object
    Dim aCarriers As Variant
    Dim aPCT() As Integer, aKNIGHT() As Integer, aPREMIER() As Integer, aHONEST() As Integer, aEBT() As Integer
    Dim aSAVANAH() As Integer, aLEGS() As Integer, aAML() As Integer, aGSC() As Integer
    Dim PCTbool As Boolean, KNIGHTbool As Boolean, PREMIERbool As Boolean, HONESTbool As Boolean, EBTbool As Boolean
    Dim SAVANAHbool As Boolean, LEGSbool As Boolean, AMLbool As Boolean, GSCbool As Boolean
    Dim pctSize As Integer, knightSize As Integer, premierSize As Integer, honestSize As Integer, ebtSize As Integer
    Dim savanaSize As Integer, legsSize As Integer, amlSize As Integer, gscSize As Integer
    PCTbool = False
    KNIGHTbool = False
    PREMIERbool = False
    HONESTbool = False
    SAVANAHbool = False
    LEGSbool = False
    AMLbool = False
    GSCbool = False
    EBTbool = False
    Dim signature As String
    
    '____________________________________________________________________________________________
    '  CARRIERS TO SEARCH FOR IN ORDER TO CREATE EMPTIES EMAIL
    aCarriers = Array("PCT", "KNIGHT", "PREMIER", "HONEST TRK", "SAVANAH", "LEGS", "AML", "GSC", "EBT")
    '____________________________________________________________________________________________
    
    'DEBUG
    'aCarriers = Array("PCT", "KNIGHT", "PREMIER")
    
    Set wkb = Workbooks.Open(path & sFound)
    wkb.Sheets("YARD").Activate
    iLastRow = Cells(Rows.Count, 3).End(xlUp).row
    
    'SORT YARD CHECK BY GATE IN DATE
    SortByDate (iLastRow)
    
    Set rng = Nothing
    Set header = Range("B1, C1, D1, E1, F1, G1, K1")
    On Error Resume Next
    
    '__________________________________________________________
    '  STORE ROW NUMBERS INTO AN ARRAY FOR EACH CARRIER
    '__________________________________________________________
    For i = 1 To iLastRow
        If Cells(i, 4).Value = "PCT" And Cells(i, 11).Value = "EMPTY" Then
            PCTbool = True
            ReDim Preserve aPCT(pctSize)
            aPCT(pctSize) = i
            pctSize = pctSize + 1
        ElseIf Cells(i, 4).Value = "KNIGHT" And Cells(i, 11).Value = "EMPTY" Then
            KNIGHTbool = True
            ReDim Preserve aKNIGHT(knightSize)
            aKNIGHT(knightSize) = i
            knightSize = knightSize + 1
        ElseIf Cells(i, 4).Value = "PREMIER" And Cells(i, 11).Value = "EMPTY" Then
            PREMIERbool = True
            ReDim Preserve aPREMIER(premierSize)
            aPREMIER(premierSize) = i
            premierSize = premierSize + 1
        ElseIf Cells(i, 4).Value = "HONEST TRK" And Cells(i, 11).Value = "EMPTY" Then
            HONESTbool = True
            ReDim Preserve aHONEST(honestSize)
            aHONEST(honestSize) = i
            honestSize = honestSize + 1
        ElseIf Cells(i, 4).Value = "SAVANAH" And Cells(i, 11).Value = "EMPTY" Then
            SAVANAHbool = True
            ReDim Preserve aSAVANAH(savanahsize)
            aSAVANAH(savanahsize) = i
            savanahsize = savanahsize + 1
        ElseIf Cells(i, 4).Value = "LEGS" And Cells(i, 11).Value = "EMPTY" Then
            LEGSbool = True
            ReDim Preserve aLEGS(legsSize)
            aLEGS(legsSize) = i
            legsSize = legsSize + 1
        ElseIf Cells(i, 4).Value = "AML" And Cells(i, 11).Value = "EMPTY" Then
            AMLbool = True
            ReDim Preserve aAML(amlSize)
            aAML(amlSize) = i
            amlSize = amlSize + 1
        ElseIf Cells(i, 4).Value = "GSC" And Cells(i, 11).Value = "EMPTY" Then
            GSCbool = True
            ReDim Preserve aGSC(gscSize)
            aGSC(gscSize) = i
            gscSize = gscSize + 1
        ElseIf Cells(i, 4).Value = "EBT" And Cells(i, 11).Value = "EMPTY" Then
            EBTbool = True
            ReDim Preserve aEBT(ebtSize)
            aEBT(ebtSize) = i
            ebtSize = ebtSize + 1
        End If
    Next i
    
    With Application
        .EnableEvents = False
        .ScreenUpdating = False
    End With
    On Error Resume Next
    
    '_____________________________________________________________________________
    '  CREATE EMAIL AND PUT EACH ROW OF SPECIFIED CARRIER INTO THAT EMAIL
    '_____________________________________________________________________________
    For Each element In aCarriers
        If element = "PCT" And PCTbool Then
            Set OutApp = CreateObject("Outlook.Application")
            Set OutMail = OutApp.CreateItem(0)
            OutMail.display
            signature = OutMail.HTMLBody
            
            With OutMail
            .To = "tchandler@pctrucking.net; markt@pctrucking.net; frank@pctrucking.net; dispatch <dispatch@pctrucking.net>"
            .body = "Hello PCT," & vbCrLf & "Please see below list of containers available for pickup."
            .Subject = "S167 Empty Containers"
            .CC = contacts
            .BCC = ""
            .HTMLBody = .HTMLBody & RangetoHTML(header)
            
            For Each cel In aPCT
                Set rng = Sheets("YARD").Range("B" & cel & ", C" & cel & ", D" & cel & ", E" & cel & ", F" & cel & ", G" & cel & ", K" & cel).SpecialCells(xlCellTypeVisible)
                .HTMLBody = .HTMLBody & RangetoHTML(rng)
            Next cel
            .HTMLBody = .HTMLBody & vbNewLine & vbNewLine & signature
            End With
        ElseIf element = "KNIGHT" And KNIGHTbool Then
            Set OutApp = CreateObject("Outlook.Application")
            Set OutMail = OutApp.CreateItem(0)
            OutMail.display
            signature = OutMail.HTMLBody
            
            With OutMail
            .To = "dave.brown@knighttrans.com; andrew.zasimovich@knighttrans.com; evan.taylor@knighttrans.com"
            .body = "Hello Knight," & vbCrLf & "Please see below list of containers available for pickup."
            .Subject = "S167 Empty Containers"
            .CC = contacts
            .BCC = ""
            .HTMLBody = .HTMLBody & RangetoHTML(header)
            
            For Each cel In aKNIGHT
                Set rng = Sheets("YARD").Range("B" & cel & ", C" & cel & ", D" & cel & ", E" & cel & ", F" & cel & ", G" & cel & ", K" & cel).SpecialCells(xlCellTypeVisible)
                .HTMLBody = .HTMLBody & RangetoHTML(rng)
            Next cel
            .HTMLBody = .HTMLBody & vbNewLine & vbNewLine & signature
            End With
        ElseIf element = "PREMIER" And PREMIERbool Then
            Set OutApp = CreateObject("Outlook.Application")
            Set OutMail = OutApp.CreateItem(0)
            OutMail.display
            signature = OutMail.HTMLBody
            
            With OutMail
            .To = "Lisa Miller <lmiller@premiertransportation.com>; pproctor@premiertransportation.com"
            .body = "Hello Premier," & vbCrLf & "Please see below list of containers available for pickup."
            .Subject = "S167 Empty Containers"
            .CC = contacts
            .BCC = ""
            .HTMLBody = .HTMLBody & RangetoHTML(header)
            For Each cel In aPREMIER
                Set rng = Sheets("YARD").Range("B" & cel & ", C" & cel & ", D" & cel & ", E" & cel & ", F" & cel & ", G" & cel & ", K" & cel).SpecialCells(xlCellTypeVisible)
                .HTMLBody = .HTMLBody & RangetoHTML(rng)
            Next cel
            
            .HTMLBody = .HTMLBody & vbNewLine & vbNewLine & signature
            End With
        ElseIf element = "HONEST TRK" And HONESTbool Then
            Set OutApp = CreateObject("Outlook.Application")
            Set OutMail = OutApp.CreateItem(0)
            OutMail.display
            signature = OutMail.HTMLBody
            
            With OutMail
            .To = "dispatch@honesttrucking.com"
            .body = "Hello Honest Trucking," & vbCrLf & "Please see below list of containers available for pickup."
            .Subject = "S167 Empty Containers"
            .CC = contacts
            .BCC = ""
            .HTMLBody = .HTMLBody & RangetoHTML(header)
            
            For Each cel In aHONEST
                Set rng = Sheets("YARD").Range("B" & cel & ", C" & cel & ", D" & cel & ", E" & cel & ", F" & cel & ", G" & cel & ", K" & cel).SpecialCells(xlCellTypeVisible)
                .HTMLBody = .HTMLBody & RangetoHTML(rng)
            Next cel
            .HTMLBody = .HTMLBody & vbNewLine & vbNewLine & signature
            End With
        ElseIf element = "SAVANAH" And SAVANAHbool Then
            Set OutApp = CreateObject("Outlook.Application")
            Set OutMail = OutApp.CreateItem(0)
            OutMail.display
            signature = OutMail.HTMLBody
            
            With OutMail
            .To = "mike@savanahlogistics.com; yosief@savanahlogistics.com; jeremy@savanahlogistics.com; gary@savanahlogistics.com; dispatch7@savanahlogistics.com"
            .body = "Hello Savanah," & vbCrLf & "Please see below list of containers available for pickup."
            .Subject = "S167 Empty Containers"
            .CC = contacts
            .BCC = ""
            .HTMLBody = .HTMLBody & RangetoHTML(header)
            
            For Each cel In aSAVANAH
                Set rng = Sheets("YARD").Range("B" & cel & ", C" & cel & ", D" & cel & ", E" & cel & ", F" & cel & ", G" & cel & ", K" & cel).SpecialCells(xlCellTypeVisible)
                .HTMLBody = .HTMLBody & RangetoHTML(rng)
            Next cel
            .HTMLBody = .HTMLBody & vbNewLine & vbNewLine & signature
            End With
        ElseIf element = "LEGS" And LEGSbool Then
            Set OutApp = CreateObject("Outlook.Application")
            Set OutMail = OutApp.CreateItem(0)
            OutMail.display
            signature = OutMail.HTMLBody
            
            With OutMail
            .To = "dispatch@legendinc.com; rlopez@newlegendinc.com"
            .body = "Hello Legend," & vbCrLf & "Please see below list of containers available for pickup."
            .Subject = "S167 Empty Containers"
            .CC = contacts
            .BCC = ""
            .HTMLBody = .HTMLBody & RangetoHTML(header)
            
            For Each cel In aLEGS
                Set rng = Sheets("YARD").Range("B" & cel & ", C" & cel & ", D" & cel & ", E" & cel & ", F" & cel & ", G" & cel & ", K" & cel).SpecialCells(xlCellTypeVisible)
                .HTMLBody = .HTMLBody & RangetoHTML(rng)
            Next cel
            .HTMLBody = .HTMLBody & vbNewLine & vbNewLine & signature
            End With
        ElseIf element = "AML" And AMLbool Then
            Set OutApp = CreateObject("Outlook.Application")
            Set OutMail = OutApp.CreateItem(0)
            OutMail.display
            signature = OutMail.HTMLBody
            
            With OutMail
            .To = "dispatch@amazonlogisticsllc.com"
            .body = "Hello AML," & vbCrLf & "Please see below list of containers available for pickup."
            .Subject = "S167 Empty Containers"
            .CC = contacts
            .BCC = ""
            .HTMLBody = .HTMLBody & RangetoHTML(header)
            
            For Each cel In aAML
                Set rng = Sheets("YARD").Range("B" & cel & ", C" & cel & ", D" & cel & ", E" & cel & ", F" & cel & ", G" & cel & ", K" & cel).SpecialCells(xlCellTypeVisible)
                .HTMLBody = .HTMLBody & RangetoHTML(rng)
            Next cel
            .HTMLBody = .HTMLBody & vbNewLine & vbNewLine & signature
            End With
        '12121'
        ElseIf element = "EBT" And EBTbool Then
            Set OutApp = CreateObject("Outlook.Application")
            Set OutMail = OutApp.CreateItem(0)
            OutMail.display
            signature = OutMail.HTMLBody
            
            With OutMail
            .To = "dispatch@elliottbaytransfer.com"
            .body = "Hello Elliott Bay," & vbCrLf & "Please see below list of containers available for pickup."
            .Subject = "S167 Empty Containers"
            .CC = contacts
            .BCC = ""
            .HTMLBody = .HTMLBody & RangetoHTML(header)
            
            For Each cel In aEBT
                Set rng = Sheets("YARD").Range("B" & cel & ", C" & cel & ", D" & cel & ", E" & cel & ", F" & cel & ", G" & cel & ", K" & cel).SpecialCells(xlCellTypeVisible)
                .HTMLBody = .HTMLBody & RangetoHTML(rng)
            Next cel
            .HTMLBody = .HTMLBody & vbNewLine & vbNewLine & signature
            End With
        ElseIf element = "GSC" And GSCbool Then
            Set OutApp = CreateObject("Outlook.Application")
            Set OutMail = OutApp.CreateItem(0)
            OutMail.display
            signature = OutMail.HTMLBody
            
            With OutMail
            .To = "EITJX@gsclogistics.com"
            .body = "Hello GSC," & vbCrLf & "Please see below list of containers available for pickup."
            .Subject = "S167 Empty Containers"
            .CC = contacts
            .BCC = ""
            .HTMLBody = .HTMLBody & RangetoHTML(header)
            
            For Each cel In aGSC
                Set rng = Sheets("YARD").Range("B" & cel & ", C" & cel & ", D" & cel & ", E" & cel & ", F" & cel & ", G" & cel & ", K" & cel).SpecialCells(xlCellTypeVisible)
                .HTMLBody = .HTMLBody & RangetoHTML(rng)
            Next cel
            .HTMLBody = .HTMLBody & vbNewLine & vbNewLine & signature
            End With
        End If
    Next element
    wkb.Close False
    On Error GoTo 0

    With Application
        .EnableEvents = True
        .ScreenUpdating = True
    End With

    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub

Function RangetoHTML(rng As Range)
    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook

    TempFile = Environ$("temp") & "\" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"

    'Copy the range and create a new workbook to paste the data in
    rng.Copy
    Set TempWB = Workbooks.Add(1)
    With TempWB.Sheets(1)
        .Cells(1).PasteSpecial Paste:=8
        .Cells(1).PasteSpecial xlPasteValues, , False, False
        .Cells(1).PasteSpecial xlPasteFormats, , False, False
        .Cells(1).Select
        Application.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.Visible = True
        .DrawingObjects.Delete
        On Error GoTo 0
    End With

    'Publish the sheet to a htm file
    With TempWB.PublishObjects.Add( _
         SourceType:=xlSourceRange, _
         FileName:=TempFile, _
         Sheet:=TempWB.Sheets(1).Name, _
         Source:=TempWB.Sheets(1).UsedRange.Address, _
         HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With

    'Read all data from the htm file into RangetoHTML
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangetoHTML = ts.readall
    ts.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
                          "align=left x:publishsource=")
    RangetoHTML = Replace(RangetoHTML, "<!--[if !excel]>&nbsp;&nbsp;<![endif]-->", "")
    'Close TempWB
    TempWB.Close savechanges:=False
    Debug.Print RangetoHTML
    'Delete the htm file we used in this function
    Kill TempFile

    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing
End Function


Function IsWorkBookOpen(FileName As String)
    Dim ff As Long, ErrNo As Long

    On Error Resume Next
    ff = FreeFile()
    Open FileName For Input Lock Read As #ff
    Close ff
    ErrNo = Err
    On Error GoTo 0

    Select Case ErrNo
    Case 0:    IsWorkBookOpen = False
    Case 70:   IsWorkBookOpen = True
    Case Else: Error ErrNo
    End Select
End Function

Private Sub SortByDate(iLastRow As Integer)
    lastrow = Cells(Rows.Count, 2).End(xlUp).row
    Range("B2:K" & iLastRow).Sort key1:=Range("F2:F" & lastrow), _
        order1:=xlAscending, header:=xlNo
End Sub
