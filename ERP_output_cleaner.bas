Attribute VB_Name = "ERP_output_cleaner"
Option Explicit

Sub GekoCleaner()
    Dim answer As Long
    Dim answer1 As Long
    Dim answer2 As Long
    answer = MsgBox("Avviare la procedura di conversione dell'output Geko in dataset? L'operazione è irreversibile e l'output Geko originale verrà perso.", vbYesNo)
    If answer = vbNo Then
        MsgBox "Annullato dall'utente"
        Exit Sub
    End If
    If InStr(Range("A3").Value, "I N V E N T A R I O    M A G A Z Z I N O") <> 0 Then
        inventario_cleaner
    ElseIf InStr(Range("A3").Value, "S T O R I C O    P R E Z Z I") <> 0 Then
        storicoPrezzi_cleaner
    ElseIf InStr(Range("A3").Value, "R U B R I C A    A R T I C O L I") <> 0 Then
        rubricaArticoli_cleaner
    ElseIf InStr(Range("A3").Value, "S C H E D A    A R T I C O L O") <> 0 Then
        schedaArticoli_cleaner
    ElseIf InStr(Range("A3").Value, "S T O R I C O    C O N T A B I L E") <> 0 Then
        storicoContabile_cleaner
    ElseIf InStr(Range("A3").Value, "E L E N C O    F O R N I T O R I") <> 0 Then
        elencoFornitori_cleaner
    ElseIf InStr(Range("A3").Value, "E L E N C O    C L I E N T I") <> 0 Then
        elencoClienti_cleaner
    Else
        MsgBox "Output Geko non riconosciuto, contattare amministratore."
    End If
End Sub

Private Sub inventario_cleaner()

    'dichiarazione variabili
   Dim lastRow As Long, lastRow2 As Long
   Dim CONTATORE As Long
   Dim rng As Range
   
    'inizializzo la variabile contatore utilizzata nello script per evitare un loop infinito (soluzione non ottimale)
   CONTATORE = 0
   
    'Pulizia delle righe iniziale manuale
    Range("A1:A6").EntireRow.Delete
    'setto la cella di inizio ciclo
    Range("A2").Select
   
   'TECNICA - LOOP RIGA PER RIGA FACCIO UN CHECK SE CI SONO CARATTERI CHE MI DICONO CHE LA RIGA VA CANCELLATA
   'Il primo loop fa pulizie di tutte le righe che interrompono la base di dati (righe vuote con trattini o con scritte varie)
   Do Until CONTATORE > 15
       If IsEmpty(ActiveCell) = True Or InStr(ActiveCell.Value, "---") <> 0 Or InStr(ActiveCell.Value, "COPPO 20") <> 0 Or InStr(ActiveCell.Value, "001 Ventimiglia") <> 0 Or InStr(ActiveCell.Value, "ESISTENZA") <> 0 Or InStr(ActiveCell.Value, "VALORE INVENTARIALE") <> 0 Or InStr(ActiveCell.Value, "ARTICOLO") <> 0 Or InStr(ActiveCell.Value, "Tipo Stampa") <> 0 Then
           Rows(ActiveCell.Row).Delete
           CONTATORE = CONTATORE + 1
       Else
           Selection.Offset(1, 0).Select
           CONTATORE = 0
       End If
   Loop

    'dati text to column
    Columns(1).TextToColumns TrailingMinusNumbers:=True, _
    DataType:=xlFixedWidth, _
    FieldInfo:=Array(Array(0, 1), Array(16, 1), Array(56, 1), Array(71, 1), Array(85, 1), Array(103, 1)), _
    TrailingMinusNumbers:=True
       
       
    Columns(2).AutoFit 'autofit per maggior leggibilita
    'Range("A1").EntireColumn.NumberFormat = 0
    Range("A3").Select 'risetto l'active cell a inizio foglio
    CONTATORE = 0 'risetto il contatore
    
    'Aggiungo una colonna per la descrizione 2
    With Range("C1")
        .EntireColumn.Insert
        .EntireColumn.HorizontalAlignment = xlLeft
    End With
    
    'Il secondo loop risolve il problema delle righe a capo che si verifica quando si utilizza la funzione text to column su record che hanno anche la descrizione 2
    Do Until CONTATORE > 15
        If IsEmpty(ActiveCell) = True Then
            ActiveCell.Select
            If IsError(ActiveCell.Offset(0, 1).Value) = True Then
                Rows(ActiveCell.Row).Delete
            End If
            'Il seguente IF determina se le primi due colonne della riga sono vuote. In tal caso esce dal loop perchè il dataset è finito
            If IsEmpty(ActiveCell.Offset(0, 1).Value) = False Then
                ActiveCell.Offset(-1, 2).Value = UCase(ActiveCell.Offset(0, 1).Value)
                Rows(ActiveCell.Row).Delete
                CONTATORE = CONTATORE + 1
            Else
                Exit Do
            End If
        ElseIf IsEmpty(ActiveCell.Offset(0, 1)) = True Then
            ActiveCell.Offset(-1, 2).Value = UCase(ActiveCell.Value)
            Rows(ActiveCell.Row).Delete
            CONTATORE = CONTATORE + 1
        Else
            Selection.Offset(1, 0).Select
            CONTATORE = 0
       End If
    Loop

    'GESTISCO IL CASO IN CUI LA FUNZIONE TESTO IN COLONNE PRODUCA 7 COLONNE (PROFUMI BHPC)
    If Range("A1", Range("A1").End(xlToRight)).Columns.Count = 7 Then
        Range("D1").EntireColumn.Insert
        lastRow = Range(Range("A1").End(xlDown)).Row - 1
        'Pulisco le celle da blank spaces con la funzione trim
        Application.ScreenUpdating = False
        Set rng = Range("B1", Range("C1").End(xlDown))
        With rng
            .Value = Evaluate(Replace("If(@="""","""",Trim(@))", "@", .Address))
        End With
        Application.ScreenUpdating = True
        Range("D2:D" & lastRow).Formula = "=CONCATENATE(B2,"" "",C2)"
        Range("D2:D" & lastRow).Copy
        Range("D2:D" & lastRow).PasteSpecial xlPasteValues
        Range("B1:C1").EntireColumn.Delete
    End If

    If Range("D1").Value = "" Then
        Range("D1").Value = "A"
    End If
    
    'GESTISCO IL CASO IN CUI LA FUNZIONE TESTO IN COLONNE PRODUCA 8 COLONNE (OCCHIALI BHPC)
    If Range("A1", Range("A1").End(xlToRight)).Columns.Count = 8 Then
        Range("E1").EntireColumn.Insert
        lastRow2 = ActiveSheet.Range(Range("A1").End(xlDown)).Row + 1
        'Pulisco le celle da blank spaces con la funzione trim
        Application.ScreenUpdating = False
        Set rng = Range("B1", Range("D1").End(xlDown))
        With rng
            .Value = Evaluate(Replace("If(@="""","""",Trim(@))", "@", .Address))
        End With
        Application.ScreenUpdating = True
        Range("E2:E" & lastRow2).Formula = "=CONCATENATE(B2,"" "",C2,"" "",D2)"
        Range("E2:E" & lastRow2).Copy
        Range("E2:E" & lastRow2).PasteSpecial xlPasteValues
        Range("B1:D1").EntireColumn.Delete
        ActiveSheet.Range("B:B").Replace What:="SU NGLAS SES", Replacement:="SUNGLASSES"
    End If

    'Reimposto i titoli per maggiore correttezza
    Range("A1").Value = "ARTICOLO"
    Range("B1").Value = "DESCRIZIONE"
    Range("C1").Value = "DESCRIZIONE 2"
    Range("D1").Value = "U.M."
    Range("E1").Value = "ESISTENZA"
    Range("F1").Value = "COSTO UNITARIO"
    Range("G1").Value = "COSTO GLOBALE"

    'imposto i formati
    Range("F1:G1").EntireColumn.NumberFormat = "$#,##0.00"
    Range("E1").EntireColumn.NumberFormat = "0"
    'autofit per maggior leggibilita
    Columns.AutoFit
    Range("A1").Select
    With ActiveWindow
        Rows("2:2").Select
        .FreezePanes = True
    End With
    With Range("C1")
        .EntireColumn.HorizontalAlignment = xlLeft
    End With
    
End Sub

Private Sub rubricaArticoli_cleaner()
   
   Dim CONTATORE As Long
   CONTATORE = 0
    'Pulizia delle righe iniziale manuale
    Range("A1:A6").EntireRow.Delete
    'setto la cella di inizio ciclo
    Range("A2").Select
   
   'Il primo loop fa pulizie di tutte le righe che interrompono la base di dati (righe vuote con trattini o con scritte varie)
   Do Until CONTATORE > 15
       If IsEmpty(ActiveCell) = True Or InStr(ActiveCell.Value, "---") <> 0 Or InStr(ActiveCell.Value, "COPPO 20") <> 0 Or InStr(ActiveCell.Value, "001 Ventimiglia") <> 0 Or InStr(ActiveCell.Value, "ESISTENZA") <> 0 Or InStr(ActiveCell.Value, "VALORE INVENTARIALE") <> 0 Or InStr(ActiveCell.Value, "ARTICOLO") <> 0 Or InStr(ActiveCell.Value, "Tipo Stampa") <> 0 Or InStr(ActiveCell.Value, "Pagina :") <> 0 Then
           Rows(ActiveCell.Row).Delete
           CONTATORE = CONTATORE + 1
       Else
           Selection.Offset(1, 0).Select
           CONTATORE = 0
       End If
   Loop
   
    'Faccio giusto qualche operazione intermedia
    Columns(1).TextToColumns 'dati text to column
    Range("B1").EntireColumn.HorizontalAlignment = xlLeft
    Range("A3:J3").EntireColumn.HorizontalAlignment = xlCenter
    Columns.AutoFit 'autofit per maggior leggibilita
    'Tolgo gli asterischi dalla colonna DESCRIZIONE
    'Range("B2", Range("B2").End(xlDown)).Replace What:="~*", Replacement:=""
    Range("A1").Select
End Sub

Private Sub storicoPrezzi_cleaner()
   
   Dim CONTATORE As Long
   CONTATORE = 0
    'Pulizia delle righe iniziale manuale
    Range("A1:A6").EntireRow.Delete
    'setto la cella di inizio ciclo
    Range("A2").Select
   
   'Il primo loop fa pulizie di tutte le righe che interrompono la base di dati (righe vuote con trattini o con scritte varie)
   Do Until CONTATORE > 15
       If IsEmpty(ActiveCell) = True Or InStr(ActiveCell.Value, "---") <> 0 Or InStr(ActiveCell.Value, "COPPO 20") <> 0 Or InStr(ActiveCell.Value, "001 Ventimiglia") <> 0 Or InStr(ActiveCell.Value, "ESISTENZA") <> 0 Or InStr(ActiveCell.Value, "VALORE INVENTARIALE") <> 0 Or InStr(ActiveCell.Value, "ARTICOLO") <> 0 Or InStr(ActiveCell.Value, "Tipo Stampa") <> 0 Then
           Rows(ActiveCell.Row).Delete
           CONTATORE = CONTATORE + 1
       Else
           Selection.Offset(1, 0).Select
           CONTATORE = 0
       End If
   Loop
   
   'Faccio il text to column
   ActiveSheet.Range("A:A").TextToColumns Destination:=Range("A1"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 1), Array(9, 1), Array(40, 1), Array(48, 1), Array(55, 1), _
        Array(62, 1), Array(95, 1), Array(112, 1)), TrailingMinusNumbers:=True
        
        
    'Correggo il layout
    With ActiveSheet
        .Range("B1").Value = "DESCRIZIONE"
        .Range("E1").Value = "NR. CLI/FOR"
        .Range("F1").Value = "NOME CLI/FOR"
        .Range("A:H").EntireColumn.AutoFit
        .Range("A1,C1:H1").EntireColumn.HorizontalAlignment = xlCenter
    End With
    
    'imposto i formati
    With ActiveSheet
        .Range("H1").EntireColumn.NumberFormat = "0.00"
        .Range("G1").EntireColumn.NumberFormat = "0"
        .Range("C1").EntireColumn.NumberFormat = "dd/mm/yy"
    End With

    'Correggo i formati data
    Dim cell As Range
    Dim ws As Worksheet
    Set ws = Worksheets("SOUND")

    ' Seleziona l'intervallo di celle che vuoi formattare
    For Each cell In ws.Range("C2", ws.Range("C2").End(xlDown))
        If IsDate(cell.Value) Then
            ' Forza il riconoscimento della data
            cell.Value = CDate(cell.Value)
            ' Applica il formato data desiderato
            cell.NumberFormat = "dd/mm/yyyy"
        End If
    Next cell
    
    'MsgBox "Pulizia effettuata. Selezionare la colonna A e fare un -text.to.columns- facendo attenzione alla colonna FORNITORE che dovrà essere allargata a 95" & vbCrLf & vbCrLf & "Aggiungere inoltre l'header DESCRIZIONE nella cella B2"
    Range("A1").Select
End Sub

Sub storicoContabile_cleaner()
    Dim c As Range
    Dim codiceConto As Long
    Dim inizio As Long
    Dim fine As Long
    Dim firstAddress As String

    Application.ScreenUpdating = False
    
    With ActiveSheet.Range("A:A")
        Set c = .Find("  CODICE   RAGIONE SOCIALE                  INDIRIZZO                        CAP     LOCALITA'                            PR", LookIn:=xlValues)
'        c.Select
        If Not c Is Nothing Then
            firstAddress = c.Address
            Do
                codiceConto = Trim(Mid(c.Offset(1, 0).Value, 4, 5))
                inizio = c.Offset(1, 0).Row
                Set c = .FindNext(c)
                c.Select
                If c.Address <> firstAddress Then
                    fine = c.Offset(-1, 0).Row
                    ActiveSheet.Range("I" & inizio, "I" & fine).Value = codiceConto
                Else
                    fine = ActiveSheet.UsedRange.Rows(ActiveSheet.UsedRange.Rows.Count).Row
                    ActiveSheet.Range("I" & inizio, "I" & fine).Value = codiceConto
                End If
            Loop While c.Address <> firstAddress
        End If
    End With
    
    Dim ws As Worksheet
    Dim searchRange As Range
    Dim foundCell As Range
    Dim searchValue As String
    Dim lastRow As Long
    
    ' Trova l'ultima riga del UsedRange
    lastRow = ActiveSheet.UsedRange.Rows(ActiveSheet.UsedRange.Rows.Count).Row
    ' Imposta l'intervallo di ricerca al UsedRange
    Set searchRange = ActiveSheet.UsedRange
    searchRange.Sort Key1:=ActiveSheet.Range("A1"), Order1:=xlAscending, Header:=xlGuess
    searchValue = "SALDI CONTABILI"

    ' Esegui la ricerca partendo dall'ultima cella del UsedRange verso l'alto
    Set foundCell = searchRange.Find(What:=searchValue, After:=ActiveSheet.Cells(lastRow, searchRange.Columns(searchRange.Columns.Count).Column), _
                                     LookIn:=xlValues, lookAt:=xlPart, SearchOrder:=xlByRows, _
                                     SearchDirection:=xlPrevious, MatchCase:=False)
    
    inizio = foundCell.Row + 1
    searchValue = " DATA REG DATA DOC    DOC   PROT DESCRIZIONE                    DESCRIZIONE                             DARE          AVERE   RIGA"
    Set foundCell = searchRange.Find(What:=searchValue, After:=ActiveSheet.Cells(lastRow, searchRange.Columns(searchRange.Columns.Count).Column), _
                                     LookIn:=xlValues, lookAt:=xlPart, SearchOrder:=xlByRows, _
                                      MatchCase:=False)
    fine = foundCell.Row - 1
    
    Dim rng As Range
    Dim newSheet As Worksheet
    
    Set ws = Worksheets(ActiveSheet.Name)
    
    Set newSheet = Worksheets.Add
    newSheet.Name = "result"
    Set rng = ws.Range("A" & inizio, "I" & fine)
    rng.Copy
    newSheet.Range("A2").PasteSpecial Paste:=xlPasteValues
    Columns("A:A").TextToColumns Destination:=Range("A2"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 1), Array(9, 1), Array(18, 1), Array(26, 1), Array(33, 1), _
        Array(86, 1), Array(109, 1), Array(124, 1)), TrailingMinusNumbers:=True
        
    'Pulizia finale
    With newSheet
        .Range("A1").Value = "DATA REG"
        .Range("B1").Value = "DATA DOC"
        .Range("C1").Value = "DOC"
        .Range("D1").Value = "PROT"
        .Range("E1").Value = "DESCRIZIONE"
        .Range("F1").Value = "DARE"
        .Range("G1").Value = "AVERE"
        .Range("H1").Value = "RIGA"
        .Range("I1").Value = "COD CLI / FOR"
        .Range("A1:I1").Font.Bold = True
        .Range("A1:I1").HorizontalAlignment = xlCenter
        .Range("A1:I1").EntireColumn.AutoFit
        'Larghezza colonne
        .Range("A1:B1").EntireColumn.ColumnWidth = 11.5
        .Range("C1:D1").EntireColumn.ColumnWidth = 7.5
        .Range("E1").EntireColumn.ColumnWidth = 35
        .Range("F1:G1").EntireColumn.ColumnWidth = 18
        .Range("H1").EntireColumn.ColumnWidth = 10
        .Range("I1").EntireColumn.ColumnWidth = 13
        'Horizontal Alignment
        .Range("A1:D1,F1:I1").EntireColumn.HorizontalAlignment = xlCenter
        .Range("E1").EntireColumn.HorizontalAlignment = xlCenter
'        .Range("E1").HorizontalAlignment = xlCenter
        'Formato
        .Range("F1:G1").EntireColumn.NumberFormat = "€#,##0.00"
        '.Range("A1", .Range("I1").End(xlDown)).Sort Key1:=.Range("I1"), Order1:=xlAscending, Header:=xlYes
    End With
    
    'Correggo il formato delle date
    With Worksheets("result")
        .Range("B2", .Range("A100000").End(xlUp)).Select
        ConvertDates .Range("B2", .Range("A100000").End(xlUp))
    End With
    
    'Imposto il corretto ordinamento
    With newSheet
        ' Definisci l'intervallo utilizzato (UsedRange)
        Set rng = .UsedRange
    
        ' Pulisci eventuali criteri di ordinamento precedenti
        .Sort.SortFields.Clear
        
        ' Aggiungi il primo criterio di ordinamento per la colonna I
        .Sort.SortFields.Add Key:=.Range("I1"), Order:=xlAscending
        
        ' Aggiungi il secondo criterio di ordinamento per la colonna A
        .Sort.SortFields.Add Key:=.Range("A1"), Order:=xlAscending
        
        ' Applica l'ordinamento
        With .Sort
            .SetRange rng
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    End With
    
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets("SOUND").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    newSheet.Range("A1").Select
    Application.ScreenUpdating = True
    
End Sub

Private Sub schedaArticoli_cleaner()
    Dim ws As Worksheet
    Dim c As Range, inizio As Range, fine As Range
    Dim firstAddress As String
    Dim riga As Long, ref_count As Long, ultimaRiga As Long
    Dim ref As String
    Dim i As Long
    Set ws = ActiveSheet
    Set c = Range("A:A").Find("ARTICOLO", LookIn:=xlValues)
    Sheets.Add.Name = "result"
    If Not c Is Nothing Then
        ws.Activate
        firstAddress = c.Address
        c.Offset(3, 0).Copy Destination:=Sheets("result").Range("A1")
        Worksheets("result").Range("M1").Value = "SKU CODE"
        Do
            If InStr(c.Offset(5, 0), "-----") = 0 Then
                ref = Left(Trim(c.Offset(1, 0).Value), 6)
                riga = Worksheets("result").Cells.SpecialCells(xlCellTypeLastCell).Row + 1
                Set inizio = c.Offset(5, 0)
                If InStr(c.End(xlDown).Offset(-1, 0).Value, "Totale") <> 0 Then
                    Set fine = c.End(xlDown).Offset(-3, 0)
                    Range(inizio, fine).Copy Destination:=Sheets("result").Range("A" & riga)
                Else
                    Set fine = c.End(xlDown).Offset(-1, 0)
                    Range(inizio, fine).Copy Destination:=Sheets("result").Range("A" & riga)
                End If
                ref_count = fine.Row - inizio.Row
                Worksheets("result").Range("M" & riga, "M" & riga + ref_count).Value = ref
            End If
            Set c = Range("A:A").FindNext(c)
        Loop While c.Address <> firstAddress
    End If
    
    
    With Worksheets("result")
        'Text to column
        .Columns(1).TextToColumns Destination:=Range("A1"), DataType:=xlFixedWidth, _
            FieldInfo:=Array(Array(0, 1), Array(3, 1), Array(12, 1), Array(21, 1), Array(28, 1), _
            Array(45, 1), Array(50, 1), Array(52, 1), Array(59, 1), Array(73, 1), Array(91, 1), Array( _
            108, 1)), TrailingMinusNumbers:=True 'dati text to column
        .Range("F1").Value = "CLI/FOR NUMBER"
        .Range("G1").Value = "CLI/FOR"
        'Elimino le colonne inutili e ordino
        .Range("A1,G1:H1,J1,L1").EntireColumn.Delete
        .Range("H:H").Copy
        .Range("A:A").Insert
        Application.CutCopyMode = False
        .Range("I:I").EntireColumn.Delete
'        .Range("B:C").EntireColumn.NumberFormat = "dd/mm/yy"
        .Range("A:H").EntireColumn.HorizontalAlignment = xlCenter
        'Aggiungo la colonna prezzo unitario
        ultimaRiga = .Range("A1").End(xlDown).Row
        .Range("I1").Value = "PRICE"
        .Range("I2", "I" & ultimaRiga).Formula = "=H2/G2"
        .Range("I1").EntireColumn.HorizontalAlignment = xlCenter
        'imposto i formati
        .Range("H:I").EntireColumn.NumberFormat = "0.00"
        .Range("G1").EntireColumn.NumberFormat = "0"
        .Columns.AutoFit 'autofit per maggior leggibilita
        
    End With
    
    'Correggo i formati data
    Dim cell As Range
    Set ws = Worksheets("result")
    ws.Activate

    ' Seleziona l'intervallo di celle che vuoi formattare
'    For Each cell In ws.Range("B2", ws.Range("C2").End(xlDown))
'        If IsDate(cell.Value) Then
'
'            ' Forza il riconoscimento della data
'            cell.Value = CDate(cell.Value)
'            ' Applica il formato data desiderato
'            cell.NumberFormat = "DD/MM/YYYY"
'        End If
'    Next cell
    
    Dim lastRow As Long
    Dim parts() As String
    Dim dt As Date
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    For Each cell In ws.Range("B2:c" & lastRow) ' Modifica B1:B in base alla colonna delle date
        If IsDate(cell.Value) Then
            ' Controlla il formato della data basato sulla lunghezza dell'anno
            parts = Split(cell.Value, "/")
            If Len(parts(2)) = 4 Then
                ' Formato americano (MM/DD/YYYY)
                dt = DateSerial(CInt(parts(2)), CInt(parts(0)), CInt(parts(1)))
            Else
                ' Formato europeo (DD/MM/YY)
                dt = DateSerial(CInt("20" & parts(2)), CInt(parts(1)), CInt(parts(0)))
            End If
            ' Converte la cella nel formato data corretto
            cell.Value = dt
            ' Applica il formato data desiderato
            cell.NumberFormat = "DD/MM/YYYY"
        End If
    Next cell
    
    'cancello lo sheet SOUND
    Application.DisplayAlerts = False
    Worksheets("SOUND").Delete
End Sub

Sub elencoFornitori_cleaner()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    Set ws = ActiveWorkbook.ActiveSheet
    lastRow = ws.UsedRange.Rows(ws.UsedRange.Rows.Count).Row
    
    ws.Range("A1:A8").EntireRow.Delete
    
    For i = lastRow To 1 Step -1
        ws.Cells(i, 1).Select
        If IsEmpty(ws.Cells(i, 1)) = True Or InStr(ws.Cells(i, 1).Value, "---") <> 0 Or InStr(ws.Cells(i, 1).Value, "===") <> 0 Or InStr(ws.Cells(i, 1).Value, "Pagina :") <> 0 Or InStr(ws.Cells(i, 1).Value, "E L E N C O    F O R N I T O R I") <> 0 Or InStr(ws.Cells(i, 1).Value, "CODICE   RAGIONE SOCIALE") <> 0 Then
            ws.Rows(i).EntireRow.Delete
        End If
    Next i
    
    ws.Columns(1).TextToColumns Destination:=Range("A1"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 1), Array(7, 1), Array(40, 1), Array(75, 1), Array(110, 1), _
        Array(120, 1)), TrailingMinusNumbers:=True
    
    lastRow = ws.UsedRange.Rows(ws.UsedRange.Rows.Count).Row
    
    For i = lastRow To 1 Step -1
        ws.Cells(i, 1).Select
        If IsEmpty(ws.Cells(i, 1)) = True Then
            ws.Rows(i).EntireRow.Delete
        End If
    Next i
    
    ws.Range("A1").EntireRow.Insert
    With ws
        .Range("A1").Value = "COD. FORNITORE"
        .Range("B1").Value = "RAG. SOCIALE"
        .Range("C1").Value = "INDIRIZZO"
        .Range("D1").Value = "LOCALITA'"
        .Range("E1").Value = "CAP"
        .Range("F1").Value = "PROVINCIA"
        .Range("A1:F1").Font.Bold = True
        .Range("A1:F1").HorizontalAlignment = xlCenter
        .Range("A1:F1").EntireColumn.AutoFit
        .Range("A1,E1,F1").EntireColumn.HorizontalAlignment = xlCenter
    End With
    
    Dim c As Range
    Set c = ws.Range("A:A").Find(5555, LookIn:=xlValues, lookAt:=xlWhole)
    If Not c Is Nothing Then
        c.EntireRow.Delete
    End If
    Set c = ws.Range("A:A").Find(6666, LookIn:=xlValues, lookAt:=xlWhole)
    If Not c Is Nothing Then
        c.EntireRow.Delete
    End If
    Set c = ws.Range("A:A").Find(7777, LookIn:=xlValues, lookAt:=xlWhole)
    If Not c Is Nothing Then
        c.EntireRow.Delete
    End If
    Set c = ws.Range("A:A").Find(8888, LookIn:=xlValues, lookAt:=xlWhole)
    If Not c Is Nothing Then
        c.EntireRow.Delete
    End If
    Set c = ws.Range("A:A").Find(9999, LookIn:=xlValues, lookAt:=xlWhole)
    If Not c Is Nothing Then
        c.EntireRow.Delete
    End If
End Sub

Sub elencoClienti_cleaner()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    Set ws = ActiveWorkbook.ActiveSheet
    lastRow = ws.UsedRange.Rows(ws.UsedRange.Rows.Count).Row
    
    ws.Range("A1:A8").EntireRow.Delete
    
    For i = lastRow To 1 Step -1
        ws.Cells(i, 1).Select
        If IsEmpty(ws.Cells(i, 1)) = True Or InStr(ws.Cells(i, 1).Value, "---") <> 0 Or InStr(ws.Cells(i, 1).Value, "===") <> 0 Or InStr(ws.Cells(i, 1).Value, "Pagina :") <> 0 Or InStr(ws.Cells(i, 1).Value, "E L E N C O    C L I E N T I") <> 0 Or InStr(ws.Cells(i, 1).Value, "CODICE   RAGIONE SOCIALE") <> 0 Then
            ws.Rows(i).EntireRow.Delete
        End If
    Next i
    
    ws.Columns(1).TextToColumns Destination:=Range("A1"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 1), Array(7, 1), Array(40, 1), Array(75, 1), Array(110, 1), _
        Array(120, 1)), TrailingMinusNumbers:=True
    
    lastRow = ws.UsedRange.Rows(ws.UsedRange.Rows.Count).Row
    
    For i = lastRow To 1 Step -1
        ws.Cells(i, 1).Select
        If IsEmpty(ws.Cells(i, 1)) = True Then
            ws.Rows(i).EntireRow.Delete
        End If
    Next i
    
    ws.Range("A1").EntireRow.Insert
    With ws
        .Range("A1").Value = "COD. FORNITORE"
        .Range("B1").Value = "RAG. SOCIALE"
        .Range("C1").Value = "INDIRIZZO"
        .Range("D1").Value = "LOCALITA'"
        .Range("E1").Value = "CAP"
        .Range("F1").Value = "PROVINCIA"
        .Range("A1:F1").Font.Bold = True
        .Range("A1:F1").HorizontalAlignment = xlCenter
        .Range("A1:F1").EntireColumn.AutoFit
        .Range("A1,E1,F1").EntireColumn.HorizontalAlignment = xlCenter
    End With
    
    Dim c As Range
    Set c = ws.Range("A:A").Find(5555, LookIn:=xlValues, lookAt:=xlWhole)
    If Not c Is Nothing Then
        c.EntireRow.Delete
    End If
    Set c = ws.Range("A:A").Find(6666, LookIn:=xlValues, lookAt:=xlWhole)
    If Not c Is Nothing Then
        c.EntireRow.Delete
    End If
    Set c = ws.Range("A:A").Find(8888, LookIn:=xlValues, lookAt:=xlWhole)
    If Not c Is Nothing Then
        c.EntireRow.Delete
    End If
    Set c = ws.Range("A:A").Find(99999, LookIn:=xlValues, lookAt:=xlWhole)
    If Not c Is Nothing Then
        c.EntireRow.Delete
    End If
End Sub

Sub ConvertDates(rng As Range)
    Dim lastRow As Long
    Dim parts() As String
    Dim dt As Date
    Dim cell As Range
    Dim ws As Worksheet
    
    lastRow = rng.End(xlDown).Row
    Set ws = Worksheets("result")
    For Each cell In rng
'        cell.Select
        If IsDate(cell.Value) Then
            ' Controlla il formato della data basato sulla lunghezza dell'anno
            parts = Split(cell.Value, "/")
            If Len(parts(2)) = 4 Then
                ' Formato americano (MM/DD/YYYY)
                dt = DateSerial(CInt(parts(2)), CInt(parts(0)), CInt(parts(1)))
            Else
                ' Formato europeo (DD/MM/YY)
                dt = DateSerial(CInt("20" & parts(2)), CInt(parts(1)), CInt(parts(0)))
            End If
            ' Converte la cella nel formato data corretto
            cell.Value = dt
            ' Applica il formato data desiderato
            cell.NumberFormat = "DD/MM/YYYY"
        End If
    Next cell
End Sub
