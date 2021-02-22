'---------------------------------------------------------------------------------------
' Tutorial creato per: https://www.forumexcel.it)
'---------------------------------------------------------------------------------------
' Module    : frmMain + modCompilaWord
' Author    : giuliano
' Date      : 11/03/2019 - 27/03/2020
' Purpose   : Questo tutorial mostra come
'             1. acquisire dati da un database Excel
'             2. usarli per compilare un file Word che contiene:
'                a) dei segnaposto, indicati con un doppio segno % (esempio: %%COGNOME%%)
'                b) un controllo TextBox
'                c) una tabella
'             3. salvare la copia del file in formato normale, oppure PDF
'                in una cartella di output scelta dall'utente
'             4. Aprire la cartella la termine dell'elaborazione
'---------------------------------------------------------------------------------------
Option Explicit

Dim sFileWord As String             ' è un semplice file di Word usato come modello
Dim sFileDatabaseExcel As String    ' file Excel che contiene il database con i dati
Dim sCartellaOutput As String       ' cartella in cui verranno salvati i nuovi file

Const wdReplaceAll = 2              '

Private Sub UserForm_Initialize()
    Rem Per comodità imposto già le variabili che mi servono
    sFileWord = ThisWorkbook.Path & "\lettera.docx"
    lblFileWord.Caption = sFileWord
    sFileDatabaseExcel = ThisWorkbook.Path & "\database.xlsx"
    lblFileExcel.Caption = sFileDatabaseExcel
    sCartellaOutput = ThisWorkbook.Path & "\Output\"
    lblCartellaOutput.Caption = sCartellaOutput

    MsgBox "I percorsi caricati sono riferiti a questo tutorial." & vbCrLf & vbCrLf & "Per non caricarli automaticamente, eliminare il codice nell'evento UserForm_Initialize.", vbInformation
End Sub
Private Sub cmdEsegui_Click()
    Dim s As String
    s = "La procedura compilerà il file di Word con i dati del file Excel," & vbCrLf
    s = s & "e creerà un nuovo file per ogni riga del database." & vbCrLf & vbCrLf
    s = s & "Eseguire la procedura?"
    If MsgBox(s, vbQuestion + vbYesNo) = vbNo Then Exit Sub

    MousePointer = fmMousePointerHourGlass
    CompilaSalvaFileConDati
    MousePointer = fmMousePointerDefault
End Sub

Private Sub cmdSelezionaCartellaOutput_Click()
    Dim tmp As String
    tmp = SelezionaCartellaOutput()
    If tmp > vbNullString Then
        sCartellaOutput = tmp
        lblCartellaOutput.Caption = sCartellaOutput
    End If
End Sub

Private Sub cmdSelezionaDatabaseExcel_Click()
    Dim tmp As String
    tmp = SelezionaDatabaseExcel()
    If tmp > vbNullString Then
        sFileDatabaseExcel = tmp
        lblFileExcel.Caption = sFileDatabaseExcel
    End If
End Sub

Private Sub cmdSelezionaFileWord_Click()
    Dim tmp As String
    tmp = SelezionaFileWord()
    If tmp > vbNullString Then
        sFileWord = tmp
        lblFileWord.Caption = sFileWord
    End If
End Sub

Public Sub ApriURL(ByVal pPercorso As String)
    Rem ------------------------------------------------------------
    Rem Routine generica che apre qualsiasi file, cartella, percorso web
    Rem usando l'applicazione predefinita nel computer in uso
    Rem ------------------------------------------------------------
    Const SW_SHOWNORMAL As Long = 1
    Rem Se è un percorso web non posso controllarlo con Dir$()
    If Left$(pPercorso, 2) <> "//" Then
        If Dir$(pPercorso) = vbNullString Then Exit Sub
    End If

    Dim objShell As Object
    Set objShell = CreateObject("shell.application")
    objShell.ShellExecute pPercorso, "", "", "open", SW_SHOWNORMAL
    Set objShell = Nothing
End Sub

Public Sub CompilaSalvaFileConDati()

    Dim i As Long, riga As Long, ur As Long, x As Long, y As Long
    Dim sCodice As String
    Dim sCognome As String
    Dim sNome As String
    Dim sPratica As String
    Dim sIndirizzo As String
    Dim sFileOutput As String

    Dim wb As Workbook
    Dim sh As Worksheet

    Dim xWord As Word.Application ' L'applicazione Word
    Dim xTabella As Word.Table         ' Oggetto Tabella
    Dim xRange As Word.Range           ' Oggetto Range
    'Dim xSelection As Word.Find        ' Oggetto Find
    'Dim xCella As Word.Cell            ' Oggetto Cella

    On Error GoTo CompilaSalvaFileConDati_Error

    Rem ------------------------------------------------------------
    Rem verifico che l'utente abbia indicato i parametri corretti
    If Left$(lblFileWord.Caption, 1) = "<" Then
        MsgBox "Selezionare il file di Word"
        Exit Sub
    End If
    If Left$(lblFileExcel.Caption, 1) = "<" Then
        MsgBox "Selezionare il file Excel"
        Exit Sub
    End If
    If Left$(lblCartellaOutput.Caption, 1) = "<" Then
        MsgBox "Selezionare la cartella di output"
        Exit Sub
    End If

    Rem ------------------------------------------------------------
    Rem Controllo l'esistenza dei file
    If Dir$(lblFileWord.Caption) = vbNullString Then
        MsgBox "Selezionare il file di Word"
        Exit Sub
    End If
    If Dir$(lblFileExcel.Caption) = vbNullString Then
        MsgBox "Selezionare il file Excel"
        Exit Sub
    End If
    If Dir$(lblCartellaOutput.Caption, vbDirectory) = vbNullString Then
        MsgBox "Selezionare la cartella di output"
        Exit Sub
    End If


    Rem ------------------------------------------------------------
    Rem Apro il file che contiene i dati in formato 'tabellare'
    Rem nel primo foglio
    Set wb = Workbooks.Open(sFileDatabaseExcel)
    wb.Windows(1).visible = False ' lo nascondo
    Set sh = wb.Sheets(1)
    ur = sh.Cells(sh.Rows.Count, 1).End(xlUp).Row ' ricavo ultima riga


    Rem ------------------------------------------------------------
    Rem Servono per mostrare l'avanzamento dell'elaborazione
    Dim counter As Long
    Dim countermax As Long
    counter = 0
    countermax = ur - riga - 1
    Rem ------------------------------------------------------------


    Rem ------------------------------------------------------------
    riga = 2 ' riga in cui iniziano i dati del database Excel
    y = 1    ' usato per indicare la riga della tabella di Word
    For i = riga To ur
        Rem ------------------------------------------------------------
        Rem  mostra il numero di file in elaborazione
        counter = counter + 1
        lblProgress.Caption = counter & " di " & countermax
        DoEvents

        Rem ------------------------------------------------------------
        Rem Per ogni riga del database compilo il modello
        Rem ------------------------------------------------------------

        Rem Apro un nuovo documento Word e lo nascondo
        Set xWord = New Word.Application
        If chkMostraWord.Value = True Then
            xWord.visible = True ' così vedo cosa accade
        End If

        Rem Apro il file Word, che mi serve da modello, tanto non verrà
        Rem modificato, perché andrà poi salvato con il nome del cliente
        xWord.Documents.Open sFileWord


        Rem ------------------------------------------------------------
        Rem Leggo i dati del cliente
        sCodice = sh.Cells(i, 1)
        sCognome = sh.Cells(i, 2)
        sNome = sh.Cells(i, 3)
        sIndirizzo = sh.Cells(i, 4)
        sPratica = sh.Cells(i, 5)

        Rem e li scrivo nell'intestazione
        Set xRange = xWord.ActiveDocument.Range
        xRange.Find.Execute "%%COGNOME%%", , , , , , , , , sCognome, wdReplaceAll
        Set xRange = xWord.ActiveDocument.Range
        xRange.Find.Execute "%%NOME%%", , , , , , , , , sNome, wdReplaceAll
        Set xRange = xWord.ActiveDocument.Range
        xRange.Find.Execute "%%CODICE%%", , , , , , , , , sCodice, wdReplaceAll
        Set xRange = xWord.ActiveDocument.Range
        xRange.Find.Execute "%%INDIRIZZO%%", , , , , , , , , sIndirizzo, wdReplaceAll


        Rem ------------------------------------------------------------
        Rem Sostituisco il titolo nel textbox, se indicato
        If lblTitolo.Caption > vbNullString Then
            xWord.ActiveDocument.Shapes.Item(1).TextFrame.TextRange = txtTitolo.Text
        End If


        Rem ------------------------------------------------------------
        Rem Aggiorno intestazioni nella tabella con i nomi delle colonne prese dal database
        Rem Diamo per scontato che vi sia una sola tabella
        Set xTabella = xWord.ActiveDocument.Tables(1)
        For x = 1 To 5
            xTabella.Rows(y).Cells(x).Range.Text = sh.Cells(y, x + 5)
        Next x


        Rem ------------------------------------------------------------
        ' aggiusto la larghezza della prima e ultima colonna
        xTabella.Rows(1).Cells(1).Width = 60
        xTabella.Rows(2).Cells(1).Width = 60
        xTabella.Rows(1).Cells(2).Width = 120
        xTabella.Rows(2).Cells(2).Width = 120
        xTabella.Rows(1).Cells(5).Width = 95
        xTabella.Rows(2).Cells(5).Width = 95


        Rem ------------------------------------------------------------
        Rem Aggiorno la tabella
        xTabella.Rows(y + 1).Cells(1).Range.Text = sh.Cells(i, 6)
        xTabella.Rows(y + 1).Cells(2).Range.Text = sh.Cells(i, 7)
        xTabella.Rows(y + 1).Cells(3).Range.Text = sh.Cells(i, 8)
        xTabella.Rows(y + 1).Cells(4).Range.Text = sh.Cells(i, 9)
        xTabella.Rows(y + 1).Cells(5).Range.Text = sh.Cells(i, 10)


        Rem ------------------------------------------------------------
        Rem Verifico che il percorso di output sia corretto
        If Right$(sCartellaOutput, 1) <> "\" Then
            sCartellaOutput = sCartellaOutput & "\"
        End If


        Rem ------------------------------------------------------------
        Rem Preparo il nome del file da salvare (senza estensione)
        sFileOutput = sCognome & " " & sNome & "-" & sPratica
        Rem e salvo il file aggiungendo l'estensione richiesta
        If chkSalvaPDF.Value = True Then
            sFileOutput = sFileOutput & ".pdf"
            xWord.ActiveDocument.SaveAs sCartellaOutput & sFileOutput, wdFormatPDF
        Else
            sFileOutput = sFileOutput & ".docx"
            xWord.ActiveDocument.SaveAs sCartellaOutput & sFileOutput, wdFormatDocumentDefault
        End If
        Rem ------------------------------------------------------------


        xWord.DisplayAlerts = False
        xWord.ActiveDocument.Saved = True
        xWord.ActiveDocument.Close wdDoNotSaveChanges
        'xWord.DisplayAlerts = True
        xWord.Quit
        Set xWord = Nothing

    Next i

    lblProgress.Caption = "Finito"
    DoEvents
    MsgBox "Elaborazione terminata.", vbInformation

    Rem ------------------------------------------------------------
    Rem Se richiesto, apro la cartella di output
    If chkApriOutput.Value = True Then
        ApriURL sCartellaOutput
    End If

    On Error GoTo 0
    Exit Sub

CompilaSalvaFileConDati_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CompilaSalvaFileConDati of Form frmMain"
    Stop        ' si ferma momentaneamente...
    Resume      ' ...e fa vedere all'utente la riga che ha generato l'errore.
End Sub

Public Function SelezionaFileWord() As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .InitialFileName = ThisWorkbook.Path
        .AllowMultiSelect = False
        .Filters.Add "Modelli Word", "*.doc; *.docx", 1
        .Show
        If .SelectedItems.Count = 0 Then Exit Function
        SelezionaFileWord = .SelectedItems(1)
    End With
End Function

Public Function SelezionaDatabaseExcel() As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .InitialFileName = ThisWorkbook.Path
        .AllowMultiSelect = False
        .Filters.Add "File Excel", "*.xls; *.xlsx", 1
        .Show
        If .SelectedItems.Count = 0 Then Exit Function
        SelezionaDatabaseExcel = .SelectedItems(1)
    End With
End Function

Public Function SelezionaCartellaOutput() As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .InitialFileName = ThisWorkbook.Path
        .AllowMultiSelect = False
        '.Filters.Add "File Excel", "*.xls; *.xlsx", 1
        .Show
        If .SelectedItems.Count = 0 Then Exit Function
        SelezionaCartellaOutput = .SelectedItems(1)
    End With
End Function
