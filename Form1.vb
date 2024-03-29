﻿

Imports System.IO
Imports System.Net
Imports System.Reflection
Imports System.Reflection.Metadata
Imports System.Runtime.CompilerServices
Imports System.Runtime.Intrinsics
Imports Windows.Win32.System
Imports System.Drawing.Printing
Imports System.Text
Imports System
Imports GemBox.Document
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.TaskbarClock

' Backed up 13 Jan 10am

'Dion   I'm not sure if this is the way to go, but I've installed the GemBox library, which includes Mail Merge.
'Having difficulty coming to grips with it,  but so  far it is the best option I've come across.
'What I want to achieve is to write a letter, inserting the Mail merge fields, then use this template to send this letter to
'all, or selected entrants in the regatta.
'I thought this would be a no brainer.  Can't believe it's so difficult.
'Hoping that with your wealth of experience you can crack it.
'Wondering if HTML   might be the way to go. Seems to be a few mail merge apps using this.

Public Class Form1

    'Public DefaultShortPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "MailMerge")
    '  Public DefaultShortConfigPath = Path.Combine(DefaultShortPath, "Config")
    Public DefaultLongPath = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData) & Path.DirectorySeparatorChar & "MailMerge"


    Dim fReader As StreamReader
    Dim fWriter As StreamWriter
    Public Letter_File_Path As String = ""
    Dim Record_Field() As String

    Dim Test_Letter As String = ""


    'These variable fields are used by the YRM program.
    'They are set by iterating through each yacht in the RegattaPlaces ListView box, which has the results of every race,
    'and also the RegAddresses.csv file, which has the competors names and addresses
    'This is just dummy data for testing.

    Public LetterMrg As String = ""
    Public MergedLetter As String = ""
    Public Letter_Text As String = ""
    Public MemName As String = "Barry Campbell"
    Public eMail As String = "barryTNU@gmail.com"
    Public YtSailNr As String = "T1527"
    Public TodaysDate As Date = Today
    Public SeriesWinner As String = "Symphony"
    Public WinnersPoints As String = "152"
    Public SeriesPoints As String = "164"
    Public SeriesPlace As String = "3"
    Public YtOwner As String = "Barry Campbell"
    Public XtianName As String = "Barry"
    Public CurrentRegatta As String = "EASTER2023"
    Public YtName As String = "Jen"
    Public YtClass As String = "Noelex"
    Public ClassWinner As String = "The Gallon"
    Public ClassPosition As String = "3"
    Public NrInClass As String = "14"
    Public ClassPoints As String = "87"
    Public Type As String = "Trailer Yacht"
    Public TypeWinner As String = "Mr Judge"
    Public YtType As String = "Trailer Yacht"
    Public NumberOfTypes As String = "26"
    Public TypePosition As String = "6"
    Public TypePoints As String = "58"
    Public YachtsInRegatta As String = "32"

    Public dataSource ' The variables used in the merge

    Public File_Path As String = DefaultLongPath + Path.DirectorySeparatorChar ' This is the default file path  accessed by all modules
    Public DocumentFilePath = File_Path & "Documents\"
    Public LetterFilePath = File_Path & "Documents\Letters\"
    Public MergeFilePath = File_Path & "Documents\Merge\"

    ReadOnly ColourBlack = System.Drawing.Color.Black
    ReadOnly ColourRed = System.Drawing.Color.Red
    ReadOnly ColourGreen = System.Drawing.Color.LawnGreen
    ReadOnly ColourAmber = System.Drawing.Color.Orange
    ReadOnly ColourWhite = System.Drawing.Color.White
    Public Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load 'Handles MyBase.Load

        'Required by Gembox library
        'If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        'set up the form
        Me.Text = "Mail Merge"
        LoadListBox()
        set_RTB_Indent(10)

        'Create all these folders if they don't exist

        CheckFolderExists(File_Path)
        CheckFolderExists(DocumentFilePath)
        CheckFolderExists(LetterFilePath)
        CheckFolderExists(MergeFilePath)

        ' Btn_Test.Visible = False
        Btn_Test.BackColor = ColourAmber
        Btn_Merge.BackColor = ColourGreen
        FileToolStripMenuItem.BackColor = ColourGreen
        LoadToolStripMenuItem1.BackColor = ColourGreen
        MergeToolStripMenuItem1.BackColor = ColourGreen
        SaveToolStripMenuItem1.BackColor = ColourGreen
        ' Btn_Merge.Visible = False
        MergeToolStripMenuItem1.Visible = False
        PrintToolStripMenuItem.Visible = False
        PrintToolStripMenuItem1.Visible = False
        SaveToolStripMenuItem.Visible = False
        SaveToolStripMenuItem1.Visible = False
        SaveAsToolStripMenuItem.Visible = False



    End Sub
    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

    Private Sub MergeToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles MergeToolStripMenuItem1.Click, Btn_Merge.Click
        Dim fPath As String = MergeFilePath
        Dim StringToMerge As String = TextBox.Text
        Dim doc = New DocumentModel()
        Dim document As New DocumentModel()


        Test_Letter = TextBox.Text

        ' Test_Letter = ParseDoc(TextBox.Text)


        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
        '' This is the Play Room
        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
        ' GoTo DateFormat '(Trivial, but it Works)
        'GoTo MergeTestLetter' (Think I had this sort of working, but not any longer. Don't know what happened.)
        GoTo MergefromTextBox '(Still playing with this.  Ideally this is what I want to happen)


        Exit Sub

        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
MergefromTextBox:
        Test_Letter = TextBox.Text
        doc = New DocumentModel()

        ' Add a new document content.

        Dim MergeTemplate As String = ParseDoc(Test_Letter)
        ' Dim MergeTemplate As String = Test_Letter

        doc.Sections.Add(New Section(doc, New Paragraph(doc, MergeTemplate)))


        dataSource = New With
            {.TodaysDate = Today,
        .SeriesWinner = "Symphony",
        .WinnersPoints = "152",
        .SeriesPoints = "164",
        .SeriesPlace = "3",
        .YtOwner = "Barry Campbell",
        .XtianName = "Barry",
        .CurrentRegatta = "EASTER2023",
        .YtName = "Jen",
        .YtClass = "Noelex",
        .ClassWinner = "The Gallon",
        .ClassPosition = "3",
        .NrInClass = "14",
        .ClassPoints = "87",
        .Type = "Trailer Yacht",
        .TypeWinner = "Mr Judge",
        .YtType = "Trailer Yacht",
        .NumberOfTypes = "26",
        .TypePosition = "6",
        .TypePoints = "58",
        .YachtsInRegatta = "32"
        }

        TextBox.Text = Test_Letter

        ' Execute mail merge.
        doc.MailMerge.Execute(dataSource)

        ' Save to DOCX and PDF files.
        doc.Save(MergeFilePath & "Document.txt")
        doc.Save(MergeFilePath & "Document.docx")
        doc.Save(MergeFilePath & "Document.pdf")

        MsgBox(MergeFilePath & "Document.txt",, "Document saved as ")

        Exit Sub

        'ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ
DateFormat:

        Dim section As New Section(doc)
        doc.Sections.Add(section)

        ' Add '{ AUTHOR }' field.
        section.Blocks.Add(New Paragraph(doc, New Run(doc, "Author: "), New Field(doc, FieldType.Author, Nothing, "Mario at GemBox")))

        ' Add '{ DATE }' field.
        section.Blocks.Add(New Paragraph(doc, New Run(doc, "Date: "), New Field(doc, FieldType.Date)))

        ' Add '{ DATE \@ "dddd, MMMM dd, yyyy" }' field.
        section.Blocks.Add(New Paragraph(doc, New Run(doc, "Date with specified format: "), New Field(doc, FieldType.Date, "\@ ""dddd, MMMM dd, yyyy""")))


        dataSource = New With {.Author = "Barry Campbell"}

        ' Execute mail merge.
        doc.MailMerge.Execute(dataSource)


        ' Save to DOCX and PDF files.
        doc.Save(MergeFilePath & "Document.txt")
        doc.Save(MergeFilePath & "Document.docx")
        doc.Save(MergeFilePath & "Document.pdf")
        MsgBox(MergeFilePath & "Document.txt",, "Document saved as ")


        Exit Sub
        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

MergeTestLetter:

        Test_Letter = "The author is " & MemName

        ' Add document content.
        doc.Sections.Add(New Section(doc, New Paragraph(doc, New Field(doc, FieldType.Author, Test_Letter))))

        ' Save the document to a file.
        doc.Save(LetterFilePath & "TemplateDocument.txt")

        ' Initialize mail merge data source.
        dataSource = New With {.Author = "Barry Campbell"}

        ' Execute mail merge.
        doc.MailMerge.Execute(dataSource)
        ' doc.MailMerge.Execute(data)

        ' Save to DOCX and PDF files.
        doc.Save(MergeFilePath & "Document.txt")
        ' doc.Save(MergeFilePath & "Document.docx")
        ' doc.Save(MergeFilePath & "Document.pdf")

        MsgBox(MergeFilePath & "Document.txt",, "Document saved as ")

        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

        '   fWriter = New StreamWriter(fPath & "MergedLetter.txt")
        '  fWriter.Write(LetterMrg)
        '   fWriter.Close()

    End Sub

    ' ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ
    'All subs below are to do with setting  up the form, Loading and saving, and other housekeeping


    Sub LoadListBox()
        With Lst_MailMerge

            .Items.Add("{TodaysDate}")
            .Items.Add("{MemName}")
            .Items.Add("{YtSailNr}")
            .Items.Add("{YtName}")
            .Items.Add("{YtOwner}")
            .Items.Add("{XtianName}")
            .Items.Add("{eMail}")

            .Items.Add("{Address1}")
            .Items.Add("{Address2}")
            .Items.Add("{Town}")
            .Items.Add("{PostCode}")
            .Items.Add("{State}")
            .Items.Add("{Country}")


            .Items.Add("{YachtsInRegatta}")
            .Items.Add("{WinnersPoints")
            .Items.Add("{SeriesWinner}")
            .Items.Add("{SeriesPoints}")
            .Items.Add("{SeriesPlace}")
            .Items.Add("{CurrentRegatta}")
            .Items.Add("{YtClass}")
            .Items.Add("{YtSkipper}")
            .Items.Add("{ClassWinner}")
            .Items.Add("{ClassPosition}")
            .Items.Add("{NrInClass}")
            .Items.Add("{ClassPoints}")
            .Items.Add("{TypeWinner}")
            .Items.Add("{YtType}")
            .Items.Add("{NumberOfTypes}")
            .Items.Add("{TypePosition}")
            .Items.Add("{TypePoints}")

        End With

        My.Computer.Clipboard.Clear()

    End Sub

    Sub Load_Data()

        dataSource = New With
                   {
        .MemName = "Barry Campbell",
        .eMail = "barryTNU@gmail.com",
        .YtSailNr = "T1527",
        .TodaysDate = Today,
        .SeriesWinner = "Symphony",
        .WinnersPoints = "152",
        .SeriesPoints = "164",
        .SeriesPlace = "3",
        .YtOwner = "Barry Campbell",
        .XtianName = "Barry",
        .CurrentRegatta = "EASTER2023",
        .YtName = "Jen",
        .YtClass = "Noelex",
        .ClassWinner = "The Gallon",
        .ClassPosition = "3",
        .NrInClass = "14",
        .ClassPoints = "87",
        .Type = "Trailer Yacht",
        .TypeWinner = "Mr Judge",
        .YtType = "Trailer Yacht",
        .NumberOfTypes = "26",
        .TypePosition = "6",
        .TypePoints = "58",
        .YachtsInRegatta = "32"
        }


    End Sub
    Function ParseDoc(StrLetter As String)
        Dim LetterText As String = ""
        If InStr(StrLetter, ChrW(123)) = False Then
            MsgBox("Not a Mail Merge Document", , "")
            Return StrLetter
            Exit Function
        End If
        StrLetter = ChrW(34) & StrLetter : StrLetter += ChrW(34) ' Add   quotes to the beginning and end of the letter
        LetterText = Replace(StrLetter, ChrW(123), ChrW(34) & ChrW(32) & ChrW(38) & ChrW(32), 1) 'Replace { with &
        StrLetter = Replace(LetterText, ChrW(125), ChrW(32) & ChrW(38) & ChrW(32) & ChrW(34), 1) ' Replace } with &
        Return StrLetter

        MsgBox(StrLetter)
    End Function

    Sub Testfile() Handles Btn_Test.Click
        ' A dummy letter for testing
        'This merge works,
        'However the same format doesn't work if it's loaded as a file,
        'It only works if it's included in the program like this.
        'Obviously not the desired result.

        Dim l As Integer = InStr(MemName, " ")
        Me.XtianName = Mid(MemName, 1, l)

        Dim Merge_Template As String = "Dear " & XtianName & "

Thank you for participating in our " & CurrentRegatta & " regatta.

Here are the results of this regatta, and your individual results.

The overall winner of the Regatta was " & SeriesWinner & " with " & WinnersPoints & " points.

Your yacht ," & YtName & ", was placed " & SeriesPlace & " with " & SeriesPoints & " points.

" & YtName & " was also entered in the " & YtClass & " class, and was competing with " & NrInClass & " other yachts in that class, and also " & NumberOfTypes & " similar yachts in the " & YtType & "
category.

The winning " & YtClass & " was " & ClassWinner & ", and your position in this Class was " & ClassPosition & ", with " & ClassPoints & " points.

The winning " & YtType & " was " & TypeWinner & ", and your position was " & TypePosition & " with " & TypePoints & " points.

Thank you again for supporting our club . We look forward to seeing you at our next regatta.

Yours sincerely

Commodore."

        TextBox.Text = Merge_Template

    End Sub

    Private Sub Lst_MailMerge_Click(sender As Object, e As EventArgs) Handles Lst_MailMerge.Click 'Handles Lst_MailMerge.Click
        My.Computer.Clipboard.SetText(Lst_MailMerge.SelectedItem, TextDataFormat.Text) 'Copy to Clipboard
    End Sub

    Private Sub TextBox_Click(sender As Object, e As EventArgs) Handles TextBox.Click 'Handles TextBox.Click
        TextBox.Paste() 'paste from clipboard
        My.Computer.Clipboard.Clear()

    End Sub

    Private Sub Btn_Test_Click(sender As Object, e As EventArgs) ' Handles Btn_Test.Click 'Handles Btn_Test.Click
        Testfile()
    End Sub

    Private Sub LoadToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LoadToolStripMenuItem.Click, LoadToolStripMenuItem1.Click 'LoadToolStripMenuItem.Click, LoadToolStripMenuItem1.Click
        Dim fPath = LetterFilePath
        Dim MergeFile As String

        OpenFileDialog1.Title = "Please select a  file"
        OpenFileDialog1.InitialDirectory = fPath
        OpenFileDialog1.Filter = "Text Files(*.txt)|*.txt|Merge Files(*.mrg)|*.mrg|All Files(*.*)|*.*"
        OpenFileDialog1.FileName = ""

        Try
            If OpenFileDialog1.ShowDialog = DialogResult.OK Then
                Letter_File_Path = OpenFileDialog1.FileName
                fReader = New StreamReader(Letter_File_Path)
                Do While fReader.Peek <> -1
                    MergeFile = fReader.ReadToEnd
                Loop
                fReader.Close()
                TextBox.Text = MergeFile

            End If

        Catch ex As Exception
            MsgBox(ex.Message, vbCritical, "File Read Error")
        End Try

    End Sub


    Private Sub SaveAsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveAsToolStripMenuItem.Click

        Dim fPath = File_Path & "Documents\Letters"
        Dim RTB_Text As String = TextBox.Text

        CheckFolderExists(fPath & "\")

        SaveFileDialog1.Title = "Please choose a file path and name"
        SaveFileDialog1.InitialDirectory = fPath
        SaveFileDialog1.Filter = "Text Files(*.txt)|*.txt|Merge Files(*.mrg)|*.mrg|All Files(*.*)|*.*"
        SaveFileDialog1.FileName = Letter_File_Path
        Try
            If SaveFileDialog1.ShowDialog() = DialogResult.OK Then
                Letter_File_Path = SaveFileDialog1.FileName
                fWriter = New StreamWriter(Letter_File_Path)
                fWriter.Write(RTB_Text)
                fWriter.Close()

            End If

        Catch EX As Exception
            MsgBox(EX.Message, vbCritical, "Error writing file")
        End Try

    End Sub



    Private Sub PrintToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles PrintToolStripMenuItem1.Click

    End Sub

    Private Sub SaveToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles SaveToolStripMenuItem.Click, SaveToolStripMenuItem1.Click 'Handles SaveToolStripMenuItem.Click, SaveToolStripMenuItem1.Click
        If Letter_File_Path = "" Then SaveAsToolStripMenuItem_Click(sender, e)

        Dim fPath As String = ""
        Dim RTB_Text As String = TextBox.Text

        Try
            fWriter = New StreamWriter(Letter_File_Path)
            fWriter.Write(RTB_Text)
            fWriter.Close()
        Catch EX As Exception
            MsgBox(EX.Message, vbCritical, "Error writing file")
        End Try
    End Sub



    Sub CheckFolderExists(FolderName)
        Dim letter As String = ""

        For i = Len(FolderName) To 0 Step -1
            letter = Mid(FolderName, i, 1)
            If letter = "\" Then
                FolderName = Mid(FolderName, 1, i - 1)
                Exit For
            End If
        Next
        Dim FolderExists As Boolean = IO.Directory.Exists(FolderName)
        If FolderExists = False Then
            MkDir(FolderName)

        End If
    End Sub
    Sub set_RTB_Indent(indent As Integer)
        'sets the Rich Text Box indent to 10
        With TextBox
            Dim OldSelStart As Integer = .SelectionStart
            Dim OldSelLen As Integer = .SelectionLength
            .SelectionStart = 0
            .SelectionLength = Len(.Text)
            .SelectionIndent = indent
            .SelectionStart = OldSelStart
            .SelectionLength = OldSelLen
        End With
    End Sub

    Private Sub TextBox_TextChanged(sender As Object, e As EventArgs) Handles TextBox.TextChanged 'Handles TextBox.TextChanged
        ' MergeToolStripMenuItem.Visible = True
        MergeToolStripMenuItem1.Visible = True
        SaveToolStripMenuItem.Visible = True
        SaveToolStripMenuItem1.Visible = True
        SaveAsToolStripMenuItem.Visible = True
        Btn_Merge.Visible = True
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click 'Handles ExitToolStripMenuItem.Click
        Me.Dispose()
    End Sub

    Private Sub NewToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NewToolStripMenuItem.Click

    End Sub

    Private Sub LetterToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LetterToolStripMenuItem.Click

    End Sub

    Private Sub MailMergeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MailMergeToolStripMenuItem.Click

    End Sub
End Class
