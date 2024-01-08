
Imports System.Diagnostics.Metrics
Imports System.IO
Imports System.Net
Imports System.Reflection
Imports System.Runtime.CompilerServices
Imports System.Runtime.Intrinsics
Imports System.Windows.Forms.AxHost
Imports Microsoft.VisualStudio.Shell.Interop
Imports System.Globalization


Public Class Form1

    'Public DefaultShortPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "MailMerge")
    '  Public DefaultShortConfigPath = Path.Combine(DefaultShortPath, "Config")
    Public DefaultLongPath = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData) & Path.DirectorySeparatorChar & "MailMerge"


    Dim fReader As StreamReader
    Dim fWriter As StreamWriter
    Public Letter_File_Path As String = ""
    Dim Record_Field() As String

    Public LetterMrg As String = ""
    Public MergedLetter As String = ""
    Public Letter_Text As String = ""
    Public MemName As String = "Barry Campbell"
    Public eMail As String = "barryTNU@gmail.com"
    Public YtSailNr As String = "T1527"
    Public TodaysDate As Date = Today
    Public SeriesWinner As String = "Symphony"
    Public WinnersPoints As String = "152"
    Public SeriesPoints As String = "57"
    Public SeriesPlace As String = "1"
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
    Public NumberOfTypes As String = "4"
    Public TypePosition As String = "6"
    Public TypePoints As String = "58"
    Public YachtsInRegatta As Integer = 32



    Public File_Path As String = DefaultLongPath + Path.DirectorySeparatorChar ' This is the default file path  accessed by all modules
    Public DocumentFilePath = File_Path & "Documents\"
    Public LetterFilePath = File_Path & "Documents\Letters\"
    Public MergeFilePath = File_Path & "Documents\Merge\"
    Public InString As InterpolatedStringHandlerArgumentAttribute

    ReadOnly ColourBlack = System.Drawing.Color.Black
    ReadOnly ColourRed = System.Drawing.Color.Red
    ReadOnly ColourGreen = System.Drawing.Color.LawnGreen
    ReadOnly ColourAmber = System.Drawing.Color.Orange
    ReadOnly ColourWhite = System.Drawing.Color.White
    Public Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load 'Handles MyBase.Load
        'set up the form
        Me.Text = "Mail Merge"
        LoadListBox()
        set_RTB_Indent(10)

        'Create all these folders if they don't exist
        CheckFolderExists(File_Path)
        CheckFolderExists(DocumentFilePath)
        CheckFolderExists(LetterFilePath)
        CheckFolderExists(MergeFilePath)

        Btn_Test.BackColor = ColourAmber
        Btn_Merge.BackColor = ColourGreen
        FileToolStripMenuItem.BackColor = ColourGreen
        LoadToolStripMenuItem1.BackColor = ColourGreen
        MergeToolStripMenuItem1.BackColor = ColourGreen
        SaveToolStripMenuItem1.BackColor = ColourGreen
        Btn_Merge.Visible = False
        ' MergeToolStripMenuItem.Visible = False
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
        CheckFolderExists(fPath)
        Dim StringToMerge As String = TextBox.Text
        PrintToolStripMenuItem.Visible = True
        PrintToolStripMenuItem1.Visible = True
        LetterMrg = ParseDoc(StringToMerge) ' Convert the MailMerge string to a concatenated string ({ and } is replaced with " & ")
        ' MessageBox.Show(LetterMrg)

        '' This is the Play Room
        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX



        Dim s2 = $"{YtName}  Is  a   {YtType}"
        MsgBox(s2)





        MsgBox(LetterMrg)
        TextBox.Text = s2






        '   fWriter = New StreamWriter(fPath & "MergedLetter.txt")
        '  fWriter.Write(LetterMrg)
        '   fWriter.Close()

        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

    End Sub
    Private Function StringBuilderTest(StringToSplit) As String
        Dim builder As New System.Text.StringBuilder
        Dim NewString As String = ""
        Record_Field = Split(StringToSplit, "&")
        Dim l As Integer = UBound(Record_Field) - 1
        For i = 0 To l
            builder.Append(Record_Field(i))
        Next


        Return builder.ToString
    End Function



    Function ParseDoc(StrLetter As String)
        Dim LetterText As String = ""
        StrLetter = ChrW(34) & StrLetter : StrLetter += ChrW(34) ' Add quotes to the beginning and end of the letter
        LetterText = Replace(StrLetter, ChrW(123), ChrW(32) & ChrW(34) & " $ " & ChrW(32), 1) 'Replace {
        StrLetter = Replace(LetterText, ChrW(125), ChrW(32) & " $ " & ChrW(34) & ChrW(32), 1) ' Replace }
        ParseDoc = StrLetter
    End Function



    ' KEEP OUT OF HERE

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


    Sub Testfile(sender As Object, e As EventArgs)


        Dim l As Integer = InStr(MemName, " ")
        Me.XtianName = Mid(MemName, 1, l)


        Dim lett2 = "Dear " & XtianName & "

Thank you for participating in our " & CurrentRegatta & " regatta.

Here are the results of this regatta, and your individual results.

The overall winner of the Regatta was " & SeriesWinner & " with " & WinnersPoints & " Points.

Your yacht " & YtName & " was placed " & SeriesPlace & " with " & SeriesPoints & " points.

Your yacht was also entered in the " & YtClass & " Class, and was competing with " & NrInClass & " other yachts in that class, and also  " & NumberOfTypes & " similar yachts in the " & YtType & " category.

The winning " & YtClass & " was " & ClassWinner & ", and your position in this class was " & ClassPosition & ", with " & ClassPoints & " points.

The winning " & YtType & " was " & TypeWinner & ", and your position was " & TypePosition & " with " & TypePoints & " Points.

Thank you again for supporting our club. We look forward to seeing you at our next regatta. 

Yours sincerely


Commodore"

        MsgBox(lett2)

    End Sub

    Private Sub Lst_MailMerge_Click(sender As Object, e As EventArgs) Handles Lst_MailMerge.Click 'Handles Lst_MailMerge.Click
        My.Computer.Clipboard.SetText(Lst_MailMerge.SelectedItem, TextDataFormat.Text) 'Copy to Clipboard
    End Sub

    Private Sub TextBox_Click(sender As Object, e As EventArgs) Handles TextBox.Click 'Handles TextBox.Click
        TextBox.Paste() 'paste from clipboard
        My.Computer.Clipboard.Clear()

    End Sub

    Private Sub Btn_Test_Click(sender As Object, e As EventArgs) Handles Btn_Test.Click 'Handles Btn_Test.Click
        Testfile(sender, e)
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
