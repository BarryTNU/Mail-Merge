<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        TextBox = New RichTextBox()
        MenuStrip1 = New MenuStrip()
        FileToolStripMenuItem = New ToolStripMenuItem()
        NewToolStripMenuItem = New ToolStripMenuItem()
        LetterToolStripMenuItem = New ToolStripMenuItem()
        MailMergeToolStripMenuItem = New ToolStripMenuItem()
        LoadToolStripMenuItem = New ToolStripMenuItem()
        SaveAsToolStripMenuItem = New ToolStripMenuItem()
        SaveToolStripMenuItem = New ToolStripMenuItem()
        PrintToolStripMenuItem = New ToolStripMenuItem()
        ExitToolStripMenuItem = New ToolStripMenuItem()
        LoadToolStripMenuItem1 = New ToolStripMenuItem()
        SaveToolStripMenuItem1 = New ToolStripMenuItem()
        MergeToolStripMenuItem1 = New ToolStripMenuItem()
        PrintToolStripMenuItem1 = New ToolStripMenuItem()
        Lst_MailMerge = New ListBox()
        ToolStripContainer1 = New ToolStripContainer()
        Btn_Test = New Button()
        Btn_Merge = New Button()
        OpenFileDialog1 = New OpenFileDialog()
        PrintPreviewControl1 = New PrintPreviewControl()
        PrintDocument1 = New Printing.PrintDocument()
        PageSetupDialog1 = New PageSetupDialog()
        PrintDialog1 = New PrintDialog()
        SaveFileDialog1 = New SaveFileDialog()
        MenuStrip1.SuspendLayout()
        ToolStripContainer1.SuspendLayout()
        SuspendLayout()
        ' 
        ' TextBox
        ' 
        TextBox.BorderStyle = BorderStyle.FixedSingle
        TextBox.Location = New Point(217, 89)
        TextBox.Margin = New Padding(4)
        TextBox.Name = "TextBox"
        TextBox.Size = New Size(1122, 678)
        TextBox.TabIndex = 0
        TextBox.Text = ""
        ' 
        ' MenuStrip1
        ' 
        MenuStrip1.Items.AddRange(New ToolStripItem() {FileToolStripMenuItem, LoadToolStripMenuItem1, SaveToolStripMenuItem1, MergeToolStripMenuItem1, PrintToolStripMenuItem1})
        MenuStrip1.Location = New Point(0, 0)
        MenuStrip1.Name = "MenuStrip1"
        MenuStrip1.Size = New Size(1406, 24)
        MenuStrip1.TabIndex = 3
        MenuStrip1.Text = "MenuStrip1"
        ' 
        ' FileToolStripMenuItem
        ' 
        FileToolStripMenuItem.DropDownItems.AddRange(New ToolStripItem() {NewToolStripMenuItem, LoadToolStripMenuItem, SaveAsToolStripMenuItem, SaveToolStripMenuItem, PrintToolStripMenuItem, ExitToolStripMenuItem})
        FileToolStripMenuItem.Name = "FileToolStripMenuItem"
        FileToolStripMenuItem.Size = New Size(37, 20)
        FileToolStripMenuItem.Text = "File"
        ' 
        ' NewToolStripMenuItem
        ' 
        NewToolStripMenuItem.DropDownItems.AddRange(New ToolStripItem() {LetterToolStripMenuItem, MailMergeToolStripMenuItem})
        NewToolStripMenuItem.Name = "NewToolStripMenuItem"
        NewToolStripMenuItem.Size = New Size(112, 22)
        NewToolStripMenuItem.Text = "New"
        ' 
        ' LetterToolStripMenuItem
        ' 
        LetterToolStripMenuItem.Name = "LetterToolStripMenuItem"
        LetterToolStripMenuItem.Size = New Size(134, 22)
        LetterToolStripMenuItem.Text = "Letter"
        ' 
        ' MailMergeToolStripMenuItem
        ' 
        MailMergeToolStripMenuItem.Name = "MailMergeToolStripMenuItem"
        MailMergeToolStripMenuItem.Size = New Size(134, 22)
        MailMergeToolStripMenuItem.Text = "Mail Merge"
        ' 
        ' LoadToolStripMenuItem
        ' 
        LoadToolStripMenuItem.Name = "LoadToolStripMenuItem"
        LoadToolStripMenuItem.Size = New Size(112, 22)
        LoadToolStripMenuItem.Text = "Load"
        ' 
        ' SaveAsToolStripMenuItem
        ' 
        SaveAsToolStripMenuItem.Name = "SaveAsToolStripMenuItem"
        SaveAsToolStripMenuItem.Size = New Size(112, 22)
        SaveAsToolStripMenuItem.Text = "Save as"
        ' 
        ' SaveToolStripMenuItem
        ' 
        SaveToolStripMenuItem.Name = "SaveToolStripMenuItem"
        SaveToolStripMenuItem.Size = New Size(112, 22)
        SaveToolStripMenuItem.Text = "Save"
        ' 
        ' PrintToolStripMenuItem
        ' 
        PrintToolStripMenuItem.Name = "PrintToolStripMenuItem"
        PrintToolStripMenuItem.Size = New Size(112, 22)
        PrintToolStripMenuItem.Text = "Print"
        ' 
        ' ExitToolStripMenuItem
        ' 
        ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
        ExitToolStripMenuItem.Size = New Size(112, 22)
        ExitToolStripMenuItem.Text = "Exit"
        ' 
        ' LoadToolStripMenuItem1
        ' 
        LoadToolStripMenuItem1.Name = "LoadToolStripMenuItem1"
        LoadToolStripMenuItem1.Size = New Size(45, 20)
        LoadToolStripMenuItem1.Text = "Load"
        ' 
        ' SaveToolStripMenuItem1
        ' 
        SaveToolStripMenuItem1.Name = "SaveToolStripMenuItem1"
        SaveToolStripMenuItem1.Size = New Size(43, 20)
        SaveToolStripMenuItem1.Text = "Save"
        ' 
        ' MergeToolStripMenuItem1
        ' 
        MergeToolStripMenuItem1.Name = "MergeToolStripMenuItem1"
        MergeToolStripMenuItem1.Size = New Size(53, 20)
        MergeToolStripMenuItem1.Text = "Merge"
        ' 
        ' PrintToolStripMenuItem1
        ' 
        PrintToolStripMenuItem1.Name = "PrintToolStripMenuItem1"
        PrintToolStripMenuItem1.Size = New Size(44, 20)
        PrintToolStripMenuItem1.Text = "Print"
        ' 
        ' Lst_MailMerge
        ' 
        Lst_MailMerge.BorderStyle = BorderStyle.FixedSingle
        Lst_MailMerge.FormattingEnabled = True
        Lst_MailMerge.ItemHeight = 21
        Lst_MailMerge.Location = New Point(25, 89)
        Lst_MailMerge.Name = "Lst_MailMerge"
        Lst_MailMerge.Size = New Size(163, 674)
        Lst_MailMerge.Sorted = True
        Lst_MailMerge.TabIndex = 4
        ' 
        ' ToolStripContainer1
        ' 
        ' 
        ' ToolStripContainer1.ContentPanel
        ' 
        ToolStripContainer1.ContentPanel.Size = New Size(8, 0)
        ToolStripContainer1.Location = New Point(426, 25)
        ToolStripContainer1.Name = "ToolStripContainer1"
        ToolStripContainer1.Size = New Size(8, 8)
        ToolStripContainer1.TabIndex = 5
        ToolStripContainer1.Text = "ToolStripContainer1"
        ' 
        ' Btn_Test
        ' 
        Btn_Test.Font = New Font("Segoe UI", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0)
        Btn_Test.Location = New Point(30, 47)
        Btn_Test.Name = "Btn_Test"
        Btn_Test.Size = New Size(49, 25)
        Btn_Test.TabIndex = 6
        Btn_Test.Text = "Test"
        Btn_Test.UseVisualStyleBackColor = True
        ' 
        ' Btn_Merge
        ' 
        Btn_Merge.Font = New Font("Segoe UI", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0)
        Btn_Merge.Location = New Point(121, 47)
        Btn_Merge.Name = "Btn_Merge"
        Btn_Merge.Size = New Size(67, 25)
        Btn_Merge.TabIndex = 7
        Btn_Merge.Text = "Merge"
        Btn_Merge.UseVisualStyleBackColor = True
        ' 
        ' OpenFileDialog1
        ' 
        OpenFileDialog1.FileName = "OpenFileDialog1"
        ' 
        ' PrintPreviewControl1
        ' 
        PrintPreviewControl1.Location = New Point(410, 25)
        PrintPreviewControl1.Name = "PrintPreviewControl1"
        PrintPreviewControl1.Size = New Size(530, 32)
        PrintPreviewControl1.TabIndex = 8
        ' 
        ' PrintDialog1
        ' 
        PrintDialog1.UseEXDialog = True
        ' 
        ' Form1
        ' 
        AutoScaleDimensions = New SizeF(10F, 21F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(1406, 812)
        Controls.Add(PrintPreviewControl1)
        Controls.Add(Btn_Merge)
        Controls.Add(Btn_Test)
        Controls.Add(ToolStripContainer1)
        Controls.Add(Lst_MailMerge)
        Controls.Add(TextBox)
        Controls.Add(MenuStrip1)
        Font = New Font("Segoe UI", 12F, FontStyle.Bold, GraphicsUnit.Point, 0)
        MainMenuStrip = MenuStrip1
        Margin = New Padding(4)
        Name = "Form1"
        StartPosition = FormStartPosition.CenterScreen
        Text = "Form1"
        MenuStrip1.ResumeLayout(False)
        MenuStrip1.PerformLayout()
        ToolStripContainer1.ResumeLayout(False)
        ToolStripContainer1.PerformLayout()
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents TextBox As RichTextBox
    Friend WithEvents MenuStrip1 As MenuStrip
    Friend WithEvents FileToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents LoadToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents SaveAsToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents SaveToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents PrintToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ExitToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents LoadToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents SaveToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents MergeToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents PrintToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents Lst_MailMerge As ListBox
    Friend WithEvents ToolStripContainer1 As ToolStripContainer
    Friend WithEvents Btn_Test As Button
    Friend WithEvents Btn_Merge As Button
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents PrintPreviewControl1 As PrintPreviewControl
    Friend WithEvents PrintDocument1 As Printing.PrintDocument
    Friend WithEvents PageSetupDialog1 As PageSetupDialog
    Friend WithEvents PrintDialog1 As PrintDialog
    Friend WithEvents SaveFileDialog1 As SaveFileDialog
    Friend WithEvents NewToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents LetterToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents MailMergeToolStripMenuItem As ToolStripMenuItem

End Class
