Imports System.IO
Public Class Form1
    Inherits System.Windows.Forms.Form
#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents LstDrives As System.Windows.Forms.ListView
    Friend WithEvents LargeImage As System.Windows.Forms.ImageList
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents LstFolder As System.Windows.Forms.TreeView
    Public WithEvents FolderImage As System.Windows.Forms.ImageList
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader5 As System.Windows.Forms.ColumnHeader
    Public WithEvents SmallImage As System.Windows.Forms.ImageList
    Friend WithEvents LstFiles As System.Windows.Forms.ListView
    Friend WithEvents ChkReadOnly As System.Windows.Forms.CheckBox
    Friend WithEvents ChkHidden As System.Windows.Forms.CheckBox
    Friend WithEvents ChkSystem As System.Windows.Forms.CheckBox
    Friend WithEvents ChkArchive As System.Windows.Forms.CheckBox
    Friend WithEvents GrpAttributes As System.Windows.Forms.GroupBox
    Friend WithEvents BttnApplyAttributes As System.Windows.Forms.Button
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents BttnDeleteFile As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.Label1 = New System.Windows.Forms.Label()
        Me.LstDrives = New System.Windows.Forms.ListView()
        Me.LargeImage = New System.Windows.Forms.ImageList(Me.components)
        Me.LstFolder = New System.Windows.Forms.TreeView()
        Me.FolderImage = New System.Windows.Forms.ImageList(Me.components)
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.LstFiles = New System.Windows.Forms.ListView()
        Me.ColumnHeader1 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader2 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader3 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader4 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader5 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.SmallImage = New System.Windows.Forms.ImageList(Me.components)
        Me.BttnApplyAttributes = New System.Windows.Forms.Button()
        Me.GrpAttributes = New System.Windows.Forms.GroupBox()
        Me.ChkReadOnly = New System.Windows.Forms.CheckBox()
        Me.ChkHidden = New System.Windows.Forms.CheckBox()
        Me.ChkSystem = New System.Windows.Forms.CheckBox()
        Me.ChkArchive = New System.Windows.Forms.CheckBox()
        Me.BttnDeleteFile = New System.Windows.Forms.Button()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.GrpAttributes.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Label1.Font = New System.Drawing.Font("MS Reference Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Red
        Me.Label1.Location = New System.Drawing.Point(5, 31)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(86, 27)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Dischi:"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'LstDrives
        '
        Me.LstDrives.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.LstDrives.BackColor = System.Drawing.Color.White
        Me.LstDrives.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.LstDrives.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LstDrives.FullRowSelect = True
        Me.LstDrives.LargeImageList = Me.LargeImage
        Me.LstDrives.Location = New System.Drawing.Point(5, 61)
        Me.LstDrives.MultiSelect = False
        Me.LstDrives.Name = "LstDrives"
        Me.LstDrives.Size = New System.Drawing.Size(86, 424)
        Me.LstDrives.Sorting = System.Windows.Forms.SortOrder.Ascending
        Me.LstDrives.TabIndex = 2
        Me.LstDrives.UseCompatibleStateImageBehavior = False
        '
        'LargeImage
        '
        Me.LargeImage.ImageStream = CType(resources.GetObject("LargeImage.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.LargeImage.TransparentColor = System.Drawing.Color.Transparent
        Me.LargeImage.Images.SetKeyName(0, "Visualpharm-Hardware-Hard-disk.ico")
        '
        'LstFolder
        '
        Me.LstFolder.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.LstFolder.BackColor = System.Drawing.Color.White
        Me.LstFolder.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.LstFolder.ImageIndex = 0
        Me.LstFolder.ImageList = Me.FolderImage
        Me.LstFolder.Location = New System.Drawing.Point(97, 61)
        Me.LstFolder.Name = "LstFolder"
        Me.LstFolder.SelectedImageIndex = 0
        Me.LstFolder.Size = New System.Drawing.Size(239, 424)
        Me.LstFolder.TabIndex = 3
        '
        'FolderImage
        '
        Me.FolderImage.ImageStream = CType(resources.GetObject("FolderImage.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.FolderImage.TransparentColor = System.Drawing.Color.Transparent
        Me.FolderImage.Images.SetKeyName(0, "Treetog-Junior-Folder-close.ico")
        Me.FolderImage.Images.SetKeyName(1, "Treetog-Junior-Folder-open.ico")
        Me.FolderImage.Images.SetKeyName(2, "Visualpharm-Hardware-Hard-disk.ico")
        '
        'Label2
        '
        Me.Label2.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Label2.Font = New System.Drawing.Font("MS Reference Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(97, 31)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(147, 27)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Seleziona cartella:"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label3
        '
        Me.Label3.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Label3.Font = New System.Drawing.Font("MS Reference Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(339, 31)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(116, 27)
        Me.Label3.TabIndex = 1
        Me.Label3.Text = "File contenuti:"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'LstFiles
        '
        Me.LstFiles.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.LstFiles.BackColor = System.Drawing.Color.White
        Me.LstFiles.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.LstFiles.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2, Me.ColumnHeader3, Me.ColumnHeader4, Me.ColumnHeader5})
        Me.LstFiles.Cursor = System.Windows.Forms.Cursors.Default
        Me.LstFiles.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LstFiles.ForeColor = System.Drawing.Color.Black
        Me.LstFiles.FullRowSelect = True
        Me.LstFiles.Location = New System.Drawing.Point(342, 61)
        Me.LstFiles.MultiSelect = False
        Me.LstFiles.Name = "LstFiles"
        Me.LstFiles.Size = New System.Drawing.Size(374, 513)
        Me.LstFiles.SmallImageList = Me.SmallImage
        Me.LstFiles.Sorting = System.Windows.Forms.SortOrder.Ascending
        Me.LstFiles.TabIndex = 2
        Me.LstFiles.UseCompatibleStateImageBehavior = False
        Me.LstFiles.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Nome File"
        Me.ColumnHeader1.Width = 202
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "Dim."
        Me.ColumnHeader2.Width = 80
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "Tipo"
        Me.ColumnHeader3.Width = 120
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "Ultima Modifica"
        Me.ColumnHeader4.Width = 150
        '
        'ColumnHeader5
        '
        Me.ColumnHeader5.Text = "Attributi"
        Me.ColumnHeader5.Width = 230
        '
        'SmallImage
        '
        Me.SmallImage.ImageStream = CType(resources.GetObject("SmallImage.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.SmallImage.TransparentColor = System.Drawing.Color.Transparent
        Me.SmallImage.Images.SetKeyName(0, "exe.ico")
        Me.SmallImage.Images.SetKeyName(1, "paint.ico")
        Me.SmallImage.Images.SetKeyName(2, "txt.ico")
        Me.SmallImage.Images.SetKeyName(3, "dll.ico")
        Me.SmallImage.Images.SetKeyName(4, "internet.ico")
        Me.SmallImage.Images.SetKeyName(5, "config.ico")
        Me.SmallImage.Images.SetKeyName(6, "winamp.ico")
        Me.SmallImage.Images.SetKeyName(7, "mediaplayer.ico")
        Me.SmallImage.Images.SetKeyName(8, "quick.ico")
        Me.SmallImage.Images.SetKeyName(9, "winzip.ico")
        Me.SmallImage.Images.SetKeyName(10, "winrar.ico")
        Me.SmallImage.Images.SetKeyName(11, "word.ico")
        Me.SmallImage.Images.SetKeyName(12, "excel.ico")
        Me.SmallImage.Images.SetKeyName(13, "access.ico")
        Me.SmallImage.Images.SetKeyName(14, "outlook.ico")
        Me.SmallImage.Images.SetKeyName(15, "pdf.ico")
        Me.SmallImage.Images.SetKeyName(16, "")
        Me.SmallImage.Images.SetKeyName(17, "help.ico")
        Me.SmallImage.Images.SetKeyName(18, "blank.ico")
        Me.SmallImage.Images.SetKeyName(19, "powerpoint.ico")
        Me.SmallImage.Images.SetKeyName(20, "png.ico")
        Me.SmallImage.Images.SetKeyName(21, "bat.ico")
        Me.SmallImage.Images.SetKeyName(22, "vlc.ico")
        Me.SmallImage.Images.SetKeyName(23, "link.ico")
        Me.SmallImage.Images.SetKeyName(24, "sistema.ico")
        Me.SmallImage.Images.SetKeyName(25, "reg.png")
        Me.SmallImage.Images.SetKeyName(26, "iso.ico")
        '
        'BttnApplyAttributes
        '
        Me.BttnApplyAttributes.Font = New System.Drawing.Font("MS Reference Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BttnApplyAttributes.ForeColor = System.Drawing.Color.Green
        Me.BttnApplyAttributes.Location = New System.Drawing.Point(255, 22)
        Me.BttnApplyAttributes.Name = "BttnApplyAttributes"
        Me.BttnApplyAttributes.Size = New System.Drawing.Size(70, 25)
        Me.BttnApplyAttributes.TabIndex = 4
        Me.BttnApplyAttributes.Text = "Applica"
        '
        'GrpAttributes
        '
        Me.GrpAttributes.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.GrpAttributes.Controls.Add(Me.ChkReadOnly)
        Me.GrpAttributes.Controls.Add(Me.ChkHidden)
        Me.GrpAttributes.Controls.Add(Me.ChkSystem)
        Me.GrpAttributes.Controls.Add(Me.ChkArchive)
        Me.GrpAttributes.Controls.Add(Me.BttnApplyAttributes)
        Me.GrpAttributes.Controls.Add(Me.BttnDeleteFile)
        Me.GrpAttributes.Enabled = False
        Me.GrpAttributes.Font = New System.Drawing.Font("MS Reference Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GrpAttributes.Location = New System.Drawing.Point(5, 491)
        Me.GrpAttributes.Name = "GrpAttributes"
        Me.GrpAttributes.Size = New System.Drawing.Size(331, 83)
        Me.GrpAttributes.TabIndex = 5
        Me.GrpAttributes.TabStop = False
        Me.GrpAttributes.Text = "Operazioni consentite:"
        '
        'ChkReadOnly
        '
        Me.ChkReadOnly.Font = New System.Drawing.Font("MS Reference Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ChkReadOnly.Location = New System.Drawing.Point(6, 24)
        Me.ChkReadOnly.Name = "ChkReadOnly"
        Me.ChkReadOnly.Size = New System.Drawing.Size(98, 24)
        Me.ChkReadOnly.TabIndex = 0
        Me.ChkReadOnly.Text = "Sola Lettura"
        '
        'ChkHidden
        '
        Me.ChkHidden.Font = New System.Drawing.Font("MS Reference Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ChkHidden.Location = New System.Drawing.Point(6, 52)
        Me.ChkHidden.Name = "ChkHidden"
        Me.ChkHidden.Size = New System.Drawing.Size(89, 24)
        Me.ChkHidden.TabIndex = 0
        Me.ChkHidden.Text = "Nascosto"
        '
        'ChkSystem
        '
        Me.ChkSystem.Font = New System.Drawing.Font("MS Reference Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ChkSystem.Location = New System.Drawing.Point(110, 52)
        Me.ChkSystem.Name = "ChkSystem"
        Me.ChkSystem.Size = New System.Drawing.Size(77, 24)
        Me.ChkSystem.TabIndex = 0
        Me.ChkSystem.Text = "System"
        '
        'ChkArchive
        '
        Me.ChkArchive.Font = New System.Drawing.Font("MS Reference Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ChkArchive.Location = New System.Drawing.Point(110, 24)
        Me.ChkArchive.Name = "ChkArchive"
        Me.ChkArchive.Size = New System.Drawing.Size(129, 24)
        Me.ChkArchive.TabIndex = 0
        Me.ChkArchive.Text = "Archivio (no .rar)"
        '
        'BttnDeleteFile
        '
        Me.BttnDeleteFile.Font = New System.Drawing.Font("MS Reference Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BttnDeleteFile.ForeColor = System.Drawing.Color.Red
        Me.BttnDeleteFile.Location = New System.Drawing.Point(221, 51)
        Me.BttnDeleteFile.Name = "BttnDeleteFile"
        Me.BttnDeleteFile.Size = New System.Drawing.Size(104, 25)
        Me.BttnDeleteFile.TabIndex = 4
        Me.BttnDeleteFile.Text = "Elimina il file"
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(12, 12)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(531, 20)
        Me.TextBox1.TabIndex = 6
        '
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("MS Reference Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Location = New System.Drawing.Point(549, 6)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 31)
        Me.Button1.TabIndex = 7
        Me.Button1.Text = "Vai"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Font = New System.Drawing.Font("MS Reference Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button2.Location = New System.Drawing.Point(630, 6)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(86, 31)
        Me.Button2.TabIndex = 8
        Me.Button2.Text = "Info"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(728, 579)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.GrpAttributes)
        Me.Controls.Add(Me.LstFolder)
        Me.Controls.Add(Me.LstDrives)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.LstFiles)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Dettaglio Cartelle   |   Jordan Merli - joe_nanoteck@hotmail.it"
        Me.GrpAttributes.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region
    Private Declare Function AssocQueryString Lib "shlwapi.dll" Alias "AssocQueryStringA" (ByVal flags As Integer, ByVal stringType As Integer, ByVal pszAssoc As String, ByVal pszExtra As String, ByVal pszOut As String, ByRef pcchOut As Integer) As Integer
    Private Const ASSOCSTR_FRIENDLYDOCNAME As Integer = 3
    Dim StrDrv, StrFldr, StrFile, StrPath As String
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Activate()
        Dim LstRoot As ListViewItem
        Dim Drv As String
        For Each Drv In Directory.GetLogicalDrives
            LstRoot = Me.LstDrives.Items.Add(Drv)
            LstRoot.ImageIndex = 0
        Next
        TextBox1.AcceptsReturn = False
    End Sub
    Private Sub LstDrives_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LstDrives.Click
        On Error GoTo Fix
        StrDrv = ""
        StrDrv = Me.LstDrives.FocusedItem.Text()
        Dim CurrFldr As New DirectoryInfo(StrDrv)
        Dim Fldr As DirectoryInfo
        Me.LstFolder.Nodes.Clear()
        For Each Fldr In CurrFldr.GetDirectories
            Me.LstFolder.Nodes.Add(Fldr.Name)
        Next
        Me.LstFiles.Items.Clear()
        Me.LoadFiles_Drv()
        StrFldr = ""
        StrFile = ""
        Me.ResetGroup()
        Me.ResetChk()
Fix:
    End Sub
    Private Sub LoadFiles_Drv()
        Dim CurrFldr As New DirectoryInfo(StrDrv)
        Dim FileInfo As FileInfo
        Dim fExt As String
        Dim FileSize As Long
        Try
            Me.LstFiles.Items.Clear()
            For Each FileInfo In CurrFldr.GetFiles
                Try
                    FileSize = FileSize + FileInfo.Length
                Catch ex As OverflowException
                    MsgBox("La dimensione del file totale supera la dimensione consentita. Le dimensioni totali dei file mostrate non saranno attendibili.")
                End Try
                Dim LstItem As ListViewItem
                Dim SzMB, SzKB As Double
                SzMB = (FileInfo.Length / 1024) / 1024
                SzKB = (FileInfo.Length / 1024)
                fExt = Microsoft.VisualBasic.LCase(FileInfo.Extension)
                If fExt = ".exe" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 0)
                ElseIf fExt = ".bmp" Or fExt = ".gif" Or fExt = ".jpg" Or fExt = ".jpeg" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 1)
                ElseIf fExt = ".txt" Or fExt = ".rtf" Or fExt = ".ttf" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 2)
                ElseIf fExt = ".dll" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 3)
                ElseIf fExt = ".htm" Or fExt = ".html" Or fExt = ".asp" Or fExt = ".aspx" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 4)
                ElseIf fExt = ".inf" Or fExt = ".ini" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 5)
                ElseIf fExt = ".wma" Or fExt = ".mp3" Or fExt = ".wav" Or fExt = ".mp1" Or fExt = ".mp2" Or fExt = ".m3u" Or fExt = ".mid" Or fExt = ".midi" Or fExt = ".wax" Or fExt = ".amr" Or fExt = ".mp4" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 7)
                ElseIf fExt = ".wmv" Or fExt = ".asf" Or fExt = ".asx" Or fExt = ".avi" Or fExt = ".dat" Or fExt = ".mpg" Or fExt = ".mpeg" Or fExt = ".m1v" Or fExt = ".mp2v" Or fExt = ".mpa" Or fExt = ".mpe" Or fExt = ".mpv2" Or fExt = ".wm" Or fExt = ".wmx" Or fExt = ".wvx" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 7)
                ElseIf fExt = ".adts" Or fExt = ".aifc" Or fExt = ".cdda" Or fExt = ".amc" Or fExt = ".dif" Or fExt = ".dv" Or fExt = ".dvd" Or fExt = ".vob" Or fExt = ".3gp" Or fExt = ".mov" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 8)
                ElseIf fExt = ".zip" Or fExt = ".7-zip" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 9)
                ElseIf fExt = ".rar" Or fExt = ".ace" Or fExt = ".arj" Or Text = ".bz2" Or fExt = ".cab" Or fExt = ".gzip" Or fExt = ".jar" Or fExt = ".lzh" Or fExt = ".tar" Or fExt = ".uue" Or fExt = ".xz" Or fExt = ".z" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 10)
                ElseIf fExt = ".doc" Or fExt = ".docx" Or fExt = ".docm" Or fExt = ".dotx" Or fExt = ".dotm" Or fExt = ".dot" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 11)
                ElseIf fExt = ".xls" Or fExt = ".xlsx" Or fExt = ".xlsm" Or fExt = ".xlsb" Or fExt = ".xltx" Or fExt = ".xltm" Or fExt = ".xlt" Or fExt = ".xml" Or fExt = ".xlam" Or fExt = ".xla" Or fExt = ".xlw" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 12)
                ElseIf fExt = ".mdb" Or fExt = ".accdb" Or fExt = ".accde" Or fExt = ".accdt" Or fExt = ".accdr" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 13)
                ElseIf fExt = ".ppt" Or fExt = ".pot" Or fExt = ".pps" Or fExt = ".pptx" Or fExt = ".pptm" Or fExt = ".potx" Or fExt = ".potm" Or fExt = ".ppam" Or fExt = ".ppsx" Or fExt = ".ppsm" Or fExt = ".sldx" Or fExt = ".sldm" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 19)
                ElseIf fExt = ".pdf" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 15)
                ElseIf fExt = ".max" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 16)
                ElseIf fExt = ".hlp" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 17)
                ElseIf fExt = ".png" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 20)
                ElseIf fExt = ".bat" Or fExt = ".com" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 21)
                ElseIf fExt = ".flv" Or fExt = ".vlc" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 22)
                ElseIf fExt = ".sys" Or fExt = ".config" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 24)
                ElseIf fExt = ".lnk" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 23)
                ElseIf fExt = ".reg" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 25)
                ElseIf fExt = ".iso" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 26)
                Else
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 18)
                End If
                If SzMB <= 1 Then
                    With LstItem.SubItems
                        .Add(Math.Round(SzKB) & " KB")
                        .Add(GetFileType(FileInfo))
                        .Add(FileInfo.LastWriteTime.ToShortDateString)
                        .Add(FileInfo.Attributes.ToString)
                    End With
                Else
                    With LstItem.SubItems
                        .Add(Math.Round(SzMB) & " MB")
                        .Add(GetFileType(FileInfo))
                        .Add(FileInfo.LastWriteTime.ToShortDateString)
                        .Add(FileInfo.Attributes.ToString)
                    End With
                End If
            Next
        Catch ex As Exception
        End Try
    End Sub
    Private Sub LstFolder_AfterSelect(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles LstFolder.AfterSelect
        'On Error GoTo Fix
        Dim CurrFldr As New DirectoryInfo(StrDrv & e.Node.FullPath)
        Dim Fldr As DirectoryInfo
        Dim newNode As TreeNode
        StrFldr = ""
        StrFldr = e.Node.FullPath
        Try
            If e.Node.Nodes.Count = 0 Then
                For Each Fldr In CurrFldr.GetDirectories
                    'Me.AddTreeNode(e.Node, Fldr.Name)
                    newNode = New TreeNode(Fldr.Name, 0, 1)
                    e.Node.Nodes.Add(newNode)
                Next Fldr
                e.Node.Expand()
            End If
        Catch ex As Exception
        End Try
        Me.LstFiles.Items.Clear()
        Me.LoadFiles_Folder(e)
        StrFile = ""
        Me.ResetGroup()
        Me.ResetChk()
    End Sub
    Private Sub LoadFiles_Folder(ByVal e As System.Windows.Forms.TreeViewEventArgs)
        Dim CurrFldr As New DirectoryInfo(StrDrv & e.Node.FullPath)
        Dim FileInfo As FileInfo
        Dim fExt As String
        Dim FileSize As Long
        Try
            Me.LstFiles.Items.Clear()
            For Each FileInfo In CurrFldr.GetFiles
                Try
                    FileSize = FileSize + FileInfo.Length
                Catch ex As OverflowException
                    MsgBox("La dimensione del file totale supera la dimensione consentita. Le dimensioni totali dei file mostrate non saranno attendibili.")
                End Try
                Dim LstItem As ListViewItem
                Dim SzMB, SzKB As Double
                SzMB = (FileInfo.Length / 1024) / 1024
                SzKB = (FileInfo.Length / 1024)
                fExt = Microsoft.VisualBasic.LCase(FileInfo.Extension)
                If fExt = ".exe" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 0)
                ElseIf fExt = ".bmp" Or fExt = ".gif" Or fExt = ".jpg" Or fExt = ".jpeg" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 1)
                ElseIf fExt = ".txt" Or fExt = ".rtf" Or fExt = ".ttf" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 2)
                ElseIf fExt = ".dll" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 3)
                ElseIf fExt = ".htm" Or fExt = ".html" Or fExt = ".asp" Or fExt = ".aspx" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 4)
                ElseIf fExt = ".inf" Or fExt = ".ini" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 5)
                ElseIf fExt = ".wma" Or fExt = ".mp3" Or fExt = ".wav" Or fExt = ".mp1" Or fExt = ".mp2" Or fExt = ".m3u" Or fExt = ".mid" Or fExt = ".midi" Or fExt = ".wax" Or fExt = ".amr" Or fExt = ".mp4" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 7)
                ElseIf fExt = ".wmv" Or fExt = ".asf" Or fExt = ".asx" Or fExt = ".avi" Or fExt = ".dat" Or fExt = ".mpg" Or fExt = ".mpeg" Or fExt = ".m1v" Or fExt = ".mp2v" Or fExt = ".mpa" Or fExt = ".mpe" Or fExt = ".mpv2" Or fExt = ".wm" Or fExt = ".wmx" Or fExt = ".wvx" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 7)
                ElseIf fExt = ".adts" Or fExt = ".aifc" Or fExt = ".cdda" Or fExt = ".amc" Or fExt = ".dif" Or fExt = ".dv" Or fExt = ".dvd" Or fExt = ".vob" Or fExt = ".3gp" Or fExt = ".mov" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 8)
                ElseIf fExt = ".zip" Or fExt = ".7-zip" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 9)
                ElseIf fExt = ".rar" Or fExt = ".ace" Or fExt = ".arj" Or Text = ".bz2" Or fExt = ".cab" Or fExt = ".gzip" Or fExt = ".jar" Or fExt = ".lzh" Or fExt = ".tar" Or fExt = ".uue" Or fExt = ".xz" Or fExt = ".z" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 10)
                ElseIf fExt = ".doc" Or fExt = ".docx" Or fExt = ".docm" Or fExt = ".dotx" Or fExt = ".dotm" Or fExt = ".dot" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 11)
                ElseIf fExt = ".xls" Or fExt = ".xlsx" Or fExt = ".xlsm" Or fExt = ".xlsb" Or fExt = ".xltx" Or fExt = ".xltm" Or fExt = ".xlt" Or fExt = ".xml" Or fExt = ".xlam" Or fExt = ".xla" Or fExt = ".xlw" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 12)
                ElseIf fExt = ".mdb" Or fExt = ".accdb" Or fExt = ".accde" Or fExt = ".accdt" Or fExt = ".accdr" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 13)
                ElseIf fExt = ".ppt" Or fExt = ".pot" Or fExt = ".pps" Or fExt = ".pptx" Or fExt = ".pptm" Or fExt = ".potx" Or fExt = ".potm" Or fExt = ".ppam" Or fExt = ".ppsx" Or fExt = ".ppsm" Or fExt = ".sldx" Or fExt = ".sldm" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 19)
                ElseIf fExt = ".pdf" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 15)
                ElseIf fExt = ".max" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 16)
                ElseIf fExt = ".hlp" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 17)
                ElseIf fExt = ".png" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 20)
                ElseIf fExt = ".bat" Or fExt = ".com" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 21)
                ElseIf fExt = ".flv" Or fExt = ".vlc" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 22)
                ElseIf fExt = ".sys" Or fExt = ".config" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 24)
                ElseIf fExt = ".lnk" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 23)
                ElseIf fExt = ".reg" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 25)
                ElseIf fExt = ".iso" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 26)
                Else
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 18)
                End If
                If SzMB <= 1 Then
                    With LstItem.SubItems
                        .Add(Math.Round(SzKB) & " KB")
                        .Add(GetFileType(FileInfo))
                        .Add(FileInfo.LastWriteTime.ToShortDateString)
                        .Add(FileInfo.Attributes.ToString)
                    End With
                Else
                    With LstItem.SubItems
                        .Add(Math.Round(SzMB) & " MB")
                        .Add(GetFileType(FileInfo))
                        .Add(FileInfo.LastWriteTime.ToShortDateString)
                        .Add(FileInfo.Attributes.ToString)
                    End With
                End If
            Next
        Catch ex As Exception
        End Try
    End Sub
    Private Function GetFileType(ByRef fl As FileInfo) As String
        Dim strRet As New String(CChar(" "), 260)
        Dim lngRet As Integer = Len(strRet)
        AssocQueryString(0, ASSOCSTR_FRIENDLYDOCNAME, fl.Extension, vbNullString, strRet, lngRet)
        strRet = strRet.Substring(0, lngRet - 1).Trim
        If Len(strRet) > 0 Then
            Return strRet
        ElseIf Len(fl.Extension) > 0 Then
            Return Mid(fl.Extension, 2) & " File"
        Else
            Return "File"
        End If
    End Function
    Private Sub LstFiles_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LstFiles.Click
        On Error GoTo Fix
        StrFile = ""
        StrFile = Me.LstFiles.FocusedItem.Text()
        Me.ResetGroup()
        If Not StrFldr = "" Then
            StrPath = StrDrv & StrFldr & "\" & StrFile
        Else
            StrPath = StrDrv & StrFldr & StrFile
        End If
        If (File.GetAttributes(StrPath) And FileAttributes.ReadOnly) = FileAttributes.ReadOnly Then
            Me.ChkReadOnly.Checked = True
        Else
            Me.ChkReadOnly.Checked = False
        End If
        If (File.GetAttributes(StrPath) And FileAttributes.Hidden) = FileAttributes.Hidden Then
            Me.ChkHidden.Checked = True
        Else
            Me.ChkHidden.Checked = False
        End If
        If (File.GetAttributes(StrPath) And FileAttributes.System) = FileAttributes.System Then
            Me.ChkSystem.Checked = True
        Else
            Me.ChkSystem.Checked = False
        End If
        If (File.GetAttributes(StrPath) And FileAttributes.Archive) = FileAttributes.Archive Then
            Me.ChkArchive.Checked = True
        Else
            Me.ChkArchive.Checked = False
        End If
Fix:
    End Sub
    Private Sub BttnApplyAttributes_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BttnApplyAttributes.Click
        If MsgBox("Vuoi veramente applicare i cambiamenti?", MsgBoxStyle.Question + vbOKCancel, "Applica Attributi") = MsgBoxResult.Ok Then
            If Me.ChkReadOnly.Checked = True And Me.ChkHidden.Checked = False And Me.ChkSystem.Checked = False And Me.ChkArchive.Checked = False Then
                'Set File Attribute (ReadOnly)
                File.SetAttributes(StrPath, FileAttributes.ReadOnly)
                Me.ResetLstFiles()
                Me.ResetAttribute()
            ElseIf Me.ChkReadOnly.Checked = False And Me.ChkHidden.Checked = True And Me.ChkSystem.Checked = False And Me.ChkArchive.Checked = False Then
                'Set File Attribute (Hidden)
                File.SetAttributes(StrPath, FileAttributes.Hidden)
                Me.ResetLstFiles()
                Me.ResetAttribute()
            ElseIf Me.ChkReadOnly.Checked = False And Me.ChkHidden.Checked = False And Me.ChkSystem.Checked = True And Me.ChkArchive.Checked = False Then
                'Set File Attribute (System)
                File.SetAttributes(StrPath, FileAttributes.System)
                Me.ResetLstFiles()
                Me.ResetAttribute()
            ElseIf Me.ChkReadOnly.Checked = False And Me.ChkHidden.Checked = False And Me.ChkSystem.Checked = False And Me.ChkArchive.Checked = True Then
                'Set File Attribute (Archive)
                File.SetAttributes(StrPath, FileAttributes.Archive)
                Me.ResetLstFiles()
                Me.ResetAttribute()
            ElseIf Me.ChkReadOnly.Checked = True And Me.ChkHidden.Checked = True And Me.ChkSystem.Checked = False And Me.ChkArchive.Checked = False Then
                'Set File Attribute (ReadOnly & Hidden)
                File.SetAttributes(StrPath, FileAttributes.ReadOnly + FileAttributes.Hidden)
                Me.ResetLstFiles()
                Me.ResetAttribute()
            ElseIf Me.ChkReadOnly.Checked = True And Me.ChkHidden.Checked = False And Me.ChkSystem.Checked = True And Me.ChkArchive.Checked = False Then
                'Set File Attribute (ReadOnly & System)
                File.SetAttributes(StrPath, FileAttributes.ReadOnly + FileAttributes.System)
                Me.ResetLstFiles()
                Me.ResetAttribute()
            ElseIf Me.ChkReadOnly.Checked = True And Me.ChkHidden.Checked = False And Me.ChkSystem.Checked = False And Me.ChkArchive.Checked = True Then
                'Set File Attribute (ReadOnly & Archive)
                File.SetAttributes(StrPath, FileAttributes.ReadOnly + FileAttributes.Archive)
                Me.ResetLstFiles()
                Me.ResetAttribute()
            ElseIf Me.ChkReadOnly.Checked = False And Me.ChkHidden.Checked = True And Me.ChkSystem.Checked = True And Me.ChkArchive.Checked = False Then
                'Set File Attribute (Hidden & System)
                File.SetAttributes(StrPath, FileAttributes.Hidden + FileAttributes.System)
                Me.ResetLstFiles()
                Me.ResetAttribute()
            ElseIf Me.ChkReadOnly.Checked = False And Me.ChkHidden.Checked = True And Me.ChkSystem.Checked = False And Me.ChkArchive.Checked = True Then
                'Set File Attribute (Hidden & Archive)
                File.SetAttributes(StrPath, FileAttributes.Hidden + FileAttributes.Archive)
                Me.ResetLstFiles()
                Me.ResetAttribute()
            ElseIf Me.ChkReadOnly.Checked = False And Me.ChkHidden.Checked = False And Me.ChkSystem.Checked = True And Me.ChkArchive.Checked = True Then
                'Set File Attribute (System & Archive)
                File.SetAttributes(StrPath, FileAttributes.System + FileAttributes.Archive)
                Me.ResetLstFiles()
                Me.ResetAttribute()
            ElseIf Me.ChkReadOnly.Checked = True And Me.ChkHidden.Checked = True And Me.ChkSystem.Checked = True And Me.ChkArchive.Checked = False Then
                'Set File Attribute (ReadOnly, Hidden & System)
                File.SetAttributes(StrPath, FileAttributes.ReadOnly + FileAttributes.Hidden + FileAttributes.System)
                Me.ResetLstFiles()
                Me.ResetAttribute()
            ElseIf Me.ChkReadOnly.Checked = True And Me.ChkHidden.Checked = True And Me.ChkSystem.Checked = False And Me.ChkArchive.Checked = True Then
                'Set File Attribute (ReadOnly, Hidden & Archive)
                File.SetAttributes(StrPath, FileAttributes.ReadOnly + FileAttributes.Hidden + FileAttributes.Archive)
                Me.ResetLstFiles()
                Me.ResetAttribute()
            ElseIf Me.ChkReadOnly.Checked = False And Me.ChkHidden.Checked = True And Me.ChkSystem.Checked = True And Me.ChkArchive.Checked = True Then
                'Set File Attribute (Hidden, System & Archive)
                File.SetAttributes(StrPath, FileAttributes.Hidden + FileAttributes.System + FileAttributes.Archive)
                Me.ResetLstFiles()
                Me.ResetAttribute()
            ElseIf Me.ChkReadOnly.Checked = True And Me.ChkHidden.Checked = False And Me.ChkSystem.Checked = True And Me.ChkArchive.Checked = True Then
                'Set File Attribute (ReadOnly, System & Archive)
                File.SetAttributes(StrPath, FileAttributes.ReadOnly + FileAttributes.System + FileAttributes.Archive)
                Me.ResetLstFiles()
                Me.ResetAttribute()
            ElseIf Me.ChkReadOnly.Checked = True And Me.ChkHidden.Checked = True And Me.ChkSystem.Checked = True And Me.ChkArchive.Checked = True Then
                'Set File Attribute (ReadOnly, Hidden, System & Archive)
                File.SetAttributes(StrPath, FileAttributes.ReadOnly + FileAttributes.Hidden + FileAttributes.System + FileAttributes.Archive)
                Me.ResetLstFiles()
                Me.ResetAttribute()
            Else
                'Set File Attribute (Normal)
                File.SetAttributes(StrPath, FileAttributes.Normal)
                Me.ResetLstFiles()
                Me.ResetAttribute()
            End If
        Else
            MsgBox("Nessun cambiamento effettuato." & vbCr & vbCr, MsgBoxStyle.Exclamation, "Indietro")
            'Reset CheckBox
            Me.LstFiles_Click(sender, New System.EventArgs)
        End If
    End Sub
    Private Sub BttnDeleteFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BttnDeleteFile.Click
        Try
            If Not StrPath = "" And Not StrFile = "" Then
                If MsgBox("Vuoi veramente cancellare il file? " & vbCr & vbCr & "Nome del file: " & StrFile & vbCr & vbCr & "Il file verrà eliminato senza passare dal cestino!", MsgBoxStyle.Question + vbYesNo, "Cancella!") = MsgBoxResult.Yes Then
                    File.Delete(StrPath)
                    StrFile = ""
                    Me.ResetGroup()
                    Me.ResetChk()
                    Me.ResetLstFiles()
                End If
            Else
                MsgBox("Seleziona il file da eliminare!", MsgBoxStyle.Exclamation, "Operazione non consentita")
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub ResetGroup()
        If Not StrFile = "" Then
            Me.GrpAttributes.Enabled = True
        Else
            Me.GrpAttributes.Enabled = False
        End If
    End Sub
    Private Sub ResetChk()
        Me.ChkArchive.Checked = False
        Me.ChkHidden.Checked = False
        Me.ChkReadOnly.Checked = False
        Me.ChkSystem.Checked = False
    End Sub
    Private Sub ResetLstFiles()
        Dim CurrFldr As New DirectoryInfo(StrDrv & StrFldr)
        Dim FileInfo As FileInfo
        Dim fExt As String
        Dim FileSize As Long
        Try
            Me.LstFiles.Items.Clear()
            For Each FileInfo In CurrFldr.GetFiles
                Try
                    FileSize = FileSize + FileInfo.Length
                Catch ex As OverflowException
                    MsgBox("La dimensione del file totale supera la dimensione consentita. Le dimensioni totali dei file mostrate non saranno attendibili.")
                End Try
                Dim LstItem As ListViewItem
                Dim SzMB, SzKB As Double
                SzMB = (FileInfo.Length / 1024) / 1024
                SzKB = (FileInfo.Length / 1024)
                fExt = Microsoft.VisualBasic.LCase(FileInfo.Extension)
                If fExt = ".exe" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 0)
                ElseIf fExt = ".bmp" Or fExt = ".gif" Or fExt = ".jpg" Or fExt = ".jpeg" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 1)
                ElseIf fExt = ".txt" Or fExt = ".rtf" Or fExt = ".ttf" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 2)
                ElseIf fExt = ".dll" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 3)
                ElseIf fExt = ".htm" Or fExt = ".html" Or fExt = ".asp" Or fExt = ".aspx" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 4)
                ElseIf fExt = ".inf" Or fExt = ".ini" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 5)
                ElseIf fExt = ".wma" Or fExt = ".mp3" Or fExt = ".wav" Or fExt = ".mp1" Or fExt = ".mp2" Or fExt = ".m3u" Or fExt = ".mid" Or fExt = ".midi" Or fExt = ".wax" Or fExt = ".amr" Or fExt = ".mp4" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 7)
                ElseIf fExt = ".wmv" Or fExt = ".asf" Or fExt = ".asx" Or fExt = ".avi" Or fExt = ".dat" Or fExt = ".mpg" Or fExt = ".mpeg" Or fExt = ".m1v" Or fExt = ".mp2v" Or fExt = ".mpa" Or fExt = ".mpe" Or fExt = ".mpv2" Or fExt = ".wm" Or fExt = ".wmx" Or fExt = ".wvx" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 7)
                ElseIf fExt = ".adts" Or fExt = ".aifc" Or fExt = ".cdda" Or fExt = ".amc" Or fExt = ".dif" Or fExt = ".dv" Or fExt = ".dvd" Or fExt = ".vob" Or fExt = ".3gp" Or fExt = ".mov" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 8)
                ElseIf fExt = ".zip" Or fExt = ".7-zip" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 9)
                ElseIf fExt = ".rar" Or fExt = ".ace" Or fExt = ".arj" Or Text = ".bz2" Or fExt = ".cab" Or fExt = ".gzip" Or fExt = ".jar" Or fExt = ".lzh" Or fExt = ".tar" Or fExt = ".uue" Or fExt = ".xz" Or fExt = ".z" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 10)
                ElseIf fExt = ".doc" Or fExt = ".docx" Or fExt = ".docm" Or fExt = ".dotx" Or fExt = ".dotm" Or fExt = ".dot" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 11)
                ElseIf fExt = ".xls" Or fExt = ".xlsx" Or fExt = ".xlsm" Or fExt = ".xlsb" Or fExt = ".xltx" Or fExt = ".xltm" Or fExt = ".xlt" Or fExt = ".xml" Or fExt = ".xlam" Or fExt = ".xla" Or fExt = ".xlw" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 12)
                ElseIf fExt = ".mdb" Or fExt = ".accdb" Or fExt = ".accde" Or fExt = ".accdt" Or fExt = ".accdr" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 13)
                ElseIf fExt = ".ppt" Or fExt = ".pot" Or fExt = ".pps" Or fExt = ".pptx" Or fExt = ".pptm" Or fExt = ".potx" Or fExt = ".potm" Or fExt = ".ppam" Or fExt = ".ppsx" Or fExt = ".ppsm" Or fExt = ".sldx" Or fExt = ".sldm" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 19)
                ElseIf fExt = ".pdf" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 15)
                ElseIf fExt = ".max" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 16)
                ElseIf fExt = ".hlp" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 17)
                ElseIf fExt = ".png" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 20)
                ElseIf fExt = ".bat" Or fExt = ".com" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 21)
                ElseIf fExt = ".flv" Or fExt = ".vlc" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 22)
                ElseIf fExt = ".sys" Or fExt = ".config" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 24)
                ElseIf fExt = ".lnk" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 23)
                ElseIf fExt = ".reg" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 25)
                ElseIf fExt = ".iso" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 26)
                Else
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 18)
                End If
                If SzMB <= 1 Then
                    With LstItem.SubItems
                        .Add(Math.Round(SzKB) & " KB")
                        .Add(GetFileType(FileInfo))
                        .Add(FileInfo.LastWriteTime.ToShortDateString)
                        .Add(FileInfo.Attributes.ToString)
                    End With
                Else
                    With LstItem.SubItems
                        .Add(Math.Round(SzMB) & " MB")
                        .Add(GetFileType(FileInfo))
                        .Add(FileInfo.LastWriteTime.ToShortDateString)
                        .Add(FileInfo.Attributes.ToString)
                    End With
                End If
            Next
            StrFile = ""
        Catch ex As Exception
        End Try
    End Sub
    Private Sub ResetAttribute()
        If (File.GetAttributes(StrPath) And FileAttributes.ReadOnly) = FileAttributes.ReadOnly Then
            Me.ChkReadOnly.Checked = True
        Else
            Me.ChkReadOnly.Checked = False
        End If
        If (File.GetAttributes(StrPath) And FileAttributes.Hidden) = FileAttributes.Hidden Then
            Me.ChkHidden.Checked = True
        Else
            Me.ChkHidden.Checked = False
        End If
        If (File.GetAttributes(StrPath) And FileAttributes.System) = FileAttributes.System Then
            Me.ChkSystem.Checked = True
        Else
            Me.ChkSystem.Checked = False
        End If
        If (File.GetAttributes(StrPath) And FileAttributes.Archive) = FileAttributes.Archive Then
            Me.ChkArchive.Checked = True
        Else
            Me.ChkArchive.Checked = False
        End If
    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim path As New DirectoryInfo(TextBox1.Text)
        Dim FileInfo As FileInfo
        Dim fExt As String
        Dim FileSize As Long
        Try
            Me.LstFiles.Items.Clear()
            For Each FileInfo In path.GetFiles
                Try
                    FileSize = FileSize + FileInfo.Length
                Catch ex As OverflowException
                    MsgBox("La dimensione del file totale supera la dimensione consentita. Le dimensioni totali dei file mostrate non saranno attendibili.")
                End Try
                Dim LstItem As ListViewItem
                Dim SzMB, SzKB As Double
                SzMB = (FileInfo.Length / 1024) / 1024
                SzKB = (FileInfo.Length / 1024)
                fExt = Microsoft.VisualBasic.LCase(FileInfo.Extension)
                If fExt = ".exe" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 0)
                ElseIf fExt = ".bmp" Or fExt = ".gif" Or fExt = ".jpg" Or fExt = ".jpeg" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 1)
                ElseIf fExt = ".txt" Or fExt = ".rtf" Or fExt = ".ttf" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 2)
                ElseIf fExt = ".dll" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 3)
                ElseIf fExt = ".htm" Or fExt = ".html" Or fExt = ".asp" Or fExt = ".aspx" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 4)
                ElseIf fExt = ".inf" Or fExt = ".ini" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 5)
                ElseIf fExt = ".wma" Or fExt = ".mp3" Or fExt = ".wav" Or fExt = ".mp1" Or fExt = ".mp2" Or fExt = ".m3u" Or fExt = ".mid" Or fExt = ".midi" Or fExt = ".wax" Or fExt = ".amr" Or fExt = ".mp4" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 7)
                ElseIf fExt = ".wmv" Or fExt = ".asf" Or fExt = ".asx" Or fExt = ".avi" Or fExt = ".dat" Or fExt = ".mpg" Or fExt = ".mpeg" Or fExt = ".m1v" Or fExt = ".mp2v" Or fExt = ".mpa" Or fExt = ".mpe" Or fExt = ".mpv2" Or fExt = ".wm" Or fExt = ".wmx" Or fExt = ".wvx" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 7)
                ElseIf fExt = ".adts" Or fExt = ".aifc" Or fExt = ".cdda" Or fExt = ".amc" Or fExt = ".dif" Or fExt = ".dv" Or fExt = ".dvd" Or fExt = ".vob" Or fExt = ".3gp" Or fExt = ".mov" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 8)
                ElseIf fExt = ".zip" Or fExt = ".7-zip" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 9)
                ElseIf fExt = ".rar" Or fExt = ".ace" Or fExt = ".arj" Or Text = ".bz2" Or fExt = ".cab" Or fExt = ".gzip" Or fExt = ".jar" Or fExt = ".lzh" Or fExt = ".tar" Or fExt = ".uue" Or fExt = ".xz" Or fExt = ".z" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 10)
                ElseIf fExt = ".doc" Or fExt = ".docx" Or fExt = ".docm" Or fExt = ".dotx" Or fExt = ".dotm" Or fExt = ".dot" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 11)
                ElseIf fExt = ".xls" Or fExt = ".xlsx" Or fExt = ".xlsm" Or fExt = ".xlsb" Or fExt = ".xltx" Or fExt = ".xltm" Or fExt = ".xlt" Or fExt = ".xml" Or fExt = ".xlam" Or fExt = ".xla" Or fExt = ".xlw" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 12)
                ElseIf fExt = ".mdb" Or fExt = ".accdb" Or fExt = ".accde" Or fExt = ".accdt" Or fExt = ".accdr" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 13)
                ElseIf fExt = ".ppt" Or fExt = ".pot" Or fExt = ".pps" Or fExt = ".pptx" Or fExt = ".pptm" Or fExt = ".potx" Or fExt = ".potm" Or fExt = ".ppam" Or fExt = ".ppsx" Or fExt = ".ppsm" Or fExt = ".sldx" Or fExt = ".sldm" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 19)
                ElseIf fExt = ".pdf" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 15)
                ElseIf fExt = ".max" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 16)
                ElseIf fExt = ".hlp" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 17)
                ElseIf fExt = ".png" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 20)
                ElseIf fExt = ".bat" Or fExt = ".com" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 21)
                ElseIf fExt = ".flv" Or fExt = ".vlc" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 22)
                ElseIf fExt = ".sys" Or fExt = ".config" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 24)
                ElseIf fExt = ".lnk" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 23)
                ElseIf fExt = ".reg" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 25)
                ElseIf fExt = ".iso" Then
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 26)
                Else
                    LstItem = Me.LstFiles.Items.Add(FileInfo.Name, 18)
                End If
                If SzMB <= 1 Then
                    With LstItem.SubItems
                        .Add(Math.Round(SzKB) & " KB")
                        .Add(GetFileType(FileInfo))
                        .Add(FileInfo.LastWriteTime.ToShortDateString)
                        .Add(FileInfo.Attributes.ToString)
                    End With
                Else
                    With LstItem.SubItems
                        .Add(Math.Round(SzMB) & " MB")
                        .Add(GetFileType(FileInfo))
                        .Add(FileInfo.LastWriteTime.ToShortDateString)
                        .Add(FileInfo.Attributes.ToString)
                    End With
                End If
            Next
            StrFile = ""
        Catch ex As Exception
        End Try
    End Sub
    Private Sub TextBox1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            Button1.PerformClick()
        End If
    End Sub
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        form2.show()
    End Sub
End Class
