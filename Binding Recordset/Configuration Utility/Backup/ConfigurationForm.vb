Public Class ConfigurationForm
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
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents DataHelp_Name As System.Windows.Forms.TextBox
    Friend WithEvents DataHelp_Id As System.Windows.Forms.TextBox
    Friend WithEvents DataHelp_Caption As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Delete_Message As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents OkButton As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.DataHelp_Name = New System.Windows.Forms.TextBox()
        Me.DataHelp_Id = New System.Windows.Forms.TextBox()
        Me.DataHelp_Caption = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.TabPage3 = New System.Windows.Forms.TabPage()
        Me.Delete_Message = New System.Windows.Forms.TextBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.OkButton = New System.Windows.Forms.Button()
        Me.TabControl1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.TabPage3.SuspendLayout()
        Me.SuspendLayout()
        '
        'TabControl1
        '
        Me.TabControl1.Controls.AddRange(New System.Windows.Forms.Control() {Me.TabPage2, Me.TabPage3})
        Me.TabControl1.Location = New System.Drawing.Point(8, 8)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(456, 352)
        Me.TabControl1.TabIndex = 23
        '
        'TabPage2
        '
        Me.TabPage2.Controls.AddRange(New System.Windows.Forms.Control() {Me.DataHelp_Name, Me.DataHelp_Id, Me.DataHelp_Caption, Me.Label5, Me.Label6, Me.Label7})
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Size = New System.Drawing.Size(448, 326)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "DataHelp"
        '
        'DataHelp_Name
        '
        Me.DataHelp_Name.Location = New System.Drawing.Point(136, 88)
        Me.DataHelp_Name.Name = "DataHelp_Name"
        Me.DataHelp_Name.Size = New System.Drawing.Size(264, 20)
        Me.DataHelp_Name.TabIndex = 21
        Me.DataHelp_Name.Text = ""
        '
        'DataHelp_Id
        '
        Me.DataHelp_Id.Location = New System.Drawing.Point(136, 56)
        Me.DataHelp_Id.Name = "DataHelp_Id"
        Me.DataHelp_Id.Size = New System.Drawing.Size(264, 20)
        Me.DataHelp_Id.TabIndex = 20
        Me.DataHelp_Id.Text = ""
        '
        'DataHelp_Caption
        '
        Me.DataHelp_Caption.Location = New System.Drawing.Point(136, 24)
        Me.DataHelp_Caption.Name = "DataHelp_Caption"
        Me.DataHelp_Caption.Size = New System.Drawing.Size(264, 20)
        Me.DataHelp_Caption.TabIndex = 19
        Me.DataHelp_Caption.Text = ""
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(24, 88)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(112, 24)
        Me.Label5.TabIndex = 18
        Me.Label5.Text = "DataHelp_Name"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(24, 56)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(112, 24)
        Me.Label6.TabIndex = 17
        Me.Label6.Text = "DataHelp_Id"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(24, 24)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(112, 24)
        Me.Label7.TabIndex = 16
        Me.Label7.Text = "DataHelp_Caption"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TabPage3
        '
        Me.TabPage3.Controls.AddRange(New System.Windows.Forms.Control() {Me.Delete_Message, Me.Label11})
        Me.TabPage3.Location = New System.Drawing.Point(4, 22)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Size = New System.Drawing.Size(448, 326)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "Messages"
        '
        'Delete_Message
        '
        Me.Delete_Message.Location = New System.Drawing.Point(128, 24)
        Me.Delete_Message.Name = "Delete_Message"
        Me.Delete_Message.Size = New System.Drawing.Size(264, 20)
        Me.Delete_Message.TabIndex = 25
        Me.Delete_Message.Text = ""
        '
        'Label11
        '
        Me.Label11.Location = New System.Drawing.Point(16, 16)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(112, 24)
        Me.Label11.TabIndex = 22
        Me.Label11.Text = "Delete_Message"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'OkButton
        '
        Me.OkButton.Location = New System.Drawing.Point(147, 368)
        Me.OkButton.Name = "OkButton"
        Me.OkButton.Size = New System.Drawing.Size(144, 24)
        Me.OkButton.TabIndex = 22
        Me.OkButton.Text = "Save Configuration"
        '
        'ConfigurationForm
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(464, 398)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.TabControl1, Me.OkButton})
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "ConfigurationForm"
        Me.Text = "Configuration Utility"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage3.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region
    Dim MyFile As String
    Dim MyFullPathFile As String
    Dim oInitValues As New InitValues()
    Dim FileNum As Byte
    Private Sub InitValuesForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim WinSys As String
        Dim WshShell As New Object()

        WshShell = CreateObject("WScript.Shell")
        WinSys = WshShell.SpecialFolders("Fonts")
        WinSys = Mid(WinSys, 1, Len(WinSys) - 5)
        WinSys += "System32\"
        FileNum = FreeFile()
        MyFullPathFile = WinSys + "DCBR10_Lang.dll"
        MyFile = Dir(WinSys + "DCBR10_Lang.dll")

        FileOpen(FileNum, MyFullPathFile, OpenMode.Random, OpenAccess.ReadWrite, OpenShare.Shared, 1000)
        If UCase(MyFile) = UCase("DCBR10_Lang.dll") Then
            FileGet(FileNum, oInitValues, 1)
            Me.DataHelp_Caption.Text = oInitValues.DataHelp_Caption
            ' Keep 5 room for Future add
            aInitValues(15) = oInitValues.DataHelp_Caption
            Me.DataHelp_Id.Text = oInitValues.DataHelp_Id
            aInitValues(16) = oInitValues.DataHelp_Id
            Me.DataHelp_Name.Text = oInitValues.DataHelp_Name
            aInitValues(17) = oInitValues.DataHelp_Name
            ' Keep 2 room for Future add
            Me.Delete_Message.Text = oInitValues.Delete_Message
            aInitValues(20) = oInitValues.Delete_Message
        End If
    End Sub

    Private Sub OkButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OkButton.Click
        If UCase(MyFile) = UCase("DCBR10_Lang.dll") Then
            FileGet(FileNum, oInitValues, 1)
            ' Keep 5 room for Future add
            oInitValues.DataHelp_Caption = Me.DataHelp_Caption.Text
            oInitValues.DataHelp_Id = Me.DataHelp_Id.Text
            oInitValues.DataHelp_Name = Me.DataHelp_Name.Text
            ' Keep 2 room for Future add
            oInitValues.Delete_Message = Me.Delete_Message.Text

            FilePut(FileNum, oInitValues, 1)
        End If
        Me.Close()
    End Sub

    Private Sub InitValuesForm_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        On Error Resume Next
        FileClose(FileNum)
    End Sub

End Class
