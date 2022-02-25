Namespace DynamicComponents
    Public Class BindingRecordset
        Inherits System.ComponentModel.Component

#Region " Component Designer generated code "

        Public Sub New(ByVal Container As System.ComponentModel.IContainer)
            MyClass.New()

            'Required for Windows.Forms Class Composition Designer support
            Container.Add(Me)
        End Sub

        'Public Sub New()
        '   MyBase.New()

        'This call is required by the Component Designer.
        '  InitializeComponent()

        'Add any initialization after the InitializeComponent() call

        ' End Sub

        'Component overrides dispose to clean up the component list.
        Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
            If disposing Then
                If Not (components Is Nothing) Then
                    components.Dispose()
                End If
            End If
            MyBase.Dispose(disposing)
        End Sub

        'Required by the Component Designer
        Private components As System.ComponentModel.IContainer

        'NOTE: The following procedure is required by the Component Designer
        'It can be modified using the Component Designer.
        'Do not modify it using the code editor.
        <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
            components = New System.ComponentModel.Container()
        End Sub

#End Region

        Private m_KeyFields As String
        Private Col_ControlName As New Collection()
        Private Col_ControlIndex As New Collection()
        Private Col_KeyValue As New Collection()
        Private Col_KeyText As New Collection()
        Private Col_KeyFields As New Collection()
        Private Col_MasterFields As New Collection()
        Private Col_DetailFields As New Collection()
        Private Col_RequiredFields As New Collection()
        Private Col_GridKeyValue As New Collection()
        Private Col_FieldsType(30, 1) As String
        Private Col_FieldsTypePos As Byte = 0
        Private m_SpecialChars As String
        Private Key_ZeroPad As Byte = 0
        Private m_KeyLeaveField As String = ""
        Private KeyLeavePos As Byte = 0
        Private FilterString As String
        Private DummyFilterString As String = ""
        Private HasGrid As Boolean = False
        Private oMaster As New ADODB.Recordset()
        Private oDetails As New ADODB.Recordset()
        Private MyForm As New System.Windows.Forms.Form()
        Private MyGrid As New AxMSDataGridLib.AxDataGrid()
        Private Col_GridFields As New Collection()
        Private m_MasterFlagField As String = ""
        Private m_DetailFlagField As String = ""
        Private m_FlagValue As String = ""
        Private HelpIdSender As String
        Private RequiredFields_Msg As String = "Uncomplete Entries"
        Private RequiredFields_ShowMsg As Boolean = True
        Private m_HoldSaving As Boolean = False
        Private m_ReleaseSaving As Boolean = False

        Public Sub New()
            MyBase.new()
        End Sub

        Private Function ZeroPad(ByVal str_String As String, ByVal int_Count As Byte) As String
            On Error GoTo EndMe
            If str_String <> "" Then
                Return (New String("0", int_Count - Len(Trim(str_String))) & Trim(str_String))
            End If
EndMe:
            Return str_String
        End Function

        Public Sub InitForm(ByRef dm_DSN As ADODB.Connection, ByRef dm_Form As System.Windows.Forms.Form, ByRef dm_MasterTable As ADODB.Recordset, Optional ByRef dm_Grid As AxMSDataGridLib.AxDataGrid = Nothing, Optional ByRef dm_DetailTable As ADODB.Recordset = Nothing)
            Dim MyProtect As New MyProtection()
            Dim ProductName As String

            ProductName = "DC Binding Recordset"
            MyProtect.SetInformation(ProductName)
            MyProtect.SetAlgorithms(1971, 15, 10, "maaat05")
            MyProtect.SetLicense(30)
            MyProtect.ShowAuthor()
            If MyProtect.NotLicensed Then
                MsgBox("Trial version expired")
                Exit Sub
            End If


            Dim TxtCtrl As New Control()
            Dim X As Byte
            
            If Not (dm_Grid Is Nothing) Then
                HasGrid = True
            End If

            MyGrid = dm_Grid
            MyForm = dm_Form
            oMaster = dm_MasterTable
            oDetails = dm_DetailTable

            If HasGrid Then
                AddHandler dm_Grid.OnAddNew, AddressOf dm_Grid_OnAddNew
                AddHandler dm_Grid.AfterColEdit, AddressOf dm_Grid_AfterColEdit
                AddHandler dm_Grid.KeyDownEvent, AddressOf dm_Grid_KeyDown
            End If
            AddHandler dm_Form.Paint, AddressOf MyForm_Paint

            ReadInitialValues()

            CN = dm_DSN
            X = 0

            For Each TxtCtrl In dm_Form.Controls
                If TypeName(TxtCtrl) = "TextBox" Or TypeName(TxtCtrl) = "ComboBox" Or TypeName(TxtCtrl) = "ListBox" Or TypeName(TxtCtrl) = "CheckBox" Or TypeName(TxtCtrl) = "RadioButton" Then
                    Col_ControlName.Add(TxtCtrl)
                    Col_ControlIndex.Add(X)
                End If
                X += 1
            Next TxtCtrl


            If HasGrid Then
                MyGrid.DataSource = oDetails
            End If

        End Sub

        Private Sub MyForm_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs)
            Static Flag As Byte = 0

            Flag += 1
            If Flag = 1 Then
                If Col_KeyValue.Count = 8 Then
                    Dim Num As Byte
                    For Num = 1 To 16
                        Col_KeyValue.Add(KeyLeavePos)
                    Next Num
                ElseIf Col_KeyValue.Count = 16 Then
                    Dim Num As Byte
                    For Num = 1 To 8
                        Col_KeyValue.Add(KeyLeavePos)
                    Next Num
                End If
                If Col_GridKeyValue.Count = 7 Then
                    Dim Num As Byte
                    For Num = 1 To 14
                        Col_GridKeyValue.Add(KeyLeavePos)
                    Next Num
                ElseIf Col_GridKeyValue.Count = 14 Then
                    Dim Num As Byte
                    For Num = 1 To 7
                        Col_GridKeyValue.Add(KeyLeavePos)
                    Next Num
                End If

            End If

            If HelpRtnID <> "" Then
                If UCase(sender.Controls(Col_KeyValue(8)).Name) = UCase(HelpIdSender) Then
                    sender.Controls(Col_KeyValue(8)).Text = HelpRtnID
                    sender.Controls(Col_KeyValue(1)).Text = HelpRtnName
                ElseIf UCase(sender.Controls(Col_KeyValue(16)).Name) = UCase(HelpIdSender) Then
                    sender.Controls(Col_KeyValue(16)).Text = HelpRtnID
                    sender.Controls(Col_KeyValue(9)).Text = HelpRtnName
                ElseIf UCase(sender.Controls(Col_KeyValue(24)).Name) = UCase(HelpIdSender) Then
                    sender.Controls(Col_KeyValue(24)).Text = HelpRtnID
                    sender.Controls(Col_KeyValue(17)).Text = HelpRtnName
                ElseIf UCase(MyGrid.Name) + "_" + UCase(MyGrid.Columns(Col_GridKeyValue(6)).DataField) = UCase(HelpIdSender) Then
                    If Col_GridKeyValue(6) + 2 <= MyGrid.Columns.Count Then
                        MyGrid.Col = Col_GridKeyValue(6) + 2
                    Else
                        MyGrid.Col = Col_GridKeyValue(6) - 2
                    End If
                    MyGrid.Columns(Col_GridKeyValue(6)).Value = HelpRtnID
                    MyGrid.Columns(Col_GridKeyValue(6) + 1).Value = HelpRtnName
                ElseIf UCase(MyGrid.Name) + "_" + UCase(MyGrid.Columns(Col_GridKeyValue(13)).DataField) = UCase(HelpIdSender) Then
                    If Col_GridKeyValue(13) + 2 <= MyGrid.Columns.Count Then
                        MyGrid.Col = Col_GridKeyValue(13) + 2
                    Else
                        MyGrid.Col = Col_GridKeyValue(13) - 2
                    End If
                    MyGrid.Columns(Col_GridKeyValue(13)).Value = HelpRtnID
                    MyGrid.Columns(Col_GridKeyValue(13) + 1).Value = HelpRtnName
                ElseIf UCase(MyGrid.Name) + "_" + UCase(MyGrid.Columns(Col_GridKeyValue(20)).DataField) = UCase(HelpIdSender) Then
                    If Col_GridKeyValue(20) + 2 <= MyGrid.Columns.Count Then
                        MyGrid.Col = Col_GridKeyValue(20) + 2
                    Else
                        MyGrid.Col = Col_GridKeyValue(20) - 2
                    End If
                    MyGrid.Columns(Col_GridKeyValue(20)).Value = HelpRtnID
                    MyGrid.Columns(Col_GridKeyValue(20) + 1).Value = HelpRtnName
                End If
            End If

        End Sub

        Public Sub PopulateForm(ByRef dm_Form As System.Windows.Forms.Form, ByRef dm_MasterTable As ADODB.Recordset, Optional ByRef dm_Grid As AxMSDataGridLib.AxDataGrid = Nothing, Optional ByRef dm_DetailTable As ADODB.Recordset = Nothing)
            Dim Num As Byte
            Dim Num2 As Integer
            Dim CtrlName As String
            Dim myCheckBox As System.Windows.Forms.CheckBox
            Dim myRadioButton As System.Windows.Forms.RadioButton
            On Error Resume Next 'keep it very important


            For Num = 1 To Col_ControlName.Count()
                CtrlName = Col_ControlName(Num).Name
                If UCase(Left(CtrlName, 1)) <> "X" Then
                    'If Not IsDBNull(dm_MasterTable(CtrlName).Value) Then
                    'If dm_MasterTable(CtrlName).Value <> "" Then
                    If TypeName(dm_Form.Controls(Col_ControlIndex(Num))) = "TextBox" Or TypeName(dm_Form.Controls(Col_ControlIndex(Num))) = "ComboBox" Or TypeName(dm_Form.Controls(Col_ControlIndex(Num))) = "ListBox" Then
                        dm_Form.Controls(Col_ControlIndex(Num)).Text = dm_MasterTable(CtrlName).Value
                    ElseIf TypeName(dm_Form.Controls(Col_ControlIndex(Num))) = "CheckBox" Then
                        myCheckBox = dm_Form.Controls(Col_ControlIndex(Num))
                        myCheckBox.Checked = IIf(dm_MasterTable(CtrlName).Value = True, True, False)
                    Else 'RadioButton
                        myRadioButton = dm_Form.Controls(Col_ControlIndex(Num))
                        myRadioButton.Checked = IIf(dm_MasterTable(CtrlName).Value = True, True, False)
                    End If
                    'End If
                    'End If
                End If
            Next Num

            For Num = 1 To Col_KeyText.Count
                Col_KeyText.Remove(Num)
            Next Num
            For Num = 1 To Col_KeyFields.Count / 2
                Col_KeyText.Add(dm_Form.Controls(Col_KeyFields(Num * 2)).Text)
            Next Num

            For Num = 1 To Col_KeyValue.Count() / 8
                Num2 = ((Num - 1) * 8) + 1
                dm_Form.Controls(Col_KeyValue(Num2)).Text = GetRelatedValue(Col_KeyValue(Num2 + 2), Col_KeyValue(Num2 + 3), dm_MasterTable(Col_KeyValue(Num2 + 4)).Value, Col_KeyValue(Num2 + 5))
            Next Num
            If HasGrid Then
                PopulateGrid(dm_Grid, dm_MasterTable, dm_DetailTable)
            End If
            'dm_Form.Controls(KeyLeavePos).Focus()
        End Sub

        Public Sub GoFirst(ByRef dm_Form As System.Windows.Forms.Form, ByRef dm_MasterTable As ADODB.Recordset, Optional ByRef dm_Grid As AxMSDataGridLib.AxDataGrid = Nothing, Optional ByRef dm_DetailTable As ADODB.Recordset = Nothing)
            Me.ClearData(dm_Form, dm_DetailTable)
            If Not dm_MasterTable.EOF Then
                dm_MasterTable.MoveFirst()
                Me.PopulateForm(dm_Form, dm_MasterTable, dm_Grid, dm_DetailTable)
            End If

        End Sub

        Public Sub GoLast(ByRef dm_Form As System.Windows.Forms.Form, ByRef dm_MasterTable As ADODB.Recordset, Optional ByRef dm_Grid As AxMSDataGridLib.AxDataGrid = Nothing, Optional ByRef dm_DetailTable As ADODB.Recordset = Nothing)
            Me.ClearData(dm_Form, dm_DetailTable)
            If Not dm_MasterTable.EOF Then
                dm_MasterTable.MoveLast()
                Me.PopulateForm(dm_Form, dm_MasterTable, dm_Grid, dm_DetailTable)
            End If

        End Sub

        Public Sub GoNext(ByRef dm_Form As System.Windows.Forms.Form, ByRef dm_MasterTable As ADODB.Recordset, Optional ByRef dm_Grid As AxMSDataGridLib.AxDataGrid = Nothing, Optional ByRef dm_DetailTable As ADODB.Recordset = Nothing)
            Me.ClearData(dm_Form, dm_DetailTable)
            If Not dm_MasterTable.EOF Then
                dm_MasterTable.MoveNext()
                If dm_MasterTable.EOF Then
                    dm_MasterTable.MoveLast()
                End If
                Me.PopulateForm(dm_Form, dm_MasterTable, dm_Grid, dm_DetailTable)
            End If
        End Sub

        Public Sub GoPrevious(ByRef dm_Form As System.Windows.Forms.Form, ByRef dm_MasterTable As ADODB.Recordset, Optional ByRef dm_Grid As AxMSDataGridLib.AxDataGrid = Nothing, Optional ByRef dm_DetailTable As ADODB.Recordset = Nothing)
            Me.ClearData(dm_Form, dm_DetailTable)
            If Not dm_MasterTable.BOF Then
                dm_MasterTable.MovePrevious()
                If dm_MasterTable.BOF Then
                    dm_MasterTable.MoveFirst()
                End If
                Me.PopulateForm(dm_Form, dm_MasterTable, dm_Grid, dm_DetailTable)
            End If
        End Sub

        Public Sub ClearData(ByRef dm_Form As System.Windows.Forms.Form, Optional ByVal dm_DetailTable As ADODB.Recordset = Nothing)
            Dim Num As Byte
            Dim myChechBox As System.Windows.Forms.CheckBox
            Dim myRadioButton As System.Windows.Forms.RadioButton

            For Num = 1 To Col_ControlIndex.Count()
                If TypeName(dm_Form.Controls(Col_ControlIndex(Num))) = "TextBox" Or TypeName(dm_Form.Controls(Col_ControlIndex(Num))) = "ComboBox" Or TypeName(dm_Form.Controls(Col_ControlIndex(Num))) = "ListBox" Then
                    dm_Form.Controls(Col_ControlIndex(Num)).Text = ""
                ElseIf TypeName(dm_Form.Controls(Col_ControlIndex(Num))) = "CheckBox" Then
                    myChechBox = dm_Form.Controls(Col_ControlIndex(Num))
                    myChechBox.Checked = False
                Else 'RadioButton
                    myRadioButton = dm_Form.Controls(Col_ControlIndex(Num))
                    myRadioButton.Checked = False
                End If
            Next Num
            If HasGrid = True Then
                dm_DetailTable.Filter = ""
            End If
        End Sub

        Public Sub NewRecord(ByRef dm_Form As System.Windows.Forms.Form, ByRef dm_MasterTable As ADODB.Recordset, Optional ByRef dm_Grid As AxMSDataGridLib.AxDataGrid = Nothing, Optional ByRef dm_DetailTable As ADODB.Recordset = Nothing)

            If HasGrid = True Then
                Me.ClearData(dm_Form, dm_DetailTable)
            Else
                Me.ClearData(dm_Form)
            End If
            If Not dm_MasterTable.EOF Then
                dm_MasterTable.MoveLast()
            End If
            If m_KeyLeaveField <> "" Then
                If Not dm_MasterTable.EOF Then
                    If Not IsDate(dm_MasterTable(m_KeyLeaveField).Value) And Val(dm_MasterTable(m_KeyLeaveField).Value) <> 0 Then
                        dm_Form.Controls(KeyLeavePos).Text = ZeroPad(dm_MasterTable(m_KeyLeaveField).Value + 1, Key_ZeroPad)
                    End If
                End If
                dm_Form.Controls(KeyLeavePos).Focus()
            End If

        End Sub

        Public Sub DeleteRecord(ByRef dm_Form As System.Windows.Forms.Form, ByRef dm_MasterTable As ADODB.Recordset, Optional ByRef dm_Grid As AxMSDataGridLib.AxDataGrid = Nothing, Optional ByRef dm_DetailTable As ADODB.Recordset = Nothing)
            Dim cFilter As String
            Dim Num As Byte

            cFilter = ""
            For Num = 1 To Col_KeyFields.Count / 2
                cFilter += dm_Form.Controls(Col_KeyFields(Num * 2)).Name + " = '" + dm_Form.Controls(Col_KeyFields(Num * 2)).Text + "'"
                If Num <> Col_KeyFields.Count / 2 Then
                    cFilter += " and "
                End If
            Next Num
            If m_MasterFlagField <> "" Then
                cFilter += " and " + m_MasterFlagField + " ='" + m_FlagValue + "'"
            End If

            If m_KeyFields = "" Then
                MsgBox("You must asign KeyField Property first ,so you can delete records", , "DC DataManger error msg")
                Return
            End If

            If cFilter <> "" Then
                If Not dm_MasterTable.EOF Then
                    dm_MasterTable.Filter = cFilter
                End If

                If Not dm_MasterTable.EOF Then
                    If MsgBox(aInitValues(20), MsgBoxStyle.YesNo) = MsgBoxResult.Yes And Not dm_MasterTable.BOF Then
                        dm_MasterTable.Delete()
                        dm_MasterTable.Filter = ""
                        If HasGrid Then
                            dm_DetailTable.MoveFirst()
                            Do While Not dm_DetailTable.EOF
                                dm_DetailTable.Delete()
                                dm_DetailTable.MoveNext()
                            Loop
                            Me.ClearData(dm_Form, dm_DetailTable)
                            dm_DetailTable.Requery()
                        End If
                        GoPrevious(dm_Form, dm_MasterTable, dm_Grid, dm_DetailTable)
                    End If
                End If
            End If
        End Sub

        Private Sub SubSaveData(ByRef dm_Form As System.Windows.Forms.Form, ByRef dm_MasterTable As ADODB.Recordset, Optional ByRef dm_Grid As AxMSDataGridLib.AxDataGrid = Nothing, Optional ByRef dm_DetailTable As ADODB.Recordset = Nothing)
            Dim Num As Byte
            Dim CtrlName As String
            Dim CtrlValue As String
            Dim myCheckBox As System.Windows.Forms.CheckBox
            Dim myRadioButton As System.Windows.Forms.RadioButton

            For Num = 1 To Col_ControlName.Count()
                CtrlName = Col_ControlName(Num).Name
                If UCase(Left(CtrlName, 1)) <> "X" Then
                    If TypeName(dm_Form.Controls(Col_ControlIndex(Num))) = "TextBox" Or TypeName(dm_Form.Controls(Col_ControlIndex(Num))) = "ComboBox" Or TypeName(dm_Form.Controls(Col_ControlIndex(Num))) = "ListBox" Then
                        CtrlValue = dm_Form.Controls(Col_ControlIndex(Num)).Text
                        If CtrlValue <> "" And Not IsDBNull(CtrlValue) Then
                            dm_MasterTable(CtrlName).Value = CtrlValue
                        ElseIf CtrlValue = "" Then
                            dm_MasterTable(CtrlName).Value = Nothing
                        End If
                    ElseIf TypeName(dm_Form.Controls(Col_ControlIndex(Num))) = "CheckBox" Then
                        myCheckBox = dm_Form.Controls(Col_ControlIndex(Num))
                        dm_MasterTable(CtrlName).Value = IIf(myCheckBox.Checked, True, False)
                    Else 'RadioButton
                        myRadioButton = dm_Form.Controls(Col_ControlIndex(Num))
                        dm_MasterTable(CtrlName).Value = IIf(myRadioButton.Checked, True, False)
                    End If

                End If
            Next Num

            If m_MasterFlagField <> "" Then
                dm_MasterTable(m_MasterFlagField).Value = m_FlagValue
            End If

        End Sub

        Public Sub HoldSaving(ByVal Mode As Boolean)
            If Mode = True Then
                m_HoldSaving = True
            Else : m_HoldSaving = False
            End If
        End Sub

        Public Sub ReleaseSaving(ByVal Mode As Boolean)
            If Mode = True Then
                m_ReleaseSaving = True
                m_HoldSaving = False
            Else : m_ReleaseSaving = False
            End If

            If m_ReleaseSaving Then
                SaveData(MyForm, oMaster, MyGrid, oDetails)
            End If
            m_ReleaseSaving = False
        End Sub

        Public Sub SaveData(ByRef dm_Form As System.Windows.Forms.Form, ByRef dm_MasterTable As ADODB.Recordset, Optional ByRef dm_Grid As AxMSDataGridLib.AxDataGrid = Nothing, Optional ByRef dm_DetailTable As ADODB.Recordset = Nothing)
            Dim cFilter As String
            Dim Num As Byte

            cFilter = ""

            If Not ValidateForm(dm_Form) Then Exit Sub
            If m_ReleaseSaving Then GoTo ReleaseSaving

            For Num = 1 To Col_KeyFields.Count / 2
                cFilter += dm_Form.Controls(Col_KeyFields(Num * 2)).Name + " = '" + dm_Form.Controls(Col_KeyFields(Num * 2)).Text + "'"
                If Num <> Col_KeyFields.Count / 2 Then
                    cFilter += " and "
                End If
            Next Num

            If m_MasterFlagField <> "" Then
                cFilter += " and " + m_MasterFlagField + " ='" + m_FlagValue + "'"
            End If

            If cFilter <> "" Then
                If Not dm_MasterTable.EOF Then
                    dm_MasterTable.Filter = cFilter
                End If

                If dm_MasterTable.EOF Then
                    dm_MasterTable.AddNew()
                End If

                Me.SubSaveData(dm_Form, dm_MasterTable)
ReleaseSaving:
                If Not m_HoldSaving Then
                    dm_MasterTable.Update()
                    dm_MasterTable.Filter = "" ' order of line is important
                    Me.ClearData(dm_Form, dm_DetailTable)
                End If
            End If

            If m_KeyLeaveField <> "" And Not m_HoldSaving Then
                If Not dm_MasterTable.EOF Then
                    dm_MasterTable.MoveLast()
                    If Not IsDate(dm_MasterTable(m_KeyLeaveField).Value) And Val(dm_MasterTable(m_KeyLeaveField).Value) <> 0 Then
                        dm_Form.Controls(KeyLeavePos).Text = ZeroPad(dm_MasterTable(m_KeyLeaveField).Value + 1, Key_ZeroPad)
                    End If
                End If
            End If

        End Sub

        Public Sub KeyLeave(ByRef dm_Form As System.Windows.Forms.Form, ByRef dm_MasterTable As ADODB.Recordset, Optional ByRef dm_Grid As AxMSDataGridLib.AxDataGrid = Nothing, Optional ByRef dm_DetailTable As ADODB.Recordset = Nothing)
            Dim Num As Byte
            Dim cFilter As String
            Dim ColFields As New Collection()

            cFilter = ""

            For Num = 1 To Col_KeyFields.Count / 2
                cFilter += dm_Form.Controls(Col_KeyFields(Num * 2)).Name + " = '" + dm_Form.Controls(Col_KeyFields(Num * 2)).Text + "'"
                ColFields.Add(dm_Form.Controls(Col_KeyFields(Num * 2)).Text)
                If Num <> Col_KeyFields.Count / 2 Then
                    cFilter += " and "
                End If
            Next Num
            If m_MasterFlagField <> "" Then
                cFilter += " and " + m_MasterFlagField + " ='" + m_FlagValue + "'"
            End If

            If cFilter <> "" Then
                If Not dm_MasterTable.EOF Then
                    dm_MasterTable.Filter = cFilter
                End If

                If HasGrid Then
                    dm_DetailTable.Filter = ""
                    Me.ClearData(dm_Form, dm_DetailTable)
                Else
                    Me.ClearData(dm_Form)
                End If

                If Not dm_MasterTable.EOF Then
                    Me.PopulateForm(dm_Form, dm_MasterTable, dm_Grid, dm_DetailTable)
                Else
                    For Num = 1 To Col_KeyFields.Count / 2
                        dm_Form.Controls(Col_KeyFields(Num * 2)).Text = ColFields(Num)
                    Next Num
                End If
                dm_MasterTable.Filter = ""
            End If

        End Sub

        Public Function GetValue(ByVal str_Table As String, ByVal str_Key As String, ByVal str_value As String, ByVal str_RetField As String) As String
            Dim oRecSet As New ADODB.Recordset()
            oRecSet.Open(str_Table, CN, oRecSet.CursorType.adOpenKeyset, oRecSet.LockType.adLockOptimistic)
            If str_value <> "" Then
                oRecSet.MoveFirst()
                oRecSet.Find(str_Key + " = '" + str_value + "'")
                If Not oRecSet.EOF Then
                    Return oRecSet(str_RetField).Value
                End If
            End If
            oRecSet.Close()
        End Function

        Public Sub AddRelatedValue(ByRef str_Table As String, ByVal str_Key As String, ByVal str_Control As String, ByVal str_RetValue As String, ByVal str_RetControl As String, Optional ByVal n_ZeroPad As Byte = 0)
            Dim Num As Byte
            Dim pos As Byte
            Static Flag As Byte = 0

            Flag += 1
            pos = ((Flag - 1) * 8) + 1
            For Num = 1 To Col_ControlName.Count
                If UCase(Col_ControlName(Num).Name) = UCase(str_RetControl) Then
                    Col_KeyValue.Add(Col_ControlIndex(Num))
                    Exit For
                End If
            Next Num

            For Num = 1 To Col_ControlName.Count
                If UCase(Col_ControlName(Num).Name) = UCase(str_Control) Then
                    If Flag = 1 Then
                        Dim txt_control As TextBox = CType(MyForm.Controls(Col_ControlIndex(Num)), TextBox)
                        AddHandler txt_control.Leave, AddressOf MyTextBox1_Leave
                        AddHandler txt_control.KeyDown, AddressOf MyTextBox1_KeyDown
                    ElseIf Flag = 2 Then
                        Dim txt_control As TextBox = CType(MyForm.Controls(Col_ControlIndex(Num)), TextBox)
                        AddHandler txt_control.Leave, AddressOf MyTextBox2_Leave
                        AddHandler txt_control.KeyDown, AddressOf MyTextBox2_KeyDown
                    ElseIf Flag = 3 Then
                        Dim txt_control As TextBox = CType(MyForm.Controls(Col_ControlIndex(Num)), TextBox)
                        AddHandler txt_control.Leave, AddressOf MyTextBox3_Leave
                        AddHandler txt_control.KeyDown, AddressOf MyTextBox3_KeyDown
                    End If
                    Exit For
                End If
            Next Num

            Col_KeyValue.Add(str_RetControl)
            Col_KeyValue.Add(str_Table)
            Col_KeyValue.Add(str_Key)
            Col_KeyValue.Add(str_Control)
            Col_KeyValue.Add(str_RetValue)
            Col_KeyValue.Add(n_ZeroPad)
            Col_KeyValue.Add(Col_ControlIndex(Num))
        End Sub


        Public Sub AddGridRelatedValue(ByVal str_Table As String, ByVal str_TableKey As String, ByVal str_Column As String, ByVal str_TableRetField As String, ByVal str_GridRetColumn As String, Optional ByVal n_ZeroPad As Byte = 0)
            Dim Num As Byte

            Col_GridKeyValue.Add(str_Table)
            Col_GridKeyValue.Add(str_TableKey)
            Col_GridKeyValue.Add(str_Column)
            Col_GridKeyValue.Add(str_TableRetField)
            Col_GridKeyValue.Add(str_GridRetColumn)

            For Num = 0 To MyGrid.Columns.Count - 1
                If UCase(MyGrid.Columns(Num).DataField) = UCase(str_Column) Then
                    Col_GridKeyValue.Add(Num)
                    Col_GridKeyValue.Add(n_ZeroPad)
                    Exit Sub
                End If
            Next
            Col_GridKeyValue.Add(-1)
            Col_GridKeyValue.Add(n_ZeroPad)
        End Sub


        Private Function GetRelatedValue(ByVal str_Table As String, ByVal str_Key As String, ByVal str_value As String, ByVal str_RetField As String) As String
            Dim oRecSet As New ADODB.Recordset()

            oRecSet.Open(str_Table, CN, oRecSet.CursorType.adOpenKeyset, oRecSet.LockType.adLockOptimistic)
            If str_value <> "" Then
                oRecSet.MoveFirst()
                oRecSet.Find(str_Key + " = '" + str_value + "'")
                If Not oRecSet.EOF Then
                    Return oRecSet(str_RetField).Value
                End If
            End If
            oRecSet.Close()
        End Function

        Public Sub FlagField(ByVal str_MasterFlagField As String, ByVal str_FlagValue As String, Optional ByVal str_DetailFlagField As String = "")
            Dim Num As Byte

            m_MasterFlagField = str_MasterFlagField
            m_FlagValue = str_FlagValue

            If HasGrid Then
                m_DetailFlagField = str_DetailFlagField
                For Num = 0 To MyGrid.Columns.Count - 1
                    If UCase(MyGrid.Columns(Num).DataField) = UCase(m_DetailFlagField) Then
                        MyGrid.Columns(Num).Visible = False
                    End If
                Next Num
            End If

        End Sub

        Public Sub KeyFields(ByVal str_KeyFields As String)
            Dim Num, Num2 As Integer
            Dim StartPos As Integer
            Dim StrPart As String
            Dim Index As Integer

            m_KeyFields = str_KeyFields
            str_KeyFields += "+"
            StartPos = 1
            Index = 1

            For Num = 1 To Len(str_KeyFields)
                If Mid(str_KeyFields, Num, 1) = "+" Then
                    StrPart = Mid(str_KeyFields, StartPos, Num - StartPos)
                    Col_KeyFields.Add(StrPart)
                    For Num2 = 1 To Col_ControlName.Count
                        If UCase(Col_ControlName(Num2).Name) = UCase(Col_KeyFields(Index)) Then
                            Col_KeyFields.Add(Col_ControlIndex(Num2))
                            Exit For
                        End If
                    Next Num2
                    StartPos = Num + 1
                    Index = Index + 2
                End If
            Next Num

        End Sub
        Public Sub SetLink(ByVal str_MasterFields As String, ByVal str_DetailFields As String)
            Dim Num, Num2 As Integer
            Dim StartPos As Integer
            Dim StrPart As String
            Dim Index As Integer

            str_MasterFields += "+"
            str_DetailFields += "+"
            StartPos = 1
            Index = 1

            For Num = 1 To Len(str_MasterFields)
                If Mid(str_MasterFields, Num, 1) = "+" Then
                    StrPart = Mid(str_MasterFields, StartPos, Num - StartPos)
                    Col_MasterFields.Add(StrPart)
                    StartPos = Num + 1
                End If
            Next Num


            StartPos = 1
            Index = 1
            For Num = 1 To Len(str_DetailFields)
                If Mid(str_DetailFields, Num, 1) = "+" Then
                    StrPart = Mid(str_DetailFields, StartPos, Num - StartPos)
                    Col_DetailFields.Add(StrPart)
                    StartPos = Num + 1
                    Index = Index + 1
                End If
            Next Num

            For Num = 0 To MyGrid.Columns.Count - 1
                For Num2 = 1 To Col_DetailFields.Count
                    If UCase(MyGrid.Columns(Num).DataField) = UCase(Col_DetailFields(Num2)) Then
                        MyGrid.Columns(Num).Visible = False
                    End If
                Next Num2
            Next Num

        End Sub

        Private Sub PopulateGrid(ByRef dm_Grid As AxMSDataGridLib.AxDataGrid, ByRef dm_MasterTable As ADODB.Recordset, ByRef dm_DetailTable As ADODB.Recordset)
            Dim cFilter As String = ""
            Dim Num As Byte
            Dim X As Byte
            Dim X_ As Byte
            Dim oRecSet As New ADODB.Recordset()
            Dim Pos As Byte


            For Num = 1 To Col_MasterFields.Count
                cFilter += Col_DetailFields(Num) + " = '" + dm_MasterTable(Col_MasterFields(Num)).Value + "'"
                If Num <> Col_MasterFields.Count Then
                    cFilter += " and "
                End If
            Next Num

            dm_DetailTable.Filter = cFilter

            For X = 1 To Col_GridKeyValue.Count / 6
                X_ = (X - 1) * 6 + 1
                oRecSet.Open(Col_GridKeyValue(X_), CN, oRecSet.CursorType.adOpenKeyset, oRecSet.LockType.adLockOptimistic)
                dm_DetailTable.MoveFirst()
                Do While Not dm_DetailTable.EOF
                    oRecSet.MoveFirst()
                    oRecSet.Find(Col_GridKeyValue(X_ + 1) + " = '" + dm_DetailTable.Fields(Col_GridKeyValue(X_ + 2)).Value + "'")
                    If Not oRecSet.EOF Then
                        dm_DetailTable.Fields(Col_GridKeyValue(X_ + 4)).Value = oRecSet(Col_GridKeyValue(X_ + 3)).Value()
                    End If
                    dm_DetailTable.MoveNext()

                Loop
                oRecSet.Close()
            Next X

            dm_Grid.AllowAddNew = True
            dm_Grid.AllowDelete = True
        End Sub

        Public Sub KeyLeaveField(ByRef dm_MasterTable As ADODB.Recordset, ByVal str_KeyLeaveField As String, Optional ByVal n_ZeroPad As Byte = 0)
            Dim Num As Byte

            m_KeyLeaveField = str_KeyLeaveField
            For Num = 1 To Col_KeyFields.Count / 2
                If UCase(Col_KeyFields(Num * 2 - 1)) = UCase(m_KeyLeaveField) Then
                    KeyLeavePos = Col_KeyFields(Num * 2)
                    Exit For
                End If
            Next Num

            Key_ZeroPad = n_ZeroPad
            AddHandler MyForm.Controls(KeyLeavePos).Leave, AddressOf MyTextBox0_Leave

        End Sub

        Public Sub RequiredFields(ByVal str_RequiredFields As String, Optional ByVal b_ShowMsgBox As Boolean = True, Optional ByVal str_Msg As String = "Uncomplete Entries")
            Dim Num, Num2 As Integer
            Dim StartPos As Integer
            Dim StrPart As String


            RequiredFields_Msg = str_Msg
            RequiredFields_ShowMsg = b_ShowMsgBox

            str_RequiredFields += "+"
            StartPos = 1

            For Num = 1 To Len(str_RequiredFields)
                If Mid(str_RequiredFields, Num, 1) = "+" Then
                    StrPart = Mid(str_RequiredFields, StartPos, Num - StartPos)
                    For Num2 = 1 To Col_ControlName.Count
                        If UCase(Col_ControlName(Num2).Name) = UCase(StrPart) Then
                            Col_RequiredFields.Add(Col_ControlIndex(Num2))
                            Exit For
                        End If
                    Next Num2
                    StartPos = Num + 1
                End If
            Next Num

        End Sub

        Public Function ValidateForm(ByRef dm_Form As System.Windows.Forms.Form) As Boolean
            Dim Num As Byte

            For Num = 1 To Col_RequiredFields.Count
                If dm_Form.Controls(Col_RequiredFields(Num)).Text = "" Then
                    dm_Form.Controls(Col_RequiredFields(Num)).Focus()
                    If RequiredFields_ShowMsg Then
                        MsgBox(RequiredFields_Msg, MsgBoxStyle.OKOnly)
                    End If
                    Return False
                    Exit For
                End If
            Next Num
            Return True

        End Function

        Public Sub Search(ByRef dm_Form As System.Windows.Forms.Form, ByRef dm_MasterTable As ADODB.Recordset, Optional ByRef dm_Grid As AxMSDataGridLib.AxDataGrid = Nothing, Optional ByRef dm_DetailTable As ADODB.Recordset = Nothing)
            Static SearchFlag As Integer = 1
            Dim cFilter As String = ""
            Dim Num As Byte
            Dim ControlText As String
            Dim myChechBox As System.Windows.Forms.CheckBox
            Dim myRadioButton As System.Windows.Forms.RadioButton

            SearchFlag += 1

            If SearchFlag Mod 2 = 0 Then
                'MyForm.Text += " Filter Mode"
                Me.ClearData(dm_Form, dm_DetailTable)
            Else
                'MyForm.Text = Mid(MyForm.Text, 1, Len(MyForm.Text) - 12)
                For Num = 1 To Col_ControlIndex.Count()
                    If Col_ControlIndex(Num) <> KeyLeavePos And UCase(Left(Col_ControlName(Num).Name, 1)) <> "X" Then
                        If TypeName(dm_Form.Controls(Col_ControlIndex(Num))) = "TextBox" Or TypeName(dm_Form.Controls(Col_ControlIndex(Num))) = "ComboBox" Or TypeName(dm_Form.Controls(Col_ControlIndex(Num))) = "ListBox" Then
                            ControlText = dm_Form.Controls(Col_ControlIndex(Num)).Text
                            If ControlText <> "" Then
                                cFilter += dm_Form.Controls(Col_ControlIndex(Num)).Name + " like '*" + ControlText + "*' and "
                            End If
                        ElseIf TypeName(dm_Form.Controls(Col_ControlIndex(Num))) = "CheckBox" Then
                            myChechBox = dm_Form.Controls(Col_ControlIndex(Num))
                            ControlText = IIf(myChechBox.Checked, "True", "False")
                            If ControlText <> "" Then
                                cFilter += dm_Form.Controls(Col_ControlIndex(Num)).Name + " = " + ControlText + " and "
                            End If
                        Else 'RadioButton
                            myRadioButton = dm_Form.Controls(Col_ControlIndex(Num))
                            ControlText = IIf(myRadioButton.Checked, "True", "False")
                            If ControlText <> "" Then
                                cFilter += dm_Form.Controls(Col_ControlIndex(Num)).Name + " = " + ControlText + " and "
                            End If
                        End If

                    End If
                Next Num
                If cFilter <> "" Then
                    cFilter = Mid(cFilter, 1, cFilter.Length - 4)
                End If
                dm_MasterTable.Filter = cFilter
                Me.PopulateForm(dm_Form, dm_MasterTable, dm_Grid, dm_DetailTable)
            End If
        End Sub
        Private Sub dm_Grid_OnAddNew(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim Num As Byte

            sender.DataSource.AddNew()
            If m_DetailFlagField <> "" Then
                sender.DataSource.Fields(m_DetailFlagField).value = m_FlagValue
            End If
            For Num = 1 To Col_KeyText.Count
                sender.DataSource.Fields(Col_KeyFields((Num * 2) - 1)).Value() = Col_KeyText(Num)
            Next Num
            sender.DataSource.Refresh()
        End Sub

        Private Sub dm_Grid_AfterColEdit(ByVal sender As Object, ByVal e As AxMSDataGridLib.DDataGridEvents_AfterColEditEvent)
            Dim Pos1 As Integer = -1
            Dim Pos2 As Integer = -1
            Dim Pos3 As Integer = -1

            On Error Resume Next ' Keep it
            Pos1 = Col_GridKeyValue(6)
            Pos2 = Col_GridKeyValue(13)
            Pos3 = Col_GridKeyValue(20)

            If e.colIndex = Pos1 Then
                MyGrid.Columns(Pos1).Value = ZeroPad(MyGrid.Columns(Pos1).Value, Col_GridKeyValue(7))
                MyGrid.Columns(Pos1 + 1).Value = GetValue(Col_GridKeyValue(1), Col_GridKeyValue(2), MyGrid.Columns(Pos1).Value, Col_GridKeyValue(4))
            ElseIf e.colIndex = Pos2 Then
                MyGrid.Columns(Pos2).Value = ZeroPad(MyGrid.Columns(Pos2).Value, Col_GridKeyValue(14))
                MyGrid.Columns(Pos2 + 1).Value = GetValue(Col_GridKeyValue(8), Col_GridKeyValue(9), MyGrid.Columns(Pos2).Value, Col_GridKeyValue(11))
            ElseIf e.colIndex = Pos3 Then
                MyGrid.Columns(Pos3).Value = ZeroPad(MyGrid.Columns(Pos3).Value, Col_GridKeyValue(21))
                MyGrid.Columns(Pos3 + 1).Value = GetValue(Col_GridKeyValue(14), Col_GridKeyValue(15), MyGrid.Columns(Pos3).Value, Col_GridKeyValue(17))
            End If

        End Sub


        Private Sub dm_Grid_KeyDown(ByVal sender As Object, ByVal e As AxMSDataGridLib.DDataGridEvents_KeyDownEvent)
            Dim Pos1 As Integer = -1
            Dim Pos2 As Integer = -1
            Dim Pos3 As Integer = -1
            Dim oHelpForm As New DataHelpForm()
            Dim Num As Byte

            On Error Resume Next ' Keep it
            If e.keyCode = Keys.F1 Then
                Pos1 = Col_GridKeyValue(6)
                Pos2 = Col_GridKeyValue(13)
                Pos3 = Col_GridKeyValue(20)

                If sender.Col = Pos1 Then
                    Num = 1
                    HelpFile = Col_GridKeyValue(Num)
                    HelpID = Col_GridKeyValue(Num + 1)
                    HelpName = Col_GridKeyValue(Num + 3)
                    HelpIdSender = sender.Name + "_" + sender.Columns(Pos1).DataField
                    oHelpForm.Show()
                ElseIf sender.Col = Pos2 Then
                    Num = 8
                    HelpFile = Col_GridKeyValue(Num)
                    HelpID = Col_GridKeyValue(Num + 1)
                    HelpName = Col_GridKeyValue(Num + 3)
                    HelpIdSender = sender.Name + "_" + sender.Columns(Pos2).DataField
                    oHelpForm.Show()
                ElseIf sender.Col = Pos3 Then
                    Num = 15
                    HelpFile = Col_GridKeyValue(Num)
                    HelpID = Col_GridKeyValue(Num + 1)
                    HelpName = Col_GridKeyValue(Num + 3)
                    HelpIdSender = sender.Name + "_" + sender.Columns(Pos3).DataField
                    oHelpForm.Show()
                End If
            End If
        End Sub

        Public Sub NavigationButtons(ByVal dm_First As String, ByVal dm_Previous As String, ByVal dm_Next As String, ByVal dm_Last As String)
            Dim MyButton As New System.Windows.Forms.Control()
            Dim Num As Byte = 0
            Dim cButton As String
            On Error Resume Next

            For Each MyButton In MyForm.Controls
                If UCase(MyButton.Name) = UCase(dm_First) Then
                    AddHandler MyForm.Controls(Num).Click, AddressOf FirstButton_Click
                ElseIf UCase(MyButton.Name) = UCase(dm_Previous) Then
                    AddHandler MyForm.Controls(Num).Click, AddressOf PreviousButton_Click
                ElseIf UCase(MyButton.Name) = UCase(dm_Next) Then
                    AddHandler MyForm.Controls(Num).Click, AddressOf NextButton_Click
                ElseIf UCase(MyButton.Name) = UCase(dm_Last) Then
                    AddHandler MyForm.Controls(Num).Click, AddressOf LastButton_Click
                End If
                Num += 1
            Next MyButton

        End Sub

        Public Sub ManipulationButtons(ByVal dm_Save As String, ByVal dm_New As String, ByVal dm_Delete As String, ByVal dm_Close As String, Optional ByVal dm_Search As String = Nothing)
            Dim MyButton As New System.Windows.Forms.Control()
            Dim Num As Byte = 0
            On Error Resume Next

            For Each MyButton In MyForm.Controls
                If UCase(MyButton.Name) = UCase(dm_Save) Then
                    AddHandler MyForm.Controls(Num).Click, AddressOf SaveButton_Click
                ElseIf UCase(MyButton.Name) = UCase(dm_New) Then
                    AddHandler MyForm.Controls(Num).Click, AddressOf NewButton_Click
                ElseIf UCase(MyButton.Name) = UCase(dm_Delete) Then
                    AddHandler MyForm.Controls(Num).Click, AddressOf DeleteButton_Click
                ElseIf UCase(MyButton.Name) = UCase(dm_Close) Then
                    AddHandler MyForm.Controls(Num).Click, AddressOf CloseButton_Click
                ElseIf UCase(MyButton.Name) = UCase(dm_Search) Then
                    AddHandler MyForm.Controls(Num).Click, AddressOf SearchButton_Click
                End If
                Num += 1
            Next MyButton
        End Sub

        ''''''''''First Button Handles
        Private Sub FirstButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
            GoFirst(MyForm, oMaster, MyGrid, oDetails)
        End Sub

        'Previous Button Handles
        Private Sub PreviousButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
            GoPrevious(MyForm, oMaster, MyGrid, oDetails)
        End Sub
        ''''''''''Next Button Handles
        Private Sub NextButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
            GoNext(MyForm, oMaster, MyGrid, oDetails)
        End Sub
        '''''''''''Last Button Hadles
        Private Sub LastButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
            GoLast(MyForm, oMaster, MyGrid, oDetails)
        End Sub
        ''''''''''Save Button Handles
        Private Sub SaveButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
            SaveData(MyForm, oMaster, MyGrid, oDetails)
        End Sub
        ''''''''''New Button Handles
        Private Sub NewButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
            NewRecord(MyForm, oMaster, MyGrid, oDetails)
        End Sub

        '''''Delete Button Handles
        Private Sub DeleteButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
            DeleteRecord(MyForm, oMaster, MyGrid, oDetails)
        End Sub
        ''''''Close Button Handles
        Private Sub CloseButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
            CloseForm(MyForm, oMaster, oDetails)
        End Sub
        ''''''Searrch Button Handles
        Private Sub SearchButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
            Search(MyForm, oMaster, MyGrid, oDetails)
        End Sub
        '''''' Close Form Handles
        Private Sub CloseForm(ByRef dm_Form As System.Windows.Forms.Form, ByRef dm_MasterTable As ADODB.Recordset, Optional ByRef dm_DetailTable As ADODB.Recordset = Nothing)
            On Error Resume Next

            dm_MasterTable.Close()
            If HasGrid Then
                dm_DetailTable.Close()
            End If
            dm_Form.Close()
        End Sub
        Private Sub MyTextBox1_Leave(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim Num As Byte

            Num = 1
            If sender.text <> "" Then
                sender.Text = ZeroPad(sender.Text, Col_KeyValue(Num + 6))
                MyForm.Controls(Col_KeyValue(Num)).Text = GetValue(Col_KeyValue(Num + 2), Col_KeyValue(Num + 3), sender.Text, Col_KeyValue(Num + 5))
            End If
        End Sub

        Private Sub MyTextBox1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
            If e.KeyCode = Keys.F1 Then
                Dim Num As Byte = 1
                Dim oHelpForm As New DataHelpForm()
                HelpFile = Col_KeyValue(Num + 2)
                HelpID = Col_KeyValue(Num + 3)
                HelpName = Col_KeyValue(Num + 5)
                HelpIdSender = sender.Name
                oHelpForm.Show()
            End If
        End Sub

        Private Sub MyTextBox2_Leave(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim Num As Byte

            Num = 9
            If sender.text <> "" Then
                sender.Text = ZeroPad(sender.Text, Col_KeyValue(Num + 6))
                MyForm.Controls(Col_KeyValue(Num)).Text = GetValue(Col_KeyValue(Num + 2), Col_KeyValue(Num + 3), sender.Text, Col_KeyValue(Num + 5))
            End If
        End Sub

        Private Sub MyTextBox2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
            If e.KeyCode = Keys.F1 Then
                Dim Num As Byte = 9
                Dim oHelpForm As New DataHelpForm()
                HelpFile = Col_KeyValue(Num + 2)
                HelpID = Col_KeyValue(Num + 3)
                HelpName = Col_KeyValue(Num + 5)
                HelpIdSender = sender.Name
                oHelpForm.Show()
            End If
        End Sub

        Private Sub MyTextBox3_Leave(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim Num As Byte

            Num = 17
            If sender.text <> "" Then
                sender.Text = ZeroPad(sender.Text, Col_KeyValue(Num + 6))
                MyForm.Controls(Col_KeyValue(Num)).Text = GetValue(Col_KeyValue(Num + 2), Col_KeyValue(Num + 3), sender.Text, Col_KeyValue(Num + 5))
            End If
        End Sub

        Private Sub MyTextBox3_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
            If e.KeyCode = Keys.F1 Then
                Dim Num As Byte = 17
                Dim oHelpForm As New DataHelpForm()
                HelpFile = Col_KeyValue(Num + 2)
                HelpID = Col_KeyValue(Num + 3)
                HelpName = Col_KeyValue(Num + 5)
                HelpIdSender = sender.Name
                oHelpForm.Show()
            End If
        End Sub


        Private Sub MyTextBox0_Leave(ByVal sender As Object, ByVal e As System.EventArgs)
            If Key_ZeroPad <> 0 Then
                sender.Text = ZeroPad(sender.Text, Key_ZeroPad)
            End If
            KeyLeave(MyForm, oMaster, MyGrid, oDetails)
        End Sub

        Private Sub ReadInitialValues()
            Dim MyFile As String
            Dim MyFullPathFile As String
            Dim oInitValues As New InitValues()
            Dim FileNum As Byte

            Dim WinSys As String
            Dim WshShell As New Object()
            On Error Resume Next

            WshShell = CreateObject("WScript.Shell")
            WinSys = WshShell.SpecialFolders("Fonts")
            WinSys = Mid(WinSys, 1, Len(WinSys) - 5)
            WinSys += "System32\"
            FileNum = FreeFile()
            MyFullPathFile = WinSys + "DCDM30_Lang.dll"
            MyFile = Dir(WinSys + "DCDM30_Lang.dll")

            FileOpen(1, MyFullPathFile, OpenMode.Random, OpenAccess.ReadWrite, OpenShare.Shared, 1000)
            If UCase(MyFile) = "DCDM30_LANG.DLL" Then
                FileGet(FileNum, oInitValues, 1)
                aInitValues(0) = oInitValues.Help_Caption.Trim + " "
                aInitValues(1) = oInitValues.Help_DataEntryType.Trim + " "
                aInitValues(2) = oInitValues.Help_MaxLenght.Trim + " "
                aInitValues(3) = oInitValues.Help_Required.Trim + " "
                aInitValues(4) = oInitValues.Help_HasDataHelp.Trim + " "
                aInitValues(5) = oInitValues.Help_Const_NotDefined.Trim + " "
                aInitValues(6) = oInitValues.Help_Const_Characters.Trim + " "
                aInitValues(7) = oInitValues.Help_Const_Description.Trim + " "
                aInitValues(8) = oInitValues.Help_Const_Yes.Trim + " "
                aInitValues(9) = oInitValues.Help_Const_No.Trim + " "
                ' Keep 5 room for Future add
                aInitValues(15) = oInitValues.DataHelp_Caption.Trim + " "
                aInitValues(16) = oInitValues.DataHelp_Id.Trim + " "
                aInitValues(17) = oInitValues.DataHelp_Name.Trim + " "
                ' Keep 2 room for Future add
                aInitValues(20) = oInitValues.Delete_Message.Trim + " "
                aInitValues(20) = "Are you sure you want to delete this record"
            End If
        End Sub
    End Class
End Namespace