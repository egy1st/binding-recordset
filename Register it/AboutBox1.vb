Public NotInheritable Class AboutBox1

    Private Sub AboutBox1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ' Set the title of the form.
        
        
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
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
    End Sub
End Class
