Option Compare Database

Private Sub cmdLogin_Click()
On Error Resume Next
'Mengecek apakah username sudah dimasukan ke combo box

    If IsNull(Me.cboEmployee) Or Me.cboEmployee = "" Then
            MsgBox "Silahkan Masukan Username ...!", vbInformation, "Konfirmasi Username !"
            Me.cboEmployee.SetFocus
        Exit Sub
    End If

'Mengecek apakah password sudah diisi

    If IsNull(Me.txtPassword) Or Me.txtPassword = "" Then
            MsgBox "Anda Harus Mengisi Password.", vbInformation, "Konfirmasi Password !"
            Me.txtPassword.SetFocus
            Me.txtPassword = ""
        Exit Sub
    End If

'Mengecek password pada Tabel User apakah cocok dengan username pada cboEmployee

    If Me.txtPassword.Value = DLookup("password", "tbl_user", "[ID]=" & Me.cboEmployee.Value) Then
        Id = Me.cboEmployee.Value

'Progressbar
    
        jml_data = 200
        Me.ProgressBar1.max = jml_data

        For i = 1 To jml_data
        Me.ProgressBar1.Value = i
        Next
    
        
'Menutup/menghide Form Logon dan membuka form utama
        
        If Me.acclevel = "Administrator" Then
            
            DoCmd.OpenForm "frmutama"
            DoCmd.Close acForm, "frmutama_user"
            Form_frmUtama.Requery
        Else
        If Me.acclevel = "User" Then
            DoCmd.OpenForm "frmutama"
            'DoCmd.Close acForm, "frmutama"
            Form_frmUtama.Requery
        End If
        End If
        Me.Visible = False
        Form_frm_Login.Visible = False
        
        

    Else
        MsgBox "Password Anda Salah !  Silahkan Coba Lagi ", vbInformation, "Konfirmasi Password !"
        Me.txtPassword.SetFocus
               
       
    End If

End Sub
