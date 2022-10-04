VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   5724
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   12264
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BtnAcep_Click()
    Dim nombre, apellido, tlfn, email As String
    Dim idnum As Integer
    
    nombre = InputBox("Ingrese el nombre", "Registro de datos personales", "Nombre")
    Range("B5").Value = nombre
    
    apellido = InputBox("Ingrese el Apellido", "Registro de datos personales", "Apellido")
    Range("C5").Value = apellido
    
    tlfn = InputBox("Ingrese el telefono", "Registro de datos personales", "Telefono")
    Range("D5").Value = tlfn
    
    email = InputBox("Ingrese el correo electronico", "Registro de datos personales", "Correo Electronico")
    Range("E5").Value = email
    
    idnum = 3
    Range("A5").Value = idnum
    
End Sub

Private Sub TxtTelf_Change()

End Sub

Private Sub UserForm_Click()

End Sub
