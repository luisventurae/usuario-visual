Attribute VB_Name = "Module1"
Option Explicit

Public fso As Object

'formulario al iniciar
Public Sub Main()

On Error GoTo err_sub

Set fso = CreateObject("scripting.filesystemobject")

'comprobacion de datos existentes
If fso.fileexists(App.Path & "/Datos.txt") Then
frmInicio.Show
End If

Set fso = Nothing
Exit Sub
err_sub:
  MsgBox Err.Description, vbCritical, "error al usar fso"
End Sub

Public Sub Datos()

On Error GoTo error

Open App.Path & "/Datos.txt" For Append As #1
Print #1, frmcrear.Text1.Text
Print #1, frmcrear.Text2.Text
Close #1
MsgBox "Se crearon los datos de usuario" & vbCrLf & _
"Usted ya puede entrar a su nueva cuenta", vbInformation
frmcrear.Text1.Text = ""
frmcrear.Text2.Text = ""
frmcrear.Text3.Text = ""
frmcrear.Hide
frmInicio.Show
Exit Sub
error:
MsgBox "Nose puede crear!", vbCritical, "Error"
End Sub

Public Sub usuario()
'Para leer los datos de entrada
On Error GoTo error1
Dim nombre As String, clave As String
Open App.Path & "/Datos.txt" For Input As #1
While Not EOF(1)
Line Input #1, nombre
Line Input #1, clave

'Verificacion si los datos en las cajas de texto son correctos
If frm0.Text1.Text = nombre And frm0.Text2.Text = clave Then
frmcargar.Shape1.FillColor = RGB(0, 255, 255)
frm0.Hide
frmcargar.Show
Else: frm0.Text1.SetFocus
End If
Wend
Close #1
Exit Sub
error1:
MsgBox "No se pudo leer", vbCritical, "Error"
End Sub
'sub para eliminar el archivo creado
Public Sub eliminar()
On Error GoTo err_sub
Set fso = CreateObject("scripting.filesystemobject")

'comprobacion de existencia del archivo
If fso.fileexists(App.Path & "/Datos.txt") Then
Kill App.Path & "/Datos.txt"
MsgBox "Todos los datos de usuario han sido eliminados" & vbCrLf & _
"El programa se iniciara normalmente la proxima vez", vbInformation
Else: MsgBox "No existe el archivo"
End If

Set fso = Nothing
Exit Sub
err_sub:
  MsgBox Err.Description, vbCritical, "Error al usar fso"
End Sub


