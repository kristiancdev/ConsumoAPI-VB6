VERSION 5.00
Begin VB.Form frmInvokeAPI 
   Caption         =   "API DOGS"
   ClientHeight    =   3195
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4785
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4785
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   2295
      Left            =   240
      ScaleHeight     =   2235
      ScaleWidth      =   4155
      TabIndex        =   1
      Top             =   720
      Width           =   4215
   End
   Begin VB.CommandButton btnInvokeApi 
      Caption         =   "Consumir API"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmInvokeAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnInvokeApi_Click()
    Dim xmlhttp As Object
    Dim jsonResponse As Object
    Dim url As String
    Dim message As String
    Dim status As String

    ' URL del API que deseas consumir
    url = "https://dog.ceo/api/breeds/image/random"

    ' Crear un objeto xmlhttp para hacer la solicitud GET
    Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP")

    ' Abrir la conexión con la URL
    xmlhttp.Open "GET", url, False

    ' Enviar la solicitud al servidor
    xmlhttp.send

    ' Verificar si la solicitud fue exitosa (código de estado 200)
    If xmlhttp.status = 200 Then
        ' Parsear la respuesta JSON
        Set jsonResponse = JsonConverter.ParseJson(xmlhttp.responseText)

        ' Obtener el mensaje y el estado (status) del JSON
        message = jsonResponse("message")
        status = jsonResponse("status")

        ' Mostrar el mensaje y el estado en un MessageBox
        MsgBox "Mensaje: " & message & vbCrLf & "Estado: " & status, vbInformation, "Respuesta del API"
    Else
        ' En caso de que la solicitud falle, mostrar un mensaje de error
        MsgBox "Error al hacer la solicitud al API. Código de estado: " & xmlhttp.status, vbExclamation, "Error"
    End If

    ' Liberar el objeto xmlhttp
    Set xmlhttp = Nothing
End Sub

