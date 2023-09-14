# Consumo de API y Parseo de JSON con Visual Basic 6
Este proyecto en Visual Basic 6 muestra cómo consumir una API web y utilizar una biblioteca externa de GitHub para analizar JSON. En este ejemplo, aprenderás cómo realizar solicitudes HTTP a una API y procesar la respuesta JSON utilizando la biblioteca externa JSONParser.

Claro, puedo ayudarte a generar un README para tu proyecto en Visual Basic 6 que muestre cómo consumir una API y usar una biblioteca externa de GitHub para analizar JSON. Aquí tienes un ejemplo de README:

## Requisitos previos

Asegúrate de tener los siguientes requisitos previos instalados antes de continuar:

1. Visual Basic 6 (VB6) instalado en tu sistema.

## Configuración del proyecto

1. Clona este repositorio en tu sistema local:

   ```
   git clone https://github.com/kristiancdev/ConsumoAPI-VB6
   ```

2. Abre el proyecto en Visual Basic 6.

## Consumo de la API

En este proyecto de ejemplo, se utiliza [Dog API](https://dog.ceo/api/breeds/image/random) para obtener datos. Puedes reemplazar la URL de la API y los parámetros según tus necesidades.

```vb
Private Sub btnConsumirAPI_Click()
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
```

## Uso de la biblioteca externa JSONParser

Para analizar JSON, este proyecto utiliza la biblioteca externa JSONParser, que puedes encontrar en [VBA-JSON](https://github.com/VBA-tools/VBA-JSON).

1. Clona el repositorio JSONParser en tu sistema local:

   ```
   git clone https://github.com/tu-usuario/JSONParser.git
   ```

2. Agrega la clase JSONParser a tu proyecto VB6. Para hacerlo, ve a "Proyecto" -> "Clic derecho agregar módulo" y selecciona "JSONParser" en el explorador de archivos.

## Contribuir y mejorar el proyecto

Si deseas contribuir a este proyecto o mejorarlo, siéntete libre de enviar pull requests o informar problemas en [ConsumoAPI-VB6](https://github.com/kristiancdev/ConsumoAPI-VB6)

.

¡Disfruta trabajando en tu proyecto de Visual Basic 6 y consumiendo APIs con facilidad!
