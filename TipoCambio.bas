Attribute VB_Name = "CustomModule"

' Cambios generados en la rama NuevosCambios

' FUNCION QUE PERMITE OBTENER EL TIPO DE CAMBIO DESDE EL DIARIO OFICIAL USD - MXN

' 22-05-2023 ACC

Public Function GetTipoCambioFromDof() As String
   
   ' Inicia solicitud GET aqui -------------------------------------------
   
   ' Variables a utilizar
   Dim objHttp As Object       ' Para crear el objeto de la solicitud HTTP
   Dim response As String      ' Para capturar la respuesta
   Dim url As String           ' Url del recurso a consultar
   Dim json As Object          ' Para parsear la respuesta a json pues aqui la obtenmos como una cadena de texto y asi no podemos hacer mucho
   Dim list As Object          ' Para extraer la lista que viene en la respuesta
   Dim firstElement As Object  ' Para extraer el primer elemento de la lista
   Dim value As String         ' Para obtener el valor deseado del primer elemento de la lista (tipo de cambio)

   ' Instanciar objeto Microsoft.XMLHTTP para hacer la peticion
   Set objHttp = CreateObject("Microsoft.XMLHTTP")

   ' Establecer la url a consultar (si no se especifica la fecha tomara la actual para consultar)
   url = "https://sidofqa.segob.gob.mx/dof/sidof/indicadores/"

   ' Especificar la solicitud
   ' Explicaci�n: Open se utiliza para especificar el m�todo de solicitud
   ' en este caso "GET", enseguida la url y si la solicitud es sincrona, true o false
   objHttp.Open "GET", url, False

   ' Enviar la solicitud
   objHttp.send

   ' Verificar si la solicitud se realiz� correctamente
   If objHttp.Status = 200 Then

       ' Obtener la respuesta
       response = objHttp.responseText
       
       ' IMPORTANTE!!!!

       ' ParseJson() es una funci�n que no est� disponible de forma nativa en Visual Basic 6,
       ' pertenece a un modulo del proyecto externo llamado VBA-JSON
       ' repo. del proyecto -> https://github.com/VBA-tools/VBA-JSON
       
       ' Para agregar este modulo solo hay que dar click derecho en el nombre del proyecto,
       ' Agregar/Agregar archivo y seleccionar el archivo JsonConverter.bas

       ' Es necesario agregar la referencia "Microsoft Scripting Runtime" para que funcione
       ' esto en -> Proyecto/Referencias

       ' Parsear la respuesta a json
       Set json = ParseJson(response)
       
       ' Extraer la lista
       Set list = json("ListaIndicadores")
       
       ' Extraer el primer elemento de la lista
       Set firstElement = list(1)
       
       ' Obtener el valor del atributo "valor" (tipo de cambio) del primer elemento de la lista
       value = firstElement("valor")
       
       ' Retornar el resultado
       GetTipoCambioFromDof = value
   Else
    
       ' Capturar mensaje de error si la solicitud no fue exitosa
       GetTipoCambioFromDof = "Error en la solicitud: " & objHttp.Status
   End If

   ' Desechar el objeto de solicitud para liberar recursos
   Set objHttp = Nothing

   ' Termina solicitud GET aqui -------------------------------------------

End Function

' FUNCION QUE PERMITE OBTENER EL TIPO DE CAMBIO CON EXCHANGERATE API

' 23-05-2023 ACC

Public Function GetTipoCambioFromExchangeAPI(ByVal bCurrency As String, ByVal tCurrency As String) As String

   ' Inicia solicitud GET aqui -------------------------------------------
   
   ' Variables a utilizar
   Dim objHttp As Object         ' Para crear el objeto de la solicitud HTTP
   Dim response As String        ' Para capturar la respuesta
   Dim url As String             ' Url del recurso a consultar
   Dim json As Object            ' Para parsear la respuesta a json pues aqui la obtenmos como una cadena de texto y asi no podemos hacer mucho
   Dim value As String           ' Para obtener el valor deseado
   Dim baseCurrency As String    ' Para obtener la moneda base
   Dim targetCurrency As String  ' Para obtener la moneda a consultar
   Dim key As String             ' Key para poder hacer peticiones a la API

   ' Instanciar objeto Microsoft.XMLHTTP para hacer la peticion
   Set objHttp = CreateObject("Microsoft.XMLHTTP")

   ' Establecer la key proporcionada por la API para poder hacer peticiones
   key = "2eb11317c2a656877b034aea"
   
   ' Obtener baseCurrency y targetCurrency
   baseCurrency = bCurrency
   targetCurrency = tCurrency
   
   ' Establecer la url a consultar
   url = "https://v6.exchangerate-api.com/v6/" + key + "/pair/" + baseCurrency + "/" + targetCurrency

   ' Especificar la solicitud
   ' Explicaci�n: Open se utiliza para especificar el m�todo de solicitud
   ' en este caso "GET", enseguida la url y si la solicitud es sincrona, true o false
   objHttp.Open "GET", url, False

   ' Enviar la solicitud
   objHttp.send

   ' Verificar si la solicitud se realiz� correctamente
   If objHttp.Status = 200 Then

       ' Obtener la respuesta
       response = objHttp.responseText
       
       ' IMPORTANTE!!!!

       ' ParseJson() es una funci�n que no est� disponible de forma nativa en Visual Basic 6,
       ' pertenece a un modulo del proyecto llamado VBA-JSON
       ' repo. del proyecto -> https://github.com/VBA-tools/VBA-JSON
       
       ' Para agregar este modulo solo hay que dar click derecho en el nombre del proyecto,
       ' Agregar/Agregar archivo y seleccionar el archivo JsonConverter.bas

       ' Es necesario agregar la referencia "Microsoft Scripting Runtime" para que funcione
       ' esto en -> Proyecto/Referencias

       ' Parsear la respuesta a json
       Set json = ParseJson(response)
       
       ' Extraer el valor que queremos de la respuesta
       value = json("conversion_rate")
       
       ' Retornar el resultado
       GetTipoCambioFromExchangeAPI = value
   Else
    
       ' Capturar mensaje de error si la solicitud no fue exitosa
       GetTipoCambioFromExchangeAPI = "Error en la solicitud: " & objHttp.Status
   End If

   ' Desechar el objeto de solicitud para liberar recursos
   Set objHttp = Nothing

   ' Termina solicitud GET aqui -------------------------------------------

End Function
