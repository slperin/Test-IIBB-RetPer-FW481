Imports System.IO
Imports System.Net
Imports System.Security.Cryptography
Imports System.Text

Module IIBBRetPer

  'Variables Input
  Public Usuario As String = "" 'CUIT asociado a la contraseña
  Public Password As String = ""
  Public Cuit As String = "" 'CUIT a consultar
  Public FechaDesde As String = ""
  Public FechaHasta As String = ""
  Public Testing As Boolean = False

  Sub Main()

    Consultar_RetPer(Usuario, Password, Cuit, FechaDesde, FechaHasta)
    End

  End Sub

  Private Sub Consultar_RetPer(ByVal Usuario As String, ByVal Password As String, ByVal Cuit As String, FechaDesde As String, ByVal FechaHasta As String)

    Try

      Const Boundary As String = "AaB03x"
      Const Archivo As String = "DFEServicioConsulta_"
      Const ExtensionXml As String = "xml"

      'Configura la URL segun el ambiente
      Dim Url As String
      If Testing Then
        Url = "https://dfe.test.arba.gov.ar/DomicilioElectronico/SeguridadCliente/dfeServicioConsulta.do"
      Else
        Url = "https://dfe.arba.gov.ar/DomicilioElectronico/SeguridadCliente/dfeServicioConsulta.do"
      End If

      'Genera XML
      Dim XMLRequest As String = "<?xml version=""1.0"" encoding=""iso-8859-1""?>"
      XMLRequest &= "<CONSULTA-ALICUOTA>"
      XMLRequest &= "  <fechaDesde>" & FechaDesde & "</fechaDesde>"
      XMLRequest &= "  <fechaHasta>" & FechaHasta & "</fechaHasta>"
      XMLRequest &= "  <cantidadContribuyentes>1</cantidadContribuyentes>"
      XMLRequest &= "  <contribuyentes class=""list"">"
      XMLRequest &= "    <contribuyente>"
      XMLRequest &= "      <cuitContribuyente>" & Cuit & "</cuitContribuyente>"
      XMLRequest &= "    </contribuyente>"
      XMLRequest &= "  </contribuyentes>"
      XMLRequest &= "</CONSULTA-ALICUOTA>"

      'Request
      Dim Contenido As String = XMLRequest & vbCrLf
      Dim Hash As String = MD5FromString(XMLRequest)
      Dim FileName As String = Archivo & Hash & "." & ExtensionXml

      Dim PostData As String = String.Format("--{0}" & vbCrLf & "Content-Disposition: form-data; name=""user""" & vbCrLf & vbCrLf & "{1}" & vbCrLf & "--{0}" & vbCrLf & "Content-Disposition: form-data; name=""password""" & vbCrLf & vbCrLf & "{2}" & vbCrLf & "--{0}" & vbCrLf & "Content-Disposition: form-data; name=""file""; filename={3}" & vbCrLf & "Content-Type: text/xml" & vbCrLf & vbCrLf & "{4}" & "--{0}--", Boundary, Usuario, Password, FileName, Contenido)
      Dim Buffer As Byte() = Encoding.ASCII.GetBytes(PostData)
      Dim Consulta As HttpWebRequest = WebRequest.CreateHttp(Url)
      Consulta.Method = "POST"
      Consulta.ContentType = String.Format("multipart/form-data;boundary={0}", Boundary)
      Consulta.ContentLength = Buffer.Length
      Consulta.Timeout = 200 * 1000 ' 200 segundos

      'Post data
      Dim NewStream As Stream = Consulta.GetRequestStream()
      If (NewStream IsNot Nothing) Then
        NewStream.Write(Buffer, 0, Buffer.Length)
        NewStream.Close()

        'Response
        Dim Respuesta As HttpWebResponse = DirectCast(Consulta.GetResponse(), HttpWebResponse)
        Dim StreamRespuesta As Stream = Respuesta.GetResponseStream()
        Dim StreamReader As New StreamReader(StreamRespuesta)
        Dim RespuestaIIBB As String = StreamReader.ReadToEnd()
        StreamRespuesta.Close()
        StreamReader.Close()
        Respuesta.Close()
        MsgBox(RespuestaIIBB)
      End If
    Catch Ex As Exception
      MsgBox(Ex.Message)
    End Try

  End Sub

  'Private Sub Consultar_RetPer(ByVal Usuario As String, ByVal Password As String, ByVal Cuit As String, FechaDesde As String, ByVal FechaHasta As String)

  '  Try

  '    Const Boundary As String = "AaB03x"
  '    Const Archivo As String = "DFEServicioConsulta_"
  '    Const ExtensionXml As String = "xml"

  '    'Configura la URL segun el ambiente
  '    Dim Url As String
  '    If Testing Then
  '      Url = "https://dfe.test.arba.gov.ar/DomicilioElectronico/SeguridadCliente/dfeServicioConsulta.do"
  '    Else
  '      Url = "https://dfe.arba.gov.ar/DomicilioElectronico/SeguridadCliente/dfeServicioConsulta.do"
  '    End If

  '    'Genera XML
  '    Dim XMLRequest As String = "<?xml version=""1.0"" encoding=""iso-8859-1""?>"
  '    XMLRequest &= "<CONSULTA-ALICUOTA>"
  '    XMLRequest &= "  <fechaDesde>" & FechaDesde & "</fechaDesde>"
  '    XMLRequest &= "  <fechaHasta>" & FechaHasta & "</fechaHasta>"
  '    XMLRequest &= "  <cantidadContribuyentes>1</cantidadContribuyentes>"
  '    XMLRequest &= "  <contribuyentes class=""list"">"
  '    XMLRequest &= "    <contribuyente>"
  '    XMLRequest &= "      <cuitContribuyente>" & Cuit & "</cuitContribuyente>"
  '    XMLRequest &= "    </contribuyente>"
  '    XMLRequest &= "  </contribuyentes>"
  '    XMLRequest &= "</CONSULTA-ALICUOTA>"

  '    'Dim XFileName As String = String.Format("{0}{1}.{2}", Path.GetTempPath(), Archivo, ExtensionXml)
  '    'Dim Writer As StreamWriter = New StreamWriter(XFileName)
  '    'Writer.Write(XMLRequest)
  '    'Writer.Flush()
  '    'Writer.Close()

  '    'Dim Hash As String = CalcularHashMD5(XFileName)
  '    'Dim FileName As String = String.Format("{0}{1}_{2}.{3}", Path.GetTempPath(), Archivo, Hash, ExtensionXml)

  '    'File.Copy(XFileName, FileName, True)
  '    'Dim Contenido As String = ContenidoDeArchivo(FileName)

  '    'Request
  '    Dim Contenido As String = XMLRequest & vbCrLf
  '    Dim Hash As String = MD5FromString(XMLRequest)
  '    'MsgBox("Hash MD5 de la cadena: " & vbCrLf & Hash & vbCrLf & Hash1 & vbCrLf & Contenido.Length.ToString & vbCrLf & XMLRequest.Length.ToString)
  '    'MsgBox("Hash MD5 de la cadena: " & vbCrLf & Hash & vbCrLf & Hash1 & vbCrLf & Contenido & "#" & vbCrLf & XMLRequest & "#")

  '    Dim FileName As String = Archivo & Hash & "." & ExtensionXml
  '    'FileName = Archivo & Hash & "." & ExtensionXml

  '    Dim PostData As String = String.Format("--{0}" & vbCrLf & "Content-Disposition: form-data; name=""user""" & vbCrLf & vbCrLf & "{1}" & vbCrLf & "--{0}" & vbCrLf & "Content-Disposition: form-data; name=""password""" & vbCrLf & vbCrLf & "{2}" & vbCrLf & "--{0}" & vbCrLf & "Content-Disposition: form-data; name=""file""; filename={3}" & vbCrLf & "Content-Type: text/xml" & vbCrLf & vbCrLf & "{4}" & "--{0}--", Boundary, Usuario, Password, FileName, Contenido)
  '    Dim Buffer As Byte() = Encoding.ASCII.GetBytes(PostData)
  '    Dim Consulta As HttpWebRequest = WebRequest.CreateHttp(Url)
  '    Consulta.Method = "POST"
  '    Consulta.ContentType = String.Format("multipart/form-data;boundary={0}", Boundary)
  '    Consulta.ContentLength = Buffer.Length
  '    Consulta.Timeout = 200 * 1000 ' 200 segundos

  '    'Post data
  '    Dim NewStream As Stream = Consulta.GetRequestStream()
  '    If (NewStream IsNot Nothing) Then
  '      NewStream.Write(Buffer, 0, Buffer.Length)
  '      NewStream.Close()

  '      'Response
  '      Dim Respuesta As HttpWebResponse = DirectCast(Consulta.GetResponse(), HttpWebResponse)
  '      Dim StreamRespuesta As Stream = Respuesta.GetResponseStream()
  '      Dim StreamReader As New StreamReader(StreamRespuesta)
  '      Dim RespuestaIIBB As String = StreamReader.ReadToEnd()
  '      StreamRespuesta.Close()
  '      StreamReader.Close()
  '      Respuesta.Close()
  '      MsgBox(RespuestaIIBB)
  '    End If
  '  Catch Ex As Exception
  '    MsgBox(Ex.Message)
  '  End Try

  'End Sub

  Private Function MD5FromString(ByVal Source As String) As String

    Dim j As Integer
    Dim Bytes() As Byte
    Dim sb As New StringBuilder()

    'Check for empty string.
    If String.IsNullOrEmpty(Source) Then
      Throw New ArgumentNullException
    End If

    'Get bytes from string.
    Bytes = Encoding.Default.GetBytes(Source)

    'Get md5 hash
    Bytes = MD5.Create.ComputeHash(Bytes)

    'Loop though the byte array and convert each byte to hex.
    For j = 0 To Bytes.Length - 1
      sb.Append(Bytes(j).ToString("x2"))
    Next

    'Return md5 hash.
    Return sb.ToString().ToUpper

  End Function

  'Private Function CalcularHashMD5(ByVal Archivo As String) As String

  '  Try
  '    Dim FS As FileStream = New FileStream(Archivo, FileMode.Open, FileAccess.Read)
  '    Dim Hash As MD5CryptoServiceProvider = New MD5CryptoServiceProvider()
  '    Dim SB As StringBuilder = New StringBuilder()

  '    FS.Seek(0, SeekOrigin.Begin)
  '    For Each b In Hash.ComputeHash(FS)
  '      SB.Append(b.ToString("X2"))
  '    Next
  '    FS.Close()
  '    Return SB.ToString()
  '  Catch Ex As Exception
  '    MsgBox(Ex.Message)
  '    Return ""
  '  End Try

  'End Function

  'Private Function ContenidoDeArchivo(ByVal Archivo As String) As String

  '  Try
  '    Dim Stream As FileStream = New FileStream(Archivo, FileMode.Open, FileAccess.Read)
  '    Dim Cadena As String = String.Empty

  '    If Stream IsNot Nothing Then
  '      Stream.Position = 0
  '      Dim Reader = New StreamReader(Stream)
  '      While Reader.Peek() > -1
  '        Cadena = Cadena & Reader.ReadLine() & vbCrLf
  '      End While
  '      Reader.Close()
  '    End If
  '    Stream.Close()

  '    Return Cadena
  '  Catch Ex As Exception
  '    MsgBox(Ex.Message)
  '    Return ""
  '  End Try

  'End Function

End Module