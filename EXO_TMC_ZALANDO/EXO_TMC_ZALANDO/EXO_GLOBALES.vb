Imports System.IO

Public Class EXO_GLOBALES
    Public Shared Sub CopiarRecurso(ByVal pAssembly As Reflection.Assembly, ByVal pNombreRecurso As String, ByVal pRuta As String)
        Dim s As Stream = pAssembly.GetManifestResourceStream(pAssembly.GetName().Name + "." + pNombreRecurso)
        If s.Length = 0 Then
            Throw New Exception("No se puede encontrar el recurso '" + pNombreRecurso + "'")
        Else
            Dim buffer(CInt(s.Length() - 1)) As Byte
            s.Read(buffer, 0, buffer.Length)

            Dim sw As BinaryWriter = New BinaryWriter(File.Open(pRuta, FileMode.Create))
            sw.Write(buffer)
            sw.Close()
        End If
    End Sub
End Class
