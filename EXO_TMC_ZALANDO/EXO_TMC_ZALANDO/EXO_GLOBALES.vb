Imports System.IO
Imports SAPbouiCOM

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
    Public Shared Sub BorrarFicheros(ByVal sPath As String)
        'Borramos los ficheros más antiguos a X días
        Dim Fecha As DateTime = DateTime.Now
        Dim sDias = 30
        For Each archivo As String In My.Computer.FileSystem.GetFiles(sPath, FileIO.SearchOption.SearchTopLevelOnly)
            Dim Fecha_Archivo As DateTime = My.Computer.FileSystem.GetFileInfo(archivo).LastWriteTime
            Dim diferencia = (CType(Fecha, DateTime) - CType(Fecha_Archivo, DateTime)).TotalDays

            If diferencia >= CDbl(sDias) Then ' Nº de días
                File.Delete(archivo)
            End If
        Next
    End Sub
    Public Shared Sub Copia_Seguridad(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByVal sArchivoOrigen As String, ByVal sArchivo As String)
        'Comprobamos el directorio de copia que exista
        Dim sPath As String = ""
        sPath = IO.Path.GetDirectoryName(sArchivo)
        If IO.Directory.Exists(sPath) = False Then
            IO.Directory.CreateDirectory(sPath)
        End If
        If IO.File.Exists(sArchivo) = True Then
            IO.File.Delete(sArchivo)
        End If
        'Subimos el archivo
        oObjGlobal.SBOApp.StatusBar.SetText("(EXO) - Comienza la Copia de seguridad del fichero - " & sArchivoOrigen & " -.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        If oObjGlobal.SBOApp.ClientType = BoClientType.ct_Browser Then
            Dim fs As FileStream = New FileStream(sArchivoOrigen, FileMode.Open, FileAccess.Read)
            Dim b(CInt(fs.Length() - 1)) As Byte
            fs.Read(b, 0, b.Length)
            fs.Close()
            Dim fs2 As New System.IO.FileStream(sArchivo, IO.FileMode.Create, IO.FileAccess.Write)
            fs2.Write(b, 0, b.Length)
            fs2.Close()
        Else
            My.Computer.FileSystem.CopyFile(sArchivoOrigen, sArchivo)
        End If
        oObjGlobal.SBOApp.StatusBar.SetText("(EXO) - Copia de Seguridad realizada Correctamente", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
    End Sub
#Region "Funciones formateos datos"
    Public Shared Function TextToDbl(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByVal sValor As String) As Double
        Dim cValor As Double = 0
        Dim sValorAux As String = "0"

        TextToDbl = 0

        Try
            sValorAux = sValor

            If oObjGlobal.SBOApp.ClientType = SAPbouiCOM.BoClientType.ct_Desktop Then
                If sValorAux <> "" Then
                    If Left(sValorAux, 1) = "." Then sValorAux = "0" & sValorAux

                    If oObjGlobal.refDi.OADM.separadorMillarB1 = "." AndAlso oObjGlobal.refDi.OADM.separadorDecimalB1 = "," Then 'Decimales ES
                        If sValorAux.IndexOf(".") > 0 AndAlso sValorAux.IndexOf(",") > 0 Then
                            cValor = CDbl(sValorAux.Replace(".", ""))
                        ElseIf sValorAux.IndexOf(".") > 0 Then
                            cValor = CDbl(sValorAux.Replace(".", ","))
                        Else
                            cValor = CDbl(sValorAux)
                        End If
                    Else 'Decimales USA
                        If sValorAux.IndexOf(".") > 0 AndAlso sValorAux.IndexOf(",") > 0 Then
                            cValor = CDbl(sValorAux.Replace(",", "").Replace(".", ","))
                        ElseIf sValorAux.IndexOf(".") > 0 Then
                            cValor = CDbl(sValorAux.Replace(".", ","))
                        Else
                            cValor = CDbl(sValorAux)
                        End If
                    End If
                End If
            Else
                If sValorAux <> "" Then
                    If Left(sValorAux, 1) = "," Then sValorAux = "0" & sValorAux

                    If oObjGlobal.refDi.OADM.separadorMillarB1 = "." AndAlso oObjGlobal.refDi.OADM.separadorDecimalB1 = "," Then 'Decimales ES
                        If sValorAux.IndexOf(",") > 0 AndAlso sValorAux.IndexOf(".") > 0 Then
                            cValor = CDbl(sValorAux.Replace(",", ""))
                        ElseIf sValorAux.IndexOf(",") > 0 Then
                            cValor = CDbl(sValorAux.Replace(",", "."))
                        Else
                            cValor = CDbl(sValorAux)
                        End If
                    Else 'Decimales USA
                        If sValorAux.IndexOf(",") > 0 AndAlso sValorAux.IndexOf(".") > 0 Then
                            cValor = CDbl(sValorAux.Replace(".", "").Replace(",", "."))
                        ElseIf sValorAux.IndexOf(",") > 0 Then
                            cValor = CDbl(sValorAux.Replace(",", "."))
                        Else
                            cValor = CDbl(sValorAux)
                        End If
                    End If
                End If
            End If

            TextToDbl = cValor

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Public Shared Function DblNumberToText(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByVal sValor As String) As String
        Dim sNumberDouble As String = "0"

        DblNumberToText = "0"

        Try
            If sValor <> "" Then
                If oObjGlobal.refDi.OADM.separadorMillarB1 = "." AndAlso oObjGlobal.refDi.OADM.separadorDecimalB1 = "," Then 'Decimales ES
                    sNumberDouble = sValor
                Else 'Decimales USA
                    sNumberDouble = sValor.Replace(",", ".")
                End If
            End If

            DblNumberToText = sNumberDouble


        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Public Shared Function FormateaString(ByVal dato As Object, ByVal tam As Integer) As String
        Dim retorno As String = String.Empty

        If dato IsNot Nothing Then
            retorno = dato.ToString
        End If

        If retorno.Length > tam Then
            retorno = retorno.Substring(0, tam)
        End If

        Return retorno.PadRight(tam, CChar(" "))
    End Function
    Public Shared Function FormateaNumero(ByVal dato As String, ByVal posiciones As Integer, ByVal decimales As Integer, ByVal obligatorio As Boolean) As String
        Dim retorno As String = String.Empty
        Dim val As Decimal
        Dim totalNum As Integer = posiciones
        Dim format As String = ""
        Dim bEsNegativo As Boolean = False
        If Left(dato, 1) = "-" Then
            dato = dato.Replace("-", "")
            bEsNegativo = True
            posiciones = posiciones - 1
            totalNum = posiciones
        End If
        Decimal.TryParse(dato.Replace(".", ","), val)
        If val = 0 AndAlso Not obligatorio Then
            retorno = New String(CChar(" "), posiciones)
        Else
            If decimales <= 0 Then
            Else
                format = "0"
                format = "0." & New String(CChar("0"), decimales)
            End If
            retorno = val.ToString(format).Replace(",", ".")
            If retorno.Length > totalNum Then
                retorno = retorno.Substring(retorno.Length - totalNum)
            End If
            retorno = retorno.PadLeft(totalNum, CChar("0"))
        End If
        If bEsNegativo = True Then
            retorno = "N" & retorno
        End If
        Return retorno
    End Function
    Public Shared Function FormateaNumeroSinPunto(ByVal dato As String, ByVal posiciones As Integer, ByVal decimales As Integer, ByVal obligatorio As Boolean) As String
        Dim retorno As String = String.Empty
        Dim val As Decimal
        Dim totalNum As Integer = posiciones
        Dim format As String = ""
        Dim bEsNegativo As Boolean = False
        If Left(dato, 1) = "-" Then
            dato = dato.Replace("-", "")
            bEsNegativo = True
            posiciones = posiciones - 1
            totalNum = posiciones
        End If
        Decimal.TryParse(dato.Replace(".", ","), val)
        If val = 0 AndAlso Not obligatorio Then
            retorno = New String(CChar(" "), posiciones)
        Else
            If decimales <= 0 Then
            Else
                format = "0"
                format = "0." & New String(CChar("0"), decimales)
            End If
            retorno = val.ToString(format).Replace(",", ".")
            retorno = retorno.Replace(".", "")

            If retorno.Length > totalNum Then
                retorno = retorno.Substring(retorno.Length - totalNum)
            End If
            retorno = retorno.PadLeft(totalNum, CChar("0"))
        End If
        If bEsNegativo = True Then
            retorno = "N" & retorno
        End If
        Return retorno
    End Function
    Public Shared Function FormateaNumeroconSigno(ByVal dato As String, ByVal posiciones As Integer, ByVal decimales As Integer, ByVal obligatorio As Boolean) As String
        Dim retorno As String = String.Empty
        Dim val As Decimal
        Dim totalNum As Integer = posiciones
        Dim format As String = ""
        Dim bEsNegativo As Boolean = False
        If Left(dato, 1) = "-" Then
            dato = dato.Replace("-", "")
            bEsNegativo = True
            posiciones = posiciones - 1
            totalNum = posiciones
        End If
        Decimal.TryParse(dato.Replace(".", ","), val)
        If val = 0 AndAlso Not obligatorio Then
            retorno = New String(CChar(" "), posiciones)
        Else
            If decimales <= 0 Then
            Else
                format = "0"
                format = "0." & New String(CChar("0"), decimales)
            End If
            retorno = val.ToString(format).Replace(",", ".")
            retorno = retorno.Replace(".", "")

            If retorno.Length > totalNum Then
                retorno = retorno.Substring(retorno.Length - totalNum)
            End If
            retorno = retorno.PadLeft(totalNum, CChar("0"))
        End If
        If bEsNegativo = True Then
            retorno = "N" & retorno
        Else
            retorno = " " & retorno
        End If
        Return retorno
    End Function
#End Region
#Region "SQL"
    Public Shared Function GetValueDB(oCompany As SAPbobsCOM.Company, ByRef sTable As String, ByRef sField As String, ByRef sCondition As String) As String
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

        Try
            If sCondition = "" Then
                sSQL = "Select " & sField & " FROM " & sTable
            Else
                sSQL = "Select " & sField & " FROM " & sTable & " WHERE " & sCondition
            End If
            oRs.DoQuery(sSQL)
            If oRs.RecordCount > 0 Then
                sField = sField.Replace("""", "")
                GetValueDB = CType(oRs.Fields.Item(sField).Value, String)
            Else
                GetValueDB = ""
            End If

        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function
#End Region
#Region "Tratar ficheros"
    Public Shared Sub TratarFichero_TXT(ByVal sArchivo As String, ByVal sDelimitador As String, ByRef oForm As SAPbouiCOM.Form, ByRef oCompany As SAPbobsCOM.Company, ByRef objglobal As EXO_UIAPI.EXO_UIAPI)
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim oRCampos As SAPbobsCOM.Recordset = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sCampo As String = ""

        Dim iDoc As Integer = 0 'Contador de Cabecera de documentos
        Dim sTFac As String = "" : Dim sTFacColumna As String = "" : Dim sTipoLineas As String = "" : Dim sTDoc As String = ""
        Dim sCliente As String = "" : Dim sCliNombre As String = "" : Dim sCodCliente As String = "" : Dim sClienteColumna As String = "" : Dim sCodClienteColumna As String = ""
        Dim sSerie As String = "" : Dim sDocNum As String = "" : Dim sManual As String = "" : Dim sSerieColumna As String = "" : Dim sDocNumColumna As String = ""
        Dim sDIR As String = "" : Dim sPob As String = "" : Dim sProv As String = "" : Dim sCPos As String = ""
        Dim sNumAtCard As String = "" : Dim sNumAtCardColumna As String = ""
        Dim sMoneda As String = "" : Dim sMonedaColumna As String = ""
        Dim sEmpleado As String = ""
        Dim sFContable As String = "" : Dim sFDocumento As String = "" : Dim sFVto As String = "" : Dim sFDocumentoColumna As String = ""
        Dim sTipoDto As String = "" : Dim sDto As String = ""
        Dim sPeyMethod As String = "" : Dim sCondPago As String = ""
        Dim sDirFac As String = "" : Dim sDirEnv As String = ""
        Dim sComent As String = "" : Dim sComentCab As String = "" : Dim sComentPie As String = ""
        Dim sCondicion As String = ""

        Dim sExiste As String = ""
        Dim bCrearCli As Boolean = False
        Dim iLinea As Integer = 0 : Dim sCodCampos As String = ""

        Dim sMensaje As String = ""
        Dim sCamposC(1, 3) As String : Dim sCamposL(1, 3) As String

        ' Apuntador libre a archivo
        Dim Apunt As Integer = FreeFile()
        ' Variable donde guardamos cada línea de texto
        Dim Texto As String = ""
        Dim sValorCampo As String = ""

        Dim sDocumento As String = "" : Dim sRef As String = ""
        'Cada Fichero es una cabecera de documento, por lo que utilizamos una variable para generar solo una cabecera
        Dim bGenerarcabecera = True : Dim sDatosCabecera As String = "" : Dim sSaltarCabecera As String = ""
        Try
            'Tengo que buscar en la tabla el último numero de documento
            iDoc = objglobal.refDi.SQL.sqlNumericaB1("SELECT isnull(MAX(cast(CODE as int)),0) FROM ""@EXO_TMPDOC"" ")
            sSaltarCabecera = objglobal.refDi.OGEN.valorVariable("Zalando_Fich_con_Cabecera")
            ' miramos si existe el fichero y cargamos
            If File.Exists(sArchivo) Then
                Using MyReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(sArchivo, System.Text.Encoding.UTF7)
                    MyReader.TextFieldType = FileIO.FieldType.Delimited
                    MyReader.SetDelimiters(",")

                    Dim currentRow As String()
                    While Not MyReader.EndOfData
                        Try
                            currentRow = MyReader.ReadFields()
                            'Comprobamos si tenemos que saltar la cabecera
                            If sSaltarCabecera = "N" Then
                                Dim currentField As String
                                Dim scampos(1) As String
                                Dim iCampo As Integer = 0
                                For Each currentField In currentRow
                                    iCampo += 1
                                    ReDim Preserve scampos(iCampo)
                                    scampos(iCampo) = currentField
                                    'SboApp.MessageBox(scampos(iCampo))
                                Next
#Region "Lectura cabecera"
                                If bGenerarcabecera = True Then
                                    bGenerarcabecera = False
                                    objglobal.SBOApp.StatusBar.SetText("(EXO) - Leyendo Valores de Cabecera...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    sDatosCabecera = Path.GetFileNameWithoutExtension(sArchivo)
                                    Dim sDatosCab() As String = sDatosCabecera.Split("_")
                                    If sTFac = "" Then
                                        Select Case sDatosCab(3)
                                            Case "SAL" : sTFac = "13" ' En el caso de no estar indicado, se ha tomado como Factura de venta
                                            Case "RET", "CAN" : sTFac = "14" ' En el caso de no estar indicado, se ha tomado como Factura de compras
                                            Case Else : sTFac = "13" ' En el caso de no estar indicado, se ha tomado como Factura de venta
                                        End Select
                                    End If
                                    If sTipoLineas = "" Then : sTipoLineas = "I" : End If ' En el caso de no estar indicado, se ha tomado como que son líneas de artículo
                                    sTDoc = objglobal.refDi.OGEN.valorVariable("Zalando_Documento")
#Region "Cliente"
                                    sCliente = objglobal.refDi.OGEN.valorVariable("Zalando_IC")
                                    sExiste = ""
                                    sExiste = EXO_GLOBALES.GetValueDB(objglobal.compañia, """OCRD""", """CardCode""", """CardCode""='" & sCliente & "'")
                                    If sExiste = "" Then
                                        objglobal.SBOApp.StatusBar.SetText("(EXO) - El Interlocutor  - " & sCliente & " - no se encuentra al buscarlo por el código de SAP. No se puede continuar. Revise los datos del parámetro - Zalando_IC -.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                        Exit Sub
                                    End If
                                    'Comprobamos las direcciones de entrega y de facturación para comprobar que si son las de defecto del desarrollo, debemos buscar las de por defecto del cliente
                                    sSQL = "SELECT ""BillToDef"" FROM ""OCRD"" WHERE ""CardCode""='" & sCliente & "' "
                                    Dim sDIrFacDef As String = objglobal.refDi.SQL.sqlStringB1(sSQL)
                                    If sDIrFacDef <> "" Then
                                        sDirFac = sDIrFacDef
                                    End If

                                    sSQL = "SELECT ""ShipToDef"" FROM ""OCRD"" WHERE ""CardCode""='" & sCliente & "' "
                                    Dim sDirEnvDef As String = objglobal.refDi.SQL.sqlStringB1(sSQL)
                                    If sDirEnvDef <> "" Then
                                        sDirEnv = sDirEnvDef
                                    End If
#End Region
#Region "NumAtCard"
                                    sNumAtCard = sDatosCab(2)
#End Region
#Region "Contador/Serie"
                                    Dim sAnno As String = scampos(1)
                                    Dim sRemark As String = objglobal.refDi.OGEN.valorVariable("Zalando_Remark")
                                    sManual = "N"
                                    sSerie = EXO_GLOBALES.GetValueDB(objglobal.compañia, """NNM1""", """SeriesName""", " ""ObjectCode""='" & sTFac & "' and  Indicator='" & sAnno & "' and Remark='" & sRemark & "'")
                                    sDocNum = ""
                                    If sSerie = "" Then
                                        objglobal.SBOApp.StatusBar.SetText("(EXO) - No se encuentra la serie para el tipo de documento a realizar con el remark - " & sRemark & "-. No se puede continuar.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                        Exit Sub
                                    End If
#End Region
#Region "Moneda"
                                    sMoneda = scampos(4)
                                    If sMoneda = "" Then
                                        sMoneda = "EUR"
                                    End If
#End Region
#Region "Fechas"
                                    Dim sFechaLectura = scampos(3)
                                    Dim sFechaDoc() As String = sFechaLectura.Split(".")
                                    sFContable = sFechaDoc(2) & "-" & sFechaDoc(1) & "-" & sFechaDoc(0)
                                    sFDocumento = sFechaDoc(2) & "-" & sFechaDoc(1) & "-" & sFechaDoc(0)
                                    sFVto = ""
#End Region
#Region "Comment"
                                    sComent = "Importado a través Zalando. Nº:  " & sNumAtCard
#End Region
                                    objglobal.SBOApp.StatusBar.SetText("(EXO) - Valores de Cabecera leída.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
#Region "Comprobar datos cabecera"
                                    If sTipoDto = "" Then : sTipoDto = "%" : End If ' Se toma si no tiene valor que el dto va en Porcentaje
                                    If sDto = "" Then : sDto = "0.00" : End If ' Se toma por defecto dto valor a 0.00                                   
                                    objglobal.SBOApp.StatusBar.SetText("(EXO) - Datos de cabecera comprobados.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
#End Region
                                    'Grabamos la cabecera
                                    sRef = sNumAtCard
                                    iDoc += 1
                                    iLinea = 0
                                    sSQL = "insert into ""@EXO_TMPDOC"" values('" & iDoc.ToString & "','" & iDoc.ToString & "'," & iDoc.ToString & ",'N','',0," & objglobal.compañia.UserSignature
                                    sSQL &= ",'','',0,'',0,'','" & objglobal.SBOApp.Company.UserName & "',"
                                    sSQL &= "'" & sTDoc & "','" & sDocNum & "','" & sTFac & "','" & sManual & "','" & sSerie & "','" & sNumAtCard & "','" & sMoneda & "','','" & sEmpleado & "',"
                                    sSQL &= "'" & sCliente & "','" & sCodCliente & "','" & sFContable & "','" & sFDocumento & "','" & sFVto & "','" & sTipoDto & "',"
                                    sSQL &= EXO_GLOBALES.DblNumberToText(objglobal, sDto.ToString) & ",'" & sPeyMethod & "','" & sDirFac & "','" & sDirEnv & "','" & sComent.Replace("'", "") & "','"
                                    sSQL &= sComentCab.Replace("'", "") & "','" & sComentPie.Replace("'", "") & "','" & sCondPago & "') "
                                    oRs.DoQuery(sSQL)
                                End If
#End Region
#Region "Lectura de Líneas"

                                Dim sCuenta As String = "" : Dim sArt As String = "" : Dim sArtDes As String = ""
                                Dim sCantidad As String = "0.00" : Dim sprecio As String = "0.00" : Dim sDtoLin As String = "0.00" : Dim sTotalServicios As String = "0.00" : Dim sPrecioBruto As String = "0.00"
                                Dim sTextoAmpliado As String = "" : Dim sLinImpuestoCod As String = "" : Dim sLinRetCodigo As String = ""

                                objglobal.SBOApp.StatusBar.SetText("(EXO) - Leyendo Valores de Líneas...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                'Ahora la Línea                           
                                sArt = scampos(8)

                                sCantidad = scampos(11)
                                sprecio = scampos(13).Replace("-", "")
                                sTextoAmpliado = scampos(5)

                                objglobal.SBOApp.StatusBar.SetText("(EXO) - Valores de líneas leídos.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
#End Region

#Region "Comprobar datos línea"
                                'Comprobamos que exista la cuenta                                  
                                If sCuenta <> "" Then
                                    sExiste = ""
                                    sExiste = EXO_GLOBALES.GetValueDB(objglobal.compañia, """OACT""", """AcctCode""", """AcctCode""='" & sCuenta & "'")
                                    If sExiste = "" Then
                                        objglobal.SBOApp.StatusBar.SetText("(EXO) - La Cuenta contable SAP  - " & sCuenta & " - no existe.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        objglobal.SBOApp.MessageBox("La Cuenta contable SAP - " & sCuenta & " - no existe.")
                                        Exit Sub
                                    End If
                                End If
                                'Comprobamos que exista el artículo
                                If sTipoLineas = "I" Then
                                    sExiste = ""
                                    sExiste = EXO_GLOBALES.GetValueDB(objglobal.compañia, """OITM""", """ItemCode""", """Codebars""= '" & sArt & "'")
                                    If sExiste = "" Then
                                        objglobal.SBOApp.StatusBar.SetText("(EXO) - El Artículo - " & sArt & " - " & sArtDes & " no existe. Borrelo de la sección concepto para poderlo crear automáticamente.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        objglobal.SBOApp.MessageBox("El Artículo - " & sArt & " - " & sArtDes & " no existe. Borrelo de la sección concepto para poderlo crear automáticamente.")
                                        Exit Sub
                                    Else
                                        sArt = sExiste
                                    End If
                                ElseIf sTipoLineas = "S" Then
                                    If sCuenta = "" Then
                                        ' No puede estar la cuenta vacía si es de tipo servicio
                                        sExiste = ""
                                        sExiste = EXO_GLOBALES.GetValueDB(objglobal.compañia, """@EXO_CFCNF""", """U_EXO_CSRV""", """Code""='" & CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "'")
                                        If sExiste = "" Then
                                            sMensaje = " La cuenta en la línea del servicio no puede estar vacía. Por favor, Revise los datos de la parametrización."
                                            objglobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            objglobal.SBOApp.MessageBox(sMensaje)
                                            Exit Sub
                                        Else
                                            sCuenta = sExiste
                                        End If
                                    End If
                                End If
                                'Comprobamos que exista el impuesto si está relleno
                                'If sLinImpuestoCod = "" Then
                                '    Select Case sTFac
                                '        Case "13", "14" 'Ventas
                                '            sLinImpuestoCod = EXO_GLOBALES.GetValueDB(objglobal.compañia, """@EXO_CFCNF""", """U_EXO_IVAV""", """Code""='" & CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "'")
                                '        Case "18", "19", "22" 'Compras
                                '            sLinImpuestoCod = EXO_GLOBALES.GetValueDB(objglobal.compañia, """@EXO_CFCNF""", """U_EXO_IVAC""", """Code""='" & CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "'")
                                '    End Select
                                'Else
                                '    sLinImpuestoCod = sLinImpuestoCod.Replace(",", ".")
                                '    Select Case sTFac
                                '        Case "13", "14" 'Ventas
                                '            sLinImpuestoCod = EXO_GLOBALES.GetValueDB(objglobal.compañia, """OVTG""", """Code""", """Rate""='" & sLinImpuestoCod & "' and  LENGTH(""Code"")=2 and left(""Code"",1)='R' and ""Category""='O' ")
                                '        Case "18", "19", "22" 'Compras
                                '            sLinImpuestoCod = EXO_GLOBALES.GetValueDB(objglobal.compañia, """OVTG""", """Code""", """Rate""='" & sLinImpuestoCod & "' and  LENGTH(""Code"")=2 and left(""Code"",1)='S' and ""Category""='I' ")
                                '    End Select
                                'End If
                                'If sLinImpuestoCod <> "" Then
                                '    sExiste = ""
                                '    sExiste = EXO_GLOBALES.GetValueDB(objglobal.compañia, """OVTG""", """Code""", """Code""='" & sLinImpuestoCod & "'")
                                '    If sExiste = "" Then
                                '        objglobal.SBOApp.StatusBar.SetText("(EXO) - El Grupo Impositivo  - " & sLinImpuestoCod & " - no existe.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                '        objglobal.SBOApp.MessageBox("El Grupo Impositivo  - " & sLinImpuestoCod & " - no existe.")
                                '        Exit Sub
                                '    End If
                                'End If
                                'Comprobamos que exista la retención si está relleno
                                'If sLinRetCodigo <> "" Then
                                '    sExiste = EXO_GLOBALES.GetValueDB(objglobal.compañia, """CRD4""", """WTCode""", """CardCode""='" & sCliente & "' and ""WTCode""='" & sLinRetCodigo & "'")
                                '    If sExiste = "" Then
                                '        objglobal.SBOApp.StatusBar.SetText("(EXO) - El indicador de Retención  - " & sLinRetCodigo & " - no no está marcado para el interlocutor " & sCliente & ". Por favor, revise el interlocutor.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                '        objglobal.SBOApp.MessageBox("El indicador de Retención  - " & sLinRetCodigo & " - no no está marcado para el interlocutor " & sCliente & ". Por favor, revise el interlocutor.")
                                '        Exit Sub
                                '    End If
                                'End If
#End Region
                                'Grabamos la línea
                                sSQL = "insert into ""@EXO_TMPDOCL"" values('" & iDoc.ToString & "','" & iLinea & "','',0,'" & objglobal.SBOApp.Company.UserName & "',"
                                sSQL &= "'" & sCuenta & "','" & sArt & "','" & sArtDes & "'," & EXO_GLOBALES.DblNumberToText(objglobal, sCantidad.ToString).Replace(",", ".") & ","
                                sSQL &= EXO_GLOBALES.DblNumberToText(objglobal, sprecio.ToString) & "," & EXO_GLOBALES.DblNumberToText(objglobal, sDtoLin.ToString)
                                sSQL &= "," & EXO_GLOBALES.DblNumberToText(objglobal, sTotalServicios.ToString).Replace(",", ".") & ",'" & sLinImpuestoCod & "','" & sLinRetCodigo & "','"
                                sSQL &= sTextoAmpliado & "','" & sTipoLineas & "'," & sPrecioBruto & ",'' ) "
                                oRs.DoQuery(sSQL)

                                iLinea += 1
                            Else
                                sSaltarCabecera = "N"
                            End If

                        Catch ex As Microsoft.VisualBasic.FileIO.MalformedLineException
                            objglobal.SBOApp.StatusBar.SetText("(EXO) - Línea " & ex.Message & " no es válida y se omitirá.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            objglobal.SBOApp.MessageBox("Línea " & ex.Message & " no es válida y se omitirá.")
                        End Try
                    End While
                End Using
            Else
                objglobal.SBOApp.StatusBar.SetText("(EXO) - No se ha encontrado el fichero txt a cargar.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            ' Cerramos el archivo
            FileClose(Apunt)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRCampos, Object))
        End Try
    End Sub
    Public Shared Function CrearDocumentos(ByRef oForm As SAPbouiCOM.Form, ByVal sData As String, ByVal sTDoc As String, ByRef oCompany As SAPbobsCOM.Company, ByRef objglobal As EXO_UIAPI.EXO_UIAPI) As Boolean
        CrearDocumentos = False
        Dim oDoc As SAPbobsCOM.Documents = Nothing
        Dim sExiste As String = "" ' Para comprobar si existen los datos

        Dim sErrorDes As String = "" : Dim sDocAdd As String = "" : Dim sDocEntryAdd As String = "" : Dim sMensaje As String = ""
        Dim sTipoFac As String = "" : Dim sModo As String = "" : Dim sTabla As String = ""

        Dim oRsCab As SAPbobsCOM.Recordset = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim oRsLin As SAPbobsCOM.Recordset = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim oRsLote As SAPbobsCOM.Recordset = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim oRsArt As SAPbobsCOM.Recordset = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim oRsCliente As SAPbobsCOM.Recordset = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim oRsSerie As SAPbobsCOM.Recordset = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim oRsSerieNumber As SAPbobsCOM.Recordset = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sSQL As String = ""
        Dim esprimeralinea As Boolean = True
        Dim esprimeraportes As Boolean = True
        Try
            'If Company.InTransaction = True Then
            '    Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            'End If
            'Company.StartTransaction()
            For i = 0 To oForm.DataSources.DataTables.Item(sData).Rows.Count - 1
                If oForm.DataSources.DataTables.Item(sData).GetValue("Sel", i).ToString = "Y" Then 'Sólo los registros que se han seleccionado
                    sSQL = "Select * FROM ""@EXO_TMPDOC"" Where ""Code""=" & oForm.DataSources.DataTables.Item(sData).GetValue("Code", i).ToString & " and ""U_EXO_USR""='" & objglobal.SBOApp.Company.UserName & "' "
                    oRsCab.DoQuery(sSQL)
                    If oRsCab.RecordCount > 0 Then
#Region "Cabecera"
                        Dim dImpTotal As Double = 0.00
                        esprimeralinea = True
#Region "Tipo Documento"
                        sTipoFac = oRsCab.Fields.Item("U_EXO_TIPOF").Value.ToString
                        sModo = oForm.DataSources.DataTables.Item(sData).GetValue("Modo", i).ToString
                        If sModo = "F" Then
                            Select Case sTipoFac
                                Case "13" 'Factura de ventas
                                    oDoc = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices), SAPbobsCOM.Documents)
                                    sTabla = "OINV"
                                Case "14" 'Abono de ventas
                                    oDoc = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes), SAPbobsCOM.Documents)
                                    sTabla = "ORIN"
                                Case "18" 'Factura de compras
                                    oDoc = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices), SAPbobsCOM.Documents)
                                    sTabla = "OPCH"
                                Case "19" 'Abono de compras
                                    oDoc = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes), SAPbobsCOM.Documents)
                                    sTabla = "ORPC"
                                Case "22" 'Pedidos de compras
                                    oDoc = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders), SAPbobsCOM.Documents)
                                    sTabla = "OPOR"
                            End Select
                        Else
                            oDoc = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts), SAPbobsCOM.Documents)
                            sTabla = "ODRF"
                        End If
                        Select Case sTipoFac
                            Case "13" 'Factura de ventas
                                oDoc.DocObjectCode = SAPbobsCOM.BoObjectTypes.oInvoices
                            Case "14" 'Abono de ventas
                                oDoc.DocObjectCode = SAPbobsCOM.BoObjectTypes.oCreditNotes
                            Case "18" 'Factura de compras
                                oDoc.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseInvoices
                            Case "19" 'Abono de compras
                                oDoc.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes
                            Case "22" 'Pedido de compras
                                oDoc.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseOrders
                        End Select
#End Region
#Region " Serie o Num Documento"
                        If oForm.DataSources.DataTables.Item(sData).GetValue("Nº Documento", i).ToString <> "" Then
                            ''Si se crea en borrador, habrá que buscar el número para no dejarlo crear
                            'If sTabla = "ODRF" Then
                            '    Dim sEncuentraDocNUm As String = ""
                            '    sEncuentraDocNUm = EXO_GLOBALES.GetValueDB(objglobal.conexionSAP.compañia, """ODRF""", """DocNum""", """DocNum""=" & oForm.DataSources.DataTables.Item(sData).GetValue("Nº Documento", i).ToString)
                            '    If sEncuentraDocNUm <> "" Then
                            '        'Como lo ha encontrado, no podemos dejar crearlo

                            '    End If
                            'End If
                            oDoc.HandWritten = SAPbobsCOM.BoYesNoEnum.tYES
                            oDoc.DocNum = oForm.DataSources.DataTables.Item(sData).GetValue("Nº Documento", i).ToString
                        Else
                            'Buscamos la Serie
                            Dim sSerie As String = oForm.DataSources.DataTables.Item(sData).GetValue("Serie", i).ToString
                            sSQL = "SELECT ""Series"" "
                            sSQL += " FROM ""NNM1"" "
                            sSQL += " WHERE ""ObjectCode""=" & sTipoFac & " and ""SeriesName""='" & sSerie & "' "
                            oRsSerie = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
                            oRsSerie.DoQuery(sSQL)
                            If oRsSerie.RecordCount > 0 Then
                                Dim sSerieDoc As String = oRsSerie.Fields.Item("Series").Value.ToString
                                oDoc.Series = sSerieDoc
                            Else
                                objglobal.SBOApp.StatusBar.SetText("(EXO) - No se ha encontrado serie para el documento.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                Exit Function
                            End If
                        End If
#End Region
                        oDoc.CardCode = oRsCab.Fields.Item("U_EXO_CLISAP").Value.ToString
                        oDoc.NumAtCard = oForm.DataSources.DataTables.Item(sData).GetValue("Referencia", i).ToString
                        oDoc.DocCurrency = oRsCab.Fields.Item("U_EXO_MONEDA").Value.ToString
                        If oRsCab.Fields.Item("U_EXO_CTABAL").Value.ToString <> "" Then
                            oDoc.ControlAccount = oRsCab.Fields.Item("U_EXO_CTABAL").Value.ToString
                        End If
                        'Hay que buscar el comercial para asignarlo
                        If oForm.DataSources.DataTables.Item(sData).GetValue("Comercial", i).ToString <> "" Then
                            Dim sCodComercial = ""
                            sCodComercial = EXO_GLOBALES.GetValueDB(objglobal.compañia, """OSLP""", """SlpCode""", """SlpName""='" & oForm.DataSources.DataTables.Item(sData).GetValue("Comercial", i).ToString & "'")
                            If sCodComercial <> "" Then
                                oDoc.SalesPersonCode = sCodComercial
                            Else
                                objglobal.SBOApp.StatusBar.SetText("(EXO) - El empleado del Dpto.  - " & oForm.DataSources.DataTables.Item(sData).GetValue("Comercial", i).ToString & " - no existe.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                If objglobal.SBOApp.MessageBox("El empleado del Dpto.  - " & oForm.DataSources.DataTables.Item(sData).GetValue("Comercial", i).ToString & " - no existe. ¿Desea Crearlo?""?", 1, "Sí", "No") = 1 Then
                                    'EXO_GLOBALES.CrearEmpleado(oForm.DataSources.DataTables.Item(sData).GetValue("Comercial", i).ToString, oCompany, oSboApp)
                                Else
                                    objglobal.SBOApp.StatusBar.SetText("(EXO) - No se puede continuar si no se crea el empleado del Dpto.  - " & oForm.DataSources.DataTables.Item(sData).GetValue("Comercial", i).ToString & ".", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    objglobal.SBOApp.MessageBox("No se puede continuar si no se crea el empleado del Dpto.  - " & oForm.DataSources.DataTables.Item(sData).GetValue("Comercial", i).ToString & ".")
                                    Exit Function
                                End If
                            End If
                        End If
#Region "Fechas"
                        oDoc.DocDate = Year(oForm.DataSources.DataTables.Item(sData).GetValue("F. Contable", i).ToString) & "-" & Month(oForm.DataSources.DataTables.Item(sData).GetValue("F. Contable", i).ToString) & "-" & Day(oForm.DataSources.DataTables.Item(sData).GetValue("F. Contable", i).ToString)
                        Try
                            oDoc.TaxDate = Year(oForm.DataSources.DataTables.Item(sData).GetValue("F. Documento", i).ToString) & "-" & Month(oForm.DataSources.DataTables.Item(sData).GetValue("F. Documento", i).ToString) & "-" & Day(oForm.DataSources.DataTables.Item(sData).GetValue("F. Documento", i).ToString)
                        Catch ex As Exception
                            oDoc.TaxDate = Year(oForm.DataSources.DataTables.Item(sData).GetValue("F. Contable", i).ToString) & "-" & Month(oForm.DataSources.DataTables.Item(sData).GetValue("F. Contable", i).ToString) & "-" & Day(oForm.DataSources.DataTables.Item(sData).GetValue("F. Contable", i).ToString)
                        End Try

                        Dim sFechaVTO As String = ""
                        Try
                            sFechaVTO = oForm.DataSources.DataTables.Item(sData).GetValue("F. Vto", i).ToString
                        Catch ex As Exception
                            sFechaVTO = "1900-01-01"
                        End Try
                        If Year(sFechaVTO) > 1950 Then
                            oDoc.DocDueDate = Year(oForm.DataSources.DataTables.Item(sData).GetValue("F. Vto", i).ToString) & "-" & Month(oForm.DataSources.DataTables.Item(sData).GetValue("F. Vto", i).ToString) & "-" & Day(oForm.DataSources.DataTables.Item(sData).GetValue("F. Vto", i).ToString)
                        Else
                            'oDoc.DocDueDate = Year(oForm.DataSources.DataTables.Item(sData).GetValue("F. Contable", i).ToString) & "-" & Month(oForm.DataSources.DataTables.Item(sData).GetValue("F. Contable", i).ToString) & "-" & Day(oForm.DataSources.DataTables.Item(sData).GetValue("F. Contable", i).ToString)
                        End If
#End Region
                        If oRsCab.Fields.Item("U_EXO_DIRFAC").Value.ToString <> "" Then : oDoc.PayToCode = oRsCab.Fields.Item("U_EXO_DIRFAC").Value.ToString : End If
                        If oRsCab.Fields.Item("U_EXO_DIRENT").Value.ToString <> "" Then : oDoc.ShipToCode = oRsCab.Fields.Item("U_EXO_DIRENT").Value.ToString : End If
#Region "condición y modo de pago"
                        If oRsCab.Fields.Item("U_EXO_CPAGO").Value.ToString <> "" Then
                            oDoc.PaymentMethod = oRsCab.Fields.Item("U_EXO_CPAGO").Value.ToString
                        End If
                        If oRsCab.Fields.Item("U_EXO_GROUPNUM").Value.ToString <> "" Then
                            Dim sGroupNum As Integer = -1
                            Try
                                sGroupNum = EXO_GLOBALES.GetValueDB(objglobal.compañia, """OCTG""", """GroupNum""", """PymntGroup""='" & oRsCab.Fields.Item("U_EXO_GROUPNUM").Value.ToString & "'")
                            Catch ex As Exception
                                sGroupNum = -1
                            End Try
                            If sGroupNum >= 0 Then
                                oDoc.PaymentGroupCode = sGroupNum
                            End If
                        End If
#End Region
#Region "Comentarios"
                        oDoc.Comments = oForm.DataSources.DataTables.Item(sData).GetValue("Comentario", i).ToString
                        oDoc.OpeningRemarks = oRsCab.Fields.Item("U_EXO_CCAB").Value.ToString
                        oDoc.ClosingRemarks = oRsCab.Fields.Item("U_EXO_CPIE").Value.ToString
#End Region
#End Region

#Region "Líneas"
                        'Buscamos las líneas del documento
                        sSQL = "Select * FROM ""@EXO_TMPDOCL"" Where ""Code""=" & oRsCab.Fields.Item("Code").Value.ToString & " and ""U_EXO_USR""='" & objglobal.SBOApp.Company.UserName & "' "
                        oRsLin.DoQuery(sSQL)
                        For iLin = 1 To oRsLin.RecordCount
                            If esprimeralinea = False Then
                                oDoc.Lines.Add()
                            Else
#Region "Tipo Líneas"
                                Select Case oRsLin.Fields.Item("U_EXO_DOCTYPE").Value.ToString
                                    Case "S" : oDoc.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service
                                    Case "I" : oDoc.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items
                                End Select
#End Region
                            End If
                            esprimeralinea = False
#Region "Norma Reparto Coste"
                            Dim sReparto As String = oRsLin.Fields.Item("U_EXO_REPARTO").Value.ToString
                            If sReparto <> "" Then
                                oDoc.Lines.CostingCode = sReparto
                            End If

#End Region
                            If oRsLin.Fields.Item("U_EXO_DOCTYPE").Value.ToString = "I" Then
                                oDoc.Lines.ItemCode = oRsLin.Fields.Item("U_EXO_ART").Value
                                If Trim(oRsLin.Fields.Item("U_EXO_ARTDES").Value.ToString) <> "" Then
                                    oDoc.Lines.ItemDescription = oRsLin.Fields.Item("U_EXO_ARTDES").Value
                                End If
                                oDoc.Lines.UserFields.Fields.Item("U_EXO_PEDIDO").Value = oRsLin.Fields.Item("U_EXO_TXT").Value.ToString
                                oDoc.Lines.Quantity = oRsLin.Fields.Item("U_EXO_CANT").Value
                                oDoc.Lines.UnitPrice = oRsLin.Fields.Item("U_EXO_PRECIO").Value
                                oDoc.Lines.GrossBuyPrice = oRsLin.Fields.Item("U_EXO_PRECIOBRUTO").Value
                                oDoc.Lines.DiscountPercent = oRsLin.Fields.Item("U_EXO_DTOLIN").Value
                                dImpTotal += (oRsLin.Fields.Item("U_EXO_CANT").Value * oRsLin.Fields.Item("U_EXO_PRECIO").Value) - ((oRsLin.Fields.Item("U_EXO_DTOLIN").Value * (oRsLin.Fields.Item("U_EXO_CANT").Value * oRsLin.Fields.Item("U_EXO_PRECIO").Value)) / 100)
                                'Buscamos series disponibles
                                sSQL = "select t0.""SysNumber"" ""SysNumber"" "
                                sSQL &= " FROM ""OSRN"" t0 INNER JOIN ""OSRQ"" t1 on t0.""ItemCode""=t1.""ItemCode"" and t0.""SysNumber""=t1.""SysNumber"" "
                                sSQL &= " WHERE t0.""ItemCode""='" & oRsLin.Fields.Item("U_EXO_ART").Value.ToString & "' and t1.""Quantity"">0 ORDER BY ""SysNumber"""
                                oRsSerieNumber.DoQuery(sSQL)
                                'Incluimos los Lotes
                                sSQL = "Select * FROM ""@EXO_TMPDOCLT"" Where ""Code""=" & oRsLin.Fields.Item("Code").Value.ToString & " And ""U_EXO_USR""='" & objglobal.SBOApp.Company.UserName & "' "
                                sSQL &= " and ""U_EXO_LineId""=" & oRsLin.Fields.Item("LineId").Value.ToString
                                oRsLote.DoQuery(sSQL)
                                For iLote = 1 To oRsLote.RecordCount
                                    'tengo que buscar el artículo para saber si va por lote o serie
                                    Dim sLote As String = "" : Dim sSerie As String = ""
                                    sSQL = "SELECT ""ManSerNum"", ""ManBtchNum"" FROM ""OITM"" WHERE ""ItemCode""='" & oRsLin.Fields.Item("U_EXO_ART").Value & "'"
                                    oRsArt.DoQuery(sSQL)
                                    If oRsArt.RecordCount > 0 Then
                                        sSerie = oRsArt.Fields.Item("ManSerNum").Value.ToString
                                        sLote = oRsArt.Fields.Item("ManBtchNum").Value.ToString
                                    End If
                                    If sLote = "Y" Then
                                        'Creamos el lote de la línea del artículo
                                        oDoc.Lines.BatchNumbers.BatchNumber = oRsLote.Fields.Item("U_EXO_Lote").Value.ToString
                                        oDoc.Lines.BatchNumbers.Quantity = oRsLote.Fields.Item("U_EXO_CANT").Value.ToString
                                        oDoc.Lines.BatchNumbers.Add()
                                    ElseIf sSerie = "Y" Then
                                        'Creamos la serie de la línea del artículo
                                        Select Case sTipoFac
                                            Case "14", "18"
                                                oDoc.Lines.SerialNumbers.InternalSerialNumber = oRsLote.Fields.Item("U_EXO_Lote").Value.ToString
                                            Case "13", "19"
                                                If oRsSerieNumber.RecordCount = 0 Then
                                                    objglobal.SBOApp.StatusBar.SetText("(EXO) - No hay series disponibles para generar el documento.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    objglobal.SBOApp.MessageBox("No hay series disponibles para generar el documento.")
                                                    Exit Function
                                                End If
                                                'Tenemos que buscar el system serial que la cantidad sea superior a 0
                                                Dim iSerialNumber As Integer = 0
                                                iSerialNumber = oRsSerieNumber.Fields.Item("SysNumber").Value
                                                oDoc.Lines.SerialNumbers.SystemSerialNumber = iSerialNumber
                                                oRsSerieNumber.MoveNext()
                                        End Select
                                        oDoc.Lines.SerialNumbers.Quantity = oRsLote.Fields.Item("U_EXO_CANT").Value.ToString
                                        oDoc.Lines.SerialNumbers.Add()
                                    End If
                                    'oDoc.Lines.WarehouseCode = "01"
                                    oRsLote.MoveNext()
                                Next

                                'oDoc.Lines.FreeText = oRsLin.Fields.Item("U_EXO_TXT").Value.ToString
                            ElseIf oRsLin.Fields.Item("U_EXO_DOCTYPE").Value.ToString = "S" Then
                                oDoc.Lines.AccountCode = oRsLin.Fields.Item("U_EXO_CTA").Value
                                oDoc.Lines.LineTotal = oRsLin.Fields.Item("U_EXO_IMPSRV").Value
                                dImpTotal += oRsLin.Fields.Item("U_EXO_IMPSRV").Value

                                oDoc.Lines.ItemDescription = oRsLin.Fields.Item("U_EXO_TXT").Value
                            End If
#Region "Impuesto y Retencion de línea"
                            If oRsLin.Fields.Item("U_EXO_Impuesto").Value <> "" Then
                                oDoc.Lines.VatGroup = oRsLin.Fields.Item("U_EXO_Impuesto").Value
                            End If

                            If oRsLin.Fields.Item("U_EXO_Retencion").Value = "" Then
                                oDoc.Lines.WTLiable = SAPbobsCOM.BoYesNoEnum.tNO
                            Else
                                oDoc.Lines.WTLiable = SAPbobsCOM.BoYesNoEnum.tYES
                                If oRsLin.Fields.Item("U_EXO_Retencion").Value <> "" Then
                                    oDoc.WithholdingTaxData.WTCode = oRsLin.Fields.Item("U_EXO_Retencion").Value
                                    oDoc.WithholdingTaxData.Add()
                                End If
                            End If
#End Region
                            oRsLin.MoveNext()
                        Next
#End Region
#Region "Dto en cabecera"
                        If oRsCab.Fields.Item("U_EXO_TDTO").Value.ToString = "%" Then
                            oDoc.DiscountPercent = oForm.DataSources.DataTables.Item(sData).GetValue("Dto.", i).ToString
                        Else
                            oDoc.DiscountPercent = (oForm.DataSources.DataTables.Item(sData).GetValue("Dto.", i).ToString * 100) / dImpTotal
                        End If
#End Region
                        'grabar el documento
                        If oDoc.Add() <> 0 Then 'Si ocurre un error en la grabación entra
                            sErrorDes = oCompany.GetLastErrorCode & " / " & oCompany.GetLastErrorDescription
                            objglobal.SBOApp.StatusBar.SetText(sErrorDes, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            oForm.DataSources.DataTables.Item(sData).SetValue("Estado", i, "ERROR")
                            oForm.DataSources.DataTables.Item(sData).SetValue("Descripción Estado", i, sErrorDes)
                            oForm.DataSources.DataTables.Item(sData).SetValue("DocEntry", i, "")
                        Else
                            esprimeralinea = True
                            esprimeraportes = True
                            sDocEntryAdd = oCompany.GetNewObjectKey() 'Recoge el último documento creado 
                            oForm.DataSources.DataTables.Item(sData).SetValue("DocEntry", i, sDocEntryAdd)
                            'Buscamos el documento para crear un mensaje
                            sDocAdd = EXO_GLOBALES.GetValueDB(oCompany, """" & sTabla & """", """DocNum""", """DocEntry""=" & sDocEntryAdd)
                            If sModo = "F" Then
                                sModo = ""
                            Else
                                sModo = " borrador "
                            End If
                            oForm.DataSources.DataTables.Item(sData).SetValue("Estado", i, "OK")
                            oForm.DataSources.DataTables.Item(sData).SetValue("Nº Documento", i, sDocAdd)
                            Select Case sTipoFac
                                Case "13" 'Factura de ventas
                                    sMensaje = "(EXO) - Ha sido creada la factura " & sModo & " de ventas Nº" & sDocAdd
                                Case "14" 'Abono de ventas
                                    sMensaje = "(EXO) - Ha sido creado el abono " & sModo & " de ventas Nº" & sDocAdd
                                Case "18" 'Factura de compras
                                    sMensaje = "(EXO) - Ha sido creada la factura " & sModo & " de compras Nº" & sDocAdd
                                Case "19" 'Abono de compras
                                    sMensaje = "(EXO) - Ha sido creado el abono " & sModo & " de compras Nº" & sDocAdd
                                Case "22" 'Pedido de compras
                                    sMensaje = "(EXO) - Ha sido creado el pedido " & sModo & " de compras Nº" & sDocAdd
                            End Select
                            oForm.DataSources.DataTables.Item(sData).SetValue("Descripción Estado", i, sMensaje)
                            objglobal.SBOApp.StatusBar.SetText(sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            '##########################
                            ' Attach_SAP
                            Dim oRsDIR As SAPbobsCOM.Recordset = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
                            Dim sArchivo As String = ""
                            sSQL = "SELECT ""U_EXO_PATH"" FROM ""@EXO_OGEN"" "
                            oRsDIR.DoQuery(sSQL)
                            If oRsDIR.RecordCount > 0 Then
                                sArchivo = oRsDIR.Fields.Item("U_EXO_PATH").Value.ToString
                                sArchivo &= "\08.Historico\ZALANDO\DOC_CARGADOS\VENTAS\FACTURAS\"
                                'Comprobamos que exista el directorio y sino, lo creamos
                                If System.IO.Directory.Exists(sArchivo) = False Then
                                    System.IO.Directory.CreateDirectory(sArchivo)
                                End If
                            End If
                            Dim sFichero As String = CType(oForm.Items.Item("txt_Fich").Specific, SAPbouiCOM.EditText).Value.ToString
                            sFichero = IO.Path.GetFileName(sFichero)
                            sArchivo &= sFichero
                            Attach_SAP(objglobal, oCompany, sTabla, sDocEntryAdd, sTipoFac, sArchivo)
                            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsDIR, Object))
                            '##########################
                        End If
                    End If
                End If
            Next

            'If Company.InTransaction = True Then
            '    Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
            'End If

            CrearDocumentos = True
        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            'If Company.InTransaction = True Then
            '    Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            'End If
            oForm.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oDoc, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsCab, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsLin, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsLote, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsSerie, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsSerieNumber, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsArt, Object))
        End Try
    End Function
    Public Shared Sub Attach_SAP(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef oCompany As SAPbobsCOM.Company, ByVal sTabla As String, ByVal sDocEntryAdd As String, ByVal sTipoFac As String, ByVal sArchivoOrigen As String)
#Region "Variables"
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sAttachPath As String = ""
        Dim sAbsEntry As String = ""
        Dim oDoc As SAPbobsCOM.Documents = Nothing
#End Region

        Try
            'Ruta Anexos
            sSQL = "SELECT t1.AttachPath FROM OADP t1 WITH (NOLOCK)"

            oRs.DoQuery(sSQL)

            If oRs.RecordCount > 0 Then
                sAttachPath = oRs.Fields.Item("AttachPath").Value.ToString
            End If

            If sAttachPath = "" Then
                Throw New Exception("No está definida la ruta de Anexos en SAP.")
            Else
                If Right(sAttachPath, 1) = "\" Then
                    sAttachPath &= IO.Path.GetFileName(sArchivoOrigen)
                Else
                    sAttachPath &= "\" & IO.Path.GetFileName(sArchivoOrigen)
                End If

                EXO_GLOBALES.Copia_Seguridad(oObjGlobal, sArchivoOrigen, sAttachPath)
                'Generamos el documento adjunto
                'If oCompany.InTransaction = True Then
                '    oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                'End If
                sSQL = "Select isnull(""AtcEntry"",0) as ""AtcEntry"" from """ & sTabla & """ where ""DocEntry""=" + sDocEntryAdd.ToString & " And ""ObjType""=" & sTipoFac
                Dim sAtcEntry As String = oObjGlobal.refDi.SQL.sqlStringB1(sSQL)



                'oCompany.StartTransaction()

                Dim oOATT As SAPbobsCOM.Attachments2 = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oAttachments2), SAPbobsCOM.Attachments2)

                If oOATT.GetByKey(sAtcEntry) = True Then
                    oOATT.Lines.Add()
                    oOATT.Lines.FileName = IO.Path.GetFileNameWithoutExtension(sAttachPath)
                    oOATT.Lines.FileExtension = IO.Path.GetExtension(sAttachPath).Substring(1)
                    oOATT.Lines.SourcePath = System.IO.Path.GetDirectoryName(sAttachPath)
                    oOATT.Lines.Override = SAPbobsCOM.BoYesNoEnum.tYES

                    If oOATT.Update <> 0 Then
                        Throw New Exception(oCompany.GetLastErrorCode.ToString & " / " & oCompany.GetLastErrorDescription.Replace("'", ""))
                    End If

                    sAbsEntry = sAtcEntry
                Else
                    oOATT.Lines.Add()
                    oOATT.Lines.FileName = IO.Path.GetFileNameWithoutExtension(sAttachPath)
                    oOATT.Lines.FileExtension = IO.Path.GetExtension(sAttachPath).Substring(1)
                    oOATT.Lines.SourcePath = System.IO.Path.GetDirectoryName(sAttachPath)
                    oOATT.Lines.Override = SAPbobsCOM.BoYesNoEnum.tYES

                    If oOATT.Add <> 0 Then
                        Throw New Exception(oCompany.GetLastErrorCode.ToString & " / " & oCompany.GetLastErrorDescription.Replace("'", ""))
                    End If

                    sAbsEntry = oCompany.GetNewObjectKey
                End If

                'Adjuntamos documento a la entrega
                If sTabla <> "ODRF" Then
                    Select Case sTipoFac
                        Case "13" 'Factura de ventas
                            oDoc = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices), SAPbobsCOM.Documents)

                        Case "14" 'Abono de ventas
                            oDoc = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes), SAPbobsCOM.Documents)

                    End Select
                Else
                    oDoc = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts), SAPbobsCOM.Documents)
                End If

                If oDoc.GetByKey(sDocEntryAdd) = True Then
                    oDoc.AttachmentEntry = sAbsEntry

                    If oDoc.Update <> 0 Then
                        Throw New Exception(oCompany.GetLastErrorCode.ToString & " / " & oCompany.GetLastErrorDescription.Replace("'", ""))
                    End If
                End If

                'If oCompany.InTransaction = True Then
                '    oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                'End If

                oObjGlobal.SBOApp.StatusBar.SetText("(EXO) -  Se ha anexado al documento el fichero " & sAttachPath & " - " & "Operación finalizada con éxito.", EXO_Log.EXO_Log.Tipo.informacion)
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Sub


#End Region
End Class
