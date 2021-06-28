Imports System.IO
Imports System.Xml
Imports SAPbouiCOM
Imports System.Windows.Forms
Public Class EXO_OWTR
    Inherits EXO_UIAPI.EXO_DLLBase

    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, usaLicencia, idAddOn)

        If actualizar Then
            ' cargaCampos()
        End If
    End Sub
#Region "Inicialización"
    Public Overrides Function filtros() As SAPbouiCOM.EventFilters
        Dim fXML As String = objGlobal.funciones.leerEmbebido(Me.GetType(), "XML_Filtros.xml")
        Dim filtro As SAPbouiCOM.EventFilters = New SAPbouiCOM.EventFilters()
        filtro.LoadFromXML(fXML)
        Return filtro
    End Function
    Public Overrides Function menus() As System.Xml.XmlDocument
        Return Nothing
    End Function
    'Private Sub cargaCampos()
    '    Dim sXML As String = ""
    '    Dim res As String = ""

    '    If objGlobal.refDi.comunes.esAdministrador Then
    '        'UDO Datos maestros marcas
    '        sXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDO_EXO004.xml")
    '        objGlobal.refDi.comunes.LoadBDFromXML(sXML)
    '        objGlobal.SBOApp.StatusBar.SetText("Validando: UDO EXO004", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
    '        res = objGlobal.SBOApp.GetLastBatchResults
    '    End If
    'End Sub
#End Region
#Region "Eventos"
    Public Overrides Function SBOApp_MenuEvent(ByVal infoEvento As MenuEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        Try
            If infoEvento.BeforeAction = True Then

            Else
                Select Case infoEvento.MenuUID

                End Select
            End If

            Return MyBase.SBOApp_MenuEvent(infoEvento)

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)

            Return False
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)

            Return False
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function

    Public Overrides Function SBOApp_FormDataEvent(ByVal infoEvento As BusinessObjectInfo) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        'Dim sCardCode As String = ""
        'Dim oXml As New Xml.XmlDocument

        Try
            If infoEvento.BeforeAction = True Then
                Select Case infoEvento.FormTypeEx
                    Case "940"
                        Select Case infoEvento.EventType

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE

                        End Select
                End Select
            Else
                Select Case infoEvento.FormTypeEx
                    Case "940"
                        Select Case infoEvento.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                                If infoEvento.ActionSuccess Then

                                End If
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                                If infoEvento.ActionSuccess Then

                                End If
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE
                                If infoEvento.ActionSuccess Then
                                End If
                        End Select
                End Select
            End If

            Return MyBase.SBOApp_FormDataEvent(infoEvento)

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)

            Return False
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)

            Return False
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function

    Public Overrides Function SBOApp_ItemEvent(ByVal infoEvento As ItemEvent) As Boolean
        Try
            If infoEvento.InnerEvent = False Then
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "940"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                    If EventHandler_ItemPressed_After(infoEvento) = False Then
                                        Return False
                                    End If

                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE

                            End Select
                    End Select
                ElseIf infoEvento.BeforeAction = True Then
                    Select Case infoEvento.FormTypeEx
                        Case "940"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE

                            End Select
                    End Select
                End If
            Else
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "940"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE

                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                    If EventHandler_Form_Load(infoEvento) = False Then
                                        Return False
                                    End If

                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                    If EventHandler_ItemPressed_After(infoEvento) = False Then
                                        Return False
                                    End If
                            End Select
                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "940"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED

                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE

                            End Select
                    End Select
                End If
            End If

            Return MyBase.SBOApp_ItemEvent(infoEvento)

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)

            Return False
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)

            Return False
        End Try
    End Function

    Private Function EventHandler_Form_Load(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        EventHandler_Form_Load = False

        Try
            'Recuperar el formulario
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            oForm.Visible = False

            objGlobal.SBOApp.StatusBar.SetText("Presentando información...Espere por favor", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            'crear Botón
            Dim oItem As SAPbouiCOM.Item
            oItem = oForm.Items.Add("btnFich", BoFormItemTypes.it_BUTTON)
            oItem.Top = oForm.Items.Item("2").Top
            oItem.Left = oForm.Items.Item("2").Left + oForm.Items.Item("2").Width + 10
            oItem.Height = oForm.Items.Item("2").Height
            oItem.Width = oForm.Items.Item("2").Width + 20

            CType(oForm.Items.Item("btnFich").Specific, SAPbouiCOM.Button).Caption = "Crear Fichero"
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            If oForm IsNot Nothing Then oForm.Visible = True

            EventHandler_Form_Load = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            If oForm IsNot Nothing Then oForm.Visible = True

            Throw exCOM
        Catch ex As Exception
            If oForm IsNot Nothing Then oForm.Visible = True

            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function

    Private Function EventHandler_ItemPressed_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oRs As SAPbobsCOM.Recordset = Nothing : Dim oRsDIR As SAPbobsCOM.Recordset = Nothing
        Dim sSQL As String = ""
        Dim sDocEntry As String = ""
        Dim sPath As String = "" : Dim sRutaFich As String = "" : Dim sNomFich As String = ""
        Dim sLinea As String = ""
        EventHandler_ItemPressed_After = False

        Try
            oRs = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            oRsDIR = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            If pVal.ItemUID = "btnFich" Then 'Botón crear fichero zalando
                sDocEntry = oForm.DataSources.DBDataSources.Item("OWTR").GetValue("DocEntry", 0).Trim
                sSQL = "SELECT isnull(IT.""codebars"",'') ""codebars"", TR1.""Quantity"" FROM ""WTR1"" TR1 "
                sSQL &= " INNER JOIN ""OITM"" IT ON TR1.""ItemCode""= IT.""ItemCode"" "
                sSQL &= " WHERE TR1.DocEntry=" & sDocEntry
                oRs.DoQuery(sSQL)
                If oRs.RecordCount > 0 Then
                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - Creando documento ... Espere por favor.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    'Guardamos el fichero en Hcos
                    sSQL = "SELECT ""U_EXO_PATH"" FROM ""@EXO_OGEN"" "
                    oRsDIR.DoQuery(sSQL)
                    If oRsDIR.RecordCount > 0 Then
                        sPath = oRsDIR.Fields.Item("U_EXO_PATH").Value.ToString
                        sPath &= "\08.Historico\ZALANDO"
                        If Not System.IO.Directory.Exists(sPath) Then
                            System.IO.Directory.CreateDirectory(sPath)
                        End If
                        'creamos fichero
                        sNomFich = "Traslado_" & oForm.DataSources.DBDataSources.Item("OWTR").GetValue("DocNum", 0).Trim
                        sRutaFich = Path.Combine(sPath & "\" & sNomFich & ".csv")
                        FileOpen(1, sRutaFich, OpenMode.Output)
                        sLinea = "EAN;Quantity"
                        PrintLine(1, sLinea)
                        For i = 0 To oRs.RecordCount - 1
                            sLinea = oRs.Fields.Item("codebars").Value.ToString & ";" & oRs.Fields.Item("Quantity").Value.ToString
                            PrintLine(1, sLinea)
                            oRs.MoveNext()
                        Next
                        'cerramos fichero
                        FileClose(1)
                        objGlobal.SBOApp.StatusBar.SetText("Fichero creado: " & sRutaFich, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        'objGlobal.SBOApp.MessageBox("Fichero creado: " & sRutaFich)
#Region "Pedimos directorio para guardar donde quiere el usuario"
                        'pedimos directorio para guardar
                        If objGlobal.SBOApp.ClientType = BoClientType.ct_Browser Then
                            sPath = objGlobal.SBOApp.SendFileToBrowser(IO.Path.GetFileName(sRutaFich))
                        Else
                            sPath = objGlobal.funciones.SaveDialogFiles("Guardar archivo como", "CSV|*.csv|Fichero TXT|*.txt", IO.Path.GetFileName(sRutaFich), Environment.SpecialFolder.Desktop)
                        End If
                        If Len(sPath) = 0 Then
                            objGlobal.SBOApp.MessageBox("No ha indicado un directorio para guardar. EL fichero está guardado en los Hcos. en el servidor: " & sRutaFich)
                            objGlobal.SBOApp.StatusBar.SetText("(EXO) - No ha indicado un directorio para guardar. EL fichero está guardado en los Hcos. en el servidor: " & sRutaFich, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        Else
                            'Guardamos
                            Copia_Seguridad(sRutaFich, sPath)
                        End If
#End Region
                    Else
                        objGlobal.SBOApp.StatusBar.SetText("(EXO) - No se encuentra la ruta del Kernel. No se puede continuar.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        objGlobal.SBOApp.MessageBox("No se encuentra la ruta del Kernel. No se puede continuar.")
                    End If
                Else
                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - No se encuentra líneas para crear documento.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End If
            End If

            EventHandler_ItemPressed_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            oForm.Freeze(False)

            Throw exCOM
        Catch ex As Exception
            oForm.Freeze(False)

            Throw ex
        Finally

            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function


#End Region
    Private Sub Copia_Seguridad(ByVal sArchivoOrigen As String, ByVal sArchivo As String)
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
        objGlobal.SBOApp.StatusBar.SetText("(EXO) - Comienza la Copia del fichero - " & sArchivoOrigen & " -.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        If objGlobal.SBOApp.ClientType = BoClientType.ct_Browser Then
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
        objGlobal.SBOApp.StatusBar.SetText("(EXO) - Fichero guardado: " & sArchivo, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        objGlobal.SBOApp.MessageBox("Fichero guardado." & sArchivo)
    End Sub
End Class
