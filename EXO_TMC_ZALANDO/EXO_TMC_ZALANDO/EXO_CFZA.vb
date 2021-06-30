﻿Imports System.IO
Imports System.Xml
Imports SAPbouiCOM
Public Class EXO_CFZA
    Inherits EXO_UIAPI.EXO_DLLBase
    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, usaLicencia, idAddOn)
        cargamenu()
        If actualizar Then
            CargaCampos()
            GenerarParametros()
        End If
    End Sub
#Region "Inicialización"
    Private Sub cargamenu()
        Dim sPath As String = ""
        Dim oRsDIR As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sSQL As String = ""

        Dim menuXML As String = objGlobal.funciones.leerEmbebido(Me.GetType(), "XML_MENU.xml")
        objGlobal.SBOApp.LoadBatchActions(menuXML)
        Dim res As String = objGlobal.SBOApp.GetLastBatchResults

        sSQL = "SELECT ""U_EXO_PATH"" FROM ""@EXO_OGEN"" "
        oRsDIR.DoQuery(sSQL)
        If oRsDIR.RecordCount > 0 Then
            sPath = oRsDIR.Fields.Item("U_EXO_PATH").Value.ToString
            sPath &= "\02.Menus"
            If Not System.IO.Directory.Exists(sPath) Then
                System.IO.Directory.CreateDirectory(sPath)
            End If
            If objGlobal.SBOApp.Menus.Exists("EXO-MnCFZA") = True Then
                If sPath <> "" Then
                    If IO.File.Exists(sPath & "\MnCFZA.png") = True Then
                        objGlobal.SBOApp.Menus.Item("EXO-MnCFZA").Image = sPath & "\MnCFZA.png"
                    Else
                        'Sino existe lo copiamos y asignamos
                        EXO_GLOBALES.CopiarRecurso(Reflection.Assembly.GetExecutingAssembly(), "MnCFZA.png", sPath & "\MnCFZA.png")

                        objGlobal.SBOApp.Menus.Item("EXO-MnCFZA").Image = sPath & "\MnCFZA.png"
                    End If
                End If
            End If
        Else
            objGlobal.SBOApp.StatusBar.SetText("(EXO) - No se encuentra la ruta del Kernel. No se puede cargar la imágen del menú.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End If


    End Sub
    Public Overrides Function filtros() As SAPbouiCOM.EventFilters
        Dim fXML As String = objGlobal.funciones.leerEmbebido(Me.GetType(), "XML_Filtros.xml")
        Dim filtro As SAPbouiCOM.EventFilters = New SAPbouiCOM.EventFilters()
        filtro.LoadFromXML(fXML)
        Return filtro
    End Function
    Public Overrides Function menus() As System.Xml.XmlDocument
        Return Nothing
    End Function
    Private Sub CargaCampos()
        If objGlobal.refDi.comunes.esAdministrador Then
            Dim oXML As String = "" : Dim res As String = ""
            Dim udoObj As EXO_Generales.EXO_UDO = Nothing
            'EXO_TMPDOC Cabecera
            oXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UT_EXO_TMPDOC.xml")
            objGlobal.SBOApp.StatusBar.SetText("Validando: UT_EXO_TMPDOC", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            objGlobal.refDi.comunes.LoadBDFromXML(oXML)
            res = objGlobal.SBOApp.GetLastBatchResults

            'EXO_TMPDOC Líneas
            oXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UT_EXO_TMPDOCL.xml")
            objGlobal.SBOApp.StatusBar.SetText("Validando: UT_EXO_TMPDOCL", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            objGlobal.refDi.comunes.LoadBDFromXML(oXML)
            res = objGlobal.SBOApp.GetLastBatchResults

            'EXO_TMPDOC Líneas Lotes
            oXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UT_EXO_TMPDOCLT.xml")
            objGlobal.SBOApp.StatusBar.SetText("Validando: UT_EXO_TMPDOCLT", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            objGlobal.refDi.comunes.LoadBDFromXML(oXML)
            res = objGlobal.SBOApp.GetLastBatchResults

            'UDFs_OINV 
            oXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDFs_OINV.xml")
            objGlobal.SBOApp.StatusBar.SetText("Validando: UDFs_OINV", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            objGlobal.refDi.comunes.LoadBDFromXML(oXML)
            res = objGlobal.SBOApp.GetLastBatchResults
        End If
    End Sub
    Private Sub GenerarParametros()
        If objGlobal.refDi.comunes.esAdministrador Then
            If Not objGlobal.funcionesUI.refDi.OGEN.existeVariable("Zalando_IC") Then
                objGlobal.funcionesUI.refDi.OGEN.fijarValorVariable("Zalando_IC", "Z000000")
                objGlobal.SBOApp.StatusBar.SetText("Creado Parñametro: Zalando_IC", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            End If

            If Not objGlobal.funcionesUI.refDi.OGEN.existeVariable("Zalando_Remark") Then
                objGlobal.funcionesUI.refDi.OGEN.fijarValorVariable("Zalando_Remark", "MARKETPLACE")
                objGlobal.SBOApp.StatusBar.SetText("Creado Parñametro: Zalando_Remark", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            End If
            If Not objGlobal.funcionesUI.refDi.OGEN.existeVariable("Zalando_Fich_con_Cabecera") Then
                objGlobal.funcionesUI.refDi.OGEN.fijarValorVariable("Zalando_Fich_con_Cabecera", "Y")
                objGlobal.SBOApp.StatusBar.SetText("Creado Parñametro: Zalando_Fich_con_Cabecera", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            End If
        End If
    End Sub

#End Region
#Region "Menu Eventos"
    Public Overrides Function SBOApp_MenuEvent(infoEvento As MenuEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        Try
            If infoEvento.BeforeAction = True Then
            Else
                Select Case infoEvento.MenuUID
                    Case "EXO-MnCFZA"
                        'Cargamos pantalla de gestión.
                        If CargarFormCDOC() = False Then
                            Exit Function
                        End If
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
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Public Function CargarFormCDOC() As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim oFP As SAPbouiCOM.FormCreationParams = Nothing

        CargarFormCDOC = False

        Try
            oFP = CType(objGlobal.SBOApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams), SAPbouiCOM.FormCreationParams)
            oFP.XmlData = objGlobal.funciones.leerEmbebido(Me.GetType(), "EXO_CFZA.srf")

            Try
                oForm = objGlobal.SBOApp.Forms.AddEx(oFP)
            Catch ex As Exception
                If ex.Message.StartsWith("Form - already exists") = True Then
                    objGlobal.SBOApp.StatusBar.SetText("El formulario ya está abierto.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit Function
                ElseIf ex.Message.StartsWith("Se produjo un error interno") = True Then 'Falta de autorización
                    Exit Function
                End If
            End Try


            CargarFormCDOC = True
        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oForm.Visible = True
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function
#End Region
#Region "Eventos"
    Public Overrides Function SBOApp_ItemEvent(ByVal infoEvento As ItemEvent) As Boolean
        Try
            If infoEvento.InnerEvent = False Then
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_CFZA"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                    If EventHandler_ItemPressed_After(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE

                            End Select
                    End Select
                ElseIf infoEvento.BeforeAction = True Then
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_CFZA"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED
                                    If EventHandler_Matrix_Link_Press_Before(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                            End Select
                    End Select
                End If
            Else
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_CFZA"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE
                                    If EventHandler_FORM_VISIBLE(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS

                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                            End Select

                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_CFZA"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED

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
    Private Function EventHandler_FORM_VISIBLE(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        EventHandler_FORM_VISIBLE = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            If oForm.Visible = True Then
                oForm.Items.Item("btn_Carga").Enabled = False
            End If

            EventHandler_FORM_VISIBLE = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Private Function EventHandler_Matrix_Link_Press_Before(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sModo As String = ""
        EventHandler_Matrix_Link_Press_Before = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            If pVal.ItemUID = "grd_DOC" Then
                If pVal.ColUID = "DocEntry" Then
                    sModo = CType(oForm.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).DataTable.GetValue("Modo", pVal.Row).ToString
                    If sModo = "F" Then
                        CType(CType(oForm.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item("DocEntry"), SAPbouiCOM.EditTextColumn).LinkedObjectType = CType(oForm.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).DataTable.GetValue("Tipo", pVal.Row).ToString
                    Else
                        CType(CType(oForm.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item("DocEntry"), SAPbouiCOM.EditTextColumn).LinkedObjectType = 112
                    End If
                End If
            End If
            EventHandler_Matrix_Link_Press_Before = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Private Function EventHandler_ItemPressed_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim oRsDIR As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sSQL As String = ""
        Dim oColumnTxt As SAPbouiCOM.EditTextColumn = Nothing
        Dim oColumnChk As SAPbouiCOM.CheckBoxColumn = Nothing
        Dim sTipoArchivo As String = ""
        Dim sArchivoOrigen As String = ""
        Dim sArchivo As String = ""
        Dim sNomFICH As String = ""
        EventHandler_ItemPressed_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
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


            Select Case pVal.ItemUID
                Case "btn_Carga"
                    If objGlobal.SBOApp.MessageBox("¿Está seguro que quiere generar los Documentos seleccionados?", 1, "Sí", "No") = 1 Then
                        If ComprobarDOC(oForm, "DT_DOC") = True Then
                            oForm.Items.Item("btn_Carga").Enabled = False
                            'Generamos facturas
                            objGlobal.SBOApp.StatusBar.SetText("Creando Documentos ... Espere por favor.", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            oForm.Freeze(True)
                            If EXO_GLOBALES.CrearDocumentos(oForm, "DT_DOC", "FACTURA", objGlobal.compañia, objGlobal) = False Then
                                Exit Function
                            End If
                            oForm.Freeze(False)
                            objGlobal.SBOApp.StatusBar.SetText("Fin del proceso.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            objGlobal.SBOApp.MessageBox("Fin del Proceso" & ChrW(10) & ChrW(13) & "Por favor, revise el Log para ver las operaciones realizadas.")
                            oForm.Items.Item("btn_Carga").Enabled = True
                        End If
                    End If
                Case "btn_Fich"
                    Limpiar_Grid(oForm)
                    'Cargar Fichero para leer
                    sTipoArchivo = "Ficheros CSV|*.csv|Texto|*.txt"

                    'Tenemos que controlar que es cliente o web
                    If objGlobal.SBOApp.ClientType = SAPbouiCOM.BoClientType.ct_Browser Then
                        sArchivoOrigen = objGlobal.SBOApp.GetFileFromBrowser() 'Modificar
                    Else
                        'Controlar el tipo de fichero que vamos a abrir según campo de formato
                        sArchivoOrigen = objGlobal.funciones.OpenDialogFiles("Abrir archivo como", sTipoArchivo)
                    End If

                    If Len(sArchivoOrigen) = 0 Then
                        CType(oForm.Items.Item("txt_Fich").Specific, SAPbouiCOM.EditText).Value = ""
                        objGlobal.SBOApp.MessageBox("Debe indicar un archivo a importar.")
                        objGlobal.SBOApp.StatusBar.SetText("(EXO) - Debe indicar un archivo a importar.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                        oForm.Items.Item("btn_Carga").Enabled = False
                        Exit Function
                    Else
                        ''Borramos los ficheros de más de 1 mes
                        'EXO_GLOBALES.BorrarFicheros(sArchivo)

                        CType(oForm.Items.Item("txt_Fich").Specific, SAPbouiCOM.EditText).Value = sArchivoOrigen
                        sNomFICH = IO.Path.GetFileName(sArchivoOrigen)
                        sArchivo = sArchivo & sNomFICH

                        'Hacemos copia de seguridad para tratarlo
                        Copia_Seguridad(sArchivoOrigen, sArchivo)
                        'Ahora abrimos el fichero para tratarlo
                        TratarFichero(sArchivo, oForm)
                        oForm.Items.Item("btn_Carga").Enabled = True
                    End If
            End Select

            EventHandler_ItemPressed_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oForm.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oColumnTxt, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oColumnChk, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsDIR, Object))
        End Try
    End Function
    Private Function ComprobarDOC(ByRef oForm As SAPbouiCOM.Form, ByVal sFra As String) As Boolean
        Dim bLineasSel As Boolean = False

        ComprobarDOC = False

        Try
            For i As Integer = 0 To oForm.DataSources.DataTables.Item(sFra).Rows.Count - 1
                If oForm.DataSources.DataTables.Item(sFra).GetValue("Sel", i).ToString = "Y" Then
                    bLineasSel = True
                    Exit For
                End If
            Next

            If bLineasSel = False Then
                objGlobal.SBOApp.MessageBox("Debe seleccionar al menos una línea.")
                Exit Function
            End If

            ComprobarDOC = bLineasSel

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Private Sub Limpiar_Grid(ByRef oForm As SAPbouiCOM.Form)
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Try
            oForm.Freeze(True)
            'Limpiamos grid
            'Borrar tablas temporales por usuario activo
            sSQL = "DELETE FROM ""@EXO_TMPDOC"" where ""U_EXO_USR""='" & objGlobal.SBOApp.Company.UserName & "'  "
            oRs.DoQuery(sSQL)
            sSQL = "DELETE FROM ""@EXO_TMPDOCL"" where ""U_EXO_USR""='" & objGlobal.SBOApp.Company.UserName & "'  "
            oRs.DoQuery(sSQL)
            sSQL = "DELETE FROM ""@EXO_TMPDOCLT"" where ""U_EXO_USR""='" & objGlobal.SBOApp.Company.UserName & "'  "
            oRs.DoQuery(sSQL)
            'Ahora cargamos el Grid con los datos guardados
            objGlobal.SBOApp.StatusBar.SetText("Cargando Documentos en pantalla ... Espere por favor", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            sSQL = "SELECT 'Y' as ""Sel"",""Code"",""U_EXO_MODO"" as ""Modo"", '     ' as ""Estado"",""U_EXO_TIPOF"" As ""Tipo"",'      ' as ""DocEntry"", ""U_EXO_Serie"" as ""Serie"",""U_EXO_DOCNUM"" as ""Nº Documento"","
            sSQL &= " ""U_EXO_REF"" as ""Referencia"", ""U_EXO_MONEDA"" as ""Moneda"", ""U_EXO_COMER"" as ""Comercial"", ""U_EXO_CLISAP"" as ""Interlocutor SAP"", ""U_EXO_ADDID"" as ""Interlocutor Ext."", "
            sSQL &= " ""U_EXO_FCONT"" as ""F. Contable"", ""U_EXO_FDOC"" as ""F. Documento"", ""U_EXO_FVTO"" as ""F. Vto"", ""U_EXO_TDTO"" as ""T. Dto."", ""U_EXO_DTO"" as ""Dto."",  "
            sSQL &= " ""U_EXO_CPAGO"" as ""Vía Pago"", ""U_EXO_GROUPNUM"" as ""Cond. Pago"", ""U_EXO_COMENT"" as ""Comentario"", "
            sSQL &= " CAST('' as varchar(254)) as ""Descripción Estado"" "
            sSQL &= " From ""@EXO_TMPDOC"" "
            sSQL &= " WHERE ""U_EXO_USR""='" & objGlobal.SBOApp.Company.UserName & "' "
            sSQL &= " ORDER BY ""U_EXO_FDOC"", ""U_EXO_MODO"", ""U_EXO_TIPOF"" "
            'Cargamos grid
            oForm.DataSources.DataTables.Item("DT_DOC").ExecuteQuery(sSQL)
            FormateaGrid(oForm)
        Catch ex As Exception
            Throw ex
        Finally
            oForm.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Sub
    Private Sub FormateaGrid(ByRef oform As SAPbouiCOM.Form)
        Dim oColumnTxt As SAPbouiCOM.EditTextColumn = Nothing
        Dim oColumnChk As SAPbouiCOM.CheckBoxColumn = Nothing
        Dim oColumnCb As SAPbouiCOM.ComboBoxColumn = Nothing
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Try
            oform.Freeze(True)
            CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(0).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
            oColumnChk = CType(CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(0), SAPbouiCOM.CheckBoxColumn)
            oColumnChk.Editable = True
            oColumnChk.Width = 30

            CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(1).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
            oColumnTxt = CType(CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(1), SAPbouiCOM.EditTextColumn)
            oColumnTxt.Editable = False
            oColumnTxt.Width = 40

            CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(2).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
            oColumnCb = CType(CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(2), SAPbouiCOM.ComboBoxColumn)
            oColumnCb.ValidValues.Add("F", "Factura")
            oColumnCb.ValidValues.Add("B", "Borrador")
            oColumnCb.Editable = True
            oColumnCb.Width = 70
            oColumnCb.DisplayType = BoComboDisplayType.cdt_Description

            CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(3).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
            oColumnTxt = CType(CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(3), SAPbouiCOM.EditTextColumn)
            oColumnTxt.Editable = False
            oColumnTxt.Width = 50

            CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(4).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
            oColumnCb = CType(CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(4), SAPbouiCOM.ComboBoxColumn)
            oColumnCb.ValidValues.Add("13", "Factura de Ventas")
            oColumnCb.ValidValues.Add("14", "Abonos de Venta")
            oColumnCb.ValidValues.Add("18", "Factura de Compras")
            oColumnCb.ValidValues.Add("19", "Abono de Compras")
            oColumnCb.ValidValues.Add("22", "Pedido de Compras")
            oColumnCb.DisplayType = BoComboDisplayType.cdt_Description
            oColumnCb.Editable = False
            oColumnCb.Width = 100

            For i = 5 To 10
                If i <> 8 Then
                    CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                    oColumnTxt = CType(CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                    If i <> 10 Then
                        oColumnTxt.Editable = False
                    End If
                End If
                If i = 5 Then
                    oColumnTxt.LinkedObjectType = "22"
                ElseIf i = 10 Then
                    'Comercial
                    oColumnTxt.ChooseFromListUID = "CFL_0"
                    oColumnTxt.ChooseFromListAlias = "SlpName"
                    oColumnTxt.Width = 150
                End If
            Next

            For i = 11 To 21
                CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                oColumnTxt = CType(CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                oColumnTxt.Editable = False
                Select Case i
                    Case 11, 12, 13, 14, 15 : oColumnTxt.Width = 70
                    Case 16, 17 : oColumnTxt.Width = 45
                    Case 21 : oColumnTxt.Width = 300
                End Select

                If i = 11 Then
                    oColumnTxt.LinkedObjectType = "2"
                End If
                Select Case i
                    Case 16, 17 : oColumnTxt.RightJustified = True
                End Select
            Next
        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oform.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Sub
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
        objGlobal.SBOApp.StatusBar.SetText("(EXO) - Comienza la Copia de seguridad del fichero - " & sArchivoOrigen & " -.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
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
        objGlobal.SBOApp.StatusBar.SetText("(EXO) - Copia de Seguridad realizada Correctamente", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
    End Sub
    Private Sub TratarFichero(ByVal sArchivo As String, ByRef oForm As SAPbouiCOM.Form)
        Dim myStream As StreamReader = Nothing
        Dim Reader As XmlTextReader = New XmlTextReader(myStream)
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sExiste As String = "" ' Para comprobar si existen los datos
        Dim sDelimitador As String = ""
        Try

#Region "TXT|CSV"
            EXO_GLOBALES.TratarFichero_TXT(sArchivo, sDelimitador, oForm, objGlobal.compañia, objGlobal)
#End Region
            objGlobal.SBOApp.StatusBar.SetText("(EXO) - Se ha leido correctamente el fichero.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

#Region "cargar Grid con los datos leidos"
            'Ahora cargamos el Grid con los datos guardados
            objGlobal.SBOApp.StatusBar.SetText("Cargando Documentos en pantalla ... Espere por favor", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            sSQL = "SELECT 'Y' as ""Sel"",""Code"",""U_EXO_MODO"" as ""Modo"", '     ' as ""Estado"",""U_EXO_TIPOF"" As ""Tipo"",'      ' as ""DocEntry"",""U_EXO_Serie"" as ""Serie"",""U_EXO_DOCNUM"" as ""Nº Documento"","
            sSQL &= " ""U_EXO_REF"" as ""Referencia"", ""U_EXO_MONEDA"" as ""Moneda"", ""U_EXO_COMER"" as ""Comercial"", ""U_EXO_CLISAP"" as ""Interlocutor SAP"", ""U_EXO_ADDID"" as ""Interlocutor Ext."", "
            sSQL &= " ""U_EXO_FCONT"" as ""F. Contable"", ""U_EXO_FDOC"" as ""F. Documento"", ""U_EXO_FVTO"" as ""F. Vto"", ""U_EXO_TDTO"" as ""T. Dto."", ""U_EXO_DTO"" as ""Dto."",  "
            sSQL &= " ""U_EXO_CPAGO"" as ""Vía Pago"", ""U_EXO_GROUPNUM"" as ""Cond. Pago"", ""U_EXO_COMENT"" as ""Comentario"", "
            sSQL &= " CAST('' as varchar(254)) as ""Descripción Estado"" "
            sSQL &= " From ""@EXO_TMPDOC"" "
            sSQL &= " WHERE ""U_EXO_USR""='" & objGlobal.SBOApp.Company.UserName & "' "
            sSQL &= " ORDER BY ""U_EXO_FDOC"",""U_EXO_MODO"", ""U_EXO_TIPOF"" "
            'Cargamos grid
            oForm.DataSources.DataTables.Item("DT_DOC").ExecuteQuery(sSQL)
            FormateaGrid(oForm)
#End Region
            oForm.Freeze(True)
            objGlobal.SBOApp.StatusBar.SetText("Fin del proceso.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            objGlobal.SBOApp.MessageBox("Se ha leido correctamente el fichero. Fin del proceso")
        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oForm.Freeze(False)
            myStream = Nothing
            Reader.Close()
            Reader = Nothing
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Sub
#End Region
End Class
