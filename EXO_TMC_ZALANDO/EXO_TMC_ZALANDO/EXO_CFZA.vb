Imports System.Xml
Imports SAPbouiCOM
Public Class EXO_CFZA
    Inherits EXO_UIAPI.EXO_DLLBase
    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, usaLicencia, idAddOn)
        cargamenu()
        If actualizar Then
            CargaCampos()
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
            Res = objGlobal.SBOApp.GetLastBatchResults

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
        End If
    End Sub

#End Region
End Class
