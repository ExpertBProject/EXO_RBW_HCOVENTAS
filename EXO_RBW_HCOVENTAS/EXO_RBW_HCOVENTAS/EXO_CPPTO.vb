Imports System.IO
Imports System.Xml
Imports OfficeOpenXml
Imports SAPbouiCOM
Public Class EXO_CPPTO
    Inherits EXO_UIAPI.EXO_DLLBase
    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, False, idAddOn)

        cargamenu()

    End Sub
    Public Overrides Function filtros() As EventFilters
        Dim filtrosXML As Xml.XmlDocument = New Xml.XmlDocument
        filtrosXML.LoadXml(objGlobal.funciones.leerEmbebido(Me.GetType(), "XML_FILTROS.xml"))
        Dim filtro As SAPbouiCOM.EventFilters = New SAPbouiCOM.EventFilters()
        filtro.LoadFromXML(filtrosXML.OuterXml)

        Return filtro
    End Function
    Public Overrides Function menus() As System.Xml.XmlDocument
        Return Nothing
    End Function

    Private Sub cargamenu()
        Dim Path As String = ""
        Dim menuXML As String = objGlobal.funciones.leerEmbebido(Me.GetType(), "XML_MENU.xml")
        objGlobal.SBOApp.LoadBatchActions(menuXML)
        Dim res As String = objGlobal.SBOApp.GetLastBatchResults
    End Sub

    Public Overrides Function SBOApp_MenuEvent(infoEvento As MenuEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        Try
            If infoEvento.BeforeAction = True Then

            Else

                Select Case infoEvento.MenuUID
                    Case "EXO-MnCPPTO"
                        If CargarForm() = False Then
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
    Public Function CargarForm() As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sSQL As String = ""
        Dim oFP As SAPbouiCOM.FormCreationParams = Nothing
        Dim EXO_Xml As New EXO_UIAPI.EXO_XML(objGlobal)

        CargarForm = False

        Try
            oFP = CType(objGlobal.SBOApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams), SAPbouiCOM.FormCreationParams)
            oFP.XmlData = objGlobal.leerEmbebido(Me.GetType(), "EXO_CPPTO.srf")

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
            CargaComboFormato(oForm)

            CargarForm = True
        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            oForm.Visible = True
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Private Function CargaComboFormato(ByRef oForm As SAPbouiCOM.Form) As Boolean

        CargaComboFormato = False
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Try
            oForm.Freeze(True)
            sSQL = "(Select 'EXCEL' as ""Code"",'EXCEL' as ""Name"" FROM ""DUMMY"") "
            oRs.DoQuery(sSQL)
            If oRs.RecordCount > 0 Then
                objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
                CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Select("EXCEL", BoSearchKey.psk_ByValue)
            Else
                objGlobal.SBOApp.StatusBar.SetText("(EXO) - Por favor, antes de continuar, revise la parametrización.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If
            CargaComboFormato = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oForm.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function
    Public Overrides Function SBOApp_ItemEvent(infoEvento As ItemEvent) As Boolean
        Try
            If infoEvento.InnerEvent = False Then
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_CPPTO"
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
                        Case "EXO_CPPTO"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                            End Select
                    End Select
                End If
            Else
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_CPPTO"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE

                                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS

                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                            End Select

                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_CPPTO"
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
    Private Function EventHandler_ItemPressed_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sTipoArchivo As String = ""
        Dim sArchivoOrigen As String = ""
        Dim sArchivo As String = objGlobal.pathHistorico & "\DOC_CARGADOS\PPTOVTAS\"
        Dim sNomFICH As String = ""
        EventHandler_ItemPressed_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            'Comprobamos que exista el directorio y sino, lo creamos
            If System.IO.Directory.Exists(sArchivo) = False Then
                System.IO.Directory.CreateDirectory(sArchivo)
            End If
            Select Case pVal.ItemUID
                Case "btn_Carga"
                    If objGlobal.SBOApp.MessageBox("¿Está seguro que quiere tratar el fichero seleccionado?", 1, "Sí", "No") = 1 Then
                        sArchivoOrigen = CType(oForm.Items.Item("txt_Fich").Specific, SAPbouiCOM.EditText).Value
                        sNomFICH = IO.Path.GetFileName(sArchivoOrigen)
                        sArchivo = sArchivo & sNomFICH
                        'Hacemos copia de seguridad para tratarlo
                        Copia_Seguridad(sArchivoOrigen, sArchivo)
                        oForm.Items.Item("btn_Carga").Enabled = False
                        objGlobal.SBOApp.StatusBar.SetText("Cargando datos de Ppto. de Ventas ... Espere por favor.", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        oForm.Freeze(True)
                        'Ahora abrimos el fichero para tratarlo
                        TratarFichero(sArchivo, CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString, oForm)
                        oForm.Freeze(False)
                        objGlobal.SBOApp.StatusBar.SetText("Fin del proceso.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        objGlobal.SBOApp.MessageBox("Fin del Proceso" & ChrW(10) & ChrW(13) & "Por favor, revise el Log para ver las operaciones realizadas.")
                        oForm.Items.Item("btn_Carga").Enabled = True
                    Else
                        objGlobal.SBOApp.StatusBar.SetText("Se ha cancelado el proceso.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End If
                Case "btn_Fich"
                    'Cargar Fichero para leer
                    If CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString <> "--" Then
                        sTipoArchivo = "Libro de Excel|*.xlsx|Excel 97-2003|*.xls"

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
                            CType(oForm.Items.Item("txt_Fich").Specific, SAPbouiCOM.EditText).Value = sArchivoOrigen
                            oForm.Items.Item("btn_Carga").Enabled = True
                        End If
                    Else
                        objGlobal.SBOApp.MessageBox("No ha seleccionado el formato a importar." & ChrW(10) & ChrW(13) & " Antes de continuar Seleccione un formato de los que se ha creado en la parametrización.")
                        objGlobal.SBOApp.StatusBar.SetText("(EXO) - No ha seleccionado el formato a importar." & ChrW(10) & ChrW(13) & " Antes de continuar Seleccione un formato de los que se ha creado en la parametrización.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Active = True
                    End If
            End Select

            EventHandler_ItemPressed_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            oForm.Freeze(False)
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            oForm.Freeze(False)
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            oForm.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
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
    Private Sub TratarFichero(ByVal sArchivo As String, ByVal sTipoArchivo As String, ByRef oForm As SAPbouiCOM.Form)
        Dim myStream As StreamReader = Nothing
        Dim Reader As XmlTextReader = New XmlTextReader(myStream)
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sExiste As String = "" ' Para comprobar si existen los datos
        Dim sDelimitador As String = ""
        Try
            Select Case sTipoArchivo
                Case "EXCEL"
                    TratarFichero_Excel(sArchivo, oForm)
                Case Else
                    objGlobal.SBOApp.StatusBar.SetText("(EXO) -El tipo de fichero a importar no está contemplado. Avise a su Administrador.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    objGlobal.SBOApp.MessageBox("El tipo de fichero a importar no está contemplado. Avise a su Administrador.")
                    Exit Sub
            End Select
            objGlobal.SBOApp.StatusBar.SetText("(EXO) - Se ha leido correctamente el fichero.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            oForm.Freeze(True)

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
    Private Sub TratarFichero_Excel(ByVal sArchivo As String, ByRef oForm As SAPbouiCOM.Form)
#Region "Variables"
        Dim pck As ExcelPackage = Nothing
        Dim sTipo As String = "" : Dim sAnno As String = "" : Dim sICCod As String = "" : Dim sICName As String = "" : Dim sComercial As String = ""
        Dim sItemCode As String = "" : Dim sItemName As String = "" : Dim sDivision As String = ""
        Dim sPeriodo As String = "" : Dim sPais As String = "" : Dim sProvincia As String = ""
        Dim dCantA As Double = 0 : Dim dCantB As Double = 0 : Dim dPrecio As Double = 0 : Dim dImp As Double = 0
        Dim sExiste As String = ""
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim oRsLOG As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim oRsLOGADD As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim oDI_COM As EXO_DIAPI.EXO_UDOEntity = Nothing 'Instancia del UDO para Insertar datos
        Dim sFecha As String = "" : Dim sHora As String = "" : Dim iCode As Integer = 0
#End Region
        Try
            ' miramos si existe el fichero y cargamos
            If File.Exists(sArchivo) Then
                Dim excel As New FileInfo(sArchivo)
                pck = New ExcelPackage(excel)
                Dim workbook = pck.Workbook
                Dim worksheet = workbook.Worksheets.First()
                Dim startCell As ExcelCellAddress = worksheet.Dimension.Start
                Dim endCell As ExcelCellAddress = worksheet.Dimension.End
                For iRow As Integer = 2 To endCell.Row
                    sTipo = worksheet.Cells(iRow, 1).Text.ToUpper
                    sICCod = worksheet.Cells(iRow, 2).Text
                    If sICCod = "" Then
                        objGlobal.SBOApp.StatusBar.SetText("El registro Nº" & iRow & " tiene el cod de IC vacío. Se deja de leer el fichero. ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        Exit Sub
                    End If
                    sICName = worksheet.Cells(iRow, 3).Text.Replace(",", "")
                    sAnno = worksheet.Cells(iRow, 10).Text : sAnno = Right(sAnno.ToString, 4)
                    sItemCode = worksheet.Cells(iRow, 4).Text : sItemName = worksheet.Cells(iRow, 5).Text.Replace(",", "")
                    dCantA = worksheet.Cells(iRow, 7).Text.Replace(".", "")
                    dPrecio = worksheet.Cells(iRow, 8).Text.Replace(".", "").Replace("€", "")
                    dImp = worksheet.Cells(iRow, 9).Text.Replace(".", "").Replace("€", "")
                    sDivision = worksheet.Cells(iRow, 6).Text
                    sPeriodo = worksheet.Cells(iRow, 10).Text
                    sComercial = worksheet.Cells(iRow, 11).Text
                    sPais = worksheet.Cells(iRow, 12).Text.ToUpper
                    sProvincia = worksheet.Cells(iRow, 13).Text
                    'Buscamos si existe el IC
                    sExiste = EXO_GLOBALES.GetValueDB(objGlobal.compañia, """OCRD""", """CardCode""", """CardCode""='" & sICCod & "' ")
                    If sExiste = "" Then
                        objGlobal.SBOApp.StatusBar.SetText("El registro Nº" & iRow & " no encuentra el IC " & sICCod & " - Se buscará por el nombre " & sICName, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        sExiste = EXO_GLOBALES.GetValueDB(objGlobal.compañia, """OCRD""", """CardCode""", """CardName""='" & sICName & "' ")
                        sICCod = sExiste
                    End If
                    If sExiste <> "" Then
                        'Buscaremos si existe el artículo y la división
                        sExiste = EXO_GLOBALES.GetValueDB(objGlobal.compañia, """OITM""", """ItemCode""", """ItemCode""='" & sItemCode & "' ")
                        If sExiste <> "" Then
                            sExiste = EXO_GLOBALES.GetValueDB(objGlobal.compañia, """OITB""", """ItmsGrpCod""", """ItmsGrpNam""='" & sDivision & "' ")
                            If sExiste <> "" Then
                                sDivision = sExiste
                                'Buscamos el comercial
                                sExiste = EXO_GLOBALES.GetValueDB(objGlobal.compañia, """OSLP""", """SlpCode""", """SlpName""='" & sComercial & "' ")
                                If sExiste <> "" Then
                                    sComercial = sExiste
                                Else
                                    sComercial = EXO_GLOBALES.GetValueDB(objGlobal.compañia, """OCRD""", """SlpCode""", """CardCode""='" & sICCod & "' ")
                                    objGlobal.SBOApp.StatusBar.SetText("El registro Nº" & iRow & " no encuentra el comercial, se le asigna el del cliente - " & sComercial, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                End If

                                'Buscamos el pais
                                sExiste = EXO_GLOBALES.GetValueDB(objGlobal.compañia, """OCRY""", """Code""", "UPPER(""Name"")='" & sPais & "' ")
                                If sExiste <> "" Then
                                    sPais = sExiste
                                Else
                                    sPais = EXO_GLOBALES.GetValueDB(objGlobal.compañia, """CRD1""", """Country""", """CardCode""='" & sICCod & "' and ""AdresType""='B' ")
                                    objGlobal.SBOApp.StatusBar.SetText("El registro Nº" & iRow & " no encuentra el país, se le asigna el de la dirección de factura del cliente - " & sICCod, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                End If

                                'Buscamos la provincia
                                If sProvincia <> "0" Then
                                    sExiste = EXO_GLOBALES.GetValueDB(objGlobal.compañia, """OCST""", """Code""", """Code""='" & sProvincia & "' ")
                                    If sExiste = "" Then
                                        sProvincia = "0"
                                        objGlobal.SBOApp.StatusBar.SetText("El registro Nº" & iRow & " no encuentra la provincia, se dejará en blanco la provincia.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    End If
                                End If

                                'Ahora empezamos a generar el registro
                                oDI_COM = New EXO_DIAPI.EXO_UDOEntity(objGlobal.refDi.comunes, "EXO_PPTOS") 'UDO 
                                Dim sCode As String = sICCod & "_" & sTipo & "_" & sAnno
                                sSQL = "SELECT * FROM ""@EXO_PPTOS"" WHERE ""Code""='" & sCode & "' "
                                oRs.DoQuery(sSQL)
                                If oRs.RecordCount > 0 Then
                                    oDI_COM.GetByKey(sCode)
                                    CrearCamposLíneas(oDI_COM, sCode, sItemCode, sItemName, sDivision, dCantA, dPrecio, dImp, sPeriodo, sComercial, sPais, sProvincia)
                                    If oDI_COM.UDO_Update = False Then
                                        Throw New Exception("(EXO) - Error al añadir registro Nº" & iRow.ToString & ". " & oDI_COM.GetLastError)
                                    End If
                                Else
                                    oDI_COM.GetNew()
                                    oDI_COM.SetValue("Code") = sCode
                                    oDI_COM.SetValue("U_EXO_ANNO") = sAnno
                                    oDI_COM.SetValue("U_EXO_TIPO") = sTipo.ToUpper.Trim
                                    oDI_COM.SetValue("U_EXO_CARDCODE") = sICCod
                                    oDI_COM.SetValue("U_EXO_CARDNAME") = sICName
                                    CrearCamposLíneas(oDI_COM, sCode, sItemCode, sItemName, sDivision, dCantA, dPrecio, dImp, sPeriodo, sComercial, sPais, sProvincia)
                                    If oDI_COM.UDO_Add = False Then
                                        Throw New Exception("(EXO) - Error al añadir registro Nº" & iRow.ToString & ". " & oDI_COM.GetLastError)
                                    End If
                                End If
                                'SI ha grabado o actualizado nos buscará las líneas que no estén en el LOG para crearlas"
                                'Vemos primero el código max.
                                iCode = objGlobal.refDi.SQL.sqlNumericaB1("SELECT ifnull(MAX(""Code""),0) ""MAXCODE"" FROM ""@EXO_PPTOSLOG"" ")
                                sSQL = "SELECT L.* FROM ""@EXO_PPTOSL"" L LEFT JOIN ""@EXO_PPTOSLOG"" LO ON L.""Code""=LO.""U_EXO_CODE"" and L.""LineId""=LO.""U_EXO_LINEA"" WHERE IFNULL(LO.""U_EXO_CODE"",'')='' and L.""Code""='" & sCode & "' "
                                oRsLOG.DoQuery(sSQL)
                                For i = 0 To oRsLOG.RecordCount - 1
                                    iCode += 1
                                    sFecha = Now.Year.ToString("0000") & Now.Month.ToString("00") & Now.Day.ToString("00") : sHora = Now.Hour.ToString("00") & Now.Minute.ToString("00")
                                    sSQL = "insert into ""@EXO_PPTOSLOG""(""Code"",""U_EXO_CODE"", ""U_EXO_LINEA"", ""U_EXO_LinId"", ""U_EXO_ACCION"",""U_EXO_FECHA"",""U_EXO_HORA"", ""U_EXO_ITEMCODE"","
                                    sSQL &= " ""U_EXO_ITEMNAME"",""U_EXO_DIV"",""U_EXO_CANTA"", ""U_EXO_PRECIO"", ""U_EXO_IMP"", ""U_EXO_PERIODO"", ""U_EXO_PAIS"", ""U_EXO_PROVINCIA"", ""U_EXO_COMERCIAL"") "
                                    sSQL &= " VALUES (" & iCode.ToString & ", '" & oRsLOG.Fields.Item("Code").Value.ToString & "'," & oRsLOG.Fields.Item("LineId").Value.ToString & ", 0, 'C', '" & sFecha & "'," & sHora & ","
                                    sSQL &= "'" & oRsLOG.Fields.Item("U_EXO_ITEMCODE").Value.ToString & "', '" & oRsLOG.Fields.Item("U_EXO_ITEMNAME").Value.ToString & "', '" & oRsLOG.Fields.Item("U_EXO_DIV").Value.ToString & "', "
                                    sSQL &= oRsLOG.Fields.Item("U_EXO_CANTA").Value.ToString & ", " & oRsLOG.Fields.Item("U_EXO_PRECIO").Value.ToString & ", "
                                    Dim dPeriodo As Date = CDate(oRsLOG.Fields.Item("U_EXO_PERIODO").Value.ToString)
                                    sSQL &= oRsLOG.Fields.Item("U_EXO_IMP").Value.ToString & ", '" & dPeriodo.Year.ToString("0000") & dPeriodo.Month.ToString("00") & dPeriodo.Day.ToString("00") & "', "
                                    sSQL &= "'" & oRsLOG.Fields.Item("U_EXO_PAIS").Value.ToString & "', '" & oRsLOG.Fields.Item("U_EXO_PROVINCIA").Value.ToString & "', '" & oRsLOG.Fields.Item("U_EXO_COMERCIAL").Value.ToString & "')"
                                    oRsLOGADD.DoQuery(sSQL)
                                Next
                            Else
                                objGlobal.SBOApp.StatusBar.SetText("El registro Nº" & iRow & " no se cargará, pues no encuentra la división " & sDivision & " del artículo " & sItemCode, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            End If
                        Else
                            objGlobal.SBOApp.StatusBar.SetText("El registro Nº" & iRow & " no se cargará, pues no encuentra el artículo " & sItemCode, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End If
                    Else
                        objGlobal.SBOApp.StatusBar.SetText("El registro Nº" & iRow & " no se cargará, pues no encuentra el IC " & sICCod & " - " & sICName, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End If
                Next
            Else
                objGlobal.SBOApp.StatusBar.SetText("No se ha encontrado el fichero excel a cargar.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oDI_COM, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsLOG, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsLOGADD, Object))
        End Try
    End Sub
    Private Sub CrearCamposLíneas(ByRef oDI_COM As EXO_DIAPI.EXO_UDOEntity, ByVal sCodigo As String, ByVal sItemCode As String, ByVal sItemName As String, ByVal sDiv As String, ByVal dCantA As Double,
                                       ByVal dPrecio As Double, ByVal dImp As Double, ByVal sPeriodo As String, ByVal sComercial As String, ByVal sPais As String, ByVal sProvincia As String)
        Try
            oDI_COM.GetNewChild("EXO_PPTOSL")
            oDI_COM.SetValueChild("U_EXO_ITEMCODE") = sItemCode
            oDI_COM.SetValueChild("U_EXO_ITEMNAME") = sItemName
            oDI_COM.SetValueChild("U_EXO_DIV") = sDiv
            oDI_COM.SetValueChild("U_EXO_CANTA") = dCantA
            oDI_COM.SetValueChild("U_EXO_PRECIO") = dPrecio
            oDI_COM.SetValueChild("U_EXO_IMP") = dImp
            Dim dFecha As Date = CDate(sPeriodo)
            oDI_COM.SetValueChild("U_EXO_PERIODO") = sPeriodo
            oDI_COM.SetValueChild("U_EXO_COMERCIAL") = sComercial
            oDI_COM.SetValueChild("U_EXO_PAIS") = sPais
            oDI_COM.SetValueChild("U_EXO_PROVINCIA") = sProvincia
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
End Class
