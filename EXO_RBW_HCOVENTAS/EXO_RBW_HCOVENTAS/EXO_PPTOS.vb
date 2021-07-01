Imports SAPbouiCOM
Public Class EXO_PPTOS
    Inherits EXO_UIAPI.EXO_DLLBase
    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, False, idAddOn)

        If actualizar Then
            cargaCampos()
        End If
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
    Private Sub cargaCampos()
        If objGlobal.refDi.comunes.esAdministrador Then
            Dim oXML As String = ""
            Dim udoObj As EXO_Generales.EXO_UDO = Nothing

            oXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDO_EXO_PPTOS.xml")
            objGlobal.refDi.comunes.LoadBDFromXML(oXML)
            objGlobal.SBOApp.StatusBar.SetText("Validado: UDO_EXO_PPTOS", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            oXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "EXO_PPTOLOG.xml")
            objGlobal.refDi.comunes.LoadBDFromXML(oXML)
            objGlobal.SBOApp.StatusBar.SetText("Validado: EXO_PPTOLOG", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        End If
    End Sub
    Public Overrides Function SBOApp_MenuEvent(infoEvento As MenuEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Try
            If infoEvento.BeforeAction = True Then
                'Aqui podremos guardar la línea que hemos borrado
                'podemos controlar que si se pulsa cancelar elimine los registros de borrado.
                Select Case infoEvento.MenuUID
                    Case "EXO_PPTOS_Remove_Line"
                        'Insertamos línea en el LOG
                        oForm = objGlobal.SBOApp.Forms.ActiveForm
                        EXO_GLOBALES.Borrar_LOG(objGlobal, oForm)
                End Select
            Else
                Select Case infoEvento.MenuUID
                    Case "EXO-MnPTOVTAS"
                        'Cargamos UDO
                        objGlobal.funcionesUI.cargaFormUdoBD("EXO_PPTOS")
                    Case "1282"
                        oForm = objGlobal.SBOApp.Forms.ActiveForm
                        If oForm.Visible = True And oForm.TypeEx = "UDO_FT_EXO_PPTOS" Then
                            CargaCombos(oForm)

                            Dim oItem As SAPbouiCOM.Item
                            oItem = oForm.Items.Item("cmdDupli")
                            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

                            CType(oForm.Items.Item("13_U_E").Specific, SAPbouiCOM.EditText).Active = True
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
    Private Function CargaDatoDivision(ByRef oForm As SAPbouiCOM.Form, ByVal iLinea As Integer) As Boolean
        CargaDatoDivision = False
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sArticulo As String = ""
        Try
            oForm.Freeze(True)
            sArticulo = oForm.DataSources.DBDataSources.Item("@EXO_PPTOSL").GetValue("U_EXO_ITEMCODE", iLinea - 1)
            sSQL = " Select ""ItmsGrpCod"" FROM ""OITM"" WHERE ""ItemCode""='" & sArticulo & "' "
            oRs.DoQuery(sSQL)
            If oRs.RecordCount > 0 Then
                Dim sValor As String = oRs.Fields.Item("ItmsGrpCod").Value.ToString
                oForm.DataSources.DBDataSources.Item("@EXO_PPTOSL").SetValue("U_EXO_DIV", iLinea - 1, oRs.Fields.Item("ItmsGrpCod").Value.ToString)
                'CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_3").Cells.Item(iLinea).Specific, SAPbouiCOM.ComboBox).Select(sValor, BoSearchKey.psk_ByValue)
            End If

            CargaDatoDivision = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            oForm.Freeze(False)
            'Throw exCOM
        Catch ex As Exception
            oForm.Freeze(False)
            ' Throw ex
        Finally
            oForm.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function
    Private Function CargaDatoPrecio(ByRef oForm As SAPbouiCOM.Form, ByVal iLinea As Integer) As Boolean

        CargaDatoPrecio = False
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sArticulo As String = ""
        Try
            oForm.Freeze(True)
            sArticulo = oForm.DataSources.DBDataSources.Item("@EXO_PPTOSL").GetValue("U_EXO_ITEMCODE", iLinea - 1)
            sSQL = " Select ""AvgPrice"" FROM ""OITW"" WHERE ""ItemCode""='" & sArticulo & "' and ""WhsCode""='RW1' "
            oRs.DoQuery(sSQL)
            If oRs.RecordCount > 0 Then
                oForm.DataSources.DBDataSources.Item("@EXO_PPTOSL").SetValue("U_EXO_PRECIO", iLinea - 1, oRs.Fields.Item("AvgPrice").Value.ToString)
            Else
                oForm.DataSources.DBDataSources.Item("@EXO_PPTOSL").SetValue("U_EXO_PRECIO", iLinea - 1, "0")
            End If
            CargaDatoPrecio = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oForm.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function
    Private Function CargaCombos(ByRef oForm As SAPbouiCOM.Form) As Boolean

        CargaCombos = False
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Try
            oForm.Freeze(True)

            'División
            sSQL = " Select ""ItmsGrpCod"",""ItmsGrpNam"" FROM ""OITB"" Order by ""ItmsGrpNam"" "
            oRs.DoQuery(sSQL)
            If oRs.RecordCount > 0 Then
                objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_3").ValidValues, sSQL)
            End If

            'Pais
            sSQL = " Select ""Code"",""Name"" FROM ""OCRY"" Order by ""Name"" "
            oRs.DoQuery(sSQL)
            If oRs.RecordCount > 0 Then
                objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_9").ValidValues, sSQL)
            End If

            'PROVINCIA
            sSQL = " SELECT '0' ""Code"", ' ' ""Name"" FROM DUMMY "
            sSQL &= " UNION ALL "
            sSQL = " Select ""Code"",""Name"" FROM ""OCST"" Order by ""Name"" "
            oRs.DoQuery(sSQL)
            If oRs.RecordCount > 0 Then
                objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_10").ValidValues, sSQL)
            End If

            'Comercial
            sSQL = " Select ""SlpCode"",""SlpName"" FROM ""OSLP"" Order by ""SlpName"" "
            oRs.DoQuery(sSQL)
            If oRs.RecordCount > 0 Then
                objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_11").ValidValues, sSQL)
            End If
            CargaCombos = True

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
                        Case "UDO_FT_EXO_PPTOS"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                    If EventHandler_ITEM_PRESSED_After(objGlobal, infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                                    If EventHandler_VALIDATE_After(objGlobal, infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE

                            End Select
                    End Select
                ElseIf infoEvento.BeforeAction = True Then
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_EXO_PPTOS"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                    If EventHandler_ITEM_PRESSED_Before(objGlobal, infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK

                            End Select

                    End Select
                End If
            Else
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_EXO_PPTOS"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE
                                    If EventHandler_FORM_VISIBLE(objGlobal, infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                    If EventHandler_Choose_FromList_After(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                            End Select
                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_EXO_PPTOS"
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
    Private Function EventHandler_ITEM_PRESSED_After(ByRef objGlobal As EXO_UIAPI.EXO_UIAPI, ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sSQL As String = ""
        EventHandler_ITEM_PRESSED_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            Dim sTable_Origen As String = CType(oForm.Items.Item("0_U_E").Specific, SAPbouiCOM.EditText).DataBind.TableName
            Dim sCode As String = oForm.DataSources.DBDataSources.Item(sTable_Origen).GetValue("Code", 0).ToString

            If pVal.ItemUID = "1" Then
                ' al pulsar aceptar, tendremos que aceptar el borrado de las líneas marcadas como pdte de aceptar
                sSQL = "UPDATE ""@EXO_PPTOSLOG"" SET ""U_EXO_ACEPTADO""='Y' WHERE ""U_EXO_CODE""='" & sCode & "' and ""U_EXO_ACEPTADO""='N' and ""U_EXO_ACCION""='B' "
                objGlobal.refDi.SQL.sqlUpdB1(sSQL)
            End If
            EventHandler_ITEM_PRESSED_After = True


        Catch exCOM As System.Runtime.InteropServices.COMException
            oForm.Freeze(False)
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            oForm.Freeze(False)
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            oForm.Freeze(False)

        End Try
    End Function
    Private Function EventHandler_ITEM_PRESSED_Before(ByRef objGlobal As EXO_UIAPI.EXO_UIAPI, ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sSQL As String = "" : Dim sPais As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        EventHandler_ITEM_PRESSED_Before = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            Dim sTable_Origen As String = CType(oForm.Items.Item("0_U_E").Specific, SAPbouiCOM.EditText).DataBind.TableName
            Dim sCode As String = oForm.DataSources.DBDataSources.Item(sTable_Origen).GetValue("Code", 0).ToString

            If pVal.ItemUID = "2" Then
                'Al pulsar cancelar, tendremos que borrar las líneas creadas como borradas en el mismo día y pdte de aceptar.
                sSQL = "DELETE FROM ""@EXO_PPTOSLOG"" WHERE ""U_EXO_CODE""='" & sCode & "' and ""U_EXO_ACEPTADO""='N' and ""U_EXO_ACCION""='B'"
                objGlobal.refDi.SQL.sqlUpdB1(sSQL)
            ElseIf pVal.ItemUID = "cmdDupli" Then

                Dim sCardCode As String = oForm.DataSources.DBDataSources.Item(sTable_Origen).GetValue("U_EXO_CARDCODE", 0).ToString
                Dim sCardName As String = oForm.DataSources.DBDataSources.Item(sTable_Origen).GetValue("U_EXO_CARDNAME", 0).ToString
                oForm.Mode = BoFormMode.fm_ADD_MODE
                oForm.DataSources.DBDataSources.Item(sTable_Origen).SetValue("U_EXO_CARDCODE", 0, sCardCode)
                oForm.DataSources.DBDataSources.Item(sTable_Origen).SetValue("U_EXO_CARDNAME", 0, sCardName)
                'rellenamos la matrix
                oForm.Freeze(True)
                sSQL = "SELECT * FROM ""@EXO_PPTOSL"" WHERE ""Code""='" & sCode & "' "
                oRs.DoQuery(sSQL)
                If oRs.RecordCount > 0 Then
                    CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).FlushToDataSource()
                    For i = 0 To oRs.RecordCount - 1
                        oForm.DataSources.DBDataSources.Item("@EXO_PPTOSL").InsertRecord(i)
                        oForm.DataSources.DBDataSources.Item("@EXO_PPTOSL").Offset = i

                        oForm.DataSources.DBDataSources.Item("@EXO_PPTOSL").SetValue("U_EXO_ITEMCODE", oForm.DataSources.DBDataSources.Item("@EXO_PPTOSL").Offset, oRs.Fields.Item("U_EXO_ITEMCODE").Value.ToString)
                        oForm.DataSources.DBDataSources.Item("@EXO_PPTOSL").SetValue("U_EXO_ITEMNAME", oForm.DataSources.DBDataSources.Item("@EXO_PPTOSL").Offset, oRs.Fields.Item("U_EXO_ITEMNAME").Value.ToString)
                        oForm.DataSources.DBDataSources.Item("@EXO_PPTOSL").SetValue("U_EXO_DIV", oForm.DataSources.DBDataSources.Item("@EXO_PPTOSL").Offset, oRs.Fields.Item("U_EXO_DIV").Value.ToString)
                        oForm.DataSources.DBDataSources.Item("@EXO_PPTOSL").SetValue("U_EXO_CANTA", oForm.DataSources.DBDataSources.Item("@EXO_PPTOSL").Offset, oRs.Fields.Item("U_EXO_CANTA").Value.ToString)
                        oForm.DataSources.DBDataSources.Item("@EXO_PPTOSL").SetValue("U_EXO_CANTB", oForm.DataSources.DBDataSources.Item("@EXO_PPTOSL").Offset, oRs.Fields.Item("U_EXO_CANTB").Value.ToString)
                        oForm.DataSources.DBDataSources.Item("@EXO_PPTOSL").SetValue("U_EXO_PRECIO", oForm.DataSources.DBDataSources.Item("@EXO_PPTOSL").Offset, oRs.Fields.Item("U_EXO_PRECIO").Value.ToString)
                        oForm.DataSources.DBDataSources.Item("@EXO_PPTOSL").SetValue("U_EXO_IMP", oForm.DataSources.DBDataSources.Item("@EXO_PPTOSL").Offset, oRs.Fields.Item("U_EXO_IMP").Value.ToString)
                        Dim dFecha As Date = CDate(oRs.Fields.Item("U_EXO_PERIODO").Value.ToString)
                        Dim sFecha As String = dFecha.Year.ToString("0000") & dFecha.Month.ToString("00") & dFecha.Day.ToString("00")
                        oForm.DataSources.DBDataSources.Item("@EXO_PPTOSL").SetValue("U_EXO_PERIODO", oForm.DataSources.DBDataSources.Item("@EXO_PPTOSL").Offset, sFecha)
                        oForm.DataSources.DBDataSources.Item("@EXO_PPTOSL").SetValue("U_EXO_PAIS", oForm.DataSources.DBDataSources.Item("@EXO_PPTOSL").Offset, oRs.Fields.Item("U_EXO_PAIS").Value.ToString)
                        oForm.DataSources.DBDataSources.Item("@EXO_PPTOSL").SetValue("U_EXO_PROVINCIA", oForm.DataSources.DBDataSources.Item("@EXO_PPTOSL").Offset, oRs.Fields.Item("U_EXO_PROVINCIA").Value.ToString)

                    Next
                    oForm.DataSources.DBDataSources.Item("@EXO_PPTOSL").RemoveRecord(oRs.RecordCount)
                End If
                CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).LoadFromDataSource()
                oForm.Freeze(False)
                CType(oForm.Items.Item("13_U_E").Specific, SAPbouiCOM.EditText).Active = True
            End If

            EventHandler_ITEM_PRESSED_Before = True


        Catch exCOM As System.Runtime.InteropServices.COMException
            oForm.Freeze(False)
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            oForm.Freeze(False)
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            oForm.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function
    Private Function EventHandler_VALIDATE_After(ByRef objGlobal As EXO_UIAPI.EXO_UIAPI, ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sTable_Origen As String = ""
        Dim sAnno As String = "" : Dim sIC As String = "" : Dim sTipo As String = ""
        EventHandler_VALIDATE_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            sTable_Origen = CType(oForm.Items.Item("0_U_E").Specific, SAPbouiCOM.EditText).DataBind.TableName
            If pVal.ItemUID = "13_U_E" Or pVal.ItemUID = "14_U_E" And oForm.Mode = BoFormMode.fm_ADD_MODE Then
                sAnno = oForm.DataSources.DBDataSources.Item(sTable_Origen).GetValue("U_EXO_ANNO", 0).ToString
                sIC = oForm.DataSources.DBDataSources.Item(sTable_Origen).GetValue("U_EXO_CARDCODE", 0).ToString
                sTipo = oForm.DataSources.DBDataSources.Item(sTable_Origen).GetValue("U_EXO_TIPO", 0).ToString
                oForm.DataSources.DBDataSources.Item(sTable_Origen).SetValue("Code", 0, sIC & "_" & sTipo & "_" & sAnno)
            ElseIf pVal.ItemUID = "0_U_G" And (pVal.ColUID = "C_0_4" Or pVal.ColUID = "C_0_6") Then
                'Si cambia cantidad A o precio debe multiplicar y ponerlo en los importes
                Dim dCantidad As Double = 0 : Dim dPrecio As Double = 0 : Dim dImporte As Double = 0
                dCantidad = EXO_GLOBALES.DblTextToNumber(objGlobal.compañia, CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_4").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value)
                dPrecio = EXO_GLOBALES.DblTextToNumber(objGlobal.compañia, CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_6").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value)
                dImporte = dCantidad * dPrecio
                CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_7").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value = EXO_GLOBALES.DblNumberToText(objGlobal.compañia, dImporte, 2)
                ' oForm.DataSources.DBDataSources.Item("@EXO_PPTOSL").SetValue("U_EXO_IMP", pVal.Row - 1, EXO_GLOBALES.DblNumberToText(objGlobal.compañia, dImporte, 2))
                CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_7").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Active = True
                CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_6").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Active = True
            End If

            EventHandler_VALIDATE_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            oForm.Freeze(False)
        End Try
    End Function
    Private Function EventHandler_FORM_VISIBLE(ByRef objGlobal As EXO_UIAPI.EXO_UIAPI, ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

        EventHandler_FORM_VISIBLE = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            If oForm.Visible = True Then
                oForm.Freeze(True)
                CargaCombos(oForm)

                Dim oItem As SAPbouiCOM.Item
                oItem = oForm.Items.Item("cmdDupli")
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

                CType(oForm.Items.Item("13_U_E").Specific, SAPbouiCOM.EditText).Active = True
            End If

            EventHandler_FORM_VISIBLE = True


        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            oForm.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))

        End Try
    End Function
    Private Function EventHandler_Choose_FromList_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sTable_Origen As String = ""
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        EventHandler_Choose_FromList_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent = CType(pVal, SAPbouiCOM.IChooseFromListEvent)
            Dim oDataTable As DataTable
            sTable_Origen = CType(oForm.Items.Item("0_U_E").Specific, SAPbouiCOM.EditText).DataBind.TableName
            oDataTable = oCFLEvento.SelectedObjects
            If pVal.ItemUID = "14_U_E" AndAlso pVal.ChooseFromListUID = "CFL_IC" AndAlso (oForm.Mode = BoFormMode.fm_ADD_MODE Or oForm.Mode = BoFormMode.fm_UPDATE_MODE) Then
                If oDataTable IsNot Nothing Then
                    Try
                        oForm.DataSources.DBDataSources.Item(sTable_Origen).SetValue("U_EXO_CARDNAME", 0, oDataTable.GetValue("CardName", 0).ToString)
                    Catch ex As Exception
                        oForm.DataSources.DBDataSources.Item(sTable_Origen).SetValue("U_EXO_CARDNAME", 0, oDataTable.GetValue("CardName", 0).ToString)
                    End Try
                End If
            ElseIf pVal.ItemUID = "0_U_G" AndAlso pVal.ChooseFromListUID = "CFL_ART" Then
                If oDataTable IsNot Nothing Then
                    Try
                        oForm.DataSources.DBDataSources.Item("@EXO_PPTOSL").SetValue("U_EXO_ITEMNAME", pVal.Row - 1, oDataTable.GetValue("ItemName", 0).ToString)
                        CargaDatoDivision(oForm, pVal.Row)
                        'CargaDatoPrecio(oForm, pVal.Row)
                        CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_3").Cells.Item(pVal.Row).Specific, SAPbouiCOM.ComboBox).Item.Enabled = True
                        'CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_3").Cells.Item(pVal.Row).Specific, SAPbouiCOM.ComboBox).Active = True
                    Catch ex As Exception
                        oForm.DataSources.DBDataSources.Item("@EXO_PPTOSL").SetValue("U_EXO_ITEMNAME", pVal.Row - 1, oDataTable.GetValue("ItemName", 0).ToString)
                        CargaDatoDivision(oForm, pVal.Row)
                    End Try
                End If
            ElseIf pVal.ItemUID = "0_U_G" AndAlso pVal.ChooseFromListUID = "CFLART2" Then
                If oDataTable IsNot Nothing Then
                    Try
                        oForm.DataSources.DBDataSources.Item("@EXO_PPTOSL").SetValue("U_EXO_ITEMCODE", pVal.Row - 1, oDataTable.GetValue("ItemCode", 0).ToString)
                        CargaDatoDivision(oForm, pVal.Row)
                        'CargaDatoPrecio(oForm, pVal.Row)
                        CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_3").Cells.Item(pVal.Row).Specific, SAPbouiCOM.ComboBox).Active = True
                    Catch ex As Exception
                        oForm.DataSources.DBDataSources.Item("@EXO_PPTOSL").SetValue("U_EXO_ITEMCODE", pVal.Row - 1, oDataTable.GetValue("ItemCode", 0).ToString)
                        CargaDatoDivision(oForm, pVal.Row)
                    End Try
                End If
            End If

            EventHandler_Choose_FromList_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            oForm.Freeze(False)
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            oForm.Freeze(False)
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function

    Public Overrides Function SBOApp_FormDataEvent(infoEvento As BusinessObjectInfo) As Boolean
#Region "Variables"
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim resultado As Boolean = True
        Dim sTable_Origen As String = ""
        Dim sCode As String = ""
#End Region

        Try
            oForm = objGlobal.SBOApp.Forms.Item(infoEvento.FormUID)
            sTable_Origen = CType(oForm.Items.Item("0_U_E").Specific, SAPbouiCOM.EditText).DataBind.TableName
            sCode = oForm.DataSources.DBDataSources.Item(sTable_Origen).GetValue("Code", 0).ToString
            If infoEvento.BeforeAction = False Then
                Select Case infoEvento.FormTypeEx
                    Case "UDO_FT_EXO_PPTOS"
                        Select Case infoEvento.EventType
                            Case BoEventTypes.et_FORM_DATA_ADD
                                If sCode <> "" Then
                                    If EXO_GLOBALES.Añadir_LOG(objGlobal, oForm, sCode) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                                End If
                            Case BoEventTypes.et_FORM_DATA_UPDATE
                                If sCode <> "" Then
                                    If EXO_GLOBALES.Modificar_LOG(objGlobal, oForm, sCode) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                                End If
                        End Select
                End Select
            Else
                Select Case infoEvento.FormTypeEx
                    Case "UDO_FT_EXO_PPTOS"
                        Select Case infoEvento.EventType

                        End Select
                End Select
            End If
        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
        Return resultado
    End Function
End Class
