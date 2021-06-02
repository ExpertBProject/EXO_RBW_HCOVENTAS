Public Class EXO_GLOBALES
    Public Enum FuenteInformacion
        Visual = 1
        Otros = 2
    End Enum
#Region "Métodos auxiliares"
    Public Shared Function DblNumberToText(ByRef oCompany As SAPbobsCOM.Company, ByVal cValor As Double, ByVal oDestino As FuenteInformacion) As String
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim sSQL As String = ""
        Dim sNumberDouble As String = "0"
        Dim sSeparadorMillarB1 As String = "."
        Dim sSeparadorDecimalB1 As String = ","
        Dim sSeparadorDecimalSO As String = System.Threading.Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator

        DblNumberToText = "0"

        Try
            oRs = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

            sSQL = "SELECT COALESCE(""DecSep"", ',') ""DecSep"", COALESCE(""ThousSep"", '.') ""ThousSep"" " &
                   "FROM ""OADM"" " &
                   "WHERE ""Code"" = 1"

            oRs.DoQuery(sSQL)

            If oRs.RecordCount > 0 Then
                sSeparadorMillarB1 = oRs.Fields.Item("ThousSep").Value.ToString
                sSeparadorDecimalB1 = oRs.Fields.Item("DecSep").Value.ToString
            End If

            If cValor.ToString <> "" Then
                If sSeparadorMillarB1 = "." AndAlso sSeparadorDecimalB1 = "," Then 'Decimales ES
                    sNumberDouble = cValor.ToString
                Else 'Decimales USA
                    sNumberDouble = cValor.ToString.Replace(",", ".")
                End If
            End If

            If oDestino = FuenteInformacion.Visual Then
                If sSeparadorDecimalSO = "," Then
                    DblNumberToText = sNumberDouble
                Else
                    DblNumberToText = sNumberDouble.Replace(".", ",")
                End If
            Else
                DblNumberToText = sNumberDouble.Replace(",", ".")
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function
    Public Shared Function DblTextToNumber(ByRef oCompany As SAPbobsCOM.Company, ByVal sValor As String) As Double
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim sSQL As String = ""
        Dim cValor As Double = 0
        Dim sValorAux As String = "0"
        Dim sSeparadorMillarB1 As String = "."
        Dim sSeparadorDecimalB1 As String = ","
        Dim sSeparadorDecimalSO As String = System.Threading.Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator

        DblTextToNumber = 0

        Try
            oRs = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

            sSQL = "SELECT COALESCE(""DecSep"", ',') ""DecSep"", COALESCE(""ThousSep"", '.') ""ThousSep"" " &
                   "FROM ""OADM"" " &
                   "WHERE ""Code"" = 1"

            oRs.DoQuery(sSQL)

            If oRs.RecordCount > 0 Then
                sSeparadorMillarB1 = oRs.Fields.Item("ThousSep").Value.ToString
                sSeparadorDecimalB1 = oRs.Fields.Item("DecSep").Value.ToString
            End If

            sValorAux = sValor

            If sSeparadorDecimalSO = "," Then
                If sValorAux <> "" Then
                    If Left(sValorAux, 1) = "." Then sValorAux = "0" & sValorAux

                    If sSeparadorMillarB1 = "." AndAlso sSeparadorDecimalB1 = "," Then 'Decimales ES
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

                    If sSeparadorMillarB1 = "." AndAlso sSeparadorDecimalB1 = "," Then 'Decimales ES
                        If sValorAux.IndexOf(",") > 0 AndAlso sValorAux.IndexOf(".") > 0 Then
                            cValor = CDbl(sValorAux.Replace(".", "").Replace(",", "."))
                        ElseIf sValorAux.IndexOf(",") > 0 Then
                            cValor = CDbl(sValorAux.Replace(",", "."))
                        Else
                            cValor = CDbl(sValorAux)
                        End If
                    Else 'Decimales USA
                        If sValorAux.IndexOf(",") > 0 AndAlso sValorAux.IndexOf(".") > 0 Then
                            cValor = CDbl(sValorAux.Replace(",", ""))
                        ElseIf sValorAux.IndexOf(",") > 0 Then
                            cValor = CDbl(sValorAux.Replace(",", "."))
                        Else
                            cValor = CDbl(sValorAux)
                        End If
                    End If
                End If
            End If

            DblTextToNumber = cValor

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function

#End Region
#Region "Datos"
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
    Public Shared Function Añadir_LOG(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef oForm As SAPbouiCOM.Form, ByVal sCode As String) As Boolean
        Añadir_LOG = False
        Dim iCode As Integer = 0 : Dim sFecha As String = "" : Dim sHora As String = ""
        Dim sSQL As String = ""
        Dim oRsLOG As SAPbobsCOM.Recordset = CType(oObjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim oRsLOGADD As SAPbobsCOM.Recordset = CType(oObjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Try
            'SI ha grabado o actualizado nos buscará las líneas que no estén en el LOG para crearlas"
            'Vemos primero el código max.
            iCode = oObjGlobal.refDi.SQL.sqlNumericaB1("SELECT ifnull(MAX(""Code""),0) ""MAXCODE"" FROM ""@EXO_PPTOSLOG"" ")
            sSQL = "SELECT L.* FROM ""@EXO_PPTOSL"" L LEFT JOIN ""@EXO_PPTOSLOG"" LO ON L.""Code""=LO.""U_EXO_CODE"" and L.""LineId""=LO.""U_EXO_LINEA"" WHERE IFNULL(LO.""U_EXO_CODE"",'')='' and L.""Code""='" & sCode & "' "
            oRsLOG.DoQuery(sSQL)
            For i = 0 To oRsLOG.RecordCount - 1
                iCode += 1
                sFecha = Now.Year.ToString("0000") & Now.Month.ToString("00") & Now.Day.ToString("00") : sHora = Now.Hour.ToString("00") & Now.Minute.ToString("00")
                sSQL = "insert into ""@EXO_PPTOSLOG""(""Code"",""U_EXO_CODE"", ""U_EXO_LINEA"", ""U_EXO_LinId"", ""U_EXO_ACCION"",""U_EXO_FECHA"",""U_EXO_HORA"", ""U_EXO_ITEMCODE"","
                sSQL &= " ""U_EXO_ITEMNAME"",""U_EXO_DIV"",""U_EXO_CANTA"",  ""U_EXO_PRECIO"", ""U_EXO_IMP"", ""U_EXO_PERIODO"", ""U_EXO_PAIS"", ""U_EXO_PROVINCIA"", ""U_EXO_COMERCIAL"") "
                sSQL &= " VALUES (" & iCode.ToString & ", '" & oRsLOG.Fields.Item("Code").Value.ToString & "'," & oRsLOG.Fields.Item("LineId").Value.ToString & ", 0, 'C', '" & sFecha & "'," & sHora & ","
                sSQL &= "'" & oRsLOG.Fields.Item("U_EXO_ITEMCODE").Value.ToString & "', '" & oRsLOG.Fields.Item("U_EXO_ITEMNAME").Value.ToString & "', '" & oRsLOG.Fields.Item("U_EXO_DIV").Value.ToString & "', "
                sSQL &= oRsLOG.Fields.Item("U_EXO_CANTA").Value.ToString & ", " & oRsLOG.Fields.Item("U_EXO_PRECIO").Value.ToString & ", "
                Dim dPeriodo As Date = CDate(oRsLOG.Fields.Item("U_EXO_PERIODO").Value.ToString)
                sSQL &= oRsLOG.Fields.Item("U_EXO_IMP").Value.ToString & ", '" & dPeriodo.Year.ToString("0000") & dPeriodo.Month.ToString("00") & dPeriodo.Day.ToString("00") & "', "
                sSQL &= "'" & oRsLOG.Fields.Item("U_EXO_PAIS").Value.ToString & "', '" & oRsLOG.Fields.Item("U_EXO_PROVINCIA").Value.ToString & "', '" & oRsLOG.Fields.Item("U_EXO_COMERCIAL").Value.ToString & "')"
                oRsLOGADD.DoQuery(sSQL)
                oRsLOG.MoveNext()
            Next
            Añadir_LOG = True
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsLOG, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsLOGADD, Object))
        End Try
    End Function
    Public Shared Function Modificar_LOG(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef oForm As SAPbouiCOM.Form, ByVal sCode As String) As Boolean
        Modificar_LOG = False
        Dim iCode As Integer = 0 : Dim sFecha As String = "" : Dim sHora As String = "" : Dim iVersion As Integer = 0 : Dim sAccion As String = "M"
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(oObjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim oRsLOGADD As SAPbobsCOM.Recordset = CType(oObjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

        Try
            'SI ha grabado o actualizado nos buscará las líneas que no estén en el LOG para crearlas"
            'Vemos primero el código max.
            iCode = oObjGlobal.refDi.SQL.sqlNumericaB1("SELECT ifnull(MAX(""Code""),0) ""MAXCODE"" FROM ""@EXO_PPTOSLOG"" ")
            sSQL = "SELECT * FROM ""@EXO_PPTOSL"" WHERE ""Code""='" & sCode & "' Order By ""LineId"" "
            oRs.DoQuery(sSQL)
            For i = 0 To oRs.RecordCount - 1
                iCode += 1
                sSQL = "SELECT MAX(ifnull(""U_EXO_LinId"",0)) ""VERSION"" FROM ""@EXO_PPTOSLOG"" WHERE ""U_EXO_CODE""='" & oRs.Fields.Item("Code").Value.ToString & "' "
                sSQL &= " and ""U_EXO_LINEA""= " & oRs.Fields.Item("LineId").Value.ToString
                iVersion = oObjGlobal.refDi.SQL.sqlNumericaB1(sSQL)
                sSQL = "SELECT ifnull(""U_EXO_CODE"",'C') ""ACCION"" FROM ""@EXO_PPTOSLOG"" WHERE ""U_EXO_CODE""='" & oRs.Fields.Item("Code").Value.ToString & "' "
                sSQL &= " and ""U_EXO_LINEA""= " & oRs.Fields.Item("LineId").Value.ToString
                sAccion = oObjGlobal.refDi.SQL.sqlStringB1(sSQL)

                sFecha = Now.Year.ToString("0000") & Now.Month.ToString("00") & Now.Day.ToString("00") : sHora = Now.Hour.ToString("00") & Now.Minute.ToString("00")

                If iVersion = 0 And sAccion = "" Then
                    sSQL = "insert into ""@EXO_PPTOSLOG""(""Code"",""U_EXO_CODE"", ""U_EXO_LINEA"", ""U_EXO_LinId"", ""U_EXO_ACCION"",""U_EXO_FECHA"",""U_EXO_HORA"", ""U_EXO_ITEMCODE"","
                    sSQL &= " ""U_EXO_ITEMNAME"",""U_EXO_DIV"",""U_EXO_CANTA"", ""U_EXO_PRECIO"", ""U_EXO_IMP"", ""U_EXO_PERIODO"", ""U_EXO_PAIS"", ""U_EXO_PROVINCIA"", ""U_EXO_COMERCIAL"") "
                    sSQL &= " VALUES (" & iCode.ToString & ", '" & oRs.Fields.Item("Code").Value.ToString & "'," & oRs.Fields.Item("LineId").Value.ToString & ", 0, 'C', '" & sFecha & "'," & sHora & ","
                    sSQL &= "'" & oRs.Fields.Item("U_EXO_ITEMCODE").Value.ToString & "', '" & oRs.Fields.Item("U_EXO_ITEMNAME").Value.ToString & "', '" & oRs.Fields.Item("U_EXO_DIV").Value.ToString & "', "
                    sSQL &= oRs.Fields.Item("U_EXO_CANTA").Value.ToString & ", " & oRs.Fields.Item("U_EXO_PRECIO").Value.ToString & ", "
                    Dim dPeriodo As Date = CDate(oRs.Fields.Item("U_EXO_PERIODO").Value.ToString)
                    sSQL &= oRs.Fields.Item("U_EXO_IMP").Value.ToString & ", '" & dPeriodo.Year.ToString("0000") & dPeriodo.Month.ToString("00") & dPeriodo.Day.ToString("00") & "', "
                    sSQL &= "'" & oRs.Fields.Item("U_EXO_PAIS").Value.ToString & "', '" & oRs.Fields.Item("U_EXO_PROVINCIA").Value.ToString & "', '" & oRs.Fields.Item("U_EXO_COMERCIAL").Value.ToString & "')"
                Else
                    iVersion += 1
                    sSQL = "insert into ""@EXO_PPTOSLOG""(""Code"",""U_EXO_CODE"", ""U_EXO_LINEA"", ""U_EXO_LinId"", ""U_EXO_ACCION"",""U_EXO_FECHA"",""U_EXO_HORA"", ""U_EXO_ITEMCODE"","
                    sSQL &= " ""U_EXO_ITEMNAME"",""U_EXO_DIV"",""U_EXO_CANTA"",  ""U_EXO_PRECIO"", ""U_EXO_IMP"", ""U_EXO_PERIODO"", ""U_EXO_PAIS"", ""U_EXO_PROVINCIA"", ""U_EXO_COMERCIAL"") "
                    sSQL &= " VALUES (" & iCode.ToString & ", '" & oRs.Fields.Item("Code").Value.ToString & "'," & oRs.Fields.Item("LineId").Value.ToString & ", " & iVersion.ToString & ", 'M', '" & sFecha & "'," & sHora & ","
                    sSQL &= "'" & oRs.Fields.Item("U_EXO_ITEMCODE").Value.ToString & "', '" & oRs.Fields.Item("U_EXO_ITEMNAME").Value.ToString & "', '" & oRs.Fields.Item("U_EXO_DIV").Value.ToString & "', "
                    sSQL &= oRs.Fields.Item("U_EXO_CANTA").Value.ToString & ", " & oRs.Fields.Item("U_EXO_PRECIO").Value.ToString & ", "
                    Dim dPeriodo As Date = CDate(oRs.Fields.Item("U_EXO_PERIODO").Value.ToString)
                    sSQL &= oRs.Fields.Item("U_EXO_IMP").Value.ToString & ", '" & dPeriodo.Year.ToString("0000") & dPeriodo.Month.ToString("00") & dPeriodo.Day.ToString("00") & "', "
                    sSQL &= "'" & oRs.Fields.Item("U_EXO_PAIS").Value.ToString & "', '" & oRs.Fields.Item("U_EXO_PROVINCIA").Value.ToString & "', '" & oRs.Fields.Item("U_EXO_COMERCIAL").Value.ToString & "')"
                End If
                oRsLOGADD.DoQuery(sSQL)
                oRs.MoveNext()
            Next
            Modificar_LOG = True
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsLOGADD, Object))
        End Try
    End Function
    Public Shared Function Borrar_LOG(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef oForm As SAPbouiCOM.Form) As Boolean
        Borrar_LOG = False
        Dim iCode As Integer = 0 : Dim sFecha As String = "" : Dim sHora As String = "" : Dim iVersion As Integer = 0
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(oObjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim oRsLOGADD As SAPbobsCOM.Recordset = CType(oObjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

        Try
            'Vemos primero el código max.
            iCode = oObjGlobal.refDi.SQL.sqlNumericaB1("SELECT ifnull(MAX(""Code""),0) ""MAXCODE"" FROM ""@EXO_PPTOSLOG"" ")
            Dim sTable_Origen As String = CType(oForm.Items.Item("0_U_E").Specific, SAPbouiCOM.EditText).DataBind.TableName
            Dim sCode As String = oForm.DataSources.DBDataSources.Item(sTable_Origen).GetValue("Code", 0).ToString
            Dim irow As Integer = CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).GetNextSelectedRow
            sSQL = "SELECT * FROM ""@EXO_PPTOSL"" WHERE ""Code""='" & sCode & "' and ""LineId""=" & irow.ToString
            oRs.DoQuery(sSQL)
            For i = 0 To oRs.RecordCount - 1
                iCode += 1
                sSQL = "SELECT MAX(ifnull(""U_EXO_LinId"",0)) ""VERSION"" FROM ""@EXO_PPTOSLOG"" WHERE ""U_EXO_CODE""='" & oRs.Fields.Item("Code").Value.ToString & "' "
                sSQL &= " and ""U_EXO_LINEA""= " & oRs.Fields.Item("LineId").Value.ToString
                iVersion = oObjGlobal.refDi.SQL.sqlNumericaB1(sSQL) : iVersion += 1
                sFecha = Now.Year.ToString("0000") & Now.Month.ToString("00") & Now.Day.ToString("00") : sHora = Now.Hour.ToString("00") & Now.Minute.ToString("00")
                sSQL = "insert into ""@EXO_PPTOSLOG""(""Code"",""U_EXO_CODE"", ""U_EXO_LINEA"", ""U_EXO_LinId"", ""U_EXO_ACCION"",""U_EXO_FECHA"",""U_EXO_HORA"", ""U_EXO_ITEMCODE"","
                sSQL &= " ""U_EXO_ITEMNAME"",""U_EXO_DIV"",""U_EXO_CANTA"",  ""U_EXO_PRECIO"", ""U_EXO_IMP"", ""U_EXO_PERIODO"", ""U_EXO_PAIS"", ""U_EXO_PROVINCIA"",""U_EXO_COMERCIAL"",""U_EXO_ACEPTADO"") "
                sSQL &= " VALUES (" & iCode.ToString & ", '" & oRs.Fields.Item("Code").Value.ToString & "'," & oRs.Fields.Item("LineId").Value.ToString & ", " & iVersion.ToString & ", 'B', '" & sFecha & "'," & sHora & ","
                sSQL &= "'" & oRs.Fields.Item("U_EXO_ITEMCODE").Value.ToString & "', '" & oRs.Fields.Item("U_EXO_ITEMNAME").Value.ToString & "', '" & oRs.Fields.Item("U_EXO_DIV").Value.ToString & "', "
                sSQL &= oRs.Fields.Item("U_EXO_CANTA").Value.ToString & ", " & oRs.Fields.Item("U_EXO_PRECIO").Value.ToString & ", "
                Dim dPeriodo As Date = CDate(oRs.Fields.Item("U_EXO_PERIODO").Value.ToString)
                sSQL &= oRs.Fields.Item("U_EXO_IMP").Value.ToString & ", '" & dPeriodo.Year.ToString("0000") & dPeriodo.Month.ToString("00") & dPeriodo.Day.ToString("00") & "', "
                sSQL &= "'" & oRs.Fields.Item("U_EXO_PAIS").Value.ToString & "', '" & oRs.Fields.Item("U_EXO_PROVINCIA").Value.ToString & "', '" & oRs.Fields.Item("U_EXO_COMERCIAL").Value.ToString & "','N')"
                oRsLOGADD.DoQuery(sSQL)
            Next
            Borrar_LOG = True
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsLOGADD, Object))
        End Try
    End Function
#End Region
End Class
