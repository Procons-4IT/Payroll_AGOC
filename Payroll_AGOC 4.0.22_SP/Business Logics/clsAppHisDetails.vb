Imports System.IO
Public Class clsAppHisDetails
    Inherits clsBase
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox As SAPbouiCOM.ComboBox
    Private oCombo As SAPbouiCOM.ComboBoxColumn
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oGrid As SAPbouiCOM.Grid
    Private dtDocumentList As SAPbouiCOM.DataTable
    Private dtHistoryList As SAPbouiCOM.DataTable
    Private InvForConsumedItems As Integer
    Private blnFlag As Boolean = False

    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub
    Public Sub LoadForm(ByVal oForm As SAPbouiCOM.Form, ByVal DocNo As String)
        Try
            oForm = oApplication.Utilities.LoadForm(xml_Pay_AppHisDetails, frm_Pay_AppHisDetails)
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Items.Item("3").Visible = True
            oForm.DataSources.UserDataSources.Add("fileName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oEditText = oForm.Items.Item("6").Specific
            oEditText.DataBind.SetBound(True, "", "fileName")
            LoadViewHistory(oForm, DocNo)
        Catch ex As Exception
            oForm.Freeze(False)
            Throw ex
        End Try
    End Sub
    Public Sub LoadForm1(ByVal oForm As SAPbouiCOM.Form, ByVal Month As String, ByVal Year As String, ByVal Comp As String)
        Try
            oForm = oApplication.Utilities.LoadForm(xml_Pay_AppHisDetails, frm_Pay_AppHisDetails)
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Items.Item("3").Visible = True
            oForm.DataSources.UserDataSources.Add("fileName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oEditText = oForm.Items.Item("6").Specific
            oEditText.DataBind.SetBound(True, "", "fileName")
            LoadViewHistoryMYC(oForm, Month, Year, Comp)
        Catch ex As Exception
            oForm.Freeze(False)
            Throw ex
        End Try
    End Sub
    Public Sub LoadViewHistoryMYC(ByVal aForm As SAPbouiCOM.Form, ByVal Month As String, ByVal Year As String, ByVal Comp As String)
        Try
            aForm.Freeze(True)
            Dim sQuery As String
            oGrid = aForm.Items.Item("3").Specific
            ' sQuery = " Select T0.DocEntry,T0.U_Z_DocEntry,T0.U_Z_DocType,T0.U_Z_ApproveBy,T0.U_Z_EmpId,T0.U_Z_EmpName,T0.CreateDate as 'Action Date' ,T0.CreateTime as 'Action Time',T0.UpdateDate,T0.UpdateTime,T0.U_Z_AppStatus,"
            sQuery = " Select T0.DocEntry,T0.U_Z_DocEntry,T0.U_Z_DocType,T0.U_Z_ApproveBy,T0.U_Z_EmpId,T0.U_Z_EmpName,T0.CreateDate as 'Action Date' ,T0.CreateTime as 'Action Time',T0.U_Z_AppStatus,"
            sQuery += " T0.U_Z_Remarks,T0.U_Z_MONTH,T0.U_Z_YEAR,T0.U_Z_CompNo,T0.U_Z_Attachment,T1.U_Z_FileName,T0.U_Z_NextApprover 'Next Approver' From [@Z_PAY_APHIS] T0 Right outer Join [@Z_PAY_APPROVAL] T1 on T1.Code=T0.U_Z_DocEntry"
            sQuery += " Where T0.U_Z_DocType = 'R'"
            sQuery += " And T0.U_Z_MONTH = " + Month + " and T0.U_Z_YEAR=" + Year + " and T0.U_Z_CompNo='" & Comp & "'"
            oGrid.DataTable.ExecuteQuery(sQuery)

            Dim oTest As SAPbobsCOM.Recordset
            oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTest.DoQuery("Select * from [@Z_PAY_Approval] where U_Z_Month=" & Month & " and U_Z_Year=" & Year & " and U_Z_CompNo='" & Comp & "' and Convert(Varchar,U_Z_FileName) <> '' order by Convert(Numeric,Code) Desc")
            If oTest.RecordCount > 0 Then
                oApplication.Utilities.setEdittextvalue(aForm, "6", oTest.Fields.Item("U_Z_FileName").Value)
            End If


            formatHistory(aForm)
            assignMatrixLineno(oGrid, aForm)
            aForm.Freeze(False)
        Catch ex As Exception
            aForm.Freeze(False)
            Throw ex
        End Try
    End Sub

    Public Sub assignMatrixLineno(ByVal aGrid As SAPbouiCOM.Grid, ByVal aform As SAPbouiCOM.Form)
        aform.Freeze(True)
        For intNo As Integer = 0 To aGrid.DataTable.Rows.Count - 1
            aGrid.RowHeaders.SetText(intNo, intNo + 1)
        Next
        aGrid.Columns.Item("RowsHeader").TitleObject.Caption = "#"
        aform.Freeze(False)
    End Sub
    Public Sub LoadViewHistory(ByVal aForm As SAPbouiCOM.Form, ByVal strDocEntry As String)
        Try
            aForm.Freeze(True)
            Dim sQuery As String
            oGrid = aForm.Items.Item("3").Specific

            'sQuery = " Select T0.DocEntry,T0.U_Z_DocEntry,T0.U_Z_DocType,T0.U_Z_EmpId,T0.U_Z_EmpName,T0.U_Z_ApproveBy,T0.CreateDate ,T0.CreateTime,T0.UpdateDate,T0.UpdateTime,T0.U_Z_AppStatus,"
            'sQuery += " T0.U_Z_Remarks,T0.U_Z_MONTH,T0.U_Z_YEAR,T0.U_Z_CompNo,T0.U_Z_Attachment,T1.U_Z_FileName,T0.U_Z_NextApprover 'Next Approver' From [@Z_PAY_APHIS] T0 Right outer Join [@Z_PAY_APPROVAL] T1 on T1.Code=T0.U_Z_DocEntry"
            'sQuery += " Where T0.U_Z_DocType = 'R'"
            'sQuery += " And T0.U_Z_MONTH = " + Month() + " and T0.U_Z_YEAR=" + Year() + " and T0.U_Z_CompNo='" & Comp & "'"


            ' sQuery = " Select DocEntry,U_Z_DocEntry,U_Z_DocType,U_Z_EmpId,U_Z_EmpName,U_Z_ApproveBy,CreateDate ,CreateTime,UpdateDate,UpdateTime,U_Z_AppStatus,U_Z_Remarks,U_Z_MONTH,U_Z_YEAR,U_Z_CompNo,U_Z_Attachment,U_Z_FileName,U_Z_NextApprover 'Next Approver' From [@Z_PAY_APHIS] "
            ' sQuery += " Where U_Z_DocType = 'R'"
            'sQuery = " Select T0.DocEntry,T0.U_Z_DocEntry,T0.U_Z_DocType,T0.U_Z_ApproveBy,T0.U_Z_EmpId,T0.U_Z_EmpName,T0.CreateDate as 'Action Date' ,T0.CreateTime as 'Action Time',T0.UpdateDate,T0.UpdateTime,T0.U_Z_AppStatus,"
            sQuery = " Select T0.DocEntry,T0.U_Z_DocEntry,T0.U_Z_DocType,T0.U_Z_ApproveBy,T0.U_Z_EmpId,T0.U_Z_EmpName,T0.CreateDate as 'Action Date' ,T0.CreateTime as 'Action Time',T0.U_Z_AppStatus,"
            sQuery += " T0.U_Z_Remarks,T0.U_Z_MONTH,T0.U_Z_YEAR,T0.U_Z_CompNo,T0.U_Z_Attachment,T1.U_Z_FileName,T0.U_Z_NextApprover 'Next Approver' From [@Z_PAY_APHIS] T0 Right outer Join [@Z_PAY_APPROVAL] T1 on T1.Code=T0.U_Z_DocEntry"
            sQuery += " Where T0.U_Z_DocType = 'R'"
            sQuery += " And U_Z_DocEntry = '" + strDocEntry + "'"
            oGrid.DataTable.ExecuteQuery(sQuery)
            formatHistory(aForm)
            assignMatrixLineno(oGrid, aForm)
            aForm.Freeze(False)
        Catch ex As Exception
            aForm.Freeze(False)
            Throw ex
        End Try
    End Sub
    Private Sub formatHistory(ByVal aForm As SAPbouiCOM.Form)
        Try
            aForm.Freeze(True)
            Dim oGrid As SAPbouiCOM.Grid
            Dim oComboBox As SAPbouiCOM.ComboBox
            Dim oGridCombo As SAPbouiCOM.ComboBoxColumn
            Dim oEditTextColumn As SAPbouiCOM.EditTextColumn
            oGrid = aForm.Items.Item("3").Specific
            oGrid.Columns.Item("DocEntry").Visible = False
            oGrid.Columns.Item("U_Z_DocEntry").TitleObject.Caption = "Reference No."
            oGrid.Columns.Item("U_Z_DocEntry").Visible = False
            oGrid.Columns.Item("U_Z_DocType").Visible = False
            oGrid.Columns.Item("U_Z_EmpId").TitleObject.Caption = "Employee ID"
            oEditTextColumn = oGrid.Columns.Item("U_Z_EmpId")
            oEditTextColumn.LinkedObjectType = "171"
            oGrid.Columns.Item("U_Z_EmpId").Visible = False
            oGrid.Columns.Item("U_Z_EmpName").TitleObject.Caption = "Approver  Name"
            oGrid.Columns.Item("U_Z_EmpName").Visible = True
            oGrid.Columns.Item("U_Z_CompNo").TitleObject.Caption = "Company Code"
            oGrid.Columns.Item("U_Z_YEAR").TitleObject.Caption = "Payroll Year"
            oGrid.Columns.Item("U_Z_MONTH").TitleObject.Caption = "Payroll Month"
            'oGrid.Columns.Item("U_Z_MONTH").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
            'oGridCombo = oGrid.Columns.Item("U_Z_MONTH")
            'For intRow As Integer = 1 To 12
            '    oGridCombo.ValidValues.Add(intRow, MonthName(intRow))
            'Next
            'oGridCombo.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description
            'oGridCombo.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_both
            oGrid.Columns.Item("U_Z_ApproveBy").TitleObject.Caption = "Approved By"
            oGrid.Columns.Item("U_Z_AppStatus").TitleObject.Caption = "Approved Status"
            oGrid.Columns.Item("U_Z_AppStatus").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
            oGridCombo = oGrid.Columns.Item("U_Z_AppStatus")
            oGridCombo.ValidValues.Add("A", "Approved")
            oGridCombo.ValidValues.Add("R", "Rejected")
            oGridCombo.ValidValues.Add("P", "Pending")
            oGridCombo.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description
            oGridCombo.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_both
            oGrid.Columns.Item("U_Z_Remarks").TitleObject.Caption = "Remarks"
            oGrid.Columns.Item("U_Z_Attachment").TitleObject.Caption = "Attachment"
            oGrid.Columns.Item("U_Z_Attachment").Editable = False
            oEditTextColumn = oGrid.Columns.Item("U_Z_Attachment")
            oEditTextColumn.LinkedObjectType = "Z_PAY_APHIS"


            oGrid.Columns.Item("U_Z_FileName").TitleObject.Caption = "Worksheet Attachment"
            oGrid.Columns.Item("U_Z_FileName").Editable = False
            oEditTextColumn = oGrid.Columns.Item("U_Z_FileName")
            oEditTextColumn.LinkedObjectType = "Z_PAY_APHIS"
            oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            oGrid.AutoResizeColumns()
            aForm.Freeze(False)
            For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                If oGrid.DataTable.GetValue("U_Z_ApproveBy", intRow) = oApplication.Company.UserName Then
                    oGrid.Columns.Item("RowsHeader").Click(intRow, False, False)
                    aForm.Freeze(False)
                    Exit Sub
                End If
            Next
            aForm.Items.Item("8").Enabled = True
            aForm.Items.Item("10").Enabled = True
        Catch ex As Exception
            aForm.Freeze(False)
            Throw ex
        End Try
    End Sub
    Private Sub LoadFiles(ByVal aform As SAPbouiCOM.Form, ByVal GridId As String)
        oGrid = aform.Items.Item(GridId).Specific
        For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
            If oGrid.Rows.IsSelected(intRow) Then
                Dim strFilename, strFilePath As String
                strFilename = oGrid.DataTable.GetValue("U_Z_Attachment", intRow)
                Dim Filename As String = Path.GetFileName(strFilename)
                strFilePath = oGrid.DataTable.GetValue("U_Z_Attachment", intRow)

                If File.Exists(strFilePath) = False Then
                    Dim oRec As SAPbobsCOM.Recordset
                    oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Dim strQry = "Select ""AttachPath"" From OADP"
                    oRec.DoQuery(strQry)
                    strFilePath = oRec.Fields.Item(0).Value

                    If Filename = "" Then
                        strFilePath = strFilePath
                    Else
                        strFilePath = strFilePath & Filename
                    End If
                    If File.Exists(strFilePath) = False Then
                        oApplication.Utilities.Message("File does not exists ", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Exit Sub
                    End If
                    strFilename = strFilePath
                Else
                    strFilename = strFilePath
                End If

                Dim x As System.Diagnostics.ProcessStartInfo
                x = New System.Diagnostics.ProcessStartInfo
                x.UseShellExecute = True
                x.FileName = strFilename
                System.Diagnostics.Process.Start(x)
                x = Nothing
                Exit Sub
            End If
        Next
        oApplication.Utilities.Message("No file has been selected...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    End Sub

    Private Sub LoadFiles_New(ByVal aform As SAPbouiCOM.Form)



        Dim strFilename, strFilePath As String
        strFilename = oApplication.Utilities.getEdittextvalue(aform, "6")
        Dim Filename As String = Path.GetFileName(strFilename)
        strFilePath = oApplication.Utilities.getEdittextvalue(aform, "6")
        If strFilePath = "" Then
            Exit Sub
        End If
        If File.Exists(strFilePath) = False Then

            Dim oRec As SAPbobsCOM.Recordset
            oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim strQry = "Select ""AttachPath"" From OADP"
            oRec.DoQuery(strQry)
            strFilePath = oRec.Fields.Item(0).Value

            If Filename = "" Then
                strFilePath = strFilePath
            Else
                strFilePath = strFilePath & Filename
            End If
            If File.Exists(strFilePath) = False Then
                oApplication.Utilities.Message("File does not exists ", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            strFilename = strFilePath
        Else
            strFilename = strFilePath
        End If
        Dim x As System.Diagnostics.ProcessStartInfo
        x = New System.Diagnostics.ProcessStartInfo
        x.UseShellExecute = True
        x.FileName = strFilename
        System.Diagnostics.Process.Start(x)
        x = Nothing
        Exit Sub

    End Sub
    Private Sub LoadFiles1(ByVal aform As SAPbouiCOM.Form, ByVal GridId As String)
        oGrid = aform.Items.Item(GridId).Specific
        For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
            If oGrid.Rows.IsSelected(intRow) Then
                Dim strFilename, strFilePath As String
                strFilename = oGrid.DataTable.GetValue("U_Z_FileName", intRow)
                Dim Filename As String = Path.GetFileName(strFilename)
                strFilePath = oGrid.DataTable.GetValue("U_Z_FileName", intRow)

                If File.Exists(strFilePath) = False Then
                    Dim oRec As SAPbobsCOM.Recordset
                    oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Dim strQry = "Select ""AttachPath"" From OADP"
                    oRec.DoQuery(strQry)
                    strFilePath = oRec.Fields.Item(0).Value

                    If Filename = "" Then
                        strFilePath = strFilePath
                    Else
                        strFilePath = strFilePath & Filename
                    End If
                    If File.Exists(strFilePath) = False Then
                        oApplication.Utilities.Message("File does not exists ", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Exit Sub
                    End If
                    strFilename = strFilePath
                Else
                    strFilename = strFilePath
                End If

                Dim x As System.Diagnostics.ProcessStartInfo
                x = New System.Diagnostics.ProcessStartInfo
                x.UseShellExecute = True
                x.FileName = strFilename
                System.Diagnostics.Process.Start(x)
                x = Nothing
                Exit Sub
            End If
        Next
        oApplication.Utilities.Message("No file has been selected...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    End Sub
#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_Pay_AppHisDetails Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                            Case SAPbouiCOM.BoEventTypes.et_CLICK

                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "3" And (pVal.ColUID = "U_Z_Attachment") Then
                                    oGrid = oForm.Items.Item(pVal.ItemUID).Specific
                                    oGrid.Columns.Item("RowsHeader").Click(pVal.Row)
                                    LoadFiles(oForm, pVal.ItemUID)
                                    BubbleEvent = False
                                    Exit Sub
                                End If

                                If pVal.ItemUID = "3" And (pVal.ColUID = "U_Z_FileName") Then
                                    oGrid = oForm.Items.Item(pVal.ItemUID).Specific
                                    oGrid.Columns.Item("RowsHeader").Click(pVal.Row)
                                    LoadFiles1(oForm, pVal.ItemUID)
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "7" Then
                                    LoadFiles_New(oForm)
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                            Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE
                                If oForm.State = SAPbouiCOM.BoFormStateEnum.fs_Restore Or oForm.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized Then
                                    'oApplication.Utilities.Resize(oForm)
                                End If
                        End Select
                End Select
            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub
#End Region

    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD) Then
                oForm = oApplication.SBO_Application.Forms.ActiveForm()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

End Class
