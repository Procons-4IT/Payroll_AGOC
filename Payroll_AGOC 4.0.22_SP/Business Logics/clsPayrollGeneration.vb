Public Class clsPayrollGeneration
    Inherits clsBase
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox As SAPbouiCOM.ComboBox
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oGrid As SAPbouiCOM.Grid
    Private dtTemp As SAPbouiCOM.DataTable
    Private dtResult As SAPbouiCOM.DataTable
    Private oMode As SAPbouiCOM.BoFormMode
    Private oItem As SAPbobsCOM.Items
    Private oInvoice As SAPbobsCOM.Documents
    Private InvBase As DocumentType
    Private InvBaseDocNo As String
    Private InvForConsumedItems As Integer
    Private blnFlag As Boolean = False
    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub
    Private Sub LoadForm()
        If oApplication.Utilities.validateAuthorization(oApplication.Company.UserSignature, frm_PayrollGeneration) = False Then
            oApplication.Utilities.Message("You are not authorized to do this action", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        End If
        oForm = oApplication.Utilities.LoadForm(xml_Payrollgeneration, frm_PayrollGeneration)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        Try
            oForm.Freeze(True)
            oForm.EnableMenu(mnu_ADD_ROW, True)
            oForm.EnableMenu(mnu_DELETE_ROW, True)
            Databind(oForm, 0)
            oForm.PaneLevel = 1
            oForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub

#Region "LoadParroll Details"
    Private Sub LoadPayRollDetails(ByVal aform As SAPbouiCOM.Form)
        oGrid = aform.Items.Item("10").Specific


        Dim intYear, intMonth As Integer
        oCombobox = aform.Items.Item("7").Specific
        If oCombobox.Selected.Value = "" Then
            oApplication.Utilities.Message("Select year", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Else
            intYear = oCombobox.Selected.Value
            If intYear = 0 Then
                oApplication.Utilities.Message("Select year", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If
        End If
        oCombobox = aform.Items.Item("9").Specific
        If oCombobox.Selected.Value = "" Then
            oApplication.Utilities.Message("Select Month", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Else
            intMonth = oCombobox.Selected.Value
            If intMonth = 0 Then
                oApplication.Utilities.Message("Select Month", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If
        End If
        Dim strCode As String
        For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
            If oGrid.Rows.IsSelected(intRow) Then
                strCode = oGrid.DataTable.GetValue("Code", intRow)
                If strCode <> "" Then
                    Dim oOBj As New clsPayrolLDetails
                    frmSourceForm = aform
                    oOBj.LoadForm(intMonth, intYear, strCode, "Payroll")
                End If
            End If
        Next
    End Sub
#End Region
#Region "Databind"
    Private Sub Databind(ByVal aform As SAPbouiCOM.Form, ByVal intPane As Integer)
        Try
            aform.Freeze(True)
            If intPane = 0 Then
                aform.DataSources.UserDataSources.Add("intYear", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
                aform.DataSources.UserDataSources.Add("intMonth", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)

                aform.DataSources.UserDataSources.Add("intYear1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
                aform.DataSources.UserDataSources.Add("intMonth1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
                aform.DataSources.UserDataSources.Add("strComp", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)

                aform.DataSources.UserDataSources.Add("strPost", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)

                oCombobox = aform.Items.Item("24").Specific
                oCombobox.DataBind.SetBound(True, "", "strPost")
                oCombobox.ValidValues.Add("0", "")
                oCombobox.ValidValues.Add("H", "On Hold")
                oCombobox.ValidValues.Add("A", "Active")
                oCombobox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                oCombobox.Select("A", SAPbouiCOM.BoSearchKey.psk_ByValue)
                aform.Items.Item("24").DisplayDesc = True

                oCombobox = aform.Items.Item("7").Specific
                oCombobox.ValidValues.Add("0", "")
                For intRow As Integer = 2010 To 2050
                    oCombobox.ValidValues.Add(intRow, intRow)
                Next
                oCombobox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
                oCombobox.DataBind.SetBound(True, "", "intYear")
                aform.Items.Item("7").DisplayDesc = True
                oCombobox = aform.Items.Item("9").Specific
                oCombobox.ValidValues.Add("0", "")
                For intRow As Integer = 1 To 12
                    oCombobox.ValidValues.Add(intRow, MonthName(intRow))
                Next

                oCombobox.DataBind.SetBound(True, "", "intMonth")
                oCombobox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
                aform.Items.Item("9").DisplayDesc = True

                oEditText = aform.Items.Item("16").Specific
                oEditText.DataBind.SetBound(True, "", "intmonth1")
                oEditText = aform.Items.Item("18").Specific
                oEditText.DataBind.SetBound(True, "", "intYear1")

                oCombobox = aform.Items.Item("cmbCmp").Specific
                oCombobox.DataBind.SetBound(True, "", "strComp")
                oApplication.Utilities.FillCombobox(oCombobox, "Select U_Z_CompCode,U_Z_CompName from [@Z_OADM]")

            End If
            oGrid = aform.Items.Item("10").Specific
            dtTemp = oGrid.DataTable
            If intPane = 0 Then
                dtTemp.ExecuteQuery("SELECT T0.[empID], T0.[firstName]+T0.[LastName] 'Emplopyee name', T0.[jobTitle],T1.[Name], T0.[salary], T0.[salaryUnit], T2.[PrcName] FROM OHEM T0  INNER JOIN OUDP T1 ON T0.dept = T1.Code INNER JOIN OPRC T2 ON T0.U_Z_COST = T2.PrcCode where EmpId=10000000")
            Else
                dtTemp.ExecuteQuery("SELECT T0.[empID], T0.[firstName]+T0.[LastName] 'Emplopyee name', T0.[jobTitle],T1.[Name], T0.[salary], T0.[salaryUnit], T2.[PrcName] FROM OHEM T0  INNER JOIN OUDP T1 ON T0.dept = T1.Code INNER JOIN OPRC T2 ON T0.U_Z_COST = T2.PrcCode")
            End If
            oGrid.DataTable = dtTemp
            Formatgrid(oGrid, "Load")
            oApplication.Utilities.assignMatrixLineno(oGrid, aform)
            oForm.Items.Item("10").Enabled = False
            aform.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aform.Freeze(False)
        End Try
    End Sub
#End Region

#Region "Populate Payroll Worksheet Details"
    Public Function PrepareWorkSheet(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Try
            aForm.Freeze(True)
            Dim intYear, intMonth As Integer
            Dim strmonth, strPostMethod As String
            oCombobox = aForm.Items.Item("24").Specific
            If oCombobox.Selected.Value = "" Then
                oApplication.Utilities.Message("Select Posting Method", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                aForm.Freeze(False)
                Return False
            Else
                strPostMethod = oCombobox.Selected.Value
                If strPostMethod = "0" Then
                    oApplication.Utilities.Message("Select Posting Method", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    aForm.Freeze(False)
                    Return False
                End If
            End If
            oCombobox = aForm.Items.Item("7").Specific
            If oCombobox.Selected.Value = "" Then
                oApplication.Utilities.Message("Select year", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                aForm.Freeze(False)
                Return False
            Else
                intYear = oCombobox.Selected.Value
                If intYear = 0 Then
                    oApplication.Utilities.Message("Select year", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    aForm.Freeze(False)
                    Return False
                End If
            End If
            oCombobox = aForm.Items.Item("9").Specific
            If oCombobox.Selected.Value = "" Then
                oApplication.Utilities.Message("Select Month", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                aForm.Freeze(False)
                Return False
            Else
                intMonth = oCombobox.Selected.Value
                If intYear = 0 Then
                    oApplication.Utilities.Message("Select Month", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    aForm.Freeze(False)
                    Return False
                Else
                    strMonth = oCombobox.Selected.Description
                End If
            End If
            Dim strCompany As String
            oCombobox = aForm.Items.Item("cmbCmp").Specific
            If oCombobox.Selected.Value = "" Then
                oApplication.Utilities.Message("Select Company Code", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                aForm.Freeze(False)
                Return False
            Else
                strCompany = oCombobox.Selected.Value
            End If
            '  oApplication.Utilities.UpdatePayrollTotal(intMonth, intYear)
            Dim oPayrec, oTempRec As SAPbobsCOM.Recordset
            oPayrec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTempRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oCombobox = oForm.Items.Item("24").Specific
            '   MsgBox(oCombobox.Selected.Value)
            oApplication.Utilities.getRoundingDigit()
            '   oPayrec.DoQuery("Select * from [@Z_PAYROLL] where  U_Z_CompNo='" & strCompany & "' and U_Z_OffCycle='N' and  U_Z_Process='Y' and  U_Z_YEAR=" & intYear & " and U_Z_MONTH=" & intMonth)
            oPayrec.DoQuery("Select * from [@Z_PAYROLL1] where U_Z_OnHold='" & oCombobox.Selected.Value & "' and    U_Z_CompNo='" & strCompany & "' and U_Z_OffCycle='N' and  U_Z_Posted='Y' and  U_Z_YEAR=" & intYear & " and U_Z_MONTH=" & intMonth)
            If oPayrec.RecordCount > 0 Then
                oApplication.Utilities.Message("Payroll already processed for this selected period", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                aForm.Items.Item("5").Enabled = False
                '  aForm.Items.Item("5").Enabled = True
            Else
                aForm.Items.Item("5").Enabled = True
            End If


          

          
            oPayrec.DoQuery("Select * from [@Z_PAYROLL] where U_Z_CompNo='" & strCompany & "' and U_Z_OffCycle='N'  and U_Z_YEAR=" & intYear & " and U_Z_MONTH=" & intMonth)
            If oPayrec.RecordCount <= 0 Then
                oApplication.Utilities.Message("Payroll Worksheet not prepared for this selected month and year", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                aForm.Freeze(False)
                Return False
            Else
                If oApplication.Utilities.ValidateApprovalPending(strCompany, intYear, intMonth) = False Then
                    '  oApplication.Utilities.Message("Worksheet is under approval. You can not post the payroll ", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    aForm.Items.Item("5").Enabled = False
                End If
                Dim strquery As String
                Dim oRecordset As SAPbobsCOM.Recordset
                oRecordset = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                ' strquery = "Select isnull(U_Z_IsApp,'N')  from [@Z_OADM] where U_Z_CompCode='" & strCompany & "'"
                'oRecordset.DoQuery(strquery)
                strquery = "Select * from ""@Z_PAY_OAPPT"" T0 left join ""@Z_PAY_APPT1"" T1 on T0.""DocEntry""=T1.""DocEntry"" where isnull(T0.""U_Z_Active"",'N')='Y' and T0.""U_Z_DocType""='R' and T1.""U_Z_OUser""='" & strCompany & "' "
                oRecordset.DoQuery(strquery)
                If oRecordset.RecordCount > 0 Then
                    strquery = "Select  * from [@Z_PAY_Approval] where  U_Z_CompNo='" & strCompany & "' and U_Z_DocType='R' and U_Z_YEAR=" & intYear & " and U_Z_MONTH=" & intMonth
                    oRecordset.DoQuery(strquery)
                    If oRecordset.RecordCount <= 0 Then
                        oApplication.Utilities.Message("Worksheet requires approval for posting.. Please initiate the approval for worksheet and proceed...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        aForm.Items.Item("5").Enabled = False
                    End If

                End If
                oGrid = aForm.Items.Item("10").Specific
                dtTemp = oGrid.DataTable
                Dim strrefcode, strsql As String
                strrefcode = oPayrec.Fields.Item("Code").Value
                ' oApplication.Utilities.UpdatePayrollTotal_Payroll(intMonth, intYear, strrefcode)
                ' If GenerateWorkSheet(aForm) = True Then
                oApplication.Utilities.setEdittextvalue(aForm, "16", strmonth)
                oApplication.Utilities.setEdittextvalue(aForm, "18", intYear.ToString)

                oCombobox = aForm.Items.Item("cmbCmp").Specific
                oApplication.Utilities.setEdittextvalue(aForm, "cmbName", oCombobox.Selected.Description)

                'to check negative salary blocking
                Dim blnNetSalary As Boolean = False
                strsql = "SELECT T0.[Code], T0.[Name], T0.[U_Z_RefCode], T0.[U_Z_PersonalID], T0.[U_Z_TANO] 'TANO', T0.[U_Z_empid], T0.[U_Z_EmpName], Case T0.U_Z_OnHold when 'H' then 'On Hold' else 'Active' end 'Status', T0.[U_Z_JobTitle], T0.[U_Z_Department], T0.[U_Z_TermName] 'Contract Term',T0.[U_Z_Basic], T0.[U_Z_InrAmt], T0.[U_Z_BasicSalary],t0.U_Z_MonthlyBasic 'Monthly Basic', T0.[U_Z_SalaryType], T0.[U_Z_CostCentre], T0.[U_Z_Earning], T0.[U_Z_Deduction], T0.[U_Z_UnPaidLeave], T0.[U_Z_PaidLeave], T0.[U_Z_AnuLeave], T0.[U_Z_Contri], T0.[U_Z_AirAmt], T0.[U_Z_AcrAmt] ,T0.[U_Z_AcrAirAmt], T0.[U_Z_Cost], T0.[U_Z_NetSalary], isnull(T0.U_Z_MonthlyBasic,0) + isnull(T0.U_Z_Earning,0)  'GrossSalary', T0.[U_Z_Startdate], T0.[U_Z_TermDate], T0.[U_Z_JVNo], T0.[U_Z_EOSYTD] ,T0.[U_Z_EOSBalance], T0.[U_Z_EOS],T0.U_Z_WorkingDays1,T0.[U_Z_CalenderDays] 'Working Days of month',T0.[U_Z_TotalLeave] 'Leave Utilized',T0.[U_Z_ActWork] 'Total Worked days', T0.[U_Z_CompNo], T0.[U_Z_Branch], T0.[U_Z_Dept] FROM [dbo].[@Z_PAYROLL1]  T0 where isnull(U_Z_OffCycle,'N')='N' and T0.U_Z_NetSalary<0 and T0.U_Z_RefCode='" & strrefcode & "' and T0.U_Z_OnHold='" & strPostMethod & "'"
                oTempRec.DoQuery(strsql)
                If oTempRec.RecordCount > 0 Then
                    blnNetSalary = True
                End If
                oTempRec.DoQuery("Select isnull(U_Z_Block,'N') from [@Z_OADM] where U_Z_CompCode='" & strCompany & "'")
                If oTempRec.Fields.Item(0).Value = "N" Then
                    blnNetSalary = False
                End If
                If blnNetSalary = True Then
                    aForm.Items.Item("5").Enabled = False
                    oApplication.Utilities.Message("You cannot post the payroll with a negative salary.", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End If
                'end 

                If 1 = 1 Then
                    ' strsql = "Select * from [@Z_PAYROLL1] where U_Z_RefCode='" & strrefcode & "'"
                    ' strsql = "SELECT T0.[Code], T0.[Name], T0.[U_Z_RefCode], T0.[U_Z_PersonalID], T0.[U_Z_empid], T0.[U_Z_EmpName], T0.[U_Z_JobTitle], T0.[U_Z_Department], T0.[U_Z_Basic], T0.[U_Z_InrAmt], T0.[U_Z_BasicSalary], T0.[U_Z_SalaryType], T0.[U_Z_CostCentre], T0.[U_Z_Earning], T0.[U_Z_Deduction], T0.[U_Z_UnPaidLeave], T0.[U_Z_PaidLeave], T0.[U_Z_AnuLeave], T0.[U_Z_Contri], T0.[U_Z_Cost], T0.[U_Z_NetSalary], T0.[U_Z_Startdate], T0.[U_Z_TermDate], T0.[U_Z_JVNo], T0.[U_Z_EOS], T0.[U_Z_CompNo], T0.[U_Z_Branch], T0.[U_Z_Dept], T0.[U_Z_AirAmt], T0.[U_Z_AcrAmt] FROM [dbo].[@Z_PAYROLL1]  T0 where T0.U_Z_RefCode='" & strrefcode & "'"
                    ' strsql = "SELECT T0.[Code], T0.[Name], T0.[U_Z_RefCode], T0.[U_Z_PersonalID], T0.[U_Z_TANO] 'TANO', T0.[U_Z_empid], T0.[U_Z_EmpName], Case T0.U_Z_OnHold when 'H' then 'On Hold' else 'Active' end 'Status', T0.[U_Z_JobTitle], T0.[U_Z_Department], T0.[U_Z_TermName] 'Contract Term',T0.[U_Z_Basic], T0.[U_Z_InrAmt], T0.[U_Z_BasicSalary],t0.U_Z_MonthlyBasic 'Monthly Basic', T0.[U_Z_SalaryType],T0.[U_Z_Cost], T0.[U_Z_NetSalary], isnull(T0.U_Z_MonthlyBasic,0) + isnull(T0.U_Z_Earning,0)  'GrossSalary', T0.[U_Z_CostCentre], T0.[U_Z_Earning], T0.[U_Z_Deduction], T0.[U_Z_UnPaidLeave], T0.[U_Z_PaidLeave], T0.[U_Z_AnuLeave], T0.[U_Z_Contri], T0.[U_Z_AirAmt], T0.[U_Z_AcrAmt] ,T0.[U_Z_AcrAirAmt],  T0.[U_Z_Startdate], T0.[U_Z_TermDate], T0.[U_Z_JVNo], T0.[U_Z_EOSYTD] ,T0.[U_Z_EOSBalance], T0.[U_Z_EOS],T0.U_Z_WorkingDays1,T0.[U_Z_CalenderDays] 'Working Days of month',T0.[U_Z_TotalLeave] 'Leave Utilized',T0.[U_Z_ActWork] 'Total Worked days', T0.[U_Z_CompNo], T0.[U_Z_Branch], T0.[U_Z_Dept] FROM [dbo].[@Z_PAYROLL1]  T0 where isnull(U_Z_OffCycle,'N')='N' and T0.U_Z_RefCode='" & strrefcode & "' and T0.U_Z_OnHold='" & strPostMethod & "'"
                    strsql = "SELECT T0.[Code], T0.[Name],T0.[U_Z_TANO] 'TANO',T0.[U_Z_empid],T0.[U_Z_ExtNo] 'Batch No', T0.[U_Z_EmpName],T0.[U_Z_Country] 'Country', Case T0.""U_Z_OnHold"" when 'H' then 'On Hold' else 'Active' end ""Status"", T0.[U_Z_Basic], T0.[U_Z_InrAmt], T0.[U_Z_BasicSalary], T0.[U_Z_MonthlyBasic] 'Monthly Basic',T0.[U_Z_ActWork] 'Total Worked days', T0.[U_Z_Cost], isnull(T0.U_Z_MonthlyBasic,0) + isnull(T0.U_Z_Earning,0)  'GrossSalary',T0.[U_Z_Earning], T0.[U_Z_Deduction], T0.[U_Z_NetSalary], T0.[U_Z_UnPaidLeave], T0.[U_Z_PaidLeave], T0.[U_Z_AnuLeave],T0.""U_Z_CashOutAmt"", T0.[U_Z_Contri], T0.[U_Z_AirAmt], ""U_Z_NetPayAmt"",""U_Z_CmpPayAmt"", T0.[U_Z_AcrAmt] ,T0.[U_Z_AcrAirAmt], T0.[U_Z_EOSYTD] ,T0.[U_Z_EOSBalance],T0.[U_Z_EOS],T0.U_Z_WorkingDays1, T0.[U_Z_CalenderDays] 'Working Days of month',T0.[U_Z_TotalLeave] 'Leave Utilized ',T0.[U_Z_RefCode], T0.[U_Z_PersonalID],  T0.[U_Z_JobTitle], T0.[U_Z_Department],T0.[U_Z_EmpBranch], T0.[U_Z_TermName] 'Contract Term', T0.[U_Z_SalaryType], T0.[U_Z_CostCentre],  T0.[U_Z_Startdate], T0.[U_Z_TermDate], T0.[U_Z_JVNo],  T0.[U_Z_CompNo], T0.[U_Z_Branch], T0.[U_Z_Dept],T0.""U_Z_EOS1"",T0.""U_Z_Leave"",T0.""U_Z_Ticket"",T0.""U_Z_Saving"",T0.""U_Z_PaidExtraSalary"",T0.""U_Z_GOVAMT"" 'Social Gov.Amt' FROM [dbo].[@Z_PAYROLL1]  T0  Inner Join OHEM T1 on T1.empID=T0.U_Z_EmpID where T0.U_Z_RefCode='" & strrefcode & "'  and T0.U_Z_OnHold='" & strPostMethod & "'"

                    'oTempRec.DoQuery("SELECT T0.[empID], T0.[firstName]+T0.[LastName] 'Emplopyee name', T0.[jobTitle],T1.[Name], T0.[salary], T0.[salaryUnit], T2.[PrcName] FROM OHEM T0  INNER JOIN OUDP T1 ON T0.dept = T1.Code INNER JOIN OPRC T2 ON T0.U_Z_COST = T2.PrcCode")
                    oGrid.DataTable.ExecuteQuery(strsql)
                    Formatgrid(oGrid, "Payroll")
                    oApplication.Utilities.assignMatrixLineno_Payroll(oGrid, aForm)
                End If
            End If
            aForm.Freeze(False)
            Return True
            aForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)
            Return False
        End Try
        Return True
    End Function


    Private Function GenerateWorkSheet(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Try
            aForm.Freeze(True)
            Dim intYear, intMonth As Integer
            oCombobox = aForm.Items.Item("7").Specific
            If oCombobox.Selected.Value = "" Then
                oApplication.Utilities.Message("Select year", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                aForm.Freeze(False)
                Return False
            Else
                intYear = oCombobox.Selected.Value
                If intYear = 0 Then
                    oApplication.Utilities.Message("Select year", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    aForm.Freeze(False)
                    Return False
                End If
            End If
            oCombobox = aForm.Items.Item("9").Specific
            If oCombobox.Selected.Value = "" Then
                oApplication.Utilities.Message("Select Month", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                aForm.Freeze(False)
                Return False
            Else
                intMonth = oCombobox.Selected.Value
                If intMonth = 0 Then
                    oApplication.Utilities.Message("Select Month", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    aForm.Freeze(False)
                    Return False
                End If
            End If

            Dim oPayrec, oTempRec As SAPbobsCOM.Recordset
            Dim strPayrollcode As String
            oPayrec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTempRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            oPayrec.DoQuery("Select * from [@Z_PAYROLL] where   U_Z_Process='Y' and U_Z_YEAR=" & intYear & " and U_Z_MONTH=" & intMonth)
            If oPayrec.RecordCount > 0 Then
                oApplication.Utilities.Message("Payroll already generated for this selected period", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                aForm.Freeze(False)
                Return False
            End If



            oPayrec.DoQuery("Select * from [@Z_PAYROLL] where U_Z_YEAR=" & intYear & " and U_Z_MONTH=" & intMonth)
            If oPayrec.RecordCount <= 0 Then
                strPayrollcode = AddtoPayroll(intYear, intMonth)
                If strPayrollcode <> "" Then
                    If AddPayRoll1(strPayrollcode) = True Then
                        If Addearning(strPayrollcode) = True Then
                            If AddDeduction(strPayrollcode) Then
                                If AddContribution(strPayrollcode) Then
                                End If
                            End If
                        End If
                    End If
                End If

            Else
                strPayrollcode = oPayrec.Fields.Item("Code").Value
                If strPayrollcode <> "" Then
                    If AddPayRoll1(strPayrollcode) = True Then
                        If Addearning(strPayrollcode) = True Then
                            If AddDeduction(strPayrollcode) Then
                                If AddContribution(strPayrollcode) Then
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            oApplication.Utilities.UpdatePayrollTotal(intMonth, intYear)
            oApplication.Utilities.Message("Payroll Worksheet generation Completed", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            aForm.Freeze(False)
            Return True
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)
            Return False
        End Try
        Return True
    End Function
#End Region

#Region "FormatGrid"
    Private Sub Formatgrid(ByVal agrid As SAPbouiCOM.Grid, ByVal aOption As String)
        Select Case aOption
            Case "Load"
                agrid.Columns.Item(0).TitleObject.Caption = "Employee ID"
                agrid.Columns.Item(1).TitleObject.Caption = "Employee Name"
                agrid.Columns.Item(2).TitleObject.Caption = "Job Title"
                agrid.Columns.Item(3).TitleObject.Caption = "Department"
                agrid.Columns.Item(4).TitleObject.Caption = "Salary"
                agrid.Columns.Item(5).TitleObject.Caption = "Salary Type"
                agrid.Columns.Item(6).TitleObject.Caption = "Cost Center"
                oEditTextColumn = agrid.Columns.Item(0)
                oEditTextColumn.LinkedObjectType = "171"
            Case "Payroll"
                'agrid.Columns.Item(0).TitleObject.Caption = "Code"
                'agrid.Columns.Item(1).TitleObject.Caption = "Name"
                'agrid.Columns.Item(1).Visible = False
                'agrid.Columns.Item(2).TitleObject.Caption = "Reference Code"
                'agrid.Columns.Item(2).Visible = False
                'agrid.Columns.Item(3).TitleObject.Caption = "Employee ID"
                'agrid.Columns.Item(4).TitleObject.Caption = "Employee Name"
                'agrid.Columns.Item(5).TitleObject.Caption = "Job Title"
                'agrid.Columns.Item(6).TitleObject.Caption = "Department"
                'agrid.Columns.Item(7).TitleObject.Caption = "Salary"
                'agrid.Columns.Item(8).TitleObject.Caption = "Salary Type"
                'agrid.Columns.Item(9).TitleObject.Caption = "Cost Center"
                'agrid.Columns.Item(10).TitleObject.Caption = "Earnings"
                'oEditTextColumn = oGrid.Columns.Item(10)
                'oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                'agrid.Columns.Item(11).TitleObject.Caption = "Deduction"
                'oEditTextColumn = oGrid.Columns.Item(11)
                'oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto

                'agrid.Columns.Item(12).TitleObject.Caption = "UnPaid Leave"
                'oEditTextColumn = oGrid.Columns.Item(12)
                'oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto

                ''agrid.Columns.Item(13).TitleObject.Caption = "Contribution"
                ''oEditTextColumn = oGrid.Columns.Item(13)
                ''oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                ''agrid.Columns.Item(14).TitleObject.Caption = "Total Cost"
                ''oEditTextColumn = oGrid.Columns.Item(14)
                ''oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                ''agrid.Columns.Item(15).TitleObject.Caption = "Net Salary"
                ''oEditTextColumn = oGrid.Columns.Item(15)
                ''oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                ''oEditTextColumn = agrid.Columns.Item(3)
                ''oEditTextColumn.LinkedObjectType = "171"
                ''agrid.Columns.Item(16).TitleObject.Caption = "Joining Date"
                ''agrid.Columns.Item(17).TitleObject.Caption = "Termination Date"
                ''agrid.Columns.Item(18).TitleObject.Caption = "Journal Voucher Ref"
                ''oEditTextColumn = agrid.Columns.Item(18)
                ''oEditTextColumn.LinkedObjectType = "28"
                ''agrid.Columns.Item("U_Z_EOS").TitleObject.Caption = "End of Service"
                ''agrid.Columns.Item("U_Z_EOS").Editable = False
                ''agrid.Columns.Item("U_Z_CompNo").TitleObject.Caption = "Company Code"
                ''agrid.Columns.Item("U_Z_CompNo").Editable = False
                ''agrid.Columns.Item("U_Z_Branch").TitleObject.Caption = "Branch"
                ''agrid.Columns.Item("U_Z_Branch").Editable = False
                ''agrid.Columns.Item("U_Z_Dept").TitleObject.Caption = "Department "
                ''agrid.Columns.Item("U_Z_Dept").Editable = False
                ''agrid.Columns.Item("U_Z_AirAmt").TitleObject.Caption = "AirTicket Availed Amount"
                ''agrid.Columns.Item("U_Z_AirAmt").Editable = False

                'agrid.Columns.Item("U_Z_Contri").TitleObject.Caption = "Contribution"
                'oEditTextColumn = oGrid.Columns.Item("U_Z_Contri")
                'oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                'agrid.Columns.Item("U_Z_Cost").TitleObject.Caption = "Total Cost"
                'oEditTextColumn = oGrid.Columns.Item("U_Z_Cost")
                'oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                'agrid.Columns.Item("U_Z_NetSalary").TitleObject.Caption = "Net Salary"
                'oEditTextColumn = oGrid.Columns.Item("U_Z_NetSalary")
                'oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                'oEditTextColumn = agrid.Columns.Item(3)
                'oEditTextColumn.LinkedObjectType = "171"
                'agrid.Columns.Item("U_Z_Startdate").TitleObject.Caption = "Joining Date"
                'agrid.Columns.Item("U_Z_TermDate").TitleObject.Caption = "Termination Date"
                'agrid.Columns.Item("U_Z_JVNo").TitleObject.Caption = "Journal Voucher Ref"
                'oEditTextColumn = agrid.Columns.Item("U_Z_JVNo")
                'oEditTextColumn.LinkedObjectType = "28"
                'agrid.Columns.Item("U_Z_EOS").TitleObject.Caption = "End of Service"
                'agrid.Columns.Item("U_Z_EOS").Editable = False
                'agrid.Columns.Item("U_Z_CompNo").TitleObject.Caption = "Company Code"
                'agrid.Columns.Item("U_Z_CompNo").Editable = False

                'agrid.Columns.Item("U_Z_Branch").TitleObject.Caption = "Branch"
                'agrid.Columns.Item("U_Z_Branch").Editable = False
                'agrid.Columns.Item("U_Z_Dept").TitleObject.Caption = "Department "
                'agrid.Columns.Item("U_Z_Dept").Editable = False
                'agrid.Columns.Item("U_Z_AirAmt").TitleObject.Caption = "AirTicket Availed Amount"
                'agrid.Columns.Item("U_Z_AirAmt").Editable = False
                'agrid.Columns.Item("U_Z_AnuLeave").TitleObject.Caption = "Annual Leave"
                'agrid.Columns.Item("U_Z_AnuLeave").Editable = False

                agrid.Columns.Item("Code").TitleObject.Caption = "Code"
                agrid.Columns.Item("Name").TitleObject.Caption = "Name"
                agrid.Columns.Item("Name").Visible = False
                agrid.Columns.Item("TANO").TitleObject.Caption = "T & A Employee No"
                agrid.Columns.Item("U_Z_RefCode").TitleObject.Caption = "Reference Code"
                agrid.Columns.Item("U_Z_RefCode").Visible = False
                agrid.Columns.Item("U_Z_empid").TitleObject.Caption = "Employee ID"
                agrid.Columns.Item("U_Z_EmpName").TitleObject.Caption = "Employee Name"
                agrid.Columns.Item("U_Z_JobTitle").TitleObject.Caption = "Job Title"
                agrid.Columns.Item("U_Z_Department").TitleObject.Caption = "Department"
                agrid.Columns.Item("U_Z_BasicSalary").TitleObject.Caption = "Total Basic Salary"
                oEditTextColumn = oGrid.Columns.Item("U_Z_BasicSalary")
                oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                agrid.Columns.Item("U_Z_SalaryType").TitleObject.Caption = "Salary Type"
                agrid.Columns.Item("U_Z_CostCentre").TitleObject.Caption = "Cost Center"
                agrid.Columns.Item("U_Z_Earning").TitleObject.Caption = "Earnings"
                oEditTextColumn = oGrid.Columns.Item("U_Z_Earning")
                oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                agrid.Columns.Item("U_Z_Deduction").TitleObject.Caption = "Deductions"
                oEditTextColumn = oGrid.Columns.Item("U_Z_Deduction")
                oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                agrid.Columns.Item("U_Z_UnPaidLeave").TitleObject.Caption = "UnPaid Leave"
                oEditTextColumn = oGrid.Columns.Item("U_Z_UnPaidLeave")
                oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                agrid.Columns.Item("Monthly Basic").TitleObject.Caption = "Current Month Basic"
                oEditTextColumn = oGrid.Columns.Item("Monthly Basic")
                oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                agrid.Columns.Item("U_Z_PaidLeave").TitleObject.Caption = "Paid Leave"
                oEditTextColumn = oGrid.Columns.Item("U_Z_PaidLeave")
                oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto

                agrid.Columns.Item("U_Z_Contri").TitleObject.Caption = "Contributions"
                oEditTextColumn = oGrid.Columns.Item("U_Z_Contri")
                oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                agrid.Columns.Item("U_Z_Cost").TitleObject.Caption = "Total Cost"
                oEditTextColumn = oGrid.Columns.Item("U_Z_Cost")
                oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                agrid.Columns.Item("U_Z_NetSalary").TitleObject.Caption = "Net Salary"
                oEditTextColumn = oGrid.Columns.Item("U_Z_NetSalary")
                oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                oEditTextColumn = agrid.Columns.Item("U_Z_empid")
                oEditTextColumn.LinkedObjectType = "171"
                agrid.Columns.Item("U_Z_Startdate").TitleObject.Caption = "Joining Date"
                agrid.Columns.Item("U_Z_TermDate").TitleObject.Caption = "Termination Date"
                agrid.Columns.Item("U_Z_JVNo").TitleObject.Caption = "Journal Voucher Ref"


                oEditTextColumn = agrid.Columns.Item("U_Z_JVNo")
                oEditTextColumn.LinkedObjectType = "28"
                agrid.Columns.Item("U_Z_EOS").TitleObject.Caption = "End of Service"
                agrid.Columns.Item("U_Z_EOS").Editable = False
                oEditTextColumn = oGrid.Columns.Item("U_Z_EOS")
                oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                agrid.Columns.Item("U_Z_CompNo").TitleObject.Caption = "Company Code"
                agrid.Columns.Item("U_Z_CompNo").Editable = False

                agrid.Columns.Item("U_Z_Branch").TitleObject.Caption = "Branch"
                agrid.Columns.Item("U_Z_Branch").Editable = False
                agrid.Columns.Item("U_Z_Dept").TitleObject.Caption = "Department "
                agrid.Columns.Item("U_Z_Dept").Editable = False
                agrid.Columns.Item("U_Z_AirAmt").TitleObject.Caption = "AirTicket Availed Amount"
                agrid.Columns.Item("U_Z_AirAmt").Editable = False
                oEditTextColumn = oGrid.Columns.Item("U_Z_AirAmt")
                oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                agrid.Columns.Item("U_Z_AnuLeave").TitleObject.Caption = "Annual Leave"
                agrid.Columns.Item("U_Z_AnuLeave").Editable = False
                oEditTextColumn = oGrid.Columns.Item("U_Z_AnuLeave")
                oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                agrid.Columns.Item("U_Z_PersonalID").TitleObject.Caption = "Government ID"
                agrid.Columns.Item("U_Z_PersonalID").Editable = False

                agrid.Columns.Item("U_Z_Basic").TitleObject.Caption = "Basic Salary"
                agrid.Columns.Item("U_Z_Basic").Editable = False
                oEditTextColumn = oGrid.Columns.Item("U_Z_Basic")
                oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                agrid.Columns.Item("U_Z_InrAmt").TitleObject.Caption = "Increment Amount"
                agrid.Columns.Item("U_Z_InrAmt").Editable = False
                oEditTextColumn = oGrid.Columns.Item("U_Z_InrAmt")
                oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                agrid.Columns.Item("U_Z_AcrAmt").TitleObject.Caption = "Annual Leave Accural Amount"
                agrid.Columns.Item("U_Z_AcrAmt").Editable = False
                oEditTextColumn = oGrid.Columns.Item("U_Z_AcrAmt")
                oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                agrid.Columns.Item("U_Z_AcrAirAmt").TitleObject.Caption = "AirTicket Accural Amount"
                agrid.Columns.Item("U_Z_AcrAirAmt").Editable = False
                oEditTextColumn = oGrid.Columns.Item("U_Z_AcrAirAmt")
                oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                agrid.Columns.Item("U_Z_EOSYTD").TitleObject.Caption = "EOS YTD"
                agrid.Columns.Item("U_Z_EOSYTD").Editable = False
                oEditTextColumn = oGrid.Columns.Item("U_Z_EOSYTD")
                oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                agrid.Columns.Item("U_Z_EOSBalance").TitleObject.Caption = "Total EOS Accural Balance"
                agrid.Columns.Item("U_Z_EOSBalance").Editable = False
                oEditTextColumn = oGrid.Columns.Item("U_Z_EOSBalance")
                oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto

                oGrid.Columns.Item("U_Z_WorkingDays1").TitleObject.Caption = "Basic Salary Calculation Days"
                oGrid.Columns.Item("U_Z_WorkingDays1").Editable = False

                oGrid.Columns.Item("GrossSalary").TitleObject.Caption = "Gross Salary"
                oGrid.Columns.Item("GrossSalary").Editable = False
                oEditTextColumn = oGrid.Columns.Item("GrossSalary")
                oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto


                oGrid.Columns.Item("U_Z_EOS1").TitleObject.Caption = "Include EOS Amount"
                oGrid.Columns.Item("U_Z_EOS1").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
                oGrid.Columns.Item("U_Z_EOS1").Editable = False

                oGrid.Columns.Item("U_Z_Leave").TitleObject.Caption = "Include Leave Amount"
                oGrid.Columns.Item("U_Z_Leave").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
                oGrid.Columns.Item("U_Z_Leave").Editable = False
                oGrid.Columns.Item("U_Z_Ticket").TitleObject.Caption = "Include Ticket Amount"
                oGrid.Columns.Item("U_Z_Ticket").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
                oGrid.Columns.Item("U_Z_Ticket").Editable = False
                oGrid.Columns.Item("U_Z_Saving").TitleObject.Caption = "Include Saving Amount"
                oGrid.Columns.Item("U_Z_Saving").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
                oGrid.Columns.Item("U_Z_Saving").Editable = False
                oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                oGrid.Columns.Item("U_Z_PaidExtraSalary").TitleObject.Caption = "Include Extra Salary"
                oGrid.Columns.Item("U_Z_PaidExtraSalary").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
                oGrid.Columns.Item("U_Z_PaidExtraSalary").Editable = False
                oGrid.Columns.Item("U_Z_CashOutAmt").TitleObject.Caption = "Leave Cashout Amount"
                oGrid.Columns.Item("U_Z_CashOutAmt").Editable = False

                oGrid.Columns.Item("U_Z_WorkingDays1").TitleObject.Caption = "Basic Salary Calculation Days"
                oGrid.Columns.Item("U_Z_WorkingDays1").Editable = False

                oGrid.Columns.Item("GrossSalary").TitleObject.Caption = "Gross Salary"
                oGrid.Columns.Item("GrossSalary").Editable = False
                oEditTextColumn = oGrid.Columns.Item("GrossSalary")
        End Select

        agrid.AutoResizeColumns()
        agrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
    End Sub
#End Region

#Region "AddRow"
    Private Sub AddEmptyRow(ByVal aGrid As SAPbouiCOM.Grid)
        'If aGrid.DataTable.GetValue("Code", aGrid.DataTable.Rows.Count - 1) <> "" Then
        '    aGrid.DataTable.Rows.Add()
        '    aGrid.Columns.Item(0).Click(aGrid.DataTable.Rows.Count - 1, False)
        'End If
    End Sub
#End Region

#Region "CommitTrans"
    Private Sub Committrans(ByVal strChoice As String)
        Dim oTemprec, oItemRec As SAPbobsCOM.Recordset
        oTemprec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oItemRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        If strChoice = "Cancel" Then
            oTemprec.DoQuery("Update [@Z_PAY_ODED] set Name=Code where Name Like '%D'")
        Else
            oTemprec.DoQuery("Select * from [@Z_PAY_ODED] where Name like '%D'")
            For intRow As Integer = 0 To oTemprec.RecordCount - 1
                oItemRec.DoQuery("delete from [@Z_PAY_ODED] where Name='" & oTemprec.Fields.Item("Name").Value & "' and Code='" & oTemprec.Fields.Item("Code").Value & "'")
                oTemprec.MoveNext()
            Next
            oTemprec.DoQuery("Delete from  [@Z_PAY_ODED]  where Name Like '%D'")
        End If

    End Sub
#End Region

#Region "Reset Payroll Worksheet"
    Private Function ResetPayrollWorksheet(ByVal aYear As Integer, ByVal aMonth As Integer, ByVal aForm As SAPbouiCOM.Form) As Boolean
        Dim oTemp, oTemp1, oTemp2 As SAPbobsCOM.Recordset
        Dim strPayRefcod, strEmpRefCode, strPostMethod As String
        oTemp1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp2 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oCombobox = aForm.Items.Item("24").Specific
        strPostMethod = oCombobox.Selected.Value
        oApplication.Utilities.getRoundingDigit()
        oCombobox = aForm.Items.Item("cmbCmp").Specific
        If oCombobox.Selected.Value = "" Then
        End If
        ''  If oApplication.Utilities.PostJournalVoucher(aMonth, aYear, oCombobox.Selected.Value) = True Then
        'If oApplication.Utilities.PostJournalVoucher_GroupbyBranch(aMonth, aYear, oCombobox.Selected.Value) = True Then
        '    oTemp1.DoQuery("Update [@Z_PAYROLL] set U_Z_Process='Y'  where U_Z_CompNo='" & oCombobox.Selected.Value & "' and  U_Z_Year=" & aYear & " and U_Z_Month=" & aMonth & " and U_Z_Process='N'")
        '    '  LoadPayRollDetails(aForm)
        '    PrepareWorkSheet(oForm)
        '    Return True
        'Else
        '    Return False

        ' End If

        oTemp1.DoQuery("Select isnull(U_Z_PostType,'C'),isnull(U_Z_JVType,'V') from [@Z_OADM] where U_Z_CompCode='" & oCombobox.Selected.Value & "'")
        If oTemp1.Fields.Item(1).Value = "V" Then
            If oTemp1.Fields.Item(0).Value = "P" Then
                If oApplication.Utilities.PostJournalVoucher_GroupbyBranch_Project(aMonth, aYear, oCombobox.Selected.Value, strPostMethod) = True Then
                    oTemp1.DoQuery("Update [@Z_PAYROLL] set U_Z_Process='Y'  where   U_Z_OffCycle<>'Y' and  U_Z_CompNo='" & oCombobox.Selected.Value & "' and  U_Z_Year=" & aYear & " and U_Z_Month=" & aMonth & " and U_Z_Process='N'")
                    '  LoadPayRollDetails(aForm)
                    PrepareWorkSheet(oForm)
                    Return True
                Else
                    Return False

                End If
            ElseIf oTemp1.Fields.Item(0).Value = "C" Then
                If oApplication.Utilities.PostJournalVoucher_GroupbyBranch(aMonth, aYear, oCombobox.Selected.Value, strPostMethod) = True Then
                    oTemp1.DoQuery("Update [@Z_PAYROLL] set U_Z_Process='Y'  where  U_Z_OffCycle<>'Y' and  U_Z_CompNo='" & oCombobox.Selected.Value & "' and  U_Z_Year=" & aYear & " and U_Z_Month=" & aMonth & " and U_Z_Process='N'")
                    '  LoadPayRollDetails(aForm)
                    PrepareWorkSheet(oForm)
                    Return True
                Else
                    Return False
                End If
            Else
                If oApplication.Utilities.PostJournalVoucher_Employee(aMonth, aYear, oCombobox.Selected.Value, strPostMethod) = True Then
                    oTemp1.DoQuery("Update [@Z_PAYROLL] set U_Z_Process='Y'  where  U_Z_OffCycle<>'Y' and  U_Z_CompNo='" & oCombobox.Selected.Value & "' and  U_Z_Year=" & aYear & " and U_Z_Month=" & aMonth & " and U_Z_Process='N'")
                    '  LoadPayRollDetails(aForm)
                    PrepareWorkSheet(oForm)
                    Return True
                Else
                    Return False
                End If
            End If
        Else 'Journal Entry posting
            If oTemp1.Fields.Item(0).Value = "P" Then
                If oApplication.Utilities.PostJournalEntries_GroupbyBranch_Project(aMonth, aYear, oCombobox.Selected.Value, strPostMethod) = True Then
                    oTemp1.DoQuery("Update [@Z_PAYROLL] set U_Z_Process='Y'  where  U_Z_OffCycle<>'Y' and  U_Z_CompNo='" & oCombobox.Selected.Value & "' and  U_Z_Year=" & aYear & " and U_Z_Month=" & aMonth & " and U_Z_Process='N'")
                    '  LoadPayRollDetails(aForm)
                    PrepareWorkSheet(oForm)
                    Return True
                Else
                    Return False
                End If
            ElseIf oTemp1.Fields.Item(0).Value = "C" Then
                If oApplication.Utilities.PostJournalEntries_GroupbyBranch(aMonth, aYear, oCombobox.Selected.Value, strPostMethod) = True Then
                    oTemp1.DoQuery("Update [@Z_PAYROLL] set U_Z_Process='Y'  where  U_Z_OffCycle<>'Y' and  U_Z_CompNo='" & oCombobox.Selected.Value & "' and  U_Z_Year=" & aYear & " and U_Z_Month=" & aMonth & " and U_Z_Process='N'")
                    '  LoadPayRollDetails(aForm)
                    PrepareWorkSheet(oForm)
                    Return True
                Else
                    Return False
                End If
            Else
                If oApplication.Utilities.PostJournalEntries_Employee(aMonth, aYear, oCombobox.Selected.Value, strPostMethod) = True Then
                    oTemp1.DoQuery("Update [@Z_PAYROLL] set U_Z_Process='Y'  where  U_Z_OffCycle<>'Y' and  U_Z_CompNo='" & oCombobox.Selected.Value & "' and  U_Z_Year=" & aYear & " and U_Z_Month=" & aMonth & " and U_Z_Process='N'")
                    '  LoadPayRollDetails(aForm)
                    PrepareWorkSheet(oForm)
                    Return True
                Else
                    Return False

                End If


            End If

        End If



        'If oTemp1.RecordCount > 0 Then
        '    strPayRefcod = oTemp1.Fields.Item("Code").Value
        '    If strPayRefcod <> "" Then
        '        oTemp2.DoQuery("Select * from [@Z_PAYROLL1] where U_Z_RefCode='" & strPayRefcod & "'")
        '        For intRow As Integer = 0 To oTemp2.RecordCount - 1
        '            strEmpRefCode = oTemp2.Fields.Item("Code").Value
        '            If strEmpRefCode <> "" Then
        '                oTemp.DoQuery("Delete from [@Z_PAYROLL2] where U_Z_RefCode='" & strEmpRefCode & "'")
        '                oTemp.DoQuery("Delete from [@Z_PAYROLL3] where U_Z_RefCode='" & strEmpRefCode & "'")
        '                oTemp.DoQuery("Delete from [@Z_PAYROLL4] where U_Z_RefCode='" & strEmpRefCode & "'")
        '            End If
        '            oTemp2.MoveNext()
        '        Next


        '    End If
        'End If
    End Function
#End Region

#Region "AddtoUDT"
    Private Function AddtoPayroll(ByVal aYear As Integer, ByVal aMonth As Integer) As String
        Dim oUserTable, oUserTable1, ousertable2, ousertable3 As SAPbobsCOM.UserTable
        Dim strCode, strECode, strEname, strGLacc, strRefCode, strPayrollRefNo, strempID As String
        Dim oTempRec, oTemp1, otemp2, otemp3, otemp4 As SAPbobsCOM.Recordset
        oTempRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otemp2 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otemp3 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otemp4 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oUserTable = oApplication.Company.UserTables.Item("Z_PAYROLL")
        strCode = oApplication.Utilities.getMaxCode("@Z_PAYROLL", "Code")
        oUserTable.Code = strCode
        oUserTable.Name = strCode & "N"
        oUserTable.UserFields.Fields.Item("U_Z_YEAR").Value = aYear
        oUserTable.UserFields.Fields.Item("U_Z_MONTH").Value = aMonth
        oUserTable.UserFields.Fields.Item("U_Z_Process").Value = "N"
        If oUserTable.Add <> 0 Then
            oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return ""
        Else
            Return strCode
        End If
    End Function
    Private Function AddPayRoll1(ByVal arefCode) As Boolean
        Dim oUserTable, oUserTable1, ousertable2, ousertable3 As SAPbobsCOM.UserTable
        Dim strCode, strECode, strEname, strGLacc, strRefCode, strPayrollRefNo, strempID As String
        Dim oTempRec, oTemp1, otemp2, otemp3, otemp4 As SAPbobsCOM.Recordset
        oTempRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otemp2 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otemp3 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otemp4 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        strRefCode = arefCode
        'otemp2.DoQuery("Select * from [@Z_PAYROLL1] where U_Z_RefCode='" & arefCode & "'")
        'If otemp2.RecordCount > 0 Then
        '    Return True
        'End If
        oTempRec.DoQuery("SELECT T0.[empID], T0.[firstName]+T0.[LastName] 'Emplopyee name', T0.[jobTitle],T1.[Name], T0.[salary], T0.[salaryUnit], isnull(T2.[PrcName],'') FROM OHEM T0  Left Outer JOIN OUDP T1 ON T0.dept = T1.Code Left Outer JOIN OPRC T2 ON T0.U_Z_COST = T2.PrcCode")
        oUserTable1 = oApplication.Company.UserTables.Item("Z_PAYROLL1")
        For intRow As Integer = 0 To oTempRec.RecordCount - 1
            otemp2.DoQuery("Select * from [@Z_PAYROLL1] where U_Z_empid='" & oTempRec.Fields.Item(0).Value & "' and  U_Z_RefCode='" & arefCode & "'")
            If otemp2.RecordCount <= 0 Then
                oApplication.Utilities.Message("Processing..", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                strCode = oApplication.Utilities.getMaxCode("@Z_PAYROLL1", "Code")
                oUserTable1.Code = strCode
                oUserTable1.Name = strCode & "N"
                strempID = oTempRec.Fields.Item(0).Value
                oUserTable1.UserFields.Fields.Item("U_Z_RefCode").Value = strRefCode
                oUserTable1.UserFields.Fields.Item("U_Z_empid").Value = oTempRec.Fields.Item(0).Value
                oUserTable1.UserFields.Fields.Item("U_Z_EmpName").Value = oTempRec.Fields.Item(1).Value
                oUserTable1.UserFields.Fields.Item("U_Z_JobTitle").Value = oTempRec.Fields.Item(2).Value
                oUserTable1.UserFields.Fields.Item("U_Z_Department").Value = oTempRec.Fields.Item(3).Value
                oUserTable1.UserFields.Fields.Item("U_Z_BasicSalary").Value = oTempRec.Fields.Item(4).Value
                oUserTable1.UserFields.Fields.Item("U_Z_SalaryType").Value = oTempRec.Fields.Item(5).Value
                oUserTable1.UserFields.Fields.Item("U_Z_CostCentre").Value = oTempRec.Fields.Item(6).Value
                If oUserTable1.Add <> 0 Then
                    oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    If oApplication.Company.InTransaction Then
                        oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    End If
                    Return False
                Else

                End If

            End If
            oTempRec.MoveNext()
        Next

        Return True
    End Function

    Private Function Addearning(ByVal arefCode As String) As Boolean
        Dim oUserTable, oUserTable1, ousertable2, ousertable3 As SAPbobsCOM.UserTable
        Dim strCode, strECode, strEname, strGLacc, strRefCode, strPayrollRefNo, strempID As String
        Dim oTempRec, oTemp1, otemp2, otemp3, otemp4 As SAPbobsCOM.Recordset
        oTempRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otemp2 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otemp3 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otemp4 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        If oApplication.Company.InTransaction Then
            oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
        End If
        oApplication.Company.StartTransaction()
        If 1 = 1 Then
            strRefCode = arefCode
            oTempRec.DoQuery("SELECT * from [@Z_PAYROLL1] where U_Z_RefCode='" & arefCode & "'")
            oUserTable1 = oApplication.Company.UserTables.Item("Z_PAYROLL1")
            For intRow As Integer = 0 To oTempRec.RecordCount - 1
                strPayrollRefNo = oTempRec.Fields.Item("Code").Value
                strempID = oTempRec.Fields.Item("U_Z_empid").Value
                Dim stEarning As String
                oTemp1.DoQuery("Select * from [@Z_PAYROLL2] where U_Z_RefCode='" & strPayrollRefNo & "'")
                If oTemp1.RecordCount <= 0 Then
                    '  stEarning = "Select 'A' 'Type', 'Basic Salary','Basic Salary',Salary,1.00000,0.00000 from OHEM where empid=" & strempID & " Union"
                    stEarning = ""
                    stEarning = stEarning & " select 'B' 'Type',U_Z_OVTCODE,U_Z_OVTCODE,U_Z_OVTRATE,0.00000,0.00000,U_Z_GLACC from [@Z_PAY_OOVT]  UNION select 'C' 'Type',U_Z_SCODE,U_Z_SCODE,U_Z_SRATE,0.00000,0.00000 ,U_Z_GLACC from [@Z_PAY_OSHT]"
                    stEarning = stEarning & " Union Select 'D' 'Type',T0.[U_Z_CODE],T0.[U_Z_NAME],1,isnull((Select isnull(U_Z_EARN_VALUE,0) from [@Z_PAY1] "
                    stEarning = stEarning & "where U_Z_EARN_TYPE=T0.U_Z_CODE and U_Z_EMPID='" & strempID & "'),0),0.00000,T0.U_Z_EAR_GLACC from [@Z_PAY_OEAR]  T0"
                    otemp2.DoQuery(stEarning)
                    ousertable2 = oApplication.Company.UserTables.Item("Z_PAYROLL2")
                    For intRow1 As Integer = 0 To otemp2.RecordCount - 1
                        oApplication.Utilities.Message("Processing..", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        strCode = oApplication.Utilities.getMaxCode("@Z_PAYROLL2", "Code")
                        ousertable2.Code = strCode
                        ousertable2.Name = strCode & "N"
                        ousertable2.UserFields.Fields.Item("U_Z_RefCode").Value = strPayrollRefNo
                        ousertable2.UserFields.Fields.Item("U_Z_Type").Value = otemp2.Fields.Item(0).Value
                        ousertable2.UserFields.Fields.Item("U_Z_Field").Value = otemp2.Fields.Item(1).Value
                        ousertable2.UserFields.Fields.Item("U_Z_FieldName").Value = otemp2.Fields.Item(2).Value
                        ousertable2.UserFields.Fields.Item("U_Z_Rate").Value = otemp2.Fields.Item(3).Value
                        ousertable2.UserFields.Fields.Item("U_Z_Value").Value = otemp2.Fields.Item(4).Value
                        '  ousertable2.UserFields.Fields.Item("U_Z_Amount").Value = oTempRec.Fields.Item(4).Value
                        ousertable2.UserFields.Fields.Item("U_Z_GLACC").Value = otemp2.Fields.Item(6).Value
                        If ousertable2.Add <> 0 Then
                            oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            If oApplication.Company.InTransaction Then
                                oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                            End If
                            Return False
                        End If
                        otemp2.MoveNext()
                    Next
                End If
                oTempRec.MoveNext()
            Next
            otemp2.DoQuery("Update [@Z_PAYROLL2] set  U_Z_Amount=U_Z_Rate*U_Z_Value")
        End If
        If oApplication.Company.InTransaction Then
            oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
        End If
        Return True
    End Function

    Private Function AddDeduction(ByVal arefCode As String) As Boolean
        Dim oUserTable, oUserTable1, ousertable2, ousertable3 As SAPbobsCOM.UserTable
        Dim strCode, strECode, strEname, strGLacc, strRefCode, strPayrollRefNo, strempID As String
        Dim oTempRec, oTemp1, otemp2, otemp3, otemp4 As SAPbobsCOM.Recordset
        oTempRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otemp2 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otemp3 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otemp4 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        If oApplication.Company.InTransaction Then
            oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
        End If
        oApplication.Company.StartTransaction()
        If 1 = 1 Then
            strRefCode = arefCode
            oTempRec.DoQuery("SELECT * from [@Z_PAYROLL1] where U_Z_RefCode='" & arefCode & "'")
            oUserTable1 = oApplication.Company.UserTables.Item("Z_PAYROLL1")
            For intRow As Integer = 0 To oTempRec.RecordCount - 1
                strPayrollRefNo = oTempRec.Fields.Item("Code").Value
                strempID = oTempRec.Fields.Item("U_Z_empid").Value
                Dim stEarning As String
                oTemp1.DoQuery("Select * from [@Z_PAYROLL3] where U_Z_RefCode='" & strPayrollRefNo & "'")
                If oTemp1.RecordCount <= 0 Then
                    ' stEarning = "select 'A' 'Type',U_Z_OVTCODE,U_Z_OVTRATE,0.00000,0.00000 from [@Z_PAY_OOVT]  UNION select 'B' 'Type',U_Z_SCODE,U_Z_SRATE,0.00000,0.00000 from [@Z_PAY_OSHT]"
                    'stEarning = stEarning & " union Select 'C' 'Type',U_Z_EARN_TYPE,1,U_Z_EARN_VALUE,0.00000 from [@Z_PAY1] where U_Z_EMPID='" & strempID & "'"
                    ' stEarning = "select 'C' 'Type' ,U_Z_DEDUC_TYPE,1,U_Z_DEDUC_VALUE,0.00000 from  [@Z_PAY2] where U_Z_EMPID='" & strempID & "'"

                    stEarning = "Select 'C' 'Type',T0.[CODE],T0.[NAME],1,isnull((Select isnull(U_Z_DEDUC_VALUE,0) from [@Z_PAY2] "
                    stEarning = stEarning & " where U_Z_DEDUC_TYPE=T0.CODE and U_Z_EMPID='" & strempID & "'),0),0.00000,U_Z_DED_GLACC from [@Z_PAY_ODED]  T0"


                    otemp2.DoQuery(stEarning)
                    ousertable2 = oApplication.Company.UserTables.Item("Z_PAYROLL3")
                    For intRow1 As Integer = 0 To otemp2.RecordCount - 1
                        oApplication.Utilities.Message("Processing..", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        strCode = oApplication.Utilities.getMaxCode("@Z_PAYROLL3", "Code")
                        ousertable2.Code = strCode
                        ousertable2.Name = strCode & "N"
                        ousertable2.UserFields.Fields.Item("U_Z_RefCode").Value = strPayrollRefNo
                        ousertable2.UserFields.Fields.Item("U_Z_Type").Value = otemp2.Fields.Item(0).Value
                        ousertable2.UserFields.Fields.Item("U_Z_Field").Value = otemp2.Fields.Item(1).Value
                        ousertable2.UserFields.Fields.Item("U_Z_FieldName").Value = otemp2.Fields.Item(2).Value
                        ousertable2.UserFields.Fields.Item("U_Z_Rate").Value = otemp2.Fields.Item(3).Value
                        ousertable2.UserFields.Fields.Item("U_Z_Value").Value = otemp2.Fields.Item(4).Value
                        ousertable2.UserFields.Fields.Item("U_Z_GLACC").Value = otemp2.Fields.Item(6).Value
                        If ousertable2.Add <> 0 Then
                            oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            If oApplication.Company.InTransaction Then
                                oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                            End If
                            Return False
                        End If
                        otemp2.MoveNext()
                    Next
                End If
                oTempRec.MoveNext()
            Next
            otemp2.DoQuery("Update [@Z_PAYROLL3] set  U_Z_Amount=U_Z_Rate*U_Z_Value")
        End If
        If oApplication.Company.InTransaction Then
            oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
        End If
        Return True
    End Function

    Private Function AddContribution(ByVal arefCode As String) As Boolean
        Dim oUserTable, oUserTable1, ousertable2, ousertable3 As SAPbobsCOM.UserTable
        Dim strCode, strECode, strEname, strGLacc, strRefCode, strPayrollRefNo, strempID As String
        Dim oTempRec, oTemp1, otemp2, otemp3, otemp4 As SAPbobsCOM.Recordset
        oTempRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otemp2 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otemp3 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otemp4 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        If oApplication.Company.InTransaction Then
            oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
        End If
        oApplication.Company.StartTransaction()
        If 1 = 1 Then
            strRefCode = arefCode
            oTempRec.DoQuery("SELECT * from [@Z_PAYROLL1] where U_Z_RefCode='" & arefCode & "'")
            For intRow As Integer = 0 To oTempRec.RecordCount - 1
                strPayrollRefNo = oTempRec.Fields.Item("Code").Value
                strempID = oTempRec.Fields.Item("U_Z_empid").Value
                Dim stEarning As String
                oTemp1.DoQuery("Select * from [@Z_PAYROLL4] where U_Z_RefCode='" & strPayrollRefNo & "'")
                If oTemp1.RecordCount <= 0 Then
                    ' stEarning = "select 'A' 'Type',U_Z_OVTCODE,U_Z_OVTRATE,0.00000,0.00000 from [@Z_PAY_OOVT]  UNION select 'B' 'Type',U_Z_SCODE,U_Z_SRATE,0.00000,0.00000 from [@Z_PAY_OSHT]"
                    'stEarning = stEarning & " union Select 'C' 'Type',U_Z_EARN_TYPE,1,U_Z_EARN_VALUE,0.00000 from [@Z_PAY1] where U_Z_EMPID='" & strempID & "'"
                    'stEarning = "select 'C' 'Type' ,U_Z_CONTR_TYPE,1,U_Z_CONTR_VALUE,0.00000 from  [@Z_PAY3] where U_Z_EMPID='" & strempID & "'"
                    stEarning = "Select 'C' 'Type',T0.[CODE],T0.[NAME],1,isnull((Select isnull(U_Z_CONTR_VALUE,0) from [@Z_PAY3] "
                    stEarning = stEarning & " where U_Z_CONTR_TYPE=T0.CODE and U_Z_EMPID='" & strempID & "'),0),0.00000,U_Z_CON_GLACC from [@Z_PAY_OCON]  T0"
                    otemp2.DoQuery(stEarning)
                    ousertable2 = oApplication.Company.UserTables.Item("Z_PAYROLL4")
                    For intRow1 As Integer = 0 To otemp2.RecordCount - 1
                        strCode = oApplication.Utilities.getMaxCode("@Z_PAYROLL4", "Code")
                        ousertable2.Code = strCode
                        ousertable2.Name = strCode & "N"
                        ousertable2.UserFields.Fields.Item("U_Z_RefCode").Value = strPayrollRefNo
                        ousertable2.UserFields.Fields.Item("U_Z_Type").Value = otemp2.Fields.Item(0).Value
                        ousertable2.UserFields.Fields.Item("U_Z_Field").Value = otemp2.Fields.Item(1).Value
                        ousertable2.UserFields.Fields.Item("U_Z_FieldName").Value = otemp2.Fields.Item(2).Value
                        ousertable2.UserFields.Fields.Item("U_Z_Rate").Value = otemp2.Fields.Item(3).Value
                        ousertable2.UserFields.Fields.Item("U_Z_Value").Value = otemp2.Fields.Item(4).Value
                        ousertable2.UserFields.Fields.Item("U_Z_GLACC").Value = otemp2.Fields.Item(6).Value
                        If ousertable2.Add <> 0 Then
                            oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            If oApplication.Company.InTransaction Then
                                oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                            End If
                            Return False
                        End If
                        otemp2.MoveNext()
                    Next
                End If
                oTempRec.MoveNext()
            Next
            otemp2.DoQuery("Update [@Z_PAYROLL4] set  U_Z_Amount=U_Z_Rate*U_Z_Value")
        End If
        If oApplication.Company.InTransaction Then
            oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
        End If
        Return True
    End Function

    Private Function AddPayRollMaster(ByVal aYear As Integer, ByVal aMonth As Integer) As Boolean
        Dim oUserTable, oUserTable1, ousertable2, ousertable3 As SAPbobsCOM.UserTable
        Dim strCode, strECode, strEname, strGLacc, strRefCode, strPayrollRefNo, strempID As String
        Dim oTempRec, oTemp1, otemp2, otemp3, otemp4 As SAPbobsCOM.Recordset
        oTempRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otemp2 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otemp3 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otemp4 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        If oApplication.Company.InTransaction Then
            oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
        End If
        oApplication.Company.StartTransaction()
        oUserTable = oApplication.Company.UserTables.Item("Z_PAYROLL")
        strCode = oApplication.Utilities.getMaxCode("@Z_PAYROLL", "Code")
        oUserTable.Code = strCode
        oUserTable.Name = strCode & "N"
        oUserTable.UserFields.Fields.Item("U_Z_YEAR").Value = aYear
        oUserTable.UserFields.Fields.Item("U_Z_MONTH").Value = aMonth
        oUserTable.UserFields.Fields.Item("U_Z_Process").Value = "N"
        If oUserTable.Add <> 0 Then
            oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            If oApplication.Company.InTransaction Then
                oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If
            Return False
        Else
            strRefCode = strCode
            oTempRec.DoQuery("SELECT T0.[empID], T0.[firstName]+T0.[LastName] 'Emplopyee name', T0.[jobTitle],T1.[Name], T0.[salary], T0.[salaryUnit], T2.[PrcName] FROM OHEM T0  INNER JOIN OUDP T1 ON T0.dept = T1.Code INNER JOIN OPRC T2 ON T0.U_Z_COST = T2.PrcCode")
            oUserTable1 = oApplication.Company.UserTables.Item("Z_PAYROLL1")
            For intRow As Integer = 0 To oTempRec.RecordCount - 1
                strCode = oApplication.Utilities.getMaxCode("@Z_PAYROLL1", "Code")
                oUserTable1.Code = strCode
                oUserTable1.Name = strCode & "N"
                strempID = oTempRec.Fields.Item(0).Value
                oUserTable1.UserFields.Fields.Item("U_Z_RefCode").Value = strRefCode
                oUserTable1.UserFields.Fields.Item("U_Z_empid").Value = oTempRec.Fields.Item(0).Value
                oUserTable1.UserFields.Fields.Item("U_Z_EmpName").Value = oTempRec.Fields.Item(1).Value
                oUserTable1.UserFields.Fields.Item("U_Z_JobTitle").Value = oTempRec.Fields.Item(2).Value
                oUserTable1.UserFields.Fields.Item("U_Z_Department").Value = oTempRec.Fields.Item(3).Value
                oUserTable1.UserFields.Fields.Item("U_Z_BasicSalary").Value = oTempRec.Fields.Item(4).Value
                oUserTable1.UserFields.Fields.Item("U_Z_SalaryType").Value = oTempRec.Fields.Item(5).Value
                oUserTable1.UserFields.Fields.Item("U_Z_CostCentre").Value = oTempRec.Fields.Item(6).Value
                If oUserTable1.Add <> 0 Then
                    oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    If oApplication.Company.InTransaction Then
                        oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    End If
                    Return False
                Else
                    strPayrollRefNo = strCode
                    Dim stEarning As String
                    stEarning = "select 'A' 'Type',U_Z_OVTCODE,U_Z_OVTRATE,0.00000,0.00000 from [@Z_PAY_OOVT]  UNION select 'B' 'Type',U_Z_SCODE,U_Z_SRATE,0.00000,0.00000 from [@Z_PAY_OSHT]"
                    stEarning = stEarning & " union Select 'C' 'Type',U_Z_EARN_TYPE,1,U_Z_EARN_VALUE,0.00000 from [@Z_PAY1] where U_Z_EMPID='" & strempID & "'"
                    otemp2.DoQuery(stEarning)
                    ousertable2 = oApplication.Company.UserTables.Item("Z_PAYROLL2")
                    For intRow1 As Integer = 0 To otemp2.RecordCount - 1
                        strCode = oApplication.Utilities.getMaxCode("@Z_PAYROLL2", "Code")
                        ousertable2.Code = strCode
                        ousertable2.Name = strCode & "N"
                        'strempID = oTempRec.Fields.Item(0).Value
                        ousertable2.UserFields.Fields.Item("U_Z_RefCode").Value = strPayrollRefNo
                        ousertable2.UserFields.Fields.Item("U_Z_Type").Value = oTempRec.Fields.Item(0).Value
                        ousertable2.UserFields.Fields.Item("U_Z_Field").Value = oTempRec.Fields.Item(1).Value
                        ousertable2.UserFields.Fields.Item("U_Z_Rate").Value = oTempRec.Fields.Item(2).Value
                        ousertable2.UserFields.Fields.Item("U_Z_Value").Value = oTempRec.Fields.Item(3).Value
                        ousertable2.UserFields.Fields.Item("U_Z_Amount").Value = oTempRec.Fields.Item(4).Value
                        If ousertable2.Add <> 0 Then
                            oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            If oApplication.Company.InTransaction Then
                                oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                            End If
                            Return False
                        End If
                        otemp2.MoveNext()
                    Next
                End If
                oTempRec.MoveNext()
            Next
        End If
        If oApplication.Company.InTransaction Then
            oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
        End If

        'oApplication.Utilities.Message("Operation completed successfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)

        Return True
    End Function
#End Region

#Region "Remove Row"
    Private Sub RemoveRow(ByVal intRow As Integer, ByVal agrid As SAPbouiCOM.Grid)
        Dim strCode, strname As String
        Dim otemprec As SAPbobsCOM.Recordset
        For intRow = 0 To agrid.DataTable.Rows.Count - 1
            If agrid.Rows.IsSelected(intRow) Then
                strCode = agrid.DataTable.GetValue(0, intRow)
                strname = agrid.DataTable.GetValue(1, intRow)
                otemprec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                'otemprec.DoQuery("Select * from [@Z_PAY_ODED] where Code='" & strCode & "' and Name='" & strname & "'")
                'If otemprec.RecordCount > 0 And strCode <> "" Then
                '    oApplication.Utilities.Message("Transaction already exists. Can not delete the Bin Details.", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                '    Exit Sub
                'End If
                'oApplication.Utilities.ExecuteSQL(oTemp, "update [@Z_PAY_ODED] set  Name =Name +'D'  where Code='" & strCode & "'")
                agrid.DataTable.Rows.Remove(intRow)
                Exit Sub
            End If
        Next
        oApplication.Utilities.Message("No row selected", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
    End Sub
#End Region

#Region "Validate Grid details"
    Private Function validation(ByVal aGrid As SAPbouiCOM.Grid) As Boolean
        Dim strECode, strECode1, strEname, strEname1 As String
        For intRow As Integer = 0 To aGrid.DataTable.Rows.Count - 1
            strECode = aGrid.DataTable.GetValue(0, intRow)
            strEname = aGrid.DataTable.GetValue(1, intRow)
            If strECode = "" And strEname <> "" Then
                oApplication.Utilities.Message("Code is missing . Code : " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            If strECode <> "" And strEname = "" Then
                oApplication.Utilities.Message("Name is missing . Code : " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            For intInnerLoop As Integer = intRow To aGrid.DataTable.Rows.Count - 1
                strECode1 = aGrid.DataTable.GetValue(0, intInnerLoop)
                strEname1 = aGrid.DataTable.GetValue(1, intInnerLoop)
                If strECode = strECode1 And strEname = strEname1 And intRow <> intInnerLoop Then
                    oApplication.Utilities.Message("This Code and Name combination is already exists. Code no : " & intInnerLoop, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
            Next
        Next
        Return True
    End Function

#End Region

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_PayrollGeneration Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "10" And pVal.ColUID = "U_Z_JVNo" Then
                                    oGrid = oForm.Items.Item("10").Specific
                                    Dim strCmp As String = oGrid.DataTable.GetValue("U_Z_CompNo", pVal.Row)
                                    Dim oTest As SAPbobsCOM.Recordset
                                    oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oTest.DoQuery("Select isnull(U_Z_JVType,'V') from [@Z_OADM] where U_Z_CompCode='" & strCmp & "'")
                                    If oTest.Fields.Item(0).Value = "V" Then
                                        oEditTextColumn = oGrid.Columns.Item("U_Z_JVNo")
                                        oEditTextColumn.LinkedObjectType = "28"
                                    Else
                                        oEditTextColumn = oGrid.Columns.Item("U_Z_JVNo")
                                        oEditTextColumn.LinkedObjectType = "30"
                                    End If

                                End If
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "3" Then
                                    If oForm.PaneLevel = 2 Then
                                        If PrepareWorkSheet(oForm) = False Then
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                    End If
                                End If
                                If pVal.ItemUID = "5" Then
                                    Dim intYear, intMonth As Integer
                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                    If oApplication.SBO_Application.MessageBox("Do you want to  Generate Payroll for selected Month and year?", , "Yes", "No") = 1 Then
                                        oCombobox = oForm.Items.Item("7").Specific
                                        If oCombobox.Selected.Value = "" Then
                                            'oApplication.Utilities.Message("Select year", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            intYear = 0
                                        Else
                                            intYear = oCombobox.Selected.Value
                                            If intYear = 0 Then
                                                ' oApplication.Utilities.Message("Select year", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            End If
                                        End If
                                        oCombobox = oForm.Items.Item("9").Specific
                                        If oCombobox.Selected.Value = "" Then
                                            intMonth = 0
                                        Else
                                            intMonth = oCombobox.Selected.Value
                                        End If
                                        If ResetPayrollWorksheet(intYear, intMonth, oForm) = False Then
                                            BubbleEvent = False
                                            Exit Sub

                                        End If
                                    End If
                                End If
                        End Select

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                ' oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                            Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "10" And pVal.ColUID <> "RowsHeader" Then
                                    Dim strCode As String
                                    Dim intYear, intMonth As Integer
                                    oCombobox = oForm.Items.Item("7").Specific
                                    If oCombobox.Selected.Value = "" Then
                                        oApplication.Utilities.Message("Select year", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    Else
                                        intYear = oCombobox.Selected.Value
                                        If intYear = 0 Then
                                            oApplication.Utilities.Message("Select year", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        End If
                                    End If
                                    oCombobox = oForm.Items.Item("9").Specific
                                    If oCombobox.Selected.Value = "" Then
                                        oApplication.Utilities.Message("Select Month", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    Else
                                        intMonth = oCombobox.Selected.Value
                                        If intMonth = 0 Then
                                            oApplication.Utilities.Message("Select Month", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        End If
                                    End If
                                    oGrid.Columns.Item("RowsHeader").Click(pVal.Row)
                                    For intRow As Integer = pVal.Row To pVal.Row
                                        If oGrid.Rows.IsSelected(pVal.Row) Then
                                            strCode = oGrid.DataTable.GetValue("Code", intRow)
                                            If strCode <> "" Then
                                                Dim oOBj As New clsPayrolLDetails
                                                frmSourceForm = oForm
                                                oOBj.LoadForm(intMonth, intYear, strCode, "WorkSheet")
                                                Exit Sub
                                            End If
                                        End If
                                    Next
                                End If


                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Select Case pVal.ItemUID
                                    Case "4"
                                        oForm.PaneLevel = oForm.PaneLevel - 1
                                    Case "3"
                                        oForm.PaneLevel = oForm.PaneLevel + 1
                                    Case "5"
                                        oApplication.Utilities.Message("Payroll worksheet generation completed successfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                        'GenerateWorkSheet(oForm)
                                        PrepareWorkSheet(oForm)
                                        'oForm.Close()
                                    Case "13"
                                        LoadPayRollDetails(oForm)
                                    Case "11"
                                        oGrid = oForm.Items.Item("10").Specific
                                        AddEmptyRow(oGrid)
                                    Case "12"
                                        oGrid = oForm.Items.Item("10").Specific
                                        RemoveRow(1, oGrid)
                                    Case "14"
                                        If GenerateWorkSheet(oForm) = False Then
                                            Exit Sub
                                        End If


                                End Select

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
                Case mnu_PayrollPrinting
                    If pVal.BeforeAction = False Then
                        LoadForm()
                    End If

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
