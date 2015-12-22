Imports System.IO
Public Class clsPayApproval
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
    Private InvBaseDocNo, strQuery As String
    Private InvForConsumedItems As Integer
    Private blnFlag As Boolean = False
    Dim oRec As SAPbobsCOM.Recordset
    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub
    Public Sub LoadForm(ByVal aChoice As String)
        Try
            oForm = oApplication.Utilities.LoadForm(xml_Pay_Approval, frm_Pay_Approval)
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Freeze(True)
            If aChoice = "R" Then
                oForm.Title = "Payroll Worksheet Approval"
            Else
                oForm.Title = "Payroll Offcycle worksheet Approval"
            End If
            HeaderGridBind(oForm, aChoice)
            HeaderSumGridBind(oForm, aChoice)
            oForm.PaneLevel = 1
            oForm.Items.Item("5").TextStyle = 7
            oForm.Items.Item("1000001").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            oGrid = oForm.Items.Item("4").Specific
            oGrid.Columns.Item("RowsHeader").Click(0, False, False)
            oGrid = oForm.Items.Item("9").Specific
            oGrid.Columns.Item("RowsHeader").Click(0, False, False)
            oForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
            oForm.Freeze(False)
        End Try
    End Sub
    Private Sub HeaderGridBind(ByVal aForm As SAPbouiCOM.Form, ByVal aChoice As String)
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oGrid = aForm.Items.Item("4").Specific
        Try
            Select Case aChoice
                Case "R"
                    strQuery = "Select Code,Name,T0.U_Z_CompNo,T0.U_Z_YEAR,T0.U_Z_MONTH,T0.U_Z_RefCode,isnull(T0.U_Z_AppStatus,'P') as U_Z_AppStatus,T1.U_Z_Remarks,T0.U_Z_FileName,T1.U_Z_Attachment,T0.U_Z_CurApprover,T0.U_Z_NxtApprover, "
                    '  strQuery += "T0.U_Z_AppReqDate,T0.U_Z_DocType,T0.U_Z_Creater from [@Z_PAY_Approval] T0 Left Outer Join [@Z_PAY_APHIS] T1 on T0.Code=T1.U_Z_DocEntry and T1.U_Z_ApproveBy='" & oApplication.Company.UserName & "' where T0.U_Z_AppStatus='P' and T0.U_Z_DocType='R' and ( U_Z_CurApprover='" & oApplication.Company.UserName & "' OR U_Z_NxtApprover='" & oApplication.Company.UserName & "')"
                    strQuery += "T0.U_Z_AppReqDate,T0.U_Z_DocType,T0.U_Z_Creater from [@Z_PAY_Approval] T0 Left Outer Join [@Z_PAY_APHIS] T1 on T0.Code=T1.U_Z_DocEntry and T1.U_Z_ApproveBy='" & oApplication.Company.UserName & "' where T0.U_Z_AppStatus='P' and T0.U_Z_DocType='R' and ( U_Z_NxtApprover='" & oApplication.Company.UserName & "')"
            End Select
            oGrid.DataTable.ExecuteQuery(strQuery)
            FormatHeadGrid(aForm, "Header", "4")
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub HeaderSumGridBind(ByVal aForm As SAPbouiCOM.Form, ByVal aChoice As String)
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oGrid = aForm.Items.Item("9").Specific
        Try
            Select Case aChoice
                Case "R"
                    strQuery = "Select Code,Name,T0.U_Z_CompNo,T0.U_Z_YEAR,T0.U_Z_MONTH,T0.U_Z_RefCode,isnull(T0.U_Z_AppStatus,'P') as U_Z_AppStatus,T1.U_Z_Remarks,T0.U_Z_FileName,T1.U_Z_Attachment,T0.U_Z_CurApprover,T0.U_Z_NxtApprover, "
                    '  strQuery += "T0.U_Z_AppReqDate,T0.U_Z_DocType,T0.U_Z_Creater from [@Z_PAY_Approval] T0 Left Outer Join [@Z_PAY_APHIS] T1 on T0.Code=T1.U_Z_DocEntry and T1.U_Z_ApproveBy='" & oApplication.Company.UserName & "' where T0.U_Z_DocType='R' and ( U_Z_CurApprover='" & oApplication.Company.UserName & "' OR U_Z_NxtApprover='" & oApplication.Company.UserName & "')"
                    strQuery += "T0.U_Z_AppReqDate,T0.U_Z_DocType,T0.U_Z_Creater from [@Z_PAY_Approval] T0 Left Outer Join [@Z_PAY_APHIS] T1 on T0.Code=T1.U_Z_DocEntry  where T0.U_Z_DocType='R' and T1.U_Z_ApproveBy='" & oApplication.Company.UserName & "'" ' and ( U_Z_CurApprover='" & oApplication.Company.UserName & "' OR U_Z_NxtApprover='" & oApplication.Company.UserName & "')"
                    ' strQuery += "T0.U_Z_AppReqDate,T0.U_Z_DocType,T0.U_Z_Creater from [@Z_PAY_Approval] T0 Left Outer Join [@Z_PAY_APHIS] T1 on T0.Code=T1.U_Z_DocEntry and T1.U_Z_ApproveBy='" & oApplication.Company.UserName & "' where T0.U_Z_DocType='R'" ' and ( U_Z_CurApprover='" & oApplication.Company.UserName & "' OR U_Z_NxtApprover='" & oApplication.Company.UserName & "')"
            End Select
            oGrid.DataTable.ExecuteQuery(strQuery)
            FormatHeadGrid(aForm, "Header", "9")
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub FormatHeadGrid(ByVal aForm As SAPbouiCOM.Form, ByVal aChoice As String, ByVal gridId As String)
        Dim oGridCombo As SAPbouiCOM.ComboBoxColumn
        Try
            If aChoice = "Header" Then
                oGrid = aForm.Items.Item(gridId).Specific
                oGrid.Columns.Item("Code").TitleObject.Caption = "Code"
                oGrid.Columns.Item("Code").Editable = False
                oEditTextColumn = oGrid.Columns.Item("Code")
                oEditTextColumn.LinkedObjectType = "Z_PAY_OAPPT"
                oGrid.Columns.Item("Name").Visible = False
                oGrid.Columns.Item("U_Z_CompNo").TitleObject.Caption = "Company Code"
                oGrid.Columns.Item("U_Z_CompNo").Editable = False
                oGrid.Columns.Item("U_Z_YEAR").TitleObject.Caption = "Payroll Year"
                oGrid.Columns.Item("U_Z_YEAR").Editable = False
                oGrid.Columns.Item("U_Z_MONTH").TitleObject.Caption = "Payroll Month"
                oGrid.Columns.Item("U_Z_MONTH").Editable = False
                oGrid.Columns.Item("U_Z_RefCode").TitleObject.Caption = "Reference Code"
                oGrid.Columns.Item("U_Z_RefCode").Editable = False
                oGrid.Columns.Item("U_Z_AppStatus").TitleObject.Caption = "Worksheet Status"
                oGrid.Columns.Item("U_Z_AppStatus").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
                oGridCombo = oGrid.Columns.Item("U_Z_AppStatus")
                oGridCombo.ValidValues.Add("P", "Pending")
                oGridCombo.ValidValues.Add("A", "Approved")
                oGridCombo.ValidValues.Add("R", "Rejected")
                oGridCombo.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_both
                oGrid.Columns.Item("U_Z_Remarks").TitleObject.Caption = "Remarks"
                oGrid.Columns.Item("U_Z_Remarks").Editable = True

                oGrid.Columns.Item("U_Z_FileName").TitleObject.Caption = "Worksheet Attachment"
                oGrid.Columns.Item("U_Z_FileName").Editable = False
                oEditTextColumn = oGrid.Columns.Item("U_Z_FileName")
                oEditTextColumn.LinkedObjectType = "Z_PAY_APHIS"

                oGrid.Columns.Item("U_Z_Attachment").TitleObject.Caption = "Approver Attachment"
                oGrid.Columns.Item("U_Z_Attachment").Editable = False
                oEditTextColumn = oGrid.Columns.Item("U_Z_Attachment")
                oEditTextColumn.LinkedObjectType = "Z_PAY_APHIS"
                oGrid.Columns.Item("U_Z_CurApprover").TitleObject.Caption = "Current Approver"
                oGrid.Columns.Item("U_Z_CurApprover").Editable = False
                oGrid.Columns.Item("U_Z_NxtApprover").TitleObject.Caption = "Next Approver"
                oGrid.Columns.Item("U_Z_NxtApprover").Editable = False
                oGrid.Columns.Item("U_Z_AppReqDate").Visible = False
                oGrid.Columns.Item("U_Z_DocType").Visible = False
                oGrid.Columns.Item("U_Z_Creater").Visible = False
                oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
                oGrid.AutoResizeColumns()
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub ViewWorkSheet(ByVal sForm As SAPbouiCOM.Form, ByVal RefCode As String, ByVal gridId As String)
        oGrid = sForm.Items.Item(gridId).Specific
        Try
            strSQL = "SELECT T0.[Code], T0.[Name],T0.[U_Z_TANO] 'TANO',T0.[U_Z_empid], T0.[U_Z_EmpName], Case T0.""U_Z_OnHold"" when 'H' then 'On Hold' else 'Active' end ""Status"", T0.[U_Z_Basic], T0.[U_Z_InrAmt], T0.[U_Z_BasicSalary], T0.[U_Z_MonthlyBasic],T0.[U_Z_Cost], T0.[U_Z_NetSalary], isnull(T0.U_Z_MonthlyBasic,0) + isnull(T0.U_Z_Earning,0)  'GrossSalary',T0.U_Z_WorkingDays1,T0.[U_Z_Earning], T0.[U_Z_Deduction], T0.[U_Z_UnPaidLeave], T0.[U_Z_PaidLeave], T0.[U_Z_AnuLeave],T0.""U_Z_CashOutAmt"", T0.[U_Z_Contri], T0.[U_Z_AirAmt], ""U_Z_NetPayAmt"",""U_Z_CmpPayAmt"", T0.[U_Z_AcrAmt] ,T0.[U_Z_AcrAirAmt], T0.[U_Z_EOSYTD] ,T0.[U_Z_EOSBalance],T0.[U_Z_EOS],T0.[U_Z_RefCode], T0.[U_Z_PersonalID],  T0.[U_Z_JobTitle], T0.[U_Z_Department],T0.[U_Z_EmpBranch], T0.[U_Z_TermName] 'Contract Term', T0.[U_Z_SalaryType], T0.[U_Z_CostCentre],  T0.[U_Z_Startdate], T0.[U_Z_TermDate], T0.[U_Z_JVNo],  T0.[U_Z_CompNo], T0.[U_Z_Branch], T0.[U_Z_Dept],T0.""U_Z_EOS1"",T0.""U_Z_Leave"",T0.""U_Z_Ticket"",T0.""U_Z_Saving"",T0.""U_Z_PaidExtraSalary"",T0.""U_Z_GOVAMT"" 'Social Gov.Amt' FROM [dbo].[@Z_PAYROLL1]  T0 where T0.U_Z_RefCode='" & RefCode & "'"
            strSQL = "SELECT T0.[Code], T0.[Name],T0.[U_Z_TANO] 'TANO',T0.[U_Z_empid], T0.[U_Z_ExtNo] 'Batch No', T0.[U_Z_EmpName],T0.[U_Z_Country] 'Country', Case T0.""U_Z_OnHold"" when 'H' then 'On Hold' else 'Active' end ""Status"", T0.[U_Z_Basic], T0.[U_Z_InrAmt], T0.[U_Z_BasicSalary], T0.[U_Z_MonthlyBasic],T0.[U_Z_ActWork] 'Total Worked days', T0.[U_Z_Cost], isnull(T0.U_Z_MonthlyBasic,0) + isnull(T0.U_Z_Earning,0)  'GrossSalary',T0.[U_Z_Earning], T0.[U_Z_Deduction], T0.[U_Z_NetSalary], T0.[U_Z_UnPaidLeave], T0.[U_Z_PaidLeave], T0.[U_Z_AnuLeave],T0.""U_Z_CashOutAmt"", T0.[U_Z_Contri], T0.[U_Z_AirAmt], ""U_Z_NetPayAmt"",""U_Z_CmpPayAmt"", T0.[U_Z_AcrAmt] ,T0.[U_Z_AcrAirAmt], T0.[U_Z_EOSYTD] ,T0.[U_Z_EOSBalance],T0.[U_Z_EOS],T0.U_Z_WorkingDays1, T0.[U_Z_CalenderDays] 'Working Days of month',T0.[U_Z_TotalLeave] 'Leave Utilized ',T0.[U_Z_RefCode], T0.[U_Z_PersonalID],  T0.[U_Z_JobTitle], T0.[U_Z_Department],T0.[U_Z_EmpBranch], T0.[U_Z_TermName] 'Contract Term', T0.[U_Z_SalaryType], T0.[U_Z_CostCentre],  T0.[U_Z_Startdate], T0.[U_Z_TermDate], T0.[U_Z_JVNo],  T0.[U_Z_CompNo], T0.[U_Z_Branch], T0.[U_Z_Dept],T0.""U_Z_EOS1"",T0.""U_Z_Leave"",T0.""U_Z_Ticket"",T0.""U_Z_Saving"",T0.""U_Z_PaidExtraSalary"",T0.""U_Z_GOVAMT"" 'Social Gov.Amt' FROM [dbo].[@Z_PAYROLL1]  T0 where T0.U_Z_RefCode='" & RefCode & "'"

            oGrid.DataTable.ExecuteQuery(strSQL)
            Formatgrid(oGrid, "Payroll")
            oApplication.Utilities.assignMatrixLineno(oGrid, sForm)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#Region "FormatGrid"
    Private Sub Formatgrid(ByVal agrid As SAPbouiCOM.Grid, ByVal aOption As String)
        Select Case aOption
            Case "Payroll"
                agrid.Columns.Item("Code").TitleObject.Caption = "Code"
                agrid.Columns.Item("Name").TitleObject.Caption = "Name"
                agrid.Columns.Item("Code").Visible = False
                agrid.Columns.Item("Name").Visible = False
                agrid.Columns.Item("TANO").TitleObject.Caption = "T & A Employee No"
                agrid.Columns.Item("U_Z_RefCode").TitleObject.Caption = "Reference Code"
                agrid.Columns.Item("U_Z_RefCode").Visible = False
                agrid.Columns.Item("U_Z_empid").TitleObject.Caption = "Employee ID"
                agrid.Columns.Item("U_Z_EmpName").TitleObject.Caption = "Employee Name"
                agrid.Columns.Item("U_Z_JobTitle").TitleObject.Caption = "Job Title"
                agrid.Columns.Item("U_Z_Department").TitleObject.Caption = "Department"
                agrid.Columns.Item("U_Z_EmpBranch").TitleObject.Caption = "Emp.Branch"
                agrid.Columns.Item("U_Z_BasicSalary").TitleObject.Caption = " Total Basic Salary"
                oEditTextColumn = oGrid.Columns.Item("U_Z_BasicSalary")
                oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                agrid.Columns.Item("U_Z_MonthlyBasic").TitleObject.Caption = "Current Month Baisc"
                oEditTextColumn = oGrid.Columns.Item("U_Z_MonthlyBasic")
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

                agrid.Columns.Item("U_Z_PaidLeave").TitleObject.Caption = "Paid Leave"
                oEditTextColumn = oGrid.Columns.Item("U_Z_PaidLeave")
                oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto

                agrid.Columns.Item("U_Z_Contri").TitleObject.Caption = "Contribution"
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
                agrid.Columns.Item("U_Z_EOS").TitleObject.Caption = "EOS Current Month Accural"
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
                agrid.Columns.Item("U_Z_AnuLeave").TitleObject.Caption = "Annual Leave"
                agrid.Columns.Item("U_Z_AnuLeave").Editable = False
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
                oGrid.Columns.Item("U_Z_CmpPayAmt").TitleObject.Caption = "AirTicket CosttoCompany Amount"
                oGrid.Columns.Item("U_Z_NetPayAmt").TitleObject.Caption = "AirTicket NetPay Amount"

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
                oGrid.Columns.Item("U_Z_PaidExtraSalary").TitleObject.Caption = "Include Extra Salary"
                oGrid.Columns.Item("U_Z_PaidExtraSalary").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
                oGrid.Columns.Item("U_Z_PaidExtraSalary").Editable = False

                oGrid.Columns.Item("U_Z_CashOutAmt").TitleObject.Caption = "Leave Cashout Amount"
                oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto

                oGrid.Columns.Item("U_Z_WorkingDays1").TitleObject.Caption = "Basic Salary Calculation Days"
                oGrid.Columns.Item("U_Z_WorkingDays1").Editable = False

                oGrid.Columns.Item("GrossSalary").TitleObject.Caption = "Gross Salary"
                oGrid.Columns.Item("GrossSalary").Editable = False
                oEditTextColumn = oGrid.Columns.Item("GrossSalary")
                oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                'agrid.Columns.Item("Code").TitleObject.Caption = "Code"
                'oEditTextColumn = agrid.Columns.Item("Code")
                'oEditTextColumn.LinkedObjectType = "2"
                'agrid.Columns.Item("Name").TitleObject.Caption = "Name"
                'agrid.Columns.Item("Name").Visible = False
                'agrid.Columns.Item("TANO").TitleObject.Caption = "T & A Employee No"
                'agrid.Columns.Item("U_Z_RefCode").TitleObject.Caption = "Reference Code"
                'agrid.Columns.Item("U_Z_RefCode").Visible = False
                'agrid.Columns.Item("U_Z_empid").TitleObject.Caption = "Employee ID"
                'agrid.Columns.Item("U_Z_EmpName").TitleObject.Caption = "Employee Name"
                'agrid.Columns.Item("U_Z_JobTitle").TitleObject.Caption = "Job Title"
                'agrid.Columns.Item("U_Z_Department").TitleObject.Caption = "Department"
                'agrid.Columns.Item("U_Z_EmpBranch").TitleObject.Caption = "Emp.Branch"
                'agrid.Columns.Item("U_Z_BasicSalary").TitleObject.Caption = " Total Basic Salary"
                'oEditTextColumn = oGrid.Columns.Item("U_Z_BasicSalary")
                'oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                'agrid.Columns.Item("U_Z_MonthlyBasic").TitleObject.Caption = "Current Month Baisc"
                'oEditTextColumn = oGrid.Columns.Item("U_Z_MonthlyBasic")
                'oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                'agrid.Columns.Item("U_Z_SalaryType").TitleObject.Caption = "Salary Type"
                'agrid.Columns.Item("U_Z_CostCentre").TitleObject.Caption = "Cost Center"
                'agrid.Columns.Item("U_Z_Earning").TitleObject.Caption = "Earnings"
                'oEditTextColumn = oGrid.Columns.Item("U_Z_Earning")
                'oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                'agrid.Columns.Item("U_Z_Deduction").TitleObject.Caption = "Deduction"
                'oEditTextColumn = oGrid.Columns.Item("U_Z_Deduction")
                'oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                'agrid.Columns.Item("U_Z_UnPaidLeave").TitleObject.Caption = "UnPaid Leave"
                'oEditTextColumn = oGrid.Columns.Item("U_Z_UnPaidLeave")
                'oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto

                'agrid.Columns.Item("U_Z_PaidLeave").TitleObject.Caption = "Paid Leave"
                'oEditTextColumn = oGrid.Columns.Item("U_Z_PaidLeave")
                'oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto

                'agrid.Columns.Item("U_Z_Contri").TitleObject.Caption = "Contribution"
                'oEditTextColumn = oGrid.Columns.Item("U_Z_Contri")
                'oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                'agrid.Columns.Item("U_Z_Cost").TitleObject.Caption = "Total Cost"
                'oEditTextColumn = oGrid.Columns.Item("U_Z_Cost")
                'oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                'agrid.Columns.Item("U_Z_NetSalary").TitleObject.Caption = "Net Salary"
                'oEditTextColumn = oGrid.Columns.Item("U_Z_NetSalary")
                'oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                'oEditTextColumn = agrid.Columns.Item("U_Z_empid")
                'oEditTextColumn.LinkedObjectType = "171"
                'agrid.Columns.Item("U_Z_Startdate").TitleObject.Caption = "Joining Date"
                'agrid.Columns.Item("U_Z_TermDate").TitleObject.Caption = "Termination Date"
                'agrid.Columns.Item("U_Z_JVNo").TitleObject.Caption = "Journal Voucher Ref"
                'oEditTextColumn = agrid.Columns.Item("U_Z_JVNo")
                'oEditTextColumn.LinkedObjectType = "28"
                'agrid.Columns.Item("U_Z_EOS").TitleObject.Caption = "EOS Current Month Accural"
                'agrid.Columns.Item("U_Z_EOS").Editable = False
                'oEditTextColumn = oGrid.Columns.Item("U_Z_EOS")
                'oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
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
                'agrid.Columns.Item("U_Z_PersonalID").TitleObject.Caption = "Government ID"
                'agrid.Columns.Item("U_Z_PersonalID").Editable = False

                'agrid.Columns.Item("U_Z_Basic").TitleObject.Caption = "Basic Salary"
                'agrid.Columns.Item("U_Z_Basic").Editable = False
                'oEditTextColumn = oGrid.Columns.Item("U_Z_Basic")
                'oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                'agrid.Columns.Item("U_Z_InrAmt").TitleObject.Caption = "Increment Amount"
                'agrid.Columns.Item("U_Z_InrAmt").Editable = False
                'oEditTextColumn = oGrid.Columns.Item("U_Z_InrAmt")
                'oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto

                'agrid.Columns.Item("U_Z_AcrAmt").TitleObject.Caption = "Annual Leave Accural Amount"
                'agrid.Columns.Item("U_Z_AcrAmt").Editable = False
                'oEditTextColumn = oGrid.Columns.Item("U_Z_AcrAmt")
                'oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto

                'agrid.Columns.Item("U_Z_AcrAirAmt").TitleObject.Caption = "AirTicket Accural Amount"
                'agrid.Columns.Item("U_Z_AcrAirAmt").Editable = False
                'oEditTextColumn = oGrid.Columns.Item("U_Z_AcrAirAmt")
                'oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto

                'agrid.Columns.Item("U_Z_EOSYTD").TitleObject.Caption = "EOS YTD"
                'agrid.Columns.Item("U_Z_EOSYTD").Editable = False
                'oEditTextColumn = oGrid.Columns.Item("U_Z_EOSYTD")
                'oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto

                'agrid.Columns.Item("U_Z_EOSBalance").TitleObject.Caption = "Total EOS Accural Balance"
                'agrid.Columns.Item("U_Z_EOSBalance").Editable = False
                'oEditTextColumn = oGrid.Columns.Item("U_Z_EOSBalance")
                'oGrid.Columns.Item("U_Z_CmpPayAmt").TitleObject.Caption = "AirTicket CosttoCompany Amount"
                'oGrid.Columns.Item("U_Z_NetPayAmt").TitleObject.Caption = "AirTicket NetPay Amount"

                'oGrid.Columns.Item("U_Z_EOS1").TitleObject.Caption = "Include EOS Amount"
                'oGrid.Columns.Item("U_Z_EOS1").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
                'oGrid.Columns.Item("U_Z_EOS1").Editable = False

                'oGrid.Columns.Item("U_Z_Leave").TitleObject.Caption = "Include Leave Amount"
                'oGrid.Columns.Item("U_Z_Leave").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
                'oGrid.Columns.Item("U_Z_Leave").Editable = False
                'oGrid.Columns.Item("U_Z_Ticket").TitleObject.Caption = "Include Ticket Amount"
                'oGrid.Columns.Item("U_Z_Ticket").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
                'oGrid.Columns.Item("U_Z_Ticket").Editable = False
                'oGrid.Columns.Item("U_Z_Saving").TitleObject.Caption = "Include Saving Amount"
                'oGrid.Columns.Item("U_Z_Saving").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
                'oGrid.Columns.Item("U_Z_Saving").Editable = False
                'oGrid.Columns.Item("U_Z_PaidExtraSalary").TitleObject.Caption = "Include Extra Salary"
                'oGrid.Columns.Item("U_Z_PaidExtraSalary").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
                'oGrid.Columns.Item("U_Z_PaidExtraSalary").Editable = False

                'oGrid.Columns.Item("U_Z_CashOutAmt").TitleObject.Caption = "Leave Cashout Amount"
                'oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto

                'oGrid.Columns.Item("U_Z_WorkingDays1").TitleObject.Caption = "Worked Days"
                'oGrid.Columns.Item("U_Z_WorkingDays1").Editable = False

                'oGrid.Columns.Item("GrossSalary").TitleObject.Caption = "Gross Salary"
                'oGrid.Columns.Item("GrossSalary").Editable = False
                'oEditTextColumn = oGrid.Columns.Item("GrossSalary")
                'oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
        End Select

        agrid.AutoResizeColumns()
        agrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
    End Sub
#End Region
    Private Sub CopyAttachment(ByVal Sfile As String)
        Try
            oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim strQry = "Select AttachPath From OADP"
            oRec.DoQuery(strQry)
            Dim SPath As String = Sfile
            If SPath = "" Then
            Else
                Dim DPath As String = ""
                If Not oRec.EoF Then
                    DPath = oRec.Fields.Item("AttachPath").Value.ToString()
                End If
                If Not Directory.Exists(DPath) Then
                    Directory.CreateDirectory(DPath)
                End If
                Dim file = New FileInfo(SPath)
                Dim Filename As String = Path.GetFileName(SPath)
                Dim SavePath As String = Path.Combine(DPath, Filename)
                If System.IO.File.Exists(SavePath) Then
                Else
                    file.CopyTo(Path.Combine(DPath, file.Name), True)
                End If
            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

#Region "Approval Functions"
    Public Sub addUpdateDocument(ByVal aForm As SAPbouiCOM.Form)
        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCompanyService As SAPbobsCOM.CompanyService
        Dim oChildren As SAPbobsCOM.GeneralDataCollection
        oCompanyService = oApplication.Company.GetCompanyService()
        Dim otestRs As SAPbobsCOM.Recordset
        Dim oChild As SAPbobsCOM.GeneralData
        Dim strCode, strQuery As String
        Dim strEmpName As String = ""
        Dim blnRecordExists As Boolean = False
        Dim HeadDocEntry, UserLineId As Integer
        Dim oRecordSet As SAPbobsCOM.Recordset
        Dim oComboBox1, oCombobox2 As SAPbouiCOM.ComboBox
        Try
            If oApplication.SBO_Application.MessageBox("Documents once approved can not be changed. Do you want Continue?", , "Contine", "Cancel") = 2 Then
                Exit Sub
            End If
            oGeneralService = oCompanyService.GetGeneralService("Z_PAY_APHIS")
            oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
            oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            otestRs = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oGrid = aForm.Items.Item("4").Specific
            Dim strDocEntry As String = ""
            Dim strDocType1, HeaderCode, strRefCode, strComp, strFile As String
            Dim strMonth, strYear As Integer
            Dim strEmpID As String = ""
            Dim strLeaveType As String = ""
            If oGrid.DataTable.Rows.Count > 0 Then
                For index As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                    strDocEntry = oGrid.DataTable.GetValue("Code", index)
                    strRefCode = oGrid.DataTable.GetValue("U_Z_RefCode", index)
                    strEmpID = oGrid.DataTable.GetValue("U_Z_Creater", index)
                    strMonth = oGrid.DataTable.GetValue("U_Z_MONTH", index)
                    strYear = oGrid.DataTable.GetValue("U_Z_YEAR", index)
                    strComp = oGrid.DataTable.GetValue("U_Z_CompNo", index)
                    strFile = oGrid.DataTable.GetValue("U_Z_Attachment", index)

                    strQuery = "select T0.DocEntry,T1.LineId from [@Z_PAY_OAPPT] T0 JOIN [@Z_PAY_APPT2] T1 on T0.DocEntry=T1.DocEntry"
                    strQuery += " JOIN [@Z_PAY_APPT1] T2 on T1.DocEntry=T2.DocEntry"
                    strQuery += " where T0.U_Z_DocType='R' AND T1.U_Z_AUser='" & oApplication.Company.UserName & "'"
                    otestRs.DoQuery(strQuery)
                    If otestRs.RecordCount > 0 Then
                        HeadDocEntry = otestRs.Fields.Item(0).Value
                        UserLineId = otestRs.Fields.Item(1).Value
                    End If

                   

                    strQuery = "Select * from [@Z_PAY_APHIS] where U_Z_DocEntry='" & strDocEntry & "' and U_Z_DocType='R' and U_Z_ApproveBy='" & oApplication.Company.UserName & "'"
                    oRecordSet.DoQuery(strQuery)
                    Dim oTemp As SAPbobsCOM.Recordset
                    oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    If oRecordSet.RecordCount > 0 Then
                        oGeneralParams.SetProperty("DocEntry", oRecordSet.Fields.Item("DocEntry").Value)
                        oGeneralData = oGeneralService.GetByParams(oGeneralParams)
                        oGeneralData.SetProperty("U_Z_AppStatus", oGrid.DataTable.GetValue("U_Z_AppStatus", index))
                        oGeneralData.SetProperty("U_Z_Remarks", oGrid.DataTable.GetValue("U_Z_Remarks", index))
                        oGeneralData.SetProperty("U_Z_ADocEntry", HeadDocEntry)
                        oGeneralData.SetProperty("U_Z_ALineId", UserLineId)
                        '  Dim oTemp As SAPbobsCOM.Recordset
                        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oTemp.DoQuery("Select * ,isnull(""firstName"",'') +  ' ' + isnull(""middleName"",'') +  ' ' + isnull(""lastName"",'') 'EmpName' from OHEM where ""userid""=" & oApplication.Company.UserSignature)
                        If oTemp.RecordCount > 0 Then
                            oGeneralData.SetProperty("U_Z_EmpId", oTemp.Fields.Item("empID").Value.ToString())
                            oGeneralData.SetProperty("U_Z_EmpName", oTemp.Fields.Item("EmpName").Value)
                            strEmpName = oTemp.Fields.Item("EmpName").Value
                        Else
                            oGeneralData.SetProperty("U_Z_EmpId", "")
                            oGeneralData.SetProperty("U_Z_EmpName", "")
                        End If
                        oGeneralData.SetProperty("U_Z_YEAR", strYear)
                        oGeneralData.SetProperty("U_Z_MONTH", strMonth)
                        oGeneralData.SetProperty("U_Z_CompNo", strComp)
                        oGeneralData.SetProperty("U_Z_Attachment", strFile)

                        strQuery = "Select Top 1 U_Z_AUser From [@Z_PAY_APPT2] Where  DocEntry = '" & HeadDocEntry & "' And LineId > '" & UserLineId.ToString() & "' and isnull(U_Z_AMan,'')='Y'  Order By LineId Asc "
                        oTemp.DoQuery(strQuery)
                        If oTemp.RecordCount > 0 Then
                            oGeneralData.SetProperty("U_Z_NextApprover", oTemp.Fields.Item(0).Value.ToString)
                        Else
                            oGeneralData.SetProperty("U_Z_NextApprover", oApplication.Company.UserName)
                        End If

                        CopyAttachment(strFile)
                        oGeneralService.Update(oGeneralData)
                    ElseIf (strDocEntry <> "" And strDocEntry <> "0") Then
                        ' Dim oTemp As SAPbobsCOM.Recordset
                        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oTemp.DoQuery("Select * ,isnull(""firstName"",'') + ' ' + isnull(""middleName"",'') +  ' ' + isnull(""lastName"",'') 'EmpName' from OHEM where ""userid""=" & oApplication.Company.UserSignature)
                        If oTemp.RecordCount > 0 Then
                            oGeneralData.SetProperty("U_Z_EmpId", oTemp.Fields.Item("empID").Value.ToString())
                            oGeneralData.SetProperty("U_Z_EmpName", oTemp.Fields.Item("EmpName").Value)
                            strEmpName = oTemp.Fields.Item("EmpName").Value
                        Else
                            oGeneralData.SetProperty("U_Z_EmpId", "")
                            oGeneralData.SetProperty("U_Z_EmpName", "")
                        End If
                        oGeneralData.SetProperty("U_Z_DocEntry", strDocEntry.ToString())
                        oGeneralData.SetProperty("U_Z_DocType", "R")
                        oGeneralData.SetProperty("U_Z_AppStatus", oGrid.DataTable.GetValue("U_Z_AppStatus", index))
                        oGeneralData.SetProperty("U_Z_Remarks", oGrid.DataTable.GetValue("U_Z_Remarks", index))
                        oGeneralData.SetProperty("U_Z_ApproveBy", oApplication.Company.UserName)
                        oGeneralData.SetProperty("U_Z_Approvedt", System.DateTime.Now)
                        oGeneralData.SetProperty("U_Z_ADocEntry", HeadDocEntry)
                        oGeneralData.SetProperty("U_Z_ALineId", UserLineId)
                        oGeneralData.SetProperty("U_Z_YEAR", strYear)
                        oGeneralData.SetProperty("U_Z_MONTH", strMonth)
                        oGeneralData.SetProperty("U_Z_CompNo", strComp)
                        oGeneralData.SetProperty("U_Z_Attachment", strFile)
                        CopyAttachment(strFile)
                        strQuery = "Select Top 1 U_Z_AUser From [@Z_PAY_APPT2] Where  DocEntry = '" & HeadDocEntry & "' And LineId > '" & UserLineId.ToString() & "' and isnull(U_Z_AMan,'')='Y'  Order By LineId Asc "
                        oTemp.DoQuery(strQuery)
                        If oTemp.RecordCount > 0 Then
                            oGeneralData.SetProperty("U_Z_NextApprover", oTemp.Fields.Item(0).Value.ToString)
                        Else
                            oGeneralData.SetProperty("U_Z_NextApprover", oApplication.Company.UserName)
                        End If
                        oGeneralService.Add(oGeneralData)
                    End If
                    updateFinalStatus(aForm, HeadDocEntry, strDocEntry, strRefCode, oGrid.DataTable.GetValue("U_Z_AppStatus", index), oGrid.DataTable.GetValue("U_Z_Remarks", index), strEmpID)
                    If oGrid.DataTable.GetValue("U_Z_AppStatus", index) = "A" Then
                        SendMessage(strDocType1, strDocEntry, oGrid.DataTable.GetValue("U_Z_AppStatus", index), HeadDocEntry, strEmpName, oApplication.Company.UserName, "R")
                    End If

                Next
            End If
            HeaderGridBind(oForm, "R")
            HeaderSumGridBind(oForm, "R")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub updateFinalStatus(ByVal aForm As SAPbouiCOM.Form, ByVal strTemplateNo As String, ByVal strDocEntry As String, ByVal strRefCode As String, ByVal strStatus As String, ByVal Remarks As String, ByVal aEmpID As String)
        Try

            Dim StrMailMessage, sQuery As String
            Dim oRecordSet, oTemp As SAPbobsCOM.Recordset
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If strStatus = "A" Then
                sQuery = " Select T2.DocEntry "
                sQuery += " From [@Z_PAY_APPT2] T2 "
                sQuery += " JOIN [@Z_PAY_OAPPT] T3 ON T2.DocEntry = T3.DocEntry  "
                sQuery += " JOIN [@Z_PAY_APPT1] T4 ON T4.DocEntry = T3.DocEntry  "
                sQuery += " Where  U_Z_AFinal = 'Y'"
                sQuery += " And T2.U_Z_AUser = '" + oApplication.Company.UserName + "' And T3.U_Z_DocType = 'R'"
                oRecordSet.DoQuery(sQuery)
                If Not oRecordSet.EoF Then
                    strQuery = "Update [@Z_PAY_Approval] set U_Z_AppStatus='A' where Code='" & strDocEntry & "'"
                    oRecordSet.DoQuery(strQuery)
                    strQuery = "Update [@Z_PAYROLL] set U_Z_AppStatus='A' where Code='" & strRefCode & "'"
                    oRecordSet.DoQuery(strQuery)
                    strQuery = "Select * from  [@Z_PAYROLL] where Code='" & strRefCode & "'"
                    oTemp.DoQuery(strQuery)
                    strQuery = " Payroll Company : " & oTemp.Fields.Item("U_Z_CompNo").Value & "  month : " & MonthName(oTemp.Fields.Item("U_Z_MONTH").Value) & " and Year :" & oTemp.Fields.Item("U_Z_YEAR").Value & ""

                    StrMailMessage = "Payroll worksheet for " & strQuery & "  has been Approved "
                    UserMessage(StrMailMessage, strDocEntry, aEmpID)
                End If
            ElseIf strStatus = "R" Then
                sQuery = " Select T2.DocEntry "
                sQuery += " From [@Z_PAY_APPT2] T2 "
                sQuery += " JOIN [@Z_PAY_OAPPT] T3 ON T2.DocEntry = T3.DocEntry  "
                sQuery += " JOIN [@Z_PAY_APPT1] T4 ON T4.DocEntry = T3.DocEntry  "
                sQuery += " Where T2.U_Z_AUser = '" + oApplication.Company.UserName + "' And T3.U_Z_DocType = 'R'"
                oRecordSet.DoQuery(sQuery)
                If Not oRecordSet.EoF Then
                    strQuery = "Update [@Z_PAY_Approval] set U_Z_AppStatus='R',U_Z_Remarks='" & Remarks & "' where Code='" & strDocEntry & "'"
                    oRecordSet.DoQuery(strQuery)
                    strQuery = "Update [@Z_PAYROLL] set U_Z_AppStatus='R' where Code='" & strRefCode & "'"
                    oRecordSet.DoQuery(strQuery)
                    ' StrMailMessage = "Payroll worksheet has been Rejected for the Document number :" & CInt(strRefCode)
                    strQuery = "Select * from  [@Z_PAYROLL] where Code='" & strRefCode & "'"
                    oTemp.DoQuery(strQuery)
                    strQuery = " Payroll Company : " & oTemp.Fields.Item("U_Z_CompNo").Value & "  month : " & MonthName(oTemp.Fields.Item("U_Z_MONTH").Value) & " and Year :" & oTemp.Fields.Item("U_Z_YEAR").Value & ""

                    StrMailMessage = "Payroll worksheet for " & strQuery & "  has been Rejected "
                    UserMessage(StrMailMessage, strDocEntry, aEmpID)
                End If
            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub UserMessage(ByVal strMessage As String, ByVal strDocEntry As String, ByVal SAPUser As String)
        Dim oCmpSrv As SAPbobsCOM.CompanyService
        Dim oMessageService As SAPbobsCOM.MessagesService
        Dim oMessage As SAPbobsCOM.Message
        Dim pMessageDataColumns As SAPbobsCOM.MessageDataColumns
        Dim pMessageDataColumn As SAPbobsCOM.MessageDataColumn
        Dim oLines As SAPbobsCOM.MessageDataLines
        Dim oLine As SAPbobsCOM.MessageDataLine
        Dim oRecipientCollection As SAPbobsCOM.RecipientCollection
        oCmpSrv = oApplication.Company.GetCompanyService()
        oMessageService = oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.MessagesService)
        oMessage = oMessageService.GetDataInterface(SAPbobsCOM.MessagesServiceDataInterfaces.msdiMessage)
        oMessage.Subject = "Payroll worksheet Approval Notification "
        oMessage.Text = strMessage
        oRecipientCollection = oMessage.RecipientCollection
        oRecipientCollection.Add()
        oRecipientCollection.Item(0).SendInternal = SAPbobsCOM.BoYesNoEnum.tYES
        oRecipientCollection.Item(0).UserCode = SAPUser
        pMessageDataColumns = oMessage.MessageDataColumns
        pMessageDataColumn = pMessageDataColumns.Add()
        pMessageDataColumn.ColumnName = "Document Number"
        oLines = pMessageDataColumn.MessageDataLines()
        oLine = oLines.Add()
        oLine.Value = strDocEntry
        oMessageService.SendMessage(oMessage)
        oApplication.Utilities.SendMail_Approval(strMessage, "mail", SAPUser)
    End Sub

    Public Sub SendMessage(ByVal strReqType As String, ByVal strReqNo As String, ByVal strAppStatus As String _
        , ByVal strTemplateNo As String, ByVal strOrginator As String, ByVal strAuthorizer As String, ByVal enDocType As String)
        Try
            Dim strQuery As String
            Dim strMessageUser As String
            Dim intLineID As Integer
            Dim oRecordSet, oTemp As SAPbobsCOM.Recordset
            Dim oCmpSrv As SAPbobsCOM.CompanyService
            Dim oMessageService As SAPbobsCOM.MessagesService
            Dim oMessage As SAPbobsCOM.Message
            Dim pMessageDataColumns As SAPbobsCOM.MessageDataColumns
            Dim pMessageDataColumn As SAPbobsCOM.MessageDataColumn
            Dim oLines As SAPbobsCOM.MessageDataLines
            Dim oLine As SAPbobsCOM.MessageDataLine
            Dim oRecipientCollection As SAPbobsCOM.RecipientCollection
            oCmpSrv = oApplication.Company.GetCompanyService()
            oMessageService = oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.MessagesService)
            oMessage = oMessageService.GetDataInterface(SAPbobsCOM.MessagesServiceDataInterfaces.msdiMessage)
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strQuery = "Select LineId From [@Z_PAY_APPT2] Where DocEntry = '" & strTemplateNo & "' And U_Z_AUser = '" & strAuthorizer & "'"
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                intLineID = CInt(oRecordSet.Fields.Item(0).Value)
                strQuery = "Select Top 1 U_Z_AUser From [@Z_PAY_APPT2] Where  DocEntry = '" & strTemplateNo & "' And LineId > '" & intLineID.ToString() & "' and isnull(U_Z_AMan,'')='Y'  Order By LineId Asc "
                oRecordSet.DoQuery(strQuery)

                If Not oRecordSet.EoF Then
                    strMessageUser = oRecordSet.Fields.Item(0).Value
                    oMessage.Subject = "Payroll worksheet Need Your Approval "
                    Dim strMessage As String = ""
                    strMessage = " Requested by  :" & oApplication.Company.UserName & ": Document Number : " & strReqNo

                    strQuery = "Select * from  [@Z_PAY_Approval]  where Code='" & strReqNo & "'"
                    oTemp.DoQuery(strQuery)

                    strMessage = "  Generated / Approved by  :" & oApplication.Company.UserName & ": For  Payroll Company : " & oTemp.Fields.Item("U_Z_CompNo").Value & "  month : " & MonthName(oTemp.Fields.Item("U_Z_MONTH").Value) & " and Year :" & oTemp.Fields.Item("U_Z_YEAR").Value & ""

                    strQuery = "Update [@Z_PAY_Approval] set U_Z_CurApprover='" & oApplication.Company.UserName & "',U_Z_NxtApprover='" & strMessageUser & "' where Code='" & strReqNo & "'"
                    oTemp.DoQuery(strQuery)

                    oMessage.Text = "Payroll worksheet " & " " & strMessage & " Needs Your Approval "
                    strMessage = "Payroll worksheet " & " " & strMessage & " Needs Your Approval "

                    oRecipientCollection = oMessage.RecipientCollection
                    oRecipientCollection.Add()
                    oRecipientCollection.Item(0).SendInternal = SAPbobsCOM.BoYesNoEnum.tYES
                    oRecipientCollection.Item(0).UserCode = strMessageUser
                    pMessageDataColumns = oMessage.MessageDataColumns
                    pMessageDataColumn = pMessageDataColumns.Add()
                    pMessageDataColumn.ColumnName = "Document Number"
                    oLines = pMessageDataColumn.MessageDataLines()
                    oLine = oLines.Add()
                    oLine.Value = strReqNo
                    oMessageService.SendMessage(oMessage)

                    oApplication.Utilities.SendMail_Approval(strMessage, "Mail", strMessageUser)
                End If
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region


    Private Sub reDrawForm(ByVal oForm As SAPbouiCOM.Form)
        Try
            oForm.Freeze(True)
            ' oForm.Items.Item("9").Width = oForm.Width - 25
            ' oForm.Items.Item("9").Height = oForm.Items.Item("10").Height + 10
            oForm.Items.Item("4").Width = oForm.Width - 150
            oForm.Items.Item("4").Height = 150
            oForm.Items.Item("5").Top = oForm.Items.Item("4").Top + oForm.Items.Item("4").Height + 5
            oForm.Items.Item("6").Top = oForm.Items.Item("5").Top + 15
            oForm.Items.Item("6").Width = oForm.Items.Item("4").Width
            oForm.Items.Item("6").Height = oForm.Height / 2 - 20

            oForm.Items.Item("9").Width = oForm.Items.Item("4").Width
            oForm.Items.Item("9").Height = 150
            oForm.Items.Item("10").Top = oForm.Items.Item("5").Top + 15
            oForm.Items.Item("10").Width = oForm.Items.Item("4").Width
            oForm.Items.Item("10").Height = oForm.Height / 2 - 20
            oForm.Items.Item("8").Width = oForm.Width - 100
            oForm.Items.Item("8").Height = oForm.Items.Item("10").Top + oForm.Items.Item("10").Height + 2


            'oForm.Items.Item("9").Top = intHeight
            'oForm.Items.Item("9").Left = 15
            'oForm.Items.Item("9").Height = oForm.Height / 2 - 20 ' - 300

            'oForm.Items.Item("5").Left = oForm.Items.Item("9").Left '+ oForm.Items.Item("27").Width
            'oForm.Items.Item("5").Top = oForm.Items.Item("9").Top + oForm.Items.Item("9").Width + 3

            'oForm.Items.Item("10").Left = oForm.Items.Item("5").Left
            'oForm.Items.Item("10").Top = oForm.Items.Item("5").Top + oForm.Items.Item("5").Height + 1
            '  oForm.Items.Item("32").Left = oForm.Items.Item("28").Left
            ' oForm.Items.Item("32").Top = oForm.Items.Item("31").Top
            oForm.Freeze(False)
        Catch ex As Exception
            oForm.Freeze(False)
        End Try
    End Sub
#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_Pay_Approval Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                            Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If (pVal.ItemUID = "6" Or pVal.ItemUID = "10") And pVal.ColUID = "Code" Then
                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                    Dim strCode As String
                                    oGrid = oForm.Items.Item(pVal.ItemUID).Specific
                                    strCode = oGrid.DataTable.GetValue("Code", pVal.Row)
                                    If strCode <> "" Then
                                        Dim oOBj As New clsPayrolLDetails
                                        frmSourceForm = oForm
                                        Dim oRec As SAPbobsCOM.Recordset
                                        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        oRec.DoQuery("Select * from [@Z_PAYROLL1] where Code='" & strCode & "'")
                                        oOBj.LoadForm(oRec.Fields.Item("U_Z_MONTH").Value, oRec.Fields.Item("U_Z_YEAR").Value, strCode, "WorkSheet")
                                        frmSourceForm = oForm
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If

                        End Select

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                reDrawForm(oForm)
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                            Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If (pVal.ItemUID = "4" Or pVal.ItemUID = "9") And (pVal.ColUID = "U_Z_Attachment" Or pVal.ColUID = "U_Z_FileName") Then
                                    oGrid = oForm.Items.Item(pVal.ItemUID).Specific
                                    ' oGrid.Columns.Item("RowsHeader").Click(pVal.Row)
                                    '  LoadFiles(oForm, pVal.ItemUID)
                                    LoadFiles(oGrid.DataTable.GetValue(pVal.ColUID, pVal.Row))
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                                If (pVal.ItemUID = "4" Or pVal.ItemUID = "9") And pVal.ColUID = "Code" Then
                                    oGrid = oForm.Items.Item(pVal.ItemUID).Specific
                                    Dim strDocEntry As String = oGrid.DataTable.GetValue("Code", pVal.Row)
                                    Dim oOBj As New clsAppHisDetails
                                    oOBj.LoadForm(oForm, strDocEntry)
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                                If pVal.ItemUID = "4" And pVal.ColUID = "U_Z_Attachment" Then
                                    oGrid = oForm.Items.Item(pVal.ItemUID).Specific
                                    Dim strPath As String = oGrid.DataTable.Columns.Item("U_Z_Attachment").Cells.Item(pVal.Row).Value.ToString()
                                    FileOpen()
                                    If strFilepath = "" Then
                                        oApplication.Utilities.Message("Please Select a File", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        BubbleEvent = False
                                    Else
                                        oGrid.DataTable.Columns.Item("U_Z_Attachment").Cells.Item(pVal.Row).Value = strFilepath
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Select Case pVal.ItemUID
                                    Case "1000001"
                                        oForm.PaneLevel = 1
                                        oGrid = oForm.Items.Item("4").Specific
                                        oGrid.Columns.Item("RowsHeader").Click(0)
                                    Case "7"
                                        oForm.PaneLevel = 2
                                        oGrid = oForm.Items.Item("9").Specific
                                        oGrid.Columns.Item("RowsHeader").Click(0)
                                    Case "3"
                                        Dim intRet As Integer = oApplication.SBO_Application.MessageBox("Are you sure want to submit the document?", 2, "Yes", "No", "")
                                        If intRet = 1 Then
                                            addUpdateDocument(oForm)
                                        End If
                                End Select
                                If (pVal.ItemUID = "9" Or pVal.ItemUID = "4") And pVal.ColUID = "RowsHeader" And pVal.Row > -1 Then
                                    oGrid = oForm.Items.Item(pVal.ItemUID).Specific
                                    Dim strDocEntry As String = oGrid.DataTable.GetValue("U_Z_RefCode", pVal.Row)
                                    oForm.Freeze(True)
                                    If pVal.ItemUID = "4" Then
                                        ViewWorkSheet(oForm, strDocEntry, "6")
                                    Else
                                        ViewWorkSheet(oForm, strDocEntry, "10")
                                    End If
                                    oForm.Freeze(False)
                                End If
                        End Select
                End Select
            End If
        Catch ex As Exception
            oForm.Freeze(False)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case mnu_Pay_PayRA
                    LoadForm("R")
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
#Region "FileOpen"
    Private Sub FileOpen()
        Dim mythr As New System.Threading.Thread(AddressOf ShowFileDialog)
        mythr.SetApartmentState(Threading.ApartmentState.STA)
        mythr.Start()
        mythr.Join()
    End Sub

    Private Sub ShowFileDialog()
        Dim oDialogBox As New OpenFileDialog
        Dim strMdbFilePath As String
        Dim oProcesses() As Process
        Try
            Dim aform As New System.Windows.Forms.Form
            aform.TopMost = True
            oProcesses = Process.GetProcessesByName("SAP Business One")
            If oProcesses.Length <> 0 Then
                For i As Integer = 0 To oProcesses.Length - 1
                    Dim MyWindow As New clsListener.WindowWrapper(oProcesses(i).MainWindowHandle)
                    If oDialogBox.ShowDialog(MyWindow) = DialogResult.OK Then
                        strMdbFilePath = oDialogBox.FileName
                        strFilepath = oDialogBox.FileName
                        Exit For
                    Else
                        strFilepath = ""
                        Exit For
                    End If
                Next
            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
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


    Private Sub LoadFiles(aFileName As String)


        Dim strFilename, strFilePath As String
        strFilename = aFileName
        Dim Filename As String = Path.GetFileName(strFilename)
        strFilePath = aFileName

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
#End Region
End Class
