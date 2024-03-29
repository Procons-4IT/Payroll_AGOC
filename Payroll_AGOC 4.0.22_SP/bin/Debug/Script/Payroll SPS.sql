/****** Object:  StoredProcedure [dbo].[INSERTPAYROLL_LEAVEDETAILS_FOR_MONTH]    Script Date: 12/02/2014 09:36:40 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[INSERTPAYROLL_LEAVEDETAILS_FOR_MONTH]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[INSERTPAYROLL_LEAVEDETAILS_FOR_MONTH]
 
/****** Object:  StoredProcedure [dbo].[RESETPARYROLLWORKSHEET_REGULAR]    Script Date: 12/02/2014 09:36:40 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[RESETPARYROLLWORKSHEET_REGULAR]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[RESETPARYROLLWORKSHEET_REGULAR]
 
/****** Object:  StoredProcedure [dbo].[UPDATE_EMPLOYEE_LEAVEDETAILS]    Script Date: 12/02/2014 09:36:40 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[UPDATE_EMPLOYEE_LEAVEDETAILS]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[UPDATE_EMPLOYEE_LEAVEDETAILS]
 
/****** Object:  StoredProcedure [dbo].[UpdateEmployeeLeavedetails_Employee]    Script Date: 12/02/2014 09:36:40 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[UpdateEmployeeLeavedetails_Employee]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[UpdateEmployeeLeavedetails_Employee]
 
/****** Object:  StoredProcedure [dbo].[UpdateEmployeeLeavedetails_EMployee_Month_Company]    Script Date: 12/02/2014 09:36:40 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[UpdateEmployeeLeavedetails_EMployee_Month_Company]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[UpdateEmployeeLeavedetails_EMployee_Month_Company]
 
/****** Object:  StoredProcedure [dbo].[UpdateEmployeeLeavedetails_Employee1]    Script Date: 12/02/2014 09:36:40 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[UpdateEmployeeLeavedetails_Employee1]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[UpdateEmployeeLeavedetails_Employee1]
 
/****** Object:  StoredProcedure [dbo].[ADD_LEAVEDETAILS]    Script Date: 12/02/2014 09:36:40 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ADD_LEAVEDETAILS]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ADD_LEAVEDETAILS]
 
/****** Object:  StoredProcedure [dbo].[UPDATEPAYROLLTABLE]    Script Date: 12/02/2014 09:36:40 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[UPDATEPAYROLLTABLE]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[UPDATEPAYROLLTABLE]
 
/****** Object:  StoredProcedure [dbo].[UpdatePayrollTotal_Payroll_Employee]    Script Date: 12/02/2014 09:36:40 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[UpdatePayrollTotal_Payroll_Employee]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[UpdatePayrollTotal_Payroll_Employee]
 
/****** Object:  StoredProcedure [dbo].[UpdateSavingScheme]    Script Date: 12/02/2014 09:36:40 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[UpdateSavingScheme]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[UpdateSavingScheme]
 
/****** Object:  StoredProcedure [dbo].[ADD_WORKSHEET]    Script Date: 12/02/2014 09:36:40 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ADD_WORKSHEET]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ADD_WORKSHEET]
 
/****** Object:  StoredProcedure [dbo].[ADD_EMPLOYEE]    Script Date: 12/02/2014 09:36:40 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ADD_EMPLOYEE]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ADD_EMPLOYEE]
 
/****** Object:  StoredProcedure [dbo].[UPDATELEAVEBALANCE_TRANSACTION]    Script Date: 12/02/2014 09:36:40 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[UPDATELEAVEBALANCE_TRANSACTION]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[UPDATELEAVEBALANCE_TRANSACTION]
 
/****** Object:  StoredProcedure [dbo].[UPDATELEAVEBALANCE_TRANSACTION_EMPLOYEE]    Script Date: 12/02/2014 09:36:40 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[UPDATELEAVEBALANCE_TRANSACTION_EMPLOYEE]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[UPDATELEAVEBALANCE_TRANSACTION_EMPLOYEE]
 
/****** Object:  StoredProcedure [dbo].[UpdatePayroll]    Script Date: 12/02/2014 09:36:40 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[UpdatePayroll]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[UpdatePayroll]
 
/****** Object:  StoredProcedure [dbo].[RESET_AIRTICKET_LOAN]    Script Date: 12/02/2014 09:36:40 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[RESET_AIRTICKET_LOAN]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[RESET_AIRTICKET_LOAN]
 
/****** Object:  StoredProcedure [dbo].[RESET_AIRTICKET_LOAN1]    Script Date: 12/02/2014 09:36:40 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[RESET_AIRTICKET_LOAN1]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[RESET_AIRTICKET_LOAN1]
 
/****** Object:  StoredProcedure [dbo].[SP_PAY_EXECUTEQUERY]    Script Date: 12/02/2014 09:36:40 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[SP_PAY_EXECUTEQUERY]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[SP_PAY_EXECUTEQUERY]
 
/****** Object:  StoredProcedure [dbo].[SP_PAY_EXECUTEQUERY]    Script Date: 12/02/2014 09:36:40 ******/
SET ANSI_NULLS ON
 
SET QUOTED_IDENTIFIER ON
 
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[SP_PAY_EXECUTEQUERY]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE procedure [dbo].[SP_PAY_EXECUTEQUERY] 

@STRQUERY VARCHAR(MAX)
AS 
BEGIN
 exec (@strquery)


END
' 
END
 
/****** Object:  StoredProcedure [dbo].[RESET_AIRTICKET_LOAN1]    Script Date: 12/02/2014 09:36:40 ******/
SET ANSI_NULLS ON
 
SET QUOTED_IDENTIFIER ON
 
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[RESET_AIRTICKET_LOAN1]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'  
CREATE PROCEDURE [dbo].[RESET_AIRTICKET_LOAN1]  
@strRefCode Varchar(200),  
@Year Int,  
@Month Int  
As  
  
BEGIN  
Declare @strEmpRefCode Varchar(30)  
Declare @dblAccuralAmount Decimal(19,3)  
Declare @DblRedimAmount Decimal(19,3)  
Declare @dblCM Decimal(19,3)  
Declare @dblRem Decimal(19,3)  
Declare @dblCurAmt Decimal(19,3)  
Declare @dblredim Decimal(19,3)  
Declare @dblClosingBalance Decimal(19,3)  
Declare Z_PAY10 cursor For Select Code from [@Z_PAY10] where U_Z_EMPID=@strRefCode  
Open Z_PAY10  
Fetch next from Z_PAY10 into @strEmpRefCode  
while @@FETCH_STATUS=0  
BEGIN  
 
   set @dblCM=( Select isnull(Sum(U_Z_NoofDays),0)  from [@Z_Payroll6] where U_Z_EmpID=@strRefCode  and U_Z_TktCode =@strEmpRefcode )
   set @dblRem=( Select isnull(Sum(U_Z_Redim),0)  from [@Z_Payroll6] where U_Z_EmpID=@strRefCode  and U_Z_TktCode =@strEmpRefcode )
   set @dblCurAmt=( Select isnull(Sum(U_Z_CurAMount),0)  from [@Z_Payroll6] where U_Z_EmpID=@strRefCode  and U_Z_TktCode =@strEmpRefcode )
   set @dblredim=( Select isnull(Sum(U_Z_Amount),0)  from [@Z_Payroll6] where U_Z_EmpID=@strRefCode  and U_Z_TktCode =@strEmpRefcode )
   set @dblAccuralAmount=( Select isnull(Sum(U_Z_CurAmount),0)  from [@Z_Payroll6] where U_Z_EmpID=@strRefCode  and U_Z_TktCode =@strEmpRefcode )
  
  set  @dblClosingBalance = @dblCurAmt - @dblredim  
  set  @dblClosingBalance = @dblAccuralAmount - @DblRedimAmount  
  Update [@Z_PAY10] set U_Z_BalAmount=@dblClosingBalance ,U_Z_CM=@dblCM,U_Z_Redim=@dblRem  where Code=@strEmpRefcode  and  U_Z_empID=@strRefCode  
  Update [@Z_PAY10] set U_Z_BalAmount=U_Z_OBAmt + U_Z_BalAmount -U_Z_Redim,U_Z_Balance=isnull(U_Z_OB,0)+isnull(U_Z_CM,0)-isnull(U_Z_Redim,0)   
  where Code=@strEmpRefcode  and  U_Z_empID=@strRefCode   
  
Fetch Next From Z_PAY10 into @strEmpRefCode  
end  
Close Z_PAY10  
DeAllocate Z_PAY10  
  
END  ' 
END
 
/****** Object:  StoredProcedure [dbo].[RESET_AIRTICKET_LOAN]    Script Date: 12/02/2014 09:36:40 ******/
SET ANSI_NULLS ON
 
SET QUOTED_IDENTIFIER ON
 
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[RESET_AIRTICKET_LOAN]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[RESET_AIRTICKET_LOAN]
@strRefCode Varchar(200),
@Year Int,
@Month Int
As

BEGIN

Declare @strEmpRefCode Varchar(30)
Declare @dblLoanAmount Decimal(19,3)
Declare @dblCM Decimal(19,3)
Declare @dblPaidAmount Decimal(19,3)

Declare Cursor1 cursor For Select Code,U_Z_LoanAmount from [@Z_PAY5] where U_Z_EMPID=@strRefCode

Open Cursor1
Fetch next from Cursor1 into @strEmpRefCode,@dblLoanAmount
while @@FETCH_STATUS=0
BEGIN
Select @dblCM=COUNT(*),@dblPaidAmount= Sum(U_Z_Amount) from [@Z_Payroll3] where U_Z_Amount>0 and   U_Z_Type =''L'' and U_Z_Field=@strEmpRefcode 
If @dblCM > 0 
Update [@Z_PAY5] set  U_Z_Status=''Process'', U_Z_PaidEMI=@dblCM  where Code=@strEmpRefcode 
Else
  Update [@Z_PAY5] set  U_Z_Status=''Open'', U_Z_PaidEMI=''0''  where Code=@strEmpRefcode

Update "@Z_PAY15" set "U_Z_Status"=''O'' where "U_Z_Month"=@Month and "U_Z_Year"=@Year  and "U_Z_TrnsRefCode"=@strEmpRefCode 
Update [@Z_PAY5] set U_Z_Balance = U_Z_NoEMI - U_Z_PaidEMI  where Code=@strEmpRefcode 
If @dblLoanAmount <= @dblPaidAmount 
 Update [@Z_PAY5] set U_Z_Status=''Close'' where U_Z_NoEMI = U_Z_PaidEMI 
 Fetch Next From Cursor1 into @strEmpRefCode,@dblLoanAmount
end

Close Cursor1
DeAllocate Cursor1

END' 
END
 
/****** Object:  StoredProcedure [dbo].[UpdatePayroll]    Script Date: 12/02/2014 09:36:40 ******/
SET ANSI_NULLS ON
 
SET QUOTED_IDENTIFIER ON
 
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[UpdatePayroll]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE Procedure [dbo].[UpdatePayroll] @Code nvarchar(20) ,@EmpID nvarchar(200)

as

BEGIN

Declare @IntYear as numeric,@IntMonth as numeric
Declare @EmpID1 as numeric 
declare @CompNo as nvarchar(100)
 Update [@Z_PAYROLL2] set U_Z_Amount=U_Z_Rate*U_Z_Value
 Update [@Z_PAYROLL3] set U_Z_Amount=U_Z_Rate*U_Z_Value
 Update [@Z_PAYROLL4] set U_Z_Amount=U_Z_Rate*U_Z_Value
 Update [@Z_PAYROLL5] set U_Z_Amount=U_Z_DailyRate * U_Z_Redim 
 Update [@Z_PAYROLL5] set   U_Z_Amount=((U_Z_DedRate *U_Z_DailyRate)/100 * U_Z_Redim)  where U_Z_PaidLeave<>''H'' 
 Update [@Z_PAYROLL5] set  U_Z_Amount=(((U_Z_DedRate * U_Z_DailyRate)/100)/2) * U_Z_Redim where  U_Z_PaidLeave=''H'' 
 Update [@Z_PAYROLL5] set  U_Z_Amount=0 where isnull(U_Z_Basic,''N'')=''Y'' and U_Z_PaidLeave=''A'' 
 Update [@Z_PAYROLL2] set U_Z_Amount=Round(U_Z_Amount,3) 
 Update [@Z_PAYROLL3] set U_Z_Amount=Round(U_Z_Amount,3) 
 Update [@Z_PAYROLL4] set U_Z_Amount=Round(U_Z_Amount,3) 
 Update [@Z_PAYROLL5] set U_Z_Amount=Round(U_Z_Amount,3) 
 Update [@Z_PAYROLL5] set U_Z_AcrAmount=Round(U_Z_AcrAmount,3) 
 Update [@Z_PAYROLL6] set U_Z_Amount=Round(U_Z_Amount,3) ,U_Z_CurAMount=Round(U_Z_CurAMount,3) 
 
 Declare @strCode as nvarchar(30)
 
 
 
  if exists (Select * from [@Z_PAYROLL1] where Code=@Code and U_Z_EmpID=@EmpID )
  BEGIN
    Declare @DailyRate as numeric(18,4)
    Declare @Earning as numeric(18,4),@Deduction as Numeric(18,4),@Contribution as numeric(18,4)
    Declare @UnPaidLeave as Numeric(18,4),@PaidLeave as numeric(18,4),@AcrAmount as numeric(18,4),@AnnualLeave as numeric(18,4)
    
    Select @DailyRate = Sum(U_Z_Balance * U_Z_DailyRate) from [@Z_Payroll5] where U_Z_RefCode=@Code and  U_Z_PaidLeave=''A''
    Update [@Z_PAYROLL2] set  U_Z_Value=@DailyRate  where U_Z_Type=''L'' and  U_Z_RefCode=@Code
    Update [@Z_PAYROLL2] set U_Z_Amount=U_Z_Rate*U_Z_Value
    Update [@Z_PAYROLL2] set U_Z_Amount=Round(U_Z_Amount,3)
    Select  @Earning = isnull(Sum(U_Z_Amount),0) from [@Z_Payroll2] where U_Z_RefCode=@Code 
    Select @Deduction = isnull(Sum(U_Z_Amount),0) from [@Z_Payroll3] where U_Z_RefCode=@Code 
    Select @Contribution = isnull(Sum(U_Z_Amount),0) from [@Z_Payroll4] where U_Z_RefCode=@Code 
    Update [@Z_PAYROLL1] set U_Z_Earning =@Earning ,U_Z_Deduction =@Deduction ,U_Z_Contri=@Contribution where Code=@Code 
    update [@Z_PAYROLL1] set  U_Z_Cost=isnull(U_Z_BasicSalary,0)+isnull(U_Z_Earning,0) +isnull(U_Z_Contri,0) - isnull(U_Z_Deduction,0) where Code=@Code
    Update [@Z_PAYROLL1] set U_Z_Earning=Round(U_Z_Earning,3),U_Z_Deduction=Round(U_Z_Deduction,3),U_Z_Contri=Round(U_Z_Contri,3) where Code=@Code

   
   Select @UnPaidLeave = isnull(Sum(U_Z_Amount),0) from [@Z_Payroll5] where U_Z_RefCode=@Code  and  U_Z_PaidLeave<>''P'' and U_Z_PaidLeave<>''A''
   Update [@Z_PAYROLL1] set  U_Z_UnPaidLeave=@UnPaidLeave  where Code=@Code
   Select @PaidLeave = isnull(Sum(U_Z_Amount),0) from [@Z_Payroll5] where U_Z_RefCode=@Code and  U_Z_PaidLeave=''P''
   Update [@Z_PAYROLL1] set  U_Z_PaidLeave=@PaidLeave  where Code=@Code 

   Select  @AcrAmount =ISNULL(Sum(U_Z_CurAmount),0) from [@Z_Payroll5] where U_Z_RefCode=@Code  and  U_Z_PaidLeave=''A''
   Update [@Z_PAYROLL1] set  U_Z_AcrAmt=@AcrAmount  where Code=@Code 

   Select @AnnualLeave =ISNULL(Sum(U_Z_Amount),0) from [@Z_Payroll5] where  U_Z_RefCode=@Code  and  U_Z_PaidLeave=''A''
   Update [@Z_PAYROLL1] set  U_Z_AnuLeave=@AnnualLeave  where Code=@Code 
   
  Declare @NetPayAmt as numeric(18,4),@cmpPayAmt as numeric(18,4),@AirAmount as numeric(18,4)
  Select @AirAmount =ISNULL(Sum(U_Z_Amount),0),@NetPayAmt =ISNULL(sum(U_Z_NetPayAmt),0),@cmpPayAmt =ISNULL( sum(U_Z_CmpPayAmt),0) from [@Z_Payroll6] where U_Z_RefCode=@Code
        
       Update [@Z_PAYROLL1] set  U_Z_NetPayAmt=@NetPayAmt ,U_Z_CmpPayAmt=@cmpPayAmt , U_Z_AirAmt=@AirAmount  where Code=@Code
            
  END
 
END' 
END
 
/****** Object:  StoredProcedure [dbo].[UPDATELEAVEBALANCE_TRANSACTION_EMPLOYEE]    Script Date: 12/02/2014 09:36:40 ******/
SET ANSI_NULLS ON
 
SET QUOTED_IDENTIFIER ON
 
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[UPDATELEAVEBALANCE_TRANSACTION_EMPLOYEE]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'  
--Exec UPDATELEAVEBALANCE_TRANSACTION_EMPLOYEE ''1'',2016,1  
  CREATE PROCEDURE [UPDATELEAVEBALANCE_TRANSACTION_EMPLOYEE]  
@aEmpID varchar(30),@aYear int,@aMonth int  
AS  
  
BEGIN  
Declare @Code Varchar(30)  
Declare @dblCM decimal(19,3)  
Declare @dblRem decimal(19,3)  
Declare @dblBalance decimal(19,3)  
Declare @dblCurAmt decimal(19,3)  
Declare @dblIncrement decimal(19,3)  
Declare @dblredim decimal(19,3)  
Declare @dblClosingBalance decimal(19,3)  
Declare @dblyearofExperience decimal(19,3)  
Declare @dblNoofDays1 decimal(19,3)  
  
  
Declare @dblCarriedForward decimal(19,3)  
Declare @dblYearly decimal(19,3)  
Declare @dblOpeningBalance decimal(19,3)  
Declare @dblTransaction decimal(19,3)  
Declare @dblAdjustment decimal(19,3)  
Declare @dblAccurred decimal(19,3)  
  
Declare @dblUnPostedTrns1 decimal(19,3)  
Declare @dblnoofEncashment decimal(19,3)  
Declare @dblFinalBalance Decimal(19,3)  
  
Declare @strRefCode varchar(max)  
Declare @strEmpRefcode varchar(max)  
Declare @strsql varchar(max)  
Declare @strLeaveName varchar(max)  
Declare @strCompany varchar(max)  
Declare @strQuery varchar(max)  
Declare @stOvType varchar(max)  
Declare @stString varchar(max)  
Declare @strCAFW varchar(1)  
Declare @JoiningDate DateTime
Declare @intRemainingMonths int=1
Declare @blnProrated Varchar(1)
Declare @blnAccurred VarChar(1)
Declare @blnAccured Varchar(1)
Declare @blnHREnt Varchar(1)
Declare @dblTotalyearlyEnttile  Decimal(19,2)
Declare @dblyearlyEntitled Decimal(19,2)
Declare @JoiningDay int
 Declare @NDays int                            
 Declare @MonthlyEntitlement Decimal(19,3)
 Declare @dblCalenderDays Decimal(19,2)
  
Declare LeaveUpdate Cursor For Select isnull(U_Z_LeaveCode,'''') from [@Z_EMP_LEAVE] where U_Z_EmpID=@aEmpID 
   Open LeaveUpdate  
     fetch Next from LeaveUpdate   Into @Code  
        While @@Fetch_Status=0  
        Begin  
			set @strLeaveName=(Select   U_Z_Leavename from [@Z_EMP_LEAVE] where U_Z_EmpID=@aEmpID and  U_Z_LeaveCode=@Code)  
			set @JoiningDate =(Select startDate from OHEM where empid=@aEmpID)	
			set @dblCalenderDays =(SELECT max(U_Z_CalenderDays) FROM [@Z_PAYROLL1]   WHERE U_Z_MONTH=@aMonth and U_Z_YEAR=@aYear )
			if Year(@JoiningDate)=@aYear and Month(@JoiningDate)=@aMonth 
				  set @dblYearly=(Select U_Z_DaysYear from [@Z_PAY_LEAVE] where Code=@Code) 
			else
			   set @dblYearly=(Select isnull(U_Z_Entile,0) from [@Z_EMP_LEAVE_BALANCE] where U_Z_EmpID=@aEmpID  and U_Z_LeaveCode=@Code  and U_Z_Year=@ayear) 

			
			set @dblTransaction =(select  isnull(sum(U_Z_NoofDays),0)  from [@Z_PAY_OLETRANS] where U_Z_Trnscode=@code and U_Z_Year= @ayear and U_Z_EmpID=@aEmpID ) 
			set @dblAccurred =(select  isnull(sum(U_Z_NoofDays),0) from [@Z_PAYROLL5] where U_Z_Leavecode=@code and U_Z_Year=@ayear and U_Z_EmpID=@aEmpID ) 

			set @dblAdjustment=(select sum(U_Z_Adjustment) from [@Z_PAYROLL5] where U_Z_Leavecode=@code and U_Z_Year=@ayear and U_Z_EmpID=@aEmpID  )
			 
			set @dblnoofEncashment= (select  SUM(isnull(U_Z_NoofDays,0)) from [@Z_PAY_OLETRANS_OFF] where U_Z_Posted=''Y'' and  U_Z_EMPID=@aEmpID  
			and U_Z_TrnsCode=@code	and U_Z_Year=@aYear  )

			Declare @blnCAFW int=0  
			set @strCAFW= (Select  isnull(U_Z_Accured,''N'') from [@Z_PAY_LEAVE] where Code=@Code)
			If @strCAFW  =''Y''  
			set @blnCAFW = 1  
						
			set @JoiningDate =(Select startDate from OHEM where empid=@aEmpID)
			set @blnProrated=(Select   isnull(U_Z_Prorate,''N'') from  [@Z_PAY_LEAVE] where code=@Code)  
			set @blnAccured=(Select   isnull(U_Z_Accured,''N'') from  [@Z_PAY_LEAVE] where Code =@Code)  
			set @blnHREnt=(Select   isnull(U_Z_EntHR,''N'') from  [@Z_PAY_LEAVE] where code= @Code)  
			set @dblyearlyEntitled=(Select   U_Z_DaysYear from [@Z_PAY_LEAVE] where code = @Code)  
			set @dblyearlyEntitled=@dblyearlyEntitled
			set @dblTotalyearlyEnttile=@dblyearlyEntitled


					Declare @dblYe As Decimal(19,2)
					Declare @SalaryCode Varchar(Max)
									
					If Year(@JoiningDate) = @ayear And Month(@JoiningDate) = @aMonth 
				BEGIN
                    If @blnProrated = ''Y'' 
					BEGIN
                         If @blnHREnt=''Y''
						 BEGIN
								  
							set @SalaryCode = (Select isnull(U_Z_HR_SalaryCode,'''') from OHEM where empID=@aEmpID)
							If @SalaryCode<>'''' 
						       BEGIN
								 set @dblYe= (Select U_Z_Entitle from [@Z_HR_OSALST] where U_Z_SalCode=@SalaryCode and U_Z_LeaveCode=@Code)
								END
							else
								Begin
									set @dblYe=0
								 End 

                                    If @dblYe > 0 
                                      set   @dblTotalyearlyEnttile = @dblYe
                                    Else
                                        set @dblYe = @dblTotalyearlyEnttile
                           END 
					     set @intRemainingMonths = 12 - @aMonth
                        set @dblTotalyearlyEnttile = @dblTotalyearlyEnttile / 12

						
						set @JoiningDay=Day(@JoiningDate)
						set @NDays=@dblCalenderDays-@JoiningDay+1   
						if @NDays>@dblCalenderDays
							set @NDays=@dblCalenderDays

						set @MonthlyEntitlement=@dblTotalyearlyEnttile
						set @MonthlyEntitlement=(@MonthlyEntitlement/@dblCalenderDays)*@NDays
						set @dblTotalyearlyEnttile = (@dblTotalyearlyEnttile * @intRemainingMonths) + @MonthlyEntitlement
						                      
                        Update [@Z_EMP_LEAVE_BALANCE] set U_Z_Entile=@dblTotalyearlyEnttile where U_Z_LeaveCode=@Code  and U_Z_EmpID=@aEmpID  and U_Z_Year=@aYear
					 END

					 else
					 BEGIN
					    If @blnHREnt=''Y''
						 BEGIN
								  
							set @SalaryCode = (Select isnull(U_Z_HR_SalaryCode,'''') from OHEM where empID=@aEmpID)
							If @SalaryCode<>'''' 
						       BEGIN
								 set @dblYe= (Select U_Z_Entitle from [@Z_HR_OSALST] where U_Z_SalCode=@SalaryCode and U_Z_LeaveCode=@Code)
								END
							else
								Begin
									set @dblYe=0
								 End 

                                    If @dblYe > 0 
                                      set   @dblTotalyearlyEnttile = @dblYe
                                    Else
                                        set @dblYe = @dblTotalyearlyEnttile
                           END 
					    set @dblTotalyearlyEnttile = @dblTotalyearlyEnttile 
                        set @dblTotalyearlyEnttile = @dblTotalyearlyEnttile 
                        Update [@Z_EMP_LEAVE_BALANCE] set U_Z_Entile=@dblTotalyearlyEnttile where U_Z_LeaveCode=@Code  and U_Z_EmpID=@aEmpID  and U_Z_Year=@aYear
					 END
                  END
        	

 else
				  BEGIN

BEGIN
                         If @blnHREnt=''Y''
						   BEGIN
							set @SalaryCode = (Select isnull(U_Z_HR_SalaryCode,'''') from OHEM where empID=@aEmpID )
							
							If @SalaryCode<>'''' 
						     set @dblYe= (Select U_Z_Entitle from [@Z_HR_OSALST] where U_Z_SalCode=@SalaryCode and U_Z_LeaveCode=@Code )
							else
							   set @dblYe=0
							
							If @dblYe > 0 
                                set   @dblTotalyearlyEnttile = @dblYe
                            Else
                               set  @dblTotalyearlyEnttile=@dblYearlyEntitled
                           END
					
                    
                        set @dblTotalyearlyEnttile = @dblTotalyearlyEnttile 
						  if @dblYearly>0    
								set @dblTotalyearlyEnttile=@dblYearly     
                        set @dblNoofDays1=ROUND(@dblTotalyearlyEnttile,2)    

						Update [@Z_EMP_LEAVE_BALANCE] set U_Z_Entile=@dblTotalyearlyEnttile where U_Z_LeaveCode=@Code  and U_Z_EmpID=@aEmpID  and U_Z_Year=@aYear
					 end
END


	 if exists (Select Code from [@Z_EMP_LEAVE_BALANCE] where U_Z_LeaveCode=@Code  and U_Z_EmpID=@aEmpID  and U_Z_Year=@aYear )  
   
		BEGIN  
         	Declare @strCode1 Varchar(30)  
			Declare @dblOB decimal(19,3) 

			Select @dblCarriedForward=isnull("U_Z_CAFWD",0) ,@dblYearly=isnull("U_Z_Entile",0) ,@strcode1="Code",@dblOB=isnull("U_Z_OB",0)  from   
			"@Z_EMP_LEAVE_BALANCE"   where "U_Z_LeaveCode"=@Code  and "U_Z_EmpID"=@aEmpID and "U_Z_Year"=@ayear  
			Declare @dblClosing Decimal(19,3)  
			set @dblClosing=0  
			If @blnCAFW = 0   
			set @dblClosing = @dblYearly  
			set @dblFinalBalance = isnull(@dblClosing,0) + isnull(@dblOB,0) + isnull(@dblCarriedForward,0) + isnull(@dblAccurred,0) - isnull(@dblTransaction,0) +isnull(@dblAdjustment,0) - isnull(@dblnoofEncashment,0)  
			Update [@Z_EMP_LEAVE_BALANCE] set  U_Z_OB=@dblOB , U_Z_LeaveName=@strLeaveName, U_Z_CAFWD=@dblCarriedForward ,  
			U_Z_ACCR=@dblAccurred,U_Z_Adjustment=@dblAdjustment ,U_Z_Trans=@dblTransaction ,U_Z_Balance=@dblFinalBalance   
			where "Code"=@strCode1 and U_Z_LeaveCode=@code and U_Z_Year=@ayear  
                                     
		 END   
    
    else
    
		 BEGIN
    
			set @dblOB= (Select isnull(U_Z_OB,0) OB from [@Z_EMP_LEAVE_BALANCE] where U_Z_LeaveCode=@Code and U_Z_EmpID=@aEmpID   and U_Z_Year= @aYear- 1)
			set @dblCarriedForward= (Select  isnull(U_Z_Balance,0)  from [@Z_EMP_LEAVE_BALANCE] where U_Z_LeaveCode=@Code and U_Z_EmpID=@aEmpID and
			U_Z_Year= @aYear- 1)
			set @dblYearly= (Select isnull(U_Z_Entile,0)  from [@Z_EMP_LEAVE_BALANCE] where U_Z_LeaveCode=@Code and U_Z_EmpID=@aEmpID   and U_Z_Year= @aYear- 1)
			set @dblYearly=(Select U_Z_DaysYear from [@Z_PAY_LEAVE] where Code=@Code) 
			set @dblClosing=0  
			If @blnCAFW = 0   
			set @dblClosing = @dblYearly  
			set @dblOB=0
			set @dblFinalBalance = isnull(@dblClosing,0) + isnull(@dblOB,0) + isnull(@dblCarriedForward,0) + isnull(@dblAccurred,0) - isnull(@dblTransaction,0) +			isnull(@dblAdjustment,0) - isnull(@dblnoofEncashment,0)  
			DECLARE @ExistingRowCount INT = 0  
			set @ExistingRowCount=(SELECT  isnull(Max(CONVERT(numeric,Code)),0)  FROM [@Z_EMP_LEAVE_BALANCE])

			INSERT INTO [@Z_EMP_LEAVE_BALANCE] (Code, Name, U_Z_EmpID,U_Z_Year,U_Z_CAFWD,U_Z_LeaveCode,U_Z_LeaveName,U_Z_Entile)  
			SELECT  
			Code = RIGHT(''00000000'' +CAST(ROW_NUMBER() OVER(ORDER BY H.Code ) +@ExistingRowCount AS NVARCHAR),8),  
			Name = RIGHT(''00000000'' +CAST(ROW_NUMBER() OVER(ORDER BY H.Code) +@ExistingRowCount AS NVARCHAR),8),  
			@aEmpID ,@aYear ,@dblCarriedForward ,@Code ,@strLeaveName  ,@dblYearly
			FROM  
			[@Z_PAY_LEAVE] H  where H.Code=@Code 
			Update [@Z_EMP_LEAVE_BALANCE] set  U_Z_OB=isnull(@dblOB,0) , U_Z_LeaveName=@strLeaveName, U_Z_CAFWD=isnull(@dblCarriedForward,0) , 
			U_Z_Entile=@dblYearly, 
			U_Z_ACCR=isnull(@dblAccurred,0),U_Z_Adjustment=isnull(@dblAdjustment,0) ,U_Z_Trans=isnull(@dblTransaction,0) ,U_Z_Balance=isnull(@dblFinalBalance,0)
			where Convert(Numeric,Code)=@ExistingRowCount+1 and U_Z_LeaveCode=@code and U_Z_Year=@ayear  

		 END      
      
     fetch Next from LeaveUpdate   Into @Code  
     end  
     Close LeaveUpdate  
     Deallocate LeaveUpdate  
      
  
  
  
END



' 
END
 
/****** Object:  StoredProcedure [dbo].[UPDATELEAVEBALANCE_TRANSACTION]    Script Date: 12/02/2014 09:36:40 ******/
SET ANSI_NULLS ON
 
SET QUOTED_IDENTIFIER ON
 
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[UPDATELEAVEBALANCE_TRANSACTION]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'

CREATE PROCEDURE [dbo].[UPDATELEAVEBALANCE_TRANSACTION]
@aEmpID varchar(30),@Code Varchar(30),@aYear int,@aMonth int
AS

BEGIN
Declare @dblCM decimal(19,3)
Declare @dblRem decimal(19,3)
Declare @dblBalance decimal(19,3)
Declare @dblCurAmt decimal(19,3)
Declare @dblIncrement decimal(19,3)
Declare @dblredim decimal(19,3)
Declare @dblClosingBalance decimal(19,3)
Declare @dblyearofExperience decimal(19,3)
Declare @dblNoofDays1 decimal(19,3)


Declare @dblCarriedForward decimal(19,3)
Declare @dblYearly decimal(19,3)
Declare @dblOpeningBalance decimal(19,3)
Declare @dblTransaction decimal(19,3)
Declare @dblAdjustment decimal(19,3)
Declare @dblAccurred decimal(19,3)

Declare @dblUnPostedTrns1 decimal(19,3)
Declare @dblnoofEncashment decimal(19,3)

Declare @strRefCode varchar(max)
Declare @strEmpRefcode varchar(max)
Declare @strsql varchar(max)
Declare @strLeaveName varchar(max)
Declare @strCompany varchar(max)
Declare @strQuery varchar(max)
Declare @stOvType varchar(max)
Declare @stString varchar(max)
Declare @strCAFW varchar(1)

if Exists ( Select * from [@Z_EMP_LEAVE] where U_Z_EmpID=@aEmpID  and U_Z_LeaveCode=@Code )
BEGIN
Select 
@dblYearly=U_Z_DaysYear,
@strEmpRefcode=T1.U_Z_LeaveCode,
@strLeaveName=T1.U_Z_LeaveName 
from [@Z_PAY_LEAVE] T0 Inner Join [@Z_EMP_LEAVE] T1 on T1.U_Z_LeaveCode=T0.Code  where T0.Code=@Code  



set @dblUnPostedTrns1 = (select isnull(sum(U_Z_NoofDays),0)  from [@Z_PAY_OLETRANS] where Code=Name and U_Z_Trnscode=@Code   and  U_Z_Year= @aYear  and U_Z_EmpID=@strRefCode group by U_Z_EmpID)
select @dblAccurred= isnull(sum(U_Z_NoofDays),0),
@dblAdjustment=sum(U_Z_Adjustment)
from [@Z_PAYROLL5] where U_Z_Leavecode=@Code and U_Z_Year=@aYear  and U_Z_EmpID=@strRefCode  group by U_Z_EmpID


select @dblnoofEncashment=SUM(U_Z_NoofDays) from [@Z_PAY_OLETRANS_OFF] where U_Z_Posted=''Y'' and  U_Z_EMPID=@strRefCode  and U_Z_TrnsCode=@code and U_Z_Year=@aYear 
--Dim dblnoofEncashment As Double = oTst.Fields.Item(0).Value


Select @strCAFW = isnull(U_Z_Accured,''N'') from [@Z_PAY_LEAVE] where Code=@code
--Dim blnCAFW As Boolean = False
Declare @blnCAFW as int=0
If @strCAFW  = ''Y'' 
set  @blnCAFW = 1


if exists (Select * from [@Z_EMP_LEAVE_BALANCE] where U_Z_LeaveCode=@code and U_Z_EmpID=@strRefCode  and U_Z_Year=@aYear)
BEGIN

Declare @strCode1 varchar(30)
Declare @dblOB decimal(19,3)
Declare @dblClosing decimal(19,3)
Declare @dblFinalBalance Decimal(19,3)
Select @dblCarriedForward=isnull("U_Z_CAFWD",0) , @dblyearly=isnull("U_Z_Entile",0),@strcode1="Code",@dblOB= isnull("U_Z_OB",0) 
from "@Z_EMP_LEAVE_BALANCE" where "U_Z_LeaveCode"=@code and "U_Z_EmpID"=@strRefCode   and "U_Z_Year"=@aYear 
--oTst=oApplication.Utilities.ExecuteSP(strQuery)
--Dim strcode1 As String = oTst.Fields.Item("Code").Value
--dblYearly = oTst.Fields.Item("Yearly").Value
--dblOB = oTst.Fields.Item("OB").Value
--''new addition 2014-01-16
If @blnCAFW = 0 
set  @dblClosing = @dblYearly
Else
set @dblClosing = 0

--''end
--dblCarriedForward = oTst.Fields.Item("U_Z_CAFWD").Value
set @dblFinalBalance = @dblClosing + @dblOB + @dblCarriedForward + @dblAccurred - @dblTransaction + @dblAdjustment - @dblnoofEncashment --'' - dblUnPostedTrns1
Update "@Z_EMP_LEAVE_BALANCE" set  "U_Z_OB"=@dblOB, @strLeaveName="U_Z_LeaveName" ,@dblCarriedForward=isnull("U_Z_CAFWD",0), "U_Z_ACCR"=@dblAccurred ,"U_Z_Adjustment"=@dblAdjustment,"U_Z_Trans"=@dblTransaction ,"U_Z_Balance"=@dblFinalBalance  where "Code"=@strcode1 and  "U_Z_LeaveCode"=@Code and U_Z_Year=@ayear
--oTst=oApplication.Utilities.ExecuteSP(strQuery)

END

else

BEGIN
Select @dbloB=isnull("U_Z_OB",0) ,@dblCarriedForward= isnull("U_Z_Balance",0),@dblYearly= isnull("U_Z_Entile",0) from "@Z_EMP_LEAVE_BALANCE" where "U_Z_LeaveCode"=@Code  and "U_Z_EmpID"=@strRefCode  and "U_Z_Year"=@aYear-1
--oTst=oApplication.Utilities.ExecuteSP(strQuery)
--dblOB = oTst.Fields.Item("OB").Value
--dblCarriedForward = oTst.Fields.Item("U_Z_CAFWD").Value
--dblYearly = dblYearly

If @blnCAFW = 0 
set  @dblClosing = @dblYearly
Else
set @dblClosing = 0
 set @dblFinalBalance = @dblClosing + @dblOB + @dblCarriedForward + @dblAccurred - @dblTransaction + @dblAdjustment - @dblnoofEncashment --'' -   

DECLARE @ExistingRowCount INT = 0
SELECT @ExistingRowCount = isnull(Max(CONVERT(numeric,Code)),0) +1 FROM [@Z_EMP_LEAVE_BALANCE]
Declare @NewCode varchar(10)
set @NewCode=@ExistingRowCount
Declare @strCode11 Varchar-- = oApplication.Utilities.getMaxCode("@Z_EMP_LEAVE_BALANCE", "Code")
Insert into [@Z_EMP_LEAVE_BALANCE] (Code,Name,U_Z_EmpID,U_Z_Year,U_Z_CAFWD,U_Z_LeaveCode,U_Z_LeaveName) 
Values
( @NewCode,@NewCode ,
  @strRefCode,@ayear , @dblCarriedForward ,@Code,@strLeaveName )
Update [@Z_EMP_LEAVE_BALANCE] set U_Z_OB=@dblOB , U_Z_Entile= @dblYearly, U_Z_CAFWD= @dblCarriedForward ,  U_Z_ACCR=@dblAccurred,U_Z_Adjustment= @dblAdjustment,U_Z_Trans= @dblTransaction ,U_Z_Balance=@dblFinalBalance  where  Code=@code

                               
END


END



END
' 
END
 
/****** Object:  StoredProcedure [dbo].[ADD_EMPLOYEE]    Script Date: 12/02/2014 09:36:40 ******/
SET ANSI_NULLS ON
 
SET QUOTED_IDENTIFIER ON
 
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ADD_EMPLOYEE]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[ADD_EMPLOYEE]
@EmpId varchar(20)
AS

BEGIN

DECLARE @ExistingRowCount INT = 0
SELECT @ExistingRowCount = isnull(Max(CONVERT(numeric,Code)),0) FROM [@Z_EMP_LEAVE]
INSERT INTO [@Z_EMP_LEAVE] (Code, Name, U_Z_EmpID,U_Z_LeaveCode,U_Z_LeaveName,U_Z_GLACC,U_Z_GLACC1,U_Z_PaidLeave,U_Z_OB,U_Z_OBYear,U_Z_OBAmt)
SELECT
Code = RIGHT(''00000000'' +CAST(ROW_NUMBER() OVER(ORDER BY H.Code ) +@ExistingRowCount AS NVARCHAR),8),
Name = RIGHT(''00000000'' +CAST(ROW_NUMBER() OVER(ORDER BY H.Code) +@ExistingRowCount AS NVARCHAR),8),
@EmpId,H.Code,H.Name,H.U_Z_GLACC,H.U_Z_GLACC1,H.U_Z_PaidLeave,0,0,0
FROM
[@Z_PAY_LEAVE] H
where Code not in (Select U_Z_LeaveCode from [@Z_EMP_LEAVE] where U_Z_EMPID=@EmpId )
END
' 
END
 
/****** Object:  StoredProcedure [dbo].[ADD_WORKSHEET]    Script Date: 12/02/2014 09:36:40 ******/
SET ANSI_NULLS ON
 
SET QUOTED_IDENTIFIER ON
 
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ADD_WORKSHEET]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'

--exec [ADD_WORKSHEET] ''00000121'',2018,2
CREATE PROCEDURE [dbo].[ADD_WORKSHEET]
@RefCode varchar(20),
@aYear Int,
@aMonth Int,
@aCompany Varchar(100)
AS

BEGIN

DECLARE @ExistingRowCount INT = 0
 SELECT @ExistingRowCount = isnull(Max(CONVERT(numeric,Code)),0) FROM [@Z_PAYROLL1]
 --print @existingRowCount
INSERT INTO [@Z_PAYROLL1] (Code, Name,
U_Z_empid,U_Z_EmpName,U_Z_JobTitle,U_Z_CardCode,U_Z_PersonalID,U_Z_Department
,U_Z_EmpBranch,U_Z_TermCode,U_Z_SalaryType,U_Z_CostCentre,U_Z_Startdate,U_Z_TermDate 
,U_Z_Branch,U_Z_Dept,U_Z_Dim3,U_Z_Dim4,U_Z_Dim5,U_Z_MONTH,U_Z_YEAR,U_Z_RefCode,U_Z_CompNo
,U_Z_13th,U_Z_14th,U_Z_ExtraSalary,U_Z_TANO,U_Z_OffCycle,U_Z_EOS1,U_Z_Leave ,U_Z_Ticket,U_Z_Saving,U_Z_PaidExtraSalary 
,U_Z_PrjCode,U_Z_BankName,U_Z_GOVAMT,U_Z_TermName)

SELECT
Code = RIGHT(''00000000'' +CAST(ROW_NUMBER() OVER(ORDER BY T0.U_Z_EmpID ) +@ExistingRowCount AS NVARCHAR),8),
Name = RIGHT(''00000000'' +CAST(ROW_NUMBER() OVER(ORDER BY T0.U_Z_EmpID) +@ExistingRowCount AS NVARCHAR),8),
T0.U_Z_EmpID,isnull(T0.[firstName],'''') + '' '' + isnull(T0.[MiddleName],'''') + '' '' + isnull(T0.[LastName],''''),
t0.jobTitle,T0.U_Z_CardCode,T0.govID,T1.Name ,T4."Name",
ISNULL(U_Z_Terms,''''),T0.salaryUnit ,T2.PrcName,T0.startDate,T0.termDate 
,ISNULL(T0.U_Z_Cost,''''),ISNULL(T0.U_Z_Dept,''''),ISNULL(U_Z_Dim3,''''),ISNULL(U_Z_Dim4,''''),ISNULL(U_Z_Dim5,'''')
,@aYear,@aMonth,@RefCode,@aCompany,isnull(T5.U_Z_13th,''0'') ,ISNULL(T5.U_Z_14th,''0'')
,CASE T5.U_Z_ExtraSalary when ''0'' then  ''0'' when ''1'' then ''1'' when ''2'' then ''1'' else ''2'' end
,T0.U_Z_EmpID,''N'',ISNULL(T0.U_Z_EOS1,''N''),ISNULL(T0.U_Z_Leave,''N''),ISNULL(T0.U_Z_Ticket,''N''),ISNULL(T0.U_Z_Saving,''N''),ISNULL(T0.U_Z_ExtraSalary,''N'')
,isnull(T0.U_Z_PrjCode,''''),isnull(T6.[BankName],''N/A''),T0.U_Z_GOVAMT,T7.U_Z_Name  

FROM OHEM T0  Left Outer JOIN OUDP T1 ON T0.dept = T1.Code Left Outer JOIN OPRC T2 ON T0.U_Z_COST = T2.PrcCode left outer join OUBR T4 ON T0."branch" = T4."Code"
left outer JOIN [@Z_OADM] T5 on T5.U_Z_CompCode=@aCompany 
 left outer JOIN ODSC T6 ON T0.bankCode = T6.BankCode
 left outer Join [@Z_PAY_TERMS] T7 on T7.U_Z_Code=T0.U_Z_Terms 

where   T0.Active=''Y'' and  (T0.U_Z_CompNo=''KCC'') and ( isnull(T0.StartDate,''2016-12-01'') <=''2017-01-31'') order by empid
END


--SELECT T0.[empID], isnull(T0.[firstName],'''')+ '' '' + isnull(T0.[MiddleName],'''') + '' '' + isnull(T0.[LastName],'''') ''Emplopyee name'', T0.[jobTitle],T1.[Name], T0.[salary], T0.[salaryUnit], isnull(T2.[PrcName],''''),T0.[StartDate],T0.[TermDate] ,isnull(T0.U_Z_Cost,'''') ''Dim1'' , isnull(T0.U_Z_Dept,'''') ''Dim2'' ,T0.govID  ''PersonalID'' ,T0.U_Z_EmpID ''TANO'',isnull(U_Z_Terms,'''') ''Terms'',T0.U_Z_CardCode ''CustomerCode'',isnull(T0.U_Z_Dim3,'''') ''Dim3'',isnull(T0.U_Z_Dim4,'''') ''Dim4'',isnull(T0.U_Z_Dim5,'''') ''Dim5'',T4."Name" ''EmpBranch'' FROM OHEM T0  Left Outer JOIN OUDP T1 ON T0.dept = T1.Code Left Outer JOIN OPRC T2 ON T0.U_Z_COST = T2.PrcCode left outer join OUBR T4 ON T0."branch" = T4."Code"  where   T0.Active=''Y'' and  (U_Z_CompNo=''KCC'') and ( isnull(T0.StartDate,''2016-12-01'') <=''2017-01-31'') order by empid

' 
END
 
/****** Object:  StoredProcedure [dbo].[UpdateSavingScheme]    Script Date: 12/02/2014 09:36:40 ******/
SET ANSI_NULLS ON
 
SET QUOTED_IDENTIFIER ON
 
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[UpdateSavingScheme]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'  
CREATE PROCEDURE [dbo].[UpdateSavingScheme]  
 @EmpID  nvarchar(100)  
AS  
BEGIN  
 SET NOCOUNT ON;  
Declare @strRefCode As nvarchar(200), @strEmpRefcode As nvarchar(200), @strsql As nvarchar(200), @stTemp As nvarchar(200), @strEmpID As nvarchar(200)  
Declare @dblEmpProPer as Numeric(18,3) , @dblCmpProPeras Numeric(18,3) , @dblEmpConOB as Numeric(18,3), @dblEmpProOB as Numeric(18,3), @dblCmpConOB as Numeric(18,3), @dblCmpProOBas Numeric(18,3), @dblCmpCon as Numeric(18,3), @dblCmpPro as Numeric(18,3
), @dblEmpCon as Numeric(18,3), @dblEmpPro as Numeric(18,3)  
Declare @dblCmpProOB as numeric(18,3)  

if exists (Select U_Z_EmpConBalOB,U_Z_EmpConProOB,U_Z_CmpConBalOB,U_Z_CmpConProOB from OHEM where empID=@EmpID )  
begin  

set @dblempConOB=(Select  U_Z_EmpConBalOB from OHEM where empID=@EmpID )  
set @dblEmpProOB=( Select U_Z_EmpConProOB from OHEM where empID=@EmpID)
set @dblCmpConOB=( Select U_Z_CmpConBalOB from OHEM where empID=@EmpID) 
set @dblCmpProOB=( Select U_Z_CmpConProOB from OHEM where empID=@EmpID)


set @dblEmpCon=(Select  Sum(isnull(U_Z_EmpConBal,0))  from [@Z_PAY_EMP_OSAV] where U_Z_EmpID=@EmpID )  
set @dblEmpPro=(Select  Sum(isnull(U_Z_EmpConPro,0))  from [@Z_PAY_EMP_OSAV] where U_Z_EmpID=@EmpID ) 
set @dblCmpCon=(Select  Sum(isnull(U_Z_CmpConBal,0))  from [@Z_PAY_EMP_OSAV] where U_Z_EmpID=@EmpID )  
set @dblCmpPro=(Select  Sum(isnull(@dblCmpPro,0))  from [@Z_PAY_EMP_OSAV] where U_Z_EmpID=@EmpID )  


set @dblEmpCon = @dblEmpCon + @dblEmpConOB  
set @dblEmpPro = @dblEmpPro + @dblEmpProOB  
set @dblCmpCon = @dblCmpPro + @dblCmpConOB  
set @dblCmpPro = @dblCmpPro + @dblCmpProOB  
Update OHEM set U_Z_EmpConBal=@dblEmpCon,U_Z_EmpConPro=@dblEmpPro ,U_Z_CmpConBal=@dblCmpCon ,U_Z_CmpConPro=@dblCmpPro  where empID= @strEmpID   
end  
  
END  ' 
END
 
/****** Object:  StoredProcedure [dbo].[UpdatePayrollTotal_Payroll_Employee]    Script Date: 12/02/2014 09:36:40 ******/
SET ANSI_NULLS ON
 
SET QUOTED_IDENTIFIER ON
 
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[UpdatePayrollTotal_Payroll_Employee]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'    
CREATE PROCEDURE [dbo].[UpdatePayrollTotal_Payroll_Employee]    
@Year Int,@Month int,    
@RefCode varchar(100),    
@EmpID varchar(20)    
AS    
    
BEGIN    
Declare @StrRefCode varchar(20)    
Declare @strEmpRefcode varchar(20)    
Declare @strsql varchar(20)    
Declare @strcompanycode varchar(20)    
Declare @dblearning Decimal(19,2)    
Declare @dblDeduction Decimal(19,2)    
Declare @dblContribution Decimal(19,2)    
Declare @dblUnPaid  Decimal(19,3)    
Declare @dblPaid  Decimal(19,3)    
Declare @dblAcramt  Decimal(19,3)    
Declare @dblAnnualLeave  Decimal(19,3)    
Declare @dblNetPayamt Decimal(19,3)    
Declare @dblCmpPayamt  Decimal(19,3)    
Declare @dblAirAmt Decimal(19,3)    
Declare @dblAcrAirAmt Decimal(19,3)    
    
Update [@Z_PAYROLL2] set U_Z_Amount=U_Z_Rate*U_Z_Value where U_Z_RefCode=@RefCode    
Update [@Z_PAYROLL3] set U_Z_Amount=U_Z_Rate*U_Z_Value where U_Z_RefCode=@RefCode    
Update [@Z_PAYROLL4] set U_Z_Amount=U_Z_Rate*U_Z_Value where U_Z_RefCode=@RefCode    
Update [@Z_PAYROLL5] set  U_Z_Amount=((U_Z_DedRate *U_Z_DailyRate)/100 * U_Z_Redim)  where U_Z_PaidLeave<>''H'' and  U_Z_RefCode=@RefCode    
Update [@Z_PAYROLL5] set  U_Z_Amount=((U_Z_DedRate *U_Z_DailyRate)/100 * U_Z_Redim)  where U_Z_PaidLeave<>''H'' and  U_Z_RefCode=@RefCode    
Update [@Z_PAYROLL5] set  U_Z_Amount=((U_Z_DedRate *U_Z_DailyRate)/100 * U_Z_Redim)  where U_Z_PaidLeave<>''H'' and  U_Z_RefCode=@RefCode    
Update [@Z_PAYROLL5] set  U_Z_Amount=(((U_Z_DedRate * U_Z_DailyRate)/100)/2) * U_Z_Redim where  U_Z_PaidLeave=''H'' and  U_Z_RefCode=@RefCode  
Update [@Z_PAYROLL5] set  U_Z_Amount=((U_Z_DedRate *U_Z_DailyRate)/100 * (U_Z_Redim-isnull(U_Z_ExDays,0)))  where U_Z_PaidLeave<>''H''  and (U_Z_PaidLeave=''A'' or U_Z_PaidLeave=''P'') and U_Z_RefCode=@RefCode    
Update [@Z_PAYROLL5] set  U_Z_Amount=0 where isnull(U_Z_Basic,''N'')=''Y'' and (U_Z_PaidLeave=''A'' or U_Z_PaidLeave=''P'') and  U_Z_RefCode=@RefCode    
Update [@Z_PAYROLL2] set U_Z_Amount=Round(U_Z_Amount,3) where U_Z_RefCode=@RefCode    
Update [@Z_PAYROLL3] set U_Z_Amount=Round(U_Z_Amount,3) where U_Z_RefCode=@RefCode    
Update [@Z_PAYROLL4] set U_Z_Amount=Round(U_Z_Amount,3) where U_Z_RefCode=@RefCode    
Update [@Z_PAYROLL5] set U_Z_Amount=Round(U_Z_Amount,3) where U_Z_RefCode=@RefCode    
Update [@Z_PAYROLL5] set U_Z_AcrAmount=Round(U_Z_AcrAmount,3) where U_Z_RefCode=@RefCode    
Update [@Z_PAYROLL6] set U_Z_Amount=Round(U_Z_Amount,3) ,U_Z_CurAMount=Round(U_Z_CurAMount,3) where U_Z_RefCode=@RefCode    
    
Declare  @blnAccural  integer = 1    
Declare @debType VarChar(1)    
Declare @LeavePayment Decimal(19,3)    
If Exists(Select *,isnull(U_Z_DedType,''Y'') ''DedInclude'' from [@Z_PAYROLL1] where Code=@RefCode)    
BEGIN    
Set @debType = (Select isnull(U_Z_DedType,''Y'') ''DedInclude'' from [@Z_PAYROLL1] where Code=@RefCode )    
--print @debType    
if(@debType = ''N'')    
 Set @blnAccural = 0    
    
Set @LeavePayment = (Select Sum(U_Z_Balance * U_Z_DailyRate) from [@Z_Payroll5] where U_Z_RefCode = @RefCode and  U_Z_PaidLeave=''A'')    
    
Update [@Z_PAYROLL2] set  U_Z_Value= @LeavePayment  where U_Z_Type=''L'' and  U_Z_RefCode = @RefCode     
Update [@Z_PAYROLL2] set U_Z_Amount=U_Z_Rate*U_Z_Value where  U_Z_RefCode=@RefCode    
Update [@Z_PAYROLL2] set U_Z_Amount=Round(U_Z_Amount,3) where  U_Z_RefCode=@RefCode    
    
    
Set @dblearning = isnull((Select Sum(isnull(U_Z_Amount,0))  from [@Z_Payroll2] where U_Z_RefCode=@RefCode),0)    
Set @dblDeduction = isnull((Select Sum(isnull(U_Z_Amount,0))  from [@Z_Payroll3] where U_Z_RefCode=@RefCode),0)    
Set @dblContribution = isnull((Select Sum(isnull(U_Z_Amount,0))  from [@Z_Payroll4] where U_Z_RefCode=@RefCode),0)    
Set @dblUnPaid = isnull((Select Sum(isnull(U_Z_Amount,0))  from [@Z_Payroll5] where U_Z_RefCode = @RefCode and  U_Z_PaidLeave<>''P'' and U_Z_PaidLeave<>''A''),0)    
Set @dblPaid = isnull((Select Sum(isnull(U_Z_Amount,0))  from [@Z_Payroll5] where U_Z_RefCode=@RefCode and  U_Z_PaidLeave=''P''),0)    
Set @dblAcramt = isnull((Select Sum(isnull(U_Z_CurAmount,0))  from [@Z_Payroll5] where U_Z_RefCode=@RefCode and  U_Z_PaidLeave=''A''),0)    
    
if(@blnAccural = 0)    
 Set @dblAcramt = 0    
     
Update [@Z_PAYROLL1] set  U_Z_UnPaidLeave = @dblUnPaid where Code=@RefCode    
    
Set @dblAnnualLeave  =isnull( (Select Sum(U_Z_Amount) from [@Z_Payroll5] where  U_Z_RefCode=@RefCode and  U_Z_PaidLeave=''A'') ,0)    
    
 set @dblAirAmt=isnull(( Select   Sum(U_Z_Amount) from [@Z_Payroll6] where U_Z_RefCode=@RefCode ) ,0)  
   set @dblNetPayamt=isnull(( Select   Sum(U_Z_NetPayAmt) from [@Z_Payroll6] where U_Z_RefCode=@RefCode ) ,0)  
    set @dblCmpPayamt=isnull(( Select   Sum(U_Z_CmpPayAmt) from [@Z_Payroll6] where U_Z_RefCode=@RefCode ),0)   
     set @dblAcrAirAmt=isnull(( Select   Sum(U_Z_CurAmount) from [@Z_Payroll6] where U_Z_RefCode=@RefCode ) ,0)  
       
    
    
    
    
    
    
if(@blnAccural = 0)    
Begin    
 Set @dblAcramt = 0    
 Set @dblAcrAirAmt = 0    
End    
    
    
Update [@Z_PAYROLL1] set    
    
U_Z_Earning=@dblearning ,    
U_Z_Deduction=@dblDeduction ,    
U_Z_Contri=@dblContribution ,    
U_Z_UnPaidLeave=@dblUnPaid ,    
U_Z_PaidLeave=@dblPaid ,    
U_Z_AcrAmt=@dblAcramt ,    
U_Z_AnuLeave=@dblAnnualLeave,    
U_Z_AcrAirAmt=@dblAcrAirAmt,    
U_Z_NetPayAmt=@dblNetPayamt,    
U_Z_CmpPayAmt=@dblCmpPayamt,    
U_Z_AirAmt=@dblAirAmt     
where Code=@RefCode     
    
    
    
    
    
Update [@Z_PAYROLL1] set  U_Z_Cost= isnull(U_Z_AnuLeave,0) +isnull(U_Z_CashOutAmt,0)+     
isnull(U_Z_MonthlyBasic,0)+isnull(U_Z_Earning,0) + isnull(U_Z_PaidLeave,0)+isnull(U_Z_Contri,0)+ isnull(U_Z_CmpPayAmt,0) +isnull(U_Z_NetPayAmt,0)     
where Code=@RefCode    
Update [@Z_PAYROLL1] set  U_Z_Cost=Round(U_Z_Cost,3)  where Code=@RefCode    
Update  [@Z_PAYROLL1] set  U_Z_NetSalary=isnull(U_Z_MonthlyBasic,0)+isnull(U_Z_Earning,0) - isnull(U_Z_Deduction,0)-isnull(U_Z_UnPaidLeave,0)+ isnull(U_Z_AnuLeave,0)+isnull(U_Z_NetPayAmt,0) +isnull(U_Z_CashOutAmt,0)  where Code=@RefCode    
Update [@Z_PAYROLL1] set  U_Z_NetSalary=Round(U_Z_NetSalary,3)  where Code=@RefCode       
    
    
END    
    
END' 
END
 
/****** Object:  StoredProcedure [dbo].[UPDATEPAYROLLTABLE]    Script Date: 12/02/2014 09:36:40 ******/
SET ANSI_NULLS ON
 
SET QUOTED_IDENTIFIER ON
 
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[UPDATEPAYROLLTABLE]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE procedure [dbo].[UPDATEPAYROLLTABLE]
@Year Int,@Month int,    
 @aRefcode varchar(100)  
 

as

BEGIN
Declare @PayEmpRefCode varchar(20)   
Declare @EmpID varchar(20)    

Declare UPDATEPAYROLLDETAILS Cursor For Select "Code","U_Z_EmpID" from "@Z_PAYROLL1" where "U_Z_RefCode"=@aRefcode
			Open UPDATEPAYROLLDETAILS
			  fetch Next from UPDATEPAYROLLDETAILS  Into @PayEmpRefCode,@EMPID
				    While @@Fetch_Status=0
				    
				    
BEGIN

					--Exec UPDATELEAVEBALANCE_TRANSACTION_EMPLOYEE @EmpID,@year,@month
				    Exec UpdatePayrollTotal_Payroll_Employee @Year,@Month,@PayEmpRefCode,@EmpID
				    Exec RESET_AIRTICKET_LOAN @EmpID,@year,@month
				    Exec RESET_AIRTICKET_LOAN1 @EmpID,@year,@month
fetch Next from UPDATEPAYROLLDETAILS   Into @PayEmpRefCode,@EMPID
end
Close UPDATEPAYROLLDETAILS
Deallocate UPDATEPAYROLLDETAILS

END' 
END
 
/****** Object:  StoredProcedure [dbo].[ADD_LEAVEDETAILS]    Script Date: 12/02/2014 09:36:40 ******/
SET ANSI_NULLS ON
 
SET QUOTED_IDENTIFIER ON
 
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ADD_LEAVEDETAILS]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'--delete from [@Z_PAYROLL5]  where U_Z_RefCode=''00000292''     
--exec ADD_LEAVEDETAILS ''00000292'',2014,7    
    
CREATE Procedure [dbo].[ADD_LEAVEDETAILS]        
        
@aRefCode Varchar(30),        
@aYear Int,        
@aMonth Int        
AS        
         
BEGIN        
        
Declare @strRefCode varchar(max)        
Declare @strEmpRefcode varchar(max)        
Declare @strempID varchar(max)        
Declare @strPayrollRefNo varchar(max)        
Declare @strGLacc varchar(max)        
Declare @strEname varchar(max)        
Declare @strECode varchar(max)        
Declare @strCode varchar(1)        
        
Declare @dblTotalBasic Decimal(19,3)        
Declare @dblEmpBasic Decimal(19,3)        
Declare @dtPayrollDate Date        
Declare @blnTerm varchar(10)        
Declare @dblWorkingDays decimal(19,3)        
Declare @dblCalenderDays Decimal(19,3)        
Declare @strTems varchar(30)        
Declare @stOVStartDate Varchar(20)        
Declare @stOVEndDate varchar(20)        
Declare @stOVType Varchar(30)        
Declare @dblYearofExperience Decimal(19,3)        
Declare @dblNoofDays1 Decimal(19,3)        
if Exists (Select "Code" from "@Z_PAYROLL1" where "Code"=@aRefCode )        
BEGIN        
     
Set @strPayrollRefNo=@aRefCode         
set @strempID =(Select "U_Z_empid" from "@Z_PAYROLL1" where "Code"= @aRefCode)        
set @dblTotalBasic   =(Select "U_Z_BasicSalary" from "@Z_PAYROLL1" where "Code"= @aRefCode)        
set @dblEmpBasic=@dblTotalBasic         
set @dtPayrollDate  =(Select "U_Z_PayDate" from "@Z_PAYROLL1" where "Code"= @aRefCode)        
set @blnTerm  =(Select "U_Z_IsTerm" from "@Z_PAYROLL1" where "Code"= @aRefCode)        
set @dblCalenderDays  =(Select "U_Z_Calenderdays" from "@Z_PAYROLL1" where "Code"= @aRefCode)        
set @dblWorkingDays  =(Select "U_Z_WorkingDays" from "@Z_PAYROLL1" where "Code"= @aRefCode)        
set @strTems   =(Select "U_Z_TermCode" from "@Z_PAYROLL1" where "Code"= @aRefCode)        
set @dblYearofExperience   =(Select "U_Z_YOE" from "@Z_PAYROLL1" where "Code"= @aRefCode)        
        
if not exists (Select "Code" from "@Z_PAYROLL5" where "U_Z_RefCode"=@aRefCode )        
BEGIN        
Exec ADD_EMPLOYEE @strempid         
DECLARE @ExistingRowCount INT = 0          
        
if exists (Select "U_Z_LeaveCode" from "@Z_PAY_OALMP" where "U_Z_Terms"=@strTems )        
        
BEGIN     
     
       
SELECT @ExistingRowCount = isnull(Max(CONVERT(numeric,Code)),0) FROM [@Z_PAYROLL5]          
INSERT INTO [@Z_PAYROLL5] (Code, Name, U_Z_Refcode,U_Z_EmpID,U_Z_LeaveCode,U_Z_LeaveName,U_Z_Year,U_Z_Month      
,U_Z_Basic,U_Z_PostType,U_Z_GLACC,U_Z_GLACC1,U_Z_DedRate,U_Z_PaidLeave,U_Z_Amount,U_Z_OB,U_Z_OBAmt,U_Z_CM,U_Z_CMAmt,U_Z_NoofDays    
,U_Z_TotalAvDays,U_Z_DailyRate,U_Z_CurAMount,U_Z_Increment,U_Z_AcrAmount,U_Z_Redim ,U_Z_Balance,U_Z_BalanceAmt,U_Z_YTDAMount,U_Z_Adjustment,U_Z_CashOutAmt,U_Z_CashoutDays,U_Z_EnCashment          )           
SELECT          
Code = RIGHT(''00000000'' +CAST(ROW_NUMBER() OVER(ORDER BY H.Code ) +@ExistingRowCount AS NVARCHAR),8),          
Name = RIGHT(''00000000'' +CAST(ROW_NUMBER() OVER(ORDER BY H.Code) +@ExistingRowCount AS NVARCHAR),8),          
@strPayrollRefNo , @strempID ,H.U_Z_LeaveCode ,H.U_Z_LeaveName,@aYear,@aMonth       
 ,isnull(N1.U_Z_Basic,''N'') ,    
 CASE N1.U_Z_PaidLeave when ''A'' then ''C'' else ''D'' end,      
 case ISNULL(N2.U_Z_GLACC,'''') when '''' then N1.U_Z_GLACC else N2.U_Z_GLACC End,      
 case ISNULL(N2.U_Z_GLACC1,'''') when '''' then N1.U_Z_GLACC1 else N2.U_Z_GLACC1  END,    
 N1.U_Z_DedRate,    
 N1.U_Z_PaidLeave       
 ,0  ,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0  
FROM          
[@Z_EMP_LEAVE] H       
Inner Join OHEM N2 on N2.empID=H.U_Z_EmpID       
LEFT OUTER JOIN [@Z_PAY_LEAVE] N1 On N1.Code=H.U_Z_LeaveCode       
where H."U_Z_EmpID" = @strempID and N1.Code in  (Select "U_Z_LeaveCode" from "@Z_PAY_OALMP" where "U_Z_Terms"=@strTems )         END        
        
Else        
        
BEGIN        
        
--select * from [@Z_PAYROLL5]   
SELECT @ExistingRowCount = isnull(Max(CONVERT(numeric,Code)),0) FROM [@Z_PAYROLL5]          
INSERT INTO [@Z_PAYROLL5] (Code, Name, U_Z_Refcode,U_Z_EmpID,U_Z_LeaveCode,U_Z_LeaveName,U_Z_Year,U_Z_Month      
,U_Z_Basic,U_Z_PostType,U_Z_GLACC,U_Z_GLACC1,U_Z_DedRate,U_Z_PaidLeave,U_Z_Amount,U_Z_OB,U_Z_OBAmt,U_Z_CM,U_Z_CMAmt,U_Z_NoofDays    
,U_Z_TotalAvDays,U_Z_DailyRate,U_Z_CurAMount,U_Z_Increment,U_Z_AcrAmount,U_Z_Redim ,U_Z_Balance,U_Z_BalanceAmt,U_Z_YTDAMount,U_Z_Adjustment,U_Z_CashOutAmt,U_Z_CashoutDays,U_Z_EnCashment          )           
SELECT          
Code = RIGHT(''00000000'' +CAST(ROW_NUMBER() OVER(ORDER BY H.Code ) +@ExistingRowCount AS NVARCHAR),8),          
Name = RIGHT(''00000000'' +CAST(ROW_NUMBER() OVER(ORDER BY H.Code) +@ExistingRowCount AS NVARCHAR),8),          
@strPayrollRefNo , @strempID ,H.U_Z_LeaveCode ,H.U_Z_LeaveName,@aYear,@aMonth       
 ,isnull(N1.U_Z_Basic,''N'') ,    
 CASE N1.U_Z_PaidLeave when ''A'' then ''C'' else ''D'' end,      
 case ISNULL(N2.U_Z_GLACC,'''') when '''' then N1.U_Z_GLACC else N2.U_Z_GLACC End,      
 case ISNULL(N2.U_Z_GLACC1,'''') when '''' then N1.U_Z_GLACC1 else N2.U_Z_GLACC1  END,    
 N1.U_Z_DedRate,    
 N1.U_Z_PaidLeave       
 ,0  ,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0  
FROM          
[@Z_EMP_LEAVE] H       
Inner Join OHEM N2 on N2.empID=H.U_Z_EmpID       
LEFT OUTER JOIN [@Z_PAY_LEAVE] N1 On N1.Code=H.U_Z_LeaveCode       
where H."U_Z_EmpID" = @strempID         
END        
         
        
END        
        
        
END        
        
        
END        
' 
END
 
/****** Object:  StoredProcedure [dbo].[UpdateEmployeeLeavedetails_Employee1]    Script Date: 12/02/2014 09:36:40 ******/
SET ANSI_NULLS ON
 
SET QUOTED_IDENTIFIER ON
 
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[UpdateEmployeeLeavedetails_Employee1]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'
CREATE Procedure [dbo].[UpdateEmployeeLeavedetails_Employee1]  
  
@strPayRefCode Varchar(30),  
@IntYear Integer,  
@IntMonth Integer  
  
As  
  
Begin  
Declare @EmpID Varchar(30) 
Declare @strPayroll1RefCode varchar(30) 
Declare PurchaseCUR12 Cursor For Select U_Z_EMPID,Code from [@Z_PAYROLL1] where U_Z_RefCode=@strPayRefCode  
   Open PurchaseCur12  
     fetch Next from PurchaseCur12   Into @EmpID  ,@strPayroll1RefCode
        While @@Fetch_Status=0  
        BEgin  
			--  Exec UpdatePayrollTotal_Payroll_Employee @intyear,@intmonth,@strPayroll1RefCode,@EmpID
              Exec  UpdateSavingScheme @EmpID  
			  Exec  UPDATELEAVEBALANCE_TRANSACTION_EMPLOYEE @EmpID,@IntYear,@IntMonth  
			  Exec  RESET_AIRTICKET_LOAN @EmpID,@IntYear,@IntMonth  
			  Exec  RESET_AIRTICKET_LOAN1 @EmpID,@IntYear,@IntMonth  
        fetch Next from PurchaseCur12   Into @EmpID, @strPayroll1RefCode  
     end  
     Close PurchaseCur12  
     Deallocate PurchaseCur12 
 END' 
END
 
/****** Object:  StoredProcedure [dbo].[UpdateEmployeeLeavedetails_EMployee_Month_Company]    Script Date: 12/02/2014 09:36:40 ******/
SET ANSI_NULLS ON
 
SET QUOTED_IDENTIFIER ON
 
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[UpdateEmployeeLeavedetails_EMployee_Month_Company]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'
--select * from [@Z_PAYROLL] 

--exec UpdateEmployeeLeavedetails_EMployee_Month_Company ''00000052'',2014,1
CREATE PROCEDURE [dbo].[UpdateEmployeeLeavedetails_EMployee_Month_Company]

@CODE VARCHAR(30),
@Year int,
@Month int

AS

BEGIN
Declare @strRefCode VarChar(30)

DECLARE vendor_cursor CURSOR FOR 
SELECT U_Z_EMPID
FROM [@Z_PAYROLL1]
WHERE U_Z_REFCODE=@Code
ORDER BY U_Z_EMPID;

OPEN vendor_cursor

FETCH NEXT FROM vendor_cursor 
INTO @strRefCode

WHILE @@FETCH_STATUS = 0
BEGIN
   set @strrefcode=''1''
    exec UPDATELEAVEBALANCE_TRANSACTION_EMPLOYEE @strrefcode,@year,@month
    exec RESET_AIRTICKET_LOAN   @strrefcode,@year,@month
    exec RESET_AIRTICKET_LOAN1   @strrefcode,@year,@month
    exec UpdateSavingScheme @strRefCode
    FETCH NEXT FROM vendor_cursor 
    INTO @strRefCode
END 
CLOSE vendor_cursor;
DEALLOCATE vendor_cursor;

END' 
END
 
/****** Object:  StoredProcedure [dbo].[UpdateEmployeeLeavedetails_Employee]    Script Date: 12/02/2014 09:36:40 ******/
SET ANSI_NULLS ON
 
SET QUOTED_IDENTIFIER ON
 
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[UpdateEmployeeLeavedetails_Employee]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'

Create Procedure [dbo].[UpdateEmployeeLeavedetails_Employee]  
  
@strPayRefCode Varchar(30),  
@IntYear Integer,  
@IntMonth Integer  
  
As  
  
Begin  
Declare @EmpID Varchar(30)  
Declare PurchaseCUR1 Cursor For Select U_Z_EMPID from [@Z_PAYROLL1] where U_Z_RefCode=@strPayRefCode  
   Open PurchaseCur1  
     fetch Next from PurchaseCur1   Into @EmpID  
        While @@Fetch_Status=0  
        BEgin  
      Exec UpdateSavingScheme @EmpID  
      Exec  UPDATELEAVEBALANCE_TRANSACTION_EMPLOYEE @EmpID,@IntYear,@IntMonth  
      Exec  RESET_AIRTICKET_LOAN @EmpID,@IntYear,@IntMonth  
      Exec  RESET_AIRTICKET_LOAN1 @EmpID,@IntYear,@IntMonth  
        fetch Next from PurchaseCur1   Into @strPayRefCode  
     end  
     Close PurchaseCur1  
       Deallocate PurchaseCur1  
  
END' 
END
 
/****** Object:  StoredProcedure [dbo].[UPDATE_EMPLOYEE_LEAVEDETAILS]    Script Date: 12/02/2014 09:36:40 ******/
SET ANSI_NULLS ON
 
SET QUOTED_IDENTIFIER ON
 
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[UPDATE_EMPLOYEE_LEAVEDETAILS]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'--Exec UPDATE_EMPLOYEE_LEAVEDETAILS ''00002838'',2014,2   
    
CREATE  Procedure [dbo].[UPDATE_EMPLOYEE_LEAVEDETAILS]    
@aRefCode varchar(30),    
@aYear int,    
@aMonth int    
    
AS    
    
BEGIN    
Declare @payroll1RefCode Varchar(30)    
Declare @PayEmpRefCode varchar(30)    
Declare @strRefCode varchar(max)    
Declare @strEmpRefcode varchar(max)    
Declare @strempID varchar(max)    
Declare @strPayrollRefNo varchar(max)    
Declare @strGLacc varchar(max)    
Declare @strEname varchar(max)    
Declare @strECode varchar(max)    
Declare @strCode varchar(1)    
    
Declare @dblTotalBasic Decimal(19,3)    
Declare @dblEmpBasic Decimal(19,3)    
Declare @dtPayrollDate Date    
Declare @blnTerm varchar(10)    
Declare @dblWorkingDays decimal(19,3)    
Declare @dblCalenderDays Decimal(19,3)    
Declare @strTems varchar(30)    
Declare @stOVStartDate Varchar(20)    
Declare @stOVEndDate varchar(20)    
Declare @stOVType Varchar(30)    
Declare @dblYearofExperience Decimal(19,3)    
Declare @dblNoofDays1 Decimal(19,3)    
Declare @dblCarriedForward Decimal(19,3)    
Declare @dblYearly Decimal(19,3)    
Declare @dblOpeningBalanced Decimal(19,3)    
Declare @dblTransaction Decimal(19,3)    
Declare @dblAdjustment Decimal(19,3)    
Declare @dblAccurred Decimal(19,3)    
Declare @dblClosingBalance Decimal(19,3)    
Declare @dblUnPostedTrns Decimal(19,3)    
Declare @strLeaveCode Varchar(50)    
Declare @dblOpeningBalance Decimal(19,3)    
Declare @isAccured Varchar(1)    
Declare @dblYearlyEntitled Decimal(19,3)    
Declare @strOVStartDate Varchar(20)    
Declare @strOVEndDate Varchar(20)    
Declare @JoiningDate DateTime
 Declare @blnHREnt Varchar(Max)
  Declare @blnProrated1 Varchar(1)
  Declare @blnAccured1 VarChar(1)
  Declare @SalaryCode Varchar(Max)
  Declare @dblYe Decimal(19,2)
  Declare @dblTotalyearlyEnttile Decimal(19,2)
 Declare @intRemainingMonths int=1   
 Declare @JoiningDay int
 Declare @NDays int                            
 Declare @MonthlyEntitlement Decimal(19,3)   
Declare UpdateEmpLeaveDetails Cursor For Select "Code","U_Z_LeaveCode","U_Z_EmpID","U_Z_RefCode" from "@Z_PAYROLL5" where "U_Z_RefCode"=@aRefcode    
   Open UpdateEmpLeaveDetails    
     fetch Next from UpdateEmpLeaveDetails  Into @PayEmpRefCode,@strLeaveCode,@strEmpID,@payroll1RefCode    
        While @@Fetch_Status=0    
            
BEGIN    
set @dblCarriedForward =(Select isnull(U_Z_CAFWD,0) from [@Z_EMP_LEAVE_BALANCE] where U_Z_EmpID=@strempID  and U_Z_LeaveCode=@strLeaveCode  and U_Z_Year=@aYear )    
set @dblYearly  =(Select isnull(U_Z_Entile,0)  from [@Z_EMP_LEAVE_BALANCE] where U_Z_EmpID=@strempID  and U_Z_LeaveCode=@strLeaveCode  and U_Z_Year=@aYear )    
set @dblClosingBalance =(Select isnull(U_Z_Balance,0) from [@Z_EMP_LEAVE_BALANCE] where U_Z_EmpID=@strempID  and U_Z_LeaveCode=@strLeaveCode  and U_Z_Year=@aYear )    
set @dblOpeningBalance =(Select isnull(U_Z_OB,0)  from [@Z_EMP_LEAVE_BALANCE] where U_Z_EmpID=@strempID  and U_Z_LeaveCode=@strLeaveCode  and U_Z_Year=@aYear )    
set @isAccured=(Select ISNULL(U_Z_Accured,''N'') from "@Z_PAY_LEAVE" where "Code"=@strLeaveCode )    
if @isAccured=''Y''    
set @dblCarriedForward =@dblCarriedForward+@dblOpeningBalance     
else    
set @dblCarriedForward =@dblCarriedForward+@dblYearly     
   set @dblTransaction =isnull((select isnull(sum(isnull(U_Z_Redim,0)),0)  from [@Z_PAYROLL5] where U_Z_Leavecode=@strLeaveCode  and U_Z_month<=@aMonth  and U_Z_Year=@ayear and U_Z_EmpID=@strempID  group by U_Z_EmpID),0)    
 set @dblAdjustment =isnull((select isnull(sum(isnull(U_Z_Adjustment,0)),0)  from [@Z_PAYROLL5] where U_Z_Leavecode=@strLeaveCode  and U_Z_month<=@aMonth  and U_Z_Year=@ayear and U_Z_EmpID=@strempID  group by U_Z_EmpID),0)    
 set @dblAccurred =isnull((select isnull(sum(isnull(U_Z_NoofDays,0)),0)  from [@Z_PAYROLL5] where U_Z_Leavecode=@strLeaveCode  and U_Z_month<=@aMonth  and U_Z_Year=@ayear and U_Z_EmpID=@strempID  group by U_Z_EmpID),0)    
   
set @strOVStartDate   =(Select "U_Z_OVTSTART" from "@Z_PAYROLL1" where "Code"= @payroll1RefCode)    
set @strOVEndDate   =(Select "U_Z_OVTEND" from "@Z_PAYROLL1" where "Code"= @payroll1RefCode)                                
set @dblYearofExperience   =(Select "U_Z_YOE" from "@Z_PAYROLL1" where "Code"= @payroll1RefCode)    
set @strempID =(Select "U_Z_empid" from "@Z_PAYROLL1" where "Code"= @payroll1RefCode)    
set @dblTotalBasic   =(Select "U_Z_BasicSalary" from "@Z_PAYROLL1" where "Code"= @payroll1RefCode)    
set @dblEmpBasic=@dblTotalBasic     
set @dtPayrollDate  =(Select "U_Z_PayDate" from "@Z_PAYROLL1" where "Code"= @payroll1RefCode)    
set @blnTerm  =(Select isnull("U_Z_IsTerm",''N'') from "@Z_PAYROLL1" where "Code"= @payroll1RefCode)    
set @dblCalenderDays  =(Select "U_Z_Calenderdays" from "@Z_PAYROLL1" where "Code"= @payroll1RefCode)    
set @dblWorkingDays  =(Select "U_Z_WorkingDays" from "@Z_PAYROLL1" where "Code"= @payroll1RefCode)    
set @strTems   =(Select "U_Z_TermCode" from "@Z_PAYROLL1" where "Code"= @payroll1RefCode)    
set @dblYearofExperience   =(Select "U_Z_YOE" from "@Z_PAYROLL1" where "Code"= @payroll1RefCode)    
 
 
 
    
if exists (Select T0.DocEntry from [@Z_PAY_ALMP1] T0 inner Join [@Z_PAY_OALMP] T1 on T1.DocEntry=T0.DocEntry where @dblYearofExperience  between U_Z_FromYear and U_Z_ToYear  and T1.U_Z_Terms=@strTems  and T1.U_Z_LeaveCode=@strLeaveCode )    
BEGIN    
Set @dblNoofDays1=(Select T0.U_Z_NoofDays from [@Z_PAY_ALMP1] T0 inner Join [@Z_PAY_OALMP] T1 on T1.DocEntry=T0.DocEntry where @dblYearofExperience  between U_Z_FromYear and U_Z_ToYear  and T1.U_Z_Terms=@strTems  and T1.U_Z_LeaveCode=@strLeaveCode )    
Set @dblYearlyEntitled=@dblNoofDays1     
if @dblYearly>0    
set @dblYearlyEntitled=@dblYearly     
set @dblNoofDays1=@dblYearlyEntitled/12.0    
set @dblNoofDays1=ROUND(@dblNoofDays1,2)    
END    
else    
BEGIN    
set @dblNoofDays1=(Select "U_Z_NoofDays" from "@Z_PAY_LEAVE" where "Code"=@strLeaveCode )    
set @dblYearlyEntitled=(Select U_Z_DaysYear from "@Z_PAY_LEAVE" where "Code"=@strLeaveCode )    
if @dblYearly>0    
set @dblYearlyEntitled=@dblYearly     
set @dblNoofDays1=@dblYearlyEntitled/12.0    
set @dblNoofDays1=ROUND(@dblNoofDays1,2)    
  END   

           
  
 

  set @blnHREnt=(Select ISNULL(U_Z_EntHR,''N'') from "@Z_PAY_LEAVE" where "Code"=@strLeaveCode )     
  set @JoiningDate =(Select startDate from OHEM where empid=@strEmpID)


  set @blnProrated1=(Select   isnull(U_Z_Prorate,''N'') from  [@Z_PAY_LEAVE] where code=@strLeaveCode)  
  set @blnAccured1=(Select   isnull(U_Z_Accured,''N'') from  [@Z_PAY_LEAVE] where Code =@strLeaveCode)  

	
		If Year(@JoiningDate) = @ayear And Month(@JoiningDate) = @aMonth 
				BEGIN
					set @blnTerm=''Y''

                    If @blnProrated1 = ''Y''
					BEGIN
                         If @blnHREnt=''Y''
						   BEGIN
								set @SalaryCode = (Select isnull(U_Z_HR_SalaryCode,'''') from OHEM where empID=@strEmpID)
							
								If @SalaryCode<>'''' 
									set @dblYe= (Select U_Z_Entitle from [@Z_HR_OSALST] where U_Z_SalCode=@SalaryCode and U_Z_LeaveCode=@strLeaveCode)
								else	
								   set @dblYe=0
							
								If @dblYe > 0 
									set   @dblTotalyearlyEnttile = @dblYe
								Else
									set @dblTotalyearlyEnttile = @dblYearlyEntitled
                         END
					
                        set @intRemainingMonths = 12 - @aMonth
                        set @dblTotalyearlyEnttile = @dblTotalyearlyEnttile / 12
						
						set @JoiningDay=Day(@JoiningDate)
						set @NDays=@dblCalenderDays-@JoiningDay+1   
						if @NDays>@dblCalenderDays
							set @NDays=@dblCalenderDays

						set @MonthlyEntitlement=@dblTotalyearlyEnttile
						set @MonthlyEntitlement=(@MonthlyEntitlement/@dblCalenderDays)*@NDays
						set @dblTotalyearlyEnttile = (@dblTotalyearlyEnttile * @intRemainingMonths) + @MonthlyEntitlement

                     	set @dblNoofDays1=ROUND(@dblTotalyearlyEnttile,2)    
						set @dblNoofDays1=@dblTotalyearlyEnttile/(12-@aMonth)    
						set @dblNoofDays1=ROUND(@dblNoofDays1,2) 
					 END

					 else

				BEGIN
					 If @blnHREnt=''Y''
						   BEGIN
								set @SalaryCode = (Select isnull(U_Z_HR_SalaryCode,'''') from OHEM where empID=@strEmpID)
							    If @SalaryCode<>'''' 
									set @dblYe= (Select U_Z_Entitle from [@Z_HR_OSALST] where U_Z_SalCode=@SalaryCode and U_Z_LeaveCode=@strLeaveCode)
								else	
								   set @dblYe=0
							
								If @dblYe > 0 
									set   @dblTotalyearlyEnttile = @dblYe
								Else
									set @dblTotalyearlyEnttile = @dblYearlyEntitled
                           END
					
                        set @dblTotalyearlyEnttile =@dblTotalyearlyEnttile 
                   		set @dblNoofDays1=@dblTotalyearlyEnttile/12 
						set @dblNoofDays1=ROUND(@dblNoofDays1,2) 
					 END
				END


	else

		BEGIN
		BEGIN
                         If @blnHREnt=''Y''
						   BEGIN
								set @SalaryCode = (Select isnull(U_Z_HR_SalaryCode,'''') from OHEM where empID=@strEmpID)
							
								If @SalaryCode<>'''' 
									 set @dblYe= (Select U_Z_Entitle from [@Z_HR_OSALST] where U_Z_SalCode=@SalaryCode and U_Z_LeaveCode=@strLeaveCode)
								else
									  set @dblYe=0
							
								If @dblYe > 0 
								    set   @dblTotalyearlyEnttile = @dblYe
								Else
								   set  @dblTotalyearlyEnttile=@dblYearlyEntitled
							END
					
                    
							  set @dblTotalyearlyEnttile = @dblYearlyEntitled
							  if @dblYearly>0    
								set @dblTotalyearlyEnttile=@dblYearly     
							    set @dblNoofDays1=@dblTotalyearlyEnttile/12.0   
							    set @dblNoofDays1=ROUND(@dblNoofDays1,2)   
				end
END

if @blnTerm=''Y''     
BEGIN    
set @dblNoofDays1=@dblNoofDays1/@dblCalenderDays     
set @dblNoofDays1=@dblNoofDays1 * @dblWorkingDays     
END    
    
Declare @dblU_Z_OB Decimal(19,3)    
Declare @dblU_Z_CM Decimal(19,3)    
Declare @dblU_Z_NoofDays Decimal(19,3)    
Declare @dblRedimdays Decimal(19,3)    
Declare @dblCashout Decimal(19,3)    
Declare @dblAdjustmentDays Decimal(19,3)    
    
if @isAccured=''Y''    
set @dblU_Z_NoofDays=@dblNoofDays1     
else    
Begin
set @dblU_Z_NoofDays=0    
set @dblNoofDays1=0
End    
set @dblU_Z_OB=@dblCarriedForward + @dblAccurred    
set @dblU_Z_CM=@dblCarriedForward + @dblAccurred - @dblTransaction + @dblAdjustment    
    
    
set @dblRedimdays  = (select isnull(sum(U_Z_NoofDays),0)  from [@Z_PAY_OLETRANS] where U_Z_OffCycle=''N'' and  U_Z_Trnscode=@strLeaveCode  and U_Z_month=@aMonth  and U_Z_Year=@aYear  and U_Z_EmpID=@strempID  group by U_Z_EmpID)    
     
    
Declare @dblTACOunt Decimal(19,3)    
 Set @dblTACOunt =  (select isnull(Count(*),0)  from [@Z_TIAT]  where  (U_Z_DateIn between @strOVStartDate  and @strOVEndDate ) and  U_Z_Status=''A''  and (U_Z_LeaveType= @strLeaveCode ) and U_Z_employeeID=@strempID )    
  Set @dblRedimdays = @dblRedimdays +@dblTACOunt     
     
     
    Set @dblCashout=(select isnull(sum(U_Z_NoofDays),0)  from [@Z_PAY_OLADJTRANS] where isnull(U_Z_CashOut,''N'')=''Y'' and  U_Z_Trnscode=@strLeaveCode  and Month(U_Z_StartDate)=@aMonth  and year(U_Z_StartDate)=@aYear  and U_Z_EmpID=@strempID )    
        
                             set @dblAdjustmentDays=(select isnull(sum(U_Z_NoofDays),0)  from [@Z_PAY_OLADJTRANS] where isnull(U_Z_CashOut,''N'')=''N'' and  U_Z_Trnscode=@strLeaveCode  and Month(U_Z_StartDate)=@aMonth  and year(U_Z_StartDate)=@aYear  and   
                             U_Z_EmpID=@strempID )    
    
                  Declare @dblOverTimeAdjustable Decimal(19,3)              
            
                         set @dblOverTimeAdjustable=(select isnull(sum(U_Z_OverTime)/8.00,0) from [@Z_TIAT]  where  Month(U_Z_DateIn)=@amonth and    
                          Year(U_Z_DateIn)=@aYear and  U_Z_Status=''A''  and isnull(U_Z_LeaveBalance,''N'')=''Y''  and U_Z_LeaveType=@strLeaveCode     
                           and U_Z_employeeID=@strempID )    
                          
     set @dblAdjustmentDays = @dblAdjustmentDays + @dblOverTimeAdjustable    
                                
Declare @dblnoofEncashment1 Decimal(19,3)                      
set @dblnoofEncashment1= (select SUM(U_Z_NoofDays) from [@Z_PAY_OLETRANS_OFF] where isnull(U_Z_CashOut,''N'')=''N'' and  U_Z_Posted=''Y'' and  U_Z_EMPID=@strempid  and U_Z_TrnsCode=@strLeaveCode  and U_Z_month <@aMonth   and U_Z_Year=@aYear )    
    
  Declare @DedType Varchar(1)        
   set @DedType=(Select ISNULL(U_Z_DedType,''Y'') from "@Z_PAYROLL1" where "Code"= @payroll1RefCode)    
   if @DedType=''N''    
                    begin    
                    
       set  @dblCashout = 0    
       set  @dblAdjustmentDays = 0    
       set  @dblnoofEncashment1 = 0    
                    end                      
         
    
  Declare @dblBal Decimal(19,3)    
  Set @dblBal = @dblCarriedForward + @dblAccurred - @dblTransaction + @dblAdjustmentDays - @dblnoofEncashment1    
      
      
  Declare @NoofworkingDays Decimal(19,3)    
  Declare @strWorkCode Varchar(30)    
  set @strWorkCode=(Select ISNULL(U_Z_WOrkCode,'''') from OHEM where empID=@strempID )    
  if @strWorkCode=''''     
      
  BEGIN    
   set @NoofworkingDays =(Select U_Z_Days from [@Z_EWO1] T0 inner join [@Z_OEWO] T1 on T1.DocEntry=T0.DocEntry and T1.U_Z_Code=@strWorkCode    
       where T0.U_Z_Month=@aMonth)    
   END    
       
   If @NoofworkingDays =0    
   set @NoofworkingDays=(Select U_Z_DAYS from [@Z_WORK] where U_Z_MONTH=@aMonth and U_Z_YEAR=@aYear)    
   else    
   set @NoofworkingDays=22    
       
  Declare @dblEarning Decimal(19,3)  
  Declare @strDate varchar(10)  
  Declare @dblDailyRate Decimal(19,3)  
  Declare @dblBasic Decimal(19,3)  
  set @dblBasic=(Select ISNULL(U_Z_DailyRate,0) from [@Z_PAY_LEAVE] where Code=@strLeaveCode)  
  set @strDate = (SELECT convert(varchar(10),@dtPayrollDate, 120)) --convert(varchar,YEAR(@dtPayrollDate)) +''-'' + CONVERT(varchar,Month(@dtPayrollDate)) + ''-'' + CONVERT(Varchar,Day(@dtPayrollDate))  
 -- print @dblbasic   
   
  set @dblEarning =(Select sum(isnull(U_Z_EARN_VALUE,0)) from [@Z_PAY1] where U_Z_EMPID=@strempID   
  and  U_Z_EARN_TYPE in (Select U_Z_CODE from [@Z_PAY_OLEMAP] where isnull(U_Z_EFFPAY,''N'')=''Y'' and U_Z_LEVCODE=@strLeaveCode ) and @strDate between ISNULL(U_Z_StartDate,@strDate) and ISNULL(U_Z_EndDate,@strdate))  
    
  set @dblDailyRate =@dblTotalBasic+isnull(@dblEarning ,0)  
  set @dblDailyRate=@dblDailyRate/@dblBasic   
  
    
    
Update [@Z_PAYROLL5]    
set U_Z_DailyRate=ISNULL(@dblDailyRate,0), U_Z_OB=isnull(@dblU_Z_OB,0) ,U_Z_CM=isnull(@dblU_Z_CM,0) ,U_Z_NoofDays =@dblNoofDays1 , U_Z_CashoutDays =isnull(@dblCashout,0) ,U_Z_Redim=isnull(@dblRedimdays,0)  ,U_Z_Adjustment=isnull(@dblAdjustmentDays,0),U_Z_EnCashment=isnull(@dblnoofEncashment1,0) ,  
U_Z_Balance =isnull(@dblBal ,0)    
where Code=@PayEmpRefCode    
       
       
       
Update [@Z_PAYROLL5] set U_Z_TotalAvDays=isnull(U_Z_CM,0)+isnull(U_Z_NoofDays,0), U_Z_Balance = isnull(U_Z_CM,0) + isnull(U_Z_NoofDays,0)-isnull(U_Z_Redim,0) + isnull(U_Z_Adjustment,0)-isnull(U_Z_EnCashment,0) ,    
U_Z_Amount=((isnull(U_Z_DedRate,0) * isnull(U_Z_DailyRate,0))/100) * isnull(U_Z_Redim,0),U_Z_CurAmount=(isnull(U_Z_DailyRate,0) * isnull(U_Z_NoofDays,0)) where Code=@PayEmpRefCode    
Update [@Z_PAYROLL5] set  U_Z_CashOutAmt=((isnull(U_Z_DedRate,0) * isnull(U_Z_DailyRate,0))/100) * isnull(U_Z_CashOutDays,0) where Code=@PayEmpRefCode    
Update [@Z_PAYROLL5] set  U_Z_Amount=Round(isnull(U_Z_Amount,0),3),U_Z_CurAmount=Round(isnull(U_Z_CurAmount,0),3) where Code=@PayEmpRefCode    
Update [@Z_PAYROLL5] set  U_Z_AcrAmount = (isnull(U_Z_CurAmount,0) + isnull(U_Z_CMAmt,0)+isnull(U_Z_Increment,0))   
 where U_Z_PaidLeave=''A'' and Code=@PayEmpRefCode    
Update [@Z_PAYROLL5] set  U_Z_BalanceAmt = isnull(U_Z_AcrAmount,0)-isnull(U_Z_Amount,0) where U_Z_PaidLeave=''A'' and Code=@PayEmpRefCode    
    
fetch Next from UpdateEmpLeaveDetails   Into @PayEmpRefCode,@strLeaveCode,@strEmpID,@payroll1RefCode    
end    
Close UpdateEmpLeaveDetails    
Deallocate UpdateEmpLeaveDetails    
  
Exec UPDATELEAVEBALANCE_TRANSACTION_EMPLOYEE @strempid,@aYear,@aMonth   
        
END' 
END
 
/****** Object:  StoredProcedure [dbo].[RESETPARYROLLWORKSHEET_REGULAR]    Script Date: 12/02/2014 09:36:40 ******/
SET ANSI_NULLS ON
 
SET QUOTED_IDENTIFIER ON
 
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[RESETPARYROLLWORKSHEET_REGULAR]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N' -- exec RESETPARYROLLWORKSHEET_REGULAR 2014,5,''KCC''  
      
      
CREATE PROCEDURE [dbo].[RESETPARYROLLWORKSHEET_REGULAR]      
      
@YEAR INT,      
@MONTH INT,      
@ACOMPANY VARCHAR(MAX)      
AS      
      
BEGIN      
      
DECLARE @strPayRefCod Varchar(200)      
Declare @strEmpRefCode Varchar(200)      
Declare @strQuery varchar(Max)      
      
if exists (Select Code from [@Z_PAYROLL] where  isnull(U_Z_OffCycle,''N'')=''N'' and U_Z_CompNo=@ACOMPANY  and U_Z_Year=@YEAR  and U_Z_Month=@MONTH  and U_Z_Process=''N'')      
      
      
BEGIN    
  --print ''Test1''     
   set @strPayRefCod= (Select Code from [@Z_PAYROLL] where  isnull(U_Z_OffCycle,''N'')=''N'' and U_Z_CompNo=@ACOMPANY  and U_Z_Year=@YEAR    
  
   and U_Z_Month=@MONTH  and U_Z_Process=''N'' )  
  --  print @strpayrefcod   
       Delete from [@Z_PAYROLL22] where U_Z_RefCode in (Select Code from [@Z_PAYROLL1] where  isnull(U_Z_Posted,''N'')=''N''  and U_Z_RefCode=@strPayRefCod)                  Delete from [@Z_PAYROLL2] where U_Z_RefCode in (Select Code from [@Z_PAYROLL1] where 
 isnull(U_Z_Posted,''N'')=''N''  and U_Z_RefCode=@strPayRefCod)    
Delete  from [@Z_PAYROLL3] where U_Z_RefCode in (Select  Code from [@Z_PAYROLL1] where U_Z_RefCode =''00000101'')  
   Delete from [@Z_PAYROLL3] where U_Z_RefCode in (Select Code from [@Z_PAYROLL1] where  isnull(U_Z_Posted,''N'')=''N''  and U_Z_RefCode=@strPayRefCod)      
   Delete from [@Z_PAYROLL4] where U_Z_RefCode in (Select Code from [@Z_PAYROLL1] where  isnull(U_Z_Posted,''N'')=''N''  and U_Z_RefCode=@strPayRefCod)      
   Delete from [@Z_PAYROLL5] where U_Z_RefCode in (Select Code from [@Z_PAYROLL1] where  isnull(U_Z_Posted,''N'')=''N''  and U_Z_RefCode=@strPayRefCod)      
   Delete from [@Z_PAYROLL6] where U_Z_RefCode in (Select Code from [@Z_PAYROLL1] where  isnull(U_Z_Posted,''N'')=''N''  and U_Z_RefCode=@strPayRefCod)      
   Delete from [@Z_PAY_BANK] where U_Z_RefCode in (Select Code from [@Z_PAYROLL1] where  isnull(U_Z_Posted,''N'')=''N''  and U_Z_RefCode=@strPayRefCod)      
   Delete from [@Z_PAYROLL12] where U_Z_RefCode in (Select Code from [@Z_PAYROLL1] where  isnull(U_Z_Posted,''N'')=''N''  and U_Z_RefCode=@strPayRefCod)                  Delete from [@Z_PAY_EMP_OSAV] where U_Z_RefCode in (Select Code from [@Z_PAYROLL1] where 
 isnull(U_Z_Posted,''N'')=''N''  and U_Z_RefCode=@strPayRefCod)               Exec UpdateEmployeeLeavedetails_Employee1 @strPayRefCod,@YEAR,@MONTH    
   -- print ''test''  
   Delete from [@Z_PAYROLL1] where U_Z_RefCode=@strPayRefcod       
   Delete from [@Z_PAYROLL] where isnull(U_Z_OffCycle,''N'')=''N'' and U_Z_CompNo=@ACOMPANY   and  U_Z_Year=@year and U_Z_Month=@MONTH  and U_Z_Process=''N''                 
END      
      
END' 
END
 
/****** Object:  StoredProcedure [dbo].[INSERTPAYROLL_LEAVEDETAILS_FOR_MONTH]    Script Date: 12/02/2014 09:36:40 ******/
SET ANSI_NULLS ON
 
SET QUOTED_IDENTIFIER ON
 
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[INSERTPAYROLL_LEAVEDETAILS_FOR_MONTH]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'    
--exec INSERTPAYROLL_LEAVEDETAILS_FOR_MONTH ''00000005'',2014,6    
    
CREATE  Procedure [dbo].[INSERTPAYROLL_LEAVEDETAILS_FOR_MONTH]    
    
@aRefCode varchar(30),    
@aYear int,    
@aMonth int    
    
AS    
    
BEGIN    
Declare @PayEmpRefCode varchar(30)    
    
Declare AddLeaveDetails Cursor For Select "Code" from "@Z_PAYROLL1" where "U_Z_RefCode"=@aRefcode    
   Open AddLeaveDetails    
     fetch Next from AddLeaveDetails   Into @PayEmpRefCode    
        While @@Fetch_Status=0    
            
        BEGIN    
       Exec ADD_LEAVEDETAILS @PayEmpRefCode,@aYear,@aMonth    
      Exec UPDATE_EMPLOYEE_LEAVEDETAILS @PayEmpRefCode,@ayear,@aMonth   
     --   exec PROCONS_SP_PAY_AddAirFare_Emp @ayear,@amonth,@PayEmpRefCode
        fetch Next from AddLeaveDetails   Into @PayEmpRefCode    
     end    
     Close AddLeaveDetails    
       Deallocate AddLeaveDetails    
        
END' 
END
 
