/****** Object:  StoredProcedure [dbo].[Insert_ContDetails]    Script Date: 03/22/2015 11:38:48 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Insert_ContDetails]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[Insert_ContDetails]
 
/****** Object:  StoredProcedure [dbo].[Insert_DeductionDetails]    Script Date: 03/22/2015 11:38:48 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Insert_DeductionDetails]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[Insert_DeductionDetails]
 
/****** Object:  StoredProcedure [dbo].[Insert_EarAccrual]    Script Date: 03/22/2015 11:38:48 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Insert_EarAccrual]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[Insert_EarAccrual]
 
/****** Object:  StoredProcedure [dbo].[Insert_EarningDetails]    Script Date: 03/22/2015 11:38:48 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Insert_EarningDetails]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[Insert_EarningDetails]
 
/****** Object:  StoredProcedure [dbo].[Insert_LeaveDetails]    Script Date: 03/22/2015 11:38:48 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Insert_LeaveDetails]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[Insert_LeaveDetails]
 
/****** Object:  StoredProcedure [dbo].[Insert_ProjectDetails]    Script Date: 03/22/2015 11:38:48 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Insert_ProjectDetails]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[Insert_ProjectDetails]
 
/****** Object:  StoredProcedure [dbo].[Insert_ProjectDetails]    Script Date: 03/22/2015 11:38:48 ******/
SET ANSI_NULLS ON
 
SET QUOTED_IDENTIFIER ON
 
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Insert_ProjectDetails]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[Insert_ProjectDetails]
    @sXML NTEXT    
AS    
BEGIN
 
    SET NOCOUNT ON;
    DECLARE @docHandle int
	
	
DECLARE @RowCount INT = 0
SELECT @RowCount = isnull(Max(CONVERT(numeric,Code)),0) FROM [@Z_PAYROLL12]


   exec sp_xml_preparedocument @docHandle OUTPUT, @sXML
	 
    INSERT INTO [@Z_PAYROLL12](Code,Name,U_Z_RefCode,U_Z_Type,U_Z_Field,U_Z_FieldName,U_Z_Rate,U_Z_Value,U_Z_PostType,U_Z_GLACC,
    U_Z_CardCode,U_Z_PrjCode)
   	SELECT  Code = RIGHT(''00000000'' +CAST(ROW_NUMBER() OVER(ORDER BY RefCode ) +@RowCount AS NVARCHAR),8),
    Name = RIGHT(''00000000'' +CAST(ROW_NUMBER() OVER(ORDER BY RefCode) +@RowCount AS NVARCHAR),8)
	,RefCode,Type,Field,FieldName,Rate,Value,PostType,GLACC,CardCode,PrjCode
    FROM OPENXML (@docHandle, ''Worksheet/Project'', 2)
    WITH (RefCode Varchar(20),Type Varchar(20),Field Varchar(20),FieldName Varchar(20),Rate Varchar(20),Value Varchar(20),PostType Varchar(20),GLACC Varchar(20),CardCode Varchar(20),PrjCode Varchar(20)
    ) 
	   
	exec sp_xml_removedocument @docHandle     
	Update [@Z_PAYROLL12] set  U_Z_CardCode='' '' where U_Z_CardCode=''-''
	SET NOCOUNT OFF;
END' 
END
 
/****** Object:  StoredProcedure [dbo].[Insert_LeaveDetails]    Script Date: 03/22/2015 11:38:48 ******/
SET ANSI_NULLS ON
 
SET QUOTED_IDENTIFIER ON
 
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Insert_LeaveDetails]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[Insert_LeaveDetails]
    @sXML NTEXT    
AS    
BEGIN
 
    SET NOCOUNT ON;
    DECLARE @docHandle int
	
	
DECLARE @RowCount INT = 0
SELECT @RowCount = isnull(Max(CONVERT(numeric,Code)),0) FROM [@Z_PAYROLL5]


   exec sp_xml_preparedocument @docHandle OUTPUT, @sXML
	 
    INSERT INTO [@Z_PAYROLL5](Code,Name,U_Z_RefCode,U_Z_EmpID,U_Z_LeaveCode,U_Z_LeaveName,U_Z_PaidLeave,U_Z_OB,U_Z_OBAmt,U_Z_CM,U_Z_CMAmt,U_Z_NoofDays
    ,U_Z_DailyRate,U_Z_Redim,U_Z_Amount,U_Z_GLACC,U_Z_PostType,U_Z_GLACC1,U_Z_Year,U_Z_Adjustment,U_Z_Month,U_Z_DedRate,U_Z_CashoutDays,U_Z_Basic,U_Z_Encashment)
   	SELECT  Code = RIGHT(''00000000'' +CAST(ROW_NUMBER() OVER(ORDER BY RefCode ) +@RowCount AS NVARCHAR),8),
    Name = RIGHT(''00000000'' +CAST(ROW_NUMBER() OVER(ORDER BY RefCode) +@RowCount AS NVARCHAR),8)
	,RefCode,EmpID,LeaveCode,LeaveName,PaidLeave,OB,OBAmt,CM,CMAmt,NoofDays,DailyRate,Redim,Amount,GLACC,PostType,GLACC1,Year,Adjustment,Month,DedRate,CashoutDays,Basic,Encashment
    FROM OPENXML (@docHandle, ''Worksheet/LeaveDetails'', 2)
    WITH (RefCode Varchar(20),EmpID VarChar(20), LeaveCode VarChar(100),LeaveName VarChar(100),PaidLeave VarChar(100),OB VarChar(100)
    ,OBAmt VarChar(100),CM varChar(100),CMAmt VarChar(100),NoofDays VarChar(100),DailyRate Varchar(100),Redim Decimal(18,2),Amount Varchar(100),GLACC Varchar(100)
    ,PostType Varchar(100),GLACC1 Varchar(100),Year Varchar(100),Adjustment Varchar(100),Month Varchar(100),DedRate Varchar(100),CashoutDays Varchar(100),Basic Varchar(10),Encashment Varchar(100)
    ) 
	   
	exec sp_xml_removedocument @docHandle     
	
	
	SET NOCOUNT OFF;
END' 
END
 
/****** Object:  StoredProcedure [dbo].[Insert_EarningDetails]    Script Date: 03/22/2015 11:38:48 ******/
SET ANSI_NULLS ON
 
SET QUOTED_IDENTIFIER ON
 
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Insert_EarningDetails]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[Insert_EarningDetails]
    @sXML NTEXT    
AS    
BEGIN
 
    SET NOCOUNT ON;
    DECLARE @docHandle int
	
	
DECLARE @RowCount INT = 0
SELECT @RowCount = isnull(Max(CONVERT(numeric,Code)),0) FROM [@Z_PAYROLL2]


   exec sp_xml_preparedocument @docHandle OUTPUT, @sXML
	 
    INSERT INTO [@Z_PAYROLL2](Code,Name,U_Z_RefCode,U_Z_Type,U_Z_Field,U_Z_FieldName,U_Z_Rate,U_Z_Value,U_Z_PostType,U_Z_GLACC,
    U_Z_CardCode,U_Z_EarValue)
   	SELECT  Code = RIGHT(''00000000'' +CAST(ROW_NUMBER() OVER(ORDER BY RefCode ) +@RowCount AS NVARCHAR),8),
    Name = RIGHT(''00000000'' +CAST(ROW_NUMBER() OVER(ORDER BY RefCode) +@RowCount AS NVARCHAR),8)
	,RefCode,Type,Field,FieldName,Rate,Value,PostType,GLACC,CardCode,EarValue
    FROM OPENXML (@docHandle, ''Worksheet/Earning'', 2)
    WITH (RefCode Varchar(20),Type Varchar(20),Field Varchar(20),FieldName Varchar(20),Rate Varchar(20),Value Varchar(20),PostType Varchar(20),GLACC Varchar(20),CardCode Varchar(20),EarValue Varchar(20)
    ) 
	   
	exec sp_xml_removedocument @docHandle     
	Update [@Z_PAYROLL2] set  U_Z_CardCode='' '' where U_Z_CardCode=''-''
	SET NOCOUNT OFF;
END' 
END
 
/****** Object:  StoredProcedure [dbo].[Insert_EarAccrual]    Script Date: 03/22/2015 11:38:48 ******/
SET ANSI_NULLS ON
 
SET QUOTED_IDENTIFIER ON
 
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Insert_EarAccrual]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[Insert_EarAccrual]
    @sXML NTEXT    
AS    
BEGIN
 
    SET NOCOUNT ON;
    DECLARE @docHandle int
	
	
DECLARE @RowCount INT = 0
SELECT @RowCount = isnull(Max(CONVERT(numeric,Code)),0) FROM [@Z_PAYROLL22]


   exec sp_xml_preparedocument @docHandle OUTPUT, @sXML
	 
    INSERT INTO [@Z_PAYROLL22](Code,Name,U_Z_RefCode,U_Z_Type,U_Z_Field,U_Z_FieldName,U_Z_Rate,U_Z_Value,U_Z_AccDebit,U_Z_AccCredit,U_Z_EmpID,U_Z_Month,
    U_Z_Year,U_Z_PrjCode,U_Z_OB,U_Z_ClosingBalance,U_Z_CardCode)
   	SELECT  Code = RIGHT(''00000000'' +CAST(ROW_NUMBER() OVER(ORDER BY RefCode ) +@RowCount AS NVARCHAR),8),
    Name = RIGHT(''00000000'' +CAST(ROW_NUMBER() OVER(ORDER BY RefCode) +@RowCount AS NVARCHAR),8)
	,RefCode,Type,Field,FieldName,Rate,Value,AccDebit,AccCredit,EmpID,Month,Year,PrjCode,OB,ClosingBalance,CardCode
    FROM OPENXML (@docHandle, ''Worksheet/Deductions'', 2)
    WITH (RefCode Varchar(20),Type Varchar(20),Field Varchar(20),FieldName Varchar(20),Rate Varchar(20),Value Varchar(20),AccDebit Varchar(20),AccCredit Varchar(20),EmpID Varchar(20),Month Varchar(20),Year Varchar(20),PrjCode Varchar(20),OB Varchar(20),ClosingBalance Varchar(20),CardCode Varchar(20)
    ) 
	   
	exec sp_xml_removedocument @docHandle     
	Update [@Z_PAYROLL22] set  U_Z_CardCode='' '' where U_Z_CardCode=''-''
	SET NOCOUNT OFF;
END' 
END
 
/****** Object:  StoredProcedure [dbo].[Insert_DeductionDetails]    Script Date: 03/22/2015 11:38:48 ******/
SET ANSI_NULLS ON
 
SET QUOTED_IDENTIFIER ON
 
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Insert_DeductionDetails]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[Insert_DeductionDetails]
    @sXML NTEXT    
AS    
BEGIN
 
    SET NOCOUNT ON;
    DECLARE @docHandle int
	
	
DECLARE @RowCount INT = 0
SELECT @RowCount = isnull(Max(CONVERT(numeric,Code)),0) FROM [@Z_PAYROLL3]


   exec sp_xml_preparedocument @docHandle OUTPUT, @sXML
	 
    INSERT INTO [@Z_PAYROLL3](Code,Name,U_Z_RefCode,U_Z_Type,U_Z_Field,U_Z_FieldName,U_Z_Rate,U_Z_Value,U_Z_PostType,U_Z_GLACC,
    U_Z_CardCode,U_Z_EarValue)
   	SELECT  Code = RIGHT(''00000000'' +CAST(ROW_NUMBER() OVER(ORDER BY RefCode ) +@RowCount AS NVARCHAR),8),
    Name = RIGHT(''00000000'' +CAST(ROW_NUMBER() OVER(ORDER BY RefCode) +@RowCount AS NVARCHAR),8)
	,RefCode,Type,Field,FieldName,Rate,Value,PostType,GLACC,CardCode,EarValue
    FROM OPENXML (@docHandle, ''Worksheet/Deductions'', 2)
    WITH (RefCode Varchar(20),Type Varchar(20),Field Varchar(20),FieldName Varchar(20),Rate Varchar(20),Value Varchar(20),PostType Varchar(20),GLACC Varchar(20),CardCode Varchar(20),EarValue Varchar(20)
    ) 
	   
	exec sp_xml_removedocument @docHandle     
	Update [@Z_PAYROLL3] set U_Z_CardCode='' '' where U_Z_CardCode=''-''
	SET NOCOUNT OFF;
END' 
END
 
/****** Object:  StoredProcedure [dbo].[Insert_ContDetails]    Script Date: 03/22/2015 11:38:48 ******/
SET ANSI_NULLS ON
 
SET QUOTED_IDENTIFIER ON
 
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Insert_ContDetails]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[Insert_ContDetails]  
    @sXML NTEXT      
AS      
BEGIN  
   
    SET NOCOUNT ON;  
    DECLARE @docHandle int  
   
   
DECLARE @RowCount INT = 0  
SELECT @RowCount = isnull(Max(CONVERT(numeric,Code)),0) FROM [@Z_PAYROLL4]  
  
  
   exec sp_xml_preparedocument @docHandle OUTPUT, @sXML  
    
    INSERT INTO [@Z_PAYROLL4](Code,Name,U_Z_RefCode,U_Z_Type,U_Z_Field,U_Z_FieldName,U_Z_Rate,U_Z_Value,U_Z_PostType,U_Z_GLACC  
    ,U_Z_GLACC1
   )  
    	SELECT  Code = RIGHT(''00000000'' +CAST(ROW_NUMBER() OVER(ORDER BY RefCode ) +@RowCount AS NVARCHAR),8),
    Name = RIGHT(''00000000'' +CAST(ROW_NUMBER() OVER(ORDER BY RefCode) +@RowCount AS NVARCHAR),8)
 ,RefCode,Type,Field,FieldName,Rate,Value,PostType,GLACC,GLACC1  
    FROM OPENXML (@docHandle, ''Worksheet/Contribution'', 2)  
    WITH (RefCode Varchar(20),Type Varchar(20),Field Varchar(20),FieldName Varchar(20),Rate Varchar(20),Value Varchar(20),  
    PostType Varchar(20),GLACC Varchar(20),GLACC1 Varchar(20)  
    )   
      
 exec sp_xml_removedocument @docHandle       
  
 SET NOCOUNT OFF;  
 END' 
END
 
