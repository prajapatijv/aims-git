/*select * from TerSalTrn
select * from TerSaldet

select * from Saltrn
select * from Saldet
*/
If Exists
(
	select '1' from sysobjects 
	where name = 'stpImportFromMdb2Sql_Saltrn'
	and type = 'P'
)
	Begin	Drop Procedure stpImportFromMdb2Sql_Saltrn	End

GO
/*
	Exec stpImportFromMdb2Sql_Saltrn 'E:\AUM\AumSoft\Export\TerminalImpex1002.mdb'
*/

CREATE PROCEDURE stpImportFromMdb2Sql_Saltrn
(
	 @MdbPath	VARCHAR(200)
)
AS

Begin
	SET NOCOUNT ON
	Declare @Sql as Varchar(8000)

---------------------------------------------------------------------
--	Create Temp Tables
---------------------------------------------------------------------
	If Exists
	(
		select '1' from tempdb.dbo.sysobjects where id = Object_id('tempdb.dbo.##tmp_Saltrn') 
	)
	Begin
		Drop table ##tmp_Saltrn
	End

	If Exists
	(
		select '1' from tempdb.dbo.sysobjects where id = Object_id('tempdb.dbo.##tmp_Saldet') 
	)
	Begin
		Drop table ##tmp_Saldet
	End

	CREATE TABLE ##tmp_Saltrn 
	(
		tran_id 	varchar(20) 	COLLATE Latin1_General_CI_AI NOT NULL ,
		ter_id 		smallint 	NOT NULL ,
		export_fg 	bit 		NULL  DEFAULT (0),
		paid_amt 	numeric(10, 2) 	NULL ,
		change_amt 	numeric(7, 2) 	NULL ,
		dtadat 		datetime 	NULL ,
		dtatim 		varchar (10) 	NULL ,
		dtausr 		varchar (10) 	DEFAULT (''),
		canceled	bit 		NULL  DEFAULT (0),
		Event_id	SmallInt	NULL  DEFAULT (0)
	) 

	CREATE TABLE ##tmp_Saldet
	(
		tran_id 	varchar(20) 	COLLATE Latin1_General_CI_AI NOT NULL ,
		tran_seq 	smallint 	NOT NULL ,
		itm_code 	numeric(4, 0) 	NULL ,
		rtl_prc 	numeric(10, 2) 	NULL ,
		disc_amt 	numeric(10, 2) 	NULL ,
		qty 		smallint 	NULL ,
		amt 		numeric(10, 2) 	NULL ,
	)


---------------------------------------------------------------------
--	Import Data from Mdb to SQL Temp tables
---------------------------------------------------------------------
	SET @Sql = ''

	SET @Sql = @Sql + ' Insert into ##tmp_Saltrn '
	SET @Sql = @Sql + ' ( '
	SET @Sql = @Sql + ' 		 tran_id '
	SET @Sql = @Sql + ' 		,ter_id '
	SET @Sql = @Sql + ' 		,export_fg '
	SET @Sql = @Sql + ' 		,paid_amt '
	SET @Sql = @Sql + ' 		,change_amt '
	SET @Sql = @Sql + ' 		,dtadat '
	SET @Sql = @Sql + ' 		,dtatim '
	SET @Sql = @Sql + ' 		,dtausr '
	SET @Sql = @Sql + ' 		,canceled '
	SET @Sql = @Sql + ' 		,Event_id '
	SET @Sql = @Sql + ' )'

	SET @Sql = @Sql + ' Select '
	SET @Sql = @Sql + ' 		 tran_id '
	SET @Sql = @Sql + ' 		,ter_id '
	SET @Sql = @Sql + ' 		,export_fg '
	SET @Sql = @Sql + ' 		,paid_amt '
	SET @Sql = @Sql + ' 		,change_amt '
	SET @Sql = @Sql + ' 		,dtadat '
	SET @Sql = @Sql + ' 		,dtatim '
	SET @Sql = @Sql + ' 		,dtausr '
	SET @Sql = @Sql + ' 		,canceled '
	SET @Sql = @Sql + ' 		,Event_id '
	SET @Sql = @Sql + ' From OpenDataSource(''Microsoft.Jet.OLEDB.4.0'''
	SET @Sql = @Sql + ' ,''Data Source="'+ @MdbPath + '";User ID=;Password=;'''
	SET @Sql = @Sql + ')...TerSaltrn'  

	Exec (@Sql)


	SET @Sql = ''
	SET @Sql = @Sql + ' Insert into ##tmp_Saldet '
	SET @Sql = @Sql + ' ( '
	SET @Sql = @Sql + '  		 tran_id '
	SET @Sql = @Sql + ' 		,tran_seq '
	SET @Sql = @Sql + ' 		,itm_code '
	SET @Sql = @Sql + ' 		,rtl_prc '
	SET @Sql = @Sql + ' 		,disc_amt '
	SET @Sql = @Sql + ' 		,qty '
	SET @Sql = @Sql + ' 		,amt '
	SET @Sql = @Sql + ') '

	SET @Sql = @Sql + ' Select '
	SET @Sql = @Sql + '  		 tran_id '
	SET @Sql = @Sql + ' 		,tran_seq '
	SET @Sql = @Sql + ' 		,itm_code '
	SET @Sql = @Sql + ' 		,rtl_prc '
	SET @Sql = @Sql + ' 		,disc_amt '
	SET @Sql = @Sql + ' 		,qty '
	SET @Sql = @Sql + ' 		,amt '
	SET @Sql = @Sql + ' From OpenDataSource(''Microsoft.Jet.OLEDB.4.0'''
	SET @Sql = @Sql + ' ,''Data Source="'+ @MdbPath + '";User ID=;Password=;'''
	SET @Sql = @Sql + ')...TerSaldet'  

	Exec (@Sql)


---------------------------------------------------------------------
--	Insert data from Temp table to SQL Tables
---------------------------------------------------------------------
	Insert Into Saltrn
		(
			 tran_id
			,ter_id
			,export_fg
			,paid_amt
			,change_amt
			,dtadat
			,dtatim
			,dtausr
			,canceled
			,Event_id 
		)
	Select 		 tran_id
			,ter_id
			,export_fg
			,paid_amt
			,change_amt
			,dtadat
			,dtatim
			,dtausr
			,canceled
			,Event_id 

	from ##tmp_Saltrn
	Where tran_id Not In (Select tran_id from Saltrn)


	Insert Into Saldet
		(
			 tran_id
			,tran_seq
			,itm_code
			,rtl_prc
			,disc_amt
			,qty
			,amt
		)
	Select 		 tran_id
			,tran_seq
			,itm_code
			,rtl_prc
			,disc_amt
			,qty
			,amt
	from ##tmp_Saldet
	Where tran_id Not In (Select tran_id from Saldet)

---------------------------------------------------------------------
--	End of Stp
---------------------------------------------------------------------
	Return
END