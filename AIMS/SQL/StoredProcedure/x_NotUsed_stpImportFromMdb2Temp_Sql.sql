
If Exists
(
	select '1' from sysobjects 
	where name = 'fn_ImportFromMdb2Temp_Sql'
	and type = 'FN'
)
	Begin	Drop FUNCTION fn_ImportFromMdb2Temp_Sql	End

GO
/*
	Select dbo.fn_ImportFromMdb2Temp_Sql ('E:\AUM\AumSoft\Export\TerminalImpex1002.mdb','TerSaltrn')
*/

CREATE FUNCTION fn_ImportFromMdb2Temp_Sql
(
	 @MdbPath	VARCHAR(200)
	,@TableName	VARCHAR(50)
)
Returns 	
	Varchar(8000)
AS

BEGIN

	DECLARE @fld AS VARCHAR(8000)
	DECLARE @Sql AS VARCHAR(8000)

	SET @fld = ''
	SET @Sql = ''

	SELECT @fld = name + ',' + @fld  
	FROM syscolumns 
	WHERE id = object_id(@TableName) 
	ORDER BY colid DESC

	IF LEN(@fld) > 0 
		BEGIN
			SET @fld = SubString(@fld,1,LEN(@fld)-1)
		END


	SET @Sql = @Sql + ' Select ' + @fld 
	SET @Sql = @Sql + ' From OpenDataSource(''Microsoft.Jet.OLEDB.4.0'''
	SET @Sql = @Sql + ' ,''Data Source="'+ @MdbPath + '";User ID=;Password=;'''
	SET @Sql = @Sql + ')...' + @TableName

	--Print @Sql
	Return (@Sql)
	
END