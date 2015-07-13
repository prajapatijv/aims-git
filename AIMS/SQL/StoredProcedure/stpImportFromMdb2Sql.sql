
If Exists
(
	select '1' from sysobjects 
	where name = 'stpImportFromMdb2Sql'
	and type = 'P'
)
	Begin	Drop Procedure stpImportFromMdb2Sql	End

GO
/*
	Exec stpImportFromMdb2Sql 'E:\AUM\AumSoft\Export\ServerExport.mdb','Items',1
*/

CREATE PROCEDURE stpImportFromMdb2Sql
(
	 @MdbPath	VARCHAR(200)
	,@TableName	VARCHAR(50)
	,@IsEmptyTable  Bit		= 0
)
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
	ELSE
		BEGIN	
			PRINT 'Table '+ @TableName + ' not Found!'
			RETURN
		END

	IF @IsEmptyTable <> 0
	BEGIN
		SET @Sql = 'Truncate Table ' + @TableName
	END	

	SET @Sql = @Sql + ' Insert into ' + @TableName
	SET @Sql = @Sql + ' ( ' + @fld + ' ) '

	SET @Sql = @Sql + ' Select ' + @fld 
	SET @Sql = @Sql + ' From OpenDataSource(''Microsoft.Jet.OLEDB.4.0'''
	SET @Sql = @Sql + ' ,''Data Source="'+ @MdbPath + '";User ID=;Password=;'''
	SET @Sql = @Sql + ')...' + @TableName

	--Print @Sql
	EXECUTE (@Sql)

	Select 	 @TableName AS TableName
		,@@Rowcount AS RowCnt

	RETURN
END