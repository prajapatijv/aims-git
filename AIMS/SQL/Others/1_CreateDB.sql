CREATE DATABASE [HMst]  ON 
(	NAME = N'HMst_Data'
	,FILENAME = N'C:\Program Files\Microsoft SQL Server\MSSQL\Data\HMst_Data.MDF' 
	,SIZE = 1024
	,FILEGROWTH = 10%) 

LOG ON 
(	NAME = N'HMst_Log'
	,FILENAME = N'C:\Program Files\Microsoft SQL Server\MSSQL\Data\HMst_Log.LDF' 
	,SIZE = 512
	,FILEGROWTH = 10%
)
COLLATE SQL_Latin1_General_CP1_CI_AS
