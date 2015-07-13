
USE Hinv05
exec sp_addumpdevice 'disk','Hinv05_1','d:\MsSqlDta\HInv05.mdf'

BACKUP DATABASE Hinv05 TO Hinv05_1

select * from master..sysdevices


exec sp_dropdevice 'hinv05'


