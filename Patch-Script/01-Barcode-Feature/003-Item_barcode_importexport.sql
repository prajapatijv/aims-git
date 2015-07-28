If ((Select Count(1) from ServerExport 
	Where TableType = 'Server' 
	And TableName = 'ItemBarcodes') <= 0)
Begin
	Insert into ServerExport (TableName,TableType,sWhere,actv_fg)
	Values ('ItemBarcodes','Server',NULL,1)
End 