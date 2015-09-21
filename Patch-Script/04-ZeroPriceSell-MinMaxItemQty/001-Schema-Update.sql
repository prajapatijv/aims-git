USE AIMS_DB

GO

ALTER TABLE Items
	ADD
		min_qty		smallint		default(0),
		max_qty		smallint		default(0)