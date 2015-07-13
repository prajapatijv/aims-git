

If Exists
(
	select '1' from sysobjects 
	where name = 'stpFetchItemCategory'
	and type = 'P'
)
	Begin
		Drop Procedure stpFetchItemCategory
	End
Go
/*
	stpFetchItemCategory 
*/
create procedure stpFetchItemCategory
(
	@keybrd_code	Smallint
)
As
Begin
	select 	 Distinct 
		 Categories.code 	as Code
		,Categories.shortname	as shortname
	from KeybrdItem
	Left join Items 	on (Items.code 		= KeybrdItem.itm_code)
	Inner join Categories	on (Items.category_id 	= Categories.code)
	where 	KeybrdItem.actv_fg 	= 1
	and 	KeybrdItem.keybrd_code	= @keybrd_code
	and	Categories.code is not null	
	Return
End


