

If Exists
(
	select '1' from sysobjects 
	where name = 'stpItemKeybrdList'
	and type = 'P'
)
	Begin
		Drop Procedure stpItemKeybrdList
	End
Go
/*
	stpItemKeybrdList 1001,1001
*/
create procedure stpItemKeybrdList
(
	 @iCategoryId	numeric(4)	
	,@iKbdCode	numeric(4)	
)
As
Begin
	select 	 seq					as seq
		,itm_code				as itm_code
		,isnull(Items.shortname,'<unknown>')	as shortname
		,isnull(Items.rtl_prc,0)		as rtl_prc
		,isnull(Items.disc_amt,0)		as disc_amt
		,isnull(Sizes.shortname,'<unknown>')	as SizeName
	from KeybrdItem
	Left join Items on (Items.code 	= KeybrdItem.itm_code)
	Left join Sizes on (Sizes.code 	= Items.size_id)
	where 	KeybrdItem.actv_fg 	= 1
	and	Items.category_id 	= @iCategoryId
	and 	KeybrdItem.Keybrd_code 	= @iKbdCode
	order by seq 

	Return
End


