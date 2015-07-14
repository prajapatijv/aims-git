

If Exists
(
	select '1' from sysobjects 
	where name = 'stpBarcodeItem'
	and type = 'P'
)
	Begin
		Drop Procedure stpBarcodeItem
	End
Go
/*
	Exec stpBarcodeItem '01001'
*/
create procedure stpBarcodeItem
(
	 @barcode	varchar(50)	
)
As
Begin
	Declare @RowCount as int

	Set @RowCount = 
		(
			select 	Count(1) 
			from Items
			Inner join ItemBarcodes on (Items.code = ItemBarcodes.itm_code)
			Left join Sizes on (Sizes.code 	= Items.size_id)
			where 	1 = 1
			and	ItemBarcodes.barcode = @barcode
		)
	
	If (@RowCount > 0)
	Begin
		select 	top 1 
			 itm_code				as itm_code
			,isnull(Items.shortname,'<unknown>')	as shortname
			,isnull(Items.rtl_prc,0)		as rtl_prc
			,isnull(Items.disc_amt,0)		as disc_amt
			,isnull(Sizes.shortname,'<unknown>')	as SizeName
			,isnull(ItemBarcodes.barcode,'')		as ItemBarcode
		from Items
		Inner join ItemBarcodes on (Items.code = ItemBarcodes.itm_code)
		Left join Sizes on (Sizes.code 	= Items.size_id)
		where 	1 = 1
		and	ItemBarcodes.barcode = @barcode
	End
	Else
	Begin
		select 	top 1 
			 Items.code				as itm_code
			,isnull(Items.shortname,'<unknown>')	as shortname
			,isnull(Items.rtl_prc,0)		as rtl_prc
			,isnull(Items.disc_amt,0)		as disc_amt
			,isnull(Sizes.shortname,'<unknown>')	as SizeName
			,''		as ItemBarcode
		from Items
		Left join Sizes on (Sizes.code 	= Items.size_id)
		where 	1 = 1
		and	Items.code = (case when isnumeric(@barcode) = 1 then CONVERT(int, @barcode) else 0 end)
	End

	Return
End


