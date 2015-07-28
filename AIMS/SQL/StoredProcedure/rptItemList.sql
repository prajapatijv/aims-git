

If Exists
(
	select '1' from sysobjects 
	where name = 'rptItemList'
	and type = 'P'
)
	Begin
		Drop Procedure rptItemList
	End
Go
/*
	Exec rptitemlist 'category','item',1,0
	Exec rptitemlist 'category','item',1,0,1
*/
create procedure rptItemList
(
	@IncludeCategory		as	varchar(20)	= 'Category',			
	@IncludeItem			as  varchar(20) = 'Item',
	@PreviewEnabled			as	tinyint,						--Used to view data in Sql server only for debug purpose
	@RepeatLabelCount		as	tinyint = 0,
	@isPaymentLabel			as  tinyint = 0
)
As
Begin
	

	--Set Report Source table
	If exists (select '1' from sysobjects where name = 'tmpReportSource')
	Begin
		Drop Table tmpReportSource
	End		

	if (@isPaymentLabel = 1)
	Begin
		select
			 0					as itm_code
			,'kkcej vwhw fhut.'	as itm_shortname	--Bill Puru Karo.
			, 0					as rtl_prc
			, 0					as disc_amt
			, 0					as Category_Code
			, 'Payment Barcode' as Category_Name
			, '00000'			as ItemBarcode
		into tmpReportSource
	End
	Else
	Begin
		select 	 
			 Items.code									as itm_code
			,isnull(Items.shortname,'')	 +
				(Case isnull(Sizes.Code,0)
					When 0 then ''
					Else ' (' + isnull(Sizes.shortname,'') + ')'
				 End) as itm_shortname

			,isnull(Items.rtl_prc,0)				as rtl_prc
			,isnull(Items.disc_amt,0)				as disc_amt
		
			,isnull(Categories.Code,0)				as Category_Code
			,isnull(Categories.shortname,'')		
				+ '  (' + Convert(Varchar(4),isnull(Categories.Code,0)) + ')' 	as Category_Name
			,isnull(ItemBarcodes.barcode,'')		as ItemBarcode

		into tmpReportSource

		From Items
		Inner join ItemBarcodes		on (ItemBarcodes.itm_code = Items.code)				
		Left join Sizes				on (Sizes.code 			= Items.size_id)
		Left join Categories		on (Items.Category_Id	= Categories.Code)

		where 	1 = 1
		and 	Items.Category_Id 	in (Select nVal from tmpReportFilters 
										where Type = @IncludeCategory) 

		or		Items.code 			in (Select nVal from tmpReportFilters 
									where Type = @IncludeItem)
	end

	If (@RepeatLabelCount > 0)
	Begin
		Insert into tmpReportSource
		select tmpReportSource.* 
		from tmpReportSource
		inner join tmpReportFilters on tmpReportSource.itm_code = tmpReportFilters.nVal
		--Delete top 1 from tmpReportSource
	End

	--Used to view data in Sql server only for debug purpose
	If (@PreviewEnabled = 1)
	Begin
		Select * from tmpReportSource
	End

	Return
End


