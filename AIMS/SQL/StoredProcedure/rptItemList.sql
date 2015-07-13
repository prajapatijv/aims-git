

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
	rptItemList 'Category','Item',1
*/
create procedure rptItemList
(
	@IncludeCategory		as	varchar(20)	= 'Category',			
	@IncludeItem			as  varchar(20) = 'Item',
	@PreviewEnabled			as	tinyint						--Used to view data in Sql server only for debug purpose
)
As
Begin
	

	--Set Report Source table
	If exists (select '1' from sysobjects where name = 'tmpReportSource')
	Begin
		Drop Table tmpReportSource
	End		

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

	into tmpReportSource

	From Items					
	Left join Sizes				on (Sizes.code 			= Items.size_id)
	Left join Categories		on (Items.Category_Id	= Categories.Code)

	where 	1 = 1
	and 	Items.Category_Id 	in (Select nVal from tmpReportFilters 
									where Type = @IncludeCategory) 

	or		Items.code 			in (Select nVal from tmpReportFilters 
									where Type = @IncludeItem)
	

	--Used to view data in Sql server only for debug purpose
	If (@PreviewEnabled = 1)
	Begin
		Select * from tmpReportSource
	End

	Return
End


