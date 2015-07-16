Use AIMS_DB
GO

--Tables
USE [AIMS_DB]
GO

/****** Object:  Table [dbo].[ItemBarcodes]    Script Date: 16/07/2015 22:06:42 ******/
DROP TABLE [dbo].[ItemBarcodes]
GO

/****** Object:  Table [dbo].[ItemBarcodes]    Script Date: 16/07/2015 22:06:42 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

SET ANSI_PADDING ON
GO

CREATE TABLE [dbo].[ItemBarcodes](
	[itm_code] [numeric](4, 0) NOT NULL,
	[barcode] [varchar](30) NOT NULL,
 CONSTRAINT [PK_ItemBarcodes] PRIMARY KEY CLUSTERED 
(
	[itm_code] ASC,
	[barcode] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO

SET ANSI_PADDING OFF
GO


--Stored Proc

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
			from KeybrdItem
			Inner join ItemBarcodes on (KeybrdItem.itm_code = ItemBarcodes.itm_code)
			where 	1 = 1
			and	ItemBarcodes.barcode = @barcode
		)
	
	If (@RowCount > 0)
	Begin
		select 	top 1 
			 Items.code				as itm_code
			,isnull(Items.shortname,'<unknown>')	as shortname
			,isnull(Items.rtl_prc,0)		as rtl_prc
			,isnull(Items.disc_amt,0)		as disc_amt
			,isnull(Sizes.shortname,'<unknown>')	as SizeName
			,isnull(ItemBarcodes.barcode,'')		as ItemBarcode
		from KeybrdItem
		Inner join ItemBarcodes on (KeybrdItem.itm_code = ItemBarcodes.itm_code)
		Inner join Items 	on (KeybrdItem.itm_code = Items.code)
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
		from KeybrdItem
		Inner join Items 	on (KeybrdItem.itm_code = Items.code)
		Left join Sizes on (Sizes.code 	= Items.size_id)
		where 	1 = 1
		and	Items.code = (case when isnumeric(@barcode) = 1 then CONVERT(int, @barcode) else 0 end)
	End

	Return
End

GO


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
	@PreviewEnabled			as	tinyint,						--Used to view data in Sql server only for debug purpose
	@RepeatLabelCount		as	tinyint = 0
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

GO
