
/*
select * from Saltrn
select * from Saldet
select * from UserMast
select * from Items
select * from Units
*/
If Exists
(
	select '1' from sysobjects 
	where name = 'rptSalesDetail'
	and type = 'P'
)
	Begin	Drop Procedure rptSalesDetail	End
Go

/*
	Exec rptSalesDetail '01/01/2008','31/12/2009',0,0,0,0,0,0,0,'User','0',2,1
*/

Create Procedure rptSalesDetail
(
	 @from_dt		Varchar(20) 
	,@to_dt			Varchar(20) 

	,@TerId			SmallInt
	,@UserId		SmallInt
	,@Status		TinyInt		= 0			--0:All/ 1:Active /2:Rejected	

	,@ItemId		SmallInt
	,@CategoryId	SmallInt
	,@SizeId		SmallInt
	,@UnitId		SmallInt

	,@grp_by		Varchar(30) = ''
	,@trans_id		Varchar(14) = '0'
	,@GenAt			TinyInt		= 1			--1:Server/ 2:Terminal
	,@PreviewEnabled TinyInt	= 0
)
As
Begin

---------------------------------------------------
-- Variable Declaration
---------------------------------------------------
	Declare @datfmt 	as TinyInt
	Declare @date_from 	as Datetime
	Declare @date_to 	as Datetime

	Set @datfmt = 103
	Set dateformat 'dmy'
	
	Set @date_from 	= @from_dt
	Set @date_to 	= @to_dt

---------------------------------------------------
-- Set Report Source table
---------------------------------------------------
	If exists (select '1' from sysobjects where name = 'tmpReportSource')
	Begin
		Drop Table tmpReportSource
	End		

---------------------------------------------------
-- Report Query
---------------------------------------------------
	If (@GenAt=1)		--Server
	Begin
		select 
			 Saltrn.tran_id
			,Saldet.tran_Seq

			,Saldet.Itm_Code
			,Isnull(Items.[Name],'')			As ItemName
			,Isnull(Items.[ShortName],'')		As ItemShortName

			,Items.[Category_id]				As Category_Id
			,Isnull(Categories.[Name],'')		As CategoryName
			,Isnull(Categories.[ShortName],'')	As CategoryShortName

			,Items.[Size_id]					As Size_Id
			,Isnull(Sizes.[Name],'')			As SizeName
			,Isnull(Sizes.[ShortName],'')		As SizeShortName

			,Items.[Unit_id]					As Unit_Id
			,Isnull(Units.[Name],'')			As UnitName
			,Isnull(Units.[ShortName],'')		As UnitShortName

			,Saldet.Rtl_Prc						As Rtl_Amt
			,Saldet.Disc_Amt					As Disc_Amt
			,Saldet.Qty							As Qty
			,Saldet.Amt							As SalesAmt

			,Saltrn.ter_id						As ter_id
			,Saltrn.Paid_Amt					As Customer_Amt
			,Saltrn.Change_Amt					As Change_Amt
			,Convert(Varchar(12), Saltrn.dtadat,@datfmt) 	As dtadat	--dmy
			,Convert(Varchar(12), Saltrn.dtatim,101) 		As dtatim	--hh:mm:ss
			,Saltrn.dtausr
			,UserMast.UName	+ '[' + convert(varchar(4),Saltrn.dtausr) + ']'	As UserName
			,(Case 
				IsNull(Saltrn.canceled,0)
				When 1 then 'Cancelled'
				Else ''
			 End)		As Cancelled
			,@grp_by	As Grp_by
			,(Case 	@grp_by	
				When 'Item'	Then Saldet.Itm_Code
				When 'Category' Then Items.Category_Id
				When 'User' 	Then Saltrn.dtausr
				Else 0
			 End)		As Grp_Id

			,(Case 	@grp_by	
				When 'Item'	Then Items.ShortName
				When 'Category' Then Categories.ShortName
				When 'User' 	Then Saltrn.dtausr
				Else ''
			 End)	As Grp_Name

		into tmpReportSource
		from  Saltrn
		Inner Join Saldet	On (Saldet.tran_id 	=  Saltrn.tran_id)

		Left Join Items 	On (Items.Code 		=  Saldet.Itm_Code)
		Left Join UserMast 	On (UserMast.Uid 	=  Saltrn.dtausr)

		Left Join Categories	On (Categories.Code 	=  Items.category_id)
		Left Join Sizes		On (Sizes.Code 		=  Items.size_id)
		Left Join Units		On (Units.Code 		=  Items.Unit_id)

		--Terminal Id	
		Where  Saltrn.Ter_id  =
			( 
			 Case 	when @TerId = 0 
				then  Saltrn.Ter_id
				else @TerId
			 End
			)

		--User Id	
		And  Saltrn.dtausr  =
			( 
			 Case 	when @UserId = 0 
				then Saltrn.dtausr
				else @UserId
			 End
			)

		--Item Id	
		And  Saldet.Itm_Code  =
			( 
			 Case 	when @ItemId = 0 
				then Saldet.Itm_code
				else @ItemId
			 End
			)

		--Category Id	
		And  Items.Category_Id  =
			( 
			 Case 	when @CategoryId = 0 
				then Items.Category_Id
				else @CategoryId
			 End
			)

		--Size Id	
		And  Items.Size_Id  =
			( 
			 Case 	when @SizeId = 0 
				then Items.Size_Id
				else @SizeId
			 End
			)

		--Unit Id	
		And  Items.Unit_Id  =
			( 
			 Case 	when @UnitId = 0 
				then Items.Unit_Id
				else @UnitId
			 End
			)

		--Date Range
		And  Saltrn.dtadat Between @date_from and @date_to

		--Status
		And  IsNull(Saltrn.canceled,0) In (
						Case @Status
							When 0 then IsNull(Saltrn.canceled,0)		--All
							When 1 then 0					--Active
							When 2 then 1					--Calceled
						end
						)

		--Transction Id
		And  Saltrn.tran_id =
			( 
			 Case 	
				when @trans_id = '0'  then Saltrn.tran_id
				else @trans_id
			 End
			)
	End
	Else
	Begin
		select 
			 TerSaltrn.tran_id
			,TerSaldet.tran_Seq

			,TerSaldet.Itm_Code
			,Isnull(Items.[Name],'')			As ItemName
			,Isnull(Items.[ShortName],'')		As ItemShortName

			,Items.[Category_id]				As Category_Id
			,Isnull(Categories.[Name],'')		As CategoryName
			,Isnull(Categories.[ShortName],'')	As CategoryShortName

			,Items.[Size_id]					As Size_Id
			,Isnull(Sizes.[Name],'')			As SizeName
			,Isnull(Sizes.[ShortName],'')		As SizeShortName

			,Items.[Unit_id]					As Unit_Id
			,Isnull(Units.[Name],'')			As UnitName
			,Isnull(Units.[ShortName],'')		As UnitShortName

			,TerSaldet.Rtl_Prc						As Rtl_Amt
			,TerSaldet.Disc_Amt					As Disc_Amt
			,TerSaldet.Qty							As Qty
			,TerSaldet.Amt							As SalesAmt

			,TerSaltrn.ter_id						As ter_id
			,TerSaltrn.Paid_Amt					As Customer_Amt
			,TerSaltrn.Change_Amt					As Change_Amt
			,Convert(Varchar(12), TerSaltrn.dtadat,@datfmt) 	As dtadat	--dmy
			,Convert(Varchar(12), TerSaltrn.dtatim,101) 		As dtatim	--hh:mm:ss
			,TerSaltrn.dtausr
			,UserMast.UName	+ '[' + convert(varchar(4),TerSaltrn.dtausr) + ']'	As UserName
			,(Case 
				IsNull(TerSaltrn.canceled,0)
				When 1 then 'Cancelled'
				Else ''
			 End)		As Cancelled
			,@grp_by	As Grp_by
			,(Case 	@grp_by	
				When 'Item'	Then TerSaldet.Itm_Code
				When 'Category' Then Items.Category_Id
				When 'User' 	Then TerSaltrn.dtausr
				Else 0
			 End)		As Grp_Id

			,(Case 	@grp_by	
				When 'Item'	Then Items.ShortName
				When 'Category' Then Categories.ShortName
				When 'User' 	Then TerSaltrn.dtausr
				Else ''
			 End)	As Grp_Name

		into tmpReportSource
		from  TerSaltrn
		Inner Join TerSaldet	On (TerSaldet.tran_id 	=  TerSaltrn.tran_id)

		Left Join Items 	On (Items.Code 		=  TerSaldet.Itm_Code)
		Left Join UserMast 	On (UserMast.Uid 	=  TerSaltrn.dtausr)

		Left Join Categories	On (Categories.Code 	=  Items.category_id)
		Left Join Sizes		On (Sizes.Code 		=  Items.size_id)
		Left Join Units		On (Units.Code 		=  Items.Unit_id)

		--Terminal Id	
		Where  TerSaltrn.Ter_id  =
			( 
			 Case 	when @TerId = 0 
				then  TerSaltrn.Ter_id
				else @TerId
			 End
			)

		--User Id	
		And  TerSaltrn.dtausr  =
			( 
			 Case 	when @UserId = 0 
				then TerSaltrn.dtausr
				else @UserId
			 End
			)

		--Item Id	
		And  TerSaldet.Itm_Code  =
			( 
			 Case 	when @ItemId = 0 
				then TerSaldet.Itm_code
				else @ItemId
			 End
			)

		--Category Id	
		And  Items.Category_Id  =
			( 
			 Case 	when @CategoryId = 0 
				then Items.Category_Id
				else @CategoryId
			 End
			)

		--Size Id	
		And  Items.Size_Id  =
			( 
			 Case 	when @SizeId = 0 
				then Items.Size_Id
				else @SizeId
			 End
			)

		--Unit Id	
		And  Items.Unit_Id  =
			( 
			 Case 	when @UnitId = 0 
				then Items.Unit_Id
				else @UnitId
			 End
			)

		--Date Range
		And  TerSaltrn.dtadat Between @date_from and @date_to

		--Status
		And  IsNull(TerSaltrn.canceled,0) In (
						Case @Status
							When 0 then IsNull(TerSaltrn.canceled,0)		--All
							When 1 then 0					--Active
							When 2 then 1					--Calceled
						end
						)

		--Transction Id
		And  TerSaltrn.tran_id =
			( 
			 Case 	
				when @trans_id = '0'  then TerSaltrn.tran_id
				else @trans_id
			 End
			)
	End
End

	--Used to view data in Sql server only for debug purpose
	If (@PreviewEnabled = 1)
	Begin
		Select * from tmpReportSource
	End

Return