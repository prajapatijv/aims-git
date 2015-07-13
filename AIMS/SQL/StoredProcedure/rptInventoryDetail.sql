If Exists
(
	select '1' from sysobjects 
	where name = 'rptInventoryDetail'
	and type = 'P'
)
	Begin
		Drop Procedure rptInventoryDetail
	End
Go

/*
	Exec rptInventoryDetail '01/01/2008','31/12/2009',0,0,0,0,0,0,0,'TranType',0,1
*/

Create Procedure [rptInventoryDetail]
(
	 @from_dt	Varchar(20) 
	,@to_dt		Varchar(20) 
	,@TerId		SmallInt
	,@UserId	SmallInt
	,@Tran_type			TinyInt		= 0			--0:All/ 1/2/3/4/11/12/13	

	,@ItemId		SmallInt
	,@CategoryId	SmallInt
	,@SizeId		SmallInt
	,@UnitId		SmallInt

	,@grp_by		Varchar(30) = ''
	,@vno			Int			= 0

	,@PreviewEnabled	TinyInt		= 0
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
	select 
		 Invtrn.Vno
		,Invdet.Srno

		,Invdet.qty							As Qty
		,Invdet.rtl_prc						As Rtl_Prc
		,Invdet.amt							As Amount

		,Invtrn.Doc_No						As Doc_No
		,Invtrn.ter_id						As Ter_Id

		,Invdet.Itm_Code
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

		,Convert(Varchar(12), Invtrn.rec_dat,@datfmt) As Recdat	--dmy
		,Convert(Varchar(12), Invtrn.dtadat,@datfmt) As dtadat	--dmy
		,Convert(Varchar(12), Invtrn.dtatim,101) As dtatim	--hh:mm:ss
		,Invtrn.dtausr
		,UserMast.UName	+ '[' + convert(varchar(4),Invtrn.dtausr) + ']'	As UserName
		,(Case 
			IsNull(Invtrn.tran_type,0)
			When 1 then 'WDz;tu Mxtuf'			--'Opening Balance'
			When 2 then 'lJtu Mxtuf'			--'Stock Inward'
			When 3 then 'Mxtuf mhCh (JDthtu)'	--'Stock Adjustment Up'
			When 5 then 'JuatK vh; (JDthtu)'
			When 11 then 'Mxtuf mhCh (Dxtztu)'	--'Stock Adjustment Down'
			When 12 then 'Fhtc Mxtuf'			--'Stock Waste'
			When 13 then 'Issue For Sale'
			Else ''
		 End) as TranType
		,Invtrn.remarks
		,@grp_by	As Grp_by
		,(Case 	@grp_by	
			When 'Item'		Then Invdet.Itm_Code
			When 'Category' Then Items.Category_Id
			When 'User' 	Then Invtrn.dtausr
			When 'TranType' Then Invtrn.tran_type
			Else 0
		 End)		As Grp_Id

		,(Case 	@grp_by	
				When 'Item'	Then Items.ShortName
				When 'Category' Then Categories.ShortName
				When 'User' 	Then Invtrn.dtausr
				When 'TranType' Then 
								(Case 
									IsNull(Invtrn.tran_type,0)
									When 1 then 'Opening Balance'
									When 2 then 'Stock Inward'
									When 3 then 'Stock Adjustment Up'
									When 5 then 'Sales Return'
									When 11 then 'Stock Adjustment Down'
									When 12 then 'Stock Waste'
									When 13 then 'Issue For Sale'
									Else ''
								End)
				Else ''
			 End)	As Grp_Name

	into tmpReportSource
	from  Invtrn
	Inner Join Invdet	On (Invdet.vno =  Invtrn.vno)

	Left Join Items 	On (Items.Code 		=  Invdet.Itm_Code)
	Left Join UserMast 	On (UserMast.Uid 	=  Invtrn.dtausr)
	Left Join Categories	On (Categories.Code 	=  Items.category_id)
	Left Join Sizes		On (Sizes.Code 		=  Items.size_id)
	Left Join Units		On (Units.Code 		=  Items.Unit_id)


	--Terminal Id	
	Where  Invtrn.Ter_id  =
		( 
		 Case 	when @TerId = 0 
			then  Invtrn.Ter_id
			else @TerId
		 End
		)

	--User Id	
	And  Invtrn.dtausr  =
		( 
		 Case 	when @UserId = 0 
			then Invtrn.dtausr
			else @UserId
		 End
		)

		--Item Id	
		And  Invdet.Itm_Code  =
			( 
			 Case 	when @ItemId = 0 
				then Invdet.Itm_code
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
	And  Invtrn.dtadat Between @date_from and @date_to
	

	--Transction Type
	And  IsNull(Invtrn.tran_type,0) In (
					Case @Tran_type
						When 0 then IsNull(Invtrn.tran_type,0)		--All
						else @Tran_type
					end
					)


	--Vno
	And  Invdet.vno  =
		( 
		 Case 	when @vno = 0 
			then Invdet.vno
			else @vno
		 End
		)


	--Used to view data in Sql server only for debug purpose
	If (@PreviewEnabled = 1)
	Begin
		Select * from tmpReportSource
	End

Return

End