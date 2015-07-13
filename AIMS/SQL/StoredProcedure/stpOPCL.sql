If Exists
(
	select '1' from sysobjects 
	where name = 'stpOPCL'
	and type = 'P'
)
	Begin	Drop Procedure stpOPCL	End
Go

/*
	Exec stpOPCL '12/12/2009','12/12/2009',0,0,1
*/

Create Procedure stpOPCL
(
	 @from_dt		Varchar(20) 
	,@to_dt			Varchar(20) 
	,@CategoryId	SmallInt		= 0
	,@ItemId		SmallInt		= 0
	,@PreviewEnabled TinyInt		= 0
)
As

Begin

	Set dateformat 'dmy'

----------------------------------------------------------
--	Declare
----------------------------------------------------------
	Declare @datfmt 	as TinyInt
	Declare @date_from 	as Datetime
	Declare @date_to 	as Datetime

	Declare @OP_STK		as Tinyint
	Declare @INWARD		as Tinyint
	Declare @ADJ_UP		as Tinyint
	Declare @SALES_RETURN_UP		as Tinyint
	Declare @ADJ_DN		as Tinyint
	Declare @STK_WASTE	as Tinyint
	
	Set @datfmt		= 103
	Set @date_from 	= @from_dt
	Set @date_to 	= @to_dt

	Set @OP_STK		= 1
	Set @INWARD		= 2
	Set @ADJ_UP		= 3
	Set @SALES_RETURN_UP = 5
	Set @ADJ_DN		= 11
	Set @STK_WASTE	= 12

---------------------------------------------------
-- Set Report Source table
---------------------------------------------------
	If exists (select '1' from sysobjects where name = 'tmpReportSource')
	Begin
		Drop Table tmpReportSource
	End		


	Declare @opcl Table
	(
		 Seq			int			identity(1,1)
		,Itm_Code		int			Not Null
		,Op_Qty			int			default(0)
		,Inward_Qty		int			default(0)
		,Sales_Qty		int			default(0)
		,Typ			tinyint		
		,Adj_Qty		int			default(0)	---Adjustment Qty between Period (Adj. UP, DOWN and Stock Waste)
	)

----------------------------------------------------------
--	Op Stock Calculation
----------------------------------------------------------
	Insert into @opcl (Itm_code,Op_Qty,Typ)
	Select	 Itm_code
			,Sum(qty)			
			,Invtrn.tran_type	
	From Invdet
	Inner Join Invtrn on (Invtrn.vno = Invdet.vno)
	Where Invtrn.tran_type in (@OP_STK,@INWARD,@ADJ_UP,@SALES_RETURN_UP,@ADJ_DN,@STK_WASTE)
	and rec_dat < @date_from
	Group by itm_code,Invtrn.tran_type

	--Sales to be deducted from Op stock
	Insert into @opcl (Itm_code,Op_Qty,Typ)
	Select	 Itm_code
			,sum(Qty) * -1	
			,99				
	from Saldet
	Inner Join Saltrn on (Saltrn.tran_id = Saldet.tran_id and Saltrn.canceled <> 1)
	Where 1=1
	And  dtadat < @date_from 
	Group by Itm_code


----------------------------------------------------------
--	Inward Stock between dates
----------------------------------------------------------
	Insert into @opcl (Itm_code,Inward_Qty,Typ)
	Select	 Itm_code
			,Sum(qty)			
			,Invtrn.tran_type	
	From Invdet
	Inner Join Invtrn on (Invtrn.vno = Invdet.vno)
	Where Invtrn.tran_type in (@OP_STK,@INWARD) --,@ADJ_UP,@ADJ_DN,@STK_WASTE)
	and rec_dat between @date_from And @date_to
	Group by Itm_code,Invtrn.tran_type

----------------------------------------------------------
--	Adjustment Up/Down/Stock Waste/Sales Return between dates
----------------------------------------------------------
	Insert into @opcl (Itm_code,Adj_Qty,Typ)
	Select	 Itm_code
			,Sum(qty)			
			,Invtrn.tran_type	
	From Invdet
	Inner Join Invtrn on (Invtrn.vno = Invdet.vno)
	Where Invtrn.tran_type in (@ADJ_UP,@ADJ_DN,@STK_WASTE,@SALES_RETURN_UP)
	and rec_dat between @date_from And @date_to
	Group by Itm_code,Invtrn.tran_type


----------------------------------------------------------
--	Sales to between period
----------------------------------------------------------
	Insert into @opcl (Itm_code,Sales_Qty,Typ)
	Select	 Itm_code
			,sum(qty) 		
			,99				
	From Saldet
	Inner Join Saltrn on (Saltrn.tran_id = Saldet.tran_id and Saltrn.canceled <> 1)
	Where dtadat between @date_from And @date_to 
	Group by Itm_code


----------------------------------------------------------
--	Final Output
----------------------------------------------------------
	Select	 Itm_Code 
			--Opening
			,sum(op_qty)					as Op_Qty
			,sum(Inward_qty)				as Inward_Qty
			,sum(op_qty) + sum(Inward_qty)	as Op_Inward_Qty
			
			--AdjUp,Down,StockWaste between period
			,sum(Adj_Qty)					as Adj_Qty

			--Sales
			,sum(sales_qty)					as Sales_Qty
			,Items.rtl_prc					as Rtl_Prc
			,sum(sales_qty) * Items.rtl_prc	as Sales_Amt
			,Items.Disc_Amt					as Disc_Per_Item
			,sum(sales_qty) * Items.Disc_Amt					as Disc_Amt
			,(Items.rtl_prc - Items.Disc_Amt)					as Net_Rtl_Prc
			,sum(sales_qty) * (Items.rtl_prc - Items.Disc_Amt)	as Net_Sales_Amt

			--Closing Stock
			,(sum(op_qty) 		+ 
				sum(inward_qty))	- 
				sum(sales_qty)		+					
				sum(adj_qty)					as Cl_Qty

			,((sum(op_qty) 		+ 
				sum(inward_qty))	- 
				sum(sales_qty)		+
				sum(adj_qty)) * Items.rtl_prc	as Cl_Amt

			,Items.[name]						as Itm_Name
			,Items.[shortname]					as Itm_shortname
			,Items.Category_id					as Category_Id
			,Categories.[name]					as Category_Name
			,Categories.[shortname]				as Category_ShortName

	into tmpReportSource
			
	from @opcl	as Opcl
	Left Join Items				On (Opcl.Itm_Code = Items.Code)
	Left Join Categories		On (Items.Category_id = Categories.Code)

	Where Items.Category_id = 
			(Case @CategoryId 
				When 0 then Items.Category_id
				Else @CategoryId
			End)

	and	  Items.Code = 
			(Case @ItemId 
				When 0 then Items.Code
				Else @ItemId
			End)


	Group by Itm_code,Items.[name],Items.[shortname]
			,Items.Category_id,Categories.[name]
			,Categories.[shortname],Items.rtl_prc
			,Items.Disc_Amt


--------------------------------------------------------
--	Used to view data in Sql server only for debug purpose
--------------------------------------------------------
	If (@PreviewEnabled = 1)
	Begin
		Select * from tmpReportSource
		
	End

Return

End