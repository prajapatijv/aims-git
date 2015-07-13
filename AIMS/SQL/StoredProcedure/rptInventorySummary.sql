If Exists
(
	select '1' from sysobjects 
	where name = 'rptInventorySummary '
	and type = 'P'
)
	Begin
		Drop Procedure rptInventorySummary 
	End
Go

/*
	Exec rptInventorySummary '01/01/2008','31/12/2009',0,0,1,1
*/

Create Procedure [rptInventorySummary]
(
	 @from_dt	Varchar(20) 
	,@to_dt		Varchar(20) 
	,@TerId		SmallInt
	,@UserId	SmallInt
	,@Tran_type			TinyInt		= 0			--0:All/ 1/2/3/4/11/12/13	
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
		,Invtrn.Doc_No
		,Dtl.Inv_Qty
		,Dtl.Inv_Amt
		,Dtl.Inv_Avg_Rtl_prc

		,Invtrn.ter_id

		,Convert(Varchar(12), Invtrn.rec_dat,@datfmt) As Recdat	--dmy
		,Convert(Varchar(12), Invtrn.dtadat,@datfmt) As dtadat	--dmy
		,Convert(Varchar(12), Invtrn.dtatim,101) As dtatim	--hh:mm:ss
		,Invtrn.dtausr
		,UserMast.UName	+ '[' + convert(varchar(4),Invtrn.dtausr) + ']'	As UserName
		,(Case 
			IsNull(Invtrn.tran_type,0)
			When 1 then 'Opening Balance'
			When 2 then 'Stock Inward'
			When 3 then 'Stock Adjustment Up'
			When 3 then 'Receive from Store'
			When 11 then 'Stock Adjustment Down'
			When 12 then 'Stock Waste'
			When 13 then 'Issue For Sale'
			Else ''

		 End) as TranType
		,Invtrn.remarks
	into tmpReportSource
	from  Invtrn
	Inner Join 
	(
		select 
			 vno
			,Sum(qty)	as Inv_Qty
			,Sum(amt)	as Inv_Amt	
			,Avg(rtl_prc)	as Inv_Avg_Rtl_prc
		from  Invdet
		Group by vno
	) as  Dtl		On (Dtl.vno =  Invtrn.vno)
	Left Join UserMast 	On (UserMast.Uid =  Invtrn.dtausr)

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

	--Date Range
	And  Invtrn.dtadat Between @date_from and @date_to
	

	--Transction Type
	And  IsNull(Invtrn.tran_type,0) In (
					Case @Tran_type
						When 0 then IsNull(Invtrn.tran_type,0)		--All
						else @Tran_type
					end
					)


	--Used to view data in Sql server only for debug purpose
	If (@PreviewEnabled = 1)
	Begin
		Select * from tmpReportSource
	End

Return

End