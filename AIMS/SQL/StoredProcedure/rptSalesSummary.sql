
/*
select * from Saltrn
select * from Saldet
select * from UserMast
*/
If Exists
(
	select '1' from sysobjects 
	where name = 'rptSalesSummary'
	and type = 'P'
)
	Begin	Drop Procedure rptSalesSummary	End
Go

/*
	Exec rptSalesSummary '01/01/2008','31/12/2009',0,0,0,2,1
*/

Create Procedure rptSalesSummary
(
	 @from_dt	Varchar(20) 
	,@to_dt		Varchar(20) 
	,@TerId		SmallInt
	,@UserId	SmallInt
	,@Status	TinyInt		= 0			--0:All/ 1:Active /2:Rejected	
	,@GenAt		TinyInt		= 1			--1:Server/ 2:Terminal
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
	If (@GenAt=1)
	Begin
		select 
			 Saltrn.tran_id
			,Dtl.Tkt_Amt
			,Dtl.Tkt_Qty
			,Saltrn.Paid_amt
			,Dtl.Tkt_DiscAmt
			,Saltrn.change_amt

			,Saltrn.ter_id
			,Convert(Varchar(12), Saltrn.dtadat,@datfmt) As dtadat	--dmy
			,Convert(Varchar(12), Saltrn.dtatim,101) As dtatim	--hh:mm:ss
			,Saltrn.dtausr
			,UserMast.UName	+ '[' + convert(varchar(4),Saltrn.dtausr) + ']'	As UserName
			,(Case 
				IsNull(Saltrn.canceled,0)
				When 1 then 'Cancelled'
				Else ''

			 End) as Cancelled	

		into tmpReportSource
		from  Saltrn
		Inner Join 
		(
			select 
				 tran_id
				,Sum(qty)	as Tkt_Qty
				,Sum(amt)	as Tkt_Amt	
				,Sum(disc_amt)	as Tkt_DiscAmt
			from  Saldet
			Group by tran_id
		) as  Dtl		On (Dtl.tran_id =  Saltrn.tran_id)
		Left Join UserMast 	On (UserMast.Uid =  Saltrn.dtausr)

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
	End
	Else
	Begin
		select 
			 TerSaltrn.tran_id
			,Dtl.Tkt_Amt
			,Dtl.Tkt_Qty
			,TerSaltrn.Paid_amt
			,Dtl.Tkt_DiscAmt
			,TerSaltrn.change_amt

			,TerSaltrn.ter_id
			,Convert(Varchar(12), TerSaltrn.dtadat,@datfmt) As dtadat	--dmy
			,Convert(Varchar(12), TerSaltrn.dtatim,101) As dtatim	--hh:mm:ss
			,TerSaltrn.dtausr
			,UserMast.UName	+ '[' + convert(varchar(4),TerSaltrn.dtausr) + ']'	As UserName
			,(Case 
				IsNull(TerSaltrn.canceled,0)
				When 1 then 'Cancelled'
				Else ''
			 End) as Cancelled	

		into tmpReportSource
		from  TerSaltrn
		Inner Join 
		(
			select 
				 tran_id
				,Sum(qty)	as Tkt_Qty
				,Sum(amt)	as Tkt_Amt	
				,Sum(disc_amt)	as Tkt_DiscAmt
			from  TerSaldet
			Group by tran_id
		) as  Dtl		On (Dtl.tran_id =  TerSaltrn.tran_id)
		Left Join UserMast 	On (UserMast.Uid =  TerSaltrn.dtausr)

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
	End
End

	--Used to view data in Sql server only for debug purpose
	If (@PreviewEnabled = 1)
	Begin
		Select * from tmpReportSource
	End


Return