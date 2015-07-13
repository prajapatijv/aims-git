
If Exists
(
	select '1' from sysobjects 
	where name = 'stpFetchTicketList'
	and type = 'P'
)
	Begin	Drop Procedure stpFetchTicketList	End
Go

/*
	Exec stpFetchTicketList
*/

Create Procedure stpFetchTicketList
As
Begin

	select 
		 TerSaltrn.tran_id
		,TerDtl.Tkt_Amt
		,TerDtl.Tkt_Qty
		,TerSaltrn.paid_amt
		,TerDtl.Tkt_DiscAmt
		,TerSaltrn.change_amt

		,TerSaltrn.ter_id
		,Convert(Varchar(12),TerSaltrn.dtadat,103) As dtadat	--dmy
		,Convert(Varchar(12),TerSaltrn.dtatim,101) as dtatim	--hh:mm:ss
		,TerSaltrn.dtausr
		,(Case 
			IsNull(TerSaltrn.canceled,0)
			When 1 then 'Cancelled'
			Else ''
		 End) as Cancelled	

	from TerSaltrn
	Inner Join 
	(
		select 
			 tran_id
			,Sum(qty)	as Tkt_Qty
			,Sum(amt)	as Tkt_Amt	
			,Sum(disc_amt)	as Tkt_DiscAmt
		from TerSaldet
		Group by tran_id
	) as TerDtl
	On (TerDtl.tran_id = TerSaltrn.tran_id)
	Order by TerSaltrn.tran_id Desc

End
Return