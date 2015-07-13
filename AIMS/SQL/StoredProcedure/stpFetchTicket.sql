
If Exists
(
	select '1' from sysobjects 
	where name = 'stpFetchTicket'
	and type = 'P'
)
	Begin	Drop Procedure stpFetchTicket	End
Go

/*
	Exec stpFetchTicket '08122710020003'
*/
Create Procedure stpFetchTicket
(
	@Tran_Id	as Varchar(20)	
)
As
Begin

	Select 
		 TerSaltrn.tran_id	as TranId
		,TerSaldet.tran_seq	as tran_seq
		,(Case 	When Sizes.shortName is null 
			Then Items.[shortname] 
			Else Items.[shortname] + '(' + Sizes.shortName + ')'
		 End) as ItemName
		,TerSaldet.qty		as ItemQty
		,TerSaldet.rtl_prc	as Rtl_Price
		,TerSaldet.disc_amt	as Disc_Amt
		,TerSaldet.rtl_prc - TerSaldet.disc_amt as Sal_Prc
		,TerSaldet.amt		as Bill_Amt
		,TerSaltrn.paid_amt	as Paid_Amt
		,TerSaltrn.Change_Amt	as Change_Amt

		,TerSaldet.itm_code	as ItemCode
		,Convert(Varchar(12),TerSaltrn.dtadat,3)	as DtaDat
		,1			as Grp	
		

	from TerSaltrn
	Inner Join TerSaldet 	on (TerSaltrn.tran_id = TerSaldet.tran_id)
	Left Join Items 	on (TerSaldet.Itm_code = Items.code)
	Left Join Sizes 	on (Items.Size_id = Sizes.code)
	where TerSaltrn.tran_id = @Tran_Id 

	UNION ALL

	Select 
		 TerSaltrn.tran_id	
		,999			as tran_seq
		,'Total:'		as ItemName
		,SUM(TerSaldet.qty)
		,SUM(TerSaldet.rtl_prc)
		,SUM(TerSaldet.disc_amt)
		,SUM(TerSaldet.rtl_prc) - SUM(TerSaldet.disc_amt)
		,SUM(TerSaldet.amt)
		,AVG(TerSaltrn.paid_amt)
		,AVG(TerSaltrn.Change_Amt) 
		,0 			as ItemCode
		,Convert(Varchar(12),TerSaltrn.dtadat,103)	as DtaDat
		,2			as Grp	
	from TerSaltrn
	Inner Join TerSaldet 	on (TerSaltrn.tran_id = TerSaldet.tran_id)
	Left Join Items 	on (TerSaldet.Itm_code = Items.code)
	where TerSaltrn.tran_id = @Tran_Id 
	Group By TerSaltrn.tran_id,TerSaltrn.dtadat
	Order by TranId,Grp,tran_seq

End
Return