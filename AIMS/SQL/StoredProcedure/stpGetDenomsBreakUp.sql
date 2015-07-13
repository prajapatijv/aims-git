--select * from Denoms

If Exists
(
	select '1' from sysobjects 
	where name = 'stpGetDenomsBreakUp'
	and type = 'P'
)
	Begin
		Drop Procedure stpGetDenomsBreakUp
	End
Go
/*
	stpGetDenomsBreakUp 1
*/
create procedure stpGetDenomsBreakUp
(
	@Amt 	numeric(10,2)	
)
As
Begin
	select @Amt as DenomBreakUp

	union 

	select  distinct ceiling(@Amt/denom_struc) * denom_struc
	from Denoms
	where denom_id <= (case when @Amt <= 20 then 9 else 7 end)

	Return
End
