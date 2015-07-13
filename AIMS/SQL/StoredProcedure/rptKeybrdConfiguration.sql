

If Exists
(
	select '1' from sysobjects 
	where name = 'rptKeybrdConfiguration'
	and type = 'P'
)
	Begin
		Drop Procedure rptKeybrdConfiguration
	End
Go
/*
	rptKeybrdConfiguration 1001,1
*/
create procedure rptKeybrdConfiguration
(
	 @iKbdCode		numeric(4)	= 0	
	,@PreviewEnabled	tinyint		= 0
)
As
Begin

---------------------------------------------------
-- Set Report Source table
---------------------------------------------------
	If exists (select '1' from sysobjects where name = 'tmpReportSource')
	Begin
		Drop Table tmpReportSource
	End		

	select 	 
		 seq									as seq
		,itm_code								as itm_code
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
		
		,isnull(KeybrdItem.keybrd_code,0)		as keybrd_code
		,isnull(KeybrdSetup.[name],'')			
			+ '  (' + Convert(Varchar(4),isnull(KeybrdItem.keybrd_code,0)) + ')' 	as KeyBrd_Name

		,isnull(TerminalConfig.code,0)			as Ter_Code
		,isnull(TerminalConfig.[shortname],'') 
			+ '  (' + Convert(Varchar(4),isnull(TerminalConfig.code,0)) + ')' 	as Ter_Name

	into tmpReportSource

	from KeybrdItem
	Inner join KeybrdSetup		on (KeybrdSetup.code 	= KeybrdItem.keybrd_code)
	Inner join Items			on (Items.code 			= KeybrdItem.itm_code)
	Left join Sizes				on (Sizes.code 			= Items.size_id)
	Left join Categories		on (Items.Category_Id	= Categories.Code)
	Left join TerminalConfig	on (KeybrdItem.keybrd_code = TerminalConfig.Code)

	where 	KeybrdItem.actv_fg 		= 1
	and 	KeybrdItem.Keybrd_code 	= 
					case 
						when @iKbdCode = 0 then KeybrdItem.Keybrd_code  
						else @iKbdCode 
					end
	order by KeybrdItem.Keybrd_code,Items.category_id,seq 


	--Used to view data in Sql server only for debug purpose
	If (@PreviewEnabled = 1)
	Begin
		Select * from tmpReportSource
	End

	Return
End


