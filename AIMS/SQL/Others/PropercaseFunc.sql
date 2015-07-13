CREATE Function ProperCase
(
	@String 	Varchar(200)
)

Returns Varchar(200)

As 

Begin

	Declare @lString Varchar(200)
	Declare @uString Varchar(200)
	
	Declare @iLen Int 
	
	Set @lString = ''
	Set @uString = @string
	Set @ilen =0
	
	While @iLen <= Len(@uString)
	Begin
		If  Substring(@uString,@iLen,1) = ' ' 
			Begin
				Set @lString = @lString + Upper(Substring(@uString,@iLen+1,1))
			end
		else
			Begin
				Set @lString = @lString + Lower(Substring(@uString,@iLen+1,1))
			end
		Set @iLen = @iLen + 1
	end
	Return (@lstring)
end



