REM  *****  BASIC  *****

Sub Main
oCtx = createUnoService("com.sun.star.sdb.DatabaseContext")
oSrc = oCtx.getByName("New Database")
oDB = oSrc.GetConnection("","")
oStatement = oDB.CreateStatement()
sSQL = "SELECT * FROM Person" 
oResult = oStatement.executeQuery(sSQL)
Do While oResult.next()
Print oResult.getString(1) & " " _
			& oResult.getString(2) & oResult.getString(3) & oResult.getString(4)
Loop

End Sub
