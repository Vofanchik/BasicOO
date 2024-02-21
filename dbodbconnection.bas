REM  *****  BASIC  *****

Sub db()
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

Sub InsertTextToCursor()
	 oDoc = ThisComponent
	 oText = oDoc.getText()
	 oViewCursor = oDoc.getCurrentController().getViewCursor()
     oCursor = oText.createTextCursorByRange(oViewCursor.getStart())
     oText.insertString( oViewCursor, "more text ...", false )
End Sub

Sub TableCreating()
     oDoc = ThisComponent
	 oViewCursor = oDoc.getCurrentController().getViewCursor()
	 xTable = oDoc.createInstance( "com.sun.star.text.TextTable" )
	 xTable.initialize(5, 8)
	 xTable.HoriOrient = 0 'com::sun::star::text::HoriOrientation::NONE
	 xTable.LeftMargin = 2000
	 xTable.RightMargin = 1500	 
	 oDoc.getText.insertTextContent( oViewCursor, xTable, false )
	 xTable.getCellByPosition(1, 1).setValue(15)
End Sub

Sub Main()
	
End Sub
