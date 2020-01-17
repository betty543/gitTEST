<!-- #include virtual="/include/CacheNo.asp" -->
<!-- #include virtual="/include/common.asp" -->

<%

guboon = Request("guboon")
sIdx = Request("Idx")
sDnis = Request("DNIS")
sTelNo = Request("TelNo")

sUSEYN = Request("sUSEYN")	


If guboon = "INS"Then 

	Set Rs = server.createObject("ADODB.Recordset")
	SQL = "select idx from TB_Reject where Dnis = '" & sDnis & "' and TelNo ='" & sTelNo & "'"

	Rs.open SQL,db
	
	If Rs.Eof Or Rs.bof Then
		On Error Resume next
		db.begintrans


		SQL =				" insert into TB_Reject (Dnis,Telno,USEYN)"
		SQL = SQL &	"  values('" & sDnis & "', '" & sTelNo & "', '" & sUSEYN & "')"

		'Response.Write SQL
		db.execute(SQL)

		If db.Errors.count = 0 Then 
			db.CommitTrans
			response.write "<script>parent.ListFrame.location.href='Callreject_List.asp';</script>"
			Call MsgGoUrl("정상적으로 저장되었습니다.", "Callreject_Detail.asp?guboon=INS")
		Else 
			db.RollBackTrans
			Call UrlBack("저장중 에러가 발생했습니다.\n\n다시 시도해 주세요")
		End If 

	Else

		Call UrlBack("이미 존재하는 자료입니다..\n\n확인 후 다시 입력해 주세요")
	
	end If
	
	Rs.Close
	set Rs = NOTHING	



ElseIf guboon = "UP" Then

		On Error Resume next
		db.begintrans

		SQL =				" update TB_Reject set"
		SQL = SQL &	"	DNIS = '" & sDNIS & "', "
		SQL = SQL &	"	TelNo = '" & sTelNo & "', "
		SQL = SQL &	"	USEYN = '" & sUSEYN & "' "

		SQL = SQL &	" where idx = '"& sIdx & "'"
'
		db.execute(SQL)


		If db.Errors.count = 0 Then 
			db.CommitTrans
			response.write "<script>parent.ListFrame.location.reload();</script>"
			Call MsgGoUrl("정상적으로 저장되었습니다.", "Callreject_Detail.asp?guboon=INS")
		Else 
			db.RollBackTrans
			Call UrlBack("저장중 에러가 발생했습니다.\n\n다시 시도해 주세요")
		End If 


ElseIf guboon = "DEL" Then 

	On Error Resume next
	db.begintrans

	SQL ="delete from TB_Reject where idx = '" & sIdx & "'"
	db.execute(SQL)

	If db.Errors.count = 0 Then 
		db.CommitTrans
		response.write "<script>parent.ListFrame.location.href='Callreject_List.asp';</script>"
		Call MsgGoUrl("정상적으로 삭제되었습니다.", "/menu04/Submenu0410/Callreject_Detail.asp?guboon=INS")
	Else 
		db.RollBackTrans
		Call UrlBack("삭제중 에러가 발생했습니다.\n\n다시 시도해 주세요")
	End If 

End If 

%>