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
			Call MsgGoUrl("���������� ����Ǿ����ϴ�.", "Callreject_Detail.asp?guboon=INS")
		Else 
			db.RollBackTrans
			Call UrlBack("������ ������ �߻��߽��ϴ�.\n\n�ٽ� �õ��� �ּ���")
		End If 

	Else

		Call UrlBack("�̹� �����ϴ� �ڷ��Դϴ�..\n\nȮ�� �� �ٽ� �Է��� �ּ���")
	
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
			Call MsgGoUrl("���������� ����Ǿ����ϴ�.", "Callreject_Detail.asp?guboon=INS")
		Else 
			db.RollBackTrans
			Call UrlBack("������ ������ �߻��߽��ϴ�.\n\n�ٽ� �õ��� �ּ���")
		End If 


ElseIf guboon = "DEL" Then 

	On Error Resume next
	db.begintrans

	SQL ="delete from TB_Reject where idx = '" & sIdx & "'"
	db.execute(SQL)

	If db.Errors.count = 0 Then 
		db.CommitTrans
		response.write "<script>parent.ListFrame.location.href='Callreject_List.asp';</script>"
		Call MsgGoUrl("���������� �����Ǿ����ϴ�.", "/menu04/Submenu0410/Callreject_Detail.asp?guboon=INS")
	Else 
		db.RollBackTrans
		Call UrlBack("������ ������ �߻��߽��ϴ�.\n\n�ٽ� �õ��� �ּ���")
	End If 

End If 

%>