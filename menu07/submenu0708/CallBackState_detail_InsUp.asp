<!-- #include virtual="/include/CacheNo.asp" -->
<!-- #include virtual="/include/common.asp" -->


<%
	Dtype = Request("Dtype")									'�й���(1:�ڵ�, 0:����)

	On Error Resume next

	SQL_s = "UPDATE TB_USERINFO SET MANUFACTURE = '' "
	db.execute(SQL_s)
	
	Chk = Request("Chk")
	
	Chk_ok = split(Chk,",")
	
	for i = 0 to UBound(Chk_ok)
		if i = 0 then
			User = "'"&Trim(Chk_ok(i))&"'"
		else
			User = User&",'"&Trim(Chk_ok(i))&"'"
		end if

		Chk_S = Request(Trim(Chk_ok(i)))
		Chk_S_ok = split(Chk_S,",")
		for j = 0 to UBound(Chk_S_ok)
			if j = 0 then
				User_S = Trim(Chk_S_ok(j))
			else
				User_S = User_S&","&Trim(Chk_S_ok(j))
			end if

			SQL_s = "UPDATE TB_USERINFO SET MANUFACTURE = '"&User_S&"' WHERE USERID = '"&Trim(Chk_ok(i))&"' "
		
			db.execute(SQL_s)
			'LogWrite "SQL="&SQL, "CallBackState_detail_InsUp.asp", ""
			if db.Errors.count <> 0 then
				Call UrlBack("������ ������ �߻��߽��ϴ�.\n\n�ٽ� �õ��� �ּ���")
			end if
		next
	next
	
	
	SQL = "UPDATE TB_USERINFO SET CALLBACKYN = 'Y' WHERE USERID IN ("&User&")"

	db.begintrans
	db.execute(SQL)
	'LogWrite "SQL="&SQL, "CallBackState_detail_InsUp.asp", ""
	if db.Errors.count = 0 then
		SQL = "UPDATE TB_USERINFO SET CALLBACKYN = 'N' WHERE USERID NOT IN ("&User&")"
		db.execute(SQL)
		
		if db.Errors.count = 0 then
			db.CommitTrans
			'LogWrite "SQL="&SQL, "CallBackState_detail_InsUp.asp", ""
			pageUrl = "CallBackState.asp"
			Call FrameMsgGoUrl("���������� �����Ǿ����ϴ�.",pageUrl)
		else
			db.RollBackTrans
			'LogWrite "ERROR_SQL="&SQL, "CallBackState_detail_InsUp.asp", ""
			Call UrlBack("������ ������ �߻��߽��ϴ�.\n\n�ٽ� �õ��� �ּ���")
		end if
	else
		db.RollBackTrans
		'LogWrite "ERROR_SQL="&SQL, "CallBackState_detail_InsUp.asp", ""
		Call UrlBack("������ ������ �߻��߽��ϴ�.\n\n�ٽ� �õ��� �ּ���")
	end if

%>