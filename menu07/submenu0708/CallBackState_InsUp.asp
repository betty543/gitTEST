<!-- #include virtual="/include/CacheNo.asp" -->
<!-- #include virtual="/include/common.asp" -->


<%
	Dtype = Request("Dtype")									'�й���(1:�ڵ�, 0:����)

	On Error Resume next
	db.begintrans
	
	INCODE = SESSION("SS_LoginID")
	pageUrl = "CallBackState.asp"
	
	SQL = "UPDATE TB_CONFIG_CALLBACK SET USEYN = 'N', MOCODE = '" &INCODE& "', MODATE = SYSDATE"
	
	db.execute(SQL)
	
	if db.Errors.count = 0 then
		INS_SQL = "INSERT INTO TB_CONFIG_CALLBACK(SEQ, DIVIDEKIND, USEYN, INCODE) VALUES "
		INS_SQL = INS_SQL & "(CONFIG_CALLBACK_SEQ.NEXTVAL,'"&Dtype&"','Y','"&INCODE&"')"
		
		db.execute(INS_SQL)
		if db.Errors.count = 0 then
			db.CommitTrans
			Call MsgGoUrl("���������� ����Ǿ����ϴ�.",pageUrl)
		else
			db.RollBackTrans
			Call UrlBack("������ ������ �߻��߽��ϴ�.\n\n�ٽ� �õ��� �ּ���")
		end if
	else
		db.RollBackTrans
		Call UrlBack("������ ������ �߻��߽��ϴ�.\n\n�ٽ� �õ��� �ּ���")
	end if
%>