<!-- #include virtual="/include/CacheNo.asp" -->
<!-- #include virtual="/include/common.asp" -->


<%
	Dtype = Request("Dtype")									'분배방법(1:자동, 0:수동)

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
			Call MsgGoUrl("정상적으로 저장되었습니다.",pageUrl)
		else
			db.RollBackTrans
			Call UrlBack("저장중 에러가 발생했습니다.\n\n다시 시도해 주세요")
		end if
	else
		db.RollBackTrans
		Call UrlBack("수정중 에러가 발생했습니다.\n\n다시 시도해 주세요")
	end if
%>