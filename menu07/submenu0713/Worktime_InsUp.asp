<!-- #include virtual="/include/CacheNo.asp" -->
<!-- #include virtual="/include/common.asp" -->

<%

guboon = Request("guboon")						'����/����/���� FLAG
'������
sGijunFromDate = Request("GijunFromDate")


	'#################################################################################################################'
	'���Ϻ� �ٹ��ð� �����ϱ�'
	'#################################################################################################################'

	'GIJUNDATE, INCODE, JOBGB,  WORKHOUR,	WORKDESC,	CHANGEHOUR,	PROCESSHOUR,	INDATE,		MOCODE,		MODATE
	SQL = "select * from TB_PERSON_ETC where GijunDate = '" & sGijunDate & "' and INCODE = '" & INCODE & "'"
	SQL = SQL & "	AND		JOBGB = '" & sJobGb & "'"
	Set Rs = server.createObject("ADODB.Recordset")
	Rs.open SQL,db
	On Error Resume next
	db.begintrans

	If Rs.Eof Or Rs.bof Then

		SQL ="INSERT INTO TB_PERSON_ETC (GIJUNDATE, INCODE, JOBGB,  WORKHOUR,	WORKDESC,	CHANGEHOUR,	PROCESSHOUR) VALUES "
		SQL = SQL & "('" & sGijunDate & "','" & INCODE & "','" & sJobGb & "'," & sCnt & ",'" & sUseMemo & "',"&CInt(CHANGEHOUR)&","&CInt(PROCESSHOUR)&")"

		db.execute(SQL)

		if db.Errors.count = 0 then
			db.CommitTrans
			sUrl = "Online_Frame3.asp?GijunDate="&sGijunDate
			sUrl2 = "Online_Frame4.asp?whereCD1="&whereCD1&"&whereCD2="&whereCD2&"&GijunFromDate="&sGijunFromDate&"&GijunToDate="&sGijunToDate&"&INCODE=" & INCODE
			Response.Write "<SCRIPT LANGUAGE=JavaScript>alert('���������� ��ϵǾ����ϴ�');" &_
								"document.location.href = '" & sUrl & "';" &_
								"parent.OnlineFrame4.location.href = '" & sUrl2 & "';" &_
								"</SCRIPT>"

		else
			db.RollBackTrans
			Call UrlBack("������ ������ �߻��߽��ϴ�.\n\n�ٽ� �õ��� �ּ���")
		end if		
	Else

'		SQL ="	UPDATE TB_PERSON_ETC	SET	 TB_PERSON_ETC (GIJUNDATE, INCODE, JOBGB,  WORKHOUR,	WORKDESC,	CHANGEHOUR,	PROCESSHOUR) VALUES "
		SQL = "UPDATE TB_PERSON_ETC SET	WORKHOUR = " & sCnt & ", WORKDESC = '" & sUseMemo & "', CHANGEHOUR = "&CInt(CHANGEHOUR)& ", PROCESSHOUR = "&CInt(PROCESSHOUR)
		SQL = SQL & "	where GijunDate = '" & sGijunDate & "' and INCODE = '" & INCODE & "'"
		SQL = SQL & "	AND		JOBGB = '" & sJobGb & "'"

		
		db.execute(SQL)

		if db.Errors.count = 0 then
			db.CommitTrans

			sUrl = "Online_Frame3.asp?GijunDate="&sGijunDate
			sUrl2 = "Online_Frame4.asp?whereCD1="&whereCD1&"&whereCD2="&whereCD2&"&GijunFromDate="&sGijunFromDate&"&GijunToDate="&sGijunToDate&"&INCODE=" & INCODE
			Response.Write "<SCRIPT LANGUAGE=JavaScript>alert('���������� �����Ǿ����ϴ�');" &_
								"document.location.href = '" & sUrl & "';" &_
								"parent.OnlineFrame4.location.href = '" & sUrl2 & "';" &_
								"</SCRIPT>"
		else
			db.RollBackTrans
			'Response.Write SQL

			Call UrlBack("������ ������ �߻��߽��ϴ�.\n\n�ٽ� �õ��� �ּ���")
		end if		
	
	End If

%>