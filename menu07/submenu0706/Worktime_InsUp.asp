<!-- #include virtual="/include/CacheNo.asp" -->
<!-- #include virtual="/include/common.asp" -->

<%

guboon = Request("guboon")						'저장/수정/삭제 FLAG
'월요일
sGijunFromDate = Request("GijunFromDate")


	'#################################################################################################################'
	'요일별 근무시간 관리하기'
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
			Response.Write "<SCRIPT LANGUAGE=JavaScript>alert('정상적으로 등록되었습니다');" &_
								"document.location.href = '" & sUrl & "';" &_
								"parent.OnlineFrame4.location.href = '" & sUrl2 & "';" &_
								"</SCRIPT>"

		else
			db.RollBackTrans
			Call UrlBack("저장중 에러가 발생했습니다.\n\n다시 시도해 주세요")
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
			Response.Write "<SCRIPT LANGUAGE=JavaScript>alert('정상적으로 수정되었습니다');" &_
								"document.location.href = '" & sUrl & "';" &_
								"parent.OnlineFrame4.location.href = '" & sUrl2 & "';" &_
								"</SCRIPT>"
		else
			db.RollBackTrans
			'Response.Write SQL

			Call UrlBack("저장중 에러가 발생했습니다.\n\n다시 시도해 주세요")
		end if		
	
	End If

%>