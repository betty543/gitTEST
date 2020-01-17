<!-- #include virtual="/Include/Common.asp" -->
<%
	Server.ScriptTimeout = 90000
	Response.ContentType = "application/vnd.ms-excel; name='My_Excel'"
	Call Response.AddHeader("Content-Disposition", "attachment; filename=문자메시지전송내역_" &Date()& ".xls")	'바로저장하기
	Call Response.AddHeader("Content-Description", "ASP Generated Data")

%>
<%
	QueryYN = request("QueryYN")
	FromDate = request("FromDate")
	if FromDate ="" then
		FromDate = date()
	end if
	ToDate = request("ToDate")
	if ToDate ="" then
		ToDate = date()
	end if
	whereCD1 = Trim(request("whereCD1"))
	whereCD2 = Trim(request("whereCD2"))
	whereCD3 = Trim(request("whereCD3"))
%>

<table width="940"  border="1" cellpadding="0" cellspacing="0" bgcolor="#EFECE5" align="center" bordercolor="black" bordercolordark="white" bordercolorlight="black">

	<tr height="25" bgcolor="#F3F3F3" align="center">
		<td><b>순번</b></td>
		<td><b>전송요청일시</b></td>
		<td><b>그룹</b></td>
		<td><b>전송자</b></td>
		<td><b>수신휴대폰</b></td>
		<td><b>발신번호</b></td>
		<td><b>전송내용</b></td>
		<td ><b>전송결과</b></td>
	</tr>

<%
	if QueryYN = "Y" then		

		SQL = "	SELECT SM_SDMBNO,SM_RVMBNO, SM_MSG, '2' as SM_STATUS, convert(char(19),SM_Indate,121) as SM_Sdate, SM_CODE1, SM_CODE2"
		SQL = SQL & "	FROM	SMS.DBO.SMS_Reserve"
		SQL = SQL & "	WHERE	CONVERT(CHAR(10),SM_Indate,121) >= '" & FROMDATE & "'"
		SQL = SQL & "	AND		CONVERT(CHAR(10),SM_Indate,121) <= '" & ToDate & "'"
		'전송요청자
		IF whereCD1 <> "" THEN
			SQL = SQL & "	AND		SM_CODE1 = '" & whereCD1 & "'"
		END IF
		'전화번호
		IF whereCD2 <> "" THEN
			SQL = SQL & "	AND		SM_SDMBNO LIKE '%" & whereCD3 & "%'"
		END IF
		'수신자
		IF whereCD2 <> "" THEN
			SQL = SQL & "	AND		SM_CODE2 LIKE '%" & whereCD2 & "%'"
		END IF
		SQL = SQL & "	UNION ALL "
		SQL = SQL & "	SELECT SM_SDMBNO,SM_RVMBNO, SM_MSG,SM_STATUS, convert(char(19),SM_Sdate,121) as SM_Sdate, SM_CODE1, SM_CODE2"
		SQL = SQL & "	FROM	SMS.DBO.SMS_BACK"
		SQL = SQL & "	WHERE	CONVERT(CHAR(10),SM_Sdate,121) >= '" & FROMDATE & "'"
		SQL = SQL & "	AND		CONVERT(CHAR(10),SM_Sdate,121) <= '" & ToDate & "'"
		'전송요청자
		IF whereCD1 <> "" THEN
			SQL = SQL & "	AND		SM_CODE1 = '" & whereCD1 & "'"
		END IF
		'전화번호
		IF whereCD2 <> "" THEN
			SQL = SQL & "	AND		SM_SDMBNO LIKE '%" & whereCD3 & "%'"
		END IF
		'수신자
		IF whereCD2 <> "" THEN
			SQL = SQL & "	AND		SM_CODE2 LIKE '%" & whereCD2 & "%'"
		END IF
		SQL = SQL & "	ORDER BY SM_Sdate desc"
		SET RS = DB.EXECUTE(SQL)

i = 0
		DO UNTIL RS.EOF
			i = i + 1
			sDate = RS("SM_Sdate")
			sGROUP = db_getCodeName("Z04",RS("SM_CODE2"))
			sUSERID = db_GetUSERNAME(RS("SM_CODE1"))
			sCELLPHONE = RS("SM_SDMBNO")
			sREPLYPHONE = RS("SM_RVMBNO")
			sMESSAGE = RS("SM_MSG")
			if RS("SM_STATUS") = "1" then
				sRESULT = "성공"
			elseif RS("SM_STATUS") = "2" then
				sRESULT = "예약"
			else
				sRESULT = "실패"
			end if
%>
	<tr height="25" bgcolor="#ffffff" align="center">
		<td><%=i%></td>
		<td><%=sDate%></td>
		<td><%=sGROUP%></td>
		<td><%=sUSERID%></td>
		<td><%=sCELLPHONE%></td>
		<td><%=sREPLYPHONE%></td>
		<td title="<%=sMESSAGE%>" align='left'>&nbsp;<%=CutString(sMESSAGE, 30, "...")%></td>
		<td ><%=sRESULT%></td>
	</tr>

<%
			RS.MOVENEXT
		LOOP

	end if
%>


</table>


<!-- #include virtual="/Include/Bottom.asp" -->