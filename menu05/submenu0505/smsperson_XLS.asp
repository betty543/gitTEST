<!-- #include virtual="/Include/Common.asp" -->
<%
	Server.ScriptTimeout = 90000
	Response.ContentType = "application/vnd.ms-excel; name='My_Excel'"
	Call Response.AddHeader("Content-Disposition", "attachment; filename=���κ����۳���_" &Date()& ".xls")	'�ٷ������ϱ�
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

	pageWHERE = "QueryYN="&QueryYN&"&FromDate="&FromDate&"&ToDate="&ToDate

%>


<table width="940"  border="1" cellpadding="0" cellspacing="0" bgcolor="#EFECE5" align="center" bordercolor="black" bordercolordark="white" bordercolorlight="black">
	<tr height="25" bgcolor="#F3F3F3" align="center">
		<td><b>����</b></td>
		<td><b>�����</b></td>
		<td><b>����</b></td>
		<td><b>�ѰǼ�</b></td>
		<td><b>�����Ǽ�</b></td>
		<td><b>���аǼ�</b></td>
	</tr>

<%
	I = 0
	if QueryYN = "Y" then	
		'��

		SQL = "	SELECT SDATE, SGROUP, SUM(CNT1) CNT1, SUM(CNT2) CNT2, SUM(CNT3) CNT3 "
		SQL = SQL & "	FROM ( "
		SQL = SQL & "	SELECT  CONVERT(CHAR(10),SM_Sdate,121) AS SDATE, SM_CODE1  AS SGROUP, count(SM_STATUS) cnt1, 0 CNT2, 0 CNT3"
		SQL = SQL & "	FROM	SMS.DBO.SMS_BACK"
		SQL = SQL & "	WHERE	CONVERT(CHAR(10),SM_Sdate,121) >= '" & FROMDATE & "'"
		SQL = SQL & "	AND		CONVERT(CHAR(10),SM_Sdate,121) <= '" & ToDate & "'"
		SQL = SQL & "	GROUP BY 	CONVERT(CHAR(10),SM_Sdate,121), 	SM_CODE1"
		'����
		SQL = SQL & "	UNION ALL SELECT CONVERT(CHAR(10),SM_Sdate,121) AS SDATE, SM_CODE1 AS SGROUP, 0 CNT1, count(SM_STATUS) cnt2, 0 CNT3"
		SQL = SQL & "	FROM	SMS.DBO.SMS_BACK"
		SQL = SQL & "	WHERE	CONVERT(CHAR(10),SM_Sdate,121) >= '" & FROMDATE & "'"
		SQL = SQL & "	AND		CONVERT(CHAR(10),SM_Sdate,121) <= '" & ToDate & "'"
		SQL = SQL & "	AND		SM_STATUS IN ('1')"
		SQL = SQL & "	GROUP BY 	CONVERT(CHAR(10),SM_Sdate,121), 	SM_CODE1"
		'����
		SQL = SQL & "	UNION ALL SELECT  CONVERT(CHAR(10),SM_Sdate,121) AS SDATE, SM_CODE1 AS SGROUP, 0 CNT1, 0 CNT2, count(SM_STATUS) cnt3"
		SQL = SQL & "	FROM	SMS.DBO.SMS_BACK"
		SQL = SQL & "	WHERE	CONVERT(CHAR(10),SM_Sdate,121) >= '" & FROMDATE & "'"
		SQL = SQL & "	AND		CONVERT(CHAR(10),SM_Sdate,121) <= '" & ToDate & "'"
		SQL = SQL & "	AND		SM_STATUS NOT IN ('1')"
		SQL = SQL & "	GROUP BY 	CONVERT(CHAR(10),SM_Sdate,121), 	SM_CODE1 ) A GROUP BY SDATE, SGROUP ORDER BY SDATE, SGROUP"

		'RESPONSE.WRITE SQL

		SET RS = DB.EXECUTE(SQL)
i = 0
		DO UNTIL RS.EOF
			i = i + 1
			sGROUP = db_getUserName(RS("SGROUP"))
			sDate = RS("SDATE")
			CNT1 = RS("CNT1")
			CNT2 = RS("CNT2")
			CNT3 = RS("CNT3")

%>
	<tr height="25" bgcolor="#ffffff" align="center">
		<td><%=i%></td>
		<td><%=sGROUP%></td>
		<td><%=sDate%></td>
		<td><%=CNT1%></td>
		<td><%=CNT2%></td>
		<td><%=CNT3%></td>

	</tr>


<%
			RS.MOVENEXT
		LOOP

	end if
%>


</table>


<!-- #include virtual="/Include/Bottom.asp" -->