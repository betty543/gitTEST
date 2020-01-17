
<!-- #include virtual="/Include/Top.asp" -->

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

<div id='CalendarLayer' style='display:none;'><iframe name='CalendarFrame' src='/include/Calendar.asp' width='172' height='177' border='0' frameborder='0' scrolling='no'></iframe></div>

<script>

	function fn_Search() {

		document.inUpFrm.QueryYN.value = "Y";
		document.inUpFrm.submit();
	}
	
	function fn_Xls() {
		location.href="smsgroup_XLS.asp?<%=pageWHERE%>"
	}
</script>

<table border="0" width="940" cellpadding="0" cellspacing="3" bgcolor="#EFECE5" align="center">
	<tr bgcolor="#FFFFFF">
		<td>
		
			<form method="post" name="inUpFrm" style="margin:0">
			<input type="hidden" name="QueryYN" value="">
			<table width="100%" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">

			    <tr>
			        <td width="60" bgcolor="#EEF6FF" class="TDCont" align='center'>조회기간</td>
			        <td  bgcolor="#FFFFFF" colspan=3 width=200>
			        	<input value="<%=FromDate%>" name="FromDate" readonly type="text" size="10" onfocus="setFocusColor(this);" onClick="new CalendarFrame.Calendar(this);" onblur="setOutColor(this);">
				    	~
				    	<input value="<%=ToDate%>" name="ToDate" readonly type="text" size="10" onfocus="setFocusColor(this);" onClick="new CalendarFrame.Calendar(this);" onblur="setOutColor(this);">	
			        </td>


			        <td colspan='2' rowspan="2" bgcolor="#FFFFFF" align="center">
			        	<img src="/Images/Btn/BtnSearch.gif" style="cursor:hand;" onClick="fn_Search();">
			        	<%IF SS_Login_Secgroup="A" Or SS_Login_Secgroup="B" THEN%><br><br><img src="/Images/Btn/BtnExcel.gif" style="cursor:hand;" onClick="fn_Xls();"><%END IF%>
			        </td>
				</tr>

			</table>
			</form>
		</td>
	</tr>
</table>


<table width="940" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
<table border="0" width="940" cellspacing="1" cellpadding="0" width="100%" bgcolor="#CCCCCC" align="center">
	<tr height="25" bgcolor="#F3F3F3" align="center">
		<td><b>순번</b></td>
		<td><b>그룹</b></td>
		<td><b>일자</b></td>
		<td><b>총건수</b></td>
		<td><b>성공건수</b></td>
		<td><b>실패건수</b></td>
	</tr>

<%
	i = 0
	if QueryYN = "Y" then	
		'총

		SQL = "	SELECT SDATE, SGROUP, SUM(CNT1) CNT1, SUM(CNT2) CNT2, SUM(CNT3) CNT3 "
		SQL = SQL & "	FROM ( "
		SQL = SQL & "	SELECT  CONVERT(CHAR(10),SM_Sdate,121) AS SDATE, SM_CODE2  AS SGROUP, count(SM_STATUS) cnt1, 0 CNT2, 0 CNT3"
		SQL = SQL & "	FROM	SMS.DBO.SMS_BACK"
		SQL = SQL & "	WHERE	CONVERT(CHAR(10),SM_Sdate,121) >= '" & FROMDATE & "'"
		SQL = SQL & "	AND		CONVERT(CHAR(10),SM_Sdate,121) <= '" & ToDate & "'"
		SQL = SQL & "	GROUP BY 	CONVERT(CHAR(10),SM_Sdate,121), 	SM_CODE2"
		'성공
		SQL = SQL & "	UNION ALL SELECT CONVERT(CHAR(10),SM_Sdate,121) AS SDATE, SM_CODE2 AS SGROUP, 0 CNT1, count(SM_STATUS) cnt2, 0 CNT3"
		SQL = SQL & "	FROM	SMS.DBO.SMS_BACK"
		SQL = SQL & "	WHERE	CONVERT(CHAR(10),SM_Sdate,121) >= '" & FROMDATE & "'"
		SQL = SQL & "	AND		CONVERT(CHAR(10),SM_Sdate,121) <= '" & ToDate & "'"
		SQL = SQL & "	AND		SM_STATUS IN ('1')"
		SQL = SQL & "	GROUP BY 	CONVERT(CHAR(10),SM_Sdate,121), 	SM_CODE2"
		'실패
		SQL = SQL & "	UNION ALL SELECT  CONVERT(CHAR(10),SM_Sdate,121) AS SDATE, SM_CODE2 AS SGROUP, 0 CNT1, 0 CNT2, count(SM_STATUS) cnt3"
		SQL = SQL & "	FROM	SMS.DBO.SMS_BACK"
		SQL = SQL & "	WHERE	CONVERT(CHAR(10),SM_Sdate,121) >= '" & FROMDATE & "'"
		SQL = SQL & "	AND		CONVERT(CHAR(10),SM_Sdate,121) <= '" & ToDate & "'"
		SQL = SQL & "	AND		SM_STATUS NOT IN ('1')"
		SQL = SQL & "	GROUP BY 	CONVERT(CHAR(10),SM_Sdate,121), 	SM_CODE2 ) A GROUP BY SDATE, SGROUP "
		SQL = SQL & "	ORDER BY SGROUP, SDATE "

		'RESPONSE.WRITE SQL

		SET RS = DB.EXECUTE(SQL)

		DO UNTIL RS.EOF
			i = i + 1
			sGROUP = db_getCodeName("Z04",RS("SGROUP"))
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