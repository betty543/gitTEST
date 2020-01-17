<!-- #include virtual="/Include/Common.asp" -->
<%
	'Server.ScriptTimeout = 90000
	'Response.ContentType = "application/vnd.ms-excel; name='My_Excel'"
	'Call Response.AddHeader("Content-Disposition", "attachment; filename=전체상담기록대장_" &Date()& ".xls")	'바로저장하기
	'Call Response.AddHeader("Content-Description", "ASP Generated Data")

%>
<%
	'####### 파라미터 ##################################################################################
	QueryYN = request("QueryYN")
	FromDate = request("FromDate")
	ToDate = request("ToDate")
	whereCD1 = Trim(request("whereCD1")) '성별
	whereCD2 = Trim(request("whereCD2")) '상담방법
	whereCD3 = Trim(request("whereCD3")) '의뢰인
	whereCD4 = Trim(request("whereCD4")) '상담분야
	whereCD5 = Trim(request("whereCD5")) '소속
	whereCD6 = Trim(request("whereCD6")) '계급구분
	whereCD7 = Trim(request("whereCD7")) '계급구분2
	whereCD8 = Trim(request("whereCD8"))	'성명
	whereCD9 = Trim(request("whereCD9"))	'전화번호
	whereCD10 = Trim(request("whereCD10"))	'소속
	whereCD11 = Trim(request("whereCD11"))	'처리결과
	whereCD12 = Trim(request("whereCD12"))	'처리결과
	if FromDate = "" then
		FromDate = left(date,7)&"-01"
	end if
	if ToDate = "" then
		ToDate = date
	end if

	pageSize = 10
	pageSector = 10
	if curPage = "" then curPage = 1 end If

	SQL = "	SELECT *, CONVERT(VARCHAR(19),JUBTIME,121) JUBTIME1   FROM TB_LIFECALLHISTORY"
	SQL = SQL & "		WHERE	JUBDATE >= '" & FromDate & "'"
	SQL = SQL & "		AND     JUBDATE <= '" & ToDate & "'"
	SQL = SQL & "	ORDER BY JUBTIME"

	Set Rs = server.createObject("ADODB.Recordset")
	Rs.open SQL,db

%>

<table cellspacing="0" align="center" border="1" bordercolor="black" bordercolordark="white" bordercolorlight="black">

<table width="940" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
<table cellspacing="0" align="center" border="1" bordercolor="black" bordercolordark="white" bordercolorlight="black">
	<tr height="25" bgcolor="#EEF6FF" align="center">
		<td class="TDCont" align='center'>연번</td>
		<td class="TDCont" align='center'><b>상담일시</b></td>
		<td class="TDCont" align='center'><b>성별</b></td>
		<td class="TDCont" align='center'><b>상담방법</b></td>
		<td class="TDCont" align='center'><b>성명</b></td>
		<td class="TDCont" align='center'><b>소속</b></td>
		<td class="TDCont" align='center'><b>계급</b></td>
		<td class="TDCont" align='center'><b>상담분야</b></td>
		<td class="TDCont" align='center'><b>의뢰인</b></td>
		<td class="TDCont" align='center'><b>인지경로</b></td>
		<td class="TDCont" align='center'><b>가해자</b></td>
		<td class="TDCont" align='center'><b>상담관</b></td>
		<td class="TDCont" align='center'><b>조치결과</b></td>
	</tr>

				<%'####### 실제자료가 들어간다. %>
				<%
				'---------------------------------------------------------------------------------------------------------------------
				'데이타 뽑아오기
				'---------------------------------------------------------------------------------------------------------------------

	i = 0
	DO UNTIL RS.EOF
	i = i + 1


	db_JUBSEQ = rs("JUBSEQ")
	db_JUBDATE = rs("JUBDATE")
	db_JUBTIME = rs("JUBTIME1")
	db_IOFLAG = rs("IOFLAG")
	db_CUSTNO = rs("CUSTNO")
	db_CUSTNAME = rs("CUSTNAME")
	db_TELNO = rs("TELNO")
	db_TELNO2 = rs("TELNO2")
	db_SEXGB = rs("SEXGB")
	db_CHANNELGB = rs("CHANNELGB")
	db_REQUESTERGB = rs("REQUESTERGB")
	db_CONSULTGB = rs("CONSULTGB")
	db_CONSULTETCGB = rs("CONSULTETCGB")
	db_SOSOKGB = rs("SOSOKGB")
	db_SOSOKETCGB = rs("SOSOKETCGB")
	db_LEVEL1 = rs("LEVEL1")
	db_LEVEL2 = rs("LEVEL2")
	db_ACLASS = rs("ACLASS") '소속
	db_BCLASS = rs("BCLASS") '소속세분류
	db_CCLASS = rs("CCLASS")
	db_CHANNEL = rs("CHANNEL")
	db_CALLFLAG = rs("CALLFLAG")
	db_CALLKIND = rs("CALLKIND")
	db_QUESTION = rs("QUESTION")
	db_REPLY = rs("REPLY")
	db_RESULTGB = rs("RESULTGB")
	db_RESERVEDATE = rs("RESERVEDATE")
	db_RESERVETIME = rs("RESERVETIME")
	db_PROCESSGB = rs("PROCESSGB")
	db_CALLID = rs("CALLID")
	db_RECORDFILE = rs("RECORDFILE")
	db_INCODE = rs("INCODE")
	db_EMERYN = rs("EMERYN")


%>

		<tr bgcolor="#FFFFFF">

			<td align="center"><%=i%></td>
			<td align="center"><%=db_JUBTIME%><% if db_EMERYN="Y" then%><font color="#0000ff"> (긴급)</font><%else%> (일반)<%end if%></td>
			<td align="center"><%if db_SEXGB = "1" then%>남<%else%>녀<%end if%></td>
			<td align="center"><%=db_getCodeName("C00",db_ACLASS)%><br><%=db_getCodeName("C01",db_CHANNELGB)%></td>


			<td align="center"><%=db_CUSTNAME%></td>
			<td align="center"><%=db_getCateNameA_(db_SOSOKGB)%><br><%=db_getCateNameB_(db_SOSOKGB,db_SOSOKETCGB)%></td>
			<% if db_LEVEL1 = "A" then %>
			<td align="center"><%=db_getCodeName("C05",db_LEVEL1)%><br><%=db_getCodeName("C06",db_LEVEL2)%></td>
			<% elseif db_LEVEL1 = "B" then %>
			<td align="center"><%=db_getCodeName("C05",db_LEVEL1)%><br><%=db_getCodeName("C07",db_LEVEL2)%></td>
			<% else %>
			<td align="center"><%=db_getCodeName("C05",db_LEVEL1)%></td>
			<% end if %>

			<% if db_CONSULTETCGB <> "" then %>
				<td align="center"><%=db_getCodeName("C03",db_CONSULTGB)%><br><%=db_getCodeName("C31",db_CONSULTETCGB)%></td> 
			<% else %>
				<td align="center"><%=db_getCodeName("C03",db_CONSULTGB)%></td> 
			<% end if %>
			<td align="center"><%=db_getCodeName("C02",db_REQUESTERGB)%></td>
			<td align="center"><%=db_getCodeName("C10",db_CALLFLAG)%></td>
			<td align="center"><%=db_getCodeName("C08",db_CALLKIND)%></td>
			<td align="center"><%=db_getUserName(db_INCODE)%></td>
			<td align="center"><%=db_getCodeName("C09",db_PROCESSGB)%></td>
		</tr>
<%
		RS.MOVENEXT
	LOOP

%>
</table>

<!-- #include virtual="/Include/Bottom.asp" -->
