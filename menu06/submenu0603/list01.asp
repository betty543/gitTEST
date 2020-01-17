<!-- #include virtual="/Include/Top.asp" -->
<%
	'####### 파라미터 ##################################################################################
	QueryYN = request("QueryYN")
	FromDate = request("FromDate")
	ToDate = request("ToDate")
	whereCD3 = Trim(request("whereCD3"))
	whereCD7 = Trim(request("whereCD7"))

	If QueryYN = "" Then
		whereCD3 = "1"
	End if


	if FromDate = "" then FromDate =left(Date(),7)&"-01" end If
	if ToDate = "" then ToDate=date() end If

	pageWHERE = "QueryYN="&QueryYN&"&FromDate="&FromDate&"&ToDate="&ToDate&"&whereCD3="&whereCD3&"&whereCD7="&whereCD7

%>

<div id='CalendarLayer' style='display:none;'><iframe name='CalendarFrame' src='/include/Calendar.asp' width='172' height='177' border='0' frameborder='0' scrolling='no'></iframe></div>

<script>

	function fn_Search() {

		document.inUpFrm.QueryYN.value = "Y";
		document.inUpFrm.submit();
	}
	
	function fn_Xls() {
		location.href="list01_Xls.asp?<%=pageWHERE%>"
	}
</script>
<table border="0" width="940" cellpadding="0" cellspacing="3" bgcolor="#EFECE5" align="center">
	<tr bgcolor="#FFFFFF">
		<td>
			<form name="inUpFrm" method="post" action="<%=Menu_2nd%>" onsubmit="return fn_Search(this);" style="margin:0">
			<input type="hidden" name="QueryYN" value="<%=QueryYN%>">
			<table width="100%" border="0" cellspacing="1" cellpadding="0" style="border:#E1DED6 solid 1px">
			    <tr>
			        <td class="TDCont">조회기간 :
			        	<input value="<%=FromDate%>" name="FromDate" readonly type="text" size="10" onfocus="setFocusColor(this); new CalendarFrame.Calendar(this);" onblur="setOutColor(this);">
				    	~
				    	<input value="<%=ToDate%>" name="ToDate" readonly type="text" size="10" onfocus="setFocusColor(this); new CalendarFrame.Calendar(this);" onblur="setOutColor(this);">
			        </td>


			        <td class="TDR5px">
			        	<img src="/Images/Btn/BtnSearch.gif" align="absmiddle" style="cursor:hand;" onClick="fn_Search();">
			        	<img src="/Images/Btn/BtnExcel.gif" align="absmiddle" style="cursor:hand;" onClick="fn_Xls();">
			        </td>
			    </tr>
			</table>
			</form>
		</td>
	</tr>
</table>
<table border="0" width="100%" cellpadding="0" cellspacing="0" align="center"><tr height="5"><td></td></tr></table>
<%

	If QueryYN = "Y" Then

%>

<table border="0" cellpadding="0" cellspacing="0" align="center">
	<tr>
		<td>
			<!--<DIV style="OVERFLOW-Y:auto; OVERFLOW-X:auto; MARGIN: 0px 0px 0px 0px; WIDTH:940; HEIGHT:500;">-->
			<table width="940"  border="0" cellpadding="1" cellspacing="1" bgcolor="#EFECE5" align="center">
			    <tr height="30">
					<td align="left" bgcolor="#FFFFFF" class="TDCont" colspan="8">&nbsp;<img src="/Images/dot_01.gif" >&nbsp;<b><font color="#ff00ff"></font> 상담방법</b></td>
				</tr>
				<tr bgcolor='#EEF6FF'>
					<td align='center' class='TDCont'  width='150'>구분</td>
					<td align='center' class='TDCont' >상담</td>
					<td align='center' class='TDCont' >문의</td>
					<td align='center' class='TDCont' >침묵</td>
					<td align='center' class='TDCont' >인터넷</td>
					<td align='center' class='TDCont' >인트라넷</td>
					<td align='center' class='TDCont' >대면</td>
					<td align='center' class='TDCont' >총계</td>
				</tr>

<%
	'상담방법별


	SQL = "select * from ( SELECT	'1' gubun, incode, count(incode) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and ACLASS = 'A' group by incode" '상담,
	SQL = SQL & "	union all SELECT	'2' gubun, incode, count(incode) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and ACLASS = 'B' group by incode" '문의,

	SQL = SQL & "	union all SELECT	'3' gubun, incode, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and ACLASS = 'C' group by incode" '침묵
	SQL = SQL & "	union all SELECT	'4' gubun, incode, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and ACLASS = 'D' and CHANNELGB = 'C' group by incode" '사이버

	SQL = SQL & "	union all SELECT	'5' gubun, incode, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and ACLASS = 'D' and CHANNELGB = 'D' group by incode" '사이버
	'response.write SQL
	SQL = SQL & "	union all SELECT	'6' gubun, incode, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and ACLASS = 'E' group by incode" '대면

	SQL = SQL & "	union all SELECT	'7' gubun, incode, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and ACLASS is not null and ACLASS <> '' group by incode) a order by incode, gubun"

	set Rs = db.execute(SQL)

	do until rs.eof

		incode = rs("incode")
		do until incode <> rs("incode")
			if rs("gubun") = "1" then
				tot1 = rs("cnt")
			elseif rs("gubun") = "2" then
				tot2 = rs("cnt")
			elseif rs("gubun") = "3" then
				tot3 = rs("cnt")
			elseif rs("gubun") = "4" then
				tot4 = rs("cnt")
			elseif rs("gubun") = "5" then
				tot5 = rs("cnt")
			elseif rs("gubun") = "6" then
				tot6 = rs("cnt")
			elseif rs("gubun") = "7" then
				tot7 = rs("cnt")
			end if

			rs.movenext
			if rs.eof then
				exit do
			end if
		loop

%>
		<tr bgcolor='#EEF6FF'>
			<td align='center' class='TDCont'  width='150' ><%=db_getUserName(incode)%></td>
			<td bgcolor='#ffffff' align='center' class='TDCont'><%=tot1%></td>
			<td bgcolor='#ffffff' align='center' class='TDCont'><%=tot2%></td>
			<td bgcolor='#ffffff' align='center' class='TDCont'><%=tot3%></td>
			<td bgcolor='#ffffff' align='center' class='TDCont'><%=tot4%></td>
			<td bgcolor='#ffffff' align='center' class='TDCont'><%=tot5%></td>
			<td bgcolor='#ffffff' align='center' class='TDCont'><%=tot6%></td>
			<td bgcolor='#ffffff' align='center' class='TDCont'><%=tot7%></td>
		</tr>

<%
		tot1 = 0
		tot2 = 0
		tot3 = 0
		tot4 = 0
		tot5 = 0
		tot6 = 0
		tot7 = 0
		if rs.eof then
			exit do
		end if
	loop


	'상담방법별


	SQL = "SELECT	'1' gubun, count(*) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and ACLASS = 'A'" '상담,
	SQL = SQL & "	union all SELECT	'2' gubun, count(*) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and ACLASS = 'B'" '문의,

	SQL = SQL & "	union all SELECT	'3' gubun, count(*) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and ACLASS = 'C'" '침묵
	SQL = SQL & "	union all SELECT	'4' gubun, count(*) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and ACLASS = 'D' and CHANNELGB = 'C'" '사이버

	SQL = SQL & "	union all SELECT	'5' gubun, count(*) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and ACLASS = 'D' and CHANNELGB = 'D'" '사이버
	'response.write SQL
	SQL = SQL & "	union all SELECT	'6' gubun, count(*) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and ACLASS = 'E'" '대면

	SQL = SQL & "	union all SELECT	'7' gubun, count(*) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and ACLASS is not null and ACLASS <> ''"

	set Rs = db.execute(SQL)


	do until rs.eof

		if rs("gubun") = "1" then
			tot1 = rs("cnt")
		elseif rs("gubun") = "2" then
			tot2 = rs("cnt")
		elseif rs("gubun") = "3" then
			tot3 = rs("cnt")
		elseif rs("gubun") = "4" then
			tot4 = rs("cnt")
		elseif rs("gubun") = "5" then
			tot5 = rs("cnt")
		elseif rs("gubun") = "6" then
			tot6 = rs("cnt")
		elseif rs("gubun") = "7" then
			tot7 = rs("cnt")
		end if

		rs.movenext
	loop


%>
				<tr bgcolor='#FFEEF9'>
					<td align='center' class='TDCont'>계</td>
					<td align='center' class='TDCont' ><%=tot1%></td>
					<td  align='center' class='TDCont' ><%=tot2%></td>
					<td  align='center' class='TDCont' ><%=tot3%></td>
					<td  align='center' class='TDCont' ><%=tot4%></td>
					<td  align='center' class='TDCont' ><%=tot5%></td>
					<td  align='center' class='TDCont' ><%=tot6%></td>
					<td  align='center' class='TDCont' ><%=tot7%></td>
				</tr>

				<tr><td colspan="100" height="1" bgcolor="#FFFFFF"></td></tr>
				<%'####### 실제자료가 들어간다. %>
				<%
				'---------------------------------------------------------------------------------------------------------------------
				'데이타 뽑아오기
				'---------------------------------------------------------------------------------------------------------------------

				tot1 = 0
				tot2 = 0
				tot3 = 0
				tot4 = 0
				tot5 = 0
				tot6 = 0
				tot7 = 0

				%>
			</table>
			<!--계급별-->
			<table width="100%" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table  width="940"  border="0" cellpadding="1" cellspacing="1" bgcolor="#EFECE5" align="center">
			    <tr height="30">
					<td align="left" bgcolor="#FFFFFF" class="TDCont" colspan="15">&nbsp;<img src="/Images/dot_01.gif" >&nbsp;<b><font color="#ff00ff"></font> 계급별</b></td>
				</tr>
				<tr bgcolor='#EEF6FF'>
					<td align='center' class='TDCont' width='100' colspan = '2' rowspan='2' nowrap>&nbsp;</td>
					<td align='center' class='TDCont' nowrap colspan='5'>병</td>
					<td align='center' class='TDCont' nowrap colspan='6'>간부</td>
					<td align='center' class='TDCont' nowrap rowspan='2'>기타</td>
					<td align='center' class='TDCont' nowrap rowspan='2'>총계</td>
				</tr>
				<tr bgcolor='#EEF6FF'>
					<td align='center' class='TDCont' nowrap>이병</td>
					<td align='center' class='TDCont' nowrap>일병</td>
					<td align='center' class='TDCont' nowrap>상병</td>
					<td align='center' class='TDCont' nowrap>병장</td>
					<td align='center' class='TDCont' nowrap>미상</td>
					<td align='center' class='TDCont' nowrap>부사관</td>
					<td align='center' class='TDCont' nowrap>위관</td>
					<td align='center' class='TDCont' nowrap>영관</td>
					<td align='center' class='TDCont' nowrap>장군</td>
					<td align='center' class='TDCont' nowrap>병영생활<br>전문상담관</td>
					<td align='center' class='TDCont' nowrap>미상</td>
				</tr>

				<%'####### 실제자료가 들어간다. %>
				<%
				'---------------------------------------------------------------------------------------------------------------------
				'데이타 뽑아오기
				'---------------------------------------------------------------------------------------------------------------------

	'계급별

	SQL = "select * from ( SELECT	'1' gubun, incode, count(incode) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
	SQL = SQL & "	AND		level1 = 'A' and level2 = 'A' group by incode"	'이병
	SQL = SQL & " union all SELECT	'2' gubun, incode, count(incode) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
	SQL = SQL & "	AND		level1 = 'A' and level2 = 'B' group by incode"	'일병
	SQL = SQL & " union all SELECT	'3' gubun, incode, count(incode) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
	SQL = SQL & "	AND		level1 = 'A' and level2 = 'C' group by incode"	'상병
	SQL = SQL & " union all SELECT	'4' gubun, incode, count(incode) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
	SQL = SQL & "	AND		level1 = 'A' and level2 = 'D' group by incode"	'병장
	SQL = SQL & " union all SELECT	'5' gubun, incode, count(incode) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
	SQL = SQL & "	AND		level1 = 'A' and level2 = 'Y' group by incode"	'미상
	SQL = SQL & " union all SELECT	'6' gubun, incode, count(incode) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
	SQL = SQL & "	AND		level1 = 'B' and level2 = 'A' group by incode" '부사관
	SQL = SQL & " union all SELECT	'7' gubun, incode, count(incode) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
	SQL = SQL & "	AND		level1 = 'B' and level2 = 'B' group by incode" '위관
	SQL = SQL & " union all SELECT	'8' gubun, incode, count(incode) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
	SQL = SQL & "	AND		level1 = 'B' and level2 = 'C' group by incode" '영관
	SQL = SQL & " union all SELECT	'9' gubun, incode, count(incode) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
	SQL = SQL & "	AND		level1 = 'B' and level2 = 'D' group by incode" '장군
	SQL = SQL & " union all SELECT	'10' gubun, incode, count(incode) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
	SQL = SQL & "	AND		level1 = 'B' and level2 = 'E' group by incode" '병영생활전문상담관
	SQL = SQL & " union all SELECT	'11' gubun, incode, count(incode) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
	SQL = SQL & "	AND		level1 = 'B' and level2 = 'Y' group by incode" '간부미상
	SQL = SQL & " union all SELECT	'12' gubun, incode, count(incode) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
	SQL = SQL & "	AND		level1 NOT IN ('A','B') and level1 is not null and level1 <> '' group by incode"
	SQL = SQL & " union all SELECT	'13' gubun, incode, count(incode) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and level1 is not null and rtrim(level1) <> '' group by incode ) a order by incode, gubun"

'response.write SQL

	set Rs = db.execute(SQL)

	tot1 = 0
	tot2 = 0
	tot3 = 0
	tot4 = 0
	tot5 = 0
	tot6 = 0
	tot7 = 0
	tot8 = 0
	tot9 = 0
	tot10 = 0
	tot11 = 0
	tot12 = 0
	tot13 = 0

	do until rs.eof
		incode = rs("incode")
		do until incode <> rs("incode")
			if rs("gubun") = "1" then
				tot1 = rs("cnt")
			elseif rs("gubun") = "2" then
				tot2 = rs("cnt")
			elseif rs("gubun") = "3" then
				tot3 = rs("cnt")
			elseif rs("gubun") = "4" then
				tot4 = rs("cnt")
			elseif rs("gubun") = "5" then
				tot5 = rs("cnt")
			elseif rs("gubun") = "6" then
				tot6 = rs("cnt")
			elseif rs("gubun") = "7" then
				tot7 = rs("cnt")
			elseif rs("gubun") = "8" then
				tot8 = rs("cnt")
			elseif rs("gubun") = "9" then
				tot9 = rs("cnt")
			elseif rs("gubun") = "10" then
				tot10 = rs("cnt")
			elseif rs("gubun") = "11" then
				tot11 = rs("cnt")
			elseif rs("gubun") = "12" then
				tot12 = rs("cnt")
			elseif rs("gubun") = "13" then
				tot13 = rs("cnt")
			end if

			rs.movenext

			if rs.eof then
				exit do
			end if

		loop

		'-------------------------------------------------------------------------------------------------
		' 해당직원의 상담, 문의, 사이버 찾아내기
		'-------------------------------------------------------------------------------------------------
		SQL = "select * from ( SELECT	'1' gubun, ACLASS, count(ACLASS) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
		SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
		SQL = SQL & "	AND		level1 = 'A' and level2 = 'A' AND incode = '"&incode&"'	group by ACLASS"	'이병
		SQL = SQL & " union all SELECT	'2' gubun, ACLASS, count(ACLASS) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
		SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
		SQL = SQL & "	AND		level1 = 'A' and level2 = 'B' AND incode = '"&incode&"'	group by ACLASS"	'일병
		SQL = SQL & " union all SELECT	'3' gubun, ACLASS, count(ACLASS) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
		SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
		SQL = SQL & "	AND		level1 = 'A' and level2 = 'C' AND incode = '"&incode&"'	group by ACLASS"	'상병
		SQL = SQL & " union all SELECT	'4' gubun, ACLASS, count(ACLASS) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
		SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
		SQL = SQL & "	AND		level1 = 'A' and level2 = 'D' AND incode = '"&incode&"'	group by ACLASS"	'병장
		SQL = SQL & " union all SELECT	'5' gubun, ACLASS, count(ACLASS) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
		SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
		SQL = SQL & "	AND		level1 = 'A' and level2 = 'Y' AND incode = '"&incode&"'	group by ACLASS"	'미상
		SQL = SQL & " union all SELECT	'6' gubun, ACLASS, count(ACLASS) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
		SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
		SQL = SQL & "	AND		level1 = 'B' and level2 = 'A' AND incode = '"&incode&"'	group by ACLASS" '부사관
		SQL = SQL & " union all SELECT	'7' gubun, ACLASS, count(ACLASS) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
		SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
		SQL = SQL & "	AND		level1 = 'B' and level2 = 'B' AND incode = '"&incode&"'	group by ACLASS" '위관
		SQL = SQL & " union all SELECT	'8' gubun, ACLASS, count(ACLASS) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
		SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
		SQL = SQL & "	AND		level1 = 'B' and level2 = 'C' AND incode = '"&incode&"'	group by ACLASS" '영관
		SQL = SQL & " union all SELECT	'9' gubun, ACLASS, count(ACLASS) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
		SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
		SQL = SQL & "	AND		level1 = 'B' and level2 = 'D' AND incode = '"&incode&"'	group by ACLASS" '장군
		SQL = SQL & " union all SELECT	'10' gubun, ACLASS, count(ACLASS) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
		SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
		SQL = SQL & "	AND		level1 = 'B' and level2 = 'E' AND incode = '"&incode&"'	group by ACLASS" '병영생활전문상담관
		SQL = SQL & " union all SELECT	'11' gubun, ACLASS, count(ACLASS) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
		SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
		SQL = SQL & "	AND		level1 = 'B' and level2 = 'Y' AND incode = '"&incode&"'	group by ACLASS" '간부미상
		SQL = SQL & " union all SELECT	'12' gubun, ACLASS, count(ACLASS) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
		SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
		SQL = SQL & "	AND		level1 NOT IN ('A','B') and level1 is not null and level1 <> '' AND incode = '"&incode&"'	group by ACLASS"
		SQL = SQL & " union all SELECT	'13' gubun, ACLASS, count(ACLASS) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
		SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and level1 is not null and level1 <> '' AND incode = '"&incode&"' group by ACLASS ) a order by ACLASS, gubun"

		set Rs1 = db.execute(SQL)

		tot1_1 = 0
		tot1_2 = 0
		tot1_3 = 0
		tot1_4 = 0
		tot1_5 = 0
		tot1_6 = 0
		tot1_7 = 0
		tot1_8 = 0
		tot1_9 = 0
		tot1_10 = 0
		tot1_11 = 0
		tot1_12 = 0
		tot1_13 = 0

		tot2_1 = 0
		tot2_2 = 0
		tot2_3 = 0
		tot2_4 = 0
		tot2_5 = 0
		tot2_6 = 0
		tot2_7 = 0
		tot2_8 = 0
		tot2_9 = 0
		tot2_10 = 0
		tot2_11 = 0
		tot2_12 = 0
		tot2_13 = 0

		tot3_1 = 0
		tot3_2 = 0
		tot3_3 = 0
		tot3_4 = 0
		tot3_5 = 0
		tot3_6 = 0
		tot3_7 = 0
		tot3_8 = 0
		tot3_9 = 0
		tot3_10 = 0
		tot3_11 = 0
		tot3_12 = 0
		tot3_13 = 0

		do until rs1.eof
			ACLASS = rs1("ACLASS")
			do until ACLASS <> rs1("ACLASS")

				if ACLASS = "A" then
					if rs1("gubun") = "1" then
						tot1_1 = rs1("cnt")
					elseif rs1("gubun") = "2" then
						tot1_2 = rs1("cnt")
					elseif rs1("gubun") = "3" then
						tot1_3 = rs1("cnt")
					elseif rs1("gubun") = "4" then
						tot1_4 = rs1("cnt")
					elseif rs1("gubun") = "5" then
						tot1_5 = rs1("cnt")
					elseif rs1("gubun") = "6" then
						tot1_6 = rs1("cnt")
					elseif rs1("gubun") = "7" then
						tot1_7 = rs1("cnt")
					elseif rs1("gubun") = "8" then
						tot1_8 = rs1("cnt")
					elseif rs1("gubun") = "9" then
						tot1_9 = rs1("cnt")
					elseif rs1("gubun") = "10" then
						tot1_10 = rs1("cnt")
					elseif rs1("gubun") = "11" then
						tot1_11 = rs1("cnt")
					elseif rs1("gubun") = "12" then
						tot1_12 = rs1("cnt")
					elseif rs1("gubun") = "13" then
						tot1_13 = rs1("cnt")
					end if
				elseif ACLASS = "B" then
					if rs1("gubun") = "1" then
						tot2_1 = rs1("cnt")
					elseif rs1("gubun") = "2" then
						tot2_2 = rs1("cnt")
					elseif rs1("gubun") = "3" then
						tot2_3 = rs1("cnt")
					elseif rs1("gubun") = "4" then
						tot2_4 = rs1("cnt")
					elseif rs1("gubun") = "5" then
						tot2_5 = rs1("cnt")
					elseif rs1("gubun") = "6" then
						tot2_6 = rs1("cnt")
					elseif rs1("gubun") = "7" then
						tot2_7 = rs1("cnt")
					elseif rs1("gubun") = "8" then
						tot2_8 = rs1("cnt")
					elseif rs1("gubun") = "9" then
						tot2_9 = rs1("cnt")
					elseif rs1("gubun") = "10" then
						tot2_10 = rs1("cnt")
					elseif rs1("gubun") = "11" then
						tot2_11 = rs1("cnt")
					elseif rs1("gubun") = "12" then
						tot2_12 = rs1("cnt")
					elseif rs1("gubun") = "13" then
						tot2_13 = rs1("cnt")
					end if
				else
					if rs1("gubun") = "1" then
						tot3_1 = rs1("cnt")
					elseif rs1("gubun") = "2" then
						tot3_2 = rs1("cnt")
					elseif rs1("gubun") = "3" then
						tot3_3 = rs1("cnt")
					elseif rs1("gubun") = "4" then
						tot3_4 = rs1("cnt")
					elseif rs1("gubun") = "5" then
						tot3_5 = rs1("cnt")
					elseif rs1("gubun") = "6" then
						tot3_6 = rs1("cnt")
					elseif rs1("gubun") = "7" then
						tot3_7 = rs1("cnt")
					elseif rs1("gubun") = "8" then
						tot3_8 = rs1("cnt")
					elseif rs1("gubun") = "9" then
						tot3_9 = rs1("cnt")
					elseif rs1("gubun") = "10" then
						tot3_10 = rs1("cnt")
					elseif rs1("gubun") = "11" then
						tot3_11 = rs1("cnt")
					elseif rs1("gubun") = "12" then
						tot3_12 = rs1("cnt")
					elseif rs1("gubun") = "13" then
						tot3_13 = rs1("cnt")
					end if
				end if
	
				rs1.movenext
				if rs1.eof then
					exit do
				end if
			loop
			if rs1.eof then
				exit do
			end if
		loop
				%>

				<tr bgcolor='#EEF6FF'>
					<td align='center' class='TDCont' width="100" nowrap rowspan='4'><%=db_getUserName(incode)%></td>
					<td align='center' class='TDCont' width="100" nowrap>상담</td>
					<td bgcolor='#ffffff' align='center' class='TDCont' width="100"><%=tot1_1%></td>
					<td bgcolor='#ffffff' align='center' class='TDCont' width="100"><%=tot1_2%></td>
					<td bgcolor='#ffffff' align='center' class='TDCont' width="100"><%=tot1_3%></td>
					<td bgcolor='#ffffff' align='center' class='TDCont' width="100"><%=tot1_4%></td>
					<td bgcolor='#ffffff' align='center' class='TDCont' width="100"><%=tot1_5%></td>
					<td bgcolor='#ffffff' align='center' class='TDCont' width="100"><%=tot1_6%></td>
					<td bgcolor='#ffffff' align='center' class='TDCont' width="100"><%=tot1_7%></td>
					<td bgcolor='#ffffff' align='center' class='TDCont' width="100"><%=tot1_8%></td>
					<td bgcolor='#ffffff' align='center' class='TDCont' width="100"><%=tot1_9%></td>
					<td bgcolor='#ffffff' align='center' class='TDCont' width="100"><%=tot1_10%></td>
					<td bgcolor='#ffffff' align='center' class='TDCont' width="100"><%=tot1_11%></td>
					<td bgcolor='#ffffff' align='center' class='TDCont' width="100"><%=tot1_12%></td>
					<td bgcolor='#ffffff' align='center' class='TDCont' width="100"><%=tot1_13%></td>
				</tr>
				<tr bgcolor='#EEF6FF'>
					<td align='center' class='TDCont' width="100" nowrap>문의</td>
					<td bgcolor='#ffffff' align='center' class='TDCont' width="100" ><%=tot2_1%></td>
					<td bgcolor='#ffffff' align='center' class='TDCont' width="100"><%=tot2_2%></td>
					<td bgcolor='#ffffff' align='center' class='TDCont' width="100"><%=tot2_3%></td>
					<td bgcolor='#ffffff' align='center' class='TDCont' width="100"><%=tot2_4%></td>
					<td bgcolor='#ffffff' align='center' class='TDCont' width="100"><%=tot2_5%></td>
					<td bgcolor='#ffffff' align='center' class='TDCont'  width="100"><%=tot2_6%></td>
					<td bgcolor='#ffffff' align='center' class='TDCont' width="100"><%=tot2_7%></td>
					<td bgcolor='#ffffff' align='center' class='TDCont' width="100"><%=tot2_8%></td>
					<td bgcolor='#ffffff' align='center' class='TDCont'width="100" ><%=tot2_9%></td>
					<td bgcolor='#ffffff' align='center' class='TDCont'width="100" ><%=tot2_10%></td>
					<td bgcolor='#ffffff' align='center' class='TDCont' width="100"><%=tot2_11%></td>
					<td bgcolor='#ffffff' align='center' class='TDCont' width="100"><%=tot2_12%></td>
					<td bgcolor='#ffffff' align='center' class='TDCont' width="100"><%=tot2_13%></td>
				</tr>
				<tr bgcolor='#EEF6FF'>
					<td align='center' class='TDCont' width="100" nowrap>사이버</td>
					<td bgcolor='#ffffff' align='center' class='TDCont' width="100" ><%=tot3_1%></td>
					<td bgcolor='#ffffff' align='center' class='TDCont' width="100"><%=tot3_2%></td>
					<td bgcolor='#ffffff' align='center' class='TDCont' width="100"><%=tot3_3%></td>
					<td bgcolor='#ffffff' align='center' class='TDCont' width="100"><%=tot3_4%></td>
					<td bgcolor='#ffffff' align='center' class='TDCont' width="100"><%=tot3_5%></td>
					<td bgcolor='#ffffff' align='center' class='TDCont'  width="100"><%=tot3_6%></td>
					<td bgcolor='#ffffff' align='center' class='TDCont' width="100"><%=tot3_7%></td>
					<td bgcolor='#ffffff' align='center' class='TDCont' width="100"><%=tot3_8%></td>
					<td bgcolor='#ffffff' align='center' class='TDCont'width="100" ><%=tot3_9%></td>
					<td bgcolor='#ffffff' align='center' class='TDCont'width="100" ><%=tot3_10%></td>
					<td bgcolor='#ffffff' align='center' class='TDCont' width="100"><%=tot3_11%></td>
					<td bgcolor='#ffffff' align='center' class='TDCont' width="100"><%=tot3_12%></td>
					<td bgcolor='#ffffff' align='center' class='TDCont' width="100"><%=tot3_13%></td>
				</tr>
				<tr bgcolor='#EEF6FF'>
					<td align='center' class='TDCont' width="100" nowrap>계</td>
					<td align='center' class='TDCont' width="100" ><%=tot1%></td>
					<td align='center' class='TDCont' width="100"><%=tot2%></td>
					<td align='center' class='TDCont' width="100"><%=tot3%></td>
					<td align='center' class='TDCont' width="100"><%=tot4%></td>
					<td align='center' class='TDCont' width="100"><%=tot5%></td>
					<td align='center' class='TDCont'  width="100"><%=tot6%></td>
					<td align='center' class='TDCont' width="100"><%=tot7%></td>
					<td align='center' class='TDCont' width="100"><%=tot8%></td>
					<td align='center' class='TDCont'width="100" ><%=tot9%></td>
					<td align='center' class='TDCont'width="100" ><%=tot10%></td>
					<td align='center' class='TDCont' width="100"><%=tot11%></td>
					<td align='center' class='TDCont' width="100"><%=tot12%></td>
					<td align='center' class='TDCont' width="100"><%=tot13%></td>
				</tr>


<%
				tot1 = 0
				tot2 = 0
				tot3 = 0
				tot4 = 0
				tot5 = 0
				tot6 = 0
				tot7 = 0
				tot8 = 0
				tot9 = 0
				tot10 = 0
				tot11 = 0
				tot12 = 0
				tot13 = 0

			if rs.eof then
				exit do
			end if

	loop


	SQL = "SELECT	'1' gubun, count(*) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
	SQL = SQL & "	AND		level1 = 'A' and level2 = 'A'"	'이병
	SQL = SQL & " union all SELECT	'2' gubun, count(*) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
	SQL = SQL & "	AND		level1 = 'A' and level2 = 'B'"	'일병
	SQL = SQL & " union all SELECT	'3' gubun, count(*) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
	SQL = SQL & "	AND		level1 = 'A' and level2 = 'C'"	'상병
	SQL = SQL & " union all SELECT	'4' gubun, count(*) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
	SQL = SQL & "	AND		level1 = 'A' and level2 = 'D'"	'병장
	SQL = SQL & " union all SELECT	'5' gubun, count(*) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
	SQL = SQL & "	AND		level1 = 'A' and level2 = 'Y'"	'미상
	SQL = SQL & " union all SELECT	'6' gubun, count(*) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
	SQL = SQL & "	AND		level1 = 'B' and level2 = 'A'" '부사관
	SQL = SQL & " union all SELECT	'7' gubun, count(*) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
	SQL = SQL & "	AND		level1 = 'B' and level2 = 'B'" '위관
	SQL = SQL & " union all SELECT	'8' gubun, count(*) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
	SQL = SQL & "	AND		level1 = 'B' and level2 = 'C'" '영관
	SQL = SQL & " union all SELECT	'9' gubun, count(*) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
	SQL = SQL & "	AND		level1 = 'B' and level2 = 'D'" '장군
	SQL = SQL & " union all SELECT	'10' gubun, count(*) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
	SQL = SQL & "	AND		level1 = 'B' and level2 = 'E'" '병영생활전문상담관
	SQL = SQL & " union all SELECT	'11' gubun, count(*) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
	SQL = SQL & "	AND		level1 = 'B' and level2 = 'Y'" '간부미상
	SQL = SQL & " union all SELECT	'12' gubun, count(*) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
	SQL = SQL & "	AND		level1 NOT IN ('A','B') and level1 is not null and level1 <> '' "
	SQL = SQL & " union all SELECT	'13' gubun, count(*) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and level1 is not null and level1 <> ''"

	set Rs = db.execute(SQL)

	do until rs.eof

		if rs("gubun") = "1" then
			tot1 = rs("cnt")
		elseif rs("gubun") = "2" then
			tot2 = rs("cnt")
		elseif rs("gubun") = "3" then
			tot3 = rs("cnt")
		elseif rs("gubun") = "4" then
			tot4 = rs("cnt")
		elseif rs("gubun") = "5" then
			tot5 = rs("cnt")
		elseif rs("gubun") = "6" then
			tot6 = rs("cnt")
		elseif rs("gubun") = "7" then
			tot7 = rs("cnt")
		elseif rs("gubun") = "8" then
			tot8 = rs("cnt")
		elseif rs("gubun") = "9" then
			tot9 = rs("cnt")
		elseif rs("gubun") = "10" then
			tot10 = rs("cnt")
		elseif rs("gubun") = "11" then
			tot11 = rs("cnt")
		elseif rs("gubun") = "12" then
			tot12 = rs("cnt")
		elseif rs("gubun") = "13" then
			tot13 = rs("cnt")

		end if

		rs.movenext
	loop



				%>

				<tr bgcolor="#FFEEF9">
					<td align='center' class='TDCont' colspan='2'>총계</td>
					<td align='center' class='TDCont' ><%=tot1%></td>
					<td align='center' class='TDCont' ><%=tot2%></td>
					<td align='center' class='TDCont' ><%=tot3%></td>
					<td align='center' class='TDCont' ><%=tot4%></td>
					<td align='center' class='TDCont' ><%=tot5%></td>
					<td align='center' class='TDCont' ><%=tot6%></td>
					<td align='center' class='TDCont' ><%=tot7%></td>
					<td align='center' class='TDCont' ><%=tot8%></td>
					<td align='center' class='TDCont' ><%=tot9%></td>
					<td align='center' class='TDCont' ><%=tot10%></td>
					<td align='center' class='TDCont' ><%=tot11%></td>
					<td align='center' class='TDCont' ><%=tot12%></td>
					<td align='center' class='TDCont' ><%=tot13%></td>
				</tr>
			</table>

<%
				tot1 = 0
				tot2 = 0
				tot3 = 0
				tot4 = 0
				tot5 = 0
				tot6 = 0
				tot7 = 0
				tot8 = 0
				tot9 = 0
				tot10 = 0
				tot11 = 0
				tot12 = 0
				tot13 = 0


%>
			<table width="940" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table  width="940" border="0" cellpadding="1" cellspacing="1" bgcolor="#EFECE5" align="center">
			    <tr height="30">
					<td align="left" bgcolor="#FFFFFF" class="TDCont" colspan="13">&nbsp;<img src="/Images/dot_01.gif" >&nbsp;<b><font color="#ff00ff"></font> 부대별</b></td>
				</tr>
				<tr bgcolor='#EEF6FF'>
					<td align='center' class='TDCont'  width='150'>구분</td>
					<td align='center' class='TDCont' width='150'>1군</td>
					<td align='center' class='TDCont' width='150'>2작사</td>
					<td align='center' class='TDCont' width='150'>3군</td>
					<td align='center' class='TDCont' width='150'>육직</td>
					<td align='center' class='TDCont' width='150'>군수사</td>
					<td align='center' class='TDCont'width='150' >교육사</td>
					<td align='center' class='TDCont'width='150'>특전사</td>
					<td align='center' class='TDCont'width='150'>타부대</td>
					<td align='center' class='TDCont'width='150'>기타</td>
					<td align='center' class='TDCont'width='150'>미상</td>
					<td align='center' class='TDCont'  width='150'>총계</td>
				</tr>
<%

	SQL = "select * from ( SELECT	'1' gubun, incode, count(incode) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and SOSOKGB = 'A' and ACLASS in ('A','B','C') group by incode" '1군
	SQL = SQL & "	union all SELECT	'2' gubun, incode, count(incode) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and SOSOKGB = 'B' and ACLASS in ('A','B','C') group by incode" '2작사
	SQL = SQL & "	union all SELECT	'3' gubun, incode, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and SOSOKGB = 'C' and ACLASS in ('A','B','C') group by incode" '3군
	SQL = SQL & "	union all SELECT	'4' gubun, incode, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and SOSOKGB = 'D' and ACLASS in ('A','B','C') group by incode" '육직
	SQL = SQL & "	union all SELECT	'5' gubun, incode, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and SOSOKGB = 'E' and ACLASS in ('A','B','C') group by incode" '군수사	'response.write SQL
	SQL = SQL & "	union all SELECT	'6' gubun, incode, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and SOSOKGB = 'F' and ACLASS in ('A','B','C') group by incode" '교육사
	SQL = SQL & "	union all SELECT	'7' gubun, incode, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and SOSOKGB = 'G' and ACLASS in ('A','B','C') group by incode" '특전사
	SQL = SQL & "	union all SELECT	'8' gubun, incode, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and SOSOKGB = 'I' and ACLASS in ('A','B','C') group by incode" '타부대
	SQL = SQL & "	union all SELECT	'9' gubun, incode, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and SOSOKGB = 'H' and ACLASS in ('A','B','C') group by incode " '기타
	SQL = SQL & "	union all SELECT	'10' gubun, incode, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and SOSOKGB NOT IN ('A','B','C','D','E','F','G','H','I') and ACLASS in ('A','B','C') group by incode " '미상
	SQL = SQL & "	union all SELECT	'11' gubun, incode, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and ACLASS in ('A','B','C') group by incode ) a order by incode, gubun" '총계

	set Rs = db.execute(SQL)

	do until rs.eof

		incode = rs("incode")
		do until incode <> rs("incode")
			if rs("gubun") = "1" then
				tot1 = rs("cnt")
			elseif rs("gubun") = "2" then
				tot2 = rs("cnt")
			elseif rs("gubun") = "3" then
				tot3 = rs("cnt")
			elseif rs("gubun") = "4" then
				tot4 = rs("cnt")
			elseif rs("gubun") = "5" then
				tot5 = rs("cnt")
			elseif rs("gubun") = "6" then
				tot6 = rs("cnt")
			elseif rs("gubun") = "7" then
				tot7 = rs("cnt")
			elseif rs("gubun") = "8" then
				tot8 = rs("cnt")
			elseif rs("gubun") = "9" then
				tot9 = rs("cnt")
			elseif rs("gubun") = "10" then
				tot10 = rs("cnt")
			elseif rs("gubun") = "11" then
				tot11 = rs("cnt")
			end if

			rs.movenext
			if rs.eof then
				exit do
			end if
		loop

%>
		<tr bgcolor='#EEF6FF'>
			<td align='center' class='TDCont'  width='150' ><%=db_getUserName(incode)%></td>
			<td bgcolor='#ffffff' align='center' class='TDCont'><%=tot1%></td>
			<td bgcolor='#ffffff' align='center' class='TDCont'><%=tot2%></td>
			<td bgcolor='#ffffff' align='center' class='TDCont'><%=tot3%></td>
			<td bgcolor='#ffffff' align='center' class='TDCont'><%=tot4%></td>
			<td bgcolor='#ffffff' align='center' class='TDCont'><%=tot5%></td>
			<td bgcolor='#ffffff' align='center' class='TDCont'><%=tot6%></td>
			<td bgcolor='#ffffff' align='center' class='TDCont'><%=tot7%></td>
			<td bgcolor='#ffffff' align='center' class='TDCont'><%=tot8%></td>
			<td bgcolor='#ffffff' align='center' class='TDCont'><%=tot9%></td>
			<td bgcolor='#ffffff' align='center' class='TDCont'><%=tot10%></td>
			<td bgcolor='#ffffff' align='center' class='TDCont'><%=tot11%></td>
		</tr>

<%
		tot1 = 0
		tot2 = 0
		tot3 = 0
		tot4 = 0
		tot5 = 0
		tot6 = 0
		tot7 = 0

		tot8 = 0
		tot9 = 0
		tot10 = 0
		tot11 = 0
		tot12 = 0
		if rs.eof then
			exit do
		end if
	loop

	'부대별
	SQL = "select * from ( SELECT	'1' gubun, count(incode) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and SOSOKGB = 'A' and ACLASS in ('A','B','C')" '1군
	SQL = SQL & "	union all SELECT	'2' gubun, count(incode) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and SOSOKGB = 'B' and ACLASS in ('A','B','C')" '2작사
	SQL = SQL & "	union all SELECT	'3' gubun, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and SOSOKGB = 'C' and ACLASS in ('A','B','C')" '3군
	SQL = SQL & "	union all SELECT	'4' gubun, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and SOSOKGB = 'D' and ACLASS in ('A','B','C')" '육직
	SQL = SQL & "	union all SELECT	'5' gubun, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and SOSOKGB = 'E' and ACLASS in ('A','B','C')" '군수사	'response.write SQL
	SQL = SQL & "	union all SELECT	'6' gubun, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and SOSOKGB = 'F' and ACLASS in ('A','B','C')" '교육사
	SQL = SQL & "	union all SELECT	'7' gubun, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and SOSOKGB = 'G' and ACLASS in ('A','B','C')" '특전사
	SQL = SQL & "	union all SELECT	'8' gubun, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and SOSOKGB = 'I' and ACLASS in ('A','B','C')" '타부대
	SQL = SQL & "	union all SELECT	'9' gubun, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and SOSOKGB = 'H' and ACLASS in ('A','B','C')" '기타
	SQL = SQL & "	union all SELECT	'10' gubun, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and SOSOKGB NOT IN ('A','B','C','D','E','F','G','H','I') and ACLASS in ('A','B','C')" '미상
	SQL = SQL & "	union all SELECT	'11' gubun, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and ACLASS in ('A','B','C')) a order by gubun" '총계

	set Rs = db.execute(SQL)

	do until rs.eof

		if rs("gubun") = "1" then
			tot1 = rs("cnt")
		elseif rs("gubun") = "2" then
			tot2 = rs("cnt")
		elseif rs("gubun") = "3" then
			tot3 = rs("cnt")
		elseif rs("gubun") = "4" then
			tot4 = rs("cnt")
		elseif rs("gubun") = "5" then
			tot5 = rs("cnt")
		elseif rs("gubun") = "6" then
			tot6 = rs("cnt")
		elseif rs("gubun") = "7" then
			tot7 = rs("cnt")
		elseif rs("gubun") = "8" then
			tot8 = rs("cnt")
		elseif rs("gubun") = "9" then
			tot9 = rs("cnt")
		elseif rs("gubun") = "10" then
			tot10 = rs("cnt")
		elseif rs("gubun") = "11" then
			tot11 = rs("cnt")

		end if

		rs.movenext
		if rs.eof then
			exit do
		end if
	loop

%>
				<tr bgcolor="#FFEEF9">
					<td align='center' class='TDCont'  width='150' >계</td>
					<td align='center' class='TDCont'><%=tot1%></td>
					<td align='center' class='TDCont'><%=tot2%></td>
					<td align='center' class='TDCont'><%=tot3%></td>
					<td align='center' class='TDCont'><%=tot4%></td>
					<td align='center' class='TDCont'><%=tot5%></td>
					<td align='center' class='TDCont'><%=tot6%></td>
					<td align='center' class='TDCont'><%=tot7%></td>
					<td align='center' class='TDCont'><%=tot8%></td>
					<td align='center' class='TDCont'><%=tot9%></td>
					<td align='center' class='TDCont'><%=tot10%></td>
					<td align='center' class='TDCont'><%=tot11%></td>
				</tr>
			</table>

			<table width="940" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table  width="940" border="0" cellpadding="1" cellspacing="1" bgcolor="#EFECE5" align="center">
			    <tr height="30">
					<td align="left" bgcolor="#FFFFFF" class="TDCont" colspan="16">&nbsp;<img src="/Images/dot_01.gif" >&nbsp;<b><font color="#ff00ff"></font> 상담분야별</b></td>
				</tr>
				<tr bgcolor='#EEF6FF'>
					<td align='center' class='TDCont' colspan='2' width='200' nowrap></td>
					<td align='center' class='TDCont' width='150'>자살<br>문제</td>
					<td align='center' class='TDCont' width='150'>부적응<br>보직</td>
					<td align='center' class='TDCont' width='150'>인권/<br>기본권<br>침해</td>
					<td align='center' class='TDCont' width='150'>이성<br>문제</td>
					<td align='center' class='TDCont' width='150' >건강<br>문제</td>
					<td align='center' class='TDCont' width='150' >전역<br>문제</td>
					<td align='center' class='TDCont' width='150' >가정<br>문제</td>
					<td align='center' class='TDCont' width='150' >전출<br>문제</td>
					<td align='center' class='TDCont' width='150' >대체<br>복무</td>
					<td align='center' class='TDCont' width='150' >조치<br>(감사)</td>
					<td align='center' class='TDCont' width='150' >기타</td>
					<td align='center' class='TDCont' width='150'>총계</td>
				</tr>

<%

	'상담관별 총계
	SQL = "select * from ( SELECT	'1' gubun, incode, count(incode) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and CONSULTGB = 'A' and ACLASS in ('A','D') group by incode" '자살
	SQL = SQL & "	union all SELECT	'2' gubun, incode, count(incode) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and CONSULTGB = 'B' and ACLASS in ('A','D') group by incode" '부적응
	SQL = SQL & "	union all SELECT	'3' gubun, incode, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and CONSULTGB = 'C' and ACLASS in ('A','D') group by incode" '인권
	SQL = SQL & "	union all SELECT	'4' gubun, incode, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and CONSULTGB = 'D' and ACLASS in ('A','D') group by incode" '이성
	SQL = SQL & "	union all SELECT	'5' gubun, incode, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and CONSULTGB = 'F' and ACLASS in ('A','D') group by incode" '건강	'response.write SQL
	SQL = SQL & "	union all SELECT	'6' gubun, incode, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and CONSULTGB = 'G' and ACLASS in ('A','D') group by incode" '전역
	SQL = SQL & "	union all SELECT	'7' gubun, incode, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and CONSULTGB = 'H' and ACLASS in ('A','D') group by incode" '가정
	SQL = SQL & "	union all SELECT	'8' gubun, incode, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and CONSULTGB = 'I' and ACLASS in ('A','D') group by incode" '전출
	SQL = SQL & "	union all SELECT	'9' gubun, incode, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and CONSULTGB = 'J' and ACLASS in ('A','D') group by incode " '대체
	SQL = SQL & "	union all SELECT	'10' gubun, incode, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and CONSULTGB = 'K' and ACLASS in ('A','D') group by incode" '조치
	SQL = SQL & "	union all SELECT	'11' gubun, incode, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and CONSULTGB = 'Z' and ACLASS in ('A','D') group by incode" '기타
	SQL = SQL & "	union all SELECT	'12' gubun, incode, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and ACLASS in ('A','D') group by incode ) a order by incode, gubun" '총계

	set Rs = db.execute(SQL)

		tot1 = 0
		tot2 = 0
		tot3 = 0
		tot4 = 0
		tot5 = 0
		tot6 = 0
		tot7 = 0

		tot8 = 0
		tot9 = 0
		tot10 = 0
		tot11 = 0
		tot12 = 0

	do until rs.eof

		incode = rs("incode")
		do until incode <> rs("incode")
			if rs("gubun") = "1" then
				tot1 = rs("cnt")
			elseif rs("gubun") = "2" then
				tot2 = rs("cnt")
			elseif rs("gubun") = "3" then
				tot3 = rs("cnt")
			elseif rs("gubun") = "4" then
				tot4 = rs("cnt")
			elseif rs("gubun") = "5" then
				tot5 = rs("cnt")
			elseif rs("gubun") = "6" then
				tot6 = rs("cnt")
			elseif rs("gubun") = "7" then
				tot7 = rs("cnt")
			elseif rs("gubun") = "8" then
				tot8 = rs("cnt")
			elseif rs("gubun") = "9" then
				tot9 = rs("cnt")
			elseif rs("gubun") = "10" then
				tot10 = rs("cnt")
			elseif rs("gubun") = "11" then
				tot11 = rs("cnt")
			elseif rs("gubun") = "12" then
				tot12 = rs("cnt")
			end if

			rs.movenext
			if rs.eof then
				exit do
			end if
		loop


		'-------------------------------------------------------------------------------------------------
		' 해당직원의 상담, 문의, 사이버 찾아내기
		'-------------------------------------------------------------------------------------------------
		SQL = "select * from ( SELECT	'1' gubun, ACLASS, count(incode) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
		SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and CONSULTGB = 'A' and ACLASS in ('A','D') and incode ='" & incode&"' group by ACLASS" '자살
		SQL = SQL & "	union all SELECT	'2' gubun, ACLASS, count(ACLASS) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
		SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and CONSULTGB = 'B' and ACLASS in ('A','D') and incode ='" & incode&"' group by ACLASS" '부적응
		SQL = SQL & "	union all SELECT	'3' gubun, ACLASS, count(ACLASS) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
		SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and CONSULTGB = 'C' and ACLASS in ('A','D') and incode ='" & incode&"' group by ACLASS" '인권
		SQL = SQL & "	union all SELECT	'4' gubun, ACLASS, count(ACLASS) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
		SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and CONSULTGB = 'D' and ACLASS in ('A','D') and incode ='" & incode&"' group by ACLASS" '이성
		SQL = SQL & "	union all SELECT	'5' gubun, ACLASS, count(ACLASS) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
		SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and CONSULTGB = 'F' and ACLASS in ('A','D') and incode ='" & incode&"' group by ACLASS" '건강	'response.write SQL
		SQL = SQL & "	union all SELECT	'6' gubun, ACLASS, count(ACLASS) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
		SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and CONSULTGB = 'G' and ACLASS in ('A','D') and incode ='" & incode&"' group by ACLASS" '전역
		SQL = SQL & "	union all SELECT	'7' gubun, ACLASS, count(ACLASS) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
		SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and CONSULTGB = 'H' and ACLASS in ('A','D') and incode ='" & incode&"' group by ACLASS" '가정
		SQL = SQL & "	union all SELECT	'8' gubun, ACLASS, count(ACLASS) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
		SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and CONSULTGB = 'I' and ACLASS in ('A','D') and incode ='" & incode&"' group by ACLASS" '전출
		SQL = SQL & "	union all SELECT	'9' gubun, ACLASS, count(ACLASS) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
		SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and CONSULTGB = 'J' and ACLASS in ('A','D') and incode ='" & incode&"' group by ACLASS" '대체
		SQL = SQL & "	union all SELECT	'10' gubun, ACLASS, count(ACLASS) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
		SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and CONSULTGB = 'K' and ACLASS in ('A','D') and incode ='" & incode&"' group by ACLASS" '조치
		SQL = SQL & "	union all SELECT	'11' gubun, ACLASS, count(ACLASS) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
		SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and CONSULTGB = 'Z' and ACLASS in ('A','D') and incode ='" & incode&"' group by ACLASS" '기타
		SQL = SQL & "	union all SELECT	'12' gubun, ACLASS, count(ACLASS) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
		SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and ACLASS in ('A','D') and incode ='" & incode&"' group by ACLASS ) a order by ACLASS, gubun" '총계

		set Rs1 = db.execute(SQL)

		tot1_1 = 0
		tot1_2 = 0
		tot1_3 = 0
		tot1_4 = 0
		tot1_5 = 0
		tot1_6 = 0
		tot1_7 = 0
		tot1_8 = 0
		tot1_9 = 0
		tot1_10 = 0
		tot1_11 = 0
		tot1_12 = 0
		tot1_13 = 0

		tot2_1 = 0
		tot2_2 = 0
		tot2_3 = 0
		tot2_4 = 0
		tot2_5 = 0
		tot2_6 = 0
		tot2_7 = 0
		tot2_8 = 0
		tot2_9 = 0
		tot2_10 = 0
		tot2_11 = 0
		tot2_12 = 0
		tot2_13 = 0

		tot3_1 = 0
		tot3_2 = 0
		tot3_3 = 0
		tot3_4 = 0
		tot3_5 = 0
		tot3_6 = 0
		tot3_7 = 0
		tot3_8 = 0
		tot3_9 = 0
		tot3_10 = 0
		tot3_11 = 0
		tot3_12 = 0
		tot3_13 = 0

		do until rs1.eof
			ACLASS = rs1("ACLASS")
			do until ACLASS <> rs1("ACLASS")

				if ACLASS = "A" then
					if rs1("gubun") = "1" then
						tot1_1 = rs1("cnt")
					elseif rs1("gubun") = "2" then
						tot1_2 = rs1("cnt")
					elseif rs1("gubun") = "3" then
						tot1_3 = rs1("cnt")
					elseif rs1("gubun") = "4" then
						tot1_4 = rs1("cnt")
					elseif rs1("gubun") = "5" then
						tot1_5 = rs1("cnt")
					elseif rs1("gubun") = "6" then
						tot1_6 = rs1("cnt")
					elseif rs1("gubun") = "7" then
						tot1_7 = rs1("cnt")
					elseif rs1("gubun") = "8" then
						tot1_8 = rs1("cnt")
					elseif rs1("gubun") = "9" then
						tot1_9 = rs1("cnt")
					elseif rs1("gubun") = "10" then
						tot1_10 = rs1("cnt")
					elseif rs1("gubun") = "11" then
						tot1_11 = rs1("cnt")
					elseif rs1("gubun") = "12" then
						tot1_12 = rs1("cnt")
					elseif rs1("gubun") = "13" then
						tot1_13 = rs1("cnt")
					end if
				else
					if rs1("gubun") = "1" then
						tot2_1 = rs1("cnt")
					elseif rs1("gubun") = "2" then
						tot2_2 = rs1("cnt")
					elseif rs1("gubun") = "3" then
						tot2_3 = rs1("cnt")
					elseif rs1("gubun") = "4" then
						tot2_4 = rs1("cnt")
					elseif rs1("gubun") = "5" then
						tot2_5 = rs1("cnt")
					elseif rs1("gubun") = "6" then
						tot2_6 = rs1("cnt")
					elseif rs1("gubun") = "7" then
						tot2_7 = rs1("cnt")
					elseif rs1("gubun") = "8" then
						tot2_8 = rs1("cnt")
					elseif rs1("gubun") = "9" then
						tot2_9 = rs1("cnt")
					elseif rs1("gubun") = "10" then
						tot2_10 = rs1("cnt")
					elseif rs1("gubun") = "11" then
						tot2_11 = rs1("cnt")
					elseif rs1("gubun") = "12" then
						tot2_12 = rs1("cnt")
					elseif rs1("gubun") = "13" then
						tot2_13 = rs1("cnt")
					end if

				end if
	
				rs1.movenext
				if rs1.eof then
					exit do
				end if
			loop
			if rs1.eof then
				exit do
			end if
		loop


%>
		<tr bgcolor='#EEF6FF'>
			<td align='center' class='TDCont'  width='150' rowspan='3' ><%=db_getUserName(incode)%></td>
			<td align='center' class='TDCont'  width='150' >상담</td>
			<td bgcolor='#ffffff' align='center' class='TDCont'><%=tot1_1%></td>
			<td bgcolor='#ffffff' align='center' class='TDCont'><%=tot1_2%></td>
			<td bgcolor='#ffffff' align='center' class='TDCont'><%=tot1_3%></td>
			<td bgcolor='#ffffff' align='center' class='TDCont'><%=tot1_4%></td>
			<td bgcolor='#ffffff' align='center' class='TDCont'><%=tot1_5%></td>
			<td bgcolor='#ffffff' align='center' class='TDCont'><%=tot1_6%></td>
			<td bgcolor='#ffffff' align='center' class='TDCont'><%=tot1_7%></td>
			<td bgcolor='#ffffff' align='center' class='TDCont'><%=tot1_8%></td>
			<td bgcolor='#ffffff' align='center' class='TDCont'><%=tot1_9%></td>
			<td bgcolor='#ffffff' align='center' class='TDCont'><%=tot1_10%></td>
			<td bgcolor='#ffffff' align='center' class='TDCont'><%=tot1_11%></td>
			<td bgcolor='#ffffff' align='center' class='TDCont'><%=tot1_12%></td>
		</tr>

		<tr bgcolor='#EEF6FF'>
			<td align='center' class='TDCont'  width='150' >사이버</td>
			<td bgcolor='#ffffff' align='center' class='TDCont'><%=tot2_1%></td>
			<td bgcolor='#ffffff' align='center' class='TDCont'><%=tot2_2%></td>
			<td bgcolor='#ffffff' align='center' class='TDCont'><%=tot2_3%></td>
			<td bgcolor='#ffffff' align='center' class='TDCont'><%=tot2_4%></td>
			<td bgcolor='#ffffff' align='center' class='TDCont'><%=tot2_5%></td>
			<td bgcolor='#ffffff' align='center' class='TDCont'><%=tot2_6%></td>
			<td bgcolor='#ffffff' align='center' class='TDCont'><%=tot2_7%></td>
			<td bgcolor='#ffffff' align='center' class='TDCont'><%=tot2_8%></td>
			<td bgcolor='#ffffff' align='center' class='TDCont'><%=tot2_9%></td>
			<td bgcolor='#ffffff' align='center' class='TDCont'><%=tot2_10%></td>
			<td bgcolor='#ffffff' align='center' class='TDCont'><%=tot2_11%></td>
			<td bgcolor='#ffffff' align='center' class='TDCont'><%=tot2_12%></td>
		</tr>
		<tr bgcolor='#EEF6FF'>
			<td align='center' class='TDCont'  width='150' >계</td>
			<td align='center' class='TDCont'><%=tot1%></td>
			<td align='center' class='TDCont'><%=tot2%></td>
			<td align='center' class='TDCont'><%=tot3%></td>
			<td align='center' class='TDCont'><%=tot4%></td>
			<td align='center' class='TDCont'><%=tot5%></td>
			<td align='center' class='TDCont'><%=tot6%></td>
			<td align='center' class='TDCont'><%=tot7%></td>
			<td align='center' class='TDCont'><%=tot8%></td>
			<td align='center' class='TDCont'><%=tot9%></td>
			<td align='center' class='TDCont'><%=tot10%></td>
			<td align='center' class='TDCont'><%=tot11%></td>
			<td align='center' class='TDCont'><%=tot12%></td>
		</tr>
<%

		tot1 = 0
		tot2 = 0
		tot3 = 0
		tot4 = 0
		tot5 = 0
		tot6 = 0
		tot7 = 0
		tot8 = 0
		tot9 = 0
		tot10 = 0
		tot11 = 0
		tot12 = 0
		if rs.eof then
			exit do
		end if
	loop

	'분야별총계
	SQL = "select * from ( SELECT	'1' gubun, count(incode) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and CONSULTGB = 'A' and ACLASS in ('A','B')" '자살
	SQL = SQL & "	union all SELECT	'2' gubun, count(incode) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and CONSULTGB = 'B' and ACLASS in ('A','B')" '부적응
	SQL = SQL & "	union all SELECT	'3' gubun, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and CONSULTGB = 'C' and ACLASS in ('A','B')" '인권
	SQL = SQL & "	union all SELECT	'4' gubun, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and CONSULTGB = 'D' and ACLASS in ('A','B')" '이성
	SQL = SQL & "	union all SELECT	'5' gubun, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and CONSULTGB = 'F' and ACLASS in ('A','B')" '건강	'response.write SQL
	SQL = SQL & "	union all SELECT	'6' gubun, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and CONSULTGB = 'G' and ACLASS in ('A','B')" '전역
	SQL = SQL & "	union all SELECT	'7' gubun, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and CONSULTGB = 'H' and ACLASS in ('A','B')" '가정
	SQL = SQL & "	union all SELECT	'8' gubun, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and CONSULTGB = 'I' and ACLASS in ('A','B')" '전출
	SQL = SQL & "	union all SELECT	'9' gubun, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and CONSULTGB = 'J' and ACLASS in ('A','B')" '대체
	SQL = SQL & "	union all SELECT	'10' gubun, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and CONSULTGB = 'K' and ACLASS in ('A','B')" '조치
	SQL = SQL & "	union all SELECT	'11' gubun, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and CONSULTGB = 'Z' and ACLASS in ('A','B')" '기타
	SQL = SQL & "	union all SELECT	'12' gubun, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and ACLASS in ('A','B') ) a order by gubun" '총계

	set Rs = db.execute(SQL)

	do until rs.eof

		if rs("gubun") = "1" then
			tot1 = rs("cnt")
		elseif rs("gubun") = "2" then
			tot2 = rs("cnt")
		elseif rs("gubun") = "3" then
			tot3 = rs("cnt")
		elseif rs("gubun") = "4" then
			tot4 = rs("cnt")
		elseif rs("gubun") = "5" then
			tot5 = rs("cnt")
		elseif rs("gubun") = "6" then
			tot6 = rs("cnt")
		elseif rs("gubun") = "7" then
			tot7 = rs("cnt")
		elseif rs("gubun") = "8" then
			tot8 = rs("cnt")
		elseif rs("gubun") = "9" then
			tot9 = rs("cnt")
		elseif rs("gubun") = "10" then
			tot10 = rs("cnt")
		elseif rs("gubun") = "11" then
			tot11 = rs("cnt")
		elseif rs("gubun") = "12" then
			tot12 = rs("cnt")
		end if

		rs.movenext
		if rs.eof then
			exit do
		end if
	loop

%>

				<tr bgcolor='#FFEEF9'>
					<td align='center' class='TDCont' colspan='2'>총계</td>
					<td align='center' class='TDCont'><%=tot1%></td>
					<td align='center' class='TDCont'><%=tot2%></td>
					<td align='center' class='TDCont'><%=tot3%></td>
					<td align='center' class='TDCont'><%=tot4%></td>
					<td align='center' class='TDCont'><%=tot5%></td>
					<td align='center' class='TDCont'><%=tot6%></td>
					<td align='center' class='TDCont'><%=tot7%></td>
					<td align='center' class='TDCont'><%=tot8%></td>
					<td align='center' class='TDCont'><%=tot9%></td>
					<td align='center' class='TDCont'><%=tot10%></td>
					<td align='center' class='TDCont'><%=tot11%></td>
					<td align='center' class='TDCont'><%=tot12%></td>
				</tr>
			</table>

		
			<table width="940" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table  width="940" border="0" cellpadding="1" cellspacing="1" bgcolor="#EFECE5" align="center">
			    <tr height="30">
					<td align="left" bgcolor="#FFFFFF" class="TDCont" colspan="9">&nbsp;<img src="/Images/dot_01.gif" >&nbsp;<b><font color="#ff00ff"></font> 요일별</b></td>
				</tr>
				<tr bgcolor='#EEF6FF'>
					<td align='center' class='TDCont'  width='150'>구분</td>
					<td align='center' class='TDCont' width='150'>월</td>
					<td align='center' class='TDCont' width='150'>화</td>
					<td align='center' class='TDCont' width='150'>수</td>
					<td align='center' class='TDCont' width='150'>목</td>
					<td align='center' class='TDCont' width='150'>금</td>
					<td align='center' class='TDCont'width='150' >토</td>
					<td align='center' class='TDCont'width='150' >일</td>
					<td align='center' class='TDCont'  width='150'>총계</td>
				</tr>
<%

	'상담관별 총계
	SQL = "select * from ( SELECT	'1' gubun, incode, count(incode) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and datepart(weekday,jubdate) = 2 group by incode" '
	SQL = SQL & "	union all SELECT	'2' gubun, incode, count(incode) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and datepart(weekday,jubdate) = 3 group by incode" '
	SQL = SQL & "	union all SELECT	'3' gubun, incode, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and datepart(weekday,jubdate) = 4 group by incode" '
	SQL = SQL & "	union all SELECT	'4' gubun, incode, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and datepart(weekday,jubdate) = 5 group by incode" '
	SQL = SQL & "	union all SELECT	'5' gubun, incode, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and datepart(weekday,jubdate) = 6 group by incode" '
	SQL = SQL & "	union all SELECT	'6' gubun, incode, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and datepart(weekday,jubdate) = 7 group by incode" '
	SQL = SQL & "	union all SELECT	'7' gubun, incode, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and datepart(weekday,jubdate) = 8 group by incode" '
	SQL = SQL & "	union all SELECT	'8' gubun, incode, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' group by incode) a order by incode, gubun" '

	set Rs = db.execute(SQL)

	tot1 = 0
	tot2 = 0
	tot3 = 0
	tot4 = 0
	tot5 = 0
	tot6 = 0
	tot7 = 0
	tot8 = 0
	tot9 = 0
	do until rs.eof

		incode = rs("incode")
		do until incode <> rs("incode")
			if rs("gubun") = "1" then
				tot1 = rs("cnt")
			elseif rs("gubun") = "2" then
				tot2 = rs("cnt")
			elseif rs("gubun") = "3" then
				tot3 = rs("cnt")
			elseif rs("gubun") = "4" then
				tot4 = rs("cnt")
			elseif rs("gubun") = "5" then
				tot5 = rs("cnt")
			elseif rs("gubun") = "6" then
				tot6 = rs("cnt")
			elseif rs("gubun") = "7" then
				tot7 = rs("cnt")
			elseif rs("gubun") = "8" then
				tot8 = rs("cnt")
			end if

			rs.movenext
			if rs.eof then
				exit do
			end if
		loop


%>
				<tr bgcolor='#EEF6FF'>
					<td align='center' class='TDCont'  width='150'><%=db_getUserName(incode)%></td>
					<td bgcolor='#ffffff' align='center' class='TDCont'><%=tot1%></td>
					<td bgcolor='#ffffff' align='center' class='TDCont'><%=tot2%></td>
					<td bgcolor='#ffffff' align='center' class='TDCont'><%=tot3%></td>
					<td bgcolor='#ffffff' align='center' class='TDCont'><%=tot4%></td>
					<td bgcolor='#ffffff' align='center' class='TDCont'><%=tot5%></td>
					<td bgcolor='#ffffff' align='center' class='TDCont'><%=tot6%></td>
					<td bgcolor='#ffffff' align='center' class='TDCont'><%=tot7%></td>
					<td bgcolor='#ffffff' align='center' class='TDCont'><%=tot8%></td>
				</tr>
<%

		if rs.eof then
			exit do
		end if
	loop



	'상담관별 총계
	SQL = "select * from ( SELECT	'1' gubun, count(incode) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and datepart(weekday,jubdate) = 2" '
	SQL = SQL & "	union all SELECT	'2' gubun, count(incode) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and datepart(weekday,jubdate) = 3" '
	SQL = SQL & "	union all SELECT	'3' gubun, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and datepart(weekday,jubdate) = 4" '
	SQL = SQL & "	union all SELECT	'4' gubun, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and datepart(weekday,jubdate) = 5" '
	SQL = SQL & "	union all SELECT	'5' gubun, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and datepart(weekday,jubdate) = 6" '
	SQL = SQL & "	union all SELECT	'6' gubun, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and datepart(weekday,jubdate) = 7" '
	SQL = SQL & "	union all SELECT	'7' gubun, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and datepart(weekday,jubdate) = 8" '
	SQL = SQL & "	union all SELECT	'8' gubun, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "') a order by gubun" '

	set Rs = db.execute(SQL)

	tot1 = 0
	tot2 = 0
	tot3 = 0
	tot4 = 0
	tot5 = 0
	tot6 = 0
	tot7 = 0
	tot8 = 0
	tot9 = 0
	do until rs.eof


			if rs("gubun") = "1" then
				tot1 = rs("cnt")
			elseif rs("gubun") = "2" then
				tot2 = rs("cnt")
			elseif rs("gubun") = "3" then
				tot3 = rs("cnt")
			elseif rs("gubun") = "4" then
				tot4 = rs("cnt")
			elseif rs("gubun") = "5" then
				tot5 = rs("cnt")
			elseif rs("gubun") = "6" then
				tot6 = rs("cnt")
			elseif rs("gubun") = "7" then
				tot7 = rs("cnt")
			elseif rs("gubun") = "8" then
				tot8 = rs("cnt")
			end if

		rs.movenext
		if rs.eof then
			exit do
		end if
	loop


%>
				<tr bgcolor='#FFEEF9'>
					<td align='center' class='TDCont'  width='150'>계</td>
					<td align='center' class='TDCont'><%=tot1%></td>
					<td align='center' class='TDCont'><%=tot2%></td>
					<td align='center' class='TDCont'><%=tot3%></td>
					<td align='center' class='TDCont'><%=tot4%></td>
					<td align='center' class='TDCont'><%=tot5%></td>
					<td align='center' class='TDCont'><%=tot6%></td>
					<td align='center' class='TDCont'><%=tot7%></td>
					<td align='center' class='TDCont'><%=tot8%></td>
				</tr>

			</table>
			</DIV>
		</td>
	</tr>
</table>

<% End if %>

<!-- #include virtual="/Include/Bottom.asp" -->