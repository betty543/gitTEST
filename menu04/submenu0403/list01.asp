<!-- #include virtual="/Include/Top.asp" -->
<%
	'####### �Ķ���� ##################################################################################
	QueryYN = request("QueryYN")
	FromDate = request("FromDate")
	ToDate = request("ToDate")
	whereCD3 = Trim(request("whereCD3"))
	whereCD7 = Trim(request("whereCD7"))
	SS_Login_Grade = SESSION("SS_Login_Grade")

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
			        <td class="TDCont">��ȸ�Ⱓ :
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

		
			<table width="940" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table  width="940" border="0" cellpadding="1" cellspacing="1" bgcolor="#EFECE5" align="center">
			    <tr height="30">
					<td align="left" bgcolor="#FFFFFF" class="TDCont" colspan="9">&nbsp;<img src="/Images/dot_01.gif" >&nbsp;<b><font color="#ff00ff"></font> ���Ϻ�</b></td>
				</tr>
				<tr bgcolor='#EEF6FF'>
					<td align='center' class='TDCont'  width='150'>����</td>
					<td align='center' class='TDCont' width='150'>��</td>
					<td align='center' class='TDCont' width='150'>ȭ</td>
					<td align='center' class='TDCont' width='150'>��</td>
					<td align='center' class='TDCont' width='150'>��</td>
					<td align='center' class='TDCont' width='150'>��</td>
					<td align='center' class='TDCont'width='150' >��</td>
					<td align='center' class='TDCont'width='150' >��</td>
					<td align='center' class='TDCont'  width='150'>�Ѱ�</td>
				</tr>
<%

	'������ �Ѱ�
	SQL = "select * from ( SELECT	'1' gubun, telkind, count(telkind) cnt FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and datepart(weekday,jubdate) = 2 "
	if SS_Login_Grade <> "A" and SS_Login_Grade <> "C" then
		SQL = SQL & " and TELKIND = '" & SS_Login_Grade &"'"
	end if
	SQL = SQL & "	group by telkind" '
	SQL = SQL & "	union all SELECT	'2' gubun, telkind, count(telkind) cnt FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and datepart(weekday,jubdate) = 3 "
	if SS_Login_Grade <> "A" and SS_Login_Grade <> "C" then
		SQL = SQL & " and TELKIND = '" & SS_Login_Grade &"'"
	end if
	SQL = SQL & "	group by telkind" '
	SQL = SQL & "	union all SELECT	'3' gubun, telkind, count(telkind) cnt  FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and datepart(weekday,jubdate) = 4"
	if SS_Login_Grade <> "A" and SS_Login_Grade <> "C" then
		SQL = SQL & " and TELKIND = '" & SS_Login_Grade &"'"
	end if
	SQL = SQL & "	group by telkind" '
	SQL = SQL & "	union all SELECT	'4' gubun, telkind, count(telkind) cnt  FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and datepart(weekday,jubdate) = 5 "
	if SS_Login_Grade <> "A" and SS_Login_Grade <> "C" then
		SQL = SQL & " and TELKIND = '" & SS_Login_Grade &"'"
	end if
	SQL = SQL & "	group by telkind" '
	SQL = SQL & "	union all SELECT	'5' gubun, telkind, count(telkind) cnt  FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and datepart(weekday,jubdate) = 6 "
	if SS_Login_Grade <> "A" and SS_Login_Grade <> "C" then
		SQL = SQL & " and TELKIND = '" & SS_Login_Grade &"'"
	end if
	SQL = SQL & "	group by telkind" '
	SQL = SQL & "	union all SELECT	'6' gubun, telkind, count(telkind) cnt  FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and datepart(weekday,jubdate) = 7 "
	if SS_Login_Grade <> "A" and SS_Login_Grade <> "C" then
		SQL = SQL & " and TELKIND = '" & SS_Login_Grade &"'"
	end if
	SQL = SQL & "	group by telkind" '
	SQL = SQL & "	union all SELECT	'7' gubun, telkind, count(telkind) cnt  FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and datepart(weekday,jubdate) = 8 "
	if SS_Login_Grade <> "A" and SS_Login_Grade <> "C" then
		SQL = SQL & " and TELKIND = '" & SS_Login_Grade &"'"
	end if
	SQL = SQL & "	group by telkind" '
	SQL = SQL & "	union all SELECT	'8' gubun, telkind, count(telkind) cnt  FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
	if SS_Login_Grade <> "A" and SS_Login_Grade <> "C" then
		SQL = SQL & " and TELKIND = '" & SS_Login_Grade &"'"
	end if
	SQL = SQL & "	group by telkind) a order by telkind, gubun" '

	set Rs = db.execute(SQL)


	do until rs.eof

	tot1 = 0
	tot2 = 0
	tot3 = 0
	tot4 = 0
	tot5 = 0
	tot6 = 0
	tot7 = 0
	tot8 = 0
	tot9 = 0

		telkind = rs("telkind")
		do until telkind <> rs("telkind")
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
					<td align='center' class='TDCont'  width='150'><%=db_getcodename("Z04",telkind)%></td>
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



	'������ �Ѱ�
	SQL = "select * from ( SELECT	'1' gubun, count(telkind) cnt FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and datepart(weekday,jubdate) = 2" '
	SQL = SQL & "	union all SELECT	'2' gubun, count(telkind) cnt FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and datepart(weekday,jubdate) = 3" '
	SQL = SQL & "	union all SELECT	'3' gubun, count(telkind) cnt  FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and datepart(weekday,jubdate) = 4" '
	SQL = SQL & "	union all SELECT	'4' gubun, count(telkind) cnt  FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and datepart(weekday,jubdate) = 5" '
	SQL = SQL & "	union all SELECT	'5' gubun, count(telkind) cnt  FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and datepart(weekday,jubdate) = 6" '
	SQL = SQL & "	union all SELECT	'6' gubun, count(telkind) cnt  FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and datepart(weekday,jubdate) = 7" '
	SQL = SQL & "	union all SELECT	'7' gubun, count(telkind) cnt  FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and datepart(weekday,jubdate) = 8" '
	SQL = SQL & "	union all SELECT	'8' gubun, count(telkind) cnt  FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
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
					<td align='center' class='TDCont'  width='150'>��</td>
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


			<!--<DIV style="OVERFLOW-Y:auto; OVERFLOW-X:auto; MARGIN: 0px 0px 0px 0px; WIDTH:940; HEIGHT:500;">-->
			<table width="100%" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table width="940"  border="0" cellpadding="1" cellspacing="1" bgcolor="#EFECE5" align="center">
			    <tr height="30">
					<td align="left" bgcolor="#FFFFFF" class="TDCont" colspan="8">&nbsp;<img src="/Images/dot_01.gif" >&nbsp;<b><font color="#ff00ff"></font> �����</b></td>
				</tr>
				<tr bgcolor='#EEF6FF'>
					<td align='center' class='TDCont'  width='150'>����</td>
					<td align='center' class='TDCont' >���</td>
					<td align='center' class='TDCont' >����</td>
					<td align='center' class='TDCont' >ħ��</td>
					<td align='center' class='TDCont' >���ͳ�</td>
					<td align='center' class='TDCont' >��Ʈ���</td>
					<td align='center' class='TDCont' >���</td>
					<td align='center' class='TDCont' >�Ѱ�</td>
				</tr>

<%
	'�������


	SQL = "select * from ( SELECT	'1' gubun, telkind, count(telkind) cnt FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and ACLASS = 'A' group by telkind" '���,
	SQL = SQL & "	union all SELECT	'2' gubun, telkind, count(telkind) cnt FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and ACLASS = 'B' group by telkind" '����,

	SQL = SQL & "	union all SELECT	'3' gubun, telkind, count(telkind) cnt  FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and ACLASS = 'C' group by telkind" 'ħ��
	SQL = SQL & "	union all SELECT	'4' gubun, telkind, count(telkind) cnt  FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and ACLASS = 'D' and CHANNELGB = 'C' group by telkind" '���̹�

	SQL = SQL & "	union all SELECT	'5' gubun, telkind, count(telkind) cnt  FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and ACLASS = 'D' and CHANNELGB = 'D' group by telkind" '���̹�
	'response.write SQL
	SQL = SQL & "	union all SELECT	'6' gubun, telkind, count(telkind) cnt  FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and ACLASS = 'E' group by telkind" '���

	SQL = SQL & "	union all SELECT	'7' gubun, telkind, count(telkind) cnt  FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and ACLASS is not null and ACLASS <> '' group by telkind) a order by telkind, gubun"

	'response.write SQL	
	set Rs = db.execute(SQL)

	do until rs.eof


	tot1 = 0
	tot2 = 0
	tot3 = 0
	tot4 = 0
	tot5 = 0
	tot6 = 0
	tot7 = 0
	tot8 = 0
	tot9 = 0



		telkind = rs("telkind")
		do until telkind <> rs("telkind")
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
			<td align='center' class='TDCont'  width='150' ><%=db_getcodename("Z04",telkind)%></td>
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


	'�������


	SQL = "SELECT	'1' gubun, count(*) cnt FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and ACLASS = 'A'" '���,
	SQL = SQL & "	union all SELECT	'2' gubun, count(*) cnt FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and ACLASS = 'B'" '����,

	SQL = SQL & "	union all SELECT	'3' gubun, count(*) cnt FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and ACLASS = 'C'" 'ħ��
	SQL = SQL & "	union all SELECT	'4' gubun, count(*) cnt FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and ACLASS = 'D' and CHANNELGB = 'C'" '���̹�

	SQL = SQL & "	union all SELECT	'5' gubun, count(*) cnt FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and ACLASS = 'D' and CHANNELGB = 'D'" '���̹�
	'response.write SQL
	SQL = SQL & "	union all SELECT	'6' gubun, count(*) cnt FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and ACLASS = 'E'" '���

	SQL = SQL & "	union all SELECT	'7' gubun, count(*) cnt FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
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
					<td align='center' class='TDCont'>��</td>
					<td align='center' class='TDCont' ><%=tot1%></td>
					<td  align='center' class='TDCont' ><%=tot2%></td>
					<td  align='center' class='TDCont' ><%=tot3%></td>
					<td  align='center' class='TDCont' ><%=tot4%></td>
					<td  align='center' class='TDCont' ><%=tot5%></td>
					<td  align='center' class='TDCont' ><%=tot6%></td>
					<td  align='center' class='TDCont' ><%=tot7%></td>
				</tr>

				<tr><td colspan="100" height="1" bgcolor="#FFFFFF"></td></tr>
				<%'####### �����ڷᰡ ����. %>
				<%
				'---------------------------------------------------------------------------------------------------------------------
				'����Ÿ �̾ƿ���
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
			<!--��޺�-->
			<table width="100%" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table  width="940"  border="0" cellpadding="1" cellspacing="1" bgcolor="#EFECE5" align="center">
			    <tr height="30">
					<td align="left" bgcolor="#FFFFFF" class="TDCont" colspan="15">&nbsp;<img src="/Images/dot_01.gif" >&nbsp;<b><font color="#ff00ff"></font> ��޺�</b></td>
				</tr>
				<tr bgcolor='#EEF6FF'>
					<td align='center' class='TDCont' width='100' colspan = '2' rowspan='2' nowrap>&nbsp;</td>
					<td align='center' class='TDCont' nowrap colspan='5'>��</td>
					<td align='center' class='TDCont' nowrap colspan='6'>����</td>
					<td align='center' class='TDCont' nowrap rowspan='2'>��Ÿ</td>
					<td align='center' class='TDCont' nowrap rowspan='2'>�Ѱ�</td>
				</tr>
				<tr bgcolor='#EEF6FF'>
					<td align='center' class='TDCont' nowrap>�̺�</td>
					<td align='center' class='TDCont' nowrap>�Ϻ�</td>
					<td align='center' class='TDCont' nowrap>��</td>
					<td align='center' class='TDCont' nowrap>����</td>
					<td align='center' class='TDCont' nowrap>�̻�</td>
					<td align='center' class='TDCont' nowrap>�λ��</td>
					<td align='center' class='TDCont' nowrap>����</td>
					<td align='center' class='TDCont' nowrap>����</td>
					<td align='center' class='TDCont' nowrap>�屺</td>
					<td align='center' class='TDCont' nowrap>������Ȱ<br>��������</td>
					<td align='center' class='TDCont' nowrap>�̻�</td>
				</tr>

				<%'####### �����ڷᰡ ����. %>
				<%
				'---------------------------------------------------------------------------------------------------------------------
				'����Ÿ �̾ƿ���
				'---------------------------------------------------------------------------------------------------------------------

	'��޺�

	SQL = "select * from ( SELECT	'1' gubun, telkind, count(telkind) cnt FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
	SQL = SQL & "	AND		level1 = 'A' and level2 = 'A' group by telkind"	'�̺�
	SQL = SQL & " union all SELECT	'2' gubun, telkind, count(telkind) cnt FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
	SQL = SQL & "	AND		level1 = 'A' and level2 = 'B' group by telkind"	'�Ϻ�
	SQL = SQL & " union all SELECT	'3' gubun, telkind, count(telkind) cnt FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
	SQL = SQL & "	AND		level1 = 'A' and level2 = 'C' group by telkind"	'��
	SQL = SQL & " union all SELECT	'4' gubun, telkind, count(telkind) cnt FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
	SQL = SQL & "	AND		level1 = 'A' and level2 = 'D' group by telkind"	'����
	SQL = SQL & " union all SELECT	'5' gubun, telkind, count(telkind) cnt FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
	SQL = SQL & "	AND		level1 = 'A' and level2 = 'Y' group by telkind"	'�̻�
	SQL = SQL & " union all SELECT	'6' gubun, telkind, count(telkind) cnt FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
	SQL = SQL & "	AND		level1 = 'B' and level2 = 'A' group by telkind" '�λ��
	SQL = SQL & " union all SELECT	'7' gubun, telkind, count(telkind) cnt FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
	SQL = SQL & "	AND		level1 = 'B' and level2 = 'B' group by telkind" '����
	SQL = SQL & " union all SELECT	'8' gubun, telkind, count(telkind) cnt FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
	SQL = SQL & "	AND		level1 = 'B' and level2 = 'C' group by telkind" '����
	SQL = SQL & " union all SELECT	'9' gubun, telkind, count(telkind) cnt FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
	SQL = SQL & "	AND		level1 = 'B' and level2 = 'D' group by telkind" '�屺
	SQL = SQL & " union all SELECT	'10' gubun, telkind, count(telkind) cnt FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
	SQL = SQL & "	AND		level1 = 'B' and level2 = 'E' group by telkind" '������Ȱ��������
	SQL = SQL & " union all SELECT	'11' gubun, telkind, count(telkind) cnt FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
	SQL = SQL & "	AND		level1 = 'B' and level2 = 'Y' group by telkind" '���ι̻�
	SQL = SQL & " union all SELECT	'12' gubun, telkind, count(telkind) cnt FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
	SQL = SQL & "	AND		level1 NOT IN ('A','B') and level1 is not null and level1 <> '' group by telkind"
	SQL = SQL & " union all SELECT	'13' gubun, telkind, count(telkind) cnt FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and level1 is not null and rtrim(level1) <> '' group by telkind ) a order by telkind, gubun"

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
		telkind = rs("telkind")
		do until telkind <> rs("telkind")
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
		' �ش������� ���, ����, ���̹� ã�Ƴ���
		'-------------------------------------------------------------------------------------------------
		SQL = "select * from ( SELECT	'1' gubun, ACLASS, count(ACLASS) cnt FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
		SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
		SQL = SQL & "	AND		level1 = 'A' and level2 = 'A' AND telkind = '"&telkind&"'	group by ACLASS"	'�̺�
		SQL = SQL & " union all SELECT	'2' gubun, ACLASS, count(ACLASS) cnt FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
		SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
		SQL = SQL & "	AND		level1 = 'A' and level2 = 'B' AND telkind = '"&telkind&"'	group by ACLASS"	'�Ϻ�
		SQL = SQL & " union all SELECT	'3' gubun, ACLASS, count(ACLASS) cnt FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
		SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
		SQL = SQL & "	AND		level1 = 'A' and level2 = 'C' AND telkind = '"&telkind&"'	group by ACLASS"	'��
		SQL = SQL & " union all SELECT	'4' gubun, ACLASS, count(ACLASS) cnt FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
		SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
		SQL = SQL & "	AND		level1 = 'A' and level2 = 'D' AND telkind = '"&telkind&"'	group by ACLASS"	'����
		SQL = SQL & " union all SELECT	'5' gubun, ACLASS, count(ACLASS) cnt FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
		SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
		SQL = SQL & "	AND		level1 = 'A' and level2 = 'Y' AND telkind = '"&telkind&"'	group by ACLASS"	'�̻�
		SQL = SQL & " union all SELECT	'6' gubun, ACLASS, count(ACLASS) cnt FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
		SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
		SQL = SQL & "	AND		level1 = 'B' and level2 = 'A' AND telkind = '"&telkind&"'	group by ACLASS" '�λ��
		SQL = SQL & " union all SELECT	'7' gubun, ACLASS, count(ACLASS) cnt FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
		SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
		SQL = SQL & "	AND		level1 = 'B' and level2 = 'B' AND telkind = '"&telkind&"'	group by ACLASS" '����
		SQL = SQL & " union all SELECT	'8' gubun, ACLASS, count(ACLASS) cnt FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
		SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
		SQL = SQL & "	AND		level1 = 'B' and level2 = 'C' AND telkind = '"&telkind&"'	group by ACLASS" '����
		SQL = SQL & " union all SELECT	'9' gubun, ACLASS, count(ACLASS) cnt FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
		SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
		SQL = SQL & "	AND		level1 = 'B' and level2 = 'D' AND telkind = '"&telkind&"'	group by ACLASS" '�屺
		SQL = SQL & " union all SELECT	'10' gubun, ACLASS, count(ACLASS) cnt FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
		SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
		SQL = SQL & "	AND		level1 = 'B' and level2 = 'E' AND telkind = '"&telkind&"'	group by ACLASS" '������Ȱ��������
		SQL = SQL & " union all SELECT	'11' gubun, ACLASS, count(ACLASS) cnt FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
		SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
		SQL = SQL & "	AND		level1 = 'B' and level2 = 'Y' AND telkind = '"&telkind&"'	group by ACLASS" '���ι̻�
		SQL = SQL & " union all SELECT	'12' gubun, ACLASS, count(ACLASS) cnt FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
		SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
		SQL = SQL & "	AND		level1 NOT IN ('A','B') and level1 is not null and level1 <> '' AND telkind = '"&telkind&"'	group by ACLASS"
		SQL = SQL & " union all SELECT	'13' gubun, ACLASS, count(ACLASS) cnt FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
		SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and level1 is not null and level1 <> '' AND telkind = '"&telkind&"' group by ACLASS ) a order by ACLASS, gubun"

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
					<td align='center' class='TDCont' width="100" nowrap rowspan='4'><%=db_getcodename("Z04",telkind)%></td>
					<td align='center' class='TDCont' width="100" nowrap>���</td>
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
					<td align='center' class='TDCont' width="100" nowrap>����</td>
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
					<td align='center' class='TDCont' width="100" nowrap>���̹�</td>
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
					<td align='center' class='TDCont' width="100" nowrap>��</td>
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


	SQL = "SELECT	'1' gubun, count(*) cnt FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
	SQL = SQL & "	AND		level1 = 'A' and level2 = 'A'"	'�̺�
	SQL = SQL & " union all SELECT	'2' gubun, count(*) cnt FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
	SQL = SQL & "	AND		level1 = 'A' and level2 = 'B'"	'�Ϻ�
	SQL = SQL & " union all SELECT	'3' gubun, count(*) cnt FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
	SQL = SQL & "	AND		level1 = 'A' and level2 = 'C'"	'��
	SQL = SQL & " union all SELECT	'4' gubun, count(*) cnt FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
	SQL = SQL & "	AND		level1 = 'A' and level2 = 'D'"	'����
	SQL = SQL & " union all SELECT	'5' gubun, count(*) cnt FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
	SQL = SQL & "	AND		level1 = 'A' and level2 = 'Y'"	'�̻�
	SQL = SQL & " union all SELECT	'6' gubun, count(*) cnt FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
	SQL = SQL & "	AND		level1 = 'B' and level2 = 'A'" '�λ��
	SQL = SQL & " union all SELECT	'7' gubun, count(*) cnt FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
	SQL = SQL & "	AND		level1 = 'B' and level2 = 'B'" '����
	SQL = SQL & " union all SELECT	'8' gubun, count(*) cnt FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
	SQL = SQL & "	AND		level1 = 'B' and level2 = 'C'" '����
	SQL = SQL & " union all SELECT	'9' gubun, count(*) cnt FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
	SQL = SQL & "	AND		level1 = 'B' and level2 = 'D'" '�屺
	SQL = SQL & " union all SELECT	'10' gubun, count(*) cnt FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
	SQL = SQL & "	AND		level1 = 'B' and level2 = 'E'" '������Ȱ��������
	SQL = SQL & " union all SELECT	'11' gubun, count(*) cnt FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
	SQL = SQL & "	AND		level1 = 'B' and level2 = 'Y'" '���ι̻�
	SQL = SQL & " union all SELECT	'12' gubun, count(*) cnt FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "'"
	SQL = SQL & "	AND		level1 NOT IN ('A','B') and level1 is not null and level1 <> '' "
	SQL = SQL & " union all SELECT	'13' gubun, count(*) cnt FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
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
					<td align='center' class='TDCont' colspan='2'>�Ѱ�</td>
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
					<td align="left" bgcolor="#FFFFFF" class="TDCont" colspan="13">&nbsp;<img src="/Images/dot_01.gif" >&nbsp;<b><font color="#ff00ff"></font> �δ뺰</b></td>
				</tr>
				<tr bgcolor='#EEF6FF'>
					<td align='center' class='TDCont'  width='150'>����</td>
					<td align='center' class='TDCont' width='150'>1��</td>
					<td align='center' class='TDCont' width='150'>2�ۻ�</td>
					<td align='center' class='TDCont' width='150'>3��</td>
					<td align='center' class='TDCont' width='150'>����</td>
					<td align='center' class='TDCont' width='150'>������</td>
					<td align='center' class='TDCont'width='150' >������</td>
					<td align='center' class='TDCont'width='150'>Ư����</td>
					<td align='center' class='TDCont'width='150'>Ÿ�δ�</td>
					<td align='center' class='TDCont'width='150'>��Ÿ</td>
					<td align='center' class='TDCont'width='150'>�̻�</td>
					<td align='center' class='TDCont'  width='150'>�Ѱ�</td>
				</tr>
<%

	SQL = "select * from ( SELECT	'1' gubun, telkind, count(telkind) cnt FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and SOSOKGB = 'A' and ACLASS in ('A','B','C') group by telkind" '1��
	SQL = SQL & "	union all SELECT	'2' gubun, telkind, count(telkind) cnt FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and SOSOKGB = 'B' and ACLASS in ('A','B','C') group by telkind" '2�ۻ�
	SQL = SQL & "	union all SELECT	'3' gubun, telkind, count(telkind) cnt  FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and SOSOKGB = 'C' and ACLASS in ('A','B','C') group by telkind" '3��
	SQL = SQL & "	union all SELECT	'4' gubun, telkind, count(telkind) cnt  FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and SOSOKGB = 'D' and ACLASS in ('A','B','C') group by telkind" '����
	SQL = SQL & "	union all SELECT	'5' gubun, telkind, count(telkind) cnt  FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and SOSOKGB = 'E' and ACLASS in ('A','B','C') group by telkind" '������	'response.write SQL
	SQL = SQL & "	union all SELECT	'6' gubun, telkind, count(telkind) cnt  FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and SOSOKGB = 'F' and ACLASS in ('A','B','C') group by telkind" '������
	SQL = SQL & "	union all SELECT	'7' gubun, telkind, count(telkind) cnt  FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and SOSOKGB = 'G' and ACLASS in ('A','B','C') group by telkind" 'Ư����
	SQL = SQL & "	union all SELECT	'8' gubun, telkind, count(telkind) cnt  FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and SOSOKGB = 'I' and ACLASS in ('A','B','C') group by telkind" 'Ÿ�δ�
	SQL = SQL & "	union all SELECT	'9' gubun, telkind, count(telkind) cnt  FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and SOSOKGB = 'H' and ACLASS in ('A','B','C') group by telkind " '��Ÿ
	SQL = SQL & "	union all SELECT	'10' gubun, telkind, count(telkind) cnt  FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and SOSOKGB NOT IN ('A','B','C','D','E','F','G','H','I') and ACLASS in ('A','B','C') group by telkind " '�̻�
	SQL = SQL & "	union all SELECT	'11' gubun, telkind, count(telkind) cnt  FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and ACLASS in ('A','B','C') group by telkind ) a order by telkind, gubun" '�Ѱ�

	set Rs = db.execute(SQL)

	do until rs.eof

		telkind = rs("telkind")
		do until telkind <> rs("telkind")
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
			<td align='center' class='TDCont'  width='150' ><%=db_getcodename("Z04",telkind)%></td>
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

	'�δ뺰
	SQL = "select * from ( SELECT	'1' gubun, count(telkind) cnt FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and SOSOKGB = 'A' and ACLASS in ('A','B','C')" '1��
	SQL = SQL & "	union all SELECT	'2' gubun, count(telkind) cnt FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and SOSOKGB = 'B' and ACLASS in ('A','B','C')" '2�ۻ�
	SQL = SQL & "	union all SELECT	'3' gubun, count(telkind) cnt  FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and SOSOKGB = 'C' and ACLASS in ('A','B','C')" '3��
	SQL = SQL & "	union all SELECT	'4' gubun, count(telkind) cnt  FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and SOSOKGB = 'D' and ACLASS in ('A','B','C')" '����
	SQL = SQL & "	union all SELECT	'5' gubun, count(telkind) cnt  FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and SOSOKGB = 'E' and ACLASS in ('A','B','C')" '������	'response.write SQL
	SQL = SQL & "	union all SELECT	'6' gubun, count(telkind) cnt  FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and SOSOKGB = 'F' and ACLASS in ('A','B','C')" '������
	SQL = SQL & "	union all SELECT	'7' gubun, count(telkind) cnt  FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and SOSOKGB = 'G' and ACLASS in ('A','B','C')" 'Ư����
	SQL = SQL & "	union all SELECT	'8' gubun, count(telkind) cnt  FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and SOSOKGB = 'I' and ACLASS in ('A','B','C')" 'Ÿ�δ�
	SQL = SQL & "	union all SELECT	'9' gubun, count(telkind) cnt  FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and SOSOKGB = 'H' and ACLASS in ('A','B','C')" '��Ÿ
	SQL = SQL & "	union all SELECT	'10' gubun, count(telkind) cnt  FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and SOSOKGB NOT IN ('A','B','C','D','E','F','G','H','I') and ACLASS in ('A','B','C')" '�̻�
	SQL = SQL & "	union all SELECT	'11' gubun, count(telkind) cnt  FROM	TB_callhistory	where jubdate >= '" & FromDate & "'"
	SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and ACLASS in ('A','B','C')) a order by gubun" '�Ѱ�

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
					<td align='center' class='TDCont'  width='150' >��</td>
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



<% End if %>

<!-- #include virtual="/Include/Bottom.asp" -->