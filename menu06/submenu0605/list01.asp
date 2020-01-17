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


<table border="0" cellpadding="0" cellspacing="0" align="center">
	<tr>
		<td>
			<!--<DIV style="OVERFLOW-Y:auto; OVERFLOW-X:auto; MARGIN: 0px 0px 0px 0px; WIDTH:940; HEIGHT:500;">-->
			<table width="940"  border="0" cellpadding="1" cellspacing="1" bgcolor="#EFECE5" align="center">

				<tr bgcolor='#EEF6FF'>
					<td align='center' class='TDCont'  width='150' colspan='3'>구분</td>
<%

					for i = 1 to right(dateadd("d",-1,dateadd("m",1,FromDate)),2)
%>
						<td align='center' class='TDCont'><%=i%></td>
<%
					next 

%>
				</tr>
<%

	If QueryYN = "Y" Then


dim itot(30)
		'생명의전화
		'------------------------------------------------------------------------------------------
		sql = "select datepart(day,jubdate) sday ,count(*) cnt from tb_lifecallhistory where jubdate >='"&FromDate&"' and jubdate <= '"&ToDate&"' and channelgb = 'A'"	'군전화
		sql = sql & "	group by datepart(day,jubdate)"


		set Rs = db.execute(sql)
		
		sString = "<tr ><td bgcolor='FFFFFF' align='center' rowspan='4'><b>생명의전화</b></td><td bgcolor='FFFFFF' align='center'>군전화</d><td bgcolor='FFFFFF' align='center'>연결</td>"
		do until Rs.eof

			for i = 1 to 30
				if i = cint(trim(rs("sday"))) then
					sString = sString&"<td bgcolor='FFFFFF' align='center' width='20'><font color='#000000'>"&rs("cnt")&"</font></td>"
					itot(i) = itot(i) + rs("cnt")
					rs.movenext
					if rs.eof then
						exit for
					end if
				else
					sString = sString&"<td bgcolor='FFFFFF' align='center' width='20'>&nbsp;</td>"
				end if
			next

			if rs.eof then

				if i < 30 then
					for j = i + 1 to 30
						sString = sString&"<td bgcolor='FFFFFF' align='center' width='20'>&nbsp;</td>"
					next
				end if
				exit do
			end if

		loop
		response.write sString



		'생명의전화
		'------------------------------------------------------------------------------------------
		sql = "select datepart(day,jubdate) sday ,count(*) cnt from tb_lifecallhistory where jubdate >='"&FromDate&"' and jubdate <= '"&ToDate&"' and channelgb = 'B'"	'군전화
		sql = sql & "	group by datepart(day,jubdate)"

		set Rs = db.execute(sql)
		
		sString = "<tr ><td bgcolor='FFFFFF' align='center'>일반전화</d><td bgcolor='FFFFFF' align='center'>연결</td>"
		do until Rs.eof

			for i = 1 to 30
				if i = cint(trim(rs("sday"))) then
					sString = sString&"<td bgcolor='FFFFFF' align='center' width='20'><font color='#000000'>"&rs("cnt")&"</font></td>"
					itot(i) = itot(i) + rs("cnt")
					rs.movenext
					if rs.eof then
						exit for
					end if
				else
					sString = sString&"<td bgcolor='FFFFFF' align='center' width='20'>&nbsp;</td>"
				end if
			next

			if rs.eof then

				if i < 30 then
					for j = i + 1 to 30
						sString = sString&"<td bgcolor='FFFFFF' align='center' width='20'>&nbsp;</td>"
					next
				end if
				exit do
			end if

		loop
		response.write sString


		'콜백
		'------------------------------------------------------------------------------------------
		sql = "select datepart(day,requesttime) sday ,count(*) cnt from tb_callback where convert(char(10),requesttime,121) >='"&FromDate&"' and convert(char(10),requesttime,121) <= '"&ToDate&"'"	'군전화
		sql = sql & "	and dnis = 'B' group by datepart(day,requesttime)"

		set Rs = db.execute(sql)
		
		sString = "<tr ><td bgcolor='FFFFFF' align='center'>콜백</d><td bgcolor='FFFFFF' align='center'>연결</td>"
		do until Rs.eof

			for i = 1 to 30
				if i = cint(trim(rs("sday"))) then
					sString = sString&"<td bgcolor='FFFFFF' align='center' width='20'><font color='#000000'>"&rs("cnt")&"</font></td>"
					itot(i) = itot(i) + rs("cnt")
					rs.movenext
					if rs.eof then
						exit for
					end if
				else
					sString = sString&"<td bgcolor='FFFFFF' align='center' width='20'>&nbsp;</td>"
				end if
			next

			if rs.eof then

				if i < 30 then
					for j = i + 1 to 30
						sString = sString&"<td bgcolor='FFFFFF' align='center' width='20'>&nbsp;</td>"
					next
				end if
				exit do
			end if
		loop
		response.write sString

		'계
		sString = "<tr><td bgcolor='#EEF6FF' align='center' colspan='2'>계</td>"
		for i = 1 to 30
			sString = sString&"<td bgcolor='#EEF6FF' align='center' width='20'><font color='#000000'><b>"&itot(i)&"</b></font></td>"
			itot(i) = 0
		next
		response.write sString



		SQL = "SELECT * FROM ARMYCC.DBO.TB_CODE WHERE CODEGROUP = 'Z04' AND CODE >='H' ORDER BY CODE "
		set Rs1 = db.execute(sql)

		DO UNTIL RS1.EOF
		'#####################################################################################################33

				'군범죄신고
				'------------------------------------------------------------------------------------------
				sql = "select datepart(day,jubdate) sday ,count(*) cnt from tb_callhistory where jubdate >='"&FromDate&"' and jubdate <= '"&ToDate&"' and telkind = '"&RS1("CODE")&"'  and channelgb = 'A'"	'군범죄신고
				sql = sql & "	group by datepart(day,jubdate)"


				set Rs = db.execute(sql)

				sString = "<tr ><td bgcolor='FFFFFF' align='center' rowspan='4'><b>"&RS1("CODENAME")&"</b></td><td bgcolor='FFFFFF' align='center'>군전화</d><td bgcolor='FFFFFF' align='center'>연결</td>"

				if Rs.eof then
					for i = 1 to 30
						sString = sString&"<td bgcolor='FFFFFF' align='center' width='20'>&nbsp;</td>"
					next
				end if

				do until Rs.eof

					for i = 1 to 30
						if i = cint(trim(rs("sday"))) then
							sString = sString&"<td bgcolor='FFFFFF' align='center' width='20'><font color='#000000'>"&rs("cnt")&"</font></td>"
							itot(i) = itot(i) + rs("cnt")
							rs.movenext
							if rs.eof then
								exit for
							end if
						else
							sString = sString&"<td bgcolor='FFFFFF' align='center' width='20'>&nbsp;</td>"
						end if
					next

					if rs.eof then

						if i < 30 then
							for j = i + 1 to 30
								sString = sString&"<td bgcolor='FFFFFF' align='center' width='20'>&nbsp;</td>"
							next
						end if
						exit do
					end if

				loop
				response.write sString



				'생명의전화
				'------------------------------------------------------------------------------------------
				sql = "select datepart(day,jubdate) sday ,count(*) cnt from tb_callhistory where jubdate >='"&FromDate&"' and jubdate <= '"&ToDate&"' and telkind = '"&RS1("CODE")&"' and channelgb = 'B'"	'군전화
				sql = sql & "	group by datepart(day,jubdate)"

				set Rs = db.execute(sql)
				
				sString = "<tr ><td bgcolor='FFFFFF' align='center'>일반전화</d><td bgcolor='FFFFFF' align='center'>연결</td>"

				if Rs.eof then
					for i = 1 to 30
						sString = sString&"<td bgcolor='FFFFFF' align='center' width='20'>&nbsp;</td>"
					next
				end if

				do until Rs.eof

					for i = 1 to 30
						if i = cint(trim(rs("sday"))) then
							sString = sString&"<td bgcolor='FFFFFF' align='center' width='20'><font color='#000000'>"&rs("cnt")&"</font></td>"
							itot(i) = itot(i) + rs("cnt")
							rs.movenext
							if rs.eof then
								exit for
							end if
						else
							sString = sString&"<td bgcolor='FFFFFF' align='center' width='20'>&nbsp;</td>"
						end if
					next

					if rs.eof then

						if i < 30 then
							for j = i + 1 to 30
								sString = sString&"<td bgcolor='FFFFFF' align='center' width='20'>&nbsp;</td>"
							next
						end if
						exit do
					end if

				loop
				response.write sString


				'콜백
				'------------------------------------------------------------------------------------------
				sql = "select datepart(day,requesttime) sday ,count(*) cnt from tb_callback where convert(char(10),requesttime,121) >='"&FromDate&"' and convert(char(10),requesttime,121) <= '"&ToDate&"'"	'군전화
				sql = sql & "	and dnis = '"&RS1("CODE")&"' group by datepart(day,requesttime)"

				set Rs = db.execute(sql)
				
				sString = "<tr ><td bgcolor='FFFFFF' align='center'>콜백</d><td bgcolor='FFFFFF' align='center'>연결</td>"

				if Rs.eof then
					for i = 1 to 30
						sString = sString&"<td bgcolor='FFFFFF' align='center' width='20'>&nbsp;</td>"
					next
				end if

				do until Rs.eof

					for i = 1 to 30
						if i = cint(trim(rs("sday"))) then
							sString = sString&"<td bgcolor='FFFFFF' align='center' width='20'><font color='#000000'>"&rs("cnt")&"</font></td>"
							itot(i) = itot(i) + rs("cnt")
							rs.movenext
							if rs.eof then
								exit for
							end if
						else
							sString = sString&"<td bgcolor='FFFFFF' align='center' width='20'>&nbsp;</td>"
						end if
					next

					if rs.eof then

						if i < 30 then
							for j = i + 1 to 30
								sString = sString&"<td bgcolor='FFFFFF' align='center' width='20'>&nbsp;</td>"
							next
						end if
						exit do
					end if
				loop
				response.write sString

				'계
				sString = "<tr><td bgcolor='#EEF6FF' align='center' colspan='2'>계</td>"
				for i = 1 to 30
					sString = sString&"<td bgcolor='#EEF6FF' align='center' width='20'><font color='#000000'><b>"&itot(i)&"</b></font></td>"
					itot(i) = 0
				next
				response.write sString

			RS1.MOVENEXT
		LOOP


	end if
%>

			</table>
		</td>
	</tr>
</table>



<!-- #include virtual="/Include/Bottom.asp" -->