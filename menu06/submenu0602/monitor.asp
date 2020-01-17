<!-- #include virtual="/Include/Top_Frame.asp" -->

<%


	SS_LoginID = SESSION("SS_LoginID")

	SQL = "	select a.idx,sosok_name, class, name, c.cellphone,c.gunphone,processstep,successflag,datediff(second,stepdate,getdate()) as processseconds"
	SQL = SQL & "	from	temp_conference c, TB_SMSADDR a"
	SQL = SQL & "	where	addr_idx = a.idx and userid = '" & SS_LoginID & "' and datagb = '1' order by a.idx"

i = 0
j = 0
k = 0
l = 0
			set RS2 = db.execute(SQL)
%>


<script>

	function fn_end()
	{
		self.close();
	}
	function fn_reset()
	{
		DBFrame.location = "conference_monitor.asp";
		setTimeout("fn_reset()", 2000);
	}
</script>

<table width="700" border="0" cellspacing="0" cellpadding="0" align='center'>
	<tr height="10">
		<td align="left" bgcolor="#FFFFFF" class="TDCont" colspan="8"></td>
	</tr>
</table>
<table width="700" border="0" cellspacing="0" cellpadding="0" align='center'>
	<tr height="40">
		<td align="center" bgcolor="#FFFFFF" class="TDCont" colspan="8">&nbsp;<font color="#0000ff" size="5px">다자간통화 모니터링</font> </b></td>
	</tr>
</table>
<form method="post" name="inUpFrm" style="margin:0">
<table border="0" width="700" cellpadding="0" cellspacing="1" bgcolor="#EFECE5" align="center">
	<tr bgcolor="#FFFFFF">
		<td>		
			<form method="post" name="inUpFrm" style="margin:0">
			<input type="hidden" name="QueryYN" value="">
			<table width="100%" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">

			    <tr height="50">
						<td bgcolor="#EEF6FF" class="TDCont" width=80 align="center"><b><font size="5px">총인원</font></b></td>
						<td bgcolor="#FFFFFF" colspan=4 align="center"><b><input type="text" name="cnt1" value="" size=3  style="border-width:0px ; border-color:#EEF6FF ; border-style:solid; text-align:center; font-color:#ff0000;font-size:20px;font-weight:bold" readonly ></b>		
						</td>
						<td bgcolor="#EEF6FF" class="TDCont" width=80 align="center"><b><font size="5px">성공</font></b></td>
						<td bgcolor="#FFFFFF" colspan=4 align="center"><b><input type="text" name="cnt2" value="" size=3  style="border-width:0px ; border-color:#EEF6FF ; border-style:solid; text-align:center; font-color:#ff0000;font-size:20px;font-weight:bold" readonly></b>			
						</td>
						<td bgcolor="#EEF6FF" class="TDCont" width=80 align="center"><b><font size="5px">진행중</font></b></td>
						<td bgcolor="#FFFFFF" colspan=4 align="center"><b><input type="text" name="cnt3" value="" size=3  style="border-width:0px ; border-color:#EEF6FF ; border-style:solid; text-align:center; font-color:#ff0000;font-size:20px;font-weight:bold" readonly ></b>		
						</td>
						<td bgcolor="#EEF6FF" class="TDCont" width=80 align="center"><b><font size="5px">실패</font></b></td>
						<td bgcolor="#FFFFFF" colspan=4 align="center"><b><input type="text" name="cnt4" value="" size=3  style="border-width:0px ; border-color:#EEF6FF ; border-style:solid; text-align:center; font-color:#0000ff;font-size:20px;font-weight:bold" readonly ></b>		
						</td>
				</tr>
			</table>
			</form>
		</td>
	</tr>
</table>
<table width="700" height="10" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
<table border="0" width="700" cellpadding="0" cellspacing="1" bgcolor="#EFECE5" align="center">
	<tr bgcolor="#FFFFFF">
		<td>		
        	<table width="100%" border="0" cellspacing="1" cellpadding="1" bgcolor="#CCCCCC">
        		<tr height="30" bgcolor="#EEF6FF" align="center">
        			<td class="TDCont" width=30 align='center'>NO</td>
        			<td class="TDCont" width=100 align='center' nowrap>소속</td>
        			<td class="TDCont" width=100 align='center'>계급</td>
        			<td class="TDCont" width=100 align='center'>성명</td>
        			<td class="TDCont" width=80 align='center'>군전화</td>
        			<td class="TDCont" width=100 align='center'>휴대폰</td>
        			<td class="TDCont" width=80 align='center'>결과</td>
        			<td class="TDCont" width=80 align='center'>단계</td>
        			<td class="TDCont" width=80 align='center'>진행시간</td>
        		</tr>
        		<tr><td colspan="10" height="1" bgcolor="#FFFFFF"></td></tr>
<%
			do until RS2.eof

				i = i + 1
				idx = RS2("idx")
				sosok_name = RS2("sosok_name")
				sclass = RS2("class")
				sname = RS2("name")
				cellphone = RS2("cellphone")
				gunphone = RS2("gunphone")
				processstep = RS2("processstep")
				processseconds = RS2("processseconds")
				successflag = RS2("successflag")
%>
        		<tr bgcolor="#fffff" align="center" height="30">
        			<td><%=i%></td>
        			<td><%=sosok_name%></td>
        			<td><%=sclass%></td>
        			<td class="TDCont" nowrap><font size="3px"><%=sname%></font></td>
        			<td><%=gunphone%></td>
        			<td nowrap><%=FormatHPNo(cellphone)%></td>

					<%
						if successflag = "0" then
							successflagname = "대기"		
							processstep = ""
						elseif successflag = "1" then		'성공	
							successflagname = "성공"
							j = j + 1	
						elseif successflag = "2" then		'진행중		
							successflagname = "진행중"						
							k = k + 1
						else				'실패
							successflagname = "실패"
							l = l + 1			
							processseconds = ""
						end if						
					%>

        			<td><input type="hidden" name="successflag_<%=idx%>" value="<%=successflagname%>" size=10 style="border-width:0px ; border-color:#cccccc ; border-style:solid; text-align:center" readonly><span id="panresult_<%=idx%>"></span></td>
        			<td><input type="text" name="result_<%=idx%>" value="<%=processstep%>" size=10 style="border-width:0px ; border-color:#cccccc ; border-style:solid; text-align:center" readonly></td>
        			<td><input type="text" name="time_<%=idx%>" value="<%=processseconds%>" size=5 style="border-width:0px ; border-color:#cccccc ; border-style:solid; text-align:center" readonly></td>

        		</tr>
<%
				RS2.movenext

			loop

%>

        	</table>       	

		</td>

	</tr>
</table>
</form>
<table width="700" border="0" cellspacing="0" cellpadding="0" align='center'>
	<tr height="10">
		<td align="left" bgcolor="#FFFFFF" class="TDCont" colspan="8"></td>
	</tr>
</table>

 <table align='center' width="700" border="0" cellspacing="1" cellpadding="1" bgcolor="#FFFFFF">
 <tr><td bgcolor="#FFFFFF" align='center'>
		<img src="/Images/Btn/BtnClose.GIF" style="cursor:hand;" align="absmiddle" onclick="fn_end();">
		<!--<img src="/Images/Btn/BtnCallConference2.GIF" style="cursor:hand;" align="absmiddle" onclick="fn_reset();">--></td>
</tr></table>
<iframe src="about:blank" name="DBFrame" width="0" height="0" frameborder=0 marginheight=0 marginwidth=0 scrolling="no"></iframe>
<script>
	document.all.cnt1.value = "<%=i%>";	// 총건수
	document.all.cnt2.value = "<%=j%>";	// 총건수
	document.all.cnt3.value = "<%=k%>";	// 총건수
	document.all.cnt4.value = "<%=l%>";	// 총건수
	fn_reset();
</script>

<!-- #include virtual="/Include/Bottom.asp" -->


