<!-- #include virtual="/Include/Top.asp" -->

<%
QueryYN = request("QueryYN")

gb = request("gb") : if gb = "" then gb = "A" end if
FromDate = request("FromDate")
ToDate = request("ToDate")

if FromDate = "" then FromDate =left(Date(),7)&"-01" end If
if ToDate = "" then ToDate=date() end if
%>

<script>
	function fn_Search() {
		document.inUpFrm.QueryYN.value = "Y";
		document.inUpFrm.submit();
	}

	function fn_Xls() {
		location.href="state_month_xls.asp?gb=<%=gb%>&FromDate=<%=FromDate%>&ToDate=<%=ToDate%>"
	}
</script>

<div id='CalendarLayer' style='display:none;'><iframe name='CalendarFrame' src='/include/Calendar.asp' width='172' height='177' border='0' frameborder='0' scrolling='no'></iframe></div>

<form name="inUpFrm" method="post" action="" target="">
	<input type="hidden" name="QueryYN" value="<%=QueryYN%>">
		
	<table border="0" width="1200" cellpadding="0" cellspacing="3" bgcolor="#EFECE5" align="center">
		<tr bgcolor="#FFFFFF">
			<td>
				
				<table width="100%" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">
					<tr>
						<td class="TDCont" bgcolor="#EFEFEF" align="center">구분</td>
						<td bgcolor="#FFFFFF">
							<input type="radio" name="gb" value="A" <% if gb = "A" then %>checked<% end if %> /> 전체
							<input type="radio" name="gb" value="B" <% if gb = "B" then %>checked<% end if %> /> 군전화
							<input type="radio" name="gb" value="C" <% if gb = "C" then %>checked<% end if %> /> 일반전화
						</td>
						<td class="TDCont" bgcolor="#EFEFEF" align="center">조회기간</td>
						<td bgcolor="#FFFFFF">
							<input value="<%=FromDate%>" name="FromDate" readonly type="text" size="10" onfocus="setFocusColor(this); new CalendarFrame.Calendar(this);" onblur="setOutColor(this);" />
							~
							<input value="<%=ToDate%>" name="ToDate" readonly type="text" size="10" onfocus="setFocusColor(this); new CalendarFrame.Calendar(this);" onblur="setOutColor(this);">
						</td>
						<td bgcolor="#FFFFFF" align="center">
							<img src="/Images/Btn/BtnSearch.gif" style="cursor:hand;" onClick="fn_Search();">
							<img src="/Images/Btn/BtnExcel.gif" style="cursor:hand;" onClick="fn_Xls();">
						</td>
					</tr>
				</table>
				
			</td>
		</tr>
	</table>

</form>

<table border="0" width="1200" cellspacing="1" cellpadding="0" width="100%" bgcolor="#CCCCCC" align="center">
	<colgroup>
		<col width="100px" />
		<col width="100px" />
		<col width="100px" />
		<col width="100px" />
		<col width="100px" />
		
		<col width="100px" />
		<col width="100px" />
		<col width="100px" />
		<col width="100px" />
		<col width="100px" />
		<col width="100px" />
		
		<col width="100px" />
	</colgroup>
	<tr height="25" bgcolor="#EEF6FF">
		<td class="TDCont" align="center" rowspan="3">년월</td>
		<td class="TDCont" align="center" rowspan="3">총계</td>
		<td class="TDCont" align="center" colspan="9">통합전화(1303)</td>
		<td class="TDCont" bgcolor="#FDE6F3" align="center" rowspan="3">생명의전화<br />(080)</td>
	</tr>
	<tr height="25" bgcolor="#EEF6FF">
		<td class="TDCont" align="center" rowspan="2">소계</td>
		<td class="TDCont" align="center" rowspan="2">생명의전화</td>
		<td class="TDCont" align="center" rowspan="2">성범죄신고</td>
		<td class="TDCont" align="center" colspan="6">군범죄신고</td>
	</tr>
	<tr height="25" bgcolor="#EEF6FF">
		<td class="TDCont" align="center">소계</td>
		<td class="TDCont" align="center">국방부 및 국직부대</td>
		<td class="TDCont" align="center">육군</td>
		<td class="TDCont" align="center">해군</td>
		<td class="TDCont" align="center">공군</td>
		<td class="TDCont" align="center">해병</td>
	</tr>
	
	<%
	sql = " select left(bound_ymd,7) "
	sql = sql & " 	, count(*) as sum1 "
	sql = sql & " 	, count(case when dtmf <> '00' then 1 else null end) as sum2 "
	sql = sql & " 	, count(case when dtmf = '10' then 1 else null end) as sum10 "
	sql = sql & " 	, count(case when dtmf = '20' then 1 else null end) as sum20 "
	sql = sql & " 	, count(case when left(dtmf,1) = 3 then 1 else null end) as sum3 "
	sql = sql & " 	, count(case when dtmf = '31' then 1 else null end) as sum31 "
	sql = sql & " 	, count(case when dtmf = '32' then 1 else null end) as sum32 "
	sql = sql & " 	, count(case when dtmf = '33' then 1 else null end) as sum33 "
	sql = sql & " 	, count(case when dtmf = '34' then 1 else null end) as sum34 "
	sql = sql & " 	, count(case when dtmf = '35' then 1 else null end) as sum35 "
	sql = sql & " 	, count(case when dtmf = '00' then 1 else null end) as sum00 "
	sql = sql & " from tb_bound with(nolock) "
	sql = sql & " where dtmf in ('00','10','20','31','32','33','34','35') "
	sql = sql & " 	and bound_ymd between '" & FromDate & "' and '" & ToDate & "' "
	
	if gb = "B" then
		sql = sql & " 	and left(bound_dnis,1) = '5' "
	elseif gb = "C" then
		sql = sql & " 	and (left(bound_dnis,1) = '6' or left(bound_dnis,1) = '1') "
	end if
	
	sql = sql & " group by left(bound_ymd,7) "
	sql = sql & " order by left(bound_ymd,7) "
	'response.write	sql
	set rs = db.execute(sql)
	if not rs.eof then
		arrRs = rs.getRows
		arrRc = ubound(arrRs,2)
	else
		arrRc = -1
	end if
	rs.close
	set rs = nothing
	
	dim arrSum(11)
	
	for i = 0 to arrRc
		%>
		
		<tr bgcolor="#FFFFFF">
			<td align="center"><%=arrRs(0,i)%></td>
			<%
			for j = 1 to 11
				arrSum(j) = arrSum(j) + arrRs(j,i)
				%><td <% if j = 1 or j = 2 or j = 5 then %>class="TDCont" bgcolor="#EEF6FF"<% end if %> align="right"><%=formatnumber(arrRs(j,i),0)%>&nbsp;</td><%
			next
			%>
		</tr>
		<%
	next
	%>
	<tr bgcolor="#FFEEF9">
		<td class="TDCont" align="center">합계</td>
		<%
		for j = 1 to 11
			%><td align="right"><b><%=formatnumber(arrSum(j),0)%></b>&nbsp;</td><%
		next
		%>
	</tr>
	
</table>

<!-- #include virtual="/Include/Bottom.asp" -->