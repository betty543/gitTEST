<!-- #include virtual="/Include/Common.asp" -->
<%
Server.ScriptTimeout = 90000
Response.ContentType = "application/vnd.ms-excel; name='My_Excel'"
Call Response.AddHeader("Content-Disposition", "attachment; filename=HelpCall���������" &Date()& ".xls")	'�ٷ������ϱ�
Call Response.AddHeader("Content-Description", "ASP Generated Data")

gb = request("gb")
FromDate = request("FromDate")
ToDate = request("ToDate")
%>

<html>
<head>
</head>
<body>
	
<table border="1">
	<tr>
		<td align="center" rowspan="3">���</td>
		<td align="center" rowspan="3">�Ѱ�</td>
		<td align="center" colspan="9">������ȭ(1303)</td>
		<td align="center" rowspan="3">��������ȭ(080)</td>
	</tr>
	<tr>
		<td align="center" rowspan="2">�Ұ�</td>
		<td align="center" rowspan="2">��������ȭ</td>
		<td align="center" rowspan="2">�����˽Ű�</td>
		<td align="center" colspan="6">�����˽Ű�</td>
	</tr>
	<tr>
		<td align="center">�Ұ�</td>
		<td align="center">����� �� �����δ�</td>
		<td align="center">����</td>
		<td align="center">�ر�</td>
		<td align="center">����</td>
		<td align="center">�غ�</td>
	</tr>
	
	<%
	sql = " select bound_ymd "
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
	
	sql = sql & " group by bound_ymd "
	sql = sql & " order by bound_ymd "
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
		
		<tr>
			<td align="center"><%=arrRs(0,i)%></td>
			<%
			for j = 1 to 11
				arrSum(j) = arrSum(j) + arrRs(j,i)
				%><td align="right"><%=formatnumber(arrRs(j,i),0)%>&nbsp;</td><%
			next
			%>
		</tr>
		<%
	next
	%>
	<tr>
		<td align="center">�հ�</td>
		<%
		for j = 1 to 11
			%><td align="right"><b><%=formatnumber(arrSum(j),0)%></b>&nbsp;</td><%
		next
		%>
	</tr>
	
</table>

</body>
</html>