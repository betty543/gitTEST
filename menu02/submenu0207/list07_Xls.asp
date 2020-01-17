<!-- #include virtual="/Include/Adovbs.inc" -->
<!-- #include virtual="/Include/Common.asp" -->
<%
dim Filename
Filename = "종합평가내역출력_" & Right(Replace(FormatDateTime(Date,2),"-",""),10) & "_data.xls"

Response.Buffer = True
Response.CacheControl = "public"
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-disposition","attachment;filename="&Filename

dim FromDate, ToDate

FromDate = request("FromDate")
ToDate = request("ToDate")

dim oCmd1, oCmd2, iAction, Result1, Result2, prm
Set oCmd1=Server.CreateObject("ADODB.Command")
Set oCmd2=Server.CreateObject("ADODB.Command")

set oCmd1.ActiveConnection = db
oCmd1.CommandText = "armyinformix.dbo.submenu0207"
oCmd1.CommandType = adCmdStoredProc

iAction = "1"

set prm = oCmd1.CreateParameter("@iAction",adChar,adParamInput,1,iAction)
oCmd1.Parameters.Append prm
set prm = oCmd1.CreateParameter("@sDate",adChar,adParamInput,10,FromDate)
oCmd1.Parameters.Append prm
set prm = oCmd1.CreateParameter("@eDate",adChar,adParamInput,10,ToDate)
oCmd1.Parameters.Append prm

set Result1 = oCmd1.Execute

set oCmd2.ActiveConnection = db
oCmd2.CommandText = "armyinformix.dbo.submenu0207"
oCmd2.CommandType = adCmdStoredProc

iAction = "3"

set prm = oCmd2.CreateParameter("@iAction",adChar,adParamInput,1,iAction)
oCmd2.Parameters.Append prm
set prm = oCmd2.CreateParameter("@sDate",adChar,adParamInput,10,FromDate)
oCmd2.Parameters.Append prm
set prm = oCmd2.CreateParameter("@eDate",adChar,adParamInput,10,ToDate)
oCmd2.Parameters.Append prm

set Result2 = oCmd2.Execute

dim TotalCount

TotalCount = CLng(Result2("TotalCount"))

dim Array1, i
redim Array1(TotalCount,11)

dim pre_receiptfactnum, cur_receiptfactnum
dim receiptfactnum_num, array_cnt

pre_receiptfactnum = ""
cur_receiptfactnum = ""
receiptfactnum_num = 0

array_cnt = 0

Do while not Result1.EOF

	cur_receiptfactnum = Result1("receiptfactnum")

	Array1(array_cnt,0) = cur_receiptfactnum
	Array1(array_cnt,1) = 0
	Array1(array_cnt,2) = Result1("sosok")
	Array1(array_cnt,3) = Result1("class")
	Array1(array_cnt,4) = Result1("susakwanname")
	Array1(array_cnt,5) = Result1("contents")
	Array1(array_cnt,6) = Result1("codename")
	Array1(array_cnt,7) = Result1("peoplename")
	Array1(array_cnt,8) = Result1("remark")
	Array1(array_cnt,9) = 0	
	Array1(array_cnt,10) = replace(Result1("processdate_new"),"/","-")
	
	if cur_receiptfactnum <> pre_receiptfactnum then
		receiptfactnum_num = array_cnt
		Array1(CInt(receiptfactnum_num),1) = CInt(Array1(CInt(receiptfactnum_num),1)) + 1
		Array1(CInt(array_cnt),9) = CInt(Array1(CInt(array_cnt),9)) + 1
	else
		Array1(CInt(receiptfactnum_num),1) = CInt(Array1(CInt(receiptfactnum_num),1)) + 1
		Array1(CInt(array_cnt),9) = CInt(Array1(CInt(array_cnt-1),9)) + 1
	end if

	pre_receiptfactnum = cur_receiptfactnum
	array_cnt = array_cnt + 1
Result1.MoveNext
Loop
%>

<html>
<head>
<title>:: 육군본부 수사장비 모니터링 ::</title>
	<META HTTP-EQUIV="Expires" CONTENT="0">
	<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
	<META HTTP-EQUIV="Cache-Control" CONTENT="no-cache">
	<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
</head>

<style type="text/css">
<!--
BODY {scrollbar-face-color: #f7f7f7; scrollbar-shadow-color: #cccccc; scrollbar-highlight-color: #ffffff; scrollbar-3dlight-color: #ffffff; scrollbar-darkshadow-color: #ffffff; scrollbar-track-color: #ffffff;scrollbar-arrow-color: #304A6D; font-size:9pt}

td { font-family: "Verdana","굴림체"; font-size:12px; color:#464646; letter-spacing:-1px; line-height:22px;}
-->
</style>

<body bgcolor="#FAFAFA" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="940" height="10" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td align="center" colspan="4"><font size ="5"><b>종합평가내역</b></font></td></tr></table>
<table width="940" height="5" border="0" cellspacing="0" cellpadding="0" align="left"><tr><td>기간: <%=FromDate%> 부터 <%=ToDate%> 까지</td></tr></table>
<%
dim temp_total
temp_total = 0

for i = 0 to TotalCount-1
%>
	<%
	if CInt(Array1(i,1)) >= 1 then
		icount1 = icount1 + 1
		temp_total = Array1(i,1)	
	%>
<table width="940" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
<table width="940" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td><font size='2' color='#0000ff'><b>No: <%=icount1%> </b></font></td></tr></table>
<table border="1" width="940" cellspacing="0" cellpadding="0" width="100%" bgcolor="#CCCCCC" align="center">
<tr height="25" align="center">
<td width="150" bgcolor="#F3F3F3"><b>사건번호</b></td>
<td width="320" bgcolor="#FFFFFF"><b><%=Array1(i,0)%></b></td>
<td width="150" bgcolor="#F3F3F3"><b>모니터링 날짜</b></td>
<td width="320" bgcolor="#FFFFFF"><b><%=Array1(i,10)%></b></td>
</tr>
<tr height="25">
<td bgcolor="#F3F3F3" align="center"><b>담당수사관</b></td>
<td bgcolor="#FFFFFF" colspan="3">&nbsp;&nbsp;<b>[ 소속&nbsp;:&nbsp;</b><%=Array1(i,2)%><b>]</b> &nbsp;&nbsp;&nbsp; <b>[ 계급&nbsp;:&nbsp;</b><%=Array1(i,3)%><b>]</b> &nbsp;&nbsp;&nbsp; <b>[ 성명&nbsp;:&nbsp;</b><%=Array1(i,4)%><b>]</b></td>
</tr>
<tr height="25">
<td bgcolor="#F3F3F3" align="center"><b>사건개요</b></td>
<td bgcolor="#FFFFFF" colspan="3">&nbsp;&nbsp;<%=Array1(i,5)%></td>
</tr>
<tr height="25">
<td bgcolor="#F3F3F3" align="center" rowspan="<%=Array1(i,1)%>"><b>종합평가<br>(특이사항)</b></td>
<td bgcolor="#FFFFFF" colspan="3">&nbsp;&nbsp;<b><%=Array1(i,6)%>(<%=Array1(i,7)%>)&nbsp;:&nbsp;</b><%=Array1(i,8)%></td>
</tr>
	<%
	end if
	%>
	<%
	if CInt(Array1(i,9)) > 1 then
	%>
<tr height="25">
<td bgcolor="#FFFFFF" colspan="3">&nbsp;&nbsp;<b><%=Array1(i,6)%>(<%=Array1(i,7)%>)&nbsp;:&nbsp;</b><%=Array1(i,8)%></td>
</tr>
	<%
	end if
	%>
	<%
	if CInt(temp_total) = CInt(Array1(i,9)) then
	%>
</table>
<table width="940" height="20" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
	<%
	end if
	%>
<%Next%>

<%
set oCmd1.ActiveConnection = nothing
set oCmd2.ActiveConnection = nothing
set oCmd1 = nothing
set oCmd2 = nothing
set Result1 = nothing
set Result2 = nothing
set prm = nothing
%>

<!-- #include virtual="/Include/Bottom.asp" -->