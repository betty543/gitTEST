<!-- #include virtual="/Include/Adovbs.inc" -->
<!-- #include virtual="/Include/Common.asp" -->
<!-- #include virtual="/Include/use_func.asp" -->
<%
dim FromDate, ToDate, QueryYN

QueryYN = request("QueryYN")
FromDate = request("FromDate")
ToDate = request("ToDate")

if FromDate = "" then FromDate = left(Date(),7)&"-01" end If
if ToDate = "" then ToDate=date() end If

dim pageWHERE

pageWHERE = "QueryYN=N&FromDate="&FromDate&"&ToDate="&ToDate

	dim EXCEL_CHK, Table_width_and_border, mark_code1, mark_code2
	EXCEL_CHK = "Y"
	Table_width_and_border = "width='940' border='1'"
	mark_code1 = "["
	mark_code2 = "]"

dim Filename
Filename = "부대별_" & Right(Replace(FormatDateTime(Date,2),"-",""),10) & "_data.xls"

Response.Buffer = True
Response.CacheControl = "public"
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition","attachment;filename="&Filename


dim oCmd1, oCmd2, iAction, Result1, Result2, prm
Set oCmd1=Server.CreateObject("ADODB.Command")
Set oCmd2=Server.CreateObject("ADODB.Command")

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

<table <%=Table_width_and_border%> cellspacing="1" cellpadding="0" width="100%" bgcolor="#CCCCCC" align="center">
	<tr height="30">
		<td align="center" bgcolor="#FFFFFF" class="TDCont" colspan="11">&nbsp;<b><font color="#ff00ff"></font> 부대별/응답자별현황</b></td>
	</tr>
</table>

<table <%=Table_width_and_border%> cellspacing="1" cellpadding="0" width="100%" bgcolor="#CCCCCC" align="center">
	<tr height="20">
		<td align="left" bgcolor="#FFFFFF" class="TDCont" colspan="11">&nbsp;<b><font color="#ff00ff"></font> 기간:</b>&nbsp;<%=FromDate%>부터 <%=ToDate%>까지</td>
	</tr>
</table>


<!--부대별/응답자별현황 시작 -->
<!-- #include file ="./list08_1.asp" -->
<!--부대별/응답자별현황 끝 -->

<%
set prm = nothing
%>

<table width="940" height="10" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
</body>
</html>
