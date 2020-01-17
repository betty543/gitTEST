<!-- #include virtual="/Include/Adovbs.inc" -->
<!-- #include virtual="/Include/Common.asp" -->

<%
	'####### 파라미터 ##################################################################################
	QueryYN = request("QueryYN")
	FromDate = request("FromDate")
	ToDate = request("ToDate")

	if FromDate = "" then FromDate =left(Date(),7)&"-01" end If
	if ToDate = "" then ToDate=date() end If

	pageWHERE = "QueryYN="&QueryYN&"&FromDate="&FromDate&"&ToDate="&ToDate

	dim EXCEL_CHK, Table_width_and_border, mark_code1, mark_code2
	EXCEL_CHK = "Y"
	Table_width_and_border = "border='1'"
	mark_code1 = "["
	mark_code2 = "]"	

dim Filename
Filename = "종합평가_" & Right(Replace(FormatDateTime(Date,2),"-",""),10) & "_data.xls"

Response.Buffer = True
Response.CacheControl = "public"
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition","attachment;filename="&Filename
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
		<td align="center" bgcolor="#FFFFFF" class="TDCont" colspan="11">&nbsp;<b><font color="#ff00ff"></font> 종합현황</b></td>
	</tr>
</table>

<table <%=Table_width_and_border%> cellspacing="1" cellpadding="0" width="100%" bgcolor="#CCCCCC" align="center">
	<tr height="20">
		<td align="left" bgcolor="#FFFFFF" class="TDCont" colspan="11">&nbsp;<b><font color="#ff00ff"></font> 기간:</b>&nbsp;<%=FromDate%>부터 <%=ToDate%>까지</td>
	</tr>
</table>
<table border="0" cellpadding="0" cellspacing="0" align="center">
	<tr>
		<td>

<%
dim iAction
dim oCmd1, oCmd2, oCmd21, oCmd22, oCmd3, oCmd4, oCmd5, oCmd51, oCmd6, oCmd7, oCmd8
dim Result1, Result2, Result3, Result4, Result5, Result51, Result6, Result7, Result8

Set oCmd1=Server.CreateObject("ADODB.Command")
Set oCmd2=Server.CreateObject("ADODB.Command")
Set oCmd21=Server.CreateObject("ADODB.Command")
Set oCmd22=Server.CreateObject("ADODB.Command")
Set oCmd3=Server.CreateObject("ADODB.Command")
Set oCmd4=Server.CreateObject("ADODB.Command")
Set oCmd5=Server.CreateObject("ADODB.Command")
Set oCmd51=Server.CreateObject("ADODB.Command")
Set oCmd6=Server.CreateObject("ADODB.Command")
Set oCmd7=Server.CreateObject("ADODB.Command")
Set oCmd8=Server.CreateObject("ADODB.Command")

dim ArrayValue1, ArrayValue2, ArrayValue3, ArrayValue4, ArrayValue5
redim ArrayValue1(20), ArrayValue2(20), ArrayValue3(20), ArrayValue4(20), ArrayValue5(20)

dim i, j, count_sum1, count_sum2
%>
			
<!--부대별 시작 -->
<!-- #include file ="./list04_1.asp" -->
<!--부대별 끝 -->

<!--유형별 시작 -->
<!-- #include file ="./list04_2.asp" -->
<!--유형별 끝 -->

<!--사건관계자 시작 -->
<!-- #include file ="./list04_3.asp" -->
<!--사건관계자 끝 -->

<!--불만족현황(총괄) 시작 -->
<!-- #include file ="./list04_7.asp" -->
<!--불만족현황(총괄) 끝 -->

<!--불만족현황(소속) 시작 -->
<!-- #include file ="./list04_4.asp" -->
<!--불만족현황(소속) 끝 -->

<!--불만족현황(유형) 시작 -->
<!-- #include file ="./list04_5.asp" -->
<!--불만족현황(유형) 끝 -->

<!--불만족현황(관계자별) 시작 -->
<!-- #include file ="./list04_6.asp" -->
<!--불만족현황(관계자별) 끝 -->

<%
set prm=nothing
%>			

		</td>
	</tr>
</table>


<table width="940" height="10" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
</body>
</html>