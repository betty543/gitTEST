<!-- #include virtual="/Include/Adovbs.inc" -->
<!-- #include virtual="/Include/Common.asp" -->
<%
dim Filename
Filename = "기간별_" & Right(Replace(FormatDateTime(Date,2),"-",""),10) & "_data.xls"

Response.Buffer = True
Response.CacheControl = "public"
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-disposition","attachment;filename="&Filename

dim FromDate, ToDate

	dim EXCEL_CHK, Table_width_and_border, mark_code1, mark_code2
	EXCEL_CHK = "Y"
	Table_width_and_border = "border='1'"
	mark_code1 = "["
	mark_code2 = "]"	

FromDate = request("FromDate")
ToDate = request("ToDate")

dim oCmd1, oCmd2, iAction, Result1, Result2, prm
Set oCmd1=Server.CreateObject("ADODB.Command")
Set oCmd2=Server.CreateObject("ADODB.Command")

set oCmd1.ActiveConnection = db
oCmd1.CommandText = "armyinformix.dbo.submenu0203"
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
oCmd2.CommandText = "armyinformix.dbo.submenu0203"
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
redim Array1(TotalCount,9)

dim pre_processdate_new, cur_processdate_new, pre_nameoffact, cur_nameoffact
dim processdate_num, nameoffact_num, array_cnt, cur_monitorpoint

pre_processdate_new = ""
cur_processdate_new = ""
pre_nameoffact = ""
cur_nameoffact = ""
processdate_num = 0
nameoffact_num = 0

array_cnt = 0

Do while not Result1.EOF

	cur_processdate_new = Result1("processdate_new")
	cur_nameoffact = Result1("nameoffact")
	cur_monitorpoint = Result1("monitorpoint")

	Array1(array_cnt,0) = replace(cur_processdate_new,"/","-")
	Array1(array_cnt,1) = 0
	Array1(array_cnt,2) = cur_nameoffact
	Array1(array_cnt,3) = 0
	Array1(array_cnt,4) = Result1("sosok")
	Array1(array_cnt,5) = Result1("class")
	Array1(array_cnt,6) = Result1("name")
	Array1(array_cnt,7) = Result1("codename")
	Array1(array_cnt,8) = Result1("rename")
	
	
	if Result1("monitorresult") = "9" then
	
		if CDbl(cur_monitorpoint) >= 9 then
			Array1(array_cnt,9) = "만족(" & FormatNumber(cdbl(cur_monitorpoint),2) & ")"
		else
			if CDbl(cur_monitorpoint) < 9 and CDbl(cur_monitorpoint) >= 8 then
				Array1(array_cnt,9) = "보통(" & FormatNumber(cdbl(cur_monitorpoint),2) & ")"
			else
				if CDbl(cur_monitorpoint) < 8 then
					Array1(array_cnt,9) = "불만족(" & FormatNumber(cdbl(cur_monitorpoint),2) & ")"
				end if
			end if
		end if
	
	else
	
	Array1(array_cnt,9) = Result1("monitorresult_codename")
		
	end if
	
	if cur_processdate_new <> pre_processdate_new then
		processdate_num = array_cnt
		Array1(CInt(processdate_num),1) = CInt(Array1(CInt(processdate_num),1)) + 1
	else
		Array1(CInt(processdate_num),1) = CInt(Array1(CInt(processdate_num),1)) + 1
	end if

	if cur_nameoffact <> pre_nameoffact then
		nameoffact_num = array_cnt
		Array1(CInt(nameoffact_num),3) = CInt(Array1(CInt(nameoffact_num),3)) + 1
	else
		Array1(CInt(nameoffact_num),3) = CInt(Array1(CInt(nameoffact_num),3)) + 1
	end if

	pre_processdate_new = cur_processdate_new
	pre_nameoffact = cur_nameoffact
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

<table width="940" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
<table <%=Table_width_and_border%> cellspacing="1" cellpadding="0" width="100%" bgcolor="#CCCCCC" align="center">
	<tr height="30">
		<td align="center" bgcolor="#FFFFFF" class="TDCont" colspan="8">&nbsp;<b><font color="#ff00ff"></font> 기간별</b></td>
	</tr>
</table>

<table <%=Table_width_and_border%> cellspacing="1" cellpadding="0" width="100%" bgcolor="#CCCCCC" align="center">
	<tr height="30">
		<td align="left" bgcolor="#FFFFFF" class="TDCont" colspan="8">&nbsp;<b><font color="#ff00ff"></font> 기간:</b>&nbsp;<%=FromDate%>부터 <%=ToDate%>까지</td>
	</tr>
</table>

<table border="1" width="940" cellspacing="0" cellpadding="0" bgcolor="#CCCCCC" align="center">
	<tr height="25" bgcolor="#F3F3F3" align="center">
		<td rowspan=2 width=100><b>모니터링<br>일자</b></td>
		<td rowspan=2><b>사건명</b></td>
		<td colspan=3><b>수사관</b></td>
		<td colspan=2><b>응답자</b></td>
		<td rowspan=2 width=100><b>만족도</b></td>

	</tr>

	<tr height="25" bgcolor="#F3F3F3" align="center">
		<td width=100><b>소속</b></td>
		<td width=100><b>계급</b></td>
		<td width=100><b>성명</b></td>

		<td width=100><b>구분</b></td>
		<td width=100><b>성명</b></td>
	</tr>
	
<%for i = 0 to TotalCount-1%>
<tr height="25" bgcolor="#FFFFFF">
<%if CInt(Array1(i,1)) >= 1 then%>
<td align="center" rowspan="<%=Array1(i,1)%>"><b><%=Array1(i,0)%><br>[<%=Array1(i,1)%>]</b></td>
<%end if%>
<%if CInt(Array1(i,3)) >= 1 then%>
<td align="center" rowspan="<%=Array1(i,3)%>"><b><%=Array1(i,2)%></b></td>
<td align="center" rowspan="<%=Array1(i,3)%>"><b><%=Array1(i,4)%></b></td>
<td align="center" rowspan="<%=Array1(i,3)%>"><%=Array1(i,5)%></td>
<td align="center" rowspan="<%=Array1(i,3)%>"><%=Array1(i,6)%></td>
<%end if%>
<td align="center"><%=Array1(i,7)%></td>
<td align="center"><%=Array1(i,8)%></td>
<td align="center"><%=Array1(i,9)%></td>
</tr>
<%Next%>

</table>

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