<!-- #include virtual="/Include/Adovbs.inc" -->
<!-- #include virtual="/Include/Common.asp" -->
<%
dim Filename
Filename = "���ڷ�_" & Right(Replace(FormatDateTime(Date,2),"-",""),10) & "_data.xls"

Response.Buffer = True
Response.CacheControl = "public"
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-disposition","attachment;filename="&Filename

dim FromDate, ToDate, QueryYN

QueryYN = request("QueryYN")
FromDate = request("FromDate")
ToDate = request("ToDate")
Kind = request("Kind")
response.write Kind
if FromDate = "" then FromDate = left(Date(),7)&"-01" end If
if ToDate = "" then ToDate=date() end If

dim pageWHERE

pageWHERE = "QueryYN=N&FromDate="&FromDate&"&ToDate="&ToDate

dim oCmd1, oCmd2, iAction, Result1, Result2, prm
Set oCmd1=Server.CreateObject("ADODB.Command")
Set oCmd2=Server.CreateObject("ADODB.Command")

set oCmd1.ActiveConnection = db
oCmd1.CommandText = "armyinformix.dbo.submenu0203_1"
oCmd1.CommandType = adCmdStoredProc

iAction = "1"

set prm = oCmd1.CreateParameter("@iAction",adChar,adParamInput,1,iAction)
oCmd1.Parameters.Append prm
set prm = oCmd1.CreateParameter("@sDate",adChar,adParamInput,10,FromDate)
oCmd1.Parameters.Append prm
set prm = oCmd1.CreateParameter("@eDate",adChar,adParamInput,10,ToDate)
oCmd1.Parameters.Append prm
set prm = oCmd1.CreateParameter("@sKind",adChar,adParamInput,1,Kind)
oCmd1.Parameters.Append prm

set Result1 = oCmd1.Execute

set oCmd2.ActiveConnection = db
oCmd2.CommandText = "armyinformix.dbo.submenu0203_1"
oCmd2.CommandType = adCmdStoredProc

iAction = "3"

set prm = oCmd2.CreateParameter("@iAction",adChar,adParamInput,1,iAction)
oCmd2.Parameters.Append prm
set prm = oCmd2.CreateParameter("@sDate",adChar,adParamInput,10,FromDate)
oCmd2.Parameters.Append prm
set prm = oCmd2.CreateParameter("@eDate",adChar,adParamInput,10,ToDate)
oCmd2.Parameters.Append prm
set prm = oCmd2.CreateParameter("@sKind",adChar,adParamInput,1,Kind)
oCmd2.Parameters.Append prm

set Result2 = oCmd2.Execute

dim TotalCount

TotalCount = CLng(Result2("TotalCount"))

dim Array1, i
redim Array1(TotalCount,10)

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
	cur_nameoffact = "[<a href='##' onClick=""nLink('"&Result1("factnum")&"');"">"&Result1("factnum")&"</a>] "&Result1("nameoffact")

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
			Array1(array_cnt,9) = "����(" & FormatNumber(cdbl(cur_monitorpoint),2) & ")"
		else
			if CDbl(cur_monitorpoint) < 9 and CDbl(cur_monitorpoint) >= 8 then
				Array1(array_cnt,9) = "����(" & FormatNumber(cdbl(cur_monitorpoint),2) & ")"
			else
				if CDbl(cur_monitorpoint) < 8 then
					Array1(array_cnt,9) = "<font color='#ff0000'>"&"�Ҹ���(" & FormatNumber(cdbl(cur_monitorpoint),2) & ")</font>"
				end if
			end if
		end if
	
	else
		if Result1("monitorresult") = "1" then '��ȭ�Ҵ�
			Array1(array_cnt,9) ="<font color='#0000ff'>"&Result1("monitorresult_codename")&"</font>"
		elseif Result1("monitorresult") = "2" then '�����ź�
			Array1(array_cnt,9) ="<font color='#ff00ff'>"&Result1("monitorresult_codename")&"</font>"
		elseif Result1("monitorresult") = "3" then '�̽ǽ�
			Array1(array_cnt,9) ="<font color='#00ffff'>"&Result1("monitorresult_codename")&"</font>"
		else
			Array1(array_cnt,9) ="<font color='#000000'>"&Result1("monitorresult_codename")&"</font>"
		end if
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
<title>:: �������� ������� ����͸� ::</title>
	<META HTTP-EQUIV="Expires" CONTENT="0">
	<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
	<META HTTP-EQUIV="Cache-Control" CONTENT="no-cache">
	<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
</head>

<style type="text/css">
<!--
BODY {scrollbar-face-color: #f7f7f7; scrollbar-shadow-color: #cccccc; scrollbar-highlight-color: #ffffff; scrollbar-3dlight-color: #ffffff; scrollbar-darkshadow-color: #ffffff; scrollbar-track-color: #ffffff;scrollbar-arrow-color: #304A6D; font-size:9pt}

td { font-family: "Verdana","����ü"; font-size:12px; color:#464646; letter-spacing:-1px; line-height:22px;}
-->
</style>

<body bgcolor="#FAFAFA" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">

<table border="0" width="940" cellpadding="0" cellspacing="3" bgcolor="#EFECE5" align="center">
	<tr bgcolor="#FFFFFF">
		<td>
			<table width="100%" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">

			    <tr>
			        <td width="110" bgcolor="#EFEFEF" class="TDCont">��ȸ�Ⱓ :</td>
			        <td  bgcolor="#FFFFFF" colspan=3 width=300><%=FromDate%>~<%=ToDate%>	
			        </td>
				</tr>

			</table>
		</td>
	</tr>
</table>


<table width="940" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
<table border="1" width="940" cellspacing="0" cellpadding="0" bgcolor="#CCCCCC" align="center">
	<tr height="25" bgcolor="#F3F3F3" align="center">
		<td rowspan=2 width=90><b>����͸�<br>����</b></td>
		<td rowspan=2 width="350"><b>��Ǹ�</b></td>
		<td colspan=3><b>�����</b></td>
		<td colspan=2><b>������</b></td>
		<td rowspan=2 width=100><b>������</b></td>
		<td rowspan=2 width=100><b>count</b></td>

	</tr>

	<tr height="25" bgcolor="#F3F3F3" align="center">
		<td width=100><b>�Ҽ�</b></td>
		<td width=80><b>���</b></td>
		<td width=80><b>����</b></td>
		<td width=80><b>����</b></td>
		<td width=80><b>����</b></td>
	</tr>
	
<%for i = 0 to TotalCount-1%>
<tr height="25" bgcolor="#FFFFFF">
<%if CInt(Array1(i,1)) >= 1 then%>
<td align="center" rowspan="<%=Array1(i,1)%>"><b><%=Array1(i,0)%><br>(<%=Array1(i,1)%>)</b></td>
<%end if%>
<%if CInt(Array1(i,3)) >= 1 then%>
<td align="left" rowspan="<%=Array1(i,3)%>">&nbsp;<%=Array1(i,2)%></td>
<td align="center" rowspan="<%=Array1(i,3)%>"><%=Array1(i,4)%></td>
<td align="center" rowspan="<%=Array1(i,3)%>"><%=Array1(i,5)%></td>
<td align="center" rowspan="<%=Array1(i,3)%>"><%=Array1(i,6)%></td>
<%end if%>
<td align="center"><%=Array1(i,7)%></td>
<td align="center"><%=Array1(i,8)%></td>
<td align="center"><%=Array1(i,9)%></td>
<td align="center"><%=i+1%></td>
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