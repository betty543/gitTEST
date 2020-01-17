<!-- #include virtual="/Include/Adovbs.inc" -->
<!-- #include virtual="/Include/Top.asp" -->
<!-- #include virtual="/Include/PopLayer.asp" -->
<%
dim FromDate, ToDate, QueryYN

on Error Resume next

QueryYN = request("QueryYN")
FromDate = request("FromDate")
ToDate = request("ToDate")

if FromDate = "" then FromDate = left(Date(),7)&"-01" end If
if ToDate = "" then ToDate=date() end If

dim pageWHERE

pageWHERE = "QueryYN=N&FromDate="&FromDate&"&ToDate="&ToDate

dim oCmd1, oCmd2, iAction, Result1, Result2, prm
Set oCmd1=Server.CreateObject("ADODB.Command")
Set oCmd2=Server.CreateObject("ADODB.Command")

set oCmd1.ActiveConnection = db
oCmd1.CommandText = "armyinformix.dbo.submenu0202"
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
oCmd2.CommandText = "armyinformix.dbo.submenu0202"
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
redim Array1(TotalCount,12)

dim pre_sosok, cur_sosok, pre_dutyman, cur_dutyman
dim sosok_num, dutyman_num, array_cnt, cur_monitorpoint
dim dutymancount

pre_sosok = ""
cur_sosok = ""
pre_dutyman = ""
cur_dutyman = ""
sosok_num = 0
dutyman_num = 0

array_cnt = 0

Do while not Result1.EOF

	if Isnull(Result1("sosok")) then
		cur_sosok = "소속 없음"
	else
		cur_sosok = Result1("sosok")	
	end if

	cur_dutyman = Result1("dutyman")
	cur_monitorpoint = Result1("monitorpoint")

	Array1(array_cnt,0) = cur_sosok
	Array1(array_cnt,1) = 0
	Array1(array_cnt,2) = cur_dutyman
	Array1(array_cnt,3) = 0
	Array1(array_cnt,4) = Result1("class")
	Array1(array_cnt,5) = Result1("name")
	Array1(array_cnt,6) ="[<a href='##' onClick=""nLink('"&Result1("receiptfactnum")&"');"">"&Result1("receiptfactnum")&"</a>] "&Result1("nameoffact")
	Array1(array_cnt,7) = replace(Result1("processdate_new"),"/","-")
	Array1(array_cnt,8) = cur_monitorpoint
	Array1(array_cnt,9) = cur_monitorpoint
	Array1(array_cnt,10) = 0
	
	if cur_sosok <> pre_sosok then
		sosok_num = array_cnt
		Array1(CInt(sosok_num),1) = CInt(Array1(CInt(sosok_num),1)) + 1
	else
		Array1(CInt(sosok_num),1) = CInt(Array1(CInt(sosok_num),1)) + 1
	end if

	if cur_dutyman <> pre_dutyman then
		dutyman_num = array_cnt
		Array1(CInt(dutyman_num),3) = CInt(Array1(CInt(dutyman_num),3)) + 1
		Array1(CInt(sosok_num),10) = CInt(Array1(CInt(sosok_num),10)) + 1
		dutymancount = 0
		Array1(CInt(array_cnt),11) = CInt(Array1(CInt(array_cnt),11)) + 1
	else
		Array1(CInt(dutyman_num),3) = CInt(Array1(CInt(dutyman_num),3)) + 1
		Array1(CInt(array_cnt),11) = CInt(Array1(CInt(array_cnt-1),11)) + 1
	end if

	pre_sosok = cur_sosok
	pre_dutyman = cur_dutyman
	array_cnt = array_cnt + 1
Result1.MoveNext
Loop
%>

<script>

	function fn_Search() {

		document.inUpFrm.QueryYN.value = "Y";
		document.inUpFrm.submit();
	}
	
	function fn_Xls() {
		location.href="./list02_Xls.asp?<%=pageWHERE%>"
	}

	function nLink(f){
		pURL = "/menu01/submenu0101/monitoring.asp?FRM=list&factnum=" +f;
		OpenPopMenu(pURL,'monitoring');
	}

</script>

<div id='CalendarLayer' style='display:none;'><iframe name='CalendarFrame' src='/include/Calendar.asp' width='172' height='177' border='0' frameborder='0' scrolling='no'></iframe></div>

<table border="0" width="940" cellpadding="0" cellspacing="3" bgcolor="#EFECE5" align="center">
	<tr bgcolor="#FFFFFF">
		<td>

			<form name="inUpFrm" method="post" action="./list02.asp" onsubmit="return fn_Search(this);" style="margin:0">
			<input type="hidden" name="QueryYN" value="<%=QueryYN%>">
			<table width="100%" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">

			    <tr>
			        <td width="110" bgcolor="#EFEFEF" class="TDCont">조회기간 :</td>
			        <td  bgcolor="#FFFFFF" colspan=3 width=300>
			        	<input value="<%=FromDate%>" name="FromDate" readonly type="text" size="10" onfocus="setFocusColor(this);" onClick="new CalendarFrame.Calendar(this);" onblur="setOutColor(this);">
				    	~
				    	<input value="<%=ToDate%>" name="ToDate" readonly type="text" size="10" onfocus="setFocusColor(this);" onClick="new CalendarFrame.Calendar(this);" onblur="setOutColor(this);">	
			        </td>
			        <td colspan='2' rowspan="2" bgcolor="#FFFFFF" align="center">
			        	<img src="/Images/Btn/BtnSearch.gif" style="cursor:hand;" onClick="fn_Search();">
			        	<br><br><img src="/Images/Btn/BtnExcel.gif" style="cursor:hand;" onClick="fn_Xls();">
			        </td>
				</tr>

			</table>
			</form>
		</td>
	</tr>
</table>


<table width="940" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
<table border="0" width="940" cellspacing="1" cellpadding="0" width="100%" bgcolor="#CCCCCC" align="center">
	<tr height="25" bgcolor="#F3F3F3" align="center">
		<td align="center" width="150"><b>소속</b></td>
		<td align="center" width="65"><b>계급</b></td>
		<td align="center" width="95"><b>성명</b></td>
		<td align="center"><b>사건명</b></td>
		<td align="center" width="100"><b>모니터링일자</b></td>
		<td align="center" width="95"><b>만족도</b></td>
	</tr>
<%
dim temp_total, temp_manjok, temp_manjok_str
temp_total = 0

for i = 0 to TotalCount-1
%>
<tr height="25" bgcolor="#FFFFFF">
	<%
	if CInt(Array1(i,1)) >= 1 then
	%>
<td align="center" rowspan="<%=CInt(Array1(i,1))+CInt(Array1(i,10))%>"><%=Array1(i,0)%><br>(총 <%=Array1(i,1)%>건)</td>
	<%end if%>
	<%
	if CInt(Array1(i,3)) >= 1 then
		temp_total = Array1(i,3)
		temp_manjok = Array1(i,8)
	%>
<td align="center" rowspan="<%=Array1(i,3)%>"><%=Array1(i,4)%></td>
<td align="center" rowspan="<%=Array1(i,3)%>"><%=Array1(i,5)%></td>
	<%
	else
		temp_manjok = cdbl(temp_manjok) + cdbl(Array1(i,8))
	end if
	%>
<td align="left">&nbsp;<%=Array1(i,6)%></td>
<td align="center"><%=Array1(i,7)%></td>
<%
Array1_i_8 = Array1(i,8)
%>
<td align="center"><%=FormatNumber(cdbl(Array1_i_8),2)%></td>
</tr>
	<%
	if CInt(temp_total) = CInt(Array1(i,11)) then
		temp_manjok = cdbl(temp_manjok / temp_total)
		'if inStr(CStr(temp_manjok),".") > 0 then
			temp_manjok = FormatNumber(cdbl(temp_manjok),2)
		'end if
		if CDbl(temp_manjok) >= 9 then
			temp_manjok_str = "만족(" & temp_manjok & ")"
		else
			if CDbl(temp_manjok) < 9 and CDbl(temp_manjok) >= 8 then
				temp_manjok_str = "보통(" & temp_manjok & ")"
			else
				if CDbl(temp_manjok) < 8 then
					temp_manjok_str = "불만족(" & temp_manjok & ")"
				end if
			end if
		end if
	%>
<tr bgcolor="#FFFFFF">
<td colspan=5 align="right"><strong> 총 <%=temp_total%>건 / <%=temp_manjok_str%>&nbsp;</strong></td>
</tr>
	<%end if%>
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
