<!-- #include virtual="/Include/Adovbs.inc" -->
<!-- #include virtual="/Include/Top.asp" -->
<!-- #include virtual="/Include/PopLayer.asp" -->
<%
dim FromDate, ToDate, QueryYN

QueryYN = request("QueryYN")
FromDate = request("FromDate")
ToDate = request("ToDate")
iCount1 = 0
iCount2 = 0

if FromDate = "" then FromDate = left(Date(),7)&"-01" end If
if ToDate = "" then ToDate=date() end If

dim pageWHERE

pageWHERE = "QueryYN=N&FromDate="&FromDate&"&ToDate="&ToDate

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
	iCount1 = iCount1 + 1
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

<script>

	function fn_Search() {

		document.inUpFrm.QueryYN.value = "Y";
		document.inUpFrm.submit();
	}
	
	function fn_Xls() {
		location.href="./list07_Xls.asp?<%=pageWHERE%>"
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

			<form name="inUpFrm" method="post" action="./list07.asp" onsubmit="return fn_Search(this);" style="margin:0">
			<input type="hidden" name="QueryYN" value="<%=QueryYN%>">
			<table width="100%" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">

			    <tr>
			        <td width="110" bgcolor="#EFEFEF" class="TDCont" align="center">조회기간</td>
			        <td  bgcolor="#FFFFFF" colspan=3 width=300>
			        	<input value="<%=FromDate%>" name="FromDate" readonly type="text" size="10" onfocus="setFocusColor(this);" onClick="new CalendarFrame.Calendar(this);" onblur="setOutColor(this);">
				    	~
				    	<input value="<%=ToDate%>" name="ToDate" readonly type="text" size="10" onfocus="setFocusColor(this);" onClick="new CalendarFrame.Calendar(this);" onblur="setOutColor(this);">	
			        </td>
			        <td width="110" bgcolor="#EFEFEF" class="TDCont" align="center"><span id="txtcount1"></span></td>
			        <td width="110" bgcolor="#EFEFEF" class="TDCont" align="center"><span id="txtcount2"></span></td>
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

<%
dim temp_total
temp_total = 0
icount1 = 0

for i = 0 to TotalCount-1
%>
	<%
	if CInt(Array1(i,1)) >= 1 then
		icount1 = icount1 + 1
		temp_total = Array1(i,1)	
	%>
<table width="940" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
<table width="940" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td><font size='2' color='#0000ff'><b>No: <%=icount1%> </b></font></td></tr></table>

<table border="0" width="940" cellspacing="1" cellpadding="0" width="100%" bgcolor="#CCCCCC" align="center">
<tr height="25" align="center">
<td width="150" bgcolor="#F3F3F3"><b>사건번호</b></td>
<td width="320" bgcolor="#FFFFFF"><b><a href="##" onClick="nLink('<%=Array1(i,0)%>');"><%=Array1(i,0)%></a></b></td>
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
		iCount2 = iCount2 + 1
	end if

	%>
	<%
	if CInt(Array1(i,9)) > 1 then
		iCount2 = iCount2 + 1	
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
<table width="940" height="20" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td align='center'></td></tr></table>
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
if i > 1 then
%>
<script>
	document.getElementById('txtcount1').innerHTML ="사건:&nbsp;<font color='#0000ff' size='2'><%=icount1%></font>&nbsp;건";
	document.getElementById('txtcount2').innerHTML ="인원:&nbsp;<font color='#0000ff' size='2'><%=TotalCount%></font>&nbsp;명";
</script>
<% end if %>
<!-- #include virtual="/Include/Bottom.asp" -->