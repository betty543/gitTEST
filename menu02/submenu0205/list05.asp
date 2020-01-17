<!-- #include virtual="/Include/Adovbs.inc" -->
<!-- #include virtual="/Include/Top.asp" -->
<!-- #include virtual="/Include/PopLayer.asp" -->
<%
dim FromDate, ToDate, QueryYN

QueryYN = request("QueryYN")
FromDate = request("FromDate")
ToDate = request("ToDate")

if FromDate = "" then FromDate = left(Date(),7)&"-01" end If
if ToDate = "" then ToDate=date() end If

dim pageWHERE

pageWHERE = "QueryYN=N&FromDate="&FromDate&"&ToDate="&ToDate

dim oCmd1, iAction, Result1, prm
Set oCmd1=Server.CreateObject("ADODB.Command")

set oCmd1.ActiveConnection = db
oCmd1.CommandText = "armyinformix.dbo.submenu0205"
oCmd1.CommandType = adCmdStoredProc

iAction = "1"

set prm = oCmd1.CreateParameter("@iAction",adChar,adParamInput,1,iAction)
oCmd1.Parameters.Append prm
set prm = oCmd1.CreateParameter("@sDate",adChar,adParamInput,10,FromDate)
oCmd1.Parameters.Append prm
set prm = oCmd1.CreateParameter("@eDate",adChar,adParamInput,10,ToDate)
oCmd1.Parameters.Append prm

set Result1 = oCmd1.Execute
%>

<script>

	function fn_Search() {

		document.inUpFrm.QueryYN.value = "Y";
		document.inUpFrm.submit();
	}
	
	function fn_Xls() {
		location.href="./list05_Xls.asp?<%=pageWHERE%>"
	}


</script>

<div id='CalendarLayer' style='display:none;'><iframe name='CalendarFrame' src='/include/Calendar.asp' width='172' height='177' border='0' frameborder='0' scrolling='no'></iframe></div>

<table border="0" width="940" cellpadding="0" cellspacing="3" bgcolor="#EFECE5" align="center">
	<tr bgcolor="#FFFFFF">
		<td>

			<form name="inUpFrm" method="post" action="./list05.asp" onsubmit="return fn_Search(this);" style="margin:0">
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
		<td rowspan=2 ><b>구분</b></td>
		<td rowspan=2 width=100><b>계</b></td>
		<td colspan=3><b>만족도</b></td>
		<td rowspan=2 width=100><b>통화불능</b></td>
		<td rowspan=2 width=100><b>설문거부</b></td>
		<td rowspan=2 width=100><b>미실시</b></td>
	</tr>

	<tr height="25" bgcolor="#F3F3F3" align="center">
		<td width=100><b>만족</b></td>
		<td width=100><b>보통</b></td>
		<td width=100><b>불만족</b></td>
	</tr>

<%
dim tot_sect1, tot_sect2, tot_sect3, tot_sect4, tot_sect5, tot_sect6, tot_sect7
dim cur_sect1, cur_sect2, cur_sect3, cur_sect4, cur_sect5, cur_sect6, cur_sect7
dim per_sect1, per_sect2, per_sect3, per_sect4, per_sect5, per_sect6, per_sect7

Do while not Result1.EOF

	cur_sect1 = Result1("sect1")
	cur_sect2 = Result1("sect2")
	cur_sect3 = Result1("sect3")
	cur_sect4 = Result1("sect4")
	cur_sect5 = Result1("sect5")
	cur_sect6 = Result1("sect6")
	cur_sect7 = Result1("sect7")

	if Result1("code") = "B00" then
		tot_sect1 = cur_sect1
		tot_sect2 = cur_sect2
		tot_sect3 = cur_sect3
		tot_sect4 = cur_sect4
		tot_sect5 = cur_sect5
		tot_sect6 = cur_sect6
		tot_sect7 = cur_sect7
	end if
	
	if CInt(tot_sect1) = 0 then
		per_sect1 = 0
	else
		per_sect1 = CDBL((cur_sect1/tot_sect1) * 100)
		if inStr(per_sect1,".") > 0 then
			per_sect1 = FormatNumber(cdbl(per_sect1),2)
		end if
	end if
	if CInt(cur_sect1) = 0 then
		per_sect2 = 0
	else
		per_sect2 = CDBL((cur_sect2/cur_sect1) * 100)
		if inStr(per_sect2,".") > 0 then
			per_sect2 = FormatNumber(cdbl(per_sect2),2)
		end if
	end if
	if CInt(cur_sect1) = 0 then
		per_sect3 = 0
	else
		per_sect3 = CDBL((cur_sect3/cur_sect1) * 100)
		if inStr(per_sect3,".") > 0 then
			per_sect3 = FormatNumber(cdbl(per_sect3),2)
		end if
	end if
	if CInt(cur_sect1) = 0 then
		per_sect4 = 0
	else
		per_sect4 = CDBL((cur_sect4/cur_sect1) * 100)
		if inStr(per_sect4,".") > 0 then
			per_sect4 = FormatNumber(cdbl(cur_sect4),2)
		end if
	end if
	if CInt(cur_sect1) = 0 then
		per_sect5 = 0
	else
		per_sect5 = CDBL((cur_sect5/cur_sect1) * 100)
		if inStr(per_sect5,".") > 0 then
			per_sect5 = FormatNumber(cdbl(per_sect5),2)
		end if
	end if
	if CInt(cur_sect1) = 0 then
		per_sect6 = 0
	else
		per_sect6 = CDBL((cur_sect6/cur_sect1) * 100)
		if inStr(per_sect6,".") > 0 then
			per_sect6 = FormatNumber(cdbl(per_sect6),2)
		end if
	end if
	if CInt(cur_sect1) = 0 then
		per_sect7 = 0
	else
		per_sect7 = CDBL((cur_sect7/cur_sect1) * 100)
		if inStr(per_sect7,".") > 0 then
			per_sect7 = FormatNumber(cdbl(per_sect7),2)
		end if
	end if

per_sect1 = FormatNumber(cdbl(per_sect1),2)
per_sect2 = FormatNumber(cdbl(per_sect2),2)
per_sect3 = FormatNumber(cdbl(per_sect3),2)
per_sect4 = FormatNumber(cdbl(per_sect4),2)
per_sect5 = FormatNumber(cdbl(per_sect5),2)
per_sect6 = FormatNumber(cdbl(per_sect6),2)
per_sect7 = FormatNumber(cdbl(per_sect7),2)		
%>	

<%
if Result1("code") = "B00" then
%>
	
		<tr bgcolor="#FFFFFF">
			<td align="center"><b><%=Result1("codename")%></b></td>
			<td align="center"><strong><%=cur_sect1%></strong></td>
			<td align="center"><strong><%=cur_sect2%><br>(<%=per_sect2%>%)</strong></td>
			<td align="center"><strong><%=cur_sect3%><br>(<%=per_sect3%>%)</strong></td>
			<td align="center"><strong><%=cur_sect4%><br>(<%=per_sect4%>%)</strong></td>
			<td align="center"><strong><%=cur_sect5%><br>(<%=per_sect5%>%)</strong></td>
			<td align="center"><strong><%=cur_sect6%><br>(<%=per_sect6%>%)</strong></td>
			<td align="center"><strong><%=cur_sect7%><br>(<%=per_sect7%>%)</strong></td>
		</tr>
	
<%else%>

		<tr bgcolor="#FFFFFF">
			<td align="center"><b><%=Result1("codename")%></b></td>
			<td align="center"><%=cur_sect1%><br>(<%=per_sect1%>%)</td>
			<td align="center"><%=cur_sect2%><br>(<%=per_sect2%>%)</td>
			<td align="center"><%=cur_sect3%><br>(<%=per_sect3%>%)</td>
			<td align="center"><%=cur_sect4%><br>(<%=per_sect4%>%)</td>
			<td align="center"><%=cur_sect5%><br>(<%=per_sect5%>%)</td>
			<td align="center"><%=cur_sect6%><br>(<%=per_sect6%>%)</td>
			<td align="center"><%=cur_sect7%><br>(<%=per_sect7%>%)</td>
		</tr>

<%end if%>		
		
<%
Result1.MoveNext
Loop
%>		
		
</table>

<%
set oCmd1.ActiveConnection = nothing
set oCmd1 = nothing
set Result1 = nothing
set prm = nothing
%>

<!-- #include virtual="/Include/Bottom.asp" -->
