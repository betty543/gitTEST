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

	dim EXCEL_CHK, Table_width_and_border, mark_code1, mark_code2
	EXCEL_CHK = "N"
	Table_width_and_border = "width='940' border='0'"
	mark_code1 = "("
	mark_code2 = ")"


dim oCmd1, oCmd2, iAction, Result1, Result2, prm
Set oCmd1=Server.CreateObject("ADODB.Command")
Set oCmd2=Server.CreateObject("ADODB.Command")

%>

<script>

	function fn_Search() {

		document.inUpFrm.QueryYN.value = "Y";
		document.inUpFrm.submit();
	}
	
	function fn_Xls() {
		location.href="./list08_Xls.asp?<%=pageWHERE%>"
	}


</script>

<div id='CalendarLayer' style='display:none;'><iframe name='CalendarFrame' src='/include/Calendar.asp' width='172' height='177' border='0' frameborder='0' scrolling='no'></iframe></div>

<table border="0" width="940" cellpadding="0" cellspacing="3" bgcolor="#EFECE5" align="center">
	<tr bgcolor="#FFFFFF">
		<td>

			<form name="inUpFrm" method="post" action="./list08.asp" onsubmit="return fn_Search(this);" style="margin:0">
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


<!--부대별/응답자별현황 시작 -->
<!-- #include file ="./list08_1.asp" -->
<!--부대별/응답자별현황 끝 -->

<%
set prm = nothing
%>

<!-- #include virtual="/Include/Bottom.asp" -->
