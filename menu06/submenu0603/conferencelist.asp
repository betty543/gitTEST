<!-- #include virtual="/Include/Top.asp" -->
<%
	'####### �Ķ���� ##################################################################################
	QueryYN = request("QueryYN")
	FromDate = request("FromDate")
	ToDate = request("ToDate")
	whereCD3 = Trim(request("whereCD3"))
	whereCD7 = Trim(request("whereCD7"))

	If QueryYN = "" Then
		whereCD3 = "1"
	End if


	if FromDate = "" then FromDate =left(Date(),7)&"-01" end If
	if ToDate = "" then ToDate=date() end If

	pageWHERE = "QueryYN="&QueryYN&"&FromDate="&FromDate&"&ToDate="&ToDate&"&whereCD3="&whereCD3&"&whereCD7="&whereCD7

%>

<div id='CalendarLayer' style='display:none;'><iframe name='CalendarFrame' src='/include/Calendar.asp' width='172' height='177' border='0' frameborder='0' scrolling='no'></iframe></div>

<script>

	function fn_Search() {

		document.inUpFrm.QueryYN.value = "Y";
		document.inUpFrm.submit();
	}
	
	function fn_Xls() {
		location.href="list01_Xls.asp?<%=pageWHERE%>"
	}
</script>
<table border="0" width="940" cellpadding="0" cellspacing="3" bgcolor="#EFECE5" align="center">
	<tr bgcolor="#FFFFFF">
		<td>
			<form name="inUpFrm" method="post" action="<%=Menu_2nd%>" onsubmit="return fn_Search(this);" style="margin:0">
			<input type="hidden" name="QueryYN" value="<%=QueryYN%>">
			<table width="100%" border="0" cellspacing="1" cellpadding="0" style="border:#E1DED6 solid 1px">
			    <tr>
			        <td class="TDCont">��ȸ�Ⱓ :
			        	<input value="<%=FromDate%>" name="FromDate" readonly type="text" size="10" onfocus="setFocusColor(this); new CalendarFrame.Calendar(this);" onblur="setOutColor(this);">
				    	~
				    	<input value="<%=ToDate%>" name="ToDate" readonly type="text" size="10" onfocus="setFocusColor(this); new CalendarFrame.Calendar(this);" onblur="setOutColor(this);">
			        </td>


			        <td class="TDR5px">
			        	<img src="/Images/Btn/BtnSearch.gif" align="absmiddle" style="cursor:hand;" onClick="fn_Search();">
			        	<img src="/Images/Btn/BtnExcel.gif" align="absmiddle" style="cursor:hand;" onClick="fn_Xls();">
			        </td>
			    </tr>
			</table>
			</form>
		</td>
	</tr>
</table>
<table border="0" width="100%" cellpadding="0" cellspacing="0" align="center"><tr height="5"><td></td></tr></table>
<%

	If QueryYN = "Y" Then

%>

<table border="0" cellpadding="0" cellspacing="0" align="center">
	<tr>
		<td>
			<!--<DIV style="OVERFLOW-Y:auto; OVERFLOW-X:auto; MARGIN: 0px 0px 0px 0px; WIDTH:940; HEIGHT:500;">-->
			<table width="940"  border="0" cellpadding="1" cellspacing="1" bgcolor="#EFECE5" align="center">

				<tr bgcolor='#EEF6FF'>
					<td align='center' class='TDCont'  width='150'>����ð�</td>
					<td align='center' class='TDCont' >������</td>
					<td align='center' class='TDCont' >�ѰǼ�</td>
					<td align='center' class='TDCont' >���Ἲ��</td>
					<td align='center' class='TDCont' >����</td>
					<td align='center' class='TDCont' >���</td>
					<td align='center' class='TDCont' >��ȭ�ð�</td>

				</tr>

				<tr bgcolor='#ffffff'>
					<td align='center' class='TDCont'  width='150'>2009-06-18 21:37:00</td>
					<td align='center' class='TDCont' width="150" >admin</td>
					<td align='center' class='TDCont' width="150">2</td>
					<td align='center' class='TDCont' width="150">1</td>
					<td align='center' class='TDCont' width="150">1</td>
					<td align='center' class='TDCont' width="150">0</td>
					<td align='center' class='TDCont' width="150">02:00</td>

				</tr>
			</table>
		</td>
	</tr>
</table>


<% End if %>

<!-- #include virtual="/Include/Bottom.asp" -->