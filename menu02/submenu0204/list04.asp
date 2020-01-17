<!-- #include virtual="/Include/Top.asp" -->
<!-- #include virtual="/Include/Adovbs.inc" -->
<%
	'####### 파라미터 ##################################################################################
	QueryYN = request("QueryYN")
	FromDate = request("FromDate")
	ToDate = request("ToDate")

	QueryYN = "Y"
	if FromDate = "" then FromDate =left(Date(),7)&"-01" end If
	if ToDate = "" then ToDate=date() end If

	pageWHERE = "QueryYN="&QueryYN&"&FromDate="&FromDate&"&ToDate="&ToDate

	dim EXCEL_CHK, Table_width_and_border, mark_code1, mark_code2
	EXCEL_CHK = "N"
	Table_width_and_border = "width='940' border='0'"
	mark_code1 = "("
	mark_code2 = ")"
%>

<div id='CalendarLayer' style='display:none;'><iframe name='CalendarFrame' src='/include/Calendar.asp' width='172' height='177' border='0' frameborder='0' scrolling='no'></iframe></div>

<script>

	function fn_Search() {

		document.inUpFrm.QueryYN.value = "Y";
		document.inUpFrm.submit();
	}
	
	function fn_Xls() {
		location.href="list04_Xls.asp?<%=pageWHERE%>"
	}

	function nLink(f){
		//alert(f);
		//return;

		pURL = "/menu02/listDetail.asp?QueryYN=Y&FromDate="+document.inUpFrm.FromDate.value+"&ToDate="+document.inUpFrm.ToDate.value+"&Kind=" +f;
		OpenPopMenu(pURL,'ListDetail');
	}

	function nLink7(f){
		pURL = "/menu02/listDetail.asp?QueryYN=Y&FromDate="+document.inUpFrm.FromDate.value+"&ToDate="+document.inUpFrm.ToDate.value+"&Kind=1" +f;
		OpenPopMenu(pURL,'ListDetail');
	}

</script>

<table border="0" width="940" cellpadding="0" cellspacing="3" bgcolor="#EFECE5" align="center">
	<tr bgcolor="#FFFFFF">
		<td>
			<form name="inUpFrm" method="post" action="<%=Menu_2nd%>" onsubmit="return fn_Search(this);" style="margin:0">
			<input type="hidden" name="QueryYN" value="<%=QueryYN%>">
			<table width="100%" border="0" cellspacing="1" cellpadding="0" style="border:#E1DED6 solid 1px">
			    <tr>
			        <td class="TDCont">조회기간 :
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

			'----------------------------------------
			'1) 상담방법별
			'---------------------------------------


%>

<table border="0" cellpadding="0" cellspacing="0" align="center">
	<tr>
		<td>
			<!--<DIV style="OVERFLOW-Y:auto; OVERFLOW-X:auto; MARGIN: 0px 0px 0px 0px; WIDTH:940; HEIGHT:500;">-->

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
			</DIV>
		</td>
	</tr>
</table>

<% End if %>

<!-- #include virtual="/Include/Bottom.asp" -->