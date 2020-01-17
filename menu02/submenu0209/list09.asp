<!-- #include virtual="/Include/Adovbs.inc" -->
<!-- #include virtual="/Include/Top.asp" -->
<!-- #include virtual="/Include/PopLayer.asp" -->
<%
dim FromDate, ToDate, QueryYN

'on Error Resume next

QueryYN = request("QueryYN")
FromDate = request("FromDate")
ToDate = request("ToDate")

if FromDate = "" then FromDate = left(Date(),4)&"-01-01" end If
if ToDate = "" then ToDate=date() end If

%>

<script>

	function fn_Search() {

		document.inUpFrm.QueryYN.value = "Y";
		document.inUpFrm.submit();
	}
	
	function fn_Xls() {
		location.href="./list09_Xls.asp?<%=pageWHERE%>"
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

			<form name="inUpFrm" method="post" action="./list09.asp" onsubmit="return fn_Search(this);" style="margin:0">
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
		<td align="center" width="150" rowspan='2'><b>소속</b></td>
		<td align="center" width="65"  rowspan='2'><b>계급</b></td>
		<td align="center" width="95"  rowspan='2'><b>성명</b></td>
		<td align="center" width="100" rowspan='2'><b>총건수</b></td>
		<td align="center" width="300" colspan='3'><b>대기건수</b></td>
		<td align="center" width="100" rowspan='2'><b>진행건수</b></td>
		<td align="center" width="300" colspan='3'><b>완료건수</b></td>
	</tr>
	<tr height="25" bgcolor="#F3F3F3" align="center">

		<td align="center" width="100"><b>총계</b></td>
		<td align="center" width="100"><b>첨부O</b></td>
		<td align="center" width="100"><b>첨부X</b></td>

		<td align="center" width="100"><b>총계</b></td>
		<td align="center" width="100"><b>완료</b></td>
		<td align="center" width="100"><b>완료후누락</b></td>
	</tr>

<%

SQL = "select unit, d.name as 'sosok', id, c.class, c.name,"
SQL = SQL & " (select count(receiptfactnum)  from armyinformix.dbo.receiptfact where dutyman = c.id and inputdate >= '" & FromDate & "' and inputdate <= '"& ToDate & "' ) as count0,"
SQL = SQL & " (select count(receiptfactnum)  from armyinformix.dbo.receiptfact where dutyman = c.id and inputdate >= '" & FromDate & "' and inputdate <= '"& ToDate & "' and ( processgb is null or processgb ='' )) as count1,"
SQL = SQL & " (select count(receiptfactnum)  from armyinformix.dbo.receiptfact where dutyman = c.id and inputdate >= '" & FromDate & "' and inputdate <= '"& ToDate & "' and filecnt > 0 and (  processgb is null or processgb ='' )) as count2,"
SQL = SQL & " (select count(receiptfactnum)  from armyinformix.dbo.receiptfact where dutyman = c.id and inputdate >= '" & FromDate & "' and inputdate <= '"& ToDate & "' and filecnt = 0  and ( processgb is null or processgb ='' )) as count3,"
SQL = SQL & " (select count(receiptfactnum)  from armyinformix.dbo.receiptfact where dutyman = c.id and inputdate >= '" & FromDate & "' and inputdate <= '"& ToDate & "' and processgb in ('1','2' )) as count4,"

SQL = SQL & " (select count(receiptfactnum)  from armyinformix.dbo.receiptfact where dutyman = c.id and inputdate >= '" & FromDate & "' and inputdate <= '"& ToDate & "' and processgb in ('8','9','100' )) as count5,"
SQL = SQL & " (select count(receiptfactnum)  from armyinformix.dbo.receiptfact where dutyman = c.id and inputdate >= '" & FromDate & "' and inputdate <= '"& ToDate & "' and processgb in ('9','100' )) as count6,"
SQL = SQL & " (select count(receiptfactnum)  from armyinformix.dbo.receiptfact where dutyman = c.id and inputdate >= '" & FromDate & "' and inputdate <= '"& ToDate & "' and processgb in ('8' )) as count7"

SQL = SQL & " from armyinformix.dbo.user1 c"
SQL = SQL & " left outer join armyinformix.dbo.pBudae d"
SQL = SQL & " on c.unit = d.auth"
SQL = SQL & " where d.name is not null"
SQL = SQL & " order by auth asc, id asc"


	set RsGBN = db.execute(sql)
	
	do until RsGBN.eof

		sunit = RsGBN("unit")
		
		sFirstLine = "<tr height='25' bgcolor='#ffffff' align='center'><td align='center'"
		iCount = 0
		do until sunit <> RsGBN("unit")
			iCount = iCount + 1
			if iCount = 1 then
				sSecoundLine = RsGBN("sosok") & "</td>"
				sSecoundLine = sSecoundLine & "<td align='center'>" &RsGBN("class") & "</td>"
				sSecoundLine = sSecoundLine & "<td align='center'>" &RsGBN("name") & "</td>"
				sSecoundLine = sSecoundLine & "<td align='center'>" &RsGBN("count0") & "</td>"
				sSecoundLine = sSecoundLine & "<td align='center'>" &RsGBN("count1") & "</td>"
				sSecoundLine = sSecoundLine & "<td align='center'>" &RsGBN("count2") & "</td>"
				sSecoundLine = sSecoundLine & "<td align='center'>" &RsGBN("count3") & "</td>"
				sSecoundLine = sSecoundLine & "<td align='center'>" &RsGBN("count4") & "</td>"
				sSecoundLine = sSecoundLine & "<td align='center'>" &RsGBN("count5") & "</td>"
				sSecoundLine = sSecoundLine & "<td align='center'>" &RsGBN("count6") & "</td>"
				sSecoundLine = sSecoundLine & "<td align='center'>" &RsGBN("count7") & "</td></tr>"
			else

				sSecoundLine = sSecoundLine & "<tr height='25' bgcolor='#ffffff' align='center'>"
				sSecoundLine = sSecoundLine & "<td align='center'>" &RsGBN("class") & "</td>"
				sSecoundLine = sSecoundLine & "<td align='center'>" &RsGBN("name") & "</td>"
				sSecoundLine = sSecoundLine & "<td align='center'>" &RsGBN("count0") & "</td>"
				sSecoundLine = sSecoundLine & "<td align='center'>" &RsGBN("count1") & "</td>"
				sSecoundLine = sSecoundLine & "<td align='center'>" &RsGBN("count2") & "</td>"
				sSecoundLine = sSecoundLine & "<td align='center'>" &RsGBN("count3") & "</td>"
				sSecoundLine = sSecoundLine & "<td align='center'>" &RsGBN("count4") & "</td>"
				sSecoundLine = sSecoundLine & "<td align='center'>" &RsGBN("count5") & "</td>"
				sSecoundLine = sSecoundLine & "<td align='center'>" &RsGBN("count6") & "</td>"
				sSecoundLine = sSecoundLine & "<td align='center'>" &RsGBN("count7") & "</td></tr>"
			end if
				
			icount0 = icount0 + RsGBN("count0")
			icount1 = icount1 + RsGBN("count1")
			icount2 = icount2 + RsGBN("count2")
			icount3 = icount3 + RsGBN("count3")
			icount4 = icount4 + RsGBN("count4")
			icount5 = icount5 + RsGBN("count5")
			icount6 = icount6 + RsGBN("count6")
			icount7 = icount7 + RsGBN("count7")

			RsGBN.movenext
			if RsGBN.eof then
				exit do
			end if
		loop
		response.write sFirstLine & " rowspan=" & iCount+1 & ">" & sSecoundLine
		sSecoundLine = "<tr height='25' bgcolor='#EFEFEF' align='center'>"
		sSecoundLine = sSecoundLine & "<td align='center' colspan='2'>소계</td>"
		sSecoundLine = sSecoundLine & "<td align='center'>" & icount0 & "</td>"
		sSecoundLine = sSecoundLine & "<td align='center'>" & icount1 & "</td>"
		sSecoundLine = sSecoundLine & "<td align='center'>" & icount2 & "</td>"
		sSecoundLine = sSecoundLine & "<td align='center'>" & icount3 & "</td>"
		sSecoundLine = sSecoundLine & "<td align='center'>" & icount4 & "</td>"
		sSecoundLine = sSecoundLine & "<td align='center'>" & icount5 & "</td>"
		sSecoundLine = sSecoundLine & "<td align='center'>" & icount6 & "</td>"
		sSecoundLine = sSecoundLine & "<td align='center'>" & icount7 & "</td></tr>"
		response.write sSecoundLine


			icount0 = 0
			icount1 = 0
			icount2 = 0
			icount3 = 0
			icount4 = 0
			icount5 = 0
			icount6 = 0
			icount7 = 0


	loop
	%>

	
</table>
<!-- #include virtual="/Include/Bottom.asp" -->
