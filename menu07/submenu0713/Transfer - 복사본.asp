<!-- #include virtual="/Include/Top.asp" -->
<%
	'1. 파라미터 얻어오기
	curPage = request("curPage")
	'3. 쿼리 실행
	'sql = db_getSqlWithPage(sql_tb, sql_index, sql_field, sql_where, sql_orderby, pageSize, curPage)
	sql = "	select	*	from	T_BusinessControl	order by	[T_WorkDay]"
	set rs = db.execute(sql)

	do until rs.eof

		if rs("T_WorkDay") = 1 then	'일요일
			sStartTime7 = rs("T_StartTime")
			sFinishTime7 = rs("T_EndTime")
		elseif rs("T_WorkDay") = 2 then	'월요일
			sStartTime1 = rs("T_StartTime")
			sFinishTime1 = rs("T_EndTime")
		elseif rs("T_WorkDay") = 3 then	'화요일
			sStartTime2 = rs("T_StartTime")
			sFinishTime2 = rs("T_EndTime")
		elseif rs("T_WorkDay") = 4 then	'수요일
			sStartTime3 = rs("T_StartTime")
			sFinishTime3 = rs("T_EndTime")
		elseif rs("T_WorkDay") = 5 then	'목요일
			sStartTime4 = rs("T_StartTime")
			sFinishTime4 = rs("T_EndTime")
		elseif rs("T_WorkDay") = 6 then	'금요일
			sStartTime5 = rs("T_StartTime")
			sFinishTime5 = rs("T_EndTime")
		elseif rs("T_WorkDay") = 7 then	'토요일
			sStartTime6 = rs("T_StartTime")
			sFinishTime6 = rs("T_EndTime")
		elseif rs("T_WorkDay") = 8 then	'법정공휴일
			sStartTime8 = rs("T_StartTime")
			sFinishTime8 = rs("T_EndTime")
		end if

		rs.movenext
	loop

	'4. Paging HTML 작성

%>

<div id='CalendarLayer' style='display:none;'><iframe name='CalendarFrame' src='/include/Calendar.asp' width='172' height='177' border='0' frameborder='0' scrolling='no'></iframe></div>

<table border="0" width="940" cellpadding="0" cellspacing="3" bgcolor="#EFECE5" align="center">
	<tr bgcolor="#FFFFFF">
		<td>
			<form name="inUpFrm" method="post" action="/Manage/Worktime/Worktime_InsUpDel.asp">
			<table width="100%" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">
			    <tr>
			        <td width="110" bgcolor="#EFEFEF" class="TDCont" align="center">구분</td>
			        <td width="80" bgcolor="#EFEFEF" class="TDCont" align="center">시작시각</td>
			        <td width="80" bgcolor="#EFEFEF" class="TDCont" align="center">종료시각</td>
			        <td bgcolor="#EFEFEF" class="TDCont" align="center">비고</td>
				</tr>
			    <tr>
			        <td width="110" bgcolor="#EFEFEF" class="TDCont" align="center">월요일</td>
			        <td width="80" bgcolor="#FFFFFF" class="TDCont" align="center"><input name="StartTime1" type="text" size="4" onblur="setOutColor(this);" value="<%=sStartTime1%>"></td>
			        <td width="80" bgcolor="#FFFFFF" class="TDCont" align="center"><input name="FinishTime1" type="text" size="4" onblur="setOutColor(this);" value="<%=sFinishTime1%>"></td>
			        <td bgcolor="#FFFFFF" class="TDCont" rowspan=8>&nbsp;24시 기준으로 입력하십시오.<br><br>&nbsp;시각 입력 예)<br><br>&nbsp;오전9시 => 0900, 오후6시30분 => 1830</td>
				</tr>
			    <tr>
			        <td width="110" bgcolor="#EFEFEF" class="TDCont" align="center">화요일</td>
			        <td width="80" bgcolor="#FFFFFF" class="TDCont" align="center"><input name="StartTime2" type="text" size="4" onblur="setOutColor(this);" value="<%=sStartTime2%>"></td>
			        <td width="80" bgcolor="#FFFFFF" class="TDCont" align="center"><input name="FinishTime2" type="text" size="4" onblur="setOutColor(this);" value="<%=sFinishTime2%>"></td>
				</tr>
			    <tr>
			        <td width="110" bgcolor="#EFEFEF" class="TDCont" align="center">수요일</td>
			        <td width="80" bgcolor="#FFFFFF" class="TDCont" align="center"><input name="StartTime3" type="text" size="4" onblur="setOutColor(this);" value="<%=sStartTime3%>"></td>
			        <td width="80" bgcolor="#FFFFFF" class="TDCont" align="center"><input name="FinishTime3" type="text" size="4" onblur="setOutColor(this);" value="<%=sFinishTime3%>"></td>
				</tr>
			    <tr>
			        <td width="110" bgcolor="#EFEFEF" class="TDCont" align="center">목요일</td>
			        <td width="80" bgcolor="#FFFFFF" class="TDCont" align="center"><input name="StartTime4" type="text" size="4" onblur="setOutColor(this);" value="<%=sStartTime4%>"></td>
			        <td width="80" bgcolor="#FFFFFF" class="TDCont" align="center"><input name="FinishTime4" type="text" size="4" onblur="setOutColor(this);" value="<%=sFinishTime4%>"></td>
				</tr>
			    <tr>
			        <td width="110" bgcolor="#EFEFEF" class="TDCont" align="center">금요일</td>
			        <td width="80" bgcolor="#FFFFFF" class="TDCont" align="center"><input name="StartTime5" type="text" size="4" onblur="setOutColor(this);" value="<%=sStartTime5%>"></td>
			        <td width="80" bgcolor="#FFFFFF" class="TDCont" align="center"><input name="FinishTime5" type="text" size="4" onblur="setOutColor(this);" value="<%=sFinishTime5%>"></td>
				</tr>
			    <tr>
			        <td width="110" bgcolor="#EFEFEF" class="TDCont" align="center">토요일</td>
			        <td width="80" bgcolor="#FFFFFF" class="TDCont" align="center"><input name="StartTime6" type="text" size="4" onblur="setOutColor(this);" value="<%=sStartTime6%>"></td>
			        <td width="80" bgcolor="#FFFFFF" class="TDCont" align="center"><input name="FinishTime6" type="text" size="4" onblur="setOutColor(this);" value="<%=sFinishTime6%>"></td>
				</tr>
			    <tr>
			        <td width="110" bgcolor="#EFEFEF" class="TDCont" align="center">일요일</td>
			        <td width="80" bgcolor="#FFFFFF" class="TDCont" align="center"><input name="StartTime7" type="text" size="4" onblur="setOutColor(this);" value="<%=sStartTime7%>"></td>
			        <td width="80" bgcolor="#FFFFFF" class="TDCont" align="center"><input name="FinishTime7" type="text" size="4" onblur="setOutColor(this);" value="<%=sFinishTime7%>"></td>
				</tr>
			    <tr>
			        <td width="110" bgcolor="#EFEFEF" class="TDCont" align="center">법정공휴일</td>
			        <td width="80" bgcolor="#FFFFFF" class="TDCont" align="center"><input name="StartTime8" type="text" size="4" onblur="setOutColor(this);" value="<%=sStartTime8%>"></td>
			        <td width="80" bgcolor="#FFFFFF" class="TDCont" align="center"><input name="FinishTime8" type="text" size="4" onblur="setOutColor(this);" value="<%=sFinishTime8%>"></td>
				</tr>	
			</table>
			</form>
			<table border="0" cellspacing="0" width="100%" align="center">
				<tr height="30">
					<td align="center">
						<img src="/Images/Btn/BtnSubmit.gif" style="cursor:hand;" align="absmiddle" onClick="javascript:fn_inup(document.inUpFrm);">
						<img src="/Images/Btn/BtnReset.gif" style="cursor:hand;" align="absmiddle" onClick="javascript:fn_reset();">
						<img src="/Images/Btn/BtnClose.gif" style="cursor:hand;" align="absmiddle" onClick="javascript:self.close();">
					</td>
				</tr>
			</table>	
		</td>
	</tr>
</table>

<table width="940" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
<!-- #include virtual="/Include/Bottom.asp" -->


<script>
function fn_inup(inUpFrm) {

	//필수입력값 체크

	if ( inUpFrm.StartTime1.value == '' || inUpFrm.StartTime1.value.length != 4 )
	{
		alert('월요일의 시작시각을 숫자4자리로 정확히 입력하십시오!');
		inUpFrm.StartTime1.focus();
		return;
	}
	if ( inUpFrm.FinishTime1.value == '' || inUpFrm.FinishTime1.value.length != 4 )
	{
		alert('월요일의 종료시각을 숫자4자리로 정확히 입력하십시오!');
		inUpFrm.FinishTime1.focus();
		return;
	}	
	if ( inUpFrm.StartTime2.value == '' || inUpFrm.StartTime2.value.length != 4 )
	{
		alert('화요일의 시작시각을 숫자4자리로 정확히 입력하십시오!');
		inUpFrm.StartTime2.focus();
		return;
	}
	if ( inUpFrm.FinishTime2.value == '' || inUpFrm.FinishTime2.value.length != 4 )
	{
		alert('화요일의 종료시각을 숫자4자리로 정확히 입력하십시오!');
		inUpFrm.FinishTime2.focus();
		return;
	}
	if ( inUpFrm.StartTime3.value == '' || inUpFrm.StartTime3.value.length != 4 )
	{
		alert('수요일의 시작시각을 숫자4자리로 정확히 입력하십시오!');
		inUpFrm.StartTime3.focus();
		return;
	}
	if ( inUpFrm.FinishTime3.value == '' || inUpFrm.FinishTime3.value.length != 4 )
	{
		alert('수요일의 종료시각을 숫자4자리로 정확히 입력하십시오!');
		inUpFrm.FinishTime3.focus();
		return;
	}
	if ( inUpFrm.StartTime4.value == '' || inUpFrm.StartTime4.value.length != 4 )
	{
		alert('목요일의 시작시각을 숫자4자리로 정확히 입력하십시오!');
		inUpFrm.StartTime4.focus();
		return;
	}
	if ( inUpFrm.FinishTime4.value == '' || inUpFrm.FinishTime4.value.length != 4 )
	{
		alert('목요일의 종료시각을 숫자4자리로 정확히 입력하십시오!');
		inUpFrm.FinishTime4.focus();
		return;
	}
	if ( inUpFrm.StartTime5.value == '' || inUpFrm.StartTime5.value.length != 4 )
	{
		alert('금요일의 시작시각을 숫자4자리로 정확히 입력하십시오!');
		inUpFrm.StartTime5.focus();
		return;
	}
	if ( inUpFrm.FinishTime5.value == '' || inUpFrm.FinishTime5.value.length != 4 )
	{
		alert('금요일의 종료시각을 숫자4자리로 정확히 입력하십시오!');
		inUpFrm.FinishTime5.focus();
		return;
	}
	if ( inUpFrm.StartTime6.value == '' || inUpFrm.StartTime6.value.length != 4 )
	{
		alert('토요일의 시작시각을 숫자4자리로 정확히 입력하십시오!');
		inUpFrm.StartTime6.focus();
		return;
	}
	if ( inUpFrm.FinishTime6.value == '' || inUpFrm.FinishTime6.value.length != 4 )
	{
		alert('토요일의 종료시각을 숫자4자리로 정확히 입력하십시오!');
		inUpFrm.FinishTime6.focus();
		return;
	}
	if ( inUpFrm.StartTime7.value == '' || inUpFrm.StartTime7.value.length != 4 )
	{
		alert('일요일의 시작시각을 숫자4자리로 정확히 입력하십시오!');
		inUpFrm.StartTime7.focus();
		return;
	}
	if ( inUpFrm.FinishTime7.value == '' || inUpFrm.FinishTime7.value.length != 4 )
	{
		alert('일요일의 종료시각을 숫자4자리로 정확히 입력하십시오!');
		inUpFrm.FinishTime7.focus();
		return;
	}
	if ( inUpFrm.StartTime8.value == '' || inUpFrm.StartTime8.value.length != 4 )
	{
		alert('법정공휴일의 시작시각을 숫자4자리로 정확히 입력하십시오!');
		inUpFrm.StartTime8.focus();
		return;
	}
	if ( inUpFrm.FinishTime8.value == '' || inUpFrm.FinishTime8.value.length != 4 )
	{
		alert('법정공휴일의 종료시각을 숫자4자리로 정확히 입력하십시오!');
		inUpFrm.FinishTime8.focus();
		return;
	}


	if(confirm("변경된 값을 저장하시겠습니까?"))
		inUpFrm.submit();
	else
		return;
}

function fn_reset() {

		document.inUpFrm.StartTime1.value="<%=sStartTime1%>";
		document.inUpFrm.FinishTime1.value="<%=sFinishTime1%>";

		document.inUpFrm.StartTime2.value="<%=sStartTime2%>";
		document.inUpFrm.FinishTime2.value="<%=sFinishTime2%>";

		document.inUpFrm.StartTime3.value="<%=sStartTime3%>";
		document.inUpFrm.FinishTime3.value="<%=sFinishTime3%>";

		document.inUpFrm.StartTime4.value="<%=sStartTime4%>";
		document.inUpFrm.FinishTime4.value="<%=sFinishTime4%>";

		document.inUpFrm.StartTime5.value="<%=sStartTime5%>";
		document.inUpFrm.FinishTime5.value="<%=sFinishTime5%>";

		document.inUpFrm.StartTime6.value="<%=sStartTime6%>";
		document.inUpFrm.FinishTime6.value="<%=sFinishTime6%>";

		document.inUpFrm.StartTime7.value="<%=sStartTime7%>";
		document.inUpFrm.FinishTime7.value="<%=sFinishTime7%>";

		document.inUpFrm.StartTime8.value="<%=sStartTime8%>";
		document.inUpFrm.FinishTime8.value="<%=sFinishTime8%>";

		return;
}
</script>