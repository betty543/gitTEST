<!-- #include virtual="/Include/Top_PopUp.asp" -->
<%
	sHoliday = Request("Holiday")
	sHoliday_name = Request("Holiday_Name")
	sEveryYear = Request("EveryYear")
	if sEveryYear = "" then
		sEveryYear = "0"
	end if

	guboon = Request("guboon")

	if guboon = "UP" Then
		'수정하기
		sql = "update	T_Holiday	set Holiday_Name = '" & sHoliday_name & "'"
		sql = sql & "	,	Every_Year = '" & sEveryYear & "'"
		sql = sql & "	where	[Holiday] = '" & sHoliday & "'"

		db.execute(sql)

		sql = "	delete from T_Holiday "
		sql = sql & "	where	[Holiday] = '" & dateadd("yyyy",1,sHoliday) & "'"

		db.execute(sql)

		if sEveryYear = "1" then

			sql = "INSERT INTO	T_Holiday	( [Holiday], Holiday_Name, Every_Year )"
			sql = sql & "	Values ( '" & dateadd("yyyy",1,sHoliday) & "', '" & sHoliday_name & "', '" & sEveryYear & "')"

			db.execute(sql)	

		end if

		%>
			<script>
				alert('수정되었습니다');
				parent.goSearch(parent.document.inUpFrm);
			</script>
		<%

	elseif guboon = "INS" Then

		sql = "INSERT INTO	T_Holiday	( [Holiday], Holiday_Name, Every_Year )"
		sql = sql & "	Values ( '" & sHoliday & "', '" & sHoliday_name & "', '" & sEveryYear & "')"

		db.execute(sql)	

		'----------------------------------------------------------------
		'매년적용이라면 다음해 것도 등록한다.
		'----------------------------------------------------------------
		if sEveryYear = "1" then

			sql = "INSERT INTO	T_Holiday	( [Holiday], Holiday_Name, Every_Year )"
			sql = sql & "	Values ( '" & dateadd("yyyy",1,sHoliday) & "', '" & sHoliday_name & "', '" & sEveryYear & "')"

			db.execute(sql)	

		end if

		%>
			<script>
				alert('등록되었습니다');
				parent.goSearch(parent.document.inUpFrm);
			</script>
		<%
		
	end if
%>
<div id='CalendarLayer' style='display:none;'><iframe name='CalendarFrame' src='/include/Calendar.asp' width='172' height='177' border='0' frameborder='0' scrolling='no'></iframe></div>

<table width="100%" height="100" border="1" cellspacing="0" cellpadding="0" align="center">
	<tr>
		<td>
			<form method="post" name="inUpFrm" action="<%=currentURL%>" style="margin:0">
			<input type="hidden" name="guboon" value="">
			<table width="100%" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">
				<tr>
			        <td width="90" bgcolor="#EFEFEF" class="TDCont">일자</td>
			        <td bgcolor="#FFFFFF">
						<input name="Holiday" readonly type="text" size="10" onfocus="setFocusColor(this);" value="<%=sHoliday%>" <% if sHoliday = "" then %>onClick="new CalendarFrame.Calendar(this);" onblur="setOutColor(this);" <%end if%>>
			        </td>
				</tr>
				<tr>
			        <td width="90" bgcolor="#EFEFEF" class="TDCont">휴일명</td>
			        <td bgcolor="#FFFFFF"">
						<input name="Holiday_name" type="text" value="<%=sHoliday_name%>" onblur="setOutColor(this);">
			        </td>
				</tr>
				<tr>
			        <td width="90" bgcolor="#EFEFEF" class="TDCont">매년적용여부</td>
			        <td bgcolor="#FFFFFF"><input type="checkbox" name="EveryYear" value="1" class="none" <%if sEveryYear = "1" then %>checked<%end if%>></td>
			        </td>
				</tr>
			</table>
			</form>
			<table border="0" cellspacing="0" width="100%" align="center">
				<tr height="30">
					<td align="center">
						<img src="/Images/Btn/BtnSubmit.gif" style="cursor:hand;" align="absmiddle" onClick="javascript:fn_inup(document.inUpFrm);">
						<img src="/Images/Btn/BtnReset.gif" style="cursor:hand;" align="absmiddle" onClick="javascript:fn_reset();">
						<img src="/Images/Btn/BtnClose.gif" style="cursor:hand;" align="absmiddle" onClick="parent.HddnPOPLayer();">
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>


<script>
	function fn_inup(form){
		if ( form.Holiday.value == '' && "<%=sHoliday%>"=='' )
		{
			alert('☞ 등록하고자 하는 일자를 입력해 주십시오!');
			form.Holiday.focus();
			return;
		}
		if ( "<%=sHoliday%>"=='' )
			form.guboon.value ="INS";
		else
			form.guboon.value ="UP";
		form.submit();
	}
	function fn_reset(){
		if ("<%=sEveryYear%>"=='1')
			document.all.EveryYear.checked = "true";
		else
			document.all.EveryYear.checked = "";

		document.all.Holiday.value = "<%=sHoliday%>";
		document.all.Holiday_name.value = "<%=sHoliday_name%>";
	}
</script>