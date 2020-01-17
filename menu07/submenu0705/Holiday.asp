<!-- #include virtual="/Include/Top.asp" -->
<%
	'1. 파라미터 얻어오기
	'3. 쿼리 실행
	whereCD1 = Request("whereCD1")
	if whereCD1 = "" then
		whereCD1 = left(date(),4)
	end if

	'4. Paging HTML 작성

%>


<table border="0" width="940" cellpadding="0" cellspacing="3" bgcolor="#EFECE5" align="center">
	<tr bgcolor="#FFFFFF">
		<td>
			<form method="post" name="inUpFrm" action="<%=currentURL%>" style="margin:0">
			<table width="100%" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">
				<tr>
			        <td width="90" bgcolor="#EFEFEF" class="TDCont">년도 :</td>
			        <td bgcolor="#FFFFFF" width="100">
						<select name="whereCD1" size="1" class="ComboFFFCE7">
							<% for i = 10 to -1 step -1 %>
								
								<option value="<%=left(dateadd("yyyy",i*-1,now()),4)%>" <%if whereCD1 = left(dateadd("yyyy",i*-1,now()),4) then %>selected<%end if%> ><%=left(dateadd("yyyy",i*-1,now()),4)%></option>
							<% next %>
						</select>
			        </td>
			        <td colspan='2' bgcolor="#FFFFFF" align="right">
			        	<img src="/Images/Btn/BtnSearch.gif" style="cursor:hand;" onClick="goSearch(document.inUpFrm);">&nbsp;<img src="/Images/Btn/BtnPlus.gif" style="cursor:hand;" onClick="fn_popVIEW('','','0');">
			        </td>
				</tr>
			</table>
			</form>
		</td>
	</tr>
</table>

<table width="940" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>

<table width="940" border="0" cellspacing="0" cellpadding="0" align="center">
	<tr valign="top">
		<td>
        	<table width="100%" border="0" cellspacing="0" cellpadding="0">
        		<tr><td height="300"><iframe src="Holiday_List.asp?GijunYear=<%=whereCD1%>" scrolling="auto" frameborder="0" border="0" width="100%" height="100%" name="HolidayFrame"></iframe></td></tr>
        	</table>
		</td>
	</tr>
</table>

<!-- #include virtual="/Include/PopLayer.asp" -->
<!-- #include virtual="/Include/Bottom.asp" -->

<script>

	function goSearch(form)
	{
		form.submit();
	}
	 
	function fn_popVIEW(sHoliday,sHolidayName,sEveryYear){
		GoURL ="/menu07/submenu0705/holiday_Detail.asp?Holiday=" +sHoliday+"&Holiday_Name="+sHolidayName+"&EveryYear="+sEveryYear;
		ShowPOPLayer(GoURL,'300','250');
	}
		
</script>