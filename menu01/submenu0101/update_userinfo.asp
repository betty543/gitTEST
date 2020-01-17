<!-- #include virtual="/Include/Top_PopUp.asp" -->

<%
	'####### 파라미터 ##################################################################################
	
	receiptfactnum = Trim(Request("receiptfactnum"))		

	'####### 디버깅 코드 ###############################################################################
	'Response.Write("tID=" &tID& "<br>")
	JobGb = Request("JobGb")
	dutyman = Request("dutyman")


	If JobGb = "U" Then

		SQL = "	update armyinformix.dbo.receiptfact set dutyman = '" & dutyman & "'"
		SQL = SQL & " where receiptfactnum = '" & receiptfactnum & "'"

		db.beginTrans
		db.execute(SQL)	

		if db.Errors.count = 0 then
			db.CommitTrans

			Response.Write ("<script>alert('정상적으로 수정되었습니다!');parent.location.reload();</script>")	
			Response.End
		else
			db.RollBackTrans
			Call UrlBack("저장중 에러가 발생했습니다.\n\n다시 시도해 주세요")
		end if

	End if


%>

<form name="DetailForm" style="margin:0">
<input value="<%=JOBGB%>" name="JOBGB" type="hidden" size="30">
<input value="<%=receiptfactnum%>" name="receiptfactnum" type="hidden" size="30">
<table border="0" width="300" cellpadding="0" cellspacing="0" align="center">
	<tr height="25"><td colspan="2" class="FBlk"> ◈ <b>수사관코드 수정</b></td></tr>
	<tr height="10"><td colspan="2"></td></tr>
	<tr bgcolor="#FFFFFF">
		<td colspan="2" height="80">
			<table width="100%" border="0" cellspacing="1" cellpadding="1" bgcolor="#CCCCCC">
				<tr >
					<td width="30%" bgcolor="#EEF6FF" >&nbsp;수사관코드</td>
					<td width="70%" bgcolor="#ffffff"><select name="dutyman" size="1" class="ComboFFFCE7">
						<option value="">선택</option>
<%					
							SQL = "	select * from armyinformix.dbo.user1 where unit = '" & left(receiptfactnum,4) & "' order by name" '수사관정보
							SET Rs = DB.execute(SQL)
							do until Rs.eof
									CODE = Rs("id")
									CODENAME = Rs("name") & Rs("class")
								%>

									<%=printSelect("" &CODENAME& "","" &CODE& "","" &dutyman& "")%>
								<%
								Rs.movenext
							loop
%>
						</select></td>
				</tr>
            </table>
		</td>
	</tr>
	<tr height="5"><td colspan="2"></td></tr>
	<tr height="30">
		<td align="right" height="35">
			<img src="/Images/Btn/BtnSubmit.gif" style="cursor:hand;" align="absmiddle" onClick="javascript:fn_inup(document.DetailForm);">
			<img src="/Images/Btn/BtnClose.gif" style="cursor:hand;" align="absmiddle" onClick="javascript:parent.HddnPOPLayer();">
		</td>
	</tr>
</table>

</form>

<script>
<!--

function fn_inup(form)
{

	form.JOBGB.value = "U";

	form.submit();
}

//-->
</script>
<!-- #include virtual="/Include/Bottom_PopUp.asp" -->