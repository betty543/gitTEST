<!-- #include virtual="/Include/Top_PopUp.asp" -->

<%
	'####### �Ķ���� ##################################################################################
	
	receiptfactnum = Trim(Request("receiptfactnum"))		

	'####### ����� �ڵ� ###############################################################################
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

			Response.Write ("<script>alert('���������� �����Ǿ����ϴ�!');parent.location.reload();</script>")	
			Response.End
		else
			db.RollBackTrans
			Call UrlBack("������ ������ �߻��߽��ϴ�.\n\n�ٽ� �õ��� �ּ���")
		end if

	End if


%>

<form name="DetailForm" style="margin:0">
<input value="<%=JOBGB%>" name="JOBGB" type="hidden" size="30">
<input value="<%=receiptfactnum%>" name="receiptfactnum" type="hidden" size="30">
<table border="0" width="300" cellpadding="0" cellspacing="0" align="center">
	<tr height="25"><td colspan="2" class="FBlk"> �� <b>������ڵ� ����</b></td></tr>
	<tr height="10"><td colspan="2"></td></tr>
	<tr bgcolor="#FFFFFF">
		<td colspan="2" height="80">
			<table width="100%" border="0" cellspacing="1" cellpadding="1" bgcolor="#CCCCCC">
				<tr >
					<td width="30%" bgcolor="#EEF6FF" >&nbsp;������ڵ�</td>
					<td width="70%" bgcolor="#ffffff"><select name="dutyman" size="1" class="ComboFFFCE7">
						<option value="">����</option>
<%					
							SQL = "	select * from armyinformix.dbo.user1 where unit = '" & left(receiptfactnum,4) & "' order by name" '���������
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