<!-- #include virtual="/Include/Top_PopUp.asp" -->

<%
	'####### 파라미터 ##################################################################################
	idx = Trim(Request("idx"))			

	'####### 디버깅 코드 ###############################################################################
	'Response.Write("tID=" &tID& "<br>")
	JobGb = Request("JobGb")

	If JobGb = "I" Then

		INCODE = SESSION("SS_LoginID")
		groupname = Request("groupname")
		useyn = Request("useyn")
		groupgb = Request("groupgb")

		SQL = "INSERT INTO TB_SMSGROUP ( groupname, groupgb, useyn, INCODE, INDATE, MOCODE, MODATE )"
		SQL = SQL & " VALUES ( '"  & groupname & "','"  & groupgb & "','"  & useyn & "','"  & INCODE & "',getdate(),null,null)"

		db.beginTrans
		db.execute(SQL)	

		if db.Errors.count = 0 then
			db.CommitTrans

			Response.Write ("<script>alert('정상적으로 등록되었습니다!');parent.location.reload();</script>")	
			Response.End
		else
			db.RollBackTrans
			Call UrlBack("저장중 에러가 발생했습니다.\n\n다시 시도해 주세요")
		end if
	elseIf JobGb = "U" Then

		INCODE = SESSION("SS_LoginID")
		groupname = Request("groupname")
		useyn = Request("useyn")
		groupgb = Request("groupgb")

		SQL = "UPDATE	TB_SMSGROUP set  groupname = '"  & groupname & "'"
		SQL = SQL & "	, groupgb= '"  & groupgb & "'"
		SQL = SQL & "	, useyn ='"  & useyn & "'"
		SQL = SQL & "	, MOCODE='"  & INCODE & "', MODATE=getdate()"
		SQL = SQL & "	WHERE	idx = " & idx

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

	elseIf JobGb = "D" Then

		SQL = "DELETE	FROM TB_SMSGROUP "
		SQL = SQL & "	WHERE	idx = " & idx

		db.beginTrans
		db.execute(SQL)	

		if db.Errors.count = 0 then
			db.CommitTrans

			Response.Write ("<script>alert('정상적으로 삭제되었습니다!');parent.location.reload();</script>")	
			Response.End
		else
			db.RollBackTrans
			Call UrlBack("저장중 에러가 발생했습니다.\n\n다시 시도해 주세요")
		end if

	End if


	if idx <> "" then

		SqlList = "SELECT *"
		SqlList = SqlList& " FROM TB_SMSGROUP"
		SqlList = SqlList& " WHERE idx="&idx


		SET Rs = db.execute(SqlList)
		if Rs.eof=false then

			groupgb = rs("groupgb")
			groupname=rs("groupname")
			useyn = rs("useyn")
			idx = rs("idx")
		end if
	
	else
		groupgb = "1"
		useyn = "Y"

	end if



	'Response.Write(SqlList& "<br>")
	'SET RsList = db.execute(SqlList)  -- SQL QUERY
%>

<form name="DetailForm" style="margin:0">
<input value="<%=JOBGB%>" name="JOBGB" type="hidden" size="30">
<table border="0" width="490" cellpadding="0" cellspacing="0" align="center">
	<tr height="25"><td colspan="2" class="FBlk"> ◈ <b>SMS그룹관리</b></td></tr>
	<tr height="10"><td colspan="2"></td></tr>
	<tr bgcolor="#FFFFFF">
		<td colspan="2" height="110">
			<table width="100%" border="0" cellspacing="1" cellpadding="1" bgcolor="#CCCCCC">
				<tr >
					<td width="30%" bgcolor="#EEF6FF" >&nbsp;그룹id</td>
					<td width="70%" bgcolor="#ffffff"><input value="<%=idx%>" name="idx" type="text" size="5" onfocus="setFocusColor(this);" onblur="setOutColor(this);" readonly></td>
				</tr>
				<tr>
					<td width="30%" bgcolor="#EEF6FF">&nbsp;그룹명</td>
					<td width="70%" bgcolor="#ffffff"><input value="<%=groupname%>" name="groupname" type="text" size="30" onfocus="setFocusColor(this);" onblur="setOutColor(this);"></td>
				</tr>
				<tr>
					<td width="30%" bgcolor="#EEF6FF">&nbsp;구분</td>
					<td width="70%" bgcolor="#ffffff"><input type="radio" name="groupgb" value="1" <%If groupgb = "1" Then response.write "checked" End If %> class="none"> 개인
						<input type="radio" name="groupgb" value="2" <%If groupgb = "2" Then response.write "checked" End If %> class="none" > 공용</td>
				</tr>
				<tr>
					<td width="30%" bgcolor="#EEF6FF">&nbsp;사용여부</td>
					<td width="70%" bgcolor="#ffffff"><input type="radio" name="useyn" value="Y" <%If useyn = "Y" Then response.write "checked" End If %> class="none"> 사용
						<input type="radio" name="useyn" value="N" <%If useyn = "N" Then response.write "checked" End If %> class="none" > 미사용</td>
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
	if ( form.groupname.value ==  "" )
	{
		alert("그룹명을 입력하십시오!");
		return false;
	}
	if ( form.idx.value ==  "" )
		form.JOBGB.value = "I";
	else
		form.JOBGB.value = "U";

	form.submit();
}

//-->
</script>
<!-- #include virtual="/Include/Bottom_PopUp.asp" -->