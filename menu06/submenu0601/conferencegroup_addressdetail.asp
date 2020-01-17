<!-- #include virtual="/Include/Top_PopUp.asp" -->

<%
	'####### 파라미터 ##################################################################################
	idx = Trim(Request("idx"))		
	group_idx = Trim(Request("group_idx"))		

	'####### 디버깅 코드 ###############################################################################
	'Response.Write("tID=" &tID& "<br>")
	JobGb = Request("JobGb")

	If JobGb = "I" Then

		INCODE = SESSION("SS_LoginID")
		sname = Request("name")
		sclass = Request("class")
		sosok_name = Request("sosok_name")
		cellphone = Request("cellphone")
		gunphone = Request("gunphone")

		SQL = "INSERT INTO TB_SMSADDR ( group_idx, name, class, sosok_name, cellphone,gunphone,INCODE, INDATE, MOCODE, MODATE )"
		SQL = SQL & " VALUES ( " & group_idx & ",'"  & sname & "','"  & sclass & "','"  & sosok_name & "','"  & cellphone & "','"  & gunphone & "','"  & INCODE & "',getdate(),null,null)"

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
		sname = Request("name")
		sclass = Request("class")
		sosok_name = Request("sosok_name")
		cellphone = Request("cellphone")
		gunphone = Request("gunphone")
		armyno = Request("armyno")

		SQL = "UPDATE	TB_SMSADDR set  name = '"  & sname & "'"
		SQL = SQL & "	, class = '"  & sclass & "'"
		SQL = SQL & "	, sosok_name = '"  & sosok_name & "'"
		SQL = SQL & "	, cellphone = '"  & cellphone & "'"
		SQL = SQL & "	, gunphone = '"  & gunphone & "'"
		'SQL = SQL & "	, group_idx = '"  & group_idx & "'"
		SQL = SQL & "	, MOCODE='"  & INCODE & "', MODATE=getdate()"
		SQL = SQL & "	WHERE	idx = " & idx

		db.beginTrans
		db.execute(SQL)	



		if db.Errors.count = 0 then


			if armyno <> "" then
				'군번이 있는 경우 같은 군번이면 정보를 모두 수정한다.
				SQL = "UPDATE	TB_SMSADDR set  name = '"  & sname & "'"
				SQL = SQL & "	, class = '"  & sclass & "'"
				SQL = SQL & "	, sosok_name = '"  & sosok_name & "'"
				SQL = SQL & "	, cellphone = '"  & cellphone & "'"
				SQL = SQL & "	, gunphone = '"  & gunphone & "'"
				'SQL = SQL & "	, group_idx = '"  & group_idx & "'"
				SQL = SQL & "	, MOCODE='"  & INCODE & "', MODATE=getdate()"
				SQL = SQL & "	WHERE	idx <> " & idx
				SQL = SQL & "	AND		armyno = '" & armyno &"'"

				db.execute(SQL)

			end if
			db.CommitTrans

			Response.Write ("<script>alert('정상적으로 수정되었습니다!');parent.location.reload();</script>")	
			Response.End
		else
			db.RollBackTrans
			Call UrlBack("저장중 에러가 발생했습니다.\n\n다시 시도해 주세요")
		end if

	elseIf JobGb = "D" Then

		SQL = "DELETE	FROM TB_SMSADDR "
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


	if group_idx <> "" then

		SqlList = "SELECT *"
		SqlList = SqlList& " FROM TB_SMSGROUP"
		SqlList = SqlList& " WHERE idx="&group_idx

		SET Rs = db.execute(SqlList)
		if Rs.eof=false then
			groupname=rs("groupname")
		end if

		if idx <> "" then

			SqlList = "SELECT *"
			SqlList = SqlList& " FROM TB_SMSADDR"
			SqlList = SqlList& " WHERE idx="&idx

			SET Rs = db.execute(SqlList)
			if Rs.eof=false then

				armyno = rs("armyno")
				sname=rs("name")
				sclass=rs("class")
				cellphone=rs("cellphone")
				gunphone=rs("gunphone")
				sosok_name=rs("sosok_name")

			end if
			Rs.close
		end if

	end if



	'Response.Write(SqlList& "<br>")
	'SET RsList = db.execute(SqlList)  -- SQL QUERY
%>

<form name="DetailForm" style="margin:0">
<input value="<%=JOBGB%>" name="JOBGB" type="hidden" size="30">
<input value="<%=idx%>" name="idx" type="hidden" size="30">
<input value="<%=armyno%>" name="armyno" type="hidden" size="30">
<table border="0" width="490" cellpadding="0" cellspacing="0" align="center">
	<tr height="25"><td colspan="2" class="FBlk"> ◈ <b>다자간통화주소록</b></td></tr>
	<tr height="10"><td colspan="2"></td></tr>
	<tr bgcolor="#FFFFFF">
		<td colspan="2" height="110">
			<table width="100%" border="0" cellspacing="1" cellpadding="1" bgcolor="#CCCCCC">
				<tr >
					<td width="30%" bgcolor="#EEF6FF" >&nbsp;그룹</td>
					<td width="70%" bgcolor="#ffffff"><input value="<%=group_idx%>" name="group_idx" type="hidden" size="5" onfocus="setFocusColor(this);" onblur="setOutColor(this);" readonly><input value="<%=groupname%>" name="groupname" type="text" size="30" readonly></td>
				</tr>
				<tr >
					<td width="30%" bgcolor="#EEF6FF" >&nbsp;소속</td>
					<td width="70%" bgcolor="#ffffff"><input value="<%=sosok_name%>" name="sosok_name" type="text" size="30" onfocus="setFocusColor(this);" onblur="setOutColor(this);"></td>
				</tr>
				<tr >
					<td width="30%" bgcolor="#EEF6FF" >&nbsp;계급</td>
					<td width="70%" bgcolor="#ffffff"><input value="<%=sclass%>" name="class" type="text" size="20" onfocus="setFocusColor(this);" onblur="setOutColor(this);"></td>
				</tr>
				<tr >
					<td width="30%" bgcolor="#EEF6FF" >&nbsp;성명</td>
					<td width="70%" bgcolor="#ffffff"><input value="<%=sname%>" name="name" type="text" size="20" onfocus="setFocusColor(this);" onblur="setOutColor(this);"></td>
				</tr>
				<tr >
					<td width="30%" bgcolor="#EEF6FF" >&nbsp;군전화</td>
					<td width="70%" bgcolor="#ffffff"><input value="<%=gunphone%>" name="gunphone" type="text" size="15" onfocus="setFocusColor(this);" onblur="setOutColor(this);"></td>
				</tr>
				<tr >
					<td width="30%" bgcolor="#EEF6FF" >&nbsp;휴대폰번호</td>
					<td width="70%" bgcolor="#ffffff"><input value="<%=cellphone%>" name="cellphone" type="text" size="15" onfocus="setFocusColor(this);" onblur="setOutColor(this);"></td>
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
	if ( form.name.value ==  "" )
	{
		alert("성명을 입력하십시오!");
		return false;
	}
/*
	if ( form.cellphone.value ==  "" )
	{
		alert("휴대폰번호를 입력하십시오!");
		return false;
	}
	if ( form.cellphone.value.length < 10 )
	{
		alert("휴대폰번호를 정확히 입력하십시오!");
		return false;
	}
*/
	if ( form.idx.value ==  "" )
		form.JOBGB.value = "I";
	else
		form.JOBGB.value = "U";

	form.submit();
}

//-->
</script>
<!-- #include virtual="/Include/Bottom_PopUp.asp" -->