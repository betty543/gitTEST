<!-- #include virtual="/include/top_frame.asp" -->

<%
guboon = Request("guboon")
curPage = Request("curPage")
sCodegroup = Request("sCodegroup")
sGroupname = Request("sGroupname")
sCode = Request("sCode")

If guboon = "UP" Then

	SQL ="select * from TB_CODE  where CODEGROUP = '" & sCodegroup & "' and CODE = '" & sCode & "'"
	Set rs = db.execute(sql)

	if not rs.eof then 
		sCodename = rs("codename")
		sMemo = rs("memo")
		if sMemo <> "" then
			sMemo = replace(sMemo,  "<br>", chr(13)&chr(10))
		end if
		sUseyn = rs("useyn")
		sSysyn = rs("sysyn")
	end if

	rs.close
	set rs = Nothing
	
End if

%>
<script language="javascript">
<!--
	function fn_inup(f){
		if(!FieldChk(f.txtCode,"코드")) return false;
		if(!FieldChk(f.txtCodeName,"코드명")) return false;
		
		f.submit();
	}
-->
</script>
<table border="0" cellspacing="0" cellpadding="0" width="100%">
	<tr>
		<td width="100%">
<!--// //-->
<table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
<form name="inUpFrm" method="post" action="code_InsUpDel.asp">
<input type=hidden name=guboon value="<%=guboon%>">
<input type=hidden name=curPage value="<%=curPage%>">
<input type=hidden name=sCode value="<%=sCode%>">		
<input type=hidden name=sCodegroup value="<%=sCodegroup%>">
<input type=hidden name=sGroupname value="<%=sGroupname%>">

	<tr height="25" >
		<td width="10%" bgcolor="#F3F3F3" class="TDR10px"><font color="black">구분</font></td>
		<td width="40%" bgcolor="#FFFFFF" class="TDL5px"><input type="text" name="txtCodeGroup" value="<%=sCodegroup%>" maxlength="3" size="30" onfocus="setFocusColor(this)" onblur="setOutColor(this)"></td>
		<td width="10%" bgcolor="#F3F3F3" class="TDR10px" ><font color="black">구분명</font></td>
		<td width="40%" bgcolor="#FFFFFF" class="TDL5px"><input type="text" name="txtGroupName" value="<%=sGroupname%>" maxlength="25"  size="30" onfocus="setFocusColor(this)" onblur="setOutColor(this)"></td>
	</tr>
	<tr height="25">
		<td bgcolor="#F3F3F3" class="TDR10px"><font color="black">코드</font></td>
		<td bgcolor="#FFFFFF" class="TDL5px" colspan=3><input type="text" name="txtCode" value="<%=sCode%>" maxlength="11"  size="30" onfocus="setFocusColor(this)" onblur="setOutColor(this);" <%if guboon = "UP" then%>readonly><% end if %></td>

	</tr>
	<tr height="25">

		<td bgcolor="#F3F3F3" class="TDR10px"><font color="black">코드명</font></td>
		<td bgcolor="#FFFFFF" class="TDL5px" colspan=3><input type="text" name="txtCodeName" value="<%=sCodename%>" maxlength="70"  size="50" onfocus="setFocusColor(this)" onblur="setOutColor(this)"></td>
	</tr>
	<tr>
		<td bgcolor="#F3F3F3" class="TDR10px"><font color="black">메모</font></td>
		<td bgcolor="#FFFFFF" colspan="3" class="TDL5px"><textarea name="txtMemo" style="width:100%; height:120" wrap="soft" class="TextareaInput"><%=sMemo%></textarea></td>
	</tr>
	<tr height="25">
		<td bgcolor="#F3F3F3" class="TDR10px"><font color="black">사용여부</font></td>
		<td bgcolor="#FFFFFF" class="TDL5px">
			<input type="radio" name="optUseYN" value="Y" <% If sUseyn = "Y" Or sUseyn = "" Then response.write "checked" End If %> class="none">사용 
			<input type="radio" name="optUseYN" value="N" <% If sUseyn = "N" Then response.write "checked" End If %> class="none">미사용
		</td>
		<td bgcolor="#F3F3F3" class="TDR5px"><font color="black">시스템코드</font></td>
		<td bgcolor="#FFFFFF" class="TDL5px">
			<input type="radio" name="optSysYN" value="Y" <% If sSysyn = "Y" Or sSysyn = "" Then response.write "checked" End If %> class="none">사용 
			<input type="radio" name="optSysYN" value="N" <% If sSysyn = "N" Then response.write "checked" End If %> class="none">미사용
		</td>
	</tr>

</form>
</table>
<!--// //-->
		</td>
	</tr>
</table>

<table cellpadding="0" cellspacing="0" width="100%">
	<tr><td height="5"></td></tr>
	<tr>
		<td class="TDR10px">
			<img src="/Images/Btn/BtnSubmit.gif" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_inup(document.inUpFrm);">
			<img src="/Images/Btn/BtnReset.gif" style="cursor:hand;" align="absmiddle" onClick="javascript:history.back();">
		</td>
	</tr>
</table>
