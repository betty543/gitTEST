<!-- #include virtual="/include/top.asp" -->

<%
	guboon = Request("guboon")
	curPage = Request("curPage")
	seqno = Request("seqno")
	SS_Login_Grade = SESSION("SS_Login_Grade")
	
	If guboon = "UP" Then

		SQL ="SELECT * FROM TB_TRANSFERINFO WHERE seqno = '" & seqno & "'"
		Set rs = db.execute(sql)
	
		if not rs.eof then 
			seqno = rs("seqno")
			sdnis = rs("dnis")
			sstarttime = rs("starttime")
			sendtime = rs("endtime")
			stransferno = rs("transferno")
			smon = rs("mon")
			stue = rs("tue")
			swed = rs("wed")
			sthu = rs("thu")
			sfri = rs("fri")
			ssta = rs("sta")
			ssun = rs("sun")
			sholiday = rs("holiday")
			sUseyn = rs("USEYN")
			suserid = rs("userid")
		end if
	
		rs.close
		set rs = Nothing
		
	End if
%>
<script language="javascript">

function fn_inup(f)
{
	if(!FieldChk(f.txtDNIS,"������ȣ")) return false;
	if(!FieldChk(f.txtTransferno,"���Ź�ȣ")) return false;
	if(!FieldChk(f.txtStartTime,"���Ž��۽ð�")) return false;
	if(!FieldChk(f.txtEndTime,"��������ð�")) return false;		
	f.submit();
}
</script>
<table width="1000" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr><td height="10"></td></tr>
    <tr>
    	<td>
<!-- ���α׷� �Է� �� START -->
<form name="frmBody" method="post" action="transferinfo_InsUpDel.asp">
<input type=hidden name="curPage" value="<%=curPage%>">
<input type=hidden name="guboon" value="<%=guboon%>">
<input type=hidden name="seqno" value="<%=seqno%>">
<table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
	<tr>
		<td bgcolor="#F3F3F3" width="70" class="TDR10px"><font color="black">������ȣ</font></td>
		<td bgcolor="#FFFFFF" width="230" class="TDL5px"><input name="txtDNIS" type="text" value="<%=sDNIS%>" class="input"  size="25" onfocus="setFocusColor(this)" onblur="setOutColor(this)" maxlength="4"></td>
	</tr>
	<tr>
		<td bgcolor="#F3F3F3" width="70" class="TDR10px"><font color="black">���Ź�ȣ</font></td>
		<td bgcolor="#FFFFFF" width="230" class="TDL5px"><input name="txtTransferno" type="text" value="<%=sTransferno%>" class="input"  size="25" onfocus="setFocusColor(this)" onblur="setOutColor(this)" maxlength="20"></td>
	</tr>
	<tr>
		<td bgcolor="#F3F3F3" width="70" class="TDR10px"><font color="black">���Žð�</font></td>
		<td bgcolor="#FFFFFF" class="TDL5px"><input name="txtStartTime" type="text" value="<%=sStartTime%>" class="input"  size="4" onfocus="setFocusColor(this)" onblur="setOutColor(this)" maxlength="4">~<input name="txtEndTime" type="text" value="<%=sEndTime%>" class="input"  size="4" onfocus="setFocusColor(this)" onblur="setOutColor(this)" maxlength="4">&nbsp;&nbsp;�ؽð��Է��� 24�ñ�����. ����9�� = 0900, ����3��30�� = 1530</td>
	</tr>
	<tr height="25">
		<td bgcolor="#F3F3F3" width="70" class="TDR10px"><font color="black">����</font></td>
		<td bgcolor="#FFFFFF" colspan="3" class="TDL5px">
			<input type="checkbox" name="chkUseyn" value="���" class="none" <% if sUseyn="1" then Response.Write("checked") end if %>>���&nbsp;&nbsp;&nbsp;&nbsp;     
			<input type="checkbox" name="chkMon" value="������" class="none" <% if sMon="1" then Response.Write("checked") end if %>>������ &nbsp;&nbsp;&nbsp;&nbsp;
			<input type="checkbox" name="chkTue" value="ȭ����" class="none" <% if sTue="1" then Response.Write("checked") end if %>>ȭ���� &nbsp;&nbsp;&nbsp;&nbsp;
			<input type="checkbox" name="chkWed" value="������" class="none" <% if sWed="1" then Response.Write("checked") end if %>>������ &nbsp;&nbsp;&nbsp;&nbsp;
			<input type="checkbox" name="chkThu" value="�����" class="none" <% if sThu="1" then Response.Write("checked") end if %>>����� &nbsp;&nbsp;&nbsp;&nbsp;
			<input type="checkbox" name="chkFri" value="�ݿ���" class="none" <% if sFri="1" then Response.Write("checked") end if %>>�ݿ��� &nbsp;&nbsp;&nbsp;&nbsp;
			<input type="checkbox" name="chkSta" value="�����" class="none" <% if sSta="1" then Response.Write("checked") end if %>>����� &nbsp;&nbsp;&nbsp;&nbsp;
			<input type="checkbox" name="chkSun" value="�Ͽ���" class="none" <% if sSun="1" then Response.Write("checked") end if %>>�Ͽ��� &nbsp;&nbsp;&nbsp;&nbsp;
			<input type="checkbox" name="chkHoliday" value="����" class="none" <% if sHoliday="1" then Response.Write("checked") end if %>>���� &nbsp;&nbsp;&nbsp;&nbsp;
		</td>
	</tr>
	<tr>
		<td bgcolor="#F3F3F3" width="70" class="TDR10px"><font color="black">��ȭ���</font></td>
		<td bgcolor="#FFFFFF" width="230" class="TDL5px"><%
							'======= ó������ �ڵ� �������� ==================================================
							SqlCode = "SELECT USERID, USERNAME FROM TB_USERINFO"
							SqlCode = SqlCode& " WHERE USEYN='Y'"
							if SS_Login_Grade <> "A" then
								SqlCode = SqlCode& " AND	GRADE='"&SS_Login_Grade&"'" '��������ȭ �׷�
							end if
							'if SS_Login_Secgroup = "A" then
								'���͸�
								'SqlCode = SqlCode& " AND	USERID = '"&SS_LoginID&"'"
							'end if
							SqlCode = SqlCode& " ORDER BY USERID"
							set RsCode = db.execute(SqlCode)
						%>
						&nbsp;<select name="userid" size="1" class="ComboFFFCE7">
							<Option value ='' selected>��ȭ�����</option>
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("USERID")
										CODENAME = RsCode("USERNAME")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &suserid& "")%>
							<%
									RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
							%>
						</select></td>
	</tr>
</table>
</form>

<!-- ���α׷� �Է� �� END -->
		</td>
	</tr>
</table>
<table width="1000" border="0" cellspacing="0" cellpadding="0" align="center">
	<tr><td height="5"></td></tr>
	<tr>
		<td class="TDR10px">
			<img src="/Images/Btn/BtnSubmit.gif" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_inup(document.frmBody);">
			<img src="/Images/Btn/BtnReset.gif" style="cursor:hand;" align="absmiddle" onclick="javascript:history.back();">
		</td>
	</tr>
</table>
<!-- #include virtual="/include/bottom.asp" -->