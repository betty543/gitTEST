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
	if(!FieldChk(f.txtDNIS,"내선번호")) return false;
	if(!FieldChk(f.txtTransferno,"착신번호")) return false;
	if(!FieldChk(f.txtStartTime,"착신시작시각")) return false;
	if(!FieldChk(f.txtEndTime,"착신종료시각")) return false;		
	f.submit();
}
</script>
<table width="1000" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr><td height="10"></td></tr>
    <tr>
    	<td>
<!-- 프로그램 입력 폼 START -->
<form name="frmBody" method="post" action="transferinfo_InsUpDel.asp">
<input type=hidden name="curPage" value="<%=curPage%>">
<input type=hidden name="guboon" value="<%=guboon%>">
<input type=hidden name="seqno" value="<%=seqno%>">
<table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
	<tr>
		<td bgcolor="#F3F3F3" width="70" class="TDR10px"><font color="black">내선번호</font></td>
		<td bgcolor="#FFFFFF" width="230" class="TDL5px"><input name="txtDNIS" type="text" value="<%=sDNIS%>" class="input"  size="25" onfocus="setFocusColor(this)" onblur="setOutColor(this)" maxlength="4"></td>
	</tr>
	<tr>
		<td bgcolor="#F3F3F3" width="70" class="TDR10px"><font color="black">착신번호</font></td>
		<td bgcolor="#FFFFFF" width="230" class="TDL5px"><input name="txtTransferno" type="text" value="<%=sTransferno%>" class="input"  size="25" onfocus="setFocusColor(this)" onblur="setOutColor(this)" maxlength="20"></td>
	</tr>
	<tr>
		<td bgcolor="#F3F3F3" width="70" class="TDR10px"><font color="black">착신시간</font></td>
		<td bgcolor="#FFFFFF" class="TDL5px"><input name="txtStartTime" type="text" value="<%=sStartTime%>" class="input"  size="4" onfocus="setFocusColor(this)" onblur="setOutColor(this)" maxlength="4">~<input name="txtEndTime" type="text" value="<%=sEndTime%>" class="input"  size="4" onfocus="setFocusColor(this)" onblur="setOutColor(this)" maxlength="4">&nbsp;&nbsp;※시간입력은 24시기준임. 오전9시 = 0900, 오후3시30분 = 1530</td>
	</tr>
	<tr height="25">
		<td bgcolor="#F3F3F3" width="70" class="TDR10px"><font color="black">권한</font></td>
		<td bgcolor="#FFFFFF" colspan="3" class="TDL5px">
			<input type="checkbox" name="chkUseyn" value="사용" class="none" <% if sUseyn="1" then Response.Write("checked") end if %>>사용&nbsp;&nbsp;&nbsp;&nbsp;     
			<input type="checkbox" name="chkMon" value="월요일" class="none" <% if sMon="1" then Response.Write("checked") end if %>>월요일 &nbsp;&nbsp;&nbsp;&nbsp;
			<input type="checkbox" name="chkTue" value="화요일" class="none" <% if sTue="1" then Response.Write("checked") end if %>>화요일 &nbsp;&nbsp;&nbsp;&nbsp;
			<input type="checkbox" name="chkWed" value="수요일" class="none" <% if sWed="1" then Response.Write("checked") end if %>>수요일 &nbsp;&nbsp;&nbsp;&nbsp;
			<input type="checkbox" name="chkThu" value="목요일" class="none" <% if sThu="1" then Response.Write("checked") end if %>>목요일 &nbsp;&nbsp;&nbsp;&nbsp;
			<input type="checkbox" name="chkFri" value="금요일" class="none" <% if sFri="1" then Response.Write("checked") end if %>>금요일 &nbsp;&nbsp;&nbsp;&nbsp;
			<input type="checkbox" name="chkSta" value="토요일" class="none" <% if sSta="1" then Response.Write("checked") end if %>>토요일 &nbsp;&nbsp;&nbsp;&nbsp;
			<input type="checkbox" name="chkSun" value="일요일" class="none" <% if sSun="1" then Response.Write("checked") end if %>>일요일 &nbsp;&nbsp;&nbsp;&nbsp;
			<input type="checkbox" name="chkHoliday" value="휴일" class="none" <% if sHoliday="1" then Response.Write("checked") end if %>>휴일 &nbsp;&nbsp;&nbsp;&nbsp;
		</td>
	</tr>
	<tr>
		<td bgcolor="#F3F3F3" width="70" class="TDR10px"><font color="black">통화대상</font></td>
		<td bgcolor="#FFFFFF" width="230" class="TDL5px"><%
							'======= 처리구분 코드 가져오기 ==================================================
							SqlCode = "SELECT USERID, USERNAME FROM TB_USERINFO"
							SqlCode = SqlCode& " WHERE USEYN='Y'"
							if SS_Login_Grade <> "A" then
								SqlCode = SqlCode& " AND	GRADE='"&SS_Login_Grade&"'" '생명의전화 그룹
							end if
							'if SS_Login_Secgroup = "A" then
								'내것만
								'SqlCode = SqlCode& " AND	USERID = '"&SS_LoginID&"'"
							'end if
							SqlCode = SqlCode& " ORDER BY USERID"
							set RsCode = db.execute(SqlCode)
						%>
						&nbsp;<select name="userid" size="1" class="ComboFFFCE7">
							<Option value ='' selected>통화대상선택</option>
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

<!-- 프로그램 입력 폼 END -->
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