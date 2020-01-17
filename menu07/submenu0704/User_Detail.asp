<!-- #include virtual="/include/top_frame.asp" -->



<%

guboon = request("guboon")
userid = request("userid")

If guboon = "UP" Then

	sql ="select * from TB_USERINFO where userid = '" & userid & "'"
	Set rs = db.Execute(sql)

	If not rs.eof Then   
		sUSERID = rs("USERID")
		sUSERNAME = rs("USERNAME")
		sPASSWORD = rs("PASSWORD")
		sSECGROUP = rs("SECGROUP")
		sGRADE = rs("GRADE")
		sUSEYN = rs("USEYN")
		sIPDATE = rs("IPDATE")
		sOUTDATE = rs("OUTDATE")
		sCTIYN = rs("CTIYN")
		sCTIID = rs("CTIID")
		sCTIPASSWORD = rs("CTIPASSWORD")
		sEXTNO = rs("EXTNO")
		sSOSOK = rs("SOSOK")
		sLEVEL = rs("LEVEL")
	End If

End if

%>

<div id='CalendarLayer' style='display:none;'><iframe name='CalendarFrame' src='/include/Calendar.asp' width='172' height='177' border='0' frameborder='0' scrolling='no'></iframe></div>


        	<table width="100%" border="0" cellspacing="0" cellpadding="0">
        		<tr><td height="22" colspan="2" class="FBlk">◈ <b>사용자 정보</b></td></tr>
        	</table>

			<table width="100%" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">


<form name="inUpFrm" method="post" action="User_InsUpDel.asp">
	<input type=hidden name=guboon value="<%=guboon%>">

				<tr>
					<td nowrap width="100" bgcolor="#FFEEF9" class="TDCont">아이디</td>
					<td bgcolor="#FFFFFF"><input type="text" name="sUSERID" value="<%=sUSERID%>" <%If guboon = "UP" Then response.write "readonly" End If %> maxlength="20" size="20" onfocus="setFocusColor(this);" onblur="setOutColor(this);" align="absmiddle"></td>
				</tr>
				<tr>
					<td bgcolor="#FFEEF9" class="TDCont">성명</td>
					<td bgcolor="#FFFFFF"><input type="text" name="sUSERNAME" value="<%=sUSERNAME%>" maxlength="20" size="20" onfocus="setFocusColor(this);" onblur="setOutColor(this);" align="absmiddle"></td>
				</tr>
				<tr>
					<td bgcolor="#FFEEF9" class="TDCont">비밀번호</td>
					<td bgcolor="#FFFFFF"><input type="text" name="sPASSWORD" value="<%=sPASSWORD%>" maxlength="20" size="20" onfocus="setFocusColor(this);" onblur="setOutColor(this);" align="absmiddle"></td>
				</tr>
				<tr>
					<td bgcolor="#FFEEF9" class="TDCont">소속</td>
					<td bgcolor="#FFFFFF">
						<select name="sSOSOK" size="1" class="ComboFFFCE7">
							<%=db_getTBCodeSelect("C04", sSOSOK, "N")%>
						</select>					
					</td>
				</tr>
				<tr>
					<td bgcolor="#FFEEF9" class="TDCont">계급</td>
					<td bgcolor="#FFFFFF">
						<select name="sLEVEL" size="1" class="ComboFFFCE7">
							<%=db_getTBCodeSelect("Z05", sLEVEL, "N")%>
						</select>					
					</td>
				</tr>
				<tr>
					<td bgcolor="#FFEEF9" class="TDCont">군번</td>
					<td bgcolor="#FFFFFF"><input type="text" name="sGUNNUMBER" value="<%=sGUNNUMBER%>" maxlength="20" size="20" onfocus="setFocusColor(this);" onblur="setOutColor(this);" align="absmiddle"></td>
				</tr>
				<tr>
					<td bgcolor="#FFEEF9" class="TDCont">보안그룹</td>
					<td bgcolor="#FFFFFF">
						<select name="sSECGROUP" size="1" class="ComboFFFCE7">
							<%=db_getTBCodeSelect("Z02", sSECGROUP, "N")%>
						</select>					
					</td>
				</tr>
				<tr>
					<td bgcolor="#FFEEF9" class="TDCont">운영업무</td>
					<td bgcolor="#FFFFFF">
						<select name="sGRADE" size="1" class="ComboFFFCE7">
							<%=db_getTBCodeSelect("Z04", sGRADE, "N")%>
						</select>					
					</td>
				</tr>				
				<tr>
					<td bgcolor="#FFEEF9" class="TDCont">CTI 사용여부</td>
					<td bgcolor="#FFFFFF">
						<input type="radio" name="sCTIYN" value="Y" class="none" <% If sCTIYN = "Y" Then response.write "checked" End If %>> 사용
						<input type="radio" name="sCTIYN" value="N" class="none" <% If sCTIYN = "N" Or sCTIYN = "" Then response.write "checked" End If %>> 미사용
					</td>
				</tr>

				<tr>
					<td bgcolor="#FFEEF9" class="TDCont">재직여부</td>
					<td bgcolor="#FFFFFF">
						<input type="radio" name="sUSEYN" value="Y" class="none" onClick="fn_YES();" <% If sUSEYN = "Y" Or sUSEYN = "" Then response.write "checked" End If %>> 재직
						<input type="radio" name="sUSEYN" value="N" class="none" onClick="fn_YES();" <% If sUSEYN = "N" Then response.write "checked" End If %>> 퇴직
					</td>
				</tr>				
				<tr>
					<td bgcolor="#FFEEF9" class="TDCont">시작일자</td>
					<td bgcolor="#FFFFFF"><input name="sIPDATE" value="<%=sIPDATE%>" readonly type="text" size="10" onfocus="setFocusColor(this); new CalendarFrame.Calendar(this);" onblur="setOutColor(this);"></td>
				</tr>
				<tr>
					<td bgcolor="#FFEEF9" class="TDCont">퇴직일자</td>
					<td bgcolor="#FFFFFF"><input name="sOUTDATE" value="<%=sOUTDATE%>" readonly type="text" size="10" onfocus="setFocusColor(this); new CalendarFrame.Calendar(this);" onblur="setOutColor(this);"></td>
				</tr>
		<input type="hidden" name="sCTIID" value="<%=sCTIID%>" maxlength="20" size="20" onfocus="setFocusColor(this);" onblur="setOutColor(this);" align="absmiddle">
		<input type="hidden" name="sCTIPASSWORD" value="<%=sCTIPASSWORD%>" maxlength="20" size="20" onfocus="setFocusColor(this);" onblur="setOutColor(this);" align="absmiddle"><input type="hidden" name="sEXTNO" value="<%=sEXTNO%>" maxlength="4" size="4" onfocus="setFocusColor(this);" onblur="setOutColor(this);" align="absmiddle">
</form>

			</table>
			<table border="0" cellspacing="0" width="100%" align="center">
				<tr height="30">
					<td align="right">
						<img src="/Images/Btn/BtnSubmit.gif" style="cursor:hand;" align="absmiddle" onClick="javascript:fn_inup(document.inUpFrm);">
						<img src="/Images/Btn/BtnReset.gif" style="cursor:hand;" align="absmiddle" onClick="javascript:document.inUpFrm.reset();">
						<img src="/Images/Btn/BtnDel.gif" style="cursor:hand;" align="absmiddle" onClick="javascript:fn_del();">
					</td>
				</tr>
			</table>			



<script>
function fn_inup(inUpFrm) {

	if(!FieldChk(inUpFrm.sUSERID,"아이디")) return;
	if(!FieldChk(inUpFrm.sUSERNAME,"성명")) return;
	if(!FieldChk(inUpFrm.sPASSWORD,"비밀번호")) return;

	if (inUpFrm.sUSEYN(0).checked && inUpFrm.sIPDATE.value =='')
	{
		alert('입사일자를 입력하십시오!')
		return;
	}	

	/*if (inUpFrm.sUSEYN(0).checked && inUpFrm.sOUTDATE.value !='')
	{
		alert('재직여부의 재직과 미사용일자를 동시에 입력할 수 없습니다!')
		return;
	}	

	if (inUpFrm.sUSEYN(1).checked && inUpFrm.sOUTDATE.value =='')
	{
		alert('재직여부를 퇴사로 선택하셨습니다. 퇴사일자를 입력하십시오!')
		return;
	}	*/
	
	if(confirm("저장하시겠습니까?"))
		inUpFrm.submit();
	else
		return;
}
function fn_del() {
	if(confirm("삭제하시겠습니까?"))
		location.href = "User_InsUpDel.asp?guboon=DEL&sUSERID=<%=userid%>";
	else
		return;
}

function fn_YES() {

	if (inUpFrm.sUSEYN(0).checked)
	{
		inUpFrm.sOUTDATE.value ="";
	}

}

</script>