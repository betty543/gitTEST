<!-- #include virtual="/include/top_frame.asp" -->
<%
	sql = "SELECT USERID, USERNAME, GRADE, 'Y' CALLBACKYN, 'A' MANUFACTURE FROM TB_USERINFO WHERE USEYN='Y'"
	sql = sql& "ORDER BY USEYN DESC, GRADE ASC, USERNAME ASC"
	set Rs = db.execute(sql)
%>
<script>
<!--
	function fn_ChkDisabled(f,cn){
		obj1 = eval("ListForm.Chk["+f+"]");
		obj2 = eval("ListForm."+cn);

		for(i=0;i<obj2.length;i++){
			if(obj1.checked){
				obj2[i].disabled = false;
				obj2[i].checked = false;
			}else{
				obj2[i].disabled = true;
				obj2[i].checked = false;
			}
		}
	}
//-->
</script>
<table width="765" border="0" cellspacing="1" cellpadding="1" bgcolor="#CCCCCC">
<form name="ListForm" method="post">
<%
	if NOT Rs.EOF then
		i = 1
		Do until Rs.EOF
			sUserID = Rs("USERID")
			sUserName = Rs("USERNAME")
			sGrade = Rs("GRADE")
			if sGrade <> "" then
				sGradeN = db_getCodeName("Z03",sGrade)
			end if
			sCallBack = Rs("CALLBACKYN")
			sMANUFACTURE = Rs("MANUFACTURE")
			IF sCallBack = "Y" THEN
				tempBG="#FFEEF9"
				sCallBackCHK = "checked"
			ELSE
				tempBG="#FFFFFF"
				sCallBackCHK = ""
			END IF
%>
	<tr height="20" bgcolor="<%=tempBG%>" onmouseover="this.style.background='#FFFCE7'" onmouseout="this.style.background='<%=tempBG%>'">
		<td width="40" align="center"><%=i%></td>
		<td width="100" align="center"><%=sUserID%></td>
		<td width="100" align="center"><%=sUserName%></td>
		<td width="100" align="center"><%=sGradeN%></td>
		<td align="center"><input type="checkbox" name="Chk" value="<%=sUserID%>" <%=sCallBackCHK%> class="none" onClick="fn_ChkDisabled('<%=i-1%>',this.value);"></td>
		<td width="375">

			<%
        		SQLC = "SELECT CODE, CODENAME FROM TB_CODE WHERE CODEGROUP = 'A01' ORDER BY CODE ASC"
				Set RsC = db.execute(SQLC)

					IF sCallBack = "Y" THEN
						if sMANUFACTURE <> "" then
							sMANUFACTURE_ok = split(sMANUFACTURE,",")
							if NOT(rsC.EOF or rsC.BOF) then
								Do until rsC.EOF
									selects = "false"
									for j = 0 to UBound(sMANUFACTURE_ok)
										if RsC("CODE") = sMANUFACTURE_ok(j) then selects = "true" end if
									NEXT
									if selects = "true" then
										Response.Write(printChk(sUserID,RsC("CODENAME"),RsC("CODE"),RsC("CODE"),""))
									else
										Response.Write(printChk(sUserID,RsC("CODENAME"),RsC("CODE"),"",""))
									end if
									RsC.Movenext
								Loop
							End if
						else
							if NOT(rsC.EOF or rsC.BOF) then
								Do until rsC.EOF
									Response.Write(printChk(sUserID,RsC("CODENAME"),RsC("CODE"),"",""))
									RsC.Movenext
								Loop
							End if
						end if
					ELSE
						if NOT(rsC.EOF or rsC.BOF) then
							Do until rsC.EOF
								Response.Write(printChk(sUserID,RsC("CODENAME"),RsC("CODE"),"","false"))
								RsC.Movenext
							Loop
						End if
					END IF

				rsC.Close
				Set rsC = Nothing
        	%>
		</td>
	</tr>
<%
			i = i + 1
			Rs.Movenext
		Loop
	else
%>
	<tr height="20" bgcolor="#FFFFFF" onmouseover="this.style.background='#FFFCE7'" onmouseout="this.style.background='#FFFFFF'">
		<td align="center" colspan="6">등록된 사용자가 없습니다.</td>
	</tr>
<%
	End if


	Rs.Close
	Set Rs = Nothing
%>
</form>
</table>
