<!-- #include virtual="/include/top_frame.asp" -->



<%

guboon = request("guboon")
idx = request("idx")

If guboon = "UP" Then

	sql ="select * from TB_Reject where idx = '" & idx & "'"
	Set rs = db.Execute(sql)

	If not rs.eof Then   
		sIdx = rs("idx")
		sDNIS = rs("DNIS")
		sTelNo = rs("TelNo")
		sUSEYN = rs("USEYN")
	End If

End if

%>

<div id='CalendarLayer' style='display:none;'><iframe name='CalendarFrame' src='/include/Calendar.asp' width='172' height='177' border='0' frameborder='0' scrolling='no'></iframe></div>


        	<table width="100%" border="0" cellspacing="0" cellpadding="0">
        		<tr><td height="22" colspan="2" class="FBlk">�� <b>�� ����</b></td></tr>
        	</table>

			<table width="100%" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">


<form name="inUpFrm" method="post" action="Callreject_InsUpDel.asp">
	<input type=hidden name=guboon value="<%=guboon%>">

				<tr>
					<td nowrap width="100" bgcolor="#FFEEF9" class="TDCont">����</td>
					<td bgcolor="#FFFFFF"><input type="text" name="Idx" value="<%=sIdx%>" readonly maxlength="3" size="3" onfocus="setFocusColor(this);" onblur="setOutColor(this);" align="absmiddle"><-�ڵ��ο���</td>
				</tr>
				<tr>
					<td bgcolor="#FFEEF9" class="TDCont">������ȣ</td>
					<td bgcolor="#FFFFFF"><input type="text" name="DNIS" value="<%=sDNIS%>" maxlength="20" size="20" onfocus="setFocusColor(this);" onblur="setOutColor(this);" align="absmiddle"></td>
				</tr>


				<tr>
					<td bgcolor="#FFEEF9" class="TDCont">���Űźι�ȣ</td>
					<td bgcolor="#FFFFFF"><input type="text" name="TelNo" value="<%=sTelNo%>" maxlength="20" size="20" onfocus="setFocusColor(this);" onblur="setOutColor(this);" align="absmiddle"></td>
				</tr>
				

				<tr>
					<td bgcolor="#FFEEF9" class="TDCont">��뿩��</td>
					<td bgcolor="#FFFFFF">
						<input type="radio" name="sUSEYN" value="Y" class="none" <% If sUSEYN = "Y" Or sUSEYN = "" Then response.write "checked" End If %>> �ź���
						<input type="radio" name="sUSEYN" value="N" class="none" <% If sUSEYN = "N" Then response.write "checked" End If %>> �źξ���
					</td>
				</tr>				


</form>

			</table>
			<table border="0" cellspacing="0" width="100%" align="center">
				<tr height="30">
					<td align="right">
						<img src="/Images/Btn/BtnSubmit.gif" style="cursor:hand;" align="absmiddle" onClick="javascript:fn_inup(document.inUpFrm);">
						<img src="/Images/Btn/BtnDel.gif" style="cursor:hand;" align="absmiddle" onClick="javascript:fn_del();">
					</td>
				</tr>
			</table>			



<script>
function fn_inup(inUpFrm) {

	if(!FieldChk(inUpFrm.DNIS,"������ȣ")) return;
	if(!FieldChk(inUpFrm.TelNo,"���Űźι�ȣ")) return;

	if(confirm("�����Ͻðڽ��ϱ�?"))
		inUpFrm.submit();
	else
		return;
}
function fn_del() {
	if(confirm("�����Ͻðڽ��ϱ�?"))
		location.href = "Callreject_InsUpDel.asp?guboon=DEL&idx=<%=idx%>";
	else
		return;
}


</script>