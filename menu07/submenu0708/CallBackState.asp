<!-- #include virtual="/Include/Top.asp" -->
<%
	'---------------------------------------------
	sql_tb = "TB_CALLBACK"
	sql_where = "PROCESSGB IS NULL"  '�������� ���·� �ִ� �ڷ�
	CallBack_cnt = db_getCount(db, sql_tb, sql_where)


	SQL = "SELECT DIVIDEKIND, INDATE, INCODE FROM TB_CONFIG_CALLBACK"
	SQL = SQL & "	WHERE USEYN = 'Y'"

	Set RS = db.execute(SQL)

	If RS.EOF Then
		DIVIDEKIND = "0"
		checked1 = "checked"
	Else
		DIVIDEKIND = RS("DIVIDEKIND")
		checked2 = "checked"
		INDATE = FORMATDATEH(rs("INDATE"))
		INCODE = RS("INCODE")
	End IF

%>
<script language="javascript">
<!--
	function fn_set(){
		var count = 0;
		if(ifr_List.ListForm.Chk.length > 0) {
			for(i = 0; i < ifr_List.ListForm.Chk.length; i++) {
				if(ifr_List.ListForm.Chk[i].checked) { count = count+1 }
			}
		} else {
			if(ifr_List.ListForm.Chk.checked) { count = 1 }
		}
		if(count > 0) {
			if (!confirm("��������� �����Ͻðڽ��ϱ�?")) return;
			ifr_List.ListForm.action = "CallBackState_detail_InsUp.asp";
			ifr_List.ListForm.submit();
		} else { alert("������ �� �� �̻� �����ؾ� �մϴ�."); }

		ifr_List.ListForm.submit();
	}

	function fn_reset(){
		ifr_List.ListForm.reset();
	}
-->
</script>
<table width="1000" border="0" cellpadding="0" cellspacing="0" align="center">
	<tr valign="top">
		<td width="1000">


        	<table width="100%" border="0" cellspacing="0" cellpadding="0">
        		<tr><td height="22" colspan="2" class="FBlk">�� <b>���� ����Ʈ</b></td></tr>
        	</table>

        	<table width="1000" border="0" cellspacing="1" cellpadding="1" bgcolor="#CCCCCC">
        		<tr height="20" bgcolor="#EFEFEF" align="center">
        			<td width="40">NO</td>
        			<td width="100">���̵�</td>
        			<td width="100">����</td>
        			<td width="100">���</td>
        			<td width="50">�ݹ鿩��</td>
					<td width="375">�ݹ������</td>
        		</tr>
        	</table>
        	<table cellpadding="0" cellspacing="0" width="1000">
				<tr>
					<td>
        	<!-- ���� ����Ʈ -->
        	<iframe src="CallBackState_detail.asp" frameborder=0 marginheight=0 marginwidth=0 width="1000" height="400" scrolling="auto" name="ifr_List" id ="ifr_List"></iframe>
        	<!-- ���� ����Ʈ -->
        			</td>
				</tr>
			</table>
        	<table border="0" cellspacing="0" width="1000" align="center">
				<tr height="30">
					<td align="right">
						<img src="/Images/Btn/BtnSubmit.gif" style="cursor:hand;" align="absmiddle" onClick="fn_set();">
						<img src="/Images/Btn/BtnReset.gif" style="cursor:hand;" align="absmiddle" onClick="fn_reset();">
					</td>
				</tr>
			</table>


		</td>
	</tr>
</table>
<!-- #include virtual="/Include/Bottom.asp" -->