<!-- #include virtual="/Include/Top.asp" -->
<%
	'---------------------------------------------
	sql_tb = "TB_CALLBACK"
	sql_where = "PROCESSGB IS NULL"  '접수중인 상태로 있는 자료
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
			if (!confirm("변경사항을 적용하시겠습니까?")) return;
			ifr_List.ListForm.action = "CallBackState_detail_InsUp.asp";
			ifr_List.ListForm.submit();
		} else { alert("상담원을 한 명 이상 선택해야 합니다."); }

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
        		<tr><td height="22" colspan="2" class="FBlk">◈ <b>상담관 리스트</b></td></tr>
        	</table>

        	<table width="1000" border="0" cellspacing="1" cellpadding="1" bgcolor="#CCCCCC">
        		<tr height="20" bgcolor="#EFEFEF" align="center">
        			<td width="40">NO</td>
        			<td width="100">아이디</td>
        			<td width="100">성명</td>
        			<td width="100">등급</td>
        			<td width="50">콜백여부</td>
					<td width="375">콜백담당업무</td>
        		</tr>
        	</table>
        	<table cellpadding="0" cellspacing="0" width="1000">
				<tr>
					<td>
        	<!-- 상담원 리스트 -->
        	<iframe src="CallBackState_detail.asp" frameborder=0 marginheight=0 marginwidth=0 width="1000" height="400" scrolling="auto" name="ifr_List" id ="ifr_List"></iframe>
        	<!-- 상담원 리스트 -->
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