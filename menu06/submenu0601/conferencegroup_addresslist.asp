
<!-- #include virtual="/Include/Top_Frame.asp" -->
<%
	smsgroupname=Request("conferencegroupname")
	smsgroupid=Request("conferencegroupid") 
	
	if ( smsgroupname <> "") then

%>
			<form name="ListForm" method="post" style="margin:0">
			<input value="<%=smsgroupid%>" name="smsgroupid" type="hidden" size="30">
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
        		<tr><td height="22" colspan="2" class="FBlk">�� <b> [<%=smsgroupname%>] �׷��� �ּҷ�</b></td><td height=28 colspan="1" height="1" align="right"><img src="/Images/Btn/BtnAdd.gif" title="�ּҷ��߰�" style="cursor:hand;" align="absmiddle" onClick="ShowPOPLayer('conferencegroup_addressdetail.asp?group_idx=<%=smsgroupid%>','500','250');"></td></tr>
        	</table>
			<DIV style="OVERFLOW-Y:auto; OVERFLOW-X:auto; MARGIN: 0px 0px 0px 0px; 500; HEIGHT:420;">
        	<table width="100%" border="0" cellspacing="1" cellpadding="1" bgcolor="#CCCCCC">
        		<tr height="20" bgcolor="#EEF6FF" align="center">
        			<td><input type="checkbox" name="ChkALL" class="None" onClick="select();"></td>
        			<td>NO</td>
        			<td>�Ҽ�</td>
        			<td>���</td>
        			<td>����</td>
        			<td>����ȭ</td>
        			<td>�޴�����ȣ</td>
        			<td width=40 align='center'>����</td>
        		</tr>
        		<tr><td colspan="8" height="1" bgcolor="#FFFFFF"></td></tr>
<%

	SQL = "SELECT	*	FROM TB_SMSADDR	WHERE	group_idx= " & smsgroupid & " order by idx"

				i = 0
				SET Rs = db.execute(SQL)


				do until Rs.eof	
				
					i = i + 1
					idx = rs("idx")
					sosok_name = rs("sosok_name")
					sclass = rs("class")
					sname = rs("name")
					cellphone = rs("cellphone")
					gunphone = rs("gunphone")
%>
				<tr onmouseover="this.style.background='#FFFCE7'" bgcolor="#FFFFFF" onmouseout="this.style.background='#FFFFFF'" >
					<td align="center" width=5%><input type="checkbox" name="Chk" value="<%=idx%>" class="None"></td>
					<td align="center" width=5%><%=i%></td>
					<td align="center" ><%=sosok_name%></td>
					<td align="center" width=10%><%=sclass%></td>
					<td align="center" width=20%><%=sname%></td>
					<td align="center" width=15%><%=gunphone%></td>
					<td align="center" width=20%><%=FormatHPNo(cellphone)%></td>
					<td align="center">
						<img src="/Images/Btn/BtnIconModify.gif" title="�ּҷ� ����" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_update('<%=idx%>','UP');">
						<img src="/Images/Btn/BtnIconDel.gif" title="�ּҷ� ����" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_del('<%=idx%>','DEL');">
					</td>
				</tr>
<%
					rs.movenext
				loop
				if ( i = 0 ) then
%>
        			<tr><td colspan="8" height="100" bgcolor="#FFFFFF" align="center">������ �׷쿡 ��ϵ� �ڷᰡ �����ϴ�.</td></tr>
<%
				end if
%>
        	</table>    
			</DIV>
			</form>

<% end if %>

<!-- #include virtual="/Include/PopLayer.asp" -->
<!-- #include virtual="/Include/Bottom.asp" -->

<script>
<!--
	function fn_update(arg0, arg1){
		//sms�׷� ����
		ShowPOPLayer('conferencegroup_addressdetail.asp?group_idx=<%=smsgroupid%>&idx='+arg0,'500','250');
	}
	function fn_del(arg0, arg1){
		//sms�׷� ����
		if (confirm("������ �ڷḦ �����Ͻðڽ��ϱ�?"))
			ShowPOPLayer('conferencegroup_addressdetail.asp?group_idx=<%=smsgroupid%>&idx='+arg0+'&JOBGB=D','500','250');
	}

  function select() {
	alert('ź��');

}
//-->
</script>