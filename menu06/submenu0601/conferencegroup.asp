
<!-- #include virtual="/Include/Top.asp" -->

<%


%>

<table width="940" border="0" cellpadding="0" cellspacing="0" align="center">
	<tr valign="top" height="450">

		<td width="340">

			<table width="100%" border="0" cellspacing="0" cellpadding="0">
        		<tr><td height="22" colspan="2" class="FBlk">�� <b>�׷�</b></td><td colspan="5" align="right" height=28><img src="/Images/Btn/BtnAdd.gif" title="�׷��߰�" style="cursor:hand;" align="absmiddle" onClick="ShowPOPLayer('conferencegroup_detail.asp','500','210');"></td></tr>
        	</table>
			<DIV style="OVERFLOW-Y:auto; OVERFLOW-X:auto; MARGIN: 0px 0px 0px 0px; 500; HEIGHT:420;">
        	<table width="100%" border="0" cellspacing="1" cellpadding="1" bgcolor="#CCCCCC">
        		<tr height="20" bgcolor="#EEF6FF" align="center">
        			<td>NO</td>
        			<td>����</td>
        			<td>�׷��</td>
        			<td>��뿩��</td>
        			<td align='center'>����</td>
        		</tr>
        		<tr><td colspan="5" height="1" bgcolor="#FFFFFF"></td></tr>
<%

				SQL = "	select * from TB_smsgroup where groupgb = '3' order by idx"
				i = 0
				SET Rs = db.execute(SQL)
				do until Rs.eof
						
						i = i + 1
						if rs("groupgb") = "3" then
							groupgb_NM = "���ڰ���"
						else
							groupgb_NM = "����"
						end if
						groupname=rs("groupname")
						if rs("useyn") = "Y" then
							useyn_NM = "���"
						else
							useyn_NM = "�̻��"
						end if
						idx = rs("idx")
					%>

				<tr onmouseover="this.style.background='#FFFCE7'" bgcolor="#FFFFFF" onmouseout="this.style.background='#FFFFFF'">
					<td align="center"><%=i%></td>
					<td align="center"><%=groupgb_NM%></td>
					<td align="center"><%=groupname%></td>
					<td align="center"><%=useyn_NM%></td>
					<td align="center" width="60">
							<img src="/Images/Comm/IconHome.gif" title="��Ϻ���" style="cursor:hand;" align="absmiddle" onclick="document.all.ogroupidx.value=<%=idx%>; ListFrame.location.href='conferencegroup_addresslist.asp?conferencegroupid=<%=idx%>&conferencegroupname=<%=groupgb_NM%>-<%=groupname%>';">
							<img src="/Images/Btn/BtnIconModify.gif" title="�׷����" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_update('<%=idx%>','UP');">
							<img src="/Images/Btn/BtnIconDel.gif" title="�׷����" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_del('<%=idx%>','D');">
					</td>
				</tr>
					<%
					Rs.movenext
				loop


%>
				<!--<tr onmouseover="this.style.background='#FFFCE7'" bgcolor="#FFFFFF" onmouseout="this.style.background='#FFFFFF'">
					<td align="center">1</td>
					<td align="center">����</td>
					<td align="center">ģ��</td>
					<td align="center">���</td>
					<td align="center" width="60">
							<img src="/Images/Comm/IconHome.gif" title="��Ϻ���" style="cursor:hand;" align="absmiddle" onclick="ListFrame.location.href='conferencegroup_addresslist.asp?conferencegroupid=<%=conferencegroupid%>&conferencegroupname=����-ģ��';">
							<img src="/Images/Btn/BtnIconModify.gif" title="�׷����" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_update('1111','UP');">
							<img src="/Images/Btn/BtnIconDel.gif" title="�׷����" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_update('1111','DEL');">
					</td>
				</tr>
				<tr onmouseover="this.style.background='#FFFCE7'" bgcolor="#FFFFFF" onmouseout="this.style.background='#FFFFFF'">
					<td align="center">2</td>
					<td align="center">����</td>
					<td align="center">��޹߷�</td>
					<td align="center">���</td>
					<td align="center" width="60">
							<img src="/Images/Comm/IconHome.gif" title="��Ϻ���" style="cursor:hand;" align="absmiddle" onclick="ListFrame.location.href='conferencegroup_addresslist.asp?conferencegroupid=<%=conferencegroupid%>&conferencegroupname=����-��޹߷�';">
							<img src="/Images/Btn/BtnIconModify.gif" title="�׷����" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_update('1111','UP');">
							<img src="/Images/Btn/BtnIconDel.gif" title="�׷����" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_update('1111','DEL');">
					</td>
				</tr>
				<tr onmouseover="this.style.background='#FFFCE7'" bgcolor="#FFFFFF" onmouseout="this.style.background='#FFFFFF'">
					<td align="center">3</td>
					<td align="center">����</td>
					<td align="center">��޹߷�</td>
					<td align="center">���</td>

					<td align="center">
							<img src="/Images/Btn/BtnIconModify.gif" title="�׷����" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_update('1111','UP');">
							<img src="/Images/Btn/BtnIconDel.gif" title="�׷����" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_update('1111','DEL');">
					</td>
				</tr>-->			
				
			</table>
			</DIV>
		</td>
		<td width="10"></td>
		<td width="590">
			<iframe src="conferencegroup_addresslist.asp?conferencegroupid=<%=idx%>&conferencegroupname=<%=groupgb_NM%>-<%=groupname%>" name="ListFrame" width="100%" height="100%" frameborder=0 marginheight=0 marginwidth=0 scrolling="no"></iframe>
		</td>

	</tr>
	<tr>
	<td colspan=3>
		<table width="940" border="0" cellspacing="0" cellpadding="0">
			<tr><td height="22" colspan="2" class="FBlk"></b></td></tr>
		</table>			
		<table width="940"  border="0" cellspacing="1" cellpadding="1" bgcolor="#CCCCCC">
			<tr bgcolor="#EEF6FF" height="40" align=center><td height="22" colspan="2" class="FBlk" align="center"> ������ �����&nbsp;&nbsp;
						<select name="movegroup" size="1" class="ComboFFFCE7" >
<%					
							SQL = "	select * from tb_smsgroup where useyn = 'Y' and groupgb = '3'" '�ι�����
							SET Rs = DB.execute(SQL)
							do until Rs.eof
									CODE = Rs("idx")
									if Rs("groupgb")="3" then
										CODENAME = ""&Rs("groupname")
									else
										CODENAME = "����-"&Rs("groupname")
									end if
								%>

									<%=printSelect("" &CODENAME& "","" &CODE& "","" &movegroup& "")%>
								<%
								Rs.movenext
							loop
%>
						</select> �׷쿡 
						&nbsp;&nbsp;&nbsp;<img src="/Images/Btn/BtnPlus_Add.GIF" title="�߰��ϱ�" style="cursor:hand;" align="absmiddle" onclick="javascript:selectvalue();javascript:fn_insert();"> &nbsp;&nbsp;�Ǵ� &nbsp;&nbsp;<img src="/Images/Btn/BtnPlus_Mov.GIF" title="�̵��ϱ�" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_move();"></b></td></tr>
		</table>
	</td>
	</tr>
	<form name="insFrm" method="post" action="conferencegroup_move.asp">
<input value="" name="JOBGB" type="hidden" size="30">
<input value="" name="selectvalue" type="hidden" size="30">
<input value="<%=idx%>" name="ogroupidx" type="hidden" size="30">
<input value="" name="ngroupidx" type="hidden" size="30">
	</form>
</table>

<!-- #include virtual="/Include/PopLayer.asp" -->
<!-- #include virtual="/Include/Bottom.asp" -->

<script>
<!--
	function fn_update(arg0, arg1){
		//sms�׷� ����
		ShowPOPLayer('conferencegroup_detail.asp?idx='+arg0,'500','210');
	}
	function fn_del(arg0, arg1){
		//sms�׷� ����
		if (confirm("������ �׷��� �����Ͻðڽ��ϱ�?"))
			ShowPOPLayer('conferencegroup_detail.asp?idx='+arg0+'&JOBGB=D','500','210');
	}
	function fn_insert(){
		//sms�׷� ����

	if ( document.insFrm.selectvalue.value == "undefined" )
		document.insFrm.selectvalue.value == "";

		if ( document.insFrm.selectvalue.value.length==0 )
			alert('������ �ڷᰡ �����ϴ�');

		if (confirm("������ �ڷḦ �߰��Ͻðڽ��ϱ�?"))
		{
			document.insFrm.JOBGB.value="INS";
			document.all.ngroupidx.value=document.all.movegroup.value;
			insFrm.submit();
		}

	}
	function fn_move(){
		//sms�׷� ����

		if ( document.insFrm.selectvalue.value == "undefined" )
			document.insFrm.selectvalue.value == "";

		if ( document.insFrm.selectvalue.value.length==0 )
			alert('������ �ڷᰡ �����ϴ�');

		if (confirm("������ �ڷḦ �̵��Ͻðڽ��ϱ�?"))
		{
			document.insFrm.JOBGB.value="MOV";
			document.all.ngroupidx.value=document.all.movegroup.value;
			insFrm.submit();
		}
	}


  function selectvalue() {
    var i, chked=0;
	document.insFrm.selectvalue.value = ListFrame.document.ListForm.Chk.length;

    for(i=0;i<ListFrame.document.ListForm.Chk.length;i++) 
	{
		if(ListFrame.document.ListForm.Chk[i].type=='checkbox') 
		{ 
			if(ListFrame.document.ListForm.Chk[i].checked) 
			{ 				
				if ( document.insFrm.selectvalue.value == "")
				{
					document.insFrm.selectvalue.value = ListFrame.document.ListForm.Chk[i].value;	
				}
				else 
				{
					document.insFrm.selectvalue.value = document.insFrm.selectvalue.value + ':' + ListFrame.document.ListForm.Chk[i].value;	
				} 
			}

		}

   }
	if ( document.insFrm.selectvalue.value == "undefined" && ListFrame.document.ListForm.Chk.checked )
		document.insFrm.selectvalue.value = "1:" + ListFrame.document.ListForm.Chk.value;
	

		
}

//-->
</script>