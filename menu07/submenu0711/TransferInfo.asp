<!-- #include virtual="/Include/Top.asp" -->
<%

	'1. �Ķ���� ������
	curPage = request("curPage")
	SS_Login_Grade = SESSION("SS_Login_Grade")
	
	'2. ���������� ����
	pageSize = 100
	pageSector = 10
	if curPage = "" then curPage = 1 end If
	where1= "whereCD1=" & whereCD1
	
	sql_tb = "TB_TransferInfo"
	sql_index = ""
	sql_field = "*"
	sql_orderby = "DNIS ASC, StartTime"
	sql_where = " 1=1 "

	if SS_Login_Grade = "B" then
		sql_where = sql_where & "	AND		( dnis in ( select extno from tb_userinfo where GRADE = '"&SS_Login_Grade&"')"
		sql_where = sql_where & "	OR		dnis = '9192')"
	elseif SS_Login_Grade <> "A" then
'-----------------------------------------------------------------------------------------------
		sql_where = sql_where & "	AND		dnis in ( select extno from tb_userinfo where GRADE = '"&SS_Login_Grade&"')"
'-----------------------------------------------------------------------------------------------
	end if
	
	'3. ���� ����
	sql = db_getSqlWithPage(sql_tb, sql_index, sql_field, sql_where, sql_orderby, pageSize, curPage)
	set rs = db.execute(sql)
	
	
	'4. Paging HTML �ۼ�
	totalCount = db_getCount(db, sql_tb, sql_where)
	startRow = totalCount - pageSize * (curPage - 1)
	pageHtml = getPageHtml(pageSector, pageSize, totalCount, curPage, currentURL & "?" & where1)

%>

<script language="javascript">

function fn_del(arg0,arg1)
{
	var df = document.frmBody;
	var flag = confirm("�ش� ����Ÿ�� ���� �Ͻðڽ��ϱ�?");
	if(flag == true)
		{
		df.action="Transferinfo_InsUpDel.asp?seqno="+arg0+"&guboon="+arg1;
		df.submit();
		}
}

function fn_update(arg0,arg1)
{
	var df = document.frmBody;
	
	df.action="Transferinfo_Detail.asp?seqno="+arg0+"&guboon="+arg1;
	df.submit();
}

function fn_insert()
{
	var df = document.frmBody;
	df.action="Transferinfo_Detail.asp?guboon=Ins";
	df.submit();
}

</script>



<table width="1000" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr><td height="10"></td></tr>
    <tr>
    	<td>
<!-- ���α׷� ����Ʈ START -->
<form name="frmBody" method="post">
<input type=hidden name="curPage" value="<%=curPage%>">

<table cellpadding="0" cellspacing="0" width="1000">
	<tr>
		<td height="30" class="TDR10px" width="100%">
			<img src="/Images/Btn/BtnAdd.gif" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_insert();">
		</td>
	</tr>
</table>

<table border="0" cellspacing="1" cellpadding="0" width="100%" bgcolor="#CCCCCC">
	<tr height="25" bgcolor="#F3F3F3" align="center">
		<td width="7%"><b>������ȣ</b></td>
		<td width="10%"><b>���Ź�ȣ</b></td>
		<td width="7%"><b>���Žð�</b></td>
		<td width="10%"><b>�����</b></td>
		<td width="7%"><b>��뿩��</b></td>
		<td width="7%"><b>������</b></td>
		<td width="7%"><b>ȭ����</b></td>
		<td width="7%"><b>������</b></td>
		<td width="7%"><b>�����</b></td>
		<td width="7%"><b>�ݿ���</b></td>
		<td width="7%"><b>�����</b></td>
		<td width="7%"><b>�Ͽ���</b></td>
		<td width="7%"><b>����</b></td>
		<td width="7%"><b></b></td>
	</tr>
	<% 
		if rs.EOF and rs.BOF then 
	%>
	<tr><td height="30" colspan="13" bgcolor="#FFFFFF"><p align="center">�˻��� �ڷᰡ �����ϴ�.</p></td></tr>
	<%
		else	
			do until rs.EOF
	%>
		<tr bgcolor="#FFFFFF">
			<td class="TDCont" align="center"><font color="#FF0000">[<%=rs("DNIS")%>]</font></td>
			<td class="TDCont" align="center"><%=rs("Transferno")%></td>
			<td class="TDCont" align="center"><%=rs("starttime")%>~<%=rs("endtime")%></td>
			<td class="TDCont" align="center"><%=db_Getusername(rs("UserId"))%></td>
			<td align="center"><input type="checkbox" class="none"<% If Rs("useyn") = "1" Then Response.Write("checked") End If %> disabled></td>
			<td align="center"><input type="checkbox" class="none"<% If Rs("Mon") = "1" Then Response.Write("checked") End If %> disabled></td>
			<td align="center"><input type="checkbox" class="none"<% If Rs("tue") = "1" Then Response.Write("checked") End If %> disabled></td>
			<td align="center"><input type="checkbox" class="none"<% If Rs("wed") = "1" Then Response.Write("checked") End If %> disabled></td>
			<td align="center"><input type="checkbox" class="none"<% If Rs("thu") = "1" Then Response.Write("checked") End If %> disabled></td>
			<td align="center"><input type="checkbox" class="none"<% If Rs("fri") = "1" Then Response.Write("checked") End If %> disabled></td>
			<td align="center"><input type="checkbox" class="none"<% If Rs("sta") = "1" Then Response.Write("checked") End If %> disabled></td>
			<td align="center"><input type="checkbox" class="none"<% If Rs("sun") = "1" Then Response.Write("checked") End If %> disabled></td>
			<td align="center"><input type="checkbox" class="none"<% If Rs("holiday") = "1" Then Response.Write("checked") End If %> disabled></td>
			<td align="center">
				<img src="/Images/Btn/BtnIconModify.gif" title="����" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_update('<%=rs("SEQNO")%>','UP');">
				<!--<img src="/Images/Btn/BtnIconDel.gif" title="����" style="cursor:hand;" align="absmiddle"onclick="javascript:fn_del('<%=rs("SEQNO")%>','DEL');">-->
			</td>
		</tr>
	<%
				startRow = startRow - 1
				rs.MoveNext 
			Loop
			
			rs.close 
			set rs = nothing
		end if
	%>  
</table>
<table cellpadding="0" cellspacing="0" width="100%">
	<tr><td height="2" bgcolor="#f2f2f2"></td></tr>
	<tr><td height="5"></td></tr>
	<tr><td bgcolor="#D6D6D6" height="1"></td></tr>
	<tr><td bgcolor="#F7F7F7" class="TDL10px" height="25"><%=pageHtml%></td></tr>
	<tr><td bgcolor="#D6D6D6" height="1"></td></tr>
	<tr>
		<td height="30" class="TDR10px">
			<img src="/Images/Btn/BtnAdd.gif" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_insert();">
		</td>
	</tr>
</table>
</form>
<!-- ���α׷� ����Ʈ END -->
    	</td>
    </tr>
</table>  
  

<!-- #include virtual="/Include/Bottom.asp" -->