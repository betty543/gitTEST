<!-- #include virtual="/include/top_frame.asp" -->

<body leftmargin="0" marginwidth="0" topmargin="0" marginheight="0" onLoad="ifHeight('ifr_List');">

<div name="ifr" id="ifr">

<%
	sSec_group = Trim(Request("Sec_group"))
%>
<script language="JavaScript">
function click_progcode(pc){
	var df = document.frmBody;
	
	if(df.txtSec_group.value == ""){
		alert("보안그룹이 선택되지 않았습니다.\n보안그룹을 먼저 선택해주세요.")
		return false;
	}
	document.location.href="Group_Program_InsUpDel.asp?guboon=INS&Sec_group="+df.txtSec_group.value+"&Prog_Code="+pc
}

</script>

<form name="frmBody" method="post" action="Group_Program_InsUpDel.asp" style="margin:0">
<input type="hidden" name="txtSec_group" value="<%=sSec_group%>">
<table width=100% border="0" cellspacing="1" cellpadding="1" bgcolor="#CCCCCC">
	<%
		SQL1 = "select * from TB_PROGRAM order by PROG_CODE"
		set Rs1 = db.execute(SQL1)
		
		If Rs1.Eof Or Rs1.bof Then
	%>
	<tr><td height="50" align="center" bgcolor="#FFFFFF" style="color:#0000FF">설정된 보안 프로그램이 없습니다.</td></tr>
	<%
		else
			num = 1
			Do until Rs1.Eof
	%>
	<tr bgcolor="#FFFFFF">
		<td width="50" align="center" class="TDCont"><%=num%><input type="hidden" name="txtProg_Code" value="<%=rs1("PROG_CODE")%>"></TD>
		<td class="TDCont"><font color="#FF0000">[<%=rs1("PROG_CODE")%>]</font> <%=Rs1("PROG_NAME")%></a></td>
		<td width="70" align="center"><img src="/Images/Btn/BtnAdd.gif" style="cursor:hand;" align="absmiddle" onclick="javascript:click_progcode('<%=rs1("PROG_CODE")%>');"></TD>
	</tr>
	
	<%
				Rs1.movenext
				num = num + 1
			loop
			Rs1.Close
			Set Rs1 = Nothing
		End If
	%>
</table>
</form>

</div>