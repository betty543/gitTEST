<!-- #include virtual="/include/top_frame.asp" -->
<body leftmargin="0" marginwidth="0" topmargin="0" marginheight="0" onLoad="ifHeight('ifr_Available');">

<div name="ifr" id="ifr">

<%
	sSec_group = Request("sSec_group")
%>

<script language="JavaScript">
	function fn_del(arg){
		var answer = confirm("정말 삭제하시겠습니까?");
		if(answer == true){
			document.location.href="Group_Program_InsUpDel.asp?guboon=DEL&Sec_group=<%=sSec_group%>&Prog_Code="+arg
		}
	}
</script>


<table width="100%" border="0" cellspacing="1" cellpadding="1" bgcolor="#CCCCCC">
	<%
		If sSec_Group = "" then
	%>
	<tr><td height="50" colspan="3" align="center" bgcolor="#FFFFFF" style="color:#0000FF">선택된 보안 프로그램이 없습니다.</td></tr>
	<%
		else
			SQL2 = "select * from TB_SECGROUP"
			SQL2 = SQL2 & "	WHERE SEC_GROUP = '" &sSec_Group & "'"
			SQL2 = SQL2 & "order by PROG_CODE"
			set Rs2 = db.execute(SQL2)
			
			If Rs2.Eof Or Rs2.bof Then
	%>
	<tr><td height="50" colspan="3" align="center" bgcolor="#FFFFFF" style="color:#0000FF">사용가능한 프로그램이 없습니다.</td></tr>
	<%
			else
				i=0
				Do Until Rs2.EOF
					i=i+1
					ProgName = db_getProgName(Rs2("PROG_CODE"))
	%>
	<tr bgcolor="#FFFFFF">
		<td width="50" align="center" class="TDCont"><%=i%></td>
		<td class="TDCont"><font color="#FF0000">[<%=Rs2("PROG_CODE")%>]</font> <%= ProgName%></td>
		<td nowrap width="50" align="center"><img src="/Images/Btn/BtnIconDel.gif" title="삭제" style="cursor:hand;" onclick="fn_del('<%=Rs2("PROG_CODE")%>');"></td>
	</tr>
	<%
					Rs2.MoveNext
				Loop
				Rs2.Close
				Set Rs2 = Nothing
			end if
		end if
	%>
</table>

</div>