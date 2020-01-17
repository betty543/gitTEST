<!-- #include virtual="/Include/Top.asp" -->
<%
	Response.Buffer = True
	Response.Expires = -1
	
	sSec_group = Request("sSec_group")
	sSec_name = Request("sSec_name")
	if sSec_group = "" then
	sSec_group = Request("txtSec_group")
	sSec_name = Request("txtSec_name")
	end if
	sProg_code = Request("txtProg_code")
	sProg_name = Request("txtProg_name")
			
	'// 보안그룹
	SQL ="select CODE,CODENAME from TB_CODE where	CODEGROUP = 'Z04' and useyn = 'Y'  order by CODE"
	set Rs = db.execute(SQL)
	                                
	If NOT(Rs.Eof Or Rs.Bof) Then
		Do Until Rs.Eof
			lv_cnt = lv_cnt + 1
			inHtml = inHtml & "<tr bgcolor='#FFFFFF'>"
			inHtml = inHtml & "<td align='center'>"& Rs("CODE") &"</td>"
			inHtml = inHtml & "<td class='TDCont' id='tmpTd_" & lv_cnt & "'><a href=""javascript:click_code('"&Trim(Rs("CODE"))&"','" & lv_cnt &"')"">"&Rs("CODENAME")&"</a></td>"	
			inHtml = inHtml & "</tr>"
			Rs.movenext
		Loop
		Rs.Close
		Set Rs = Nothing
	End if
%>

<script language="JavaScript">

function click_code(arg1,arg2){
	for(var i=1; i<=parseInt("<%=lv_cnt%>"); i++){
		eval("document.all.tmpTd_" + i).style.backgroundColor = (i == parseInt(arg2)) ? "#ff9999" : "#ffffff"
	}

	var df = document.frmBody;
	ifr_List.document.frmBody.txtSec_group.value = arg1;
	
	fn_submit1(arg1);	// 프로그램리스트 
}

function fn_submit1(arg,arg1)
{
	var df = document.frmBody;
	df.action="Group_Program_Available.asp?sSec_group="+arg;
	df.target="ifr_Available";
	df.submit();
}

function fn_submit2(arg,arg1)
{
	var df = document.frmBody;
	df.action="Group_Program_Detail.asp?sSec_group="+arg+"&sSec_name="+arg1;
	df.target="ifr_Detail";
	df.submit();
}
</script>


<form name="frmBody" method="post" style="margin:0">
<table width="940" border="0" cellpadding="0" cellspacing="0" align="center">
	<tr valign="top">
		<td width="300">
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
        		<tr><td height="22" colspan="2" class="FBlk">◈ <b>전체 프로그램</b></td></tr>
        	</table>
        	<table width="300" border="0" cellspacing="1" cellpadding="1" bgcolor="#CCCCCC">
        		<tr bgcolor="#EEF6FF" align="center">
        			<td class="TDCont">프로그램명</td>
        			<td nowrap width="66" class="TDCont">관리</td>
        		</tr>
        	</table>
        	<table width="100%" border="0" cellspacing="0" cellpadding="0">
        		<tr><td height="1"></td></tr>
        		<tr><td><iframe src="group_program_list.asp" frameborder=0 marginheight=0 marginwidth=0 width="100%" height="200" scrolling="auto" name="ifr_List" id ="ifr_List"></iframe></td></tr>
        	</table>			
		</td>
		<td width="5"></td>
		<td width="300">
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
        		<tr><td height="22" colspan="2" class="FBlk">◈ <b>운용업무</b></td></tr>
        	</table>
        	<table width="100%" border="0" cellspacing="1" cellpadding="1" bgcolor="#CCCCCC">
        		<tr bgcolor="#FDE6F3" align="center">
        			<td nowrap width="60" class="TDCont">구분</td>
        			<td class="TDCont">구분명</td>
        		</tr>
        		<%=inHtml%>
        	</table>		
		</td>
		<td width="5"></td>
		<td width="300">
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
        		<tr><td height="22" colspan="2" class="FBlk">◈ <b>사용가능한 프로그램</b></td></tr>
        	</table>
        	<table width="300" border="0" cellspacing="1" cellpadding="1" bgcolor="#CCCCCC">
        		<tr bgcolor="#EEF6FF" align="center">
        			<td class="TDCont">프로그램명</td>
        			<td nowrap width="46" class="TDCont">관리</td>
        		</tr>
        	</table>
        	<table width="100%" border="0" cellspacing="0" cellpadding="0">
        		<tr><td height="1"></td></tr>
        		<tr><td><iframe src="group_program_available.asp" frameborder=0 marginheight=0 marginwidth=0 width="100%" height="200" scrolling="auto" name="ifr_Available" id="ifr_Available"></iframe></td></tr>
        	</table>
		</td>
	</tr>
</table>
</form>

<!-- #include virtual="/Include/Bottom.asp" -->