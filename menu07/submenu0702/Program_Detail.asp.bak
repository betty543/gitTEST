
<!-- #include virtual="/include/top_frame.asp" -->

<%
	guboon = Request("guboon")
	curPage = Request("curPage")
	w_PARENT_ID = Request("w_PARENT_ID")
	PROGRAM_ID = Request("PROGRAM_ID")


	Dim objCmd
	Set objCmd = Server.CreateObject("ADODB.Command")

	with objCmd

		.ActiveConnection = Casamiadb
		.CommandText = "USP_PROGRAM_SEL"
		.CommandType = adCmdStoredProc

		.parameters.append(.CreateParameter("@V_COMPANY_ID",advarchar,adParamInput,50,COMPANY_ID))
		.parameters.append(.CreateParameter("@V_PARENT_ID",adInteger,adParamInput,,0))
		.parameters.append(.CreateParameter("@V_SORT_TYPE",advarchar,adParamInput,1000,""))
		Set PARENT_ID_RS = .Execute

		.parameters.delete 2
		.parameters.delete 1
		.parameters.delete 0

	End with

	
	If guboon = "UP" Then

		SQL ="SELECT * FROM TB_PROGRAM WITH (NOLOCK)  WHERE PROGRAM_ID = '" & PROGRAM_ID & "'"
		Set rs = Casamiadb.execute(sql)
	
		if not rs.eof then 

			db_PROGRAM_ID = rs("PROGRAM_ID")
			db_PARENT_ID = rs("PARENT_ID")
			db_PROGRAM_IDX = rs("PROGRAM_IDX")
			db_PROGRAM_NM = rs("PROGRAM_NM")
			db_NAVIGATE_URL = rs("NAVIGATE_URL")
			db_PROGRAM_INSERT = rs("PROGRAM_INSERT")
			db_PROGRAM_UPDATE = rs("PROGRAM_UPDATE")
			db_PROGRAM_DELETE = rs("PROGRAM_DELETE")
			db_PROGRAM_PRINT = rs("PROGRAM_PRINT")
			db_USE_YN = rs("USE_YN")

		end if
	
		rs.close
		set rs = Nothing
		
	End If
	

%>
<script language="javascript">

function fn_inup(f)
{
	if(!FieldChk(f.PROGRAM_NM,"프로그램명")) return false;
	if(!FieldChk(f.PROGRAM_IDX,"순서")) return false;
	if(!FieldChk(f.NAVIGATE_URL,"URL")) return false;	
	f.submit();
}
</script>
<table width="580" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr><td height="10"></td></tr>
    <tr>
    	<td>
<!-- 프로그램 입력 폼 START -->
<form name="frmBody" method="post" action="Program_InsUpDel.asp">
<input type=hidden name="curPage" value="<%=curPage%>">
<input type=hidden name="guboon" value="<%=guboon%>">
<input type=hidden name="w_PARENT_ID" value="<%=w_PARENT_ID%>">

<table width="100%" border="0" cellspacing="0" cellpadding="0">
	<tr><td height="30" colspan="2" class="FBlk">◈ <b>프로그램 정보</b></td></tr>
</table>

<table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
	<tr>
		<td bgcolor="#F3F3F3" width="70" class="TDCont" align='center'><font color="black">상위메뉴</font></td>
		<td bgcolor="#FFFFFF" width="230" class="TDL5px">
			<select name="PARENT_ID" size="1" class="#FFFFFF" onchange="javascript:fn_Search();">
			<%
				Do Until PARENT_ID_RS.EOF
						CODE = PARENT_ID_RS("PROGRAM_ID")
						CODENAME = PARENT_ID_RS("PROGRAM_NM")
			%>
						<%=printSelect("" &CODENAME& "","" &CODE& "","" &db_PARENT_ID& "")%>
			<%
						PARENT_ID_RS.MOVENEXT
				Loop
				PARENT_ID_RS.close 
				set PARENT_ID_RS = nothing				
			%>
			</select>					
		</td>
	<tr>
		<td bgcolor="#F3F3F3" width="70" class="TDCont" align='center'><font color="black">코    드</font></td>
		<td bgcolor="#FFFFFF" width="230" class="TDL5px"><input name="PROGRAM_ID" type="text" value="<%=db_PROGRAM_ID%>" class="input"  size="25" onfocus="setFocusColor(this)" onblur="setOutColor(this)" maxlength="5" readonly></td>
	</tr>
	<tr>
		<td bgcolor="#F3F3F3" width="70" class="TDCont" align='center'><font color="black">프로그램명</font></td>
		<td bgcolor="#FFFFFF" width="221" class="TDL5px"><input name="PROGRAM_NM" type="text" value="<%=db_PROGRAM_NM%>" class="input"  size="25" onfocus="setFocusColor(this)" onblur="setOutColor(this)"></td>
	</tr>
	<tr>
		<td bgcolor="#F3F3F3" width="70" class="TDCont" align='center'><font color="black">순    서</font></td>
		<td bgcolor="#FFFFFF" width="221" class="TDL5px"><input name="PROGRAM_IDX" type="text" value="<%=db_PROGRAM_IDX%>" class="input"  size="25" onfocus="setFocusColor(this)" onblur="setOutColor(this)"></td>
	</tr>
	<tr height="25">
		<td bgcolor="#F3F3F3" width="70" class="TDCont" align='center'><font color="black">URL</font></td>
		<td bgcolor="#FFFFFF" width="517" colspan="3" class="TDL5px"><textarea name="NAVIGATE_URL" style="width:100%; height:70" wrap="soft" class="TextareaInput"><%=db_NAVIGATE_URL%></textarea></td>
	</tr>
	<tr height="25">
		<td bgcolor="#F3F3F3" width="70" class="TDCont" align='center'><font color="black">권    한</font></td>
		<td bgcolor="#FFFFFF" width="517" colspan="3" class="TDL5px">
			<input type="checkbox" name="PROGRAM_INSERT" value="Y" class="none" <% if db_PROGRAM_INSERT="Y" then Response.Write("checked") end if %>>등록 &nbsp;&nbsp;&nbsp;&nbsp;
			<input type="checkbox" name="PROGRAM_UPDATE" value="Y" class="none" <% if db_PROGRAM_UPDATE="Y" then Response.Write("checked") end if %>>수정 &nbsp;&nbsp;&nbsp;&nbsp;
			<input type="checkbox" name="PROGRAM_DELETE" value="Y" class="none" <% if db_PROGRAM_DELETE="Y" then Response.Write("checked") end if %>>삭제 &nbsp;&nbsp;&nbsp;&nbsp;
			<input type="checkbox" name="PROGRAM_PRINT" value="Y" class="none" <% if db_PROGRAM_PRINT="Y" then Response.Write("checked") end if %>>출력 &nbsp;&nbsp;&nbsp;&nbsp;
		</td>
	</tr>
	<tr height="25">
		<td bgcolor="#F3F3F3" width="70" class="TDCont" align='center'><font color="black">사용여부</font></td>
		<td bgcolor="#FFFFFF" width="517" colspan="3" class="TDL5px">
			<input type="checkbox" name="USE_YN" value="Y" class="none" <% if db_USE_YN="Y" then Response.Write("checked") end if %>>사용     
		</td>
	</tr>
</table>
</form>

<!-- 프로그램 입력 폼 END -->
		</td>
	</tr>
</table>
<table width="100%" border="1" cellspacing="0" cellpadding="0" align="center">
	<tr><td height="5" colspan='2'></td></tr>
	<tr>
		<td class="TDR10px" align='left'>
			<img src="/Images/Btn/BtnDel.gif" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_del(document.frmBody);">
		</td>
		<td class="TDR10px"  align='right'>
			<img src="/Images/Btn/BtnSubmit.gif" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_inup(document.frmBody);">
			<img src="/Images/Btn/BtnReset.gif" style="cursor:hand;" align="absmiddle" onclick="javascript:history.back();">
		</td>
	</tr>
</table>
