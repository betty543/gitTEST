<!-- #include virtual="/include/top_frame.asp" -->
<%

	'1. 파라미터 얻어오기
	curPage = request("curPage")
	w_PARENT_ID = request("w_PARENT_ID")

	If w_PARENT_ID = "" Then
		w_PARENT_ID = "1"
	End IF
	
	'2. 쿼리조건절 셋팅
	pageSize = 100
	pageSector = 10
	if curPage = "" then curPage = 1 end If
	where1= "w_PARENT_ID=" & w_PARENT_ID
	
	
	Dim objCmd
	Set objCmd = Server.CreateObject("ADODB.Command")

	with objCmd

		.ActiveConnection = Casamiadb
		.CommandText = "USP_PROGRAM_SEL"
		.CommandType = adCmdStoredProc
		'.Parameters.Append.CreateParameter("@V_COMPANY_ID",adVarChar,adParamInput,50,"Centerlink")
		'.Parameters.Append.CreateParameter("@V_PARENT_ID",adInteger,adParamInput,0)
		'.Parameters.Append.CreateParameter("@V_SORT_TYPE",adVarChar,adParamInput,1000,"")

		.parameters.append(.CreateParameter("@V_COMPANY_ID",advarchar,adParamInput,50,COMPANY_ID))
		.parameters.append(.CreateParameter("@V_PARENT_ID",adInteger,adParamInput,,0))
		.parameters.append(.CreateParameter("@V_SORT_TYPE",advarchar,adParamInput,1000,""))
		Set PARENT_RS = .Execute
		.parameters.delete 2
		.parameters.delete 1
		.parameters.delete 0

	End with


	with objCmd

		.ActiveConnection = Casamiadb
		.CommandText = "USP_PROGRAM_SEL"
		.CommandType = adCmdStoredProc

		.parameters.append(.CreateParameter("@V_COMPANY_ID",advarchar,adParamInput,50,COMPANY_ID))
		.parameters.append(.CreateParameter("@V_PARENT_ID",adInteger,adParamInput,,w_PARENT_ID))
		.parameters.append(.CreateParameter("@V_SORT_TYPE",advarchar,adParamInput,1000,""))
		Set PROGRAM_RS = .Execute

	End with

	Set objCmd = Nothing

%>

<script language="javascript">



function fn_Search()
{
	var df = document.frmBody;
	
	df.action="Program_List.asp";
	df.submit();

}


function fn_insert()
{
	parent.DetailFrame.location.href='Program_Detail.asp?guboon=INS&w_PARENT_ID=<%=w_PARENT_ID%>';
}


</script>



<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr><td height="10"></td></tr>
    <tr>
    	<td>
<!-- 프로그램 리스트 START -->
<form name="frmBody" method="post" >
<input type=hidden name="curPage" value="<%=curPage%>">

<table cellpadding="0" cellspacing="0" width="100%">
	<tr>
		<td height="30" bgcolor="#FFFFFF" class="TDCont">그룹선택: &nbsp;
			<select name="w_PARENT_ID" size="1" class="#FFFFFF" onchange="javascript:fn_Search();">
			<%
				Do Until PARENT_RS.EOF
						CODE = PARENT_RS("PROGRAM_ID")
						CODENAME = PARENT_RS("PROGRAM_NM")
			%>
						<%=printSelect("" &CODENAME& "","" &CODE& "","" &w_PARENT_ID& "")%>
			<%
						PARENT_RS.MOVENEXT
				Loop
				PARENT_RS.close 
				set PARENT_RS = nothing				
			%>
			</select>					
		</td>
		<td height="30" class="TDR10px" width="50%" align='right'>
			<img src="/Images/Btn/BtnAdd.gif" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_insert();">
		</td>
	</tr>
</table>

<DIV style="OVERFLOW-Y:auto; OVERFLOW-X:no; MARGIN: 0px 0px 0px 0px; WIDTH:100%; HEIGHT:600;">
<table border="0" cellspacing="1" cellpadding="0" width="100%" bgcolor="#CCCCCC">
	<tr height="25" bgcolor="#F3F3F3" align="center">
		<td width="3%"><b>No</b></td>
		<td width="30%"><b>프로그램명</b></td>
		<td width="3%"><b>순서</b></td>
		<td width="30%"><b>URL</b></td>
		<td width="3%"><b>추가</b></td>
		<td width="3%"><b>수정</b></td>
		<td width="3%"><b>삭제</b></td>
		<td width="3%"><b>출력</b></td>
		<td width="3%"><b>사용</b></td>
	</tr>
	<% 
		if PROGRAM_RS.EOF and PROGRAM_RS.BOF then 
	%>
	<tr><td height="30" colspan="10" bgcolor="#FFFFFF"><p align="center">검색된 자료가 없습니다.</p></td></tr>
	<%
		else
	
			do until PROGRAM_RS.EOF
	%>
		<tr height="20" bgcolor="#FFFFFF" onClick="parent.DetailFrame.location.href='Program_Detail.asp?guboon=UP&w_PARENT_ID=<%=w_PARENT_ID%>&PROGRAM_ID=<%=PROGRAM_RS("PROGRAM_ID")%>';" onmouseover="setSelectColor(this);" onmouseout="setOutColor(this);" style="cursor:hand">
			<td class="TDCont" align="right" ><%=PROGRAM_RS("No")%></td>
			<td class="TDCont" align="left"><font color="#FF0000">[<%=PROGRAM_RS("PROGRAM_ID")%>]</font> <%=PROGRAM_RS("PROGRAM_NM")%></td>
			<td class="TDCont" align="right"><%=PROGRAM_RS("PROGRAM_IDX")%></td>
			<td class="TDCont" align="left"><%=PROGRAM_RS("NAVIGATE_URL")%></td>
			<td align="center"><input type="checkbox" class="none"<% If PROGRAM_RS("PROGRAM_INSERT") = "Y" Then Response.Write("checked") End If %> disabled></td>
			<td align="center"><input type="checkbox" class="none"<% If PROGRAM_RS("PROGRAM_UPDATE") = "Y" Then Response.Write("checked") End If %> disabled></td>
			<td align="center"><input type="checkbox" class="none"<% If PROGRAM_RS("PROGRAM_DELETE") = "Y" Then Response.Write("checked") End If %> disabled></td>
			<td align="center"><input type="checkbox" class="none"<% If PROGRAM_RS("PROGRAM_PRINT") = "Y" Then Response.Write("checked") End If %> disabled></td>
			<td align="center"><input type="checkbox" class="none"<% If PROGRAM_RS("USE_YN") = "Y" Then Response.Write("checked") End If %> disabled></td>

		</tr>
	<%
				PROGRAM_RS.MoveNext 
			Loop
			
			PROGRAM_RS.close 
			set PROGRAM_RS = nothing
		end if
	%>  
</table>
</div>
</form>
<!-- 프로그램 리스트 END -->
    	</td>
    </tr>
</table>  
  

