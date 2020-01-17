<!-- #include virtual="/include/top_frame.asp" -->
<%
	'1. 파라미터 얻어오기
	curPage = request("curPage")
	sCodegroup = request("sCodegroup")
	sGroupname = request("sGroupname")
			

SS_LoginID = SESSION("SS_LoginID")
SS_Login_Secgroup = SESSION("SS_Login_Secgroup")


	'2. 쿼리조건절 셋팅
	pageSize = 100
	pageSector = 10
	if curPage = "" then curPage = 1 end If
	where1 = "sCodegroup=" & sCodegroup & "&sGroupname=" & sGroupname 
	where2 = "curPage=" & curPage & "&" & where1
	
	sql_tb = "TB_CODE"
	'sql_index = "index_desc(" & sql_tb & " IDX_TB_CALLHISTORY_JUBSEQ)"
	sql_field = "CODE, CODENAME, MEMO, USEYN, SYSYN"
	sql_orderby = "CODE"
	sql_where = " 1=1 "
	if sCodegroup <> "" then			sql_where = sql_where & " and CODEGROUP = '" & sCodegroup & "'" end If
	
	
	'3. 쿼리 실행
	sql = db_getSqlWithPage(sql_tb, sql_index, sql_field, sql_where, sql_orderby, pageSize, curPage)
	set rs = db.execute(sql)
	'LogWrite sql, "Code_List.asp", "/Setup/Code/"
	
	
	'4. Paging HTML 작성
	totalCount = db_getCount(db, sql_tb, sql_where)
	startRow = totalCount - pageSize * (curPage - 1)
	pageHtml = getPageHtml(pageSector, pageSize, totalCount, curPage, currentURL & "?" & where1)
%>

<table border="0" cellspacing="1" cellpadding="0" width="100%" bgcolor="#CCCCCC">
	<tr height="25" bgcolor="#F3F3F3" align="center">
		<td width="15%" class="TDCont"><b>코드</b></td>
		<td align="center"><b>코드명</b></td>
		<td width="48%"><b>메모</b></td>
		<td width="10%" align="center"><b>사용여부</b></td>
		<td width="5%" align="center"><b>시스템</b></td>
		<td width="6%" align="center"><b>관리</b></td>
	</tr>
	<% 
		if rs.EOF and rs.BOF then 
	%>
	<tr height="25"><td height="30" colspan="7" bgcolor="#FFFFFF"><p align="center">검색된 자료가 없습니다.</p></td></tr>
	<%
		end if
	
		do until rs.EOF
	%>
	
	<tr bgcolor="#FFFFFF" align="center">
		<td><%=rs("CODE")%></td>
		<td align="left"><%=rs("CODENAME")%></td>
		<td align="left"><%=rs("MEMO")%>&nbsp;</td>
		<td><input type="checkbox" class="none"<% If Rs("USEYN") = "Y" Then Response.Write("checked") End If %> disabled></td>
		<td><input type="checkbox" class="none"<% If Rs("SYSYN") = "Y" Then Response.Write("checked") End If %> disabled></td>
		<td>
			<img src="/Images/Btn/BtnIconModify.gif" title="수정" style="cursor:hand;" align="absmiddle" onClick="javascript:fn_update('<%=rs("CODE")%>');">
			<% if SS_Login_Secgroup <> "A" then%><img src="/Images/Btn/BtnIconDel.gif" title="삭제" style="cursor:hand;" align="absmiddle" onClick="javascript:fn_del('<%=rs("CODE")%>');"><%end if%>
		</td>
	</tr>
	
	
	<%
			startRow = startRow - 1
			rs.MoveNext 
		Loop
		
		rs.close 
		set rs = nothing
	%>  
</table>
<table cellpadding="0" cellspacing="0" width="100%">
	<tr><td height="2" bgcolor="#f2f2f2"></td></tr>
	<tr><td height="5"></td></tr>
	<tr><td bgcolor="#D6D6D6" height="1"></td></tr>
	<tr><td bgcolor="#F7F7F7" class="TDL10px" height="25"><%=pageHtml%></td></tr>
	<tr><td bgcolor="#D6D6D6" height="1"></td></tr>
	<tr>
		<td height="30" class="TDR10px"><img src="/Images/Btn/BtnAdd.gif" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_insert();"></td>
	</tr>
</table>

<script language="javascript">

function fn_insert()
{
	location.href = "code_detail.asp?guboon=INS&<%=where1%>";
}

function fn_update(code)
{
	location.href = "code_detail.asp?guboon=UP&sCode="+code+"&<%=where1%>";
}

function fn_del(code)
{
	if(confirm("해당 데이타를 삭제 하시겠습니까?")) {
		location.href = "code_InsUpDel.asp?guboon=DEL&sCode="+code+"&<%=where1%>";
	}
}

</script>