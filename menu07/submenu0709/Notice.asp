<!-- #include virtual="/Include/Top.asp" -->

<%
	'####### 1. �Ķ���� ������ �� �⺻���� ���� ###########################################################
	curPage = TRIM(request("curPage"))
	SMODE = TRIM(request("SMODE"))
	SWORD = TRIM(request("SWORD"))
	ACLASS = TRIM(request("ACLASS"))

	'Response.Write("SMODE="&SMODE&"<br>")
	'Response.Write("SWORD="&SWORD&"<br>")
	'Response.Write("ACLASS="&ACLASS&"<br>")

	pageSize = 25
	pageSector = 10
	IF curPage = "" THEN curPage = 1 END IF
	pageWHERE= "ACLASS=" &ACLASS& "&SMODE=" &SMODE& "&SWORD=" &SWORD

	'####### 2. ���������� ���� ##############################################################################
	SQL_Table = "TB_BOARD_NOTICE"
	SQL_Index = "IDX_TB_BOARD_NOTICE_SEQ"
	SQL_Field = "IDX, ACLASS, TITLE, FILENAME1, READCNT, INDATE, INCODE, FRONTYN"
	SQL_Orderby = "IDX DESC"
	SQL_Where = "1=1"
	IF NOT(ACLASS="") THEN SQL_Where=SQL_Where& " AND ACLASS='" &ACLASS& "'" END IF
	IF NOT(SMODE="") THEN SQL_Where=SQL_Where& " AND " &SMODE& " LIKE '%" &SWORD& "%'" END IF

	'####### 3. ���ڵ� ��� ������ #########################################################################
	SQL = db_getSqlWithPage(SQL_Table, SQL_Index, SQL_Field, SQL_Where, SQL_Orderby, pageSize, curPage)
	set Rs = db.execute(SQL)

	'Response.write SQL

	'####### 4. Paging HTML �ۼ� #############################################################################
	totalCount = db_getCount(db, SQL_Table, SQL_Where)
	startRow = totalCount - pageSize * (curPage - 1)
	pageHtml = getPageHtml(pageSector, pageSize, totalCount, curPage, currentURL & "?" & pageWHERE)
%>

<script>
<!--
	function goLIST(frm){
		frm.submit();
	}

	function goSearch(frm){
		frm.submit();
	}
//-->
</script>

<table border="0" cellpadding="0" cellspacing="0" width="940" align="center">
<form method="post" name="searchFrm1" action="Notice.asp">
<input type="hidden" name="SMODE" value=<%=SMODE%>>
<input type="hidden" name="SWORD" value=<%=SWORD%>>
	<tr height="30">
		<td align="right">
			<select name="ACLASS" size="1" class="ComboFFFCE7" onChange="goLIST(document.searchFrm1);">
				<option value=''>:: ��ü ::</option>
				<%=db_getTBCodeSelect("Z04", ACLASS, "N")%>
			</select>
		</td>
	</tr>
</form>
</table>

<table width="940" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">
	<tr height="20" bgcolor="#EEF6FF" align="center">
		<td nowrap width="60">����</td>
		<td nowrap width="80">�з�</td>
		<td>����</td>
		<td nowrap width="50">÷��</td>
		<td nowrap width="80">�Խ���</td>
		<td nowrap width="60">��ȸ��</td>
		<td nowrap width="150">�����</td>
	</tr>
	<tr><td colspan="7" height="1" bgcolor="#FFFFFF" align="center"></td></tr>
	<%
		IF Rs.EOF OR Rs.BOF THEN
	%>
	<tr><td height="50" colspan="7" bgcolor="#FFFFFF" align="center" style="color:#0000FF">��ϵ� ����Ÿ�� �����ϴ�.</td></tr>
	<%
		ELSE
			DO UNTIL Rs.EOF
				db_SEQ = Rs("IDX")
				db_ACLASS = Rs("ACLASS")
				db_TITLE = Rs("TITLE")
				db_FRONTYN = Rs("FRONTYN")
				db_FILENAME1 = Rs("FILENAME1")
					IF len(db_FILENAME1)>0 THEN
						Filename_Temp = split(db_FILENAME1,".")
						FileType = FormatFile(Filename_Temp(1))
					END IF
				db_READCNT = Rs("READCNT")
				db_INDATE = Rs("INDATE")
				db_INCODE = db_getUserName(Rs("INCODE"))
	%>
	<tr bgcolor="#FFFFFF">
		<td align="center"><%=startRow%></td>
		<td class="TDCont"><font color="#000000">[<%=db_getCodeName("Z04",db_ACLASS)%>]</font> </td>
		<td><input type="checkbox" name="FRONTYN" <%IF db_FRONTYN="Y" THEN%>checked<%END IF%> class="none" disabled> <a href="Notice_Detail.asp?isType=VIEW&SEQ=<%=db_SEQ%>&curPage=<%=curPage%>&<%=pageWHERE%>" class="Link1"><%=db_TITLE%></a></td>
		<td align="center"><%IF len(db_FILENAME1)>0 THEN%><a href="/Upload/Board/Notice/Download.asp?filename=<%=db_FILENAME1%>"><img src="/Images/File/<%=FileType%>" align="absmiddle" title="<%=db_FILENAME1%>"></a><%END IF%></td>
		<td align="center"><%=db_INCODE%></td>
		<td align="center"><%=db_READCNT%></td>
		<td align="center"><%=db_INDATE%></td>
	</tr>
	<%
				startRow = startRow - 1
				Rs.MoveNext
			LOOP
		END IF

		Rs.close
		set Rs = Nothing
	%>
</table>

<%'####### ����¡ ó�� ############################################################################### %>
<table border="0" cellpadding="0" cellspacing="0" width="940" align="center">
<form method="post" name="searchFrm" action="Notice.asp">
<input type="hidden" name="ACLASS" value="<%=ACLASS%>">
	<tr><td colspan="4" height="5"></td></tr>
	<tr><td colspan="4" height="1" bgcolor="#D6D6D6"></td></tr>
	<tr height="25" bgcolor="#EEF6FF">
		<td class="TDL5px"><%=pageHtml%></td>
		<td nowrap width="200" class="TDR5px">
			<select name="SMODE" size="1" class="ComboFFFCE7">
				<option value="TITLE">����</option>
				<option value="CONTENTS">����</option>
			</select>
			<input type="text" name="SWORD" value="<%=SWORD%>" size="20" onfocus="setFocusColor(this);" onblur="setOutColor(this);">
		</td>
		<td nowrap width="60" class="TDR5px"><img src="/Images/Btn/BtnSearch.gif" style="cursor:hand;" onClick="goSearch(document.searchFrm);"></td>
		<td nowrap width="60" class="TDR5px"><img src="/Images/Btn/BtnWrite.gif" style="cursor:hand;" onClick="location.href='Notice_Detail.asp?isType=INS';"></td>
	</tr>
	<tr><td  colspan="4" height="1" bgcolor="#D6D6D6"></td></tr>
</form>
</table>

<!-- #include virtual="/Include/Bottom.asp" -->