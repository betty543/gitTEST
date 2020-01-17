<!-- #include virtual="/include/top_frame.asp" -->

<%
	Seq = Request("Seq")						'KEY
	class_gb = Request("class_gb")				'클래스 구분(A:1차분류, B:2차분류, C:3차분류, D:4차분류, E:5차분류)
	db_flag = Request("db_flag")				'DB 처리 구분(INS:INSERT, UP:UPDATE, DEL:DELETE)
	
	Aclass = Request("Aclass")				'1차분류
	Bclass = Request("Bclass")				'2차분류
	Cclass = Request("Cclass")				'3차분류
	Dclass = Request("Dclass")				'3차분류
	Eclass = Request("Eclass")				'3차분류
		
	'Response.end
	pageUrl = "ProType.asp?Aclass="&Aclass&"&Bclass="&Bclass&"&Cclass="&Cclass&"&Dclass="&Dclass
	sql = "SELECT ACLASS, (SELECT CLASSNAME FROM TB_ARMYINFO WHERE ACLASS = '" &Aclass& "' AND BCLASS IS NULL AND CCLASS IS NULL ) as aClassName"
	sql = sql& ", BCLASS, (SELECT CLASSNAME FROM TB_ARMYINFO WHERE ACLASS = '" &Aclass& "' AND BCLASS = '" &Bclass& "' AND CCLASS IS NULL) as bClassName"
	sql = sql& ", CCLASS, (SELECT CLASSNAME FROM TB_ARMYINFO WHERE ACLASS = '" &Aclass& "' AND BCLASS = '" &Bclass& "' AND CCLASS = '" &Cclass& "' AND DCLASS IS NULL) as cClassName"
	sql = sql& ", DCLASS, (SELECT CLASSNAME FROM TB_ARMYINFO WHERE ACLASS = '" &Aclass& "' AND BCLASS = '" &Bclass& "' AND CCLASS = '" &Cclass& "' AND DCLASS = '" &Dclass& "' AND ECLASS IS NULL) as dClassName"
	sql = sql& ", ECLASS, (SELECT CLASSNAME FROM TB_ARMYINFO WHERE ACLASS = '" &Aclass& "' AND BCLASS = '" &Bclass& "' AND CCLASS = '" &Cclass& "' AND DCLASS = '" &Dclass& "' AND ECLASS = '" &Eclass& "') as eClassName"
	sql = sql& ", UseYN, COUNSELORYN"
	sql = sql& ", KEYWORD"
	sql = sql& ", TELNO"
	sql = sql& ", TELNO2"
	sql = sql& " FROM TB_ARMYINFO "
	if db_flag = "UP" then
		sql = sql& "WHERE SEQ = '" &Seq& "'"
	else
		sql = sql& "WHERE ISNULL(ACLASS, ' ') = '" &NullString(Aclass)& "'"
		sql = sql& " AND ISNULL(BCLASS, ' ') = '" &NullString(Bclass)& "'"
		sql = sql& " AND ISNULL(CCLASS, ' ') = '" &NullString(Cclass)& "'"
		sql = sql& " AND ISNULL(DCLASS, ' ') = '" &NullString(Dclass)& "'"
		sql = sql& " AND ISNULL(ECLASS, ' ') = '" &NullString(Eclass)& "'"
	end if
	'LogWrite "SQL="&SQL, "ProType_Detail.asp", "/Setup/ProType/"
	set Rs = db.execute(sql)
	
	if Not (Rs.EOF or Rs.BOF) then
		A_code = Rs("ACLASS")
		A_name = Rs("aClassName")
		B_code = Rs("BCLASS")
		B_name = Rs("bClassName")
		C_code = Rs("CCLASS")
		C_name = Rs("cClassName")

		D_code = Rs("DCLASS")
		D_name = Rs("dClassName")

		E_code = Rs("ECLASS")
		E_name = Rs("eClassName")

KEYWORD = Rs("KEYWORD")
TELNO = Rs("TELNO")
TELNO2 = Rs("TELNO2")

		db_COUNSELORYN = Rs("COUNSELORYN")
		db_UseYN = Rs("UseYN")
	end if
	
	Rs.Close
	Set Rs = Nothing
%>

<script>
<!--
	function fn_inup(f) {
		if(!FieldChk(f.code,"코드")) return false;
		if(!FieldChk(f.code_name,"코드명")) return false;
		
		f.submit();
	}
//-->	
</script>

<body leftmargin="0" marginwidth="0" topmargin="0" marginheight="0" onLoad="ifHeight('ProTypeDetailFrame');inUpFrm.code.focus();">

<div name="ifr" id="ifr">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
<form name="inUpFrm" method="post" action="ProType_InsUpDel.asp">
<input type="hidden" name="Seq" value="<%=Seq%>">
<input type="hidden" name="class_gb" value="<%=class_gb%>">
<input type="hidden" name="db_flag" value="<%=db_flag%>">
<input type="hidden" name="Aclass" value="<%=Aclass%>">
<input type="hidden" name="Bclass" value="<%=Bclass%>">
<input type="hidden" name="Cclass" value="<%=Cclass%>">
<input type="hidden" name="Dclass" value="<%=Dclass%>">
<input type="hidden" name="Eclass" value="<%=Eclass%>">
	<tr>
		<td nowrap width="100" bgcolor="#F3F3F3" class="TDCont">분류</td>
		<td bgcolor="#FFFFFF" class="TDL5px" colspan="7">
			<%
				PartGB = " <font face='webdings' size='2' color='#000000'>4</font> "
				Select Case class_gb
				Case "A"
					Response.Write("1차분류")
					s_code = A_code
					s_name = A_name
				Case "B"
					if (Aclass <> "") then
						Response.Write(A_name&PartGB&"2차분류")
						s_code = B_code
						s_name = B_name
					else
						Call FrameMsgGoUrl(pageUrl, "1차분류 선택이 되지않았습니다.\n1차분류를 다시 선택하신 후 추가 하시기 바랍니다.")
					end if
				Case "C"
					if (Aclass <> "" and Bclass <> "") then
						Response.Write(A_name&PartGB&B_name&PartGB&"3차분류")
						s_code = C_code
						s_name = C_name
					elseif (Mgroup = "" and Mname = "") then
						Call FrameMsgGoUrl(pageUrl, "2차분류 선택이 되지않았습니다.\2차분류를 다시 선택하신 후 추가 하시기 바랍니다.")
					end if
				Case "D"
					if (Aclass <> "" and Bclass <> "" and Cclass <> "") then
						Response.Write(A_name&PartGB&B_name&PartGB&C_name&PartGB&"4차분류")
						s_code = D_code
						s_name = D_name
					elseif (Mgroup = "" and Mname = "") then
						Call FrameMsgGoUrl(pageUrl, "3차분류 선택이 되지않았습니다.\3차분류를 다시 선택하신 후 추가 하시기 바랍니다.")
					end if
				Case "E"
					if (Aclass <> "" and Bclass <> "" and Cclass <> "" and Dclass <> "") then
						Response.Write(A_name&PartGB&B_name&PartGB&C_name&PartGB&D_name&PartGB&"5차분류")
						s_code = E_code
						s_name = E_name
					elseif (Mgroup = "" and Mname = "") then
						Call FrameMsgGoUrl(pageUrl, "4차분류 선택이 되지않았습니다.\4차분류를 다시 선택하신 후 추가 하시기 바랍니다.")
					end if
				End Select
			%>
		</td>
	</tr>
	<tr>
		<td width="104" bgcolor="#F3F3F3" class="TDCont">코드</td>
		<td bgcolor="#FFFFFF" colspan="7">
			<input name="code" type="text" value="<%=s_code%>" size="18" onfocus="setFocusColor(this)" onblur="setOutColor(this)" <% if db_flag = "UP" then Response.Write("readonly") end if%> maxlength="10" tabindex="1">
			<input type="radio" name="UseYN" value="Y" <%IF db_UseYN="Y" OR db_UseYN="" THEN%>checked<%END IF%> class="none">사용 
			<input type="radio" name="UseYN" value="N" <%IF db_UseYN="N" THEN%>checked<%END IF%> class="none">미사용	
		</td>
	</tr>
	<tr>
		<td width="104" bgcolor="#F3F3F3" class="TDCont">코드명</td>
		<td bgcolor="#FFFFFF" colspan="7">
			<input name="code_name" type="text" value="<%=s_name%>" class="input"  size="18" onfocus="setFocusColor(this)" onblur="setOutColor(this)" maxlength="50" align="absmiddle" tabindex="2">
			<% if class_gb = "A" then%>
			　<img src="/Images/Btn/BtnSubmit.gif" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_inup(inUpFrm);">
			<% end if %>
		</td>
	</tr>
	<% if class_gb = "B" or class_gb = "C" or class_gb = "D" or class_gb = "E" then%>
	<tr>
		<td width="104" bgcolor="#F3F3F3" class="TDCont">상담관배치여부</td>
		<td bgcolor="#FFFFFF" >
			<input type="radio" name="COUNSELORYN" value="Y" <%IF db_COUNSELORYN="Y" THEN%>checked<%END IF%> class="none">배치 
			<input type="radio" name="COUNSELORYN" value="N" <%IF db_COUNSELORYN="N" OR db_COUNSELORYN=""  THEN%>checked<%END IF%> class="none">미배치	
		</td>

		<td width="104" bgcolor="#F3F3F3" class="TDCont">검색키워드</td>
		<td bgcolor="#FFFFFF">
			<input name="keyword" type="text" value="<%=KEYWORD%>" size="20">
		</td>
		<td width="104" bgcolor="#F3F3F3" class="TDCont">군전화</td>
		<td bgcolor="#FFFFFF">
			<input name="telno" type="text" value="<%=TELNO%>" size="14">
		</td>
		<td width="104" bgcolor="#F3F3F3" class="TDCont">일반전화</td>
		<td bgcolor="#FFFFFF">
			<input name="telno2" type="text" value="<%=TELNO2%>" size="40">
			　　<img src="/Images/Btn/BtnSubmit.gif" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_inup(inUpFrm);">
		</td>
	</tr>

	<% end if %>
</form>		
</table>
</div>