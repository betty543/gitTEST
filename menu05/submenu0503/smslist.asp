
<!-- #include virtual="/Include/Top.asp" -->

<%
	QueryYN = request("QueryYN")
	FromDate = request("FromDate")
	if FromDate ="" then
		FromDate = date()
	end if
	ToDate = request("ToDate")
	if ToDate ="" then
		ToDate = date()
	end if
	whereCD1 = Trim(request("whereCD1"))
	whereCD2 = Trim(request("whereCD2"))
	whereCD3 = Trim(request("whereCD3"))

	pageWHERE = "QueryYN="&QueryYN&"&FromDate="&FromDate&"&ToDate="&ToDate&"&whereCD1="&whereCD1&"&whereCD2="&whereCD2&"&whereCD3="&whereCD3

%>

<div id='CalendarLayer' style='display:none;'><iframe name='CalendarFrame' src='/include/Calendar.asp' width='172' height='177' border='0' frameborder='0' scrolling='no'></iframe></div>

<script>

	function fn_Search() {

		document.inUpFrm.QueryYN.value = "Y";
		document.inUpFrm.submit();
	}
	
	function fn_Xls() {
		location.href="smslist_XLS.asp?<%=pageWHERE%>"
	}
</script>

<table border="0" width="940" cellpadding="0" cellspacing="3" bgcolor="#EFECE5" align="center">
	<tr bgcolor="#FFFFFF">
		<td>
		
			<form method="post" name="inUpFrm" style="margin:0">
			<input type="hidden" name="QueryYN" value="">
			<table width="100%" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">

			    <tr>
			        <td width="60" bgcolor="#EEF6FF" class="TDCont" align='center'>조회기간</td>
			        <td  bgcolor="#FFFFFF" colspan=3 width=200>
			        	<input value="<%=FromDate%>" name="FromDate" readonly type="text" size="10" onfocus="setFocusColor(this);" onClick="new CalendarFrame.Calendar(this);" onblur="setOutColor(this);">
				    	~
				    	<input value="<%=ToDate%>" name="ToDate" readonly type="text" size="10" onfocus="setFocusColor(this);" onClick="new CalendarFrame.Calendar(this);" onblur="setOutColor(this);">	
			        </td>

						<td width="60" bgcolor="#EEF6FF" class="TDCont" align='center'>발송자</td>
						<td bgcolor="#FFFFFF">
						<%
							'======= 상담원 가져오기 ==================================================
							SqlCode = "SELECT USERID, USERNAME FROM TB_USERINFO"
							SqlCode = SqlCode& " WHERE USEYN='Y' "
							SqlCode = SqlCode& " ORDER BY USEYN DESC, GRADE ASC, USERNAME ASC"
							set RsCode = db.execute(SqlCode)
						%>
						<select name="whereCD1" size="1" class="ComboFFFCE7">
							<option value="">선택</option>
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("USERID")
										CODENAME = RsCode("USERNAME")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &whereCD1& "")%>
							<%
									RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
							%>
						<%
							'======= 상담원 가져오기 ==================================================
							SqlCode = "SELECT USERID, USERNAME FROM TB_USERINFO"
							SqlCode = SqlCode& " WHERE USEYN='N'  and	outdate >= '"&DateAdd("d",1,DateAdd("m",-1,Date())) &"'"
							SqlCode = SqlCode& " ORDER BY USEYN DESC, GRADE ASC, USERNAME ASC"
							set RsCode = db.execute(SqlCode)

								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("USERID")
										CODENAME = "[퇴직]"&RsCode("USERNAME")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &whereCD1& "")%>
							<%
									RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
							%>
						</select>
				
						</td>

						<td width="60" bgcolor="#EEF6FF" class="TDCont" align='center'>휴대폰<br>번호</td>
						<td bgcolor="#FFFFFF"><input type="text" name="whereCD2" value="<%=whereCD2%>" maxlength="15" size="15" onfocus="setFocusColor(this);" onblur="setOutColor(this);" align="absmiddle"></td>


						<td width="60" bgcolor="#EEF6FF" class="TDCont" align='center'>수신자</td>
						<td bgcolor="#FFFFFF"><input type="text" name="whereCD3" value="<%=whereCD3%>" maxlength="15" size="15" onfocus="setFocusColor(this);" onblur="setOutColor(this);" align="absmiddle"></td>


			        <td colspan='2' rowspan="2" bgcolor="#FFFFFF" align="center">
			        	<img src="/Images/Btn/BtnSearch.gif" style="cursor:hand;" onClick="fn_Search();">
			        	<%IF SS_Login_Secgroup="A" Or SS_Login_Secgroup="B" THEN%><br><br><img src="/Images/Btn/BtnExcel.gif" style="cursor:hand;" onClick="fn_Xls();"><%END IF%>
			        </td>
				</tr>

			</table>
			</form>
		</td>
	</tr>
</table>


<table width="940" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
<table border="0" width="940" cellspacing="1" cellpadding="0" width="100%" bgcolor="#CCCCCC" align="center">
	<tr height="25" bgcolor="#F3F3F3" align="center">
		<td><b>순번</b></td>
		<td><b>전송요청일시</b></td>
		<td><b>그룹</b></td>
		<td><b>전송자</b></td>
		<td><b>수신휴대폰</b></td>
		<td><b>발신번호</b></td>
		<td><b>전송내용</b></td>
		<td ><b>전송결과</b></td>
	</tr>

<%
	if QueryYN = "Y" then		

		SQL = "	SELECT SM_SDMBNO,SM_RVMBNO, SM_MSG, '2' as SM_STATUS, convert(char(19),SM_Indate,121) as SM_Sdate, SM_CODE1, SM_CODE2"
		SQL = SQL & "	FROM	SMS.DBO.SMS_Reserve"
		SQL = SQL & "	WHERE	CONVERT(CHAR(10),SM_Indate,121) >= '" & FROMDATE & "'"
		SQL = SQL & "	AND		CONVERT(CHAR(10),SM_Indate,121) <= '" & ToDate & "'"
		'전송요청자
		IF whereCD1 <> "" THEN
			SQL = SQL & "	AND		SM_CODE1 = '" & whereCD1 & "'"
		END IF
		'전화번호
		IF whereCD2 <> "" THEN
			SQL = SQL & "	AND		SM_SDMBNO LIKE '%" & whereCD3 & "%'"
		END IF
		'수신자
		IF whereCD2 <> "" THEN
			SQL = SQL & "	AND		SM_CODE2 LIKE '%" & whereCD2 & "%'"
		END IF
		SQL = SQL & "	UNION ALL "
		SQL = SQL & "	SELECT SM_SDMBNO,SM_RVMBNO, SM_MSG,SM_STATUS, convert(char(19),SM_Sdate,121) as SM_Sdate, SM_CODE1, SM_CODE2"
		SQL = SQL & "	FROM	SMS.DBO.SMS_BACK"
		SQL = SQL & "	WHERE	CONVERT(CHAR(10),SM_Sdate,121) >= '" & FROMDATE & "'"
		SQL = SQL & "	AND		CONVERT(CHAR(10),SM_Sdate,121) <= '" & ToDate & "'"
		'전송요청자
		IF whereCD1 <> "" THEN
			SQL = SQL & "	AND		SM_CODE1 = '" & whereCD1 & "'"
		END IF
		'전화번호
		IF whereCD2 <> "" THEN
			SQL = SQL & "	AND		SM_SDMBNO LIKE '%" & whereCD3 & "%'"
		END IF
		'수신자
		IF whereCD2 <> "" THEN
			SQL = SQL & "	AND		SM_CODE2 LIKE '%" & whereCD2 & "%'"
		END IF
		SQL = SQL & "	ORDER BY SM_Sdate desc"
		SET RS = DB.EXECUTE(SQL)
		SET RS = DB.EXECUTE(SQL)

i = 0
		DO UNTIL RS.EOF
			i = i + 1
			sDate = RS("SM_Sdate")
			sGROUP = db_getCodeName("Z04",RS("SM_CODE2"))
			sUSERID = db_GetUSERNAME(RS("SM_CODE1"))
			sCELLPHONE = RS("SM_SDMBNO")
			sREPLYPHONE = RS("SM_RVMBNO")
			sMESSAGE = RS("SM_MSG")
			if RS("SM_STATUS") = "1" then
				sRESULT = "성공"
			elseif RS("SM_STATUS") = "2" then
				sRESULT = "예약"
			else
				sRESULT = "실패"
			end if
%>
	<tr height="25" bgcolor="#ffffff" align="center">
		<td><%=i%></td>
		<td><%=sDate%></td>
		<td><%=sGROUP%></td>
		<td><%=sUSERID%></td>
		<td><%=sCELLPHONE%></td>
		<td><%=sREPLYPHONE%></td>
		<td title="<%=sMESSAGE%>" align='left'>&nbsp;<%=CutString(sMESSAGE, 30, "...")%></td>
		<td ><%=sRESULT%></td>
	</tr>


<%
			RS.MOVENEXT
		LOOP

	end if
%>


</table>


<!-- #include virtual="/Include/Bottom.asp" -->