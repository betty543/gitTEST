<!-- #include virtual="/Include/Top.asp" -->
<%

SS_LoginID = SESSION("SS_LoginID")
SS_Login_Secgroup = SESSION("SS_Login_Secgroup")

	QueryYN = request("QueryYN")
	FromDate = request("FromDate")
	ToDate = request("ToDate")
	curPage = request("curPage")
	whereCD1 = Trim(request("whereCD1")) '성별
	whereCD2 = Trim(request("whereCD2")) '상담방법
	whereCD3 = Trim(request("whereCD3")) '의뢰인
	whereCD4 = Trim(request("whereCD4")) '상담분야
	whereCD5 = Trim(request("whereCD5")) '소속
	whereCD6 = Trim(request("whereCD6")) '계급구분
	whereCD7 = Trim(request("whereCD7")) '계급구분2
	whereCD8 = Trim(request("whereCD8"))	'성명
	whereCD9 = Trim(request("whereCD9"))	'전화번호
	whereCD10 = Trim(request("whereCD10"))	'소속
	whereCD11 = Trim(request("whereCD11"))	'처리결과
	whereCD12 = Trim(request("whereCD12"))	'처리결과

whereCD5_A = Trim(request("whereCD5_A")) '소속
whereCD5_B = Trim(request("whereCD5_B")) '소속
whereCD5_C = Trim(request("whereCD5_C")) '소속
whereCD5_E = Trim(request("whereCD5_E")) '소속
whereCD5_F = Trim(request("whereCD5_F")) '소속
whereCD6_A = Trim(request("whereCD6_A")) '계급구분
whereCD6_B = Trim(request("whereCD6_B")) '계급구분
whereCD6_C = Trim(request("whereCD6_C")) '계급구분

	if FromDate = "" then
		FromDate = date()
	end if
	if ToDate = "" then
		ToDate = date()
	end if

	'2. 쿼리조건절 셋팅
	pageSize = 10
	pageSector = 10
	if curPage = "" then curPage = 1 end If


	where1 = "FromDate=" & FromDate & "&ToDate=" & ToDate & "&whereCD1=" & whereCD1 & "&whereCD2=" & whereCD2 & "&whereCD3=" & whereCD3 & "&whereCD4=" & whereCD4 & "&whereCD5=" & whereCD5 & "&whereCD6=" & whereCD6 & "&whereCD7=" & whereCD7 & "&whereCD8=" & whereCD8 & "&whereCD9=" & whereCD9 & "&whereCD10=" & whereCD10 & "&whereCD11=" & whereCD11 & "&whereCD12=" & whereCD12
	where1 = where1 & "&whereCD12=" & whereCD12 & "&whereCD5_A=" & whereCD5_A& "&whereCD5_B=" & whereCD5_B& "&whereCD5_C=" & whereCD5_C& "&whereCD5_D=" & whereCD5_D& "&whereCD5_E=" & whereCD5_E
	where1 = where1 & "&whereCD6_A=" & whereCD6_A& "&whereCD6_B=" & whereCD6_B& "&whereCD6_C=" & whereCD6_C

	where2 = "curPage=" & curPage & "&" & where1

	'SQL = "	SELECT *, CONVERT(VARCHAR(19),JUBTIME,121) JUBTIME1   FROM TB_LIFECALLHISTORY"
	sql_where =	"JUBDATE >= '" & FromDate & "'"
	sql_where = sql_where & "	AND     JUBDATE <= '" & ToDate & "'"

	IF whereCD1 <> "" THEN
		sql_where = sql_where & "	AND     SEXGB = '" & whereCD1 & "'"
	END IF
	IF whereCD2 <> "" THEN
		sql_where = sql_where & "	AND     CALLKIND_B = '" & whereCD2 & "'"
	END IF
	IF whereCD3 <> "" THEN	'상담종류
		'sql_where = sql_where & "	AND     ACLASS = '" & whereCD3 & "'"
	END IF
	IF whereCD4 <> "" THEN
		'sql_where = sql_where & "	AND     CONSULTGB = '" & whereCD4 & "'"
	END IF
	IF whereCD5 <> "" THEN '소속
		'sql_where = sql_where & "	AND     SOSOKGB = '" & whereCD5 & "'"
	END IF
IF whereCD5_A <> "" THEN '소속
	sql_where = sql_where & "	AND SOSOKGB_A = '" & whereCD5_A & "'"
END IF
IF whereCD5_B <> "" THEN '소속
	sql_where = sql_where & "	AND SOSOKGB_B = '" & whereCD5_B & "'"
END IF
IF whereCD5_C <> "" THEN '소속
	sql_where = sql_where & "	AND SOSOKGB_C = '" & whereCD5_D & "'"
END IF
IF whereCD5_D <> "" THEN '소속
	sql_where = sql_where & "	AND SOSOKGB_D = '" & whereCD5_E & "'"
END IF
IF whereCD5_E <> "" THEN '소속
	sql_where = sql_where & "	AND SOSOKGB_E = '" & whereCD5_E & "'"
END IF

IF whereCD6_A <> "" THEN '계급구분
	sql_where = sql_where & "	AND LEVEL_B = '" & whereCD6_A & "'"
END IF
IF whereCD6_B <> "" THEN '계급구분
	sql_where = sql_where & "	AND LEVEL_C = '" & whereCD6_B & "'"
END IF
IF whereCD6_C <> "" THEN '계급구분
	sql_where = sql_where & "	AND LEVEL_D = '" & whereCD6_C & "'"
END IF

	IF whereCD6 <> "" THEN
		'sql_where = sql_where & "	AND     LEVEL1 = '" & whereCD6 & "'"
	END IF
	IF whereCD7 <> "" THEN
		'sql_where = sql_where & "	AND     LEVEL2 = '" & whereCD7 & "'"
	END IF
	IF whereCD8 <> "" THEN
		sql_where = sql_where & "	AND      ( CUSTNAME LIKE '%" & whereCD8 & "%' or REFERJUBSEQ in ( select JUBSEQ from TB_LIFECALLHISTORY where CUSTNAME LIKE '%" & whereCD8 & "%' ) ) "
	END IF
	IF whereCD9 <> "" THEN
		sql_where = sql_where & "	AND     (TELNO LIKE '%" & whereCD9 & "%' OR TELNO2 LIKE '%" & whereCD9 & "%')"
	END IF
	IF whereCD10 <> "" THEN
		sql_where = sql_where & "	AND     INCODE = '" & whereCD10 & "'"
	END IF
	'if SS_Login_Secgroup = "A" then
		'내것만
	'	sql_where = sql_where& " AND	INCODE = '"&SS_LoginID&"'"
	'end if

	IF whereCD11 <> "" THEN
		sql_where = sql_where & "	AND     REQUESTERGB = '" & whereCD11 & "'"
	END IF

	IF whereCD12 <> "" THEN
		'sql_where = sql_where & "	AND     EMERYN = '" & whereCD12 & "'"
	END IF



	'Set Rs = server.createObject("ADODB.Recordset")
	'Rs.open SQL,db


	sql_tb = "TB_LIFECALLHISTORY_OB"
	'sql_index = "index_desc(" & sql_tb & " IDX_TB_CALLHISTORY_JUBSEQ)"
	sql_field ="*, CONVERT(VARCHAR(19),JUBTIME,121) JUBTIME1"
	sql_orderby = "JUBTIME DESC"

	'3. 쿼리 실행
	sql = db_getSqlWithPage(sql_tb, sql_index, sql_field, sql_where, sql_orderby, pageSize, curPage)
	set Rs = db.execute(sql)

	'Response.Write sql

	'4. Paging HTML 작성
	totalCount = db_getCount(db, sql_tb, sql_where)
	startRow = totalCount - pageSize * (curPage - 1)
	pageHtml = getPageHtml(pageSector, pageSize, totalCount, curPage, currentURL & "?" & where1)


%>

<script>

	function fn_SetLevel2()
	{
		frame_level.location = "frame_level.asp?level="+document.all.whereCD6.value+"&level2=";
	}

	function fn_Search()
	{
		inUpFrm.submit();
	}

</script>


<div id='CalendarLayer' style='display:none;'><iframe name='CalendarFrame' src='/include/Calendar.asp' width='172' height='177' border='0' frameborder='0' scrolling='no'></iframe></div>

<!-- #include virtual="/Include/PopLayer.asp" -->
<table border="0" width="1200" cellpadding="0" cellspacing="3" bgcolor="#EFECE5" align="center">
		
	<form method="post" name="inUpFrm" action="<%=Menu_2nd%>" style="margin:0">
		
	<tr bgcolor="#FFFFFF">
		<td>

			<input type="hidden" name="QueryYN" value="">
			<input type="hidden" name="whereCD7" value="<%=whereCD7%>">
			<input type="hidden" name="whereCD5_A" value="<%=whereCD5_A%>">
			<input type="hidden" name="whereCD5_B" value="<%=whereCD5_B%>">
			<input type="hidden" name="whereCD5_C" value="<%=whereCD5_C%>">
			<input type="hidden" name="whereCD5_D" value="<%=whereCD5_D%>">
			<input type="hidden" name="whereCD5_E" value="<%=whereCD5_E%>">
			<input type="hidden" name="whereCD6_A" value="<%=whereCD6_A%>">
			<input type="hidden" name="whereCD6_B" value="<%=whereCD6_B%>">
			<input type="hidden" name="whereCD6_C" value="<%=whereCD6_C%>">
			
			<table width="100%" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">
				<tr>
					<td width="80" bgcolor="#EFEFEF" class="TDCont" align='center'>조회기간</td>
					<td bgcolor="#FFFFFF" colspan="3">&nbsp;
						<input value="<%=FromDate%>" name="FromDate" readonly type="text" size="10" onfocus="setFocusColor(this); new CalendarFrame.Calendar(this);" onblur="setOutColor(this);" />
						~
						<input value="<%=ToDate%>" name="ToDate" readonly type="text" size="10" onfocus="setFocusColor(this); new CalendarFrame.Calendar(this);" onblur="setOutColor(this);">
					</td>
					<td bgcolor="#EFEFEF" class="TDCont" width=80 align='center'>성별</td>
					<td bgcolor="#FFFFFF" width="130" nowrap >
						<input type="radio" name="whereCD1" value="" class="none" <%if whereCD1 = "" then%>checked<%end if%> >&nbsp;전체
						<input type="radio" name="whereCD1" value="1" class="none" <%if whereCD1 = "1" then%>checked<%end if%> >&nbsp;남
						<input type="radio" name="whereCD1" value="2" class="none" <%if whereCD1 = "2" then%>checked<%end if%> >&nbsp;녀
					</td>
					<td bgcolor="#EFEFEF" class="TDCont" width=80 align='center'>상담관</td>
					<td bgcolor="#FFFFFF"  nowrap>
						<%
						'======= 처리구분 코드 가져오기 ==================================================
						SqlCode = "SELECT USERID, USERNAME FROM TB_USERINFO"
						SqlCode = SqlCode& " WHERE USEYN='Y'"
						SqlCode = SqlCode& " AND	GRADE='B'" '생명의전화 그룹
						SqlCode = SqlCode& " AND	SECGROUP = 'A'" '생명의전화 그룹
						if SS_Login_Secgroup = "A" then
							'내것만
							'SqlCode = SqlCode& " AND	USERID = '"&SS_LoginID&"'"
						end if
						SqlCode = SqlCode& " ORDER BY USERID"
						set RsCode = db.execute(SqlCode)
						%>
						&nbsp;<select name="whereCD10" size="1" class="ComboFFFCE7">
							<Option value ='' selected>상담관선택</option>
							<%
							IF NOT(RsCode.Eof OR RsCode.bof) THEN
								DO until RsCode.EOF
									CODE = RsCode("USERID")
									CODENAME = RsCode("USERNAME")
									%>
									<%=printSelect("" &CODENAME& "","" &CODE& "","" &whereCD10& "")%>
									<%
									RsCode.MoveNext
								LOOP
							END IF
							RsCode.Close
							set RsCode = NOTHING
							%>
						</select>
					</td>
					<td rowspan="3" bgcolor="#FFFFFF" align="center">
						<img src="/Images/Btn/BtnSearch.gif" style="cursor:hand;" onClick="fn_Search();">
					</td>
				</tr>
				<tr valign="center" style="display:block;">
					<td bgcolor="#EFEFEF" class="TDCont" width=80 align='center'>성명/원상담자</td>
					<td bgcolor="#FFFFFF">&nbsp;<input value="<%=whereCD8%>" name="whereCD8" type="text" size="20" onfocus="setFocusColor(this);" onblur="setOutColor(this);"></td>
					<td bgcolor="#EFEFEF" class="TDCont" width=80 align='center'>전화번호</td>
					<td bgcolor="#FFFFFF" >&nbsp;<input value="<%=whereCD9%>" name="whereCD9" type="text" size="15" onfocus="setFocusColor(this);" onblur="setOutColor(this);"></td>
					<td bgcolor="#EFEFEF" class="TDCont" width=70 align='center'>후속확인</td>
					<td bgcolor="#FFFFFF">
						<%
						'======= 처리구분 코드 가져오기 ==================================================
						SqlCode = "SELECT CODE,  CODENAME FROM TB_CODE"
						SqlCode = SqlCode& " WHERE CODEGROUP = 'C13'"
						SqlCode = SqlCode& " ORDER BY CODE"
						set RsCode = db.execute(SqlCode)
						%>
						&nbsp;<select name="whereCD2" size="1" class="ComboFFFCE7" >
							<Option value ='' selected>후속확인선택</option>
							<%
							IF NOT(RsCode.Eof OR RsCode.bof) THEN
								DO until RsCode.EOF
									CODE = RsCode("CODE")
									CODENAME = RsCode("CODENAME")
									%>
									<%=printSelect("" &CODENAME& "","" &CODE& "","" &whereCD2& "")%>
									<%
									RsCode.MoveNext
								LOOP
							END IF
							RsCode.Close
							set RsCode = NOTHING
							%>
						</select>
					</td>
					<td bgcolor="#EFEFEF" class="TDCont" align='center' nowrap>원상담자와의관계</td>
					<td bgcolor="#FFFFFF" nowrap>
						<%
						'======= 처리구분 코드 가져오기 ==================================================
						SqlCode = "SELECT CODE, CODENAME FROM TB_CODE"
						SqlCode = SqlCode& " WHERE CODEGROUP='C14'"
						SqlCode = SqlCode& " ORDER BY CODE"
						set RsCode = db.execute(SqlCode)
						%>
						&nbsp;<select name="whereCD11" size="1" class="ComboFFFCE7">
							<Option value ='' selected>원상담자</option>
							<%
							IF NOT(RsCode.Eof OR RsCode.bof) THEN
								DO until RsCode.EOF
									CODE = RsCode("CODE")
									CODENAME = RsCode("CODENAME")
									%>
									<%=printSelect("" &CODENAME& "","" &CODE& "","" &whereCD11& "")%>
									<%
									RsCode.MoveNext
								LOOP
							END IF
							RsCode.Close
							set RsCode = NOTHING
							%>
						</select>
					</td>
				</tr>
		    <tr valign="center" style="display:block;">
					<td bgcolor="#EFEFEF" class="TDCont" width="80" align="center">소속</td>
					<td bgcolor="#FFFFFF" colspan="3" nowrap>&nbsp;
						<iframe src="Sosok_FrameA.asp?SOSOK_A=<%=whereCD5_A%>&SOSOK_B=" scrolling="no" frameborder="0" border="0" name="SosokFrameA" width="80"></iframe>
						<iframe src="Sosok_FrameB.asp?SOSOK_A=<%=whereCD5_A%>&SOSOK_B=<%=whereCD5_B%>" scrolling="no" frameborder="0" border="0" name="SosokFrameB" width="80"></iframe>
						<iframe src="Sosok_FrameC.asp?SOSOK_A=<%=whereCD5_A%>&SOSOK_B=<%=whereCD5_B%>&SOSOK_C=<%=whereCD5_C%>" scrolling="no" frameborder="0" border="0" name="SosokFrameC" width="80"></iframe>
						<iframe src="Sosok_FrameD.asp?SOSOK_A=<%=whereCD5_A%>&SOSOK_B=<%=whereCD5_B%>&SOSOK_C=<%=whereCD5_C%>&SOSOK_D=<%=whereCD5_D%>" scrolling="no" frameborder="0" border="0" name="SosokFrameD" width="80"></iframe>
						<iframe src="Sosok_FrameE.asp?SOSOK_A=<%=whereCD5_A%>&SOSOK_B=<%=whereCD5_B%>&SOSOK_C=<%=whereCD5_C%>&SOSOK_D=<%=whereCD5_D%>&SOSOK_E=<%=whereCD5_E%>" scrolling="no" frameborder="0" border="0" name="SosokFrameE" width="80"></iframe>
					</td>
					<td bgcolor="#EFEFEF" class="TDCont" width="80" align="center">계급</td>
					<td bgcolor="#FFFFFF" colspan="3" nowrap>&nbsp;
						<iframe src="Level_FrameA.asp?LEVEL_A=<%=whereCD6_A%>&LEVEL_B=" scrolling="no" frameborder="0" border="0" name="LevelFrameA" width="80"></iframe>
						<iframe src="Level_FrameB.asp?LEVEL_A=<%=whereCD6_A%>&LEVEL_B=<%=whereCD6_B%>" scrolling="no" frameborder="0" border="0" name="LevelFrameB" width="80"></iframe>
						<iframe src="Level_FrameC.asp?LEVEL_A=<%=whereCD6_A%>&LEVEL_B=<%=whereCD6_B%>&LEVEL_C=<%=whereCD6_C%>" scrolling="no" frameborder="0" border="0" name="LevelFrameC" width="80"></iframe>
					</td>
				</tr>
			</table>
			</form>
		</td>
	</tr>
</table>


<table width="1200" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
<table border="0" width="1200" cellspacing="1" cellpadding="0" width="100%" bgcolor="#CCCCCC" align="center">
	<tr height="25" bgcolor="#EEF6FF" align="center">
		<td class="TDCont" width="30" align='center'>순번</td>
		<td class="TDCont" width="40" align='center'>구분</td>
		<td class="TDCont" align='center'><b>상담일시</b></td>
		<td class="TDCont" width="60" align='center'><b>상담시간</b></td>
		<td class="TDCont" width="100" align='center'><b>전화번호</b></td>
		<td class="TDCont" width="100" align='center'><b>원상담자</b></td>
		<td class="TDCont" width="150" align='center'><b>원상담자와의관계</b></td>
		<td class="TDCont" width="80" align='center'><b>후속확인</b></td>
		<td class="TDCont" width="100" align='center'><b>소속</b></td>
		<td class="TDCont" width="100" align='center'><b>계급</b></td>
		<td class="TDCont" width="80" align='center'><b>성명</b></td>
		<td class="TDCont" width="60" align='center'><b>상담관</b></td>
		<td class="TDCont" width="40" align='center'><b>관리</b></td>
	</tr>

<%

	i = 0
	DO UNTIL RS.EOF
	i = i + 1


		db_IDX	= RS("IDX")
		db_JUBSEQ	= RS("JUBSEQ")
		db_JUBDATE	= RS("JUBDATE")
		db_JUBTIME	= RS("JUBTIME")
		db_JUBTIME1	= RS("JUBTIME1")


		db_IOFLAG	= RS("IOFLAG")
		IF db_IOFLAG = "1" THEN
			db_IOFLAG_NM = "인"
		ELSEIF db_IOFLAG = "2" THEN
			db_IOFLAG_NM = "아웃"
		ELSE
			db_IOFLAG_NM = ""
		END IF

		db_EMERYN	= RS("EMERYN")
		db_CUSTNO	= RS("CUSTNO")
		db_CUSTNAME	= RS("CUSTNAME")
		db_TELNO	= RS("TELNO")
		db_TELNO2	= RS("TELNO2")
		db_CID	= RS("CID")
		db_SEXGB	= RS("SEXGB")
		db_CHANNELGB	= RS("CHANNELGB")
		db_REQUESTERGB	= RS("REQUESTERGB")
		db_CONSULTGB	= RS("CONSULTGB")
		db_CONSULTETCGB	= RS("CONSULTETCGB")
		db_SOSOKGB_A	= RS("SOSOKGB_A")
		db_SOSOKGB_B	= RS("SOSOKGB_B")
		db_SOSOKGB_C	= RS("SOSOKGB_C")
		db_SOSOKGB_D	= RS("SOSOKGB_D")
		db_SOSOKGB_E	= RS("SOSOKGB_E")
		db_LEVEL_B	= RS("LEVEL_B")
		db_LEVEL_C	= RS("LEVEL_C")
		db_LEVEL_D	= RS("LEVEL_D")
		db_FAMILYGB	= RS("FAMILYGB")
		db_CALLCLASS_B	= RS("CALLCLASS_B")
		db_CALLCLASS_C	= RS("CALLCLASS_C")
		db_CHANNELGB_B	= RS("CHANNELGB_B")
		db_CHANNELGB_C	= RS("CHANNELGB_C")
		db_CALLFLAG	= RS("CALLFLAG")
		db_CALLKIND_B	= RS("CALLKIND_B")
		db_CALLKIND_C	= RS("CALLKIND_C")
		db_QUESTION	= RS("QUESTION")
		db_REPLY	= RS("REPLY")
		db_REMARK	= RS("REMARK")
		db_RESULTGB	= RS("RESULTGB")
		db_RESERVEDATE	= RS("RESERVEDATE")
		db_RESERVETIME	= RS("RESERVETIME")
		db_PROCESSGB	= RS("PROCESSGB")
		db_WEATHER	= RS("WEATHER")
		db_CALLID	= RS("CALLID")
		db_RECORDFILE	= RS("RECORDFILE")
		db_CALLTIMEDP	= RS("CALLTIMEDP")
		db_CALLTIME	= RS("CALLTIME")
		db_CB_SEQ	= RS("CB_SEQ")
		db_REFERJUBSEQ	= RS("REFERJUBSEQ")
		db_REFCNT	= RS("REFCNT")
		db_FILENAME	= RS("FILENAME")
		db_INCODE	= RS("INCODE")
		db_INDATE	= RS("INDATE")


		IF WEEKDAY(db_JUBDATE)=1 THEN
			JUBDAY="일"
		ELSEIF WEEKDAY(db_JUBDATE)=2 THEN
			JUBDAY="월"
		ELSEIF WEEKDAY(db_JUBDATE)=3 THEN
			JUBDAY="화"
		ELSEIF WEEKDAY(db_JUBDATE)=4 THEN
			JUBDAY="수"
		ELSEIF WEEKDAY(db_JUBDATE)=5 THEN
			JUBDAY="목"
		ELSEIF WEEKDAY(db_JUBDATE)=6 THEN
			JUBDAY="금"
		ELSEIF WEEKDAY(db_JUBDATE)=7 THEN
			JUBDAY="토"
		END IF

%>

		<tr bgcolor="#FFFFFF">

			<td align="center"><%=startRow%></td>

			<td align="center"><%=db_IOFLAG_NM%></td>

			<td align="center"><a href="javascript:fn_Detail('<%=db_JUBSEQ%>','<%=whereCD9%>');"><%=mid(db_JUBTIME1,3)%>(<%=JUBDAY%>)</a></td>
			<td align="center"><a href="javascript:fn_update('<%=db_JUBSEQ%>','UP');"><%=db_CALLTIMEDP%></a></td>
			<td align="center"><% if db_CID = "" then %><%=db_TELNO%><%else%><%=db_CID%><%end if%></td>
			<td align="center"><%IF db_REFERJUBSEQ <> db_JUBSEQ THEN %><%=db_getCustNameA(db_REFERJUBSEQ)%><%END IF%></td>
			<td align="center"><%=db_getCodeName("C14",db_REQUESTERGB)%></td>
			<td align="center"><%=db_getCodeName("C13",db_CALLKIND_B)%></td>


			<td align="left"><%=db_getCateNameA_(db_SOSOKGB_A)%><%if db_getCateNameB_(db_SOSOKGB_A,db_SOSOKGB_B) <> "" then %>><%end if%><%=db_getCateNameB_(db_SOSOKGB_A,db_SOSOKGB_B)%></td>
			<td align="left"><%=db_getCateNameB_("P",db_LEVEL_B)%><%if db_getCateNameC_("P",db_LEVEL_B,db_LEVEL_C) <> "" then %>><%end if%> <%=db_getCateNameC_("P",db_LEVEL_B,db_LEVEL_C)%><%if db_getCateNameD_("P",db_LEVEL_B,db_LEVEL_C,db_LEVEL_D) <> "" then %>><%end if%><%=db_getCateNameD_("P",db_LEVEL_B,db_LEVEL_C,db_LEVEL_D)%></td>
			<td align="center"><%=db_CUSTNAME%></td>
			<td align="center"><%=db_getUserName(db_INCODE)%></td>
			<td align="center">
				<!--<img src="/Images/Btn/BtnIconModify.gif" title="수정" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_update('<%=db_JUBSEQ%>','UP');">-->
				<% if SS_LoginID = db_INCODE or SS_Login_Secgroup <> "A" then %>
				<!--<img src="/Images/Btn/BtnIconDel.gif" title="삭제" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_del('<%=db_JUBSEQ%>','DEL');">-->
				<% end if%>
				<img src="/Images/Comm/IconWrite.gif" title="인쇄" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_print('<%=db_JUBSEQ%>');" >
				<img src="/Images/file/xls.gif" title="저장" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_save('<%=db_JUBSEQ%>');" >
			</td>
		</tr>
<%
		startRow = startRow - 1
		RS.MOVENEXT
	LOOP


%>
		<!--<tr bgcolor="#FFFFFF">
			<td class="TDCont">1</td>
			<td class="TDCont" align="center">2009-01-01 15:15</td>
			<td align="center">전화</td>
			<td align="center">1회</td>
			<td align="center" width=400>1군</td>
			<td align="center">미상</td>
			<td align="center">미상</td>
			<td align="center">김상담</td>
			<td align="center">
				<img src="/Images/Btn/BtnIconModify.gif" title="수정" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_update('1111','UP');">
				<img src="/Images/Btn/BtnIconDel.gif" title="삭제" style="cursor:hand;" align="absmiddle"onclick="javascript:fn_del('1111','DEL');">
				<img src="/Images/Comm/IconWrite.gif" title="인쇄" style="cursor:hand;" align="absmiddle"onclick="javascript:fn_del('1111','DEL');">
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td class="TDCont">2</td>
			<td class="TDCont" align="center">2009-01-01 19:00</td>
			<td align="center">전화</td>
			<td align="center">1회</td>
			<td align="center" width=400>2군</td>
			<td align="center">일병</td>
			<td align="center">손민경</td>
			<td align="center">김상담</td>
			<td align="center">
				<img src="/Images/Btn/BtnIconModify.gif" title="수정" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_update('1111','UP');">
				<img src="/Images/Btn/BtnIconDel.gif" title="삭제" style="cursor:hand;" align="absmiddle"onclick="javascript:fn_del('1111','DEL');">
				<img src="/Images/Comm/IconWrite.gif" title="인쇄" style="cursor:hand;" align="absmiddle"onclick="javascript:fn_del('1111','DEL');">
				
			</td>
		</tr>-->

</table>
<table width="1200" cellpadding="0" cellspacing="0" width="100%" align="center">
	<tr><td height="2" bgcolor="#f2f2f2"></td></tr>
	<tr><td height="5"></td></tr>
	<tr><td bgcolor="#D6D6D6" height="1"></td></tr>
	<tr><td bgcolor="#EEF6FF" class="TDL10px" height="25"><%=pageHtml%></td></tr>
	<tr><td bgcolor="#D6D6D6" height="1"></td></tr>
	<tr>
		<td height="30" class="TDR10px">
			<img src="/Images/Btn/BtnAdd.gif" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_new();">&nbsp;<img src="/Images/Btn/BtnSaveDaily.GIF" style="cursor:hand;" align="absmiddle" onClick="fn_printA();">
		</td>
	</tr>
</table>

<table width="1200" cellpadding="0" cellspacing="0" width="100%" align="center">

	<tr><td bgcolor="#D6D6D6" colspan='2'><iframe src="History_Detail.asp" scrolling="auto" frameborder="0" border="0" name="historyDetail" height='160' width="100%"></iframe></td></tr>
</table>


<script>

	function fn_Detail(arg0,arg1) {	

		historyDetail.location = "/menu03/submenu0311/History_Detail.asp?JUBSEQ="+arg0+"&Keyword="+arg1;	
	}
	function fn_Search() {
		document.inUpFrm.submit();
	}

	function fn_print(arg0) {	

		ShowPOPLayer("/menu03/submenu0311/lifecallmanage_print.asp?guboon=UP&jubseq="+arg0,'760','600');		
	}


	function fn_save(arg0) {	

		location.href="/menu03/submenu0311/lifecallmanage_save.asp?guboon=UP&jubseq="+arg0;
		
	}

	function fn_del(arg0,arg1) {	

		//alert("/menu03/submenu0301/lifecallhistory_InsUpDel.asp?guboon=DEL&jubseq="+arg0+"<%=where2%>");
		if ( confirm('선택한 자료를 삭제하시겠습니까?') )
		{
			location.href="/menu03/submenu0311/lifecallhistory_InsUpDel.asp?guboon=DEL&jubseq="+arg0+"&<%=where2%>";
		}
	}

	function fn_update(arg0,arg1) {	

			location.href="/menu03/submenu0312/lifecallmanage.asp?guboon=UP&jubseq="+arg0+"&<%=where2%>";
	}


	function fn_new() {	
		location.href="/menu03/submenu0312/lifecallmanage.asp?guboon=INS&<%=where2%>";
	}

	function fn_Xls() {
		location.href="Part_Xls.asp?<%=pageWHERE%>"
	}

	function fn_printA() {
		location.href="lifecallmanage_save_A.asp?<%=where2%>"
	}
	function pCateSelect(cn){

		Cate1 = 'A' ; //eval("inUpFrm.ACLASS"+cn).value;
		CUSTNO = '0000000000'; //parent.MemInfoFrame.frmSearch.CUSTNO.value;

		if (cn == '1')
		{//PSEQ1
			Relation = '0';//eval("inUpFrm.RELATION"+cn).value;
			RelationSeq = '0';//eval("inUpFrm.PSEQ"+cn).value;
			GoURL ="/Include/PopUp/PCategory.asp?Cate1=" +Cate1+ "&FM=" +cn+ "&CUSTNO=" +CUSTNO+"&Relation="+Relation+"&RelationSeq="+RelationSeq;
		}
		else
		{
			GoURL ="/Include/PopUp/PCategory.asp?Cate1=" +Cate1+ "&FM=" +cn+ "&CUSTNO=" +CUSTNO;
		}

		ShowPOPLayer(GoURL,'720','450');
	}
</script>



<!-- #include virtual="/Include/Bottom.asp" -->