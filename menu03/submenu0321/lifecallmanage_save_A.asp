<!-- #include virtual="/Include/Common.asp" -->

<%

	'response.write JUBSEQ
SS_LoginID = SESSION("SS_LoginID")
SS_Login_Secgroup = SESSION("SS_Login_Secgroup")

	QueryYN = request("QueryYN")
	FromDate = request("FromDate")
	ToDate = request("ToDate")
	curPage = 1

	whereCD1 = Trim(request("whereCD1")) '성별
	whereCD2 = Trim(request("whereCD2")) '상담방법
	whereCD3 = Trim(request("whereCD3")) '의뢰인
	whereCD4 = Trim(request("whereCD4")) '상담분야
	whereCD5 = Trim(request("whereCD5")) '소속
	whereCD6 = Trim(request("whereCD6")) '계급구분
	whereCD7 = Trim(request("whereCD7")) '계급구분2
	whereCD8 = Trim(request("whereCD8"))	'성명
	whereCD9 = Trim(request("whereCD9"))	'전화번호
	whereCD10 = Trim(request("whereCD10"))	'군분류
	whereCD11 = Trim(request("whereCD11"))	'처리결과
	whereCD12 = Trim(request("whereCD12"))	'처리결과
	whereCD13_A = Trim(request("whereCD13_A"))	'처리결과
	whereCD13_B = Trim(request("whereCD13_B"))	'처리결과
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
	CHANNELGB = Trim(request("CHANNELGB"))


	Server.ScriptTimeout = 90000
	Response.ContentType = "application/vnd.ms-excel; name='My_Excel'"
	Call Response.AddHeader("Content-Disposition", "attachment; filename=군범죄신고일지_" &FromDate&"_"&ToDate& ".xls")	'바로저장하기
	Call Response.AddHeader("Content-Description", "ASP Generated Data")


	'2. 쿼리조건절 셋팅

	where1 = "FromDate=" & FromDate & "&ToDate=" & ToDate & "&whereCD1=" & whereCD1 & "&whereCD2=" & whereCD2 & "&whereCD3=" & whereCD3 & "&whereCD4=" & whereCD4 & "&whereCD5=" & whereCD5 & "&whereCD6=" & whereCD6 & "&whereCD7=" & whereCD7 & "&whereCD8=" & whereCD8 & "&whereCD9=" & whereCD9 & "&whereCD10=" & whereCD10 & "&whereCD11=" & whereCD11 & "&whereCD12=" & whereCD12 & "&whereCD5_A=" & whereCD5_A& "&whereCD5_B=" & whereCD5_B& "&whereCD5_C=" & whereCD5_C& "&whereCD5_D=" & whereCD5_D& "&whereCD5_E=" & whereCD5_E& "&whereCD6_A=" & whereCD6_A& "&whereCD6_B=" & whereCD6_B& "&whereCD6_C=" & whereCD6_C & "&CHANNELGB=" & CHANNELGB
	where2 = "curPage=" & curPage & "&" & where1

	'SQL = "	SELECT *, CONVERT(VARCHAR(19),JUBTIME,121) JUBTIME1   FROM TB_LIFECALLHISTORY"
	sql_where =	"JUBDATE >= '" & FromDate & "'"
	sql_where = sql_where & "	AND     JUBDATE <= '" & ToDate & "'"

	IF whereCD1 <> "" THEN
		sql_where = sql_where & "	AND     SEXGB = '" & whereCD1 & "'"
	END IF
	IF whereCD2 <> "" THEN
		sql_where = sql_where & "	AND     CHANNELGB_B = '" & whereCD2 & "'"
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
		sql_where = sql_where & "	AND     SOSOKGB_A = '" & whereCD5_A & "'"
	END IF
	IF whereCD5_B <> "" THEN '소속
		sql_where = sql_where & "	AND     SOSOKGB_B = '" & whereCD5_B & "'"
	END IF
	IF whereCD5_C <> "" THEN '소속
		sql_where = sql_where & "	AND     SOSOKGB_C = '" & whereCD5_D & "'"
	END IF
	IF whereCD5_D <> "" THEN '소속
		sql_where = sql_where & "	AND     SOSOKGB_D = '" & whereCD5_E & "'"
	END IF
	IF whereCD5_E <> "" THEN '소속
		sql_where = sql_where & "	AND     SOSOKGB_E = '" & whereCD5_E & "'"
	END IF

	IF whereCD6 <> "" THEN
		'sql_where = sql_where & "	AND     LEVEL1 = '" & whereCD6 & "'"
	END IF
	IF whereCD7 <> "" THEN
		'sql_where = sql_where & "	AND     LEVEL2 = '" & whereCD7 & "'"
	END IF
	IF whereCD8 ="" and whereCD9 <> "" THEN
		sql_where = sql_where & "	AND     ( CUSTNAME LIKE '%" & whereCD9 & "%' or (TELNO LIKE '%" & whereCD9 & "%' OR TELNO2 LIKE '%" & whereCD9 & "%') or (Question LIKE '%" & whereCD9 & "%') or (REPLY LIKE '%" & whereCD9 & "%'))"
	END IF
	IF whereCD8 ="성명" and whereCD9 <> "" THEN
		sql_where = sql_where & "	AND     CUSTNAME LIKE '%" & whereCD9 & "%'"
	END IF
	IF whereCD8 ="전화번호" and whereCD9 <> "" THEN
		sql_where = sql_where & "	AND     (TELNO LIKE '%" & whereCD9 & "%' OR TELNO2 LIKE '%" & whereCD9 & "%')"
	END IF
	IF whereCD8 ="문의내용" and whereCD9 <> "" THEN
		sql_where = sql_where & "	AND     (Question LIKE '%" & whereCD9 & "%')"
	END IF
	IF whereCD8 ="조치내용" and whereCD9 <> "" THEN
		sql_where = sql_where & "	AND     (REPLY LIKE '%" & whereCD9 & "%')"
	END IF
	IF whereCD10 <> "" THEN
		sql_where = sql_where & "	AND     CHANNELGB = '" & whereCD10 & "'"
	END IF
	'if SS_Login_Secgroup = "A" then
		'내것만
	'	sql_where = sql_where& " AND	INCODE = '"&SS_LoginID&"'"
	'end if

	IF whereCD11 <> "" THEN
		sql_where = sql_where & "	AND     PROCESSGB = '" & whereCD11 & "'"
	END IF

	IF whereCD12 <> "" THEN
		'sql_where = sql_where & "	AND     EMERYN = '" & whereCD12 & "'"
	END IF

	IF whereCD13_A <> "" THEN
		sql_where = sql_where & "	AND     CALLCLASS_B = '" & whereCD13_A & "'"
	END If
	IF whereCD13_B <> "" THEN
		sql_where = sql_where & "	AND     CALLCLASS_C = '" & whereCD13_B & "'"
	END IF	
	
	if CHANNELGB <> "" then
		sql_where = sql_where & " and CHANNELGB = '" & CHANNELGB & "' "
	end if

	'Set Rs = server.createObject("ADODB.Recordset")
	'Rs.open SQL,db


	sql_tb = "TB_CRIMECALLHISTORY"
	'sql_index = "index_desc(" & sql_tb & " IDX_TB_CALLHISTORY_JUBSEQ)"
	sql_field ="*, CONVERT(VARCHAR(19),JUBTIME,121) JUBTIME1, convert(varchar(19),dateadd(second,calltime,JUBTIME),121) as JUBTIME2"
	sql_orderby = "JUBTIME asc"

	'3. 쿼리 실행
	sql = db_getSqlWithPage(sql_tb, sql_index, sql_field, sql_where, sql_orderby, 1000, 1)
	set Rs = db.execute(sql)
%>
<table width="600" height="10" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
<table width="600" cellspacing="0" align="center" border="1" bordercolor="black" bordercolordark="white" bordercolorlight="black">
	<tr bgcolor="#FFFFFF">
		<td>	

			<table width="600" height="80" border="1" cellspacing="1" cellpadding="1" align="center" bgcolor="#ffffff" bordercolor="black" bordercolordark="white" bordercolorlight="black">
			    <tr height="80">
					<td align='center' height="80" colspan='6'>
						<b><font color="#000000" size="5px" >군범죄신고 업무 보고</font></b>
					</td>
					<!--<td align="left" bgcolor="#FFFFFF" class="TDCont" colspan="8">&nbsp;<img src="/Images/dot_01.gif" >&nbsp;<b><font color="#0000ff" size=15px>상담일지</font></td>-->
				</tr>
			</table>
<%
	do until Rs.eof

		db_IDX	= RS("IDX")
		db_JUBSEQ	= RS("JUBSEQ")
		db_JUBDATE	= RS("JUBDATE")
		db_JUBTIME	= RS("JUBTIME")
		db_JUBTIME1 = RS("JUBTIME1")
		db_JUBTIME2 = RS("JUBTIME2")
		db_IOFLAG	= RS("IOFLAG")
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

	IF WEEKDAY(db_JUBTIME)=1 THEN
		JUBDAY="일"
	ELSEIF WEEKDAY(db_JUBTIME)=2 THEN
		JUBDAY="월"
	ELSEIF WEEKDAY(db_JUBTIME)=3 THEN
		JUBDAY="화"
	ELSEIF WEEKDAY(db_JUBTIME)=4 THEN
		JUBDAY="수"
	ELSEIF WEEKDAY(db_JUBTIME)=5 THEN
		JUBDAY="목"
	ELSEIF WEEKDAY(db_JUBTIME)=6 THEN
		JUBDAY="금"
	ELSEIF WEEKDAY(db_JUBTIME)=7 THEN
		JUBDAY="토"
	END IF

	sFirstLine = ""

	if db_getCateNameB_("S",db_CALLCLASS_B) <> "" then
		sFirstLine = db_getCateNameB_("S",db_CALLCLASS_B)
	end if

	if db_getCateNameC_("S",db_CALLCLASS_B,db_CALLCLASS_C) <> "" then
		sFirstLine = sFirstLine & ">" & db_getCateNameC_("S",db_CALLCLASS_B,db_CALLCLASS_C)
	end if


	if db_GetCodeName("C21",db_PROCESSGB) <> "" then
		if sFirstLine = "" then
			sFirstLine = db_GetCodeName("C21",db_PROCESSGB)
		else
			sFirstLine = sFirstLine & " - " & db_GetCodeName("C21",db_PROCESSGB)
		end if
	end if


		if sFirstLine = "" then
			sFirstLine = db_getUserName(db_INCODE)
		else
			sFirstLine = sFirstLine & " : " & db_getUserName(db_INCODE)
		end if



%>
<!--<body leftmargin="0" marginwidth="0" topmargin="0" marginheight="0" border="1">-->


			<table width="600" border="1" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC" bordercolor="black" bordercolordark="white" bordercolorlight="black">
			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="left" colspan='6'>[접수번호:<%=db_JUBSEQ%>]&nbsp;<%=left(db_JUBTIME1,10)%>&nbsp;<%=mid(db_JUBTIME1,12,5)%>&nbsp;~&nbsp;<%=mid(db_JUBTIME2,12,5)%>&nbsp;(<%=db_CALLTIMEDP%>)&nbsp;<%=db_getCodeName("A14",db_CHANNELGB)%>><%=db_getCateNameB_("Q",db_CHANNELGB_B)%><%if db_getCateNameC_("Q",db_CHANNELGB_B,db_CHANNELGB_C) <> "" then %>><%end if%><%=db_getCateNameC_("Q",db_CHANNELGB_B,db_CHANNELGB_C)%>&nbsp;[<%=sFirstLine%> 상담관]
					<br>
<%=db_getCateNameA_(db_SOSOKGB_A)%><%if db_getCateNameB_(db_SOSOKGB_A,db_SOSOKGB_B) <> "" then %>><%end if%><%=db_getCateNameB_(db_SOSOKGB_A,db_SOSOKGB_B)%><%if db_getCateNameC_(db_SOSOKGB_A,db_SOSOKGB_B,db_SOSOKGB_C) <> "" then %>><%end if%><%=db_getCateNameC_(db_SOSOKGB_A,db_SOSOKGB_B,db_SOSOKGB_C)%><%if db_getCateNameD_(db_SOSOKGB_A,db_SOSOKGB_B,db_SOSOKGB_C,db_SOSOKGB_D) <> "" then %>><%end if%><%=db_getCateNameD_(db_SOSOKGB_A,db_SOSOKGB_B,db_SOSOKGB_C,db_SOSOKGB_D)%><%if db_getCateNameE_(db_SOSOKGB_A,db_SOSOKGB_B,db_SOSOKGB_C,db_SOSOKGB_D,db_SOSOKGB_E) <> "" then %>><%end if%><%=db_getCateNameE_(db_SOSOKGB_A,db_SOSOKGB_B,db_SOSOKGB_C,db_SOSOKGB_D,db_SOSOKGB_E)%>)&nbsp;<%=db_getCateNameB_("P",db_LEVEL_B)%><%if db_getCateNameC_("P",db_LEVEL_B,db_LEVEL_C) <> "" then %>><%end if%> <%=db_getCateNameC_("P",db_LEVEL_B,db_LEVEL_C)%><%if db_getCateNameD_("P",db_LEVEL_B,db_LEVEL_C,db_LEVEL_D) <> "" then %>><%end if%><%=db_getCateNameD_("P",db_LEVEL_B,db_LEVEL_C,db_LEVEL_D)%><% if db_CUSTNAME <> "" then %>(<%=db_CUSTNAME%>)<%end if%>&nbsp;&nbsp;&nbsp;수신번호 : <%if db_CID = "" then%><%=db_TEL%><%else%><%=db_CID%><%end if%>
<br><br>
<b>* 상 담 내 용 : </b><%=db_QUESTION%>
<br>
<b>-><%=db_REPLY%></b>
<br><br>
<b>* 특 이 사 항 : </b>
<br>
<b>-><%=db_REMARK%></b>
				</tr>


			</table>

<%
			rs.movenext
		loop
%>

		</td>
	</tr>
</table>
