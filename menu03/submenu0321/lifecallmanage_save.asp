<!-- #include virtual="/Include/Common.asp" -->

<%
	guboon = request("guboon")
	JUBSEQ = request("JUBSEQ")
	InType = request("InType")

	Server.ScriptTimeout = 90000
	Response.ContentType = "application/vnd.ms-excel; name='My_Excel'"
	Call Response.AddHeader("Content-Disposition", "attachment; filename=상담일지" &JUBSEQ& ".xls")	'바로저장하기
	Call Response.AddHeader("Content-Description", "ASP Generated Data")

%>
<%

	'response.write JUBSEQ

	SQL = "	SELECT *, CONVERT(CHAR(19),JUBTIME,121) AS JUBTIME1 FROM TB_CRIMECALLHISTORY"
	SQL = SQL & "		WHERE	JUBSEQ = '" & JUBSEQ & "'"

	Set Rs = server.createObject("ADODB.Recordset")
	Rs.open SQL,db

	if rs.eof = false then

		db_IDX	= RS("IDX")
		db_JUBSEQ	= RS("JUBSEQ")
		db_JUBDATE	= RS("JUBDATE")
		db_JUBTIME	= RS("JUBTIME")
		db_JUBTIME1 = RS("JUBTIME1")
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

	end if
%>
<!--<body leftmargin="0" marginwidth="0" topmargin="0" marginheight="0" border="1">-->

<table width="600" height="10" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
<table width="600" cellspacing="0" align="center" border="1" bordercolor="black" bordercolordark="white" bordercolorlight="black">
	<tr bgcolor="#FFFFFF">
		<td>	

			<table width="600" height="80" border="1" cellspacing="1" cellpadding="1" align="center" bgcolor="#ffffff" bordercolor="black" bordercolordark="white" bordercolorlight="black">
			    <tr height="80">
					<td align='center' height="80" colspan='6'>
						<b><font color="#000000" size="5px" >군범죄신고일지</font></b>
					</td>
					<!--<td align="left" bgcolor="#FFFFFF" class="TDCont" colspan="8">&nbsp;<img src="/Images/dot_01.gif" >&nbsp;<b><font color="#0000ff" size=15px>상담일지</font></td>-->
				</tr>
			</table>

			<table width="600" border="1" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC" bordercolor="black" bordercolordark="white" bordercolorlight="black">
			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center">연번</td>
					<td bgcolor="#FFFFFF" width=100 nowrap>&nbsp;<b><%=db_JUBSEQ%></b>
					</td>
					<td bgcolor="#FFFFFF" width=100><%if db_EMERYN = "Y" then%><font color="#0000ff">&nbsp;긴급</font><%else%>&nbsp;<%end if%></td>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center">상담일시</td>
					<td bgcolor="#FFFFFF" colspan='2' >&nbsp;<b><%=db_JUBTIME1%>(<%=JUBDAY%>)</b>
					</td>
				</tr>

			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center">성    별</td>
					<td bgcolor="#FFFFFF">&nbsp;<b><% if db_SEXGB = "1" or SEXGB = "" then %>남<% else %>녀<%end if%></b>
					</td>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center">성    명</td>
					<td bgcolor="#FFFFFF" >&nbsp;<b><%=db_CUSTNAME%></b>
					</td>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center">통화시간</td>
					<td bgcolor="#FFFFFF">&nbsp;<b><%=db_CALLTIMEDP%></b>
					</td>

				</tr>
				<tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center">연락처1</td>
					<td bgcolor="#FFFFFF"  >&nbsp;<b><%=db_TELNO%></b></td>

					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center">연락처2</td>
					<td bgcolor="#FFFFFF"  >&nbsp;<b><%=db_TELNO2%></b></td>

					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center">발신번호</td>
					<td bgcolor="#FFFFFF" width=100  >&nbsp;<b><%=db_CID%></b></td>
				</tr>

			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center">소    속</td>
					<td bgcolor="#FFFFFF" colspan='5' nowrap>&nbsp;<b><%=db_getCateNameA_(db_SOSOKGB_A)%><%if db_getCateNameB_(db_SOSOKGB_A,db_SOSOKGB_B) <> "" then %>><%end if%><%=db_getCateNameB_(db_SOSOKGB_A,db_SOSOKGB_B)%><%if db_getCateNameC_(db_SOSOKGB_A,db_SOSOKGB_B,db_SOSOKGB_C) <> "" then %>><%end if%><%=db_getCateNameC_(db_SOSOKGB_A,db_SOSOKGB_B,db_SOSOKGB_C)%><%if db_getCateNameD_(db_SOSOKGB_A,db_SOSOKGB_B,db_SOSOKGB_C,db_SOSOKGB_D) <> "" then %>><%end if%><%=db_getCateNameD_(db_SOSOKGB_A,db_SOSOKGB_B,db_SOSOKGB_C,db_SOSOKGB_D)%><%if db_getCateNameE_(db_SOSOKGB_A,db_SOSOKGB_B,db_SOSOKGB_C,db_SOSOKGB_D,db_SOSOKGB_E) <> "" then %>><%end if%><%=db_getCateNameE_(db_SOSOKGB_A,db_SOSOKGB_B,db_SOSOKGB_C,db_SOSOKGB_D,db_SOSOKGB_E)%></td>
				</tr>
				<tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center">계    급</td>
					<td bgcolor="#FFFFFF" colspan='5' nowrap>&nbsp;<b><%=db_getCateNameB_("P",db_LEVEL_B)%><%if db_getCateNameC_("P",db_LEVEL_B,db_LEVEL_C) <> "" then %>><%end if%> <%=db_getCateNameC_("P",db_LEVEL_B,db_LEVEL_C)%><%if db_getCateNameD_("P",db_LEVEL_B,db_LEVEL_C,db_LEVEL_D) <> "" then %>><%end if%><%=db_getCateNameD_("P",db_LEVEL_B,db_LEVEL_C,db_LEVEL_D)%></b>
					</td>

				</tr>
				<tr>

					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center">상담유형</td>
					<td bgcolor="#FFFFFF" colspan='3'>&nbsp;<b><%=db_getCodeName("A14",db_CHANNELGB)%>><%=db_getCateNameB_("Q",db_CHANNELGB_B)%><%if db_getCateNameB_("Q",db_CHANNELGB_B) <> "" then %>><%end if%><%=db_getCateNameC_("Q",db_CHANNELGB_B,db_CHANNELGB_C)%>&nbsp;&nbsp;&nbsp;&nbsp;<%if EMERYN="Y" then%><font color="#ff0000"><b>긴급</b></font><%end if%>
					</td>
					
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center">관련상담번호</td>
					<td bgcolor="#FFFFFF" >&nbsp;<b><%=db_REFERJUBSEQ%></b>
					</td>		

				</tr>
				<tr>


					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center">상담분야</td>
					<td bgcolor="#FFFFFF" nowrap valign="top" colspan='3'>&nbsp;<b><%=db_getCateNameB_("S",db_CALLCLASS_B)%><%if db_getCateNameB_("S",db_CALLCLASS_B) <> "" then %>><%end if%><%=db_getCateNameC_("S",db_CALLCLASS_B,db_CALLCLASS_C)%></b>
					</td>

					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center">상담회차</td>
					<td bgcolor="#FFFFFF"  >&nbsp;<b><%=db_REFCNT%>&nbsp;회</b>
					</td>		

				</tr>
			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center">의 뢰 인</td>
					<td bgcolor="#FFFFFF" >&nbsp;<b><%=db_GetCodeName("C02",db_REQUESTERGB)%></b>
					</td>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center" >원인제공자</td>
					<td bgcolor="#FFFFFF" >&nbsp;<b><%=db_getCateNameB_("T",db_CALLKIND_B)%></b>
					</td>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center">조치결과</td>
					<td bgcolor="#FFFFFF" >&nbsp;<b><%=db_GetCodeName("C21",db_PROCESSGB)%></b>
					</td>					

				</tr>
			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center">상담내용</td>
					<td bgcolor="#FFFFFF" colspan=5 width=500>&nbsp;<b><%=db_QUESTION%></b>		
					</td>
				</tr>
			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center">조치내용</td>
					<td bgcolor="#FFFFFF" colspan=5 width=500>&nbsp;<b><%=db_REPLY%></b>		
					</td>
				</tr>
			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center">특이사항</td>
					<td bgcolor="#FFFFFF" colspan=5 width=500>&nbsp;<b><%=db_REMARK%></b>	
					</td>
				</tr>
			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center">상담관</td>
					<td bgcolor="#FFFFFF" colspan=5 width=500>&nbsp;<b><%=db_getUserName(db_INCODE)%></b>&nbsp;&nbsp;(인)	
					</td>
				</tr>

			</table>

		</td>
	</tr>
</table>
