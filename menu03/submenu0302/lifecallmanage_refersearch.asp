<!-- #include virtual="/Include/Top_Popup.asp" -->
<!-- #include virtual="/Include/PopLayer.asp" -->
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

	if FromDate = "" then
		FromDate = date()
	end if
	if ToDate = "" then
		ToDate = date()
	end if

	'2. 쿼리조건절 셋팅
	pageSize = 3
	pageSector = 10
	if curPage = "" then curPage = 1 end If

	if QueryYN = "Y" then

		where1 = "QueryYN=Y&FromDate=" & FromDate & "&ToDate=" & ToDate & "&whereCD1=" & whereCD1 & "&whereCD2=" & whereCD2 & "&whereCD3=" & whereCD3 & "&whereCD4=" & whereCD4 & "&whereCD5=" & whereCD5 & "&whereCD6=" & whereCD6 & "&whereCD7=" & whereCD7 & "&whereCD8=" & whereCD8 & "&whereCD9=" & whereCD9 & "&whereCD10=" & whereCD10 & "&whereCD11=" & whereCD11 & "&whereCD12=" & whereCD12
		where2 = "curPage=" & curPage & "&" & where1

		'SQL = "	SELECT *, CONVERT(VARCHAR(19),JUBTIME,121) JUBTIME1   FROM TB_LIFECALLHISTORY"
		'sql_where =	"JUBDATE >= '" & FromDate & "'"
		'sql_where = sql_where & "	AND     JUBDATE <= '" & ToDate & "'"
		sql_where = "1=1"	

		IF whereCD8 <> "" THEN
			sql_where = sql_where & "	AND     CUSTNAME LIKE '%" & whereCD8 & "%'"
		END IF
		IF whereCD9 <> "" THEN
			sql_where = sql_where & "	AND     (TELNO LIKE '%" & whereCD9 & "%' OR TELNO2 LIKE '%" & whereCD9 & "%')"
		END IF
		IF whereCD10 <> "" THEN	'상담내용
			sql_where = sql_where & "	AND     ( QUESTION like '%" & whereCD10 & "%' or REPLY like '%" & whereCD10 & "%')"
		END IF
		'if SS_Login_Secgroup = "A" then
			'내것만
		'	sql_where = sql_where& " AND	INCODE = '"&SS_LoginID&"'"
		'end if




	'Set Rs = server.createObject("ADODB.Recordset")
	'Rs.open SQL,db


	sql_tb = "TB_LIFECALLHISTORY"
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

	end if
%>

<script>

	function fn_SetLevel2()
	{
		frame_level.location = "frame_level.asp?level="+document.all.whereCD6.value+"&level2=";
	}



</script>


<div id='CalendarLayer' style='display:none;'><iframe name='CalendarFrame' src='/include/Calendar.asp' width='172' height='177' border='0' frameborder='0' scrolling='no'></iframe></div>

<!-- #include virtual="/Include/PopLayer.asp" -->
<table border="0" width="940" cellpadding="0" cellspacing="3" bgcolor="#EFECE5" align="center">
		
	<form method="post" name="inUpFrm" style="margin:0" action="/menu03/submenu0302/lifecallmanage_refersearch.asp">
	<tr bgcolor="#FFFFFF">
		<td>

			<input type="hidden" name="QueryYN" value="">
			<input type="hidden" name="whereCD7" value="<%=whereCD7%>">
			<table width="100%" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">

			    <tr valign="center">
					<td bgcolor="#EFEFEF" class="TDCont" width=80 align='center'>성명</td>
					<td bgcolor="#FFFFFF">&nbsp;<input value="<%=whereCD8%>" name="whereCD8" type="text" size="20" onfocus="setFocusColor(this);" onblur="setOutColor(this);">					
					</td>
					<td bgcolor="#EFEFEF" class="TDCont" width=80 align='center'>전화번호</td>
					<td bgcolor="#FFFFFF" >&nbsp;<input value="<%=whereCD9%>" name="whereCD9" type="text" size="15" onfocus="setFocusColor(this);" onblur="setOutColor(this);">
					</td>
					<td bgcolor="#EFEFEF" class="TDCont" width=80 align='center'>상담내용</td>
					<td bgcolor="#FFFFFF" >&nbsp;<input value="<%=whereCD10%>" name="whereCD10" type="text" size="30" onfocus="setFocusColor(this);" onblur="setOutColor(this);">
					</td>

			        <td colspan='2' bgcolor="#FFFFFF" align="center">
			        	<img src="/Images/Btn/BtnSearch.gif" style="cursor:hand;" onClick="fn_Search();">
			        </td>
			    </tr>
			</table>
			</form>
		</td>
	</tr>
</table>

<table width="940" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
<table border="0" width="940" cellspacing="1" cellpadding="0" width="100%" bgcolor="#CCCCCC" align="center">
	<tr height="25" bgcolor="#EEF6FF" align="center">
		<td rowspan='2' class="TDCont" width="30" align='center'>순번</td>
		<td rowspan='2' class="TDCont" width="30" align='center'>구분</td>
		<td rowspan='2' width="80" class="TDCont" align='center'><b>상담일시</b></td>
		<td class="TDCont" width="120" align='center'><b>상담유형</b></td>
		<td class="TDCont" width="100" align='center'><b>상담분야</b></td>
		<td class="TDCont" width="100" align='center'><b>조치결과</b></td>
		<td class="TDCont" width="80" align='center'><b>소속</b></td>
		<td class="TDCont" width="80" align='center'><b>계급</b></td>
		<td class="TDCont" width="80" align='center'><b>성명</b></td>
		<td class="TDCont" width="50" align='center'><b>상담관</b></td>
		<td rowspan='2' class="TDCont" width="50" align='center'><b>상담차수</b></td>
	</tr>
	<tr height="25" bgcolor="#EEF6FF" align="center">
		<td class="TDCont" align='center' colspan='3'><b>상담내용</b></td>
		<td class="TDCont" align='center' colspan='4'><b>조치결과</b></td>
	</tr>

<%

	if QueryYN = "Y" then
		i = 0
		DO UNTIL RS.EOF
		i = i + 1


		db_IDX	= RS("IDX")
		db_JUBSEQ	= RS("JUBSEQ")
		db_JUBDATE	= RS("JUBDATE")
		db_JUBTIME	= RS("JUBTIME")
		db_JUBTIME1 = RS("JUBTIME1")
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


			if db_RECORDFILE = "" and db_CALLID <> "" then
				db_RECORDFILE = db_getRecFileName(db_CALLID,left(db_JUBTIME,10))


				IF db_RECORDFILE <> "" THEN
					RecFileName_URL = "<a href='##'>녹취</a>&nbsp;<img src='/Images/Comm/IconAlert.gif' align='absmiddle' style='cursor:hand;' onClick=parent.fn_Player('"&db_RECORDFILE&"'); title='녹음내용 청취'>&nbsp;</a>"		
				END IF

			end if

		sql = " select count(*) + 1 from tb_lifecallhistory where CUSTNO = '" & db_CUSTNO & "'"
		set Rs1 = db.execute(sql)
		
		db_TotalREFCNT = Rs1(0)

		if len(db_QUESTION)> 36 then
			db_QUESTION_con = CutString(db_QUESTION, 30, "...")
		else
			db_QUESTION_con = ""
		end if

		if len(db_REPLY)> 36 then
			db_REPLY_con = CutString(db_REPLY, 30, "...")
		else
			db_REPLY_con = ""
		end if


%>

		<tr bgcolor="#FFFFFF">

			<td align="center" rowspan='2'><%=startRow%></td>

			<td align="center" rowspan='2'><%=db_IOFLAG_NM%></td>
			<td align="center" rowspan='2'><a href="javascript:fn_Detail('<%=db_JUBSEQ%>');"><%=db_JUBTIME1%></a></td>
			<td align="center"><%=db_getCateNameB_("Q",db_CHANNELGB_B)%><%if db_getCateNameB_("Q",db_CHANNELGB_B) <> "" then %>><%end if%><%=db_getCateNameC_("Q",db_CHANNELGB_B,db_CHANNELGB_C)%></td>
			<td align="center"><%=db_getCateNameB_("O",db_CALLCLASS_B)%><%if db_getCateNameB_("O",db_CALLCLASS_B) <> "" then %>><%end if%><%=db_getCateNameC_("O",db_CALLCLASS_B,db_CALLCLASS_C)%></td>
			<td align="center"><%=db_getCodeName("C09",db_PROCESSGB)%></td>

			<td align="center"><%=db_getCateNameA_(db_SOSOKGB_A)%><%if db_getCateNameA_(db_SOSOKGB_A) <> "" then %>><%end if%><%=db_getCateNameB_(db_SOSOKGB_A,db_SOSOKGB_B)%></td>
			<td align="center"><%=db_getCateNameB_("P",db_LEVEL_B)%><%if db_getCateNameB_("P",db_LEVEL_B) <> "" then %>><%end if%> <%=db_getCateNameC_("P",db_LEVEL_B,db_LEVEL_C)%></td>
			<td align="center"><%=db_CUSTNAME%></td>
			<td align="center"><%=db_getUserName(db_INCODE)%></td>
			<td align="center" rowspan='1' >
					<%=db_REFCNT%>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td align="left" colspan='3' title="">&nbsp;<%=db_QUESTION_Con%></td>
			<td align="left" colspan='4' title="">&nbsp;<%=db_REPLY_Con%></td>

			<td align="center" >
				<% if db_REFCNT >= 1 then %>
				<img src="/Images/Btn/BtnSelectOK.gif" title="관련상담으로 선택" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_select('<%=db_JUBSEQ%>','<%=db_TotalREFCNT%>','<%=db_CUSTNAME%>','<%=db_SOSOKGB_A%>','<%=db_SOSOKGB_B%>','<%=db_SOSOKGB_C%>','<%=db_SOSOKGB_D%>','<%=db_SOSOKGB_E%>','<%=db_CUSTNO%>');">
				<% else %>
					&nbsp;
				<% end if%>
			</td>

		</tr>
<%
			startRow = startRow - 1
			RS.MOVENEXT
		LOOP

	end if
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
<table width="940" cellpadding="0" cellspacing="0" width="100%" align="center">
	<tr><td height="2" bgcolor="#f2f2f2"></td></tr>
	<tr><td height="5"></td></tr>
	<tr><td bgcolor="#D6D6D6" height="1"></td></tr>
	<tr><td bgcolor="#EEF6FF" class="TDL10px" height="25"><%=pageHtml%></td></tr>
	<tr><td bgcolor="#D6D6D6" height="1"></td></tr>
	<!--<tr>
		<td height="30" class="TDR10px">
			<img src="/Images/Btn/BtnAdd.gif" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_new();">
		</td>
	</tr>-->
</table>

<table width="940" cellpadding="0" cellspacing="0" width="100%" align="center" border="1">
<!--
	<tr><td bgcolor="#f2f2f2" width="50%" class="TDCont" align='center'>문의내용</td><td bgcolor="#f2f2f2" width="50%" class="TDCont" align='center'>조치내용</td></tr>
	<tr><td bgcolor="#f2f2f2" width="50%"><textarea name="QUESTION" style="width:100%; height:150" wrap="soft" class="TextareaInput"></textarea></td>
	<td  bgcolor="#f2f2f2" width="50%"><textarea name="REPLY" style="width:100%; height:150" wrap="soft" class="TextareaInput"></textarea></td></tr>
-->
	<tr><td bgcolor="#D6D6D6" colspan='2'><iframe src="History_Detail.asp" scrolling="auto" frameborder="0" border="0" name="historyDetail" height='220' width="100%"></iframe></td></tr>

</table>



<script>


	/*function fn_detail(arg1,arg2) {
		document.all.QUESTION.value = arg1;
		document.all.REPLY.value = arg2;

	}*/
	function fn_Detail(arg0) {	

		historyDetail.location = "/menu03/submenu0301/History_Detail.asp?JUBSEQ="+arg0+"&Keyword=";	
	}
	function fn_Search() {
		document.inUpFrm.QueryYN.value = "Y";
		document.inUpFrm.submit();
	}

	function fn_select(arg0,arg1,arg2,arg3,arg4,arg5,arg6,arg7,arg8) {	
		parent.document.all.REFCNT.value = arg1;
		parent.document.all.REFERJUBSEQ.value =arg0;	//자기자신이됨.

		if ( eval(" parent.document.all.OriginalCustnm") != null )
			parent.document.all.OriginalCustnm.value =arg2;	//자기자신이됨.

		parent.document.all.SOSOKGB_A.value = arg3;
		parent.document.all.SOSOKGB_B.value = arg4;
		parent.document.all.SOSOKGB_C.value = arg5;
		parent.document.all.SOSOKGB_D.value = arg6;
		parent.document.all.SOSOKGB_E.value = arg7;

		parent.frame_sosok_A.location.href = "/menu03/submenu0301/frame_CLASS.asp?frame_nm=frame_sosok_A&CLASSNM=SOSOK&CLASSGB=A&ACLASS="+arg3;
		parent.frame_sosok_B.location.href = "/menu03/submenu0301/frame_CLASS.asp?frame_nm=frame_sosok_B&CLASSNM=SOSOK&CLASSGB=B&ACLASS="+arg3+"&BCLASS="+arg4;
		parent.frame_sosok_C.location.href = "/menu03/submenu0301/frame_CLASS.asp?frame_nm=frame_sosok_C&CLASSNM=SOSOK&CLASSGB=C&ACLASS="+arg3+"&BCLASS="+arg4+"&CCLASS="+arg5;
		parent.frame_sosok_D.location.href = "/menu03/submenu0301/frame_CLASS.asp?frame_nm=frame_sosok_D&CLASSNM=SOSOK&CLASSGB=D&ACLASS="+arg3+"&BCLASS="+arg4+"&CCLASS="+arg5+"&DCLASS="+arg6;
		parent.frame_sosok_E.location.href = "/menu03/submenu0301/frame_CLASS.asp?frame_nm=frame_sosok_E&CLASSNM=SOSOK&CLASSGB=E&ACLASS="+arg3+"&BCLASS="+arg4+"&CCLASS="+arg5+"&DCLASS="+arg6;+"&ECLASS="+arg7;

		parent.document.getElementById('txtREFERJUBSEQ').innerHTML = "<a href='##'>"+arg0+"</a>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"+arg0+" 삭제' style='cursor:hand;' align='absmiddle' onClick=ReferDel('inUpFrm','"+arg0+"')>&nbsp;";

		parent.IframeHistory.location.href ="lifecallhistory_list.asp?CUSTNO="+arg8;



		parent.HddnPOPLayer();
	}

</script>


<!-- #include virtual="/Include/Bottom.asp" -->