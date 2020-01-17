<!-- #include virtual="/Include/Top_Frame.asp" -->
<%

	SS_LoginID = SESSION("SS_LoginID")
	SS_Login_Secgroup = SESSION("SS_Login_Secgroup")


	'2. 쿼리조건절 셋팅
	pageSize = 10
	pageSector = 10
	if curPage = "" then curPage = 1 end If

	CUSTNO = Request("CUSTNO")

	sql_where =	"CUSTNO='" &CUSTNO&"'"

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

<!-- #include virtual="/Include/PopLayer.asp" -->

<table width="1180" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
<table border="0" width="1180" cellspacing="1" cellpadding="0" bgcolor="#CCCCCC" align="center">
	<tr height="25" bgcolor="#EEF6FF" align="center">
		<td class="TDCont" width="30" align='center'>순번</td>
		<td class="TDCont" width="50" align='center'>구분</td>
		<td class="TDCont" align='center'><b>상담일시</b></td>
		<td class="TDCont" width="150" align='center'><b>상담유형</b></td>
		<td class="TDCont" width="150" align='center'><b>상담분야</b></td>
		<td class="TDCont" width="150" align='center'><b>조치결과</b></td>
		<td class="TDCont" width="100" align='center'><b>소속</b></td>
		<td class="TDCont" width="100" align='center'><b>계급</b></td>
		<td class="TDCont" width="80" align='center'><b>성명</b></td>
		<td class="TDCont" width="80" align='center'><b>상담관</b>
		<td class="TDCont" width="80" align='center'><b>녹취</b>		</td>
		<!--<td class="TDCont" width="50" align='center'><b>관리</b></td>-->
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


			if db_RECORDFILE = "" and db_CALLID <> "" then
				db_RECORDFILE = db_getRecFileName(db_CALLID,left(db_JUBTIME,10))


				IF db_RECORDFILE <> "" THEN
					RecFileName_URL = "<a href='##'>녹취</a>&nbsp;<img src='/Images/Comm/IconAlert.gif' align='absmiddle' style='cursor:hand;' onClick=parent.fn_Player('"&db_RECORDFILE&"'); title='녹음내용 청취'>&nbsp;</a>"		
				END IF

			end if


%>

		<tr bgcolor="#FFFFFF">

			<td align="center"><%=startRow%></td>

			<td align="center"><%=db_IOFLAG_NM%></td>
			<td align="center"><a href="javascript:fn_print('<%=db_JUBSEQ%>')"><%=db_JUBTIME1%></a></td>
			<td align="center"><%=db_getCateNameB_("Q",db_CHANNELGB_B)%><%if db_getCateNameB_("Q",db_CHANNELGB_B) <> "" then %>><%end if%><%=db_getCateNameC_("Q",db_CHANNELGB_B,db_CHANNELGB_C)%></td>
			<td align="center"><%=db_getCateNameB_("O",db_CALLCLASS_B)%><%if db_getCateNameB_("O",db_CALLCLASS_B) <> "" then %>><%end if%><%=db_getCateNameC_("O",db_CALLCLASS_B,db_CALLCLASS_C)%></td>
			<td align="center"><%=db_getCodeName("C09",db_PROCESSGB)%></td>

			<td align="center"><%=db_getCateNameA_(db_SOSOKGB_A)%><%if db_getCateNameB_(db_SOSOKGB_A,db_SOSOKGB_B) <> "" then %>><%end if%><%=db_getCateNameB_(db_SOSOKGB_A,db_SOSOKGB_B)%><%if db_getCateNameC_(db_SOSOKGB_A,db_SOSOKGB_B,db_SOSOKGB_C) <> "" then %>><%end if%><%=db_getCateNameC_(db_SOSOKGB_A,db_SOSOKGB_B,db_SOSOKGB_C)%><%if db_getCateNameD_(db_SOSOKGB_A,db_SOSOKGB_B,db_SOSOKGB_C,db_SOSOKGB_D) <> "" then %>><%end if%><%=db_getCateNameD_(db_SOSOKGB_A,db_SOSOKGB_B,db_SOSOKGB_C,db_SOSOKGB_D)%><%if db_getCateNameE_(db_SOSOKGB_A,db_SOSOKGB_B,db_SOSOKGB_C,db_SOSOKGB_D,db_SOSOKGB_E) <> "" then %>><%end if%><%=db_getCateNameE_(db_SOSOKGB_A,db_SOSOKGB_B,db_SOSOKGB_C,db_SOSOKGB_D,db_SOSOKGB_E)%></td>
			<td align="center"><%=db_getCateNameB_("P",db_LEVEL_B)%><%if db_getCateNameC_("P",db_LEVEL_B,db_LEVEL_C) <> "" then %>><%end if%> <%=db_getCateNameC_("P",db_LEVEL_B,db_LEVEL_C)%><%if db_getCateNameD_("P",db_LEVEL_B,db_LEVEL_C,db_LEVEL_D) <> "" then %>><%end if%><%=db_getCateNameD_("P",db_LEVEL_B,db_LEVEL_C,db_LEVEL_D)%></td>
			<td align="center"><%=db_CUSTNAME%></td>
			<td align="center"><%=db_getUserName(db_INCODE)%></td>
			<td align="center"><%=RecFileName_URL%></td>
			<!--<td align="center">
				<img src="/Images/Btn/BtnIconDel.gif" title="삭제" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_del('<%=db_JUBSEQ%>','DEL');">
			</td>-->
		</tr>
<%

		if db_JUBSEQ <> db_REFERJUBSEQ then
			'관련건 찾아오기


		end if
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
<table width="1180" cellpadding="0" cellspacing="0" align="center">
	<tr><td height="2" bgcolor="#f2f2f2"></td></tr>
	<tr><td height="5"></td></tr>
	<tr><td bgcolor="#D6D6D6" height="1"></td></tr>
	<tr><td bgcolor="#EEF6FF" class="TDL10px" height="25"><%=pageHtml%></td></tr>
	<tr><td bgcolor="#D6D6D6" height="1"></td></tr>
	<tr>
		<td height="30" class="TDR10px"></td>
	</tr>
</table>
<script>

	function fn_Search() {
		document.inUpFrm.submit();
	}

	function fn_print(arg0) {	

		parent.ShowPOPLayer("/menu03/submenu0301/lifecallmanage_print.asp?guboon=UP&jubseq="+arg0,'960','650');		
	}

	function fn_del(arg0,arg1) {	

		//alert("/menu03/submenu0301/lifecallhistory_InsUpDel.asp?guboon=DEL&jubseq="+arg0+"<%=where2%>");
		if ( confirm('선택한 자료를 삭제하시겠습니까?') )
		{
			location.href="/menu03/submenu0301/lifecallhistory_InsUpDel_frame.asp?guboon=DEL&jubseq="+arg0+"&<%=where2%>&CUSTNO=<%=CUSTNO%>";
		}
	}

	function fn_update(arg0,arg1) {	

		parent.ShowPOPLayer("/menu03/submenu0302/lifecallmanage_pop.asp?guboon=UP&jubseq="+arg0+"&<%=where2%>",'960','520');	
	}


	function fn_new() {	
		location.href="/menu03/submenu0302/lifecallmanage.asp?guboon=INS&<%=where2%>";
	}

	function fn_Xls() {
		location.href="Part_Xls.asp?<%=pageWHERE%>"
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