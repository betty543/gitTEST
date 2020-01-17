<!-- #include virtual="/Include/Top_Frame.asp" -->
<%

SS_LoginID = SESSION("SS_LoginID")
SS_Login_Secgroup = SESSION("SS_Login_Secgroup")
	SS_Login_Grade = SESSION("SS_Login_Grade")


	'2. 쿼리조건절 셋팅
	pageSize = 10
	pageSector = 10
	if curPage = "" then curPage = 1 end If

	CUSTNO = Request("CUSTNO")

	sql_where =	"CUSTNO='" &CUSTNO&"'"
	if SS_Login_Grade <> "A" and SS_Login_Grade <> "C" then
		sql_where = sql_where & " and TELKIND = '" & SS_Login_Grade &"'"
	end if

	sql_tb = "TB_CALLHISTORY"

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

<table width="924" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
<table border="0" width="924" cellspacing="1" cellpadding="0" bgcolor="#CCCCCC" align="center">
	<tr height="25" bgcolor="#EEF6FF" align="center">
		<td class="TDCont" width="30" align='center'>순번</td>
		<td class="TDCont" align='center'><b>상담일시</b></td>
		<td class="TDCont" width="80" align='center'>전화구분</td>
		<td class="TDCont" width="80" align='center'><b>상담종류</b></td>
		<td class="TDCont" width="80" align='center'><b>상담방법</b></td>
		<td class="TDCont" width="80" align='center'><b>소속</b></td>
		<td class="TDCont" width="80" align='center'><b>계급</b></td>
		<td class="TDCont" width="80" align='center'><b>성명</b></td>
		<td class="TDCont" width="80" align='center'><b>상담관</b></td>
	</tr>

<%

	i = 0
	DO UNTIL RS.EOF
	i = i + 1


	db_JUBSEQ = rs("JUBSEQ")
	db_TELKIND = rs("TELKIND")
	db_JUBDATE = rs("JUBDATE")
	db_JUBTIME = rs("JUBTIME1")
	db_IOFLAG = rs("IOFLAG")
	db_CUSTNO = rs("CUSTNO")
	db_CUSTNAME = rs("CUSTNAME")
	db_TELNO = rs("TELNO")
	db_TELNO2 = rs("TELNO2")
	db_SEXGB = rs("SEXGB")
	db_CHANNELGB = rs("CHANNELGB")
	db_REQUESTERGB = rs("REQUESTERGB")
	db_CONSULTGB = rs("CONSULTGB")
	db_CONSULTETCGB = rs("CONSULTETCGB")
	db_SOSOKGB = rs("SOSOKGB")
	db_SOSOKETCGB = rs("SOSOKETCGB")
	db_LEVEL1 = rs("LEVEL1")
	db_LEVEL2 = rs("LEVEL2")
	db_ACLASS = rs("ACLASS") '소속
	db_BCLASS = rs("BCLASS") '소속세분류
	db_CCLASS = rs("CCLASS")
	db_CHANNEL = rs("CHANNEL")
	db_CALLFLAG = rs("CALLFLAG")
	db_CALLKIND = rs("CALLKIND")
	db_QUESTION = rs("QUESTION")
	db_REPLY = rs("REPLY")
	db_RESULTGB = rs("RESULTGB")
	db_RESERVEDATE = rs("RESERVEDATE")
	db_RESERVETIME = rs("RESERVETIME")
	db_PROCESSGB = rs("PROCESSGB")
	db_CALLID = rs("CALLID")
	db_RECORDFILE = rs("RECORDFILE")
	db_INCODE = rs("INCODE")


%>

		<tr bgcolor="#FFFFFF">

			<td align="center"><%=startRow%></td>

			<td align="center"><%=db_JUBSEQ%><br><%=db_JUBTIME%></td>
			<td align="center"><%=db_getCodeName("Z04",db_TELKIND)%></td>
			<td align="center"><%=db_getCodeName("C00",db_ACLASS)%></td>

			<td align="center"><%=db_getCodeName("C01",db_CHANNELGB)%></td> 



			<% if db_SOSOKETCGB <> "" then %>
				<td align="center"><%=db_getCodeName("C04",db_SOSOKGB)%><br><%=db_getCodeName("C41",db_SOSOKETCGB)%></td>
			<% else %>
				<td align="center"><%=db_getCodeName("C04",db_SOSOKGB)%></td>
			<% end if%>
			<% if db_LEVEL1 = "A" then %>
			<td align="center"><%=db_getCodeName("C05",db_LEVEL1)%><br><%=db_getCodeName("C06",db_LEVEL2)%></td>
			<% elseif db_LEVEL1 = "B" then %>
			<td align="center"><%=db_getCodeName("C05",db_LEVEL1)%><br><%=db_getCodeName("C07",db_LEVEL2)%></td>
			<% else %>
			<td align="center"><%=db_getCodeName("C05",db_LEVEL1)%></td>
			<% end if %>
			<td align="center"><%=db_CUSTNAME%></td>
			<td align="center"><%=db_getUserName(db_INCODE)%></td>

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
<table width="924" cellpadding="0" cellspacing="0" align="center">
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