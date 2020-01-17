<!-- #include virtual="/Include/Top.asp" -->
<!-- #include virtual="/Include/PopLayer.asp" -->
<%
	QueryYN = request("QueryYN")
	FromDate = request("FromDate")
	ToDate = request("ToDate")
	whereCD1 = Trim(request("whereCD1"))	'성별
	whereCD2 = Trim(request("whereCD2"))	'전화구분
	whereCD3 = Trim(request("whereCD3"))	'소속
	whereCD4 = Trim(request("whereCD4"))	'계급
	whereCD5 = Trim(request("whereCD5"))	'성명


	SS_Login_Grade = SESSION("SS_Login_Grade")

	If QueryYN = "" Then
		whereCD1 = ""
	End if
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
	where2 = "curPage=" & curPage & "&" & where1


	sql_where =	"JUBDATE >= '" & FromDate & "'"
	sql_where = sql_where & "	AND     JUBDATE <= '" & ToDate & "'"

	IF whereCD1 <> "" THEN
		sql_where = sql_where & "	AND     SEXGB = '" & whereCD1 & "'"
	END IF
	IF whereCD2 <> "" THEN
		sql_where = sql_where & "	AND     CHANNELGB = '" & whereCD2 & "'"
	END IF
	IF whereCD3 <> "" THEN
		sql_where = sql_where & "	AND     SOSOKGB = '" & whereCD5 & "'"
	END IF
	IF whereCD4 <> "" THEN
		sql_where = sql_where & "	AND     LEVEL1 = '" & whereCD6 & "'"
	END IF
	IF whereCD5 <> "" THEN
		sql_where = sql_where & "	AND     CUSTNAME LIKE '%" & whereCD8 & "%'"
	END IF

	if SS_Login_Grade <> "A" and SS_Login_Grade <> "C" then
		sql_where = sql_where & " and TELKIND = '" & SS_Login_Grade &"'"
	end if
	'Set Rs = server.createObject("ADODB.Recordset")
	'Rs.open SQL,db


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
<div id='CalendarLayer' style='display:none;'><iframe name='CalendarFrame' src='/include/Calendar.asp' width='172' height='177' border='0' frameborder='0' scrolling='no'></iframe></div>


<table border="0" width="940" cellpadding="0" cellspacing="3" bgcolor="#EFECE5" align="center">
	<tr bgcolor="#FFFFFF">
		<td>
		
			<form method="post" name="inUpFrm" style="margin:0">
			<input type="hidden" name="QueryYN" value="">
			<table width="100%" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">

			    <tr>
			        <td width="80" bgcolor="#EFEFEF" class="TDCont" align="center">조회기간</td>
			        <td  bgcolor="#FFFFFF" colspan=3 width=250 nowrap>&nbsp;
			        	<input value="<%=FromDate%>" name="FromDate" readonly type="text" size="10" onfocus="setFocusColor(this);" onClick="new CalendarFrame.Calendar(this);" onblur="setOutColor(this);">
				    	~
				    	<input value="<%=ToDate%>" name="ToDate" readonly type="text" size="10" onfocus="setFocusColor(this);" onClick="new CalendarFrame.Calendar(this);" onblur="setOutColor(this);">	
			        </td>

					<td bgcolor="#EFEFEF" class="TDCont" width=80 align="center">성별</td>
					<td bgcolor="#FFFFFF" >&nbsp;
						<input type="radio" name="whereCD1" value="" class="none" <% if whereCD1 ="" then%> checked <%end if%>> 전체
						<input type="radio" name="whereCD1" value="1" class="none" <% if whereCD1 ="1" then%> checked <%end if%> > 남
						<input type="radio" name="whereCD1" value="2" class="none" <% if whereCD1 ="2" then%> checked <%end if%>> 녀
					</td>

					<td bgcolor="#EFEFEF" class="TDCont" width=80 align="center">전화구분</td>
					<td bgcolor="#FFFFFF" width=200 colspan=1>&nbsp;<%
							'======= 처리구분 코드 가져오기 ==================================================
							SqlCode = "SELECT CODE, CODENAME FROM TB_CODE"
							SqlCode = SqlCode& " WHERE USEYN='Y' AND SYSYN='Y' AND CODEGROUP='Z04'"
	if SS_Login_Grade <> "A" and SS_Login_Grade <> "C" then
		SqlCode = SqlCode & " and CODE = '" & SS_Login_Grade &"'"
	end if
							SqlCode = SqlCode& " ORDER BY CODE"
							set RsCode = db.execute(SqlCode)
						%>
						<select name="whereCD2" size="1" class="ComboFFFCE7">
						<option value="">선택</option>
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
							%>	</select>				
					</td>

			        <td colspan='2' rowspan="2" bgcolor="#FFFFFF" align="center" width="150">
			        	<img src="/Images/Btn/BtnSearch.gif" style="cursor:hand;" onClick="fn_Search();">
			        	<%IF SS_Login_Secgroup="A" Or SS_Login_Secgroup="B" THEN%><br><br><img src="/Images/Btn/BtnExcel.gif" style="cursor:hand;" onClick="fn_Xls();"><%END IF%>
			        </td>
				</tr>
			    <tr>

					<td bgcolor="#EFEFEF" class="TDCont" width=80 align="center">소속</td>
					<td bgcolor="#FFFFFF" nowrap colspan='3'>&nbsp;<%
							'======= 처리구분 코드 가져오기 ==================================================
							SqlCode = "SELECT CODE, CODENAME FROM TB_CODE"
							SqlCode = SqlCode& " WHERE USEYN='Y' AND SYSYN='Y' AND CODEGROUP='C04'"
							SqlCode = SqlCode& " ORDER BY CODE"
							set RsCode = db.execute(SqlCode)
						%>
						<select name="whereCD3" size="1" class="ComboFFFCE7">
						<option value="">선택</option>
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("CODE")
										CODENAME = RsCode("CODENAME")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &whereCD3& "")%>
							<%
									RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
							%>	</select>
					</td>
					<td bgcolor="#EFEFEF" class="TDCont" align='center'>계급</td>
					<td bgcolor="#FFFFFF" width=200>&nbsp;
<%
							'======= 처리구분 코드 가져오기 ==================================================
							SqlCode = "SELECT CODE, CODENAME FROM TB_CODE"
							SqlCode = SqlCode& " WHERE USEYN='Y' AND SYSYN='Y' AND CODEGROUP='C05'"
							SqlCode = SqlCode& " ORDER BY CODE"
							set RsCode = db.execute(SqlCode)
						%>
						<select name="whereCD4" size="1" class="ComboFFFCE7">
							<Option value ='' selected>계급구분</option>
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("CODE")
										CODENAME = RsCode("CODENAME")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &whereCD4& "")%>
							<%
									RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
							%>
						</select>			
					</td>
			        <td width="80" bgcolor="#EFEFEF" class="TDCont" align='center'>성명</td>
			        <td bgcolor="#FFFFFF">&nbsp;
			        	<input value="<%=whereCD5%>" name="whereCD5" type="text" size="20" onfocus="setFocusColor(this);" onblur="setOutColor(this);" onKeypress="if (event.keyCode==13) {fn_Search();}"></td>
			    </tr>
			</table>
			</form>
		</td>
	</tr>
</table>


<table width="940" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
<table border="0" width="940" cellspacing="1" cellpadding="0" width="100%" bgcolor="#CCCCCC" align="center">
	<tr height="25" bgcolor="#EEF6FF" align="center">
		<td><b>순번</b></td>
		<td><b>전화구분</b></td>
		<td><b>상담일시</b></td>
		<td><b>상담종류</b></td>
		<td><b>상담방법</b></td>
		<td><b>소속</b></td>
		<td><b>계급</b></td>
		<td><b>성명</b></td>
		<td><b>상담관</b></td>
		<td><b>관리</b></td>
	</tr>


<%

	i = 0
	DO UNTIL RS.EOF
	i = i + 1


	db_JUBSEQ = rs("JUBSEQ")
	
	db_DNIS = rs("DNIS")
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
	db_TELKIND = rs("TELKIND")


%>

		<tr bgcolor="#FFFFFF">

			<td align="center"><%=startRow%></td>
			<td align="center"><%=db_getCodeName("Z04",db_TELKIND)%></td>
			<td align="center"><%=db_JUBTIME%></td>
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
			<td align="center">
				<img src="/Images/Btn/BtnIconModify.gif" title="수정" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_update('<%=db_JUBSEQ%>','UP');">
				<img src="/Images/Btn/BtnIconDel.gif" title="삭제" style="cursor:hand;" align="absmiddle"onclick="javascript:fn_del('<%=db_JUBSEQ%>','DEL');">
			</td>
		</tr>
<%
		startRow = startRow - 1
		RS.MOVENEXT
	LOOP


%>
</table>
<table width="940" cellpadding="0" cellspacing="0" width="100%" align="center">
	<tr><td height="2" bgcolor="#f2f2f2"></td></tr>
	<tr><td height="5"></td></tr>
	<tr><td bgcolor="#D6D6D6" height="1"></td></tr>
	<tr><td bgcolor="#EEF6FF" class="TDL10px" height="25">1 </td></tr>
	<tr><td bgcolor="#D6D6D6" height="1"></td></tr>
	<tr>
		<td height="30" class="TDR10px">
			<img src="/Images/Btn/BtnAdd.gif" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_insert();">
		</td>
	</tr>
</table>


<script>

	function fn_Search() {

		//document.inUpFrm.QueryYN.value = "Y";
		document.inUpFrm.submit();
	}

	function fn_insert() {	
		location.href="/menu04/submenu0402/callmanage.asp";
	}

	function fn_del(arg0,arg1) {	

		//alert("/menu03/submenu0301/lifecallhistory_InsUpDel.asp?guboon=DEL&jubseq="+arg0+"<%=where2%>");
		if ( confirm('선택한 자료를 삭제하시겠습니까?') )
		{
			location.href="/menu04/submenu0402/callhistory_InsUpDel.asp?guboon=DEL&jubseq="+arg0+"&<%=where2%>";
		}
	}

	function fn_update(arg0,arg1) {	

			location.href="/menu04/submenu0402/callmanage.asp?guboon=UP&jubseq="+arg0+"&<%=where2%>";
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