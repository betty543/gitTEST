<!-- #include virtual="/Include/Top.asp" -->
<!-- #include virtual="/Include/PopLayer.asp" -->
<%
	QueryYN = request("QueryYN")
	FromDate = request("FromDate")
	ToDate = request("ToDate")
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
		FromDate = left(date,7)&"-01"
	end if
	if ToDate = "" then
		ToDate = date
	end if

	pageSize = 10
	pageSector = 10
	if curPage = "" then curPage = 1 end If



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


<table border="0" width="940" cellpadding="0" cellspacing="3" bgcolor="#EFECE5" align="center">
		
	<form method="post" name="inUpFrm" style="margin:0">
	<tr bgcolor="#FFFFFF">
		<td>

			<input type="hidden" name="QueryYN" value="">
			<input type="hidden" name="whereCD7" value="<%=whereCD7%>">
			<table width="100%" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">

			    <tr>
			        <td width="80" bgcolor="#EFEFEF" class="TDCont" align='center'>조회기간</td>
			        <td colspan="3" bgcolor="#FFFFFF" >&nbsp;<input value="<%=FromDate%>" name="FromDate" readonly type="text" size="10" onfocus="setFocusColor(this);" onClick="window.open('/Include/Calendar_view.asp?goMonth='+document.inUpFrm.FromDate.value+'&firstcode=&FlightCode=&StartEndGu=S&ControlName=opener.document.inUpFrm.FromDate','','width=800,height=300,top=200,left=300,scrollbars=no,status=no');" onblur="setOutColor(this);">&nbsp;<img src="/Images/Btn/BtnIconDel.gif" title="삭제" style="cursor:hand;" align="absmiddle"onclick="document.all.FromDate.value='';">
				    	~
				    	<input value="<%=ToDate%>" name="ToDate" readonly type="text" size="10" onfocus="setFocusColor(this);" onClick="window.open('/Include/Calendar_view.asp?goMonth='+document.inUpFrm.ToDate.value+'&firstcode=&FlightCode=&StartEndGu=S&ControlName=opener.document.inUpFrm.ToDate','','width=800,height=300,top=200,left=300,scrollbars=no,status=no');" onblur="setOutColor(this);">&nbsp;<img src="/Images/Btn/BtnIconDel.gif" title="삭제" style="cursor:hand;" align="absmiddle"onclick="document.all.ToDate.value='';">
			        </td>
			        <td colspan='2' bgcolor="#FFFFFF" align="center">
			        	<img src="/Images/Btn/BtnSearch.gif" style="cursor:hand;" onClick="fn_Search();">
			        	<%IF SS_Login_Secgroup="A" Or SS_Login_Secgroup="B" THEN%><br><br><img src="/Images/Btn/BtnExcel.gif" style="cursor:hand;" onClick="fn_Xls();"><%END IF%>
			        </td>
				</tr>
			    </tr>
			</table>
			</form>
		</td>
	</tr>
</table>

<table width="940" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
<table border="0" width="940" cellspacing="1" cellpadding="0" width="100%" bgcolor="#CCCCCC" align="center">
	<tr height="25" bgcolor="#EEF6FF" align="center">
		<td class="TDCont" align='center'>연번</td>
		<td class="TDCont" align='center'><b>상담일시</b></td>
		<td class="TDCont" align='center'><b>성별</b></td>
		<td class="TDCont" align='center'><b>상담방법</b></td>
		<td class="TDCont" align='center'><b>성명</b></td>
		<td class="TDCont" align='center'><b>소속</b></td>
		<td class="TDCont" align='center'><b>계급</b></td>
		<td class="TDCont" align='center'><b>상담분야</b></td>
		<td class="TDCont" align='center'><b>의뢰인</b></td>
		<td class="TDCont" align='center'><b>인지경로</b></td>
		<td class="TDCont" align='center'><b>가해자</b></td>
		<td class="TDCont" align='center'><b>상담관</b></td>
		<td class="TDCont" align='center'><b>조치결과</b></td>
	</tr>

<%


	if QueryYN = "Y" then
		SQL = "	SELECT *, CONVERT(VARCHAR(19),JUBTIME,121) JUBTIME1   FROM TB_LIFECALLHISTORY"
		SQL = SQL & "		WHERE	JUBDATE >= '" & FromDate & "'"
		SQL = SQL & "		AND     JUBDATE <= '" & ToDate & "'"
		SQL = SQL & "	ORDER BY JUBTIME"

		Set Rs = server.createObject("ADODB.Recordset")
		Rs.open SQL,db

		i = 0
		DO UNTIL RS.EOF
		i = i + 1


		db_JUBSEQ = rs("JUBSEQ")
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
		db_EMERYN = rs("EMERYN")

%>

		<tr bgcolor="#FFFFFF">

			<td align="center"><%=i%></td>
			<td align="center"><%=db_JUBTIME%><% if db_EMERYN="Y" then%><font color="#0000ff"> (긴급)</font><%else%> (일반)<%end if%></td>
			<td align="center"><%if db_SEXGB = "1" then%>남<%else%>녀<%end if%></td>
			<td align="center"><%=db_getCodeName("C00",db_ACLASS)%><br><%=db_getCodeName("C01",db_CHANNELGB)%></td>


			<td align="center"><%=db_CUSTNAME%></td>
			<td align="center"><%=db_getCateNameA_(db_SOSOKGB)%><br><%=db_getCateNameB_(db_SOSOKGB,db_SOSOKETCGB)%></td>
			<% if db_LEVEL1 = "A" then %>
			<td align="center"><%=db_getCodeName("C05",db_LEVEL1)%><br><%=db_getCodeName("C06",db_LEVEL2)%></td>
			<% elseif db_LEVEL1 = "B" then %>
			<td align="center"><%=db_getCodeName("C05",db_LEVEL1)%><br><%=db_getCodeName("C07",db_LEVEL2)%></td>
			<% else %>
			<td align="center"><%=db_getCodeName("C05",db_LEVEL1)%></td>
			<% end if %>

			<% if db_CONSULTETCGB <> "" then %>
				<td align="center"><%=db_getCodeName("C03",db_CONSULTGB)%><br><%=db_getCodeName("C31",db_CONSULTETCGB)%></td> 
			<% else %>
				<td align="center"><%=db_getCodeName("C03",db_CONSULTGB)%></td> 
			<% end if %>
			<td align="center"><%=db_getCodeName("C02",db_REQUESTERGB)%></td>
			<td align="center"><%=db_getCodeName("C10",db_CALLFLAG)%></td>
			<td align="center"><%=db_getCodeName("C08",db_CALLKIND)%></td>
			<td align="center"><%=db_getUserName(db_INCODE)%></td>
			<td align="center"><%=db_getCodeName("C09",db_PROCESSGB)%></td>
		</tr>
<%
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



<!-- #include virtual="/Include/Bottom.asp" -->


<script>

	function fn_Search() {
		document.all.QueryYN.value = "Y";
		document.inUpFrm.submit();
	}

	function fn_del(arg0,arg1) {	
		if ( confirm('선택한 자료를 삭제하시겠습니까?') )
		{
			location.href="/menu03/submenu0301/lifecallhistory_InsUpDel.asp?guboon=DEL&jubseq="+arg0;
		}
	}

	function fn_update(arg0,arg1) {	

			location.href="/menu03/submenu0302/lifecallmanage.asp?guboon=UP&jubseq="+arg0;
	}


	function fn_new() {	
		location.href="/menu03/submenu0302/lifecallmanage.asp?guboon=INS";
	}

	function fn_Xls() {
		location.href="list02_Xls.asp?FromDate=<%=FromDate%>&ToDate=<%=ToDate%>"
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