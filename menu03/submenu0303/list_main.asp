<!-- #include virtual="/Include/Top.asp" -->
<%
	'####### 파라미터 ##################################################################################
	QueryYN = request("QueryYN")
	FromDate = request("FromDate")
	ToDate = request("ToDate")
	whereCD3 = Trim(request("whereCD3"))
	whereCD7 = Trim(request("whereCD7"))

	If QueryYN = "" Then
		whereCD3 = "1"
	End if

	Dim bound_code()

	if FromDate = "" then FromDate =Date() end If
	if ToDate = "" then ToDate=date() end If

	pageWHERE = "QueryYN="&QueryYN&"&FromDate="&FromDate&"&ToDate="&ToDate&"&whereCD3="&whereCD3&"&whereCD7="&whereCD7

%>

<div id='CalendarLayer' style='display:none;'><iframe name='CalendarFrame' src='/include/Calendar.asp' width='172' height='177' border='0' frameborder='0' scrolling='no'></iframe></div>

<script>

	function fn_Search() {

		if (inUpFrm.whereCD3(1).checked && inUpFrm.whereCD7.value =='' )
		{
			alert('제품유형별 검색을 하실때는 1차분류를 꼭 선택하셔야 합니다!');
			return;
		}
		document.inUpFrm.QueryYN.value = "Y";
		document.inUpFrm.submit();
	}
	
	function fn_Xls() {
		location.href="Part_Xls.asp?<%=pageWHERE%>"
	}

	function fn_whereCD3(){
		if (inUpFrm.whereCD3(0).checked){ 
			inUpFrm.whereCD7.disabled = true;
			//CallFlagFrame1.location.href="/Manage/AsRegi/AsRegiForm_CallFlag.asp?FM=1&Cate1=&Cate2=&Cate3=";
		}
		
		if (inUpFrm.whereCD3(1).checked){
			inUpFrm.whereCD7.disabled = false;
			//CallFlagFrame1.location.href="/Manage/AsRegi/AsRegiForm_CallFlag.asp?FM=1&Cate1=A";
		}

		if (inUpFrm.whereCD3(2).checked){
			inUpFrm.whereCD7.disabled = true;
			//CallFlagFrame1.location.href="/Manage/AsRegi/AsRegiForm_CallFlag.asp?FM=1&Cate1=A";
		}
	}

</script>
<table border="0" width="1200" cellpadding="0" cellspacing="3" bgcolor="#EFECE5" align="center">
	<tr bgcolor="#FFFFFF">
		<td>
			<form name="inUpFrm" method="post" action="<%=Menu_2nd%>" onsubmit="return fn_Search(this);" style="margin:0">
			<input type="hidden" name="QueryYN" value="<%=QueryYN%>">
			<table width="100%" border="0" cellspacing="1" cellpadding="0" style="border:#E1DED6 solid 1px">
			    <tr>
			        <td class="TDCont">조회기간 :
			        	<input value="<%=FromDate%>" name="FromDate" readonly type="text" size="10" onfocus="setFocusColor(this); new CalendarFrame.Calendar(this);" onblur="setOutColor(this);">
				    	~
				    	<input value="<%=ToDate%>" name="ToDate" readonly type="text" size="10" onfocus="setFocusColor(this); new CalendarFrame.Calendar(this);" onblur="setOutColor(this);">
			        </td>
			        <td class="TDCont">
						<input type="radio" name="whereCD3" value="1" <%If whereCD3 = "1" Then response.write "checked" End If %> class="none" onClick="fn_whereCD3();"> 제품별 상담구분별
						<input type="radio" name="whereCD3" value="2" <%If whereCD3 = "2" Then response.write "checked" End If %> class="none" onClick="fn_whereCD3();"> 제품별 상담유형별
						<input type="radio" name="whereCD3" value="3" <%If whereCD3 = "3" Then response.write "checked" End If %> class="none" onClick="fn_whereCD3();"> 제품별 매체별
			        </td>

			        <td class="TDCont"> 제품분류 :</td>
			        <td class="TDCont">
			        	<select name="whereCD7" <%If whereCD3 = "1" Or whereCD3 = "3" Then %>Disabled<% End if %> size="1" class="ComboFFFCE7" style="width:120">
							<option value="">선택</option>
							<%=db_getTBCodeSelect("A01", whereCD7, "N")%>
						</select>
			        </td>

			        <td class="TDR5px">
			        	<img src="/Images/Btn/BtnSearch.gif" align="absmiddle" style="cursor:hand;" onClick="fn_Search();">
			        	<img src="/Images/Btn/BtnExcel.gif" align="absmiddle" style="cursor:hand;" onClick="fn_Xls();">
			        </td>
			    </tr>
			</table>
			</form>
		</td>
	</tr>
</table>
<table border="0" width="100%" cellpadding="0" cellspacing="0" align="center"><tr height="5"><td></td></tr></table>
<%

	If QueryYN = "Y" Then

		If whereCD3 = "2" then
	
			'----------------------------------------
			'1) 제품분류 2차, 상담유형
			'----------------------------------------

			SQL = "	SELECT	COUNT(0) AS CNT	FROM	TB_CODE	WHERE CODEGROUP = 'A04'"
			SQL = SQL & "	ORDER BY CODE"

			Set rs_cnt = db.execute(sql)

			If not rs_cnt.eof Then
				totalCount = CDBL(rs_cnt("cnt"))
			Else
				totalCount = 0
			End if

			rs_cnt.close
			Set rs_cnt = Nothing

			sFistLine = "<tr bgcolor='#EEF6FF'>"
			sFistLine = sFistLine& "<td colspan='2' align='center' class='TDCont'>구분</td>"

			sSecondLine = "<tr bgcolor='#EEF6FF'>"
			sSecondLine = sSecondLine& "<td align='center' class='TDCont' width='80' nowrap>매체</td>"
			sSecondLine = sSecondLine& "<td align='center' class='TDCont' width='100' nowrap>유형</td>"

			If whereCD7 = "A" Then

				'----------------------------------------------------------------------------------------
				'상담구분
				'----------------------------------------------------------------------------------------
				SQL = "	SELECT ACLASS,	BCLASS		"
				SQL = SQL & "	FROM	TB_GOODBUNU"
				SQL = SQL &	"	WHERE	1=1 "
				SQL = SQL & "	AND		ACLASS = 'A'	AND	BCLASS = '201'	AND	CCLASS IS NULL"
				SQL = SQL & "	UNION	"
				SQL = SQL & "	SELECT ACLASS,	'202'		"
				SQL = SQL & "	FROM	TB_GOODBUNU"
				SQL = SQL &	"	WHERE	1=1 "
				SQL = SQL & "	AND		ACLASS = 'A'	AND	BCLASS > '201'	AND	CCLASS IS NULL"	
				
				totalCount = CDBL(totalCount) * 2

			Else

				SQL = "	SELECT count(0) as cnt		"
				SQL = SQL & "	FROM	TB_GOODBUNU"
				SQL = SQL &	"	WHERE	1=1 "
				SQL = SQL & "	AND		ACLASS ='" & whereCD7 &"'	AND	BCLASS IS NOT NULL	AND	CCLASS IS NULL"

				Set rs_cnt = db.execute(sql)

				If not rs_cnt.eof Then
					totalCount = CDBL(rs_cnt("cnt")) * CDBL(totalCount)
				End if

				rs_cnt.close
				Set rs_cnt = Nothing

				'----------------------------------------
				'2) 상담원 현황(가로타이틀)
				'----------------------------------------
				SQL = "	SELECT ACLASS,	BCLASS		"
				SQL = SQL & "	FROM	TB_GOODBUNU"
				SQL = SQL &	"	WHERE	1=1 "
				SQL = SQL & "	AND		ACLASS ='" & whereCD7 &"'	AND	BCLASS IS NOT NULL	AND	CCLASS IS NULL"

			End If

			totalCount = totalCount + 1
			
			ReDim bound_code((totalCount)+2)
			ReDim bound_sub((totalCount)+3)
			ReDim bound_tot((totalCount)+3)

			Set rs = Server.CreateObject("ADODB.Recordset")
			rs.open SQL,db,1

			i = 1
			Do Until RS.EOF 

				'--------------------------------------------------------------------------------------------------------
				SQL = "	SELECT	*	FROM	TB_CODE	WHERE CODEGROUP = 'A04'"
				SQL = SQL & "	ORDER BY CODE"

				Set codeRs = Server.CreateObject("ADODB.Recordset")
				codeRs.open SQL,db,1
				iColspan = 0

				Do Until codeRs.eof
					i = i + 1
					iColspan = iColspan + 1
					bound_code(i) = rs("BCLASS")&"_"&codeRs("Code")
					sSecondLine = sSecondLine& "<td align='center' class='TDCont' width='80' nowrap>" & codeRs("CodeName") & "</td>"
					codeRs.movenext
				Loop

				If rs("ACLASS") = "A" And rs("BCLASS") = "202" Then
					sFistLine = sFistLine& "<td colspan='"&iColspan&"' align='center' class='TDCont'>하위버전</td>"				
				else
					sFistLine = sFistLine& "<td colspan='"&iColspan&"' align='center' class='TDCont'>"& db_getCateNameB_(rs("ACLASS"),rs("BCLASS")) &"</td>"
				End if
				rs.movenext

			loop

			sFistLine = sFistLine& "<td bgcolor='#EEF6FF' rowspan='2' align='center' class='TDCont' width='80' nowrap>소계</td>"

			rs.Close
			Set rs = Nothing
			sFistLine = sFistLine & "</tr>"
			sSecondLine = sSecondLine & "</tr>"

		else

			'----------------------------------------
			'1) 상단에는 제품군
			'----------------------------------------

			SQL ="	SELECT	count(0) AS CNT	FROM	TB_GOODBUNU"
			SQL = SQL &	"							WHERE	1=1 "
			SQL = SQL & "							AND		ACLASS NOT IN ('A')	AND	BCLASS IS NOT NULL	AND	CCLASS IS NULL"        				'2차분류갯수 카운드

			Set rs_cnt = db.execute(sql)

			If not rs_cnt.eof Then
				totalCount = CDBL(rs_cnt("cnt")) + 3	'오피스군은 무조건 2개로
			Else
				totalCount = 0
			End if

			rs_cnt.close
			Set rs_cnt = Nothing

			ReDim bound_code((totalCount)+2)
			ReDim bound_sub((totalCount)+3)
			ReDim bound_tot((totalCount)+3)

			'----------------------------------------
			'2) 상담원 현황(가로타이틀)
			'----------------------------------------
			SQL = "	SELECT ACLASS,	BCLASS		"
			SQL = SQL & "	FROM	TB_GOODBUNU"
			SQL = SQL &	"	WHERE	1=1 "
			SQL = SQL & "	AND		ACLASS NOT IN ('A')	AND	BCLASS IS NOT NULL	AND	CCLASS IS NULL"

			Set rs = Server.CreateObject("ADODB.Recordset")
			rs.open SQL,db,1

			sFistLine = "<tr bgcolor='#EEF6FF'>"
			sFistLine = sFistLine& "<td colspan='2' align='center' class='TDCont'>구분</td>"
			sSecondLine = "<tr bgcolor='#EEF6FF'>"
			sSecondLine = sSecondLine& "<td align='center' class='TDCont' width='80' nowrap>매체</td>"
			sSecondLine = sSecondLine& "<td align='center' class='TDCont' width='100' nowrap>구분</td>"
			sFistLine = sFistLine& "<td colspan='2' align='center' class='TDCont'>"& db_getCodeName("A01","A") &"</td>"
			sSecondLine = sSecondLine& "<td align='center' class='TDCont' width='80' nowrap>" & db_getCateNameB_("A","201") & "</td>"
			sSecondLine = sSecondLine& "<td align='center' class='TDCont' width='80' nowrap>하위버전</td>"
			bound_code(2) ="A_201"
			bound_code(3) ="A_202"

			i = 3
			Do Until rs.EOF
				ACLASS = rs("ACLASS")
				iColspan = 0
				Do Until ACLASS <> rs("ACLASS")
					iColspan = iColspan + 1
					i = i + 1
					bound_code(i) = rs("ACLASS") & "_" & rs("BCLASS")
					sSecondLine = sSecondLine& "<td bgcolor='#EEF6FF' align='center' class='TDCont' width='80' nowrap>" &db_getCateNameB_(rs("ACLASS"),rs("BCLASS"))& "</td>"
					'sSecondLine = sSecondLine& "<td bgcolor='#EEF6FF' align='center' class='TDCont'>" &rs("ACLASS") & "_" & rs("BCLASS")& "</td>"
					rs.MoveNext
					If rs.EOF Then	Exit DO					
				Loop
				sFistLine = sFistLine& "<td colspan='"&iColspan&"' align='center' class='TDCont'>"& db_getCodeName("A01",ACLASS) &"</td>"			
				If rs.EOF Then	Exit DO
			Loop

			sFistLine = sFistLine& "<td bgcolor='#EEF6FF' rowspan='2' align='center' class='TDCont' width='80' nowrap>소계</td>"

			rs.Close
			Set rs = Nothing
			sFistLine = sFistLine & "</tr>"
			sSecondLine = sSecondLine & "</tr>"

			'-------------------------------------------------------------
			'1차제품별 합계를 위한 소계배열 초기화
			'-------------------------------------------------------------
			bound_cnt = (totalCount) + 3
			For i = 3 To bound_cnt
				bound_sub(i) = 0
				bound_tot(i) = 0
			Next 

			lv_Line_Off = 0
			lv_Line_On = 0
	
		End if

%>

<table border="0" cellpadding="0" cellspacing="0" align="center">
	<tr>
		<td>
			<DIV style="OVERFLOW-Y:auto; OVERFLOW-X:auto; MARGIN: 0px 0px 0px 0px; WIDTH:1200; HEIGHT:700;">
			<table  border="0" cellpadding="1" cellspacing="1" bgcolor="#EFECE5" align="center">
				<%=sFistLine%>
				<%=sSecondLine%>
				<tr><td colspan="100" height="1" bgcolor="#FFFFFF"></td></tr>
				<%'####### 실제자료가 들어간다. %>
				<%
				'---------------------------------------------------------------------------------------------------------------------
				'데이타 뽑아오기
				'---------------------------------------------------------------------------------------------------------------------

				If whereCD3 = "1" Then '제품별 상담구분

				    sSQL = "	SELECT	ACLASS,   BCLASS,     CALLKIND,	CHANNEL,   SUM(CNT) AS CNT	FROM	("
				    sSQL = sSQL&vbCr& " SELECT ACLASS,   BCLASS,     CALLKIND,	CHANNEL,   COUNT(*) AS CNT"
				    sSQL = sSQL&vbCr& " FROM     TB_CALLHISTORY"
				    sSQL = sSQL&vbCr& " WHERE   JUBTIME >= '" &FromDate& "' "
				    sSQL = sSQL&vbCr& " AND     JUBTIME < '" &DateAdd("d",1,ToDate)& "' "
				    sSQL = sSQL&vbCr& " AND     ACLASS = 'A' AND     BCLASS      = '201'"
				    sSQL = sSQL&vbCr& " GROUP BY ACLASS,  BCLASS, CALLKIND,	CHANNEL"
					sSQL = sSQL&vbCr& " Union all"
				    sSQL = sSQL&vbCr& " SELECT ACLASS,   '202' BCLASS,     CALLKIND,		CHANNEL,		COUNT(*) AS CNT"
				    sSQL = sSQL&vbCr& " FROM     TB_CALLHISTORY"
				    sSQL = sSQL&vbCr& " WHERE   JUBTIME >= '" &FromDate& "' "
				    sSQL = sSQL&vbCr& " AND     JUBTIME < '" &DateAdd("d",1,ToDate)& "' "
				    sSQL = sSQL&vbCr& " AND     ACLASS = 'A' AND     BCLASS  > '201'"
				    sSQL = sSQL&vbCr& " GROUP BY ACLASS,  CALLKIND,		CHANNEL"
					sSQL = sSQL&vbCr& " Union all"
				    sSQL = sSQL&vbCr& " SELECT ACLASS,   BCLASS,     CALLKIND,	CHANNEL,   COUNT(*) AS CNT"
				    sSQL = sSQL&vbCr& " FROM     TB_CALLHISTORY"
				    sSQL = sSQL&vbCr& " WHERE   JUBTIME >= '" &FromDate& "' "
				    sSQL = sSQL&vbCr& " AND     JUBTIME < '" &DateAdd("d",1,ToDate)& "' "
				    sSQL = sSQL&vbCr& " AND     ACLASS NOT IN ('A')"
				    'sSQL = sSQL&vbCr& " AND		CHANNEL = 'A'"		'채널(A:전화,B:온라인)
				    sSQL = sSQL&vbCr& " GROUP BY ACLASS,  BCLASS,    CALLKIND,	CHANNEL"
					sSQL = sSQL&vbCr& " Union all"
				    sSQL = sSQL&vbCr& " SELECT ACLASS, BCLASS, CALLKIND,	CHANNEL, SUM(CNT) AS CNT"
				    sSQL = sSQL&vbCr& " From TB_ONLINE"
				    sSQL = sSQL&vbCr& " WHERE    GIJUNDATE >= '" &FromDate& "' "
				    sSQL = sSQL&vbCr& " AND      GIJUNDATE < '" &DateAdd("d",1,ToDate)& "' "
				    sSQL = sSQL&vbCr& " AND     ACLASS = 'A' AND     BCLASS      = '201'"
				    sSQL = sSQL&vbCr& " GROUP BY ACLASS,  BCLASS, CALLKIND,		CHANNEL"
					sSQL = sSQL&vbCr& " Union all"
				    sSQL = sSQL&vbCr& " SELECT ACLASS, '202' BCLASS, CALLKIND,	CHANNEL, SUM(CNT) AS CNT"
				    sSQL = sSQL&vbCr& " From TB_ONLINE"
				    sSQL = sSQL&vbCr& " WHERE    GIJUNDATE >= '" &FromDate& "' "
				    sSQL = sSQL&vbCr& " AND      GIJUNDATE < '" &DateAdd("d",1,ToDate)& "' "
				    sSQL = sSQL&vbCr& " AND     ACLASS = 'A' AND     BCLASS  > '201'"
				    sSQL = sSQL&vbCr& " GROUP BY ACLASS,  BCLASS, CALLKIND,		CHANNEL"
					sSQL = sSQL&vbCr& " Union all"
				    sSQL = sSQL&vbCr& " SELECT ACLASS, BCLASS, CALLKIND,	CHANNEL, SUM(CNT) AS CNT"
				    sSQL = sSQL&vbCr& " From TB_ONLINE"
				    sSQL = sSQL&vbCr& " WHERE    GIJUNDATE >= '" &FromDate& "' "
				    sSQL = sSQL&vbCr& " AND      GIJUNDATE < '" &DateAdd("d",1,ToDate)& "' "
					sSQL = sSQL&vbCr& " AND     ACLASS NOT IN ('A')"
				    sSQL = sSQL&vbCr& " GROUP BY ACLASS,  BCLASS, CALLKIND,		CHANNEL	) A"
				    sSQL = sSQL&vbCr& " GROUP BY CHANNEL,	CALLKIND,	ACLASS,  BCLASS"
				    sSQL = sSQL&vbCr& " ORDER BY CHANNEL,	CALLKIND,	ACLASS,  BCLASS"	
		
					'Response.Write sSQL

					Set rs = Server.CreateObject("ADODB.Recordset")
					rs.open sSQL,db,1

					Do Until RS.EOF
						'---------------------------------------------------------
						'채널, 상담구분별로 열찾아서 뿌리기
						'---------------------------------------------------------
						sCHANNEL = RS("CHANNEL")

						Do Until sCHANNEL <> RS("CHANNEL")

							sFistLine = "<tr bgcolor='#ffffff'>"
							sFistLine = sFistLine& "<td align='center' class='TDCont'>"& db_getCodeName("A02",RS("CHANNEL")) & "</td>"
							sFistLine = sFistLine& "<td align='center' class='TDCont'>"& db_getCodeName("A04",RS("CALLKIND")) & "</td>"
							lv_loop1 = 2
							sCALLKIND = RS("CALLKIND")	
							Do Until sCHANNEL <> RS("CHANNEL") Or sCALLKIND <> RS("CALLKIND")
							
								For lv_loop = lv_loop1 To totalCount+1

									If bound_code(lv_loop) = RS("ACLASS")&"_"&RS("BCLASS") Then

										sFistLine = sFistLine& "<td class='TDCont' align='right'>"& FormatNumber(RS("CNT"),0) & "&nbsp;</td>"
										lv_LineTot = CDbl(lv_LineTot) + CDbl(RS("CNT"))
										bound_sub(lv_loop) = CDbl(bound_sub(lv_loop)) + CDbl(RS("CNT"))
										lv_loop = lv_loop + 1
										Exit for
									Else
										sFistLine = sFistLine& "<td class='TDCont' align='right'>&nbsp</td>"
									End if
								Next
								lv_loop1 = lv_loop

								rs.movenext
								If RS.EOF Then Exit DO
							Loop
	'RESPONSE.WRITE lv_loop1&","&totalCount&"<br>"
							If lv_loop1 <= totalCount Then
								For lv_loop = lv_loop1 To totalCount
									sFistLine = sFistLine & "<td>&nbsp;</td>"
								Next
							End If
							sFistLine = sFistLine& "<td class='TDCont' align='right'>"&FormatNumber(lv_LineTot,0)&"&nbsp;</td>"
							lv_LineTot = 0
							sFistLine = sFistLine & "</tr>"
							Response.Write sFistLine
							If RS.EOF Then Exit DO
						Loop
						'---------------------------------------------------------
						'채널별로 소계
						'---------------------------------------------------------
						sFistLine = "<tr bgcolor='#FCFAED'>"
						sFistLine = sFistLine & "<td class='TDCont' colspan='2' align='center'>소계</td>"
						For lv_loop = 2 To totalCount
							sFistLine = sFistLine& "<td class='TDR5px' align='right'>" &FormatNumber(bound_sub(lv_loop),0)& "</td>"
							bound_tot(lv_loop) = Cdbl(bound_tot(lv_loop)) + CDbl(bound_sub(lv_loop))
							lv_LineTot = lv_LineTot + CDbl(bound_sub(lv_loop))
							bound_sub(lv_loop) = 0
						Next
						sFistLine = sFistLine& "<td class='TDCont' align='right'>"&FormatNumber(lv_LineTot,0)&"&nbsp;</td>"
						lv_LineTot = 0
						sFistLine = sFistLine & "</tr>"
						Response.Write sFistLine

						If RS.EOF Then Exit DO
					Loop
					'---------------------------------------------------------
					'총합계
					'---------------------------------------------------------	
					lv_LineTot = 0
					lv_LineTot1 = 0
					lv_Colspan = 0
					sFistLine = "<tr bgcolor='#FCFAED'>"
					sFistLine = sFistLine & "<td class='TDCont' colspan='2' align='center'>합계</td>"

					For lv_loop = 2 To totalCount
						sFistLine = sFistLine& "<td class='TDR5px' align='right'>" &FormatNumber(bound_tot(lv_loop),0)& "</td>"
						lv_LineTot = lv_LineTot + CDbl(bound_tot(lv_loop))
						bound_sub(lv_loop) = 0
					Next
					sFistLine = sFistLine& "<td class='TDCont' align='right'>"&FormatNumber(lv_LineTot,0)&"&nbsp;</td>"
					lv_LineTot = 0
					sFistLine = sFistLine & "</tr>"
					Response.Write sFistLine

					'-----------------------------------------------------------------------------------------
					'콜백 찍어주기
					'-----------------------------------------------------------------------------------------
				    sSQL = "	SELECT	A.*	FROM	("
				    sSQL = sSQL&vbCr& " SELECT ACLASS,   BCLASS,     COUNT(*) AS CNT"
				    sSQL = sSQL&vbCr& " FROM     TB_CALLHISTORY"
				    sSQL = sSQL&vbCr& " WHERE   JUBTIME >= '" &FromDate& "' "
				    sSQL = sSQL&vbCr& " AND     JUBTIME < '" &DateAdd("d",1,ToDate)& "' "
				    sSQL = sSQL&vbCr& " AND     ACLASS = 'A' AND     BCLASS      = '201'"
				    sSQL = sSQL&vbCr& " AND     CALLBACKYN = 'Y'"
				    sSQL = sSQL&vbCr& " GROUP BY ACLASS,  BCLASS"
					sSQL = sSQL&vbCr& " Union all"
				    sSQL = sSQL&vbCr& " SELECT ACLASS,   '202' BCLASS,    COUNT(*) AS CNT"
				    sSQL = sSQL&vbCr& " FROM     TB_CALLHISTORY"
				    sSQL = sSQL&vbCr& " WHERE   JUBTIME >= '" &FromDate& "' "
				    sSQL = sSQL&vbCr& " AND     JUBTIME < '" &DateAdd("d",1,ToDate)& "' "
				    sSQL = sSQL&vbCr& " AND     ACLASS = 'A' AND     BCLASS      > '201'"
				    sSQL = sSQL&vbCr& " AND     CALLBACKYN = 'Y'"
				    sSQL = sSQL&vbCr& " GROUP BY ACLASS"
					sSQL = sSQL&vbCr& " Union all"
				    sSQL = sSQL&vbCr& " SELECT ACLASS,   BCLASS,     COUNT(*) AS CNT"
				    sSQL = sSQL&vbCr& " FROM     TB_CALLHISTORY"
				    sSQL = sSQL&vbCr& " WHERE   JUBTIME >= '" &FromDate& "' "
				    sSQL = sSQL&vbCr& " AND     JUBTIME < '" &DateAdd("d",1,ToDate)& "' "
				    sSQL = sSQL&vbCr& " AND     ACLASS NOT IN ('A')"
				    sSQL = sSQL&vbCr& " AND     CALLBACKYN = 'Y'"
				    sSQL = sSQL&vbCr& " GROUP BY ACLASS,  BCLASS) A"
					sSQL = sSQL&vbCr& " ORDER BY ACLASS,  BCLASS"

					Set rs = Server.CreateObject("ADODB.Recordset")
					rs.open sSQL,db,1

					lv_loop1 = 2
					lv_LineTot = 0

					If RS.EOF Then
						sFistLine = "<tr bgcolor='#ffffff'>"
						sFistLine = sFistLine& "<td align='center' class='TDCont' colspan='2'>CallBack</td>"		
						For lv_loop = lv_loop1 To totalCount+1
							sFistLine = sFistLine& "<td class='TDCont' align='right'>&nbsp;&nbsp;</td>"
						Next
						sFistLine = sFistLine & "</tr>"
						Response.Write sFistLine							
					Else

						sFistLine = "<tr bgcolor='#ffffff'>"
						sFistLine = sFistLine& "<td align='center' class='TDCont' colspan='2'>CallBack</td>"	

						Do Until RS.EOF
							'---------------------------------------------------------
							'채널, 상담구분별로 열찾아서 뿌리기
							'---------------------------------------------------------			
							For lv_loop = lv_loop1 To totalCount
								If Trim(bound_code(lv_loop)) = Trim(RS("ACLASS")&"_"&RS("BCLASS")) Then
									sFistLine = sFistLine& "<td class='TDCont' align='right'>"& FormatNumber(CDbl(RS("CNT")),0) & "&nbsp;</td>"
									lv_LineTot = CDbl(lv_LineTot) + CDbl(RS("CNT"))
									lv_loop = lv_loop + 1
									Exit for
								Else
									sFistLine = sFistLine& "<td class='TDCont' align='right'>&nbsp;</td>"
								End if
							Next
							lv_loop1 = lv_loop
							rs.movenext
						Loop
						'---------------------------------------------------------
						'채널별로 소계
						'---------------------------------------------------------
						If lv_loop1 < totalCount + 1 Then
							For lv_loop = lv_loop1 To totalCount
								sFistLine = sFistLine & "<td>&nbsp;</td>"
							Next
						End If
						sFistLine = sFistLine& "<td class='TDCont' align='right'>"&FormatNumber(lv_LineTot,0)&"&nbsp;</td>"
						lv_LineTot = 0
						sFistLine = sFistLine & "</tr>"
						Response.Write sFistLine
					End If

				ElseIf whereCD3 = "2" Then'제품별 상담구분

					'----------------------------------------------------------------------------------------------------------
					'제품 단계별로 다르다. 온라인게임군은 3차분류를 유형으로 가져온다.
					'----------------------------------------------------------------------------------------------------------
					If whereCD7 = "B" Or whereCD7 = "D" Then

						sSQL = "	SELECT	CHANNEL,	CALLFLAG,		ACLASS,  BCLASS,	CALLKIND, SUM(CNT) AS CNT	FROM	("
						sSQL = sSQL&vbCr& " SELECT ACLASS,   BCLASS,     CALLKIND,	CHANNEL,   CCLASS AS CALLFLAG,	COUNT(*) AS CNT"
						sSQL = sSQL&vbCr& " FROM     TB_CALLHISTORY"
						sSQL = sSQL&vbCr& " WHERE   JUBTIME >= '" &FromDate& "' "
						sSQL = sSQL&vbCr& " AND     JUBTIME <  '" &DateAdd("d",1,ToDate)& "' "
						sSQL = sSQL&vbCr& " AND     ACLASS NOT IN ('A')"
						'AND	CHANNEL = 'A'"
						sSQL = sSQL&vbCr& " GROUP BY ACLASS,  BCLASS,    CALLKIND,	CHANNEL,	CCLASS"
						sSQL = sSQL&vbCr& " Union all"
						sSQL = sSQL&vbCr& " SELECT ACLASS, BCLASS, CALLKIND,	CHANNEL, CCLASS AS CALLFLAG,	SUM(CNT) AS CNT"
						sSQL = sSQL&vbCr& " From TB_ONLINE"
						sSQL = sSQL&vbCr& " WHERE    GIJUNDATE >= '" &FromDate& "' "
						sSQL = sSQL&vbCr& " AND      GIJUNDATE <  '" &DateAdd("d",1,ToDate)& "' "
						sSQL = sSQL&vbCr& " AND     ACLASS NOT IN ('A')"
						sSQL = sSQL&vbCr& " GROUP BY ACLASS,  BCLASS, CALLKIND,		CHANNEL	,	CCLASS) A"
						sSQL = sSQL&vbCr& " WHERE	ACLASS = '" & whereCD7 & "'"
						sSQL = sSQL&vbCr& " GROUP BY CHANNEL,	CALLFLAG,		ACLASS,  BCLASS,	CALLKIND"	
						sSQL = sSQL&vbCr& " ORDER BY CHANNEL,	CALLFLAG,		ACLASS,  BCLASS,	CALLKIND"	

						'Response.Write sSQL

					Else

						sSQL = "	SELECT	CHANNEL,	CALLFLAG,	ACLASS,  BCLASS,	CALLKIND, SUM(CNT1) AS CNT	FROM	("
						sSQL = sSQL&vbCr& " SELECT ACLASS,   BCLASS,     CALLKIND,	CHANNEL,   CALLFLAG,	COUNT(*) AS CNT1"
						sSQL = sSQL&vbCr& " FROM     TB_CALLHISTORY"
						sSQL = sSQL&vbCr& " WHERE   JUBTIME >= '" &FromDate& "' "
						sSQL = sSQL&vbCr& " AND     JUBTIME < '" &DateAdd("d",1,ToDate)& "' "
						sSQL = sSQL&vbCr& " AND     ACLASS = 'A' AND     BCLASS      = '201'"
						sSQL = sSQL&vbCr& " GROUP BY ACLASS,  BCLASS, CALLKIND,	CHANNEL,	CALLFLAG"
						sSQL = sSQL&vbCr& " Union all"
						sSQL = sSQL&vbCr& " SELECT ACLASS,   '202' BCLASS,     CALLKIND,		CHANNEL,		CALLFLAG,	COUNT(*) AS CNT1"
						sSQL = sSQL&vbCr& " FROM     TB_CALLHISTORY"
						sSQL = sSQL&vbCr& " WHERE   JUBTIME >=	'" &FromDate& "' "
						sSQL = sSQL&vbCr& " AND     JUBTIME <	'" &DateAdd("d",1,ToDate)& "' "
						sSQL = sSQL&vbCr& " AND     ACLASS = 'A' AND     BCLASS  > '201'"
						sSQL = sSQL&vbCr& " GROUP BY ACLASS,  CALLKIND,		CHANNEL,	CALLFLAG"
						sSQL = sSQL&vbCr& " Union all"
						sSQL = sSQL&vbCr& " SELECT ACLASS,   BCLASS,     CALLKIND,	CHANNEL,   CALLFLAG,	COUNT(*) AS CNT1"
						sSQL = sSQL&vbCr& " FROM     TB_CALLHISTORY"
						sSQL = sSQL&vbCr& " WHERE   JUBTIME >= '" &FromDate& "' "
						sSQL = sSQL&vbCr& " AND     JUBTIME <	'" &DateAdd("d",1,ToDate)& "' "
						sSQL = sSQL&vbCr& " AND     ACLASS NOT IN ('A')"'			AND	CHANNEL = 'A'"
						sSQL = sSQL&vbCr& " GROUP BY ACLASS,  BCLASS,    CALLKIND,	CHANNEL,	CALLFLAG"

						sSQL = sSQL&vbCr& " Union all"
						sSQL = sSQL&vbCr& " SELECT ACLASS, BCLASS, CALLKIND,	CHANNEL, CALLFLAG,	SUM(CNT) AS CNT1"
						sSQL = sSQL&vbCr& " From TB_ONLINE"
						sSQL = sSQL&vbCr& " WHERE    GIJUNDATE >= '" &FromDate& "' "
						sSQL = sSQL&vbCr& " AND      GIJUNDATE < '" &DateAdd("d",1,ToDate)& "' "
						sSQL = sSQL&vbCr& " AND     ACLASS = 'A' AND     BCLASS      = '201'"
						sSQL = sSQL&vbCr& " GROUP BY ACLASS,  BCLASS, CALLKIND,		CHANNEL	,	CALLFLAG"
						sSQL = sSQL&vbCr& " Union all"
						sSQL = sSQL&vbCr& " SELECT ACLASS, '202' BCLASS, CALLKIND,	CHANNEL, CALLFLAG,	SUM(CNT) AS CNT1"
						sSQL = sSQL&vbCr& " From TB_ONLINE"
						sSQL = sSQL&vbCr& " WHERE    GIJUNDATE >= '" &FromDate& "' "
						sSQL = sSQL&vbCr& " AND      GIJUNDATE < '" &DateAdd("d",1,ToDate)& "' "
						sSQL = sSQL&vbCr& " AND     ACLASS = 'A' AND     BCLASS  > '201'"
						sSQL = sSQL&vbCr& " GROUP BY ACLASS,  BCLASS, CALLKIND,		CHANNEL	,	CALLFLAG"
						sSQL = sSQL&vbCr& " Union all"
						sSQL = sSQL&vbCr& " SELECT ACLASS, BCLASS, CALLKIND,	CHANNEL, CALLFLAG,	SUM(CNT) AS CNT1"
						sSQL = sSQL&vbCr& " From TB_ONLINE"
						sSQL = sSQL&vbCr& " WHERE    GIJUNDATE >= '" &FromDate& "' "
						sSQL = sSQL&vbCr& " AND      GIJUNDATE < '" &DateAdd("d",1,ToDate)& "' "
						sSQL = sSQL&vbCr& " AND     ACLASS NOT IN ('A')"
						sSQL = sSQL&vbCr& " GROUP BY ACLASS,  BCLASS, CALLKIND,		CHANNEL	,	CALLFLAG) A"
						sSQL = sSQL&vbCr& " WHERE	ACLASS = '" & whereCD7 & "'"
						'sSQL = sSQL&vbCr& " AND	CALLFLAG = 'H'"
						sSQL = sSQL&vbCr& " GROUP BY CHANNEL,	CALLFLAG,	ACLASS,  BCLASS,	CALLKIND"	
						sSQL = sSQL&vbCr& " ORDER BY CHANNEL,	CALLFLAG,	ACLASS,  BCLASS,	CALLKIND"	
						
					End if
		

					'Response.Write sSQL

					Set rs = Server.CreateObject("ADODB.Recordset")
					rs.open sSQL,db,1

					Do Until RS.EOF
						'---------------------------------------------------------
						'채널, 상담구분별로 열찾아서 뿌리기
						'---------------------------------------------------------
						sCHANNEL = RS("CHANNEL")

						Do Until sCHANNEL <> RS("CHANNEL")

							sFistLine = "<tr bgcolor='#ffffff'>"
							sFistLine = sFistLine& "<td align='center' class='TDCont'>"& db_getCodeName("A02",RS("CHANNEL")) & "</td>"
							If RS("ACLASS") = "B" then
								sFistLine = sFistLine& "<td align='center' class='TDCont'>"& db_getCateNameC_(RS("ACLASS"),RS("BCLASS"),RS("CALLFLAG")) & "</td>"
							Else
								sFistLine = sFistLine& "<td align='center' class='TDCont'>"& db_getCallFlagName(RS("ACLASS"),RS("BCLASS"),"",RS("CALLFLAG")) & "</td>"
							End if
							lv_loop1 = 2
							sCALLFLAG = RS("CALLFLAG")	
							Do Until sCHANNEL <> RS("CHANNEL") Or sCALLFLAG <> RS("CALLFLAG")								
								For lv_loop = lv_loop1 To totalCount

									'Response.Write lv_loop1 & ":"& bound_code(lv_loop) & "," & RS("BCLASS")&"_"&RS("CALLFLAG") &"_"&RS("CNT") &":<br>"
									If bound_code(lv_loop) = RS("BCLASS")&"_"&RS("CALLKIND") Then

										sFistLine = sFistLine& "<td class='TDCont' align='right'>"& FormatNumber(RS("CNT"),0) & "&nbsp;</td>"
										lv_LineTot = CDbl(lv_LineTot) + CDbl(RS("CNT"))
										bound_sub(lv_loop) = CDbl(bound_sub(lv_loop)) + CDbl(RS("CNT"))
										lv_loop = lv_loop + 1
										Exit for
									Else
										sFistLine = sFistLine& "<td class='TDCont' align='right'>&nbsp;&nbsp;</td>"
									End if
								Next
								lv_loop1 = lv_loop

								rs.movenext
								If RS.EOF Then Exit DO
							Loop

							If lv_loop1 <= totalCount Then
								For lv_loop = lv_loop1 To totalCount
									sFistLine = sFistLine & "<td>&nbsp;</td>"
								Next
							End If
							sFistLine = sFistLine& "<td class='TDCont' align='right'>"&FormatNumber(lv_LineTot,0)&"&nbsp;</td>"
							lv_LineTot = 0
							sFistLine = sFistLine & "</tr>"
							Response.Write sFistLine
							If RS.EOF Then Exit DO
						Loop
						'---------------------------------------------------------
						'채널별로 소계
						'---------------------------------------------------------

						sFistLine = "<tr bgcolor='#FCFAED'>"
						sFistLine = sFistLine & "<td class='TDCont' colspan='2' align='center'>소계</td>"
						For lv_loop = 2 To totalCount
							sFistLine = sFistLine& "<td class='TDR5px' align='right'>" &FormatNumber(bound_sub(lv_loop),0)& "</td>"
							bound_tot(lv_loop) = Cdbl(bound_tot(lv_loop)) + CDbl(bound_sub(lv_loop))
							lv_LineTot = lv_LineTot + CDbl(bound_sub(lv_loop))
							bound_sub(lv_loop) = 0
						Next
						sFistLine = sFistLine& "<td class='TDCont' align='right'>"&FormatNumber(lv_LineTot,0)&"&nbsp;</td>"
						lv_LineTot = 0
						sFistLine = sFistLine & "</tr>"
						Response.Write sFistLine

						If RS.EOF Then Exit DO
					Loop
					'---------------------------------------------------------
					'총합계
					'---------------------------------------------------------	
					lv_LineTot = 0
					lv_LineTot1 = 0
					lv_Colspan = 0
					sFistLine = "<tr bgcolor='#FCFAED'>"
					sFistLine = sFistLine & "<td class='TDCont' colspan='2' align='center'>합계</td>"

					For lv_loop = 2 To totalCount
						sFistLine = sFistLine& "<td class='TDR5px' align='right'>" &FormatNumber(bound_tot(lv_loop),0)& "</td>"
						lv_LineTot = lv_LineTot + CDbl(bound_tot(lv_loop))
						bound_sub(lv_loop) = 0
					Next
					sFistLine = sFistLine& "<td class='TDCont' align='right'>"&FormatNumber(lv_LineTot,0)&"&nbsp;</td>"
					lv_LineTot = 0
					sFistLine = sFistLine & "</tr>"
					Response.Write sFistLine


				ElseIf whereCD3 = "3" Then '제품별 매체별
	
					'채널, 상세접수경로
				    sSQL = "	SELECT	CHANNEL,		ONLINEGB,	ACLASS,  BCLASS, SUM(CNT) AS CNT	FROM	("
				    sSQL = sSQL&vbCr& " SELECT ACLASS,   BCLASS,     CHANNEL,   CASE WHEN CHANNEL = 'A' THEN CALLBACKYN ELSE ONLINEGB END AS ONLINEGB,	COUNT(*) AS CNT"
				    sSQL = sSQL&vbCr& " FROM     TB_CALLHISTORY"
				    sSQL = sSQL&vbCr& " WHERE   JUBTIME >= '" &FromDate& "' "
				    sSQL = sSQL&vbCr& " AND     JUBTIME < '" &DateAdd("d",1,ToDate)& "' "
				    sSQL = sSQL&vbCr& " AND     ACLASS = 'A' AND     BCLASS      = '201'"
				    sSQL = sSQL&vbCr& " GROUP BY ACLASS,  BCLASS,	CHANNEL,	(CASE WHEN CHANNEL = 'A' THEN CALLBACKYN ELSE ONLINEGB END)"
					sSQL = sSQL&vbCr& " Union all"
				    sSQL = sSQL&vbCr& " SELECT ACLASS,   '202' BCLASS,     CHANNEL,   CASE WHEN CHANNEL = 'A' THEN CALLBACKYN ELSE ONLINEGB END AS ONLINEGB,		COUNT(*) AS CNT"
				    sSQL = sSQL&vbCr& " FROM     TB_CALLHISTORY"
				    sSQL = sSQL&vbCr& " WHERE   JUBTIME >= '" &FromDate& "' "
				    sSQL = sSQL&vbCr& " AND     JUBTIME < '" &DateAdd("d",1,ToDate)& "' "
				    sSQL = sSQL&vbCr& " AND     ACLASS = 'A' AND     BCLASS  > '201'"
				    sSQL = sSQL&vbCr& " GROUP BY ACLASS,     CHANNEL,	(CASE WHEN CHANNEL = 'A' THEN CALLBACKYN ELSE ONLINEGB END)"
					sSQL = sSQL&vbCr& " Union all"
				    sSQL = sSQL&vbCr& " SELECT ACLASS,   BCLASS,     CHANNEL,   CASE WHEN CHANNEL = 'A' THEN CALLBACKYN ELSE ONLINEGB END AS ONLINEGB,   COUNT(*) AS CNT"
				    sSQL = sSQL&vbCr& " FROM     TB_CALLHISTORY"
				    sSQL = sSQL&vbCr& " WHERE   JUBTIME >= '" &FromDate& "' "
				    sSQL = sSQL&vbCr& " AND     JUBTIME < '" &DateAdd("d",1,ToDate)& "' "
				    sSQL = sSQL&vbCr& " AND     ACLASS NOT IN ('A')"'		AND	CHANNEL = 'A'"
				    sSQL = sSQL&vbCr& " GROUP BY ACLASS,  BCLASS,    CHANNEL,	(CASE WHEN CHANNEL = 'A' THEN CALLBACKYN ELSE ONLINEGB END)"
					sSQL = sSQL&vbCr& " Union all"
				    sSQL = sSQL&vbCr& " SELECT ACLASS, BCLASS,		CHANNEL,   ONLINEGB, SUM(CNT) AS CNT"
				    sSQL = sSQL&vbCr& " From TB_ONLINE"
				    sSQL = sSQL&vbCr& " WHERE    GIJUNDATE >= '" &FromDate& "' "
				    sSQL = sSQL&vbCr& " AND      GIJUNDATE < '" &DateAdd("d",1,ToDate)& "' "
				    sSQL = sSQL&vbCr& " AND     ACLASS = 'A' AND     BCLASS      = '201'"
				    sSQL = sSQL&vbCr& " GROUP BY ACLASS,  BCLASS, CHANNEL,		ONLINEGB"
					sSQL = sSQL&vbCr& " Union all"
				    sSQL = sSQL&vbCr& " SELECT ACLASS, '202' BCLASS,		CHANNEL,   ONLINEGB, SUM(CNT) AS CNT"
				    sSQL = sSQL&vbCr& " From TB_ONLINE"
				    sSQL = sSQL&vbCr& " WHERE    GIJUNDATE >= '" &FromDate& "' "
				    sSQL = sSQL&vbCr& " AND      GIJUNDATE < '" &DateAdd("d",1,ToDate)& "' "
				    sSQL = sSQL&vbCr& " AND     ACLASS = 'A' AND     BCLASS  > '201'"
				    sSQL = sSQL&vbCr& " GROUP BY ACLASS,  BCLASS, CHANNEL,		ONLINEGB"
					sSQL = sSQL&vbCr& " Union all"
				    sSQL = sSQL&vbCr& " SELECT ACLASS, BCLASS,		CHANNEL,   ONLINEGB, SUM(CNT) AS CNT"
				    sSQL = sSQL&vbCr& " From TB_ONLINE"
				    sSQL = sSQL&vbCr& " WHERE    GIJUNDATE >= '" &FromDate& "' "
				    sSQL = sSQL&vbCr& " AND      GIJUNDATE < '" &DateAdd("d",1,ToDate)& "' "
				    sSQL = sSQL&vbCr& " AND     ACLASS NOT IN ('A')"
				    sSQL = sSQL&vbCr& " GROUP BY ACLASS,  BCLASS, CHANNEL,		ONLINEGB) A"

					sSQL = sSQL&vbCr& " GROUP BY CHANNEL,		ONLINEGB,	ACLASS,  BCLASS"
					sSQL = sSQL&vbCr& " ORDER BY CHANNEL,		ONLINEGB,	ACLASS,  BCLASS"
'---------------------------------------------------------------------------------------------------------------------------
					'채널, 상세접수경로
'---------------------------------------------------------------------------------------------------------------------------		
					'Response.Write sSQL
					Set rs = Server.CreateObject("ADODB.Recordset")
					rs.open sSQL,db,1

					Do Until RS.EOF
						'---------------------------------------------------------
						'채널, 상담구분별로 열찾아서 뿌리기
						'---------------------------------------------------------
						sCHANNEL = RS("CHANNEL")

						Do Until sCHANNEL <> RS("CHANNEL")

							sFistLine = "<tr bgcolor='#ffffff'>"
							sFistLine = sFistLine& "<td align='center' class='TDCont'>"& db_getCodeName("A02",RS("CHANNEL")) & "</td>"
							If RS("CHANNEL") = "A" Then
								If RS("ONLINEGB") = "Y" THEN
									sFistLine = sFistLine& "<td align='center' class='TDCont'>콜백</td>"
								Else
									sFistLine = sFistLine& "<td align='center' class='TDCont'>인바운드</td>"
								End IF
							ELSE
								sFistLine = sFistLine& "<td align='center' class='TDCont'>"& db_getCodeName("A03",RS("ONLINEGB")) & "</td>"
							End IF
							lv_loop1 = 2
							sONLINEGB = RS("ONLINEGB")	
							Do Until sCHANNEL <> RS("CHANNEL") Or sONLINEGB <> RS("ONLINEGB")								
								For lv_loop = lv_loop1 To totalCount
									If bound_code(lv_loop) = RS("ACLASS")&"_"&RS("BCLASS") Then

										sFistLine = sFistLine& "<td class='TDCont' align='right'>"& FormatNumber(RS("CNT"),0) & "&nbsp;</td>"
										lv_LineTot = CDbl(lv_LineTot) + CDbl(RS("CNT"))
										bound_sub(lv_loop) = CDbl(bound_sub(lv_loop)) + CDbl(RS("CNT"))
										lv_loop = lv_loop + 1
										Exit for
									Else
										sFistLine = sFistLine& "<td class='TDCont' align='right'>&nbsp;&nbsp;</td>"
									End if
								Next
								lv_loop1 = lv_loop

								rs.movenext
								If RS.EOF Then Exit DO
							Loop

							If lv_loop1 < totalCount + 1 Then
								For lv_loop = lv_loop1 To totalCount
									sFistLine = sFistLine & "<td>&nbsp;</td>"
								Next
							End If
							sFistLine = sFistLine& "<td class='TDCont' align='right'>"&FormatNumber(lv_LineTot,0)&"&nbsp;</td>"
							lv_LineTot = 0
							sFistLine = sFistLine & "</tr>"
							Response.Write sFistLine
							If RS.EOF Then Exit DO
						Loop
						'---------------------------------------------------------
						'채널별로 소계
						'---------------------------------------------------------

						sFistLine = "<tr bgcolor='#FCFAED'>"
						sFistLine = sFistLine & "<td class='TDCont' colspan='2' align='center'>소계</td>"
						For lv_loop = 2 To totalCount
							sFistLine = sFistLine& "<td class='TDR5px' align='right'>" &FormatNumber(bound_sub(lv_loop),0)& "</td>"
							bound_tot(lv_loop) = Cdbl(bound_tot(lv_loop)) + CDbl(bound_sub(lv_loop))
							lv_LineTot = lv_LineTot + CDbl(bound_sub(lv_loop))
							bound_sub(lv_loop) = 0
						Next
						sFistLine = sFistLine& "<td class='TDCont' align='right'>"&FormatNumber(lv_LineTot,0)&"&nbsp;</td>"
						lv_LineTot = 0
						sFistLine = sFistLine & "</tr>"
						Response.Write sFistLine

						If RS.EOF Then Exit DO
					Loop
					'---------------------------------------------------------
					'총합계
					'---------------------------------------------------------	
					lv_LineTot = 0
					lv_LineTot1 = 0
					lv_Colspan = 0
					sFistLine = "<tr bgcolor='#FCFAED'>"
					sFistLine = sFistLine & "<td class='TDCont' colspan='2' align='center'>합계</td>"

					For lv_loop = 2 To totalCount
						sFistLine = sFistLine& "<td class='TDR5px' align='right'>" &FormatNumber(bound_tot(lv_loop),0)& "</td>"
						lv_LineTot = lv_LineTot + CDbl(bound_tot(lv_loop))
						bound_sub(lv_loop) = 0
					Next
					sFistLine = sFistLine& "<td class='TDCont' align='right'>"&FormatNumber(lv_LineTot,0)&"&nbsp;</td>"
					lv_LineTot = 0
					sFistLine = sFistLine & "</tr>"
					Response.Write sFistLine
		

				End IF 
				%>

			</table>
			</DIV>
		</td>
	</tr>
</table>

<% End if %>

<!-- #include virtual="/Include/Bottom.asp" -->