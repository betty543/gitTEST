<!-- #include virtual="/Include/Top.asp" -->
<%
	'####### 파라미터 ##################################################################################
	FromDate = Trim(Request("FromDate"))
	ToDate = Trim(Request("ToDate"))
	sProcessYN = Trim(Request("sProcessYN"))
	whereCD2 = Trim(Request("whereCD2"))
	cboClassA = Request("cboClassA")

	SS_Login_Grade = SESSION("SS_Login_Grade")
	SS_LoginID = SESSION("SS_LoginID")
	SS_Login_Secgroup = SESSION("SS_Login_Secgroup") 

	
	'####### 디버깅 코드 ###############################################################################
	'Response.Write("FromDate=" &FromDate& "<br>")
	'Response.Write("ToDate=" &ToDate& "<br>")
	'Response.Write("ProcessYN=" &sProcessYN& "<br>")
	'Response.Write("whereCD2=" &whereCD2& "<br>")
	'Response.Write("DelYN=" &sDelYN& "<br>")
	
	
	if FromDate = "" then FromDate = date() end If
	if ToDate = "" then ToDate = Date() end If

	'####### 1. 파라미터 얻어오기 및 기본정보 셋팅 ###########################################################
	curPage = request("curPage")
	pageSize = 25
	pageSector = 20
	IF curPage = "" THEN curPage = 1 END IF
	pageWhere = "FromDate=" & FromDate & "&ToDate=" & ToDate
	pageWhere = pageWhere & "&sProcessYN=" & sProcessYN
	pageWhere = pageWhere & "&whereCD2=" & whereCD2
	pageWhere = pageWhere & "&cboClassA=" & cboClassA


	'pageWhere = pageWhere & "&a.JubSeq=b.JubSeq "
	
	'####### 2. 쿼리조건절 셋팅 ##############################################################################
	SQL_Table = "TB_CALLBACK"
	SQL_Field = " IDX,DNIS,RequestTime,CallBankNo,Cid,CustNo,DivideTime,ProcessGB,NONPROCESSGB,ProcessTime,ProcessCode,Memo, Jubseq, LINEKIND"
	SQL_Orderby = "IDX" '먼저 접수된것을 먼저

	sql_where = " 1=1"

	If cboClassA <> "" then
		sql_where = sql_where & " and DNIS = '" & cboClassA &"'"
	End if
	if SS_Login_Grade <> "A" then	'운영업무가
		sql_where = sql_where & " and DNIS = '" & SS_Login_Grade &"'"
	end if
	if FromDate <> "" then			sql_where = sql_where & " and REQUESTTIME >= '" & FromDate & "' " end If
	if ToDate <> "" then				sql_where = sql_where & " and REQUESTTIME < '" & DateAdd("d",1,ToDate) & "' " end If
	
	IF NOT(sProcessYN="") THEN	'처리여부
		sql_Where = sql_Where & " AND ProcessGB='" & sProcessYN &"'"
	END If
	if whereCD2 <> "" then
		sql_Where = sql_Where & "	and PROCESSCODE='" & whereCD2 &"'"
	end if

	
	'####### 3. 레코드 목록 얻어오기 #########################################################################
	SQL = db_getSqlWithPage(SQL_Table, SQL_Index, SQL_Field, SQL_Where, SQL_Orderby, pageSize, curPage)
	'Response.Write(SQL)
	set Rs = db.execute(SQL)

	'####### 4. Paging HTML 작성 #############################################################################

	totalCount = db_getCount(db, SQL_Table, SQL_Where)
	startRow = totalCount - pageSize * (curPage - 1)
	pageHtml = getPageHtml(pageSector, pageSize, totalCount, curPage, currentURL & "?" & pageWHERE)	
%>

<!-- #include virtual="/Include/PopLayer.asp" -->
<div id='CalendarLayer' style='display:none;'><iframe name='CalendarFrame' src='/include/Calendar.asp' width='172' height='177' border='0' frameborder='0' scrolling='no'></iframe></div>


<table width="1200" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">
	<form method="post" name="searchFrm" action="<%=currentURL%>">	
	<tr>
		<td bgcolor="#FFFFFF">
			
			<table width="1200" border="0" cellspacing="1" cellpadding="1" align="center">
			    <tr>
			        <td nowrap width="280">
			        	조회기간 :
						<input value="<%IF FromDate="" THEN%><%=date()%><%ElSE%><%=FromDate%><%END IF%>" name="FromDate" readonly type="text" size="10" onfocus="setFocusColor(this);" onClick="new CalendarFrame.Calendar(this);" onblur="setOutColor(this);">
			        	~
			        	<input value="<%IF ToDate="" THEN%><%=date()%><%ElSE%><%=ToDate%><%END IF%>" name="ToDate" readonly type="text" size="10" onfocus="setFocusColor(this);" onClick="new CalendarFrame.Calendar(this);" onblur="setOutColor(this);">        	
			        </td>
			        <td nowrap width="280" nowrap>
						업무구분 :
					<%
						'======= 제품분류1차 가져오기 ==================================================
						SqlCode = "SELECT Code,		CodeName	FROM TB_Code"
						SqlCode = SqlCode& " WHERE USEYN='Y'	and	codegroup = 'Z04'" '운영업무
						if 	SS_Login_Secgroup = "A" or SS_Login_Secgroup = "B" then
							SqlCode = SqlCode& " and	code = '" & SS_Login_Grade & "'"
						end if
						SqlCode = SqlCode& " ORDER BY Code ASC"

						set RsCode = db.execute(SqlCode)
					%>
					<select name="cboClassA" size="1" align="absmiddle" class="ComboFFFCE7">
						<option value="">업무선택</option>
						<%
							IF NOT(RsCode.Eof OR RsCode.bof) THEN
								DO until RsCode.EOF
									CALLFLAG = RsCode("Code")
									CALLFLAGNAME = RsCode("CodeName")
						%>
						<%=printSelect("" &CALLFLAGNAME& "","" &CALLFLAG& "",""&cboClassA&"")%>
						<%
								RsCode.MoveNext
								LOOP
							END IF
							RsCode.Close
							set RsCode = NOTHING
						%>
					</select>

			        </td>

			        <td nowrap width="150">
						처리여부 :
						<%
							SQL1 = "Select * From TB_CODE where CODEGROUP ='Z14' AND useYN = 'Y' "
							set RsCode = db.execute(SQL1)
						%>

						<select name="sProcessYN" size="1" class="ComboFFFCE7">
							<option value="">전체</option>
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("CODE")
										CODENAME = RsCode("CODENAME")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &sProcessYN& "")%>
							<%
									RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
							%>
						</select>	

			        </td>

					<td nowrap width="150">						
						처리자 :
						<%
							'======= 상담원 가져오기 ==================================================
							SqlCode = "SELECT USERID, USERNAME FROM TB_USERINFO"
							SqlCode = SqlCode& " WHERE USEYN='Y'"
							if 	SS_Login_Secgroup = "A" or SS_Login_Secgroup = "B" then
								SqlCode = SqlCode& " and	grade = '" & SS_Login_Grade & "'"
							end if

							SqlCode = SqlCode& " ORDER BY USEYN DESC, GRADE ASC, USERNAME ASC"
							set RsCode = db.execute(SqlCode)
						%>
						<select name="whereCD2" size="1" class="ComboFFFCE7">
							<option value="">전체</option>
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("USERID")
										CODENAME = RsCode("USERNAME")
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


			        <td align="right"><img src="/Images/Btn/BtnSearch.gif" style="cursor:hand;" onClick="javascript:goSearch(document.searchFrm);">&nbsp;</td>
			    </tr>
			</table>
		</td>
	</tr>
	</form>
</table>
<table width="1200" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>

<table width="1200" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">
	<tr height="20" bgcolor="#EEF6FF" align="center">
		<td align="center" width="40">No</td>
		<td align="center" width=60 nowrap>업무구분</td>
		<td align="center" width="60">전화구분</td>
		<td align="center" width="120" nowrap>콜백 요청시간</td>
		<td align="center" width="60">고객명</td>
		<td align="center" width="120">콜백 전화번호</td>
		<td align="center" width="120">발신번호</td>
		<td align="center" width="30">상태</td>	
		<td align="center" width="70">처리여부</td>	
		<!--<td align="center" width="100">분배시각</td>-->
		<td align="center" width="100">처리시각</td>
		<td align="center" width="80">처리자</td>
		<td align="center" width="250">메모</td>
	</tr>
	<tr><td colspan="12" height="1" bgcolor="#FFFFFF"></td></tr>

	<% IF FromDate="" THEN %>
	<tr><td height="50" colspan="12" align="center" bgcolor="#FFFFFF" style="color:#0000FF">기간별 검색을 해주시기 바랍니다.</td></tr>
	<% ELSE	%>
		<%
			IF Rs.EOF OR Rs.BOF THEN
		%>
	<tr><td height="50" colspan="12" align="center" bgcolor="#FFFFFF" style="color:#0000FF">조건에 만족하는 데이타가 없습니다.</td></tr>
		<%
			ELSE

				DO UNTIL Rs.EOF
					SEQ = Rs("IDX")
					ACLASS = Rs("DNIS")
					REQUESTTIME = FormatDateH(Rs("REQUESTTIME"))
					CALLBANKNO = Rs("CALLBANKNO")
					CID = ""
					If IsNull(Rs("CID")) then
						'CID = CALLBANKNO
					Else
						CID = Rs("CID")
						If CID = "000000000"  Then
							CID = "" 
						End if
					End If
					If CID = "" Then
						CID = CALLBANKNO
					End If
					
					if CID <> "" and isnull(Rs("CustNo")) then
						SQL = "select top 1 * from tb_custinfo where ( cellphone = '"&CID&"' or homephone = '"&CID&"' or sendphone = '"&CID&"')"

						set RsCode = db.execute(SQL)
						if RsCode.eof = false then
							CUSTNO = RsCode("CUSTNO")
							CUSTNAME = ""
						end if
					else
						CUSTNAME = db_getCustName(Rs("CustNo"))
						CUSTNO = Rs("CustNo")
					end if


					If IsNull( Rs("DIVIDETIME")) = false Then
						DIVIDETIME = FormatDateH(Rs("DIVIDETIME"))
					Else
						DIVIDETIME =""
					End If
					sPROCESSGB = Rs("PROCESSGB")
					PROCESSGB = db_getCodeName("Z14", Rs("PROCESSGB"))
					If IsNull( Rs("PROCESSTIME")) = false Then
						PROCESSTIME = FormatDateH(Rs("PROCESSTIME"))
					Else
						PROCESSTIME =""
					End if

					PROCESSNAME = db_getUserName(Rs("PROCESSCODE"))
					MEMO = Rs("MEMO")
					LINEKIND = Rs("LINEKIND")'"SIP-DigitalE1"
					if instr(LINEKIND,"sip:5001") > 0 or instr(LINEKIND,"sip:5002") > 0 then
						LINEKIND_NAME = "군전화"
					else
						LINEKIND_NAME = "일반"
					end if

					'URL ="/manage/AsRegi/AsRegi.asp?InType=CALLBACK&Cate1="&ACLASS&"&Channel=A&CUSTNO="&CUSTNO&"&telNo="&CALLBANKNO&"&Pid="&PID&"&CB_SEQ="&SEQ
					URL ="/manage/AsRegi/AsRegi.asp?InType=CALLBACK&Cate1="&ACLASS&"&Channel=A&CUSTNO="&CUSTNO&"&telNo="&CID&"&Pid="&PID&"&CB_SEQ="&SEQ&"&CALLBACKPHONE="&CALLBANKNO


					'------------------------------------------------------------------
					'업무에 따라서 움직인다.
					if Rs("DNIS") = "B" then
						URL ="/menu03/submenu0302/lifecallmanage.asp?InType=CALLBACK&LINEKIND="&LINEKIND&"&telNo="&CID&"&CB_SEQ="&SEQ
					else
						URL ="/menu04/submenu0402/callmanage.asp?TELKIND="&ACLASS&"&InType=CALLBACK&LINEKIND="&LINEKIND&"&telNo="&CID&"&CB_SEQ="&SEQ
					end if

					If (sPROCESSGB="C" Or sPROCESSGB="D") THEN TRBGColor="#FFFFFF" ELSE TRBGColor="#FDE6F3" END IF
		%>

					<tr bgcolor="<%=TRBGColor%>" onmouseover="this.style.background='#FFFCE7'" onmouseout="this.style.background='<%=TRBGColor%>' ">
						<td align="center"><%=startRow%></td>
						<td align="center"><%=LINEKIND_NAME%></td>
						<td align="center"><%=mid(LINEKIND,5,4)%></td>
						<td class="TDCont" nowrap align='center'><a href ='<%=URL%>'><%=left(REQUESTTIME,8)%>&nbsp;<%=right(REQUESTTIME,5)%></a></td>
						<td align="center"><%=CUSTNAME%></td>
						<td class="TDCont" nowrap><%=FormatCallNo(CALLBANKNO)%>&nbsp;<img src="/Images/Cti/icon_tel.gif" style="cursor:hand;" onClick="callTel('<%=CALLBANKNO%>');" align="absmiddle" title="<%=CALLBANKNO%> 전화걸기"></td>
						<td class="TDCont" nowrap><%=FormatCallNo(CID)%>&nbsp;<img src="/Images/Cti/icon_tel.gif" style="cursor:hand;" onClick="callTel('<%=CID%>');" align="absmiddle" title="<%=CID%> 전화걸기"></td>


						<!--<td align="center"><%=DIVIDETIME%></td>-->

						<td align="center">
							<%IF IsNull(rs("Jubseq")) THEN%>
							<img src="/Images/Btn/BtnIconModify.gif" title='콜백결과 저장' style="cursor:hand;" onClick="javascript:goDetail('<%=SEQ%>');">
						<!--<img src="/Images/Btn/BtnIconDel.gif" title='콜백결과 삭체' style="cursor:hand;" onClick="javascript:goDetail('<%=SEQ%>');">-->
							<%END IF%>
						</td>
						<td align="center"><%=PROCESSGB%></td>
						<td align="center"><%=PROCESSTIME%></td>
						<td align="center"><%=PROCESSNAME%></td>						
						<td title="<%=MEMO%>" class="TDCont"><%=Left(MEMO,40)%><%If Len(MEMO) > 40 Then %>...<%End if%></td>
					</tr>
		<%
							startRow = startRow - 1
					Rs.MoveNext
				LOOP
			END IF
			
			Rs.close
			set Rs = Nothing
		%>
	<% END IF %>

</table>

<table border="0" cellpadding="0" cellspacing="0" width="1200" align="center">
	<tr><td height="5"></td></tr>
	<tr><td height="1" bgcolor="#D6D6D6"></td></tr>
	<tr height="22" bgcolor="#EEF6FF"><td align="center"><%=pageHtml%></td></tr>
	<tr><td height="1" bgcolor="#D6D6D6"></td></tr>
</table>

<!--</form>//-->

<script>
<!--

	function goSearch(form)
	{
		form.submit();
	}
	

	function callTel(sTel)
	{
		//전화걸기

		top.CallStateFrame.document.all.txtCID.value = sTel;

		if ( top.CallStateFrame.document.all.txtCID.value == "" )
			alert('전화걸기 실패 : 전화번호가 입력되지 않음');
		else
			top.CallStateFrame.vfn_MakeCall(top.CallStateFrame.document.all.txtCID.value,'');
	}


	function MovePageConsel(sURL)
	{

		location.href = sURL;
	}

	function goDetail(_seq){		
		ShowPOPLayer('CallbackUp.asp?curPage=<%=curPage%>&<%=pageWhere%>&seq='+_seq,'500','230');
	}
	
//-->
</script>

<!-- #include virtual="/Include/Bottom.asp" -->