<!-- #include virtual="/Include/Top.asp" -->
<!-- #include virtual="/Include/PopLayer.asp" -->
<div id='CalendarLayer' style='display:none;'><iframe name='CalendarFrame' src='/include/Calendar.asp' width='172' height='177' border='0' frameborder='0' scrolling='no'></iframe></div>


<table width="1200" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">
	<form method="post" name="searchFrm" action="<%=currentURL%>">	
	<tr>
		<td bgcolor="#FFFFFF">
			
			<table width="1190" border="0" cellspacing="1" cellpadding="1" align="center">
			    <tr>
			        <td nowrap width="300">
			        	조회기간 :
						<input value="<%IF FromDate="" THEN%><%=date()%><%ElSE%><%=FromDate%><%END IF%>" name="FromDate" readonly type="text" size="10" onfocus="setFocusColor(this);" onClick="new CalendarFrame.Calendar(this);" onblur="setOutColor(this);">
			        	~
			        	<input value="<%IF ToDate="" THEN%><%=date()%><%ElSE%><%=ToDate%><%END IF%>" name="ToDate" readonly type="text" size="10" onfocus="setFocusColor(this);" onClick="new CalendarFrame.Calendar(this);" onblur="setOutColor(this);">        	
			        </td>
			        <td nowrap width="150">
						전화구분 :
					<%
						'======= 제품분류1차 가져오기 ==================================================
						SqlCode = "SELECT Code,		CodeName	FROM TB_Code"
						SqlCode = SqlCode& " WHERE USEYN='Y'	and	codegroup = 'A01'"
						SqlCode = SqlCode& " ORDER BY Code ASC"

						set RsCode = db.execute(SqlCode)
					%>
					<select name="cboClassA" size="1" align="absmiddle" class="ComboFFFCE7">
						<option value="">성범죄신고</option>
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
							SQL1 = "Select * From TB_CODE where CODEGROUP ='Z14' "
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
						<%IF SS_Login_Grade="A" THEN%>
						처리자 :
						<%
							'======= 상담원 가져오기 ==================================================
							SqlCode = "SELECT USERID, USERNAME FROM TB_USERINFO"
							SqlCode = SqlCode& " WHERE USEYN='Y'"
							SqlCode = SqlCode& " ORDER BY USEYN DESC, GRADE ASC, USERNAME ASC"
							set RsCode = db.execute(SqlCode)
						%>
						<select name="whereCD2" size="1" class="ComboFFFCE7">
							<option value="">상담관선택----</option>
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
						<%ELSE%>
						<input type="hidden" name="whereCD2" value="<%=whereCD2%>">
						<%END IF%>
					</td>


			        <td align="right"><img src="/Images/Btn/BtnSearch.gif" style="cursor:hand;" onClick="javascript:goSearch(document.searchFrm);"></td>
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
		<td align="center" width="80">전화구분</td>
		<td align="center" width="100">콜백 요청시간</td>
		<td align="center" width="60">요청자명</td>
		<td align="center" width="130">계급</td>
		<td align="center" width="130">소속</td>
		<td align="center" width="120">콜백 전화번호</td>
		<td align="center" width="120">발신번호</td>
		<td align="center" width="30">상태</td>	
		<td align="center" width="70">처리여부</td>	
		<!--<td align="center" width="100">분배시각</td>-->
		<td align="center" width="100">처리시각</td>
		<td align="center" width="80">처리자</td>
		<td align="center" width="250">메모</td>
	</tr>
	<tr><td colspan="13" height="1" bgcolor="#FFFFFF"></td></tr>

	<tr height="20" bgcolor="#ffffff" align="center">
		<td align="center" width="40">1</td>
		<td align="center" width="80">성범죄신고</td>
		<td align="center" width="100">2009-05-20 03:00:01</td>
		<td align="center" width="60">김00</td>
		<td align="center" width="130">00</td>
		<td align="center" width="130">소속1-소속2-소속3</td>
		<td align="center" width="120">01051850478  <img src="/Images/Comm/IconAlert.gif" style="cursor:hand;" onClick="fn_dial('3');" align="absmiddle" title="전화걸기"> </td>
		<td align="center" width="120">0422501111 <img src="/Images/Comm/IconAlert.gif" style="cursor:hand;" onClick="fn_dial('3');" align="absmiddle" title="전화걸기"></td>
		<td align="center" width="30">접수</td>	
		<td align="center" width="70">미처리 <img src="/Images/Btn/BtnIconModify.gif" title='콜백결과 저장' style="cursor:hand;" onClick="javascript:goDetail('<%=SEQ%>');"> </td>	
		<!--<td align="center" width="100">분배시각</td>-->
		<td align="center" width="100"></td>
		<td align="center" width="80"></td>
		<td align="center" width="250">1차: 2009-05-20 02:03:05</td>
	</tr>
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
		//현재 통화중인지 확인한다
		if (opener.parent.CallStateFrame.document.CallStateForm.txtStatus.value =="busy")
		{
			alert("현재 통화중 상태입니다. 통화종료후 다시!");
			return;
		}
		opener.parent.CallStateFrame.document.CallStateForm.txtTelno.value=sTel;
		opener.parent.CallStateFrame.vfn_MakeCall();
	}

	function MovePageConsel(sURL)
	{
		//alert(sURL);
		opener.location.href = sURL;
		opener.focus();
//		self.close();
	}

	function goDetail(_seq){		
		ShowPOPLayer('CallbackUp.asp?curPage=<%=curPage%>&<%=pageWhere%>&seq='+_seq,'500','230');
	}
	
//-->
</script>

<!-- #include virtual="/Include/Bottom.asp" -->

