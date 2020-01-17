<!-- #include virtual="/Include/Top.asp" -->

<%
'####### 파라미터 ##################################################################################
SS_Login_Secgroup = SESSION("SS_Login_Secgroup")
SS_Login_Grade = SESSION("SS_Login_Grade")
SS_Login_CTIID = SESSION("SS_Login_CTIID")
SS_Login_EXTNO = SESSION("SS_Login_EXTNO")
SS_LoginID = SESSION("SS_LoginID")

QueryYN = request("QueryYN")
FromDate = request("FromDate")
ToDate = request("ToDate")
whereCD1 = Trim(request("whereCD1"))
whereCD2 = Trim(request("whereCD2"))
whereCD3 = Trim(request("whereCD3"))
whereCD7 = Trim(request("whereCD7"))
whereCD8 = Trim(request("whereCD8"))
whereCD9 = Trim(request("whereCD9"))

whereCD2 = Replace(whereCD2," ","")

CHANNELGB1 = request("CHANNELGB1")
CHANNELGB2 = request("CHANNELGB2")
CHANNELGB3 = request("CHANNELGB3")
CHANNELGB4 = request("CHANNELGB4")

MAN = request("MAN")
WOMAN = request("WOMAN")

If QueryYN = "" Then
	whereCD3 = "1"
End If



if FromDate = "" then FromDate = left(Date(),7) & "-01" end If
if ToDate = "" then ToDate = date() end If

pageWHERE = "QueryYN=" & QueryYN & "&FromDate=" & FromDate & "&ToDate=" & ToDate
pageWHERE = pageWHERE & "&whereCD1=" & whereCD1 & "&whereCD2=" & whereCD2 & "&whereCD3=" & whereCD3 & "&whereCD7=" & whereCD7 & "&whereCD8=" & whereCD8 & "&whereCD9=" & whereCD9
pageWHERE = pageWHERE & "&channelGb1=" & CHANNELGB1 & "&channelGb2=" & CHANNELGB2 & "&channelGb3=" & CHANNELGB3 & "&channelGb4=" & CHANNELGB4& "&MAN="&MAN& "&WOMAN="&WOMAN

'RESPONSE.WRITE pageWHERE

If CHANNELGB1 <> "" then
	CHANNELGB = "''" & CHANNELGB1 & "''"
End If
If CHANNELGB2 <> "" And CHANNELGB = "" then
	CHANNELGB = "''" & CHANNELGB2 & "''"
ElseIf CHANNELGB2 <> "" then
	CHANNELGB = CHANNELGB & ",''" & CHANNELGB2 & "''"
End If
If CHANNELGB3 <> "" And CHANNELGB = "" then
	CHANNELGB = "''" & CHANNELGB3 & "''"
ElseIf CHANNELGB3 <> "" then
	CHANNELGB = CHANNELGB & ",''" & CHANNELGB3 & "''"
End If
If CHANNELGB4 <> "" And CHANNELGB = "" then
	CHANNELGB = "''" & CHANNELGB4 & "''"
ElseIf CHANNELGB4 <> "" then
	CHANNELGB = CHANNELGB & ",''" & CHANNELGB4 & "''"
End If

JEONDOR = ""
If MAN = "" Then
	JEONDOR = "N"
Else
	JEONDOR = "Y"
End If
If WOMAN = "" Then
	JEONDOR = JEONDOR & "N"
Else
	JEONDOR = JEONDOR & "Y"
End if

%>

<div id='CalendarLayer' style='display:none;'><iframe name='CalendarFrame' src='/include/Calendar.asp' width='172' height='177' border='0' frameborder='0' scrolling='no'></iframe></div>

<script>
	function fn_Search(){
		if (document.all.whereCD8.value == document.all.whereCD9.value){
			alert('세로항목과 가로항목을 다르게 설정하세요!')
			return false;
		}
		document.inUpFrm.QueryYN.value = "Y";
		document.inUpFrm.submit();
	}

	function fn_Xls() {
		location.href="list01_Xls.asp?<%=pageWHERE%>";
	}
</script>

<table border="0" width="1200" cellpadding="0" cellspacing="3" bgcolor="#EFECE5" align="center">
	<tr bgcolor="#FFFFFF">
		<td>

			<form name="inUpFrm" method="post" action="<%=Menu_2nd%>" onsubmit="return fn_Search(this);" style="margin:0">
				<input type="hidden" name="QueryYN" value="<%=QueryYN%>">

				<table width="100%" border="0" cellspacing="1" cellpadding="0" style="border:#E1DED6 solid 1px">
					<tr>
						<td class="TDCont" colspan="7">
							전화종류 :
							<input type="checkbox" name="CHANNELGB1" <% if CHANNELGB1 = "0179" then %>checked<% end if %> value="0179" class="none" />080생명
							<input type="checkbox" name="CHANNELGB2" <% if CHANNELGB2 = "13031" then %>checked<% end if %> value="13031" class="none"  >1303생명
							<input type="checkbox" name="CHANNELGB3" <% if CHANNELGB3 = "13032" then %>checked<% end if %> value="13032" class="none"  >1303성범죄
							<input type="checkbox" name="CHANNELGB4" <% if CHANNELGB4 = "13033" then %>checked<% end if %> value="13033" class="none"  >1303군범죄
							<!--<input type="checkbox" name="CHANNELGB4" <% if CHANNELGB4 = "CYBER" then %>checked<% end if %> value="CYBER" class="none"  >사이버-->
						</td>

					</tr>
					<tr>
						<td class="TDCont" >조회기간 :
							<input value="<%=FromDate%>" name="FromDate" readonly type="text" size="10" onfocus="setFocusColor(this); new CalendarFrame.Calendar(this);" onblur="setOutColor(this);">
							~
							<input value="<%=ToDate%>" name="ToDate" readonly type="text" size="10" onfocus="setFocusColor(this); new CalendarFrame.Calendar(this);" onblur="setOutColor(this);">
						</td>
						<td class="TDCont" COLSPAN = '6'>상담유형 :

							<%
							'======= 처리구분 코드 가져오기 ==================================================
							SqlCode = "SELECT BCLASS CODE, CLASSNAME CODENAME FROM TB_ARMYINFO"
							SqlCode = SqlCode& " WHERE ACLASS = 'Q' AND BCLASS IS NOT NULL AND CCLASS IS NULL"
							SqlCode = SqlCode& " ORDER BY ACLASS"
							set RsCode = db.execute(SqlCode)

							Do Until rsCode.eof
								sCode = RsCode("CODE")
								sCodeName = RsCode("CODENAME")
								If InStr(whereCD2,sCode) > 0 then
									sChecked = "checked"
								ElseIf whereCD2 = "" Then
									sChecked = ""
								else
									sChecked = ""
								End if
								%>
								<input type="checkbox" name="whereCD2" value="<%=sCode%>" class="none" <%=sChecked%>><%=sCodeName%>
								<%
								rsCode.movenext
							loop
							%>
							&nbsp;

						</td>

						<td class="TDR5px" rowspan='2'>
							<img src="/Images/Btn/BtnSearch.gif" align="absmiddle" style="cursor:hand;" onClick="fn_Search();">
							<img src="/Images/Btn/BtnExcel.gif" align="absmiddle" style="cursor:hand;" onClick="fn_Xls();">
						</td>
					</tr>
					<tr>
						<td class="TDCont" WIDTH="300">세로기준항목 : &nbsp;
							<select name="whereCD8" size="1" class="ComboFFFCE7" >
								<%=printSelect("상담관","상담관","" &whereCD8& "")%>
								<%=printSelect("시간","시간","" &whereCD8& "")%>
								<%=printSelect("요일","요일","" &whereCD8& "")%>
								<%=printSelect("상담유형","상담유형","" &whereCD8& "")%>
								<%=printSelect("계급","계급","" &whereCD8& "")%>
								<%=printSelect("부대1차","부대1차","" &whereCD8& "")%>
								<%=printSelect("부대2차","부대2차","" &whereCD8& "")%>
								<%=printSelect("부대3차","부대3차","" &whereCD8& "")%>
								<%=printSelect("부대4차","부대4차","" &whereCD8& "")%>
								<%=printSelect("부대","부대","" &whereCD8& "")%>
								<%=printSelect("상담분야","상담분야","" &whereCD8& "")%>
								<%=printSelect("통화시간","통화시간","" &whereCD8& "")%>
								<%=printSelect("조치별","조치별","" &whereCD8& "")%>
								<%=printSelect("날씨별","날씨별","" &whereCD8& "")%>
								<%=printSelect("인지경로","인지경로","" &whereCD8& "")%>
								<%=printSelect("가족사항","가족사항","" &whereCD8& "")%>
								<%=printSelect("원인제공자","원인제공자","" &whereCD8& "")%>
								<%=printSelect("월","월","" &whereCD8& "")%>
							</select>
						</td>
						<td class="TDCont" WIDTH="200">가로항목 : &nbsp;
							<select name="whereCD9" size="1" class="ComboFFFCE7" >
								<%=printSelect("상담유형","상담유형","" &whereCD9& "")%>
								<%=printSelect("계급","계급","" &whereCD9& "")%>
								<%=printSelect("부대1차","부대1차","" &whereCD9& "")%>
								<%=printSelect("부대2차","부대2차","" &whereCD9& "")%>
								<!--<%=printSelect("부대3차","부대3차","" &whereCD9& "")%>	-->
								<!--<%=printSelect("부대4차","부대4차","" &whereCD9& "")%>
								<%=printSelect("부대","부대","" &whereCD9& "")%>	-->
								<%=printSelect("상담분야","상담분야","" &whereCD9& "")%>
								<%=printSelect("통화시간","통화시간","" &whereCD9& "")%>
								<%=printSelect("조치별","조치별","" &whereCD9& "")%>
								<%=printSelect("날씨별","날씨별","" &whereCD9& "")%>
								<%=printSelect("인지경로","인지경로","" &whereCD9& "")%>
								<%=printSelect("가족사항","가족사항","" &whereCD9& "")%>
								<%=printSelect("원인제공자","원인제공자","" &whereCD9& "")%>
								<%=printSelect("상담관","상담관","" &whereCD9& "")%>
								<%=printSelect("시간","시간","" &whereCD9& "")%>
								<%=printSelect("요일","요일","" &whereCD9& "")%>
							</select>
						</td>
						<td class="TDCont" WIDTH="200">상담관 :
							<%
							'======= 상담원 가져오기 ==================================================
							SqlCode = "SELECT USERID, USERNAME FROM TB_USERINFO"
							SqlCode = SqlCode& " WHERE USEYN='Y' "
							SqlCode = SqlCode& " AND SECGROUP = 'A'"
							if SS_Login_Grade <> "A" then
								SqlCode = SqlCode& "	AND GRADE = '"&SS_Login_Grade&"'"
							end if
							if SS_Login_Secgroup = "A" then	'상담관일때는 내것만
								SqlCode = SqlCode& "	AND USERID = '" &SS_LoginID&"'"
							end if

							SqlCode = SqlCode& " ORDER BY USEYN DESC, GRADE ASC, USERNAME ASC"
							set RsCode = db.execute(SqlCode)
							%>
							<select name="whereCD1" size="1" class="ComboFFFCE7">
								<option value="">전체</option>
								<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("USERID")
										CODENAME = RsCode("USERNAME")
										%>
										<%=printSelect("" &CODENAME& "","" &CODE& "","" &whereCD1& "")%>
										<%
										RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING

								'======= 상담원 가져오기 ==================================================
								SqlCode = "SELECT CTIID, USERNAME FROM TB_USERINFO"
								SqlCode = SqlCode& " WHERE USEYN='N'  and	outdate >= '"&DateAdd("d",1,DateAdd("m",-1,Date())) &"'"
								if SS_Login_Grade <> "A" then
									SqlCode = SqlCode& "	AND GRADE = '"&SS_Login_Grade&"'"
								end if
								if SS_Login_Secgroup = "A" then	'상담관일때는 내것만
									SqlCode = SqlCode& "	AND USERID = '" &SS_LoginID&"'"
								end if

								SqlCode = SqlCode& " ORDER BY USEYN DESC, GRADE ASC, USERNAME ASC"
								set RsCode = db.execute(SqlCode)

								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("CTIID")
										CODENAME = "[퇴직]"&RsCode("USERNAME")
										%>
										<%=printSelect("" &CODENAME& "","" &CODE& "","" &whereCD1& "")%>
										<%
										RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
								%>
							</select>
						</td>
						<td class="TDCont" colspan="1">
							성&nbsp;&nbsp;&nbsp;&nbsp;별 : 
							<input type="checkbox" name="MAN" <% if MAN = "Y" then %>checked<% end if %> value="Y" class="none"  >남군
							<input type="checkbox" name="WOMAN" <% if WOMAN = "Y" then %>checked<% end if %> value="Y" class="none"  >여군
							<!--<input type="checkbox" name="CHANNELGB4" <% if CHANNELGB4 = "CYBER" then %>checked<% end if %> value="CYBER" class="none"  >사이버-->
						</td>
					</tr>
				</table>
			</form>
		</td>
	</tr>
</table>

<table border="0" width="1000" cellpadding="0" cellspacing="0" align="center"><tr height="5"><td></td></tr></table>

<% If QueryYN = "Y" Then %>

	<table border="0" cellpadding="0" cellspacing="0" align="center">
		<tr>
			<td>
				<%
				'세로항목
				cols = 1
				sDepth = 1
				'response.write whereCD8
				', 날씨,인지경로, 가족사항, 원인제공자, 상담관 , 시간, 요일,통화시간, - 1depth
				if whereCD8 = "상담유형" then'상담유형 - 2depth
					sSero = "<td align='center' class='TDCont'  width='300' colspan='2' "
					sSero1 = "상담유형</td>"
					cols = 2
					sDepth = 2
					sSelectCol = "CHANNELGB_B col_1, CHANNELGB_C col_2"
					sGroupBy = "CHANNELGB_B, CHANNELGB_C"
					sSelectCol1 = "CHANNELGB_B, CHANNELGB_C"
					sNullCol = " AND ( isnull(CHANNELGB_B,'''') <> '''' or isnull(CHANNELGB_C,'''') <> '''') "
				elseif whereCD8 = "계급" then'계급     - 3depth
					sSero = "<td align='center' class='TDCont'  width='450' colspan='3' "
					sSero1 = "계급</td>"
					cols = 3
					sDepth = 3
					sSelectCol = "LEVEL_B col_1, LEVEL_C col_2, LEVEL_D col_3"
					sGroupBy = "LEVEL_B, LEVEL_C, LEVEL_D"
					sSelectCol1 = "LEVEL_B, LEVEL_C, LEVEL_D"
'					sSelectCol1 = "ISNULL(LEVEL_B,'''') LEVEL_B, ISNULL(LEVEL_C,'''') LEVEL_C, ISNULL(LEVEL_D,'''') LEVEL_D"
					sNullCol = " AND ( RTRIM(isnull(LEVEL_B,'''')) <> '''' or RTRIM(isnull(LEVEL_C,'''')) <> '''' or RTRIM(isnull(LEVEL_D,'''')) <> '''') "
					sNullCol = " AND RTRIM(isnull(LEVEL_B,'''')) <> ''''"
				elseif whereCD8 = "부대" then'부대	  - 5depth
					sSero = "<td align='center' class='TDCont'  width='750' colspan='5' "
					sSero1 = "부대</td>"
					cols = 5
					sDepth = 5
					sSelectCol = "SOSOKGB_A col_1, SOSOKGB_B col_2, SOSOKGB_C col_3, SOSOKGB_D col_4, SOSOKGB_E col_5"
					sGroupBy = "SOSOKGB_A, SOSOKGB_B, SOSOKGB_C, SOSOKGB_D, SOSOKGB_E"
					sSelectCol1 = "SOSOKGB_A, SOSOKGB_B, SOSOKGB_C, SOSOKGB_D, SOSOKGB_E"
					sNullCol = " AND ( isnull(SOSOKGB_A,'''') <> '''' or isnull(SOSOKGB_B,'''') <> '''' or isnull(SOSOKGB_C,'''') <> '''' or isnull(SOSOKGB_D,'''') <> '''' or isnull(SOSOKGB_E,'''') <> '''') "
				elseif whereCD8 = "부대1차" then'부대	  - 5depth
					sSero = "<td align='center' class='TDCont'  width='750' colspan='1' "
					sSero1 = "부대</td>"
					cols = 1
					sDepth = 1
					sSelectCol = "SOSOKGB_A col_1"
					sGroupBy = "SOSOKGB_A"
					sSelectCol1 = "SOSOKGB_A"
					sNullCol = " AND isnull(SOSOKGB_A,'''') <> '''' "
				elseif whereCD8 = "부대2차" then'부대	  - 5depth
					sSero = "<td align='center' class='TDCont'  width='750' colspan='2' "
					sSero1 = "부대</td>"
					cols = 2
					sDepth = 2
					sSelectCol = "SOSOKGB_A col_1, SOSOKGB_B col_2"
					sGroupBy = "SOSOKGB_A, SOSOKGB_B"
					sSelectCol1 = "SOSOKGB_A, SOSOKGB_B"
					sNullCol = " AND ( isnull(SOSOKGB_A,'''') <> '''' or isnull(SOSOKGB_B,'''') <> '''') "
				elseif whereCD8 = "부대3차" then'부대	  - 5depth
					sSero = "<td align='center' class='TDCont'  width='750' colspan='3' "
					sSero1 = "부대</td>"
					cols = 3
					sDepth = 3
					sSelectCol = "SOSOKGB_A col_1, SOSOKGB_B col_2, SOSOKGB_C col_3"
					sGroupBy = "SOSOKGB_A, SOSOKGB_B, SOSOKGB_C"
					sSelectCol1 = "SOSOKGB_A, SOSOKGB_B, SOSOKGB_C"
					sNullCol = " AND ( isnull(SOSOKGB_A,'''') <> '''' or isnull(SOSOKGB_B,'''') <> '''' or isnull(SOSOKGB_C,'''') <> '''') "
				elseif whereCD8 = "부대4차" then'부대	  - 5depth
					sSero = "<td align='center' class='TDCont'  width='750' colspan='4' "
					sSero1 = "부대</td>"
					cols = 4
					sDepth = 4
					sSelectCol = "SOSOKGB_A col_1, SOSOKGB_B col_2, SOSOKGB_C col_3, SOSOKGB_D col_4"
					sGroupBy = "SOSOKGB_A, SOSOKGB_B, SOSOKGB_C, SOSOKGB_D"
					sSelectCol1 = "SOSOKGB_A, SOSOKGB_B, SOSOKGB_C, SOSOKGB_D"
					sNullCol = " AND ( isnull(SOSOKGB_A,'''') <> '''' or isnull(SOSOKGB_B,'''') <> '''' or isnull(SOSOKGB_C,'''') <> '''' or isnull(SOSOKGB_D,'''') <> '''') "
				elseif whereCD8 = "상담분야" then'상담분야 - 2depth
					sSero = "<td align='center' class='TDCont'  width='300' colspan='2' "
					sSero1 = "상담분야</td>"
					cols = 2
					sDepth = 2
					sSelectCol = "CALLCLASS_B col_1, CALLCLASS_C col_2"
					sGroupBy = "CALLCLASS_B, CALLCLASS_C"
					sSelectCol1 = "CALLCLASS_B, CALLCLASS_C"
					sNullCol = " AND ( isnull(CALLCLASS_B,'''') <> '''' or isnull(CALLCLASS_C,'''') <> '''') "
				elseif whereCD8 = "조치별" then'조치별 - 1depth
					sSero = "<td align='center' class='TDCont'  width='300' colspan='2' "
					sSero1 = "조치별</td>"
					cols = 2
					sDepth = 2
					sSelectCol = "PROCESSGB_B col_1, PROCESSGB_C col_2"
					sGroupBy = "PROCESSGB_B, PROCESSGB_C"
					sSelectCol1 = "PROCESSGB_B, PROCESSGB_C"
					sNullCol = " AND ( isnull(PROCESSGB_B,'''') <> '''' or isnull(PROCESSGB_C,'''') <> '''') "
				elseif whereCD8 = "날씨별" then'날씨 - 1depth
					sSero = "<td align='center' class='TDCont'  width='150' colspan='1' "
					sSero1 = "날씨별</td>"
					sSelectCol = "WEATHER col_1"
					sGroupBy = "WEATHER"
					sSelectCol1 = "WEATHER"
					sNullCol = " AND  isnull(WEATHER,'''') <> '''' "
				elseif whereCD8 = "인지경로" then'인지경로 - 1depth
					sSero = "<td align='center' class='TDCont'  width='150' colsFpan='1' "
					sSero1 = "인지경로</td>"
					sSelectCol = "CALLFLAG col_1"
					sGroupBy = "CALLFLAG"
					sSelectCol1 = "CALLFLAG"
					sNullCol = " AND  isnull(CALLFLAG,'''') <> '''' "
				elseif whereCD8 = "가족사항" then'가족사항 - 1depth
					sSero = "<td align='center' class='TDCont'  width='150' colspan='1' "
					sSero1 = "가족사항</td>"
					sSelectCol = "FAMILYGB col_1"
					sGroupBy = "FAMILYGB"
					sSelectCol1 = "FAMILYGB"
					sNullCol = " AND  isnull(FAMILYGB,'''') <> '''' "
				elseif whereCD8 = "원인제공자" then'원인제공자 - 1depth
					sSero = "<td align='center' class='TDCont'  width='150' colspan='1' "
					sSero1 = "원인제공자</td>"
					sSelectCol = "CALLKIND_B col_1"
					sGroupBy = "CALLKIND_B"
					sSelectCol1 = "CALLKIND_B"
					sNullCol = " AND  isnull(CALLKIND_B,'''') <> '''' "
				elseif whereCD8 = "상담관" then'상담관 - 1depth
					sSero = "<td align='center' class='TDCont'  width='150' colspan='1' "
					sSero1 = "상담관</td>"
					sSelectCol = "INCODE col_1"
					sGroupBy = "INCODE"
					sSelectCol1 = "INCODE"
					sNullCol = " AND  isnull(INCODE,'''') <> '''' "
				elseif whereCD8 = "시간" then'시간 - 1depth
					sSero = "<td align='center' class='TDCont'  width='150' colspan='1' "
					sSero1 = "시간</td>"
					sSelectCol = "datepart(hour,JUBTIME) COL_1"
					sSelectCol1 = "JUBTIME"
					sGroupBy = "datepart(hour,JUBTIME)"
					sNullCol = " "
				elseif whereCD8 = "요일" then'요일 - 1depth
					sSero = "<td align='center' class='TDCont'  width='150' colspan='1' "
					sSero1 = "요일</td>"
					sSelectCol = "datepart(WEEKDAY,JUBTIME) COL_1"
					sGroupBy = "datepart(WEEKDAY,JUBTIME)"
					sSelectCol1 = "JUBTIME"
					sNullCol = " "
				elseif whereCD8 = "통화시간" then'통화시간 - 1depth
					sSero = "<td align='center' class='TDCont'  width='150' colspan='1' "
					sSero1 = "통화시간</td>"
					sSelectCol = "CALLTIME COL_1"
					sSelectCol1 = "CALLTIME"
					sGroupBy = "CALLTIME"
					sNullCol = " "
				elseif whereCD8 = "월" then'월 - 1depth
					sSero = "<td align='center' class='TDCont'  width='150' colspan='1' "
					sSero1 = "월</td>"
					sSelectCol = "convert(varchar(7),JUBTIME,121) COL_1"
					sGroupBy = "convert(varchar(7),JUBTIME,121)"
					sSelectCol1 = "JUBTIME"
					sNullCol = " "
				end if

				'---- 원인제공자
				'sCOLNM = "CALLKIND_B"

				sSQL = "DELETE FROM TMP_CODE_VALUE"
				db.execute(sSQL)
				', 날씨,인지경로, 가족사항, 원인제공자, 상담관 , 시간, 요일,통화시간, - 1depth
				if whereCD9 = "상담유형" then'상담유형 - 2depth

					'-----가로항목 뿌리기
					rowspan = 2
					sSQL = "	select ACLASS, BCLASS, CCLASS, CLASSNAME from TB_ARMYINFO "
					sSQL = sSQL & "	where ACLASS = 'Q' AND BCLASS IS NOT NULL AND CCLASS IS NULL"
					sSQL = sSQL & "	ORDER BY ACLASS, BCLASS, CCLASS "
					set RsCode = db.execute(sSQL)

					Do Until rsCode.eof

						sCode = RsCode("BCLASS")
						sCodeName = RsCode("CLASSNAME")

						'2DEPTH 값 찾기
						iCol = 0

						'secondLine = ""
						sCodeList = ""
						sSQL = "	select ACLASS, BCLASS, CCLASS, CLASSNAME from TB_ARMYINFO "
						sSQL = sSQL & "	where ACLASS = 'Q' AND BCLASS = '" &sCode&"'  AND CCLASS IS NOT NULL"
						sSQL = sSQL & "	ORDER BY ACLASS, BCLASS, CCLASS "
						set RsCode1 = db.execute(sSQL)

						Do Until rsCode1.eof

							sCode = RsCode1("CCLASS")
							sCodeName = RsCode1("CLASSNAME")

							iCol = iCol + 1
							cols = cols + 1

							sSQL = "INSERT INTO TMP_CODE_VALUE VALUES ("&cols& ",'CHANNELGB_C','" & sCode & "')"
							db.execute(sSQL)

							If sCodeList = "" then
								sCodeList = sCode
							Else
								sCodeList = sCodeList & "|" & sCode
							End if

							secondLine = secondLine & "<td align='center' class='TDCont'  width='150'>" & sCodeName & "</td>"
							rsCode1.movenext
							'소계
							
						Loop

						If iCol = 0 Then
							
							sSQL = "INSERT INTO TMP_CODE_VALUE VALUES ("&cols& ",'CHANNELGB_B','" & sCode & "')"
							db.execute(sSQL)
							cols = cols + 1

							firstLine = firstLine & "<td align='center' class='TDCont' colspan="&iCol+1&" rowspan='2' width='150'>"&RsCode("CLASSNAME")&"</td>"

						Else
							
							cols = cols + 1
							secondLine = secondLine & "<td align='center' class='TDCont'  width='150'>계</td>"
							sSQL = "INSERT INTO TMP_CODE_VALUE VALUES ("&cols& ",'CHANNELGB_C','" & sCodeList & "')"
							db.execute(sSQL)

							firstLine = firstLine & "<td align='center' class='TDCont' colspan="&iCol+1&">"&RsCode("CLASSNAME")&"</td>"

						End if

						rsCode.movenext
						'소계
					Loop
					'총계
					%>
					
					<DIV style="OVERFLOW-X: auto;OVERFLOW-Y: auto; MARGIN: 0px 0px 0px 0px; WIDTH: 1200; HEIGHT=400;">
						<table width="<%=sWidth%>"  border="1" cellpadding="0" cellspacing="0" bgcolor="#EFECE5" align="center" bordercolor="black" bordercolordark="white" bordercolorlight="black">
							
							<%
							firstLine = firstLine & "<td align='center' class='TDCont' rowspan="& rowspan &" width='150'>계</td>"
							firstLine = "<tr bgcolor='#EEF6FF'>" & sSero & " rowspan="& rowspan &">"& sSero1 & firstLine &"</tr>"
							secondLine = "<tr bgcolor='#EEF6FF'>" & secondLine &"</tr>"
							response.write firstLine
							response.write secondLine
	
							''-----세로항목 뿌리기
							sCOLNM = "CHANNELGB_B"
							sCOLCD = ""
	
							sSQL = " EXEC SP_SUM_BY_HISTORY_BCLASS " & sDepth & ",'" & FromDate & "','" & ToDate & "','TB_LIFECALLHISTORY','"&whereCD2 & "','"& whereCD1 & "','"&sCOLNM&"','Q','','"&sGroupBy&"','"&sSelectCol&"','"&sSelectCol1&"','"&sNullCol&"','"&CHANNELGB&"','"&JEONDOR&"','',''"
							'response.write sSQL
							set rsSUM = db.execute(sSQL)
	
							firstLine = ""
	
							Do Until rsSUM.eof
								sBG = "#ffffff"
	
								firstLine = ""
	
								'--------------키에 해당하는 값
								For i = 1 To sDepth
									sUser = rsSUM("col_"&i)
									'sCodeName = db_GetUserName(sUser)
	
									If IsNull(rsSUM("col_"&i)) Then
										
										sBG = "#EEF6FF"
	
										firstLine = firstLine & "<td align='center' class='TDCont' colspan='"&sDepth-i+1&"' WIDTH='200'>계</td>"
	
										Exit for
										
									Else
										
										If i = 1 Then
											sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i),"","","","")
										ElseIf i = 2 Then
											sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-1),rsSUM("col_"&i),"","","")
										ElseIf i = 3 Then
											sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i),"","")
										ElseIf i = 4 Then
											sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-3),rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i),"")
										Else
											sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-4),rsSUM("col_"&i-3),rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i))
										End if
	
										firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='200'>"&sCodeName&"</td>"
	
									End If
								Next
	
								'--------------가로항목의 summary
								For i = sDepth + 1 To rsSUM.Fields.count
	
									firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='200'>"&rsSUM("col_"&i)&"</td>"
	
								Next
	
								firstLine = "<tr bgcolor='"&sBG&"'>" & firstLine
	
								response.write firstLine & "</tr>"
	
								rsSUM.movenext
								'소계
							Loop
							%>
							
						</table>
					</div>
	
					<%
	
				elseif whereCD9 = "계급" then'계급     - 3depth
	
					rowspan = 3
					sSQL = "	select  ACLASS, BCLASS, CCLASS, CLASSNAME from TB_ARMYINFO "
					sSQL = sSQL & "	where ACLASS = 'P' AND BCLASS IS NOT NULL AND CCLASS IS NULL AND DCLASS IS NULL"
					sSQL = sSQL & "	ORDER BY ACLASS, BCLASS, CCLASS "
					set RsCode = db.execute(sSQL)
	
					Do Until rsCode.eof
	
						sCode = RsCode("BCLASS")
						sCodeName = RsCode("CLASSNAME")
	
						'2DEPTH 값 찾기
						iCol = 0
						sCodeList = ""
						sCodeList_C = ""
						sSQL = "	select ACLASS, BCLASS, CCLASS, CLASSNAME from TB_ARMYINFO "
						sSQL = sSQL & "	where ACLASS = 'P' AND BCLASS = '" &sCode&"'  AND CCLASS IS NOT NULL AND DCLASS IS NULL"
						sSQL = sSQL & "	ORDER BY ACLASS, BCLASS, CCLASS "
						set RsCode1 = db.execute(sSQL)
	
						'response.write sSQL & vbcrlf
						Do Until rsCode1.eof
	
							sCode = RsCode1("CCLASS")
							sCodeName = RsCode1("CLASSNAME")
							iCol = iCol + 1
							iCol1 = 0
	
							'3DEPTH 값 찾기
							If sCodeList_C = "" then
								sCodeList_C = RsCode1("CCLASS")
							Else
								sCodeList_C = sCodeList_C & "|" & RsCode1("CCLASS")
							End if
	
							'secondLine = ""
							sCodeList = ""
							sSQL = "	select ACLASS, BCLASS, CCLASS, DCLASS, CLASSNAME from TB_ARMYINFO "
							sSQL = sSQL & "	where ACLASS = 'P' AND BCLASS = '" &RsCode("BCLASS")&"'  AND CCLASS = '" & RsCode1("CCLASS")&"' AND DCLASS IS NOT NULL"
							sSQL = sSQL & "	ORDER BY ACLASS, BCLASS, CCLASS, DCLASS"
							set RsCode2 = db.execute(sSQL)
							'-------------------------------------------------------------------------------------------------------------------------------------
	
							'response.write sSQL & vbcrlf
							Do Until rsCode2.eof
	
								iCol1 = iCol1 + 1
								cols = cols + 1
								iCol = iCol + 1
								sSQL = "INSERT INTO TMP_CODE_VALUE VALUES ("&cols& ",'LEVEL_D','" & RsCode2("DCLASS") & "')"
								db.execute(sSQL)
	
								If sCodeList = "" then
									sCodeList = RsCode2("DCLASS")
								Else
									sCodeList = sCodeList & "|" & RsCode2("DCLASS")
								End if
	
								threeLine = threeLine & "<td align='center' class='TDCont'  width='150'>" & RsCode2("CLASSNAME") & "</td>"
								rsCode2.movenext
								'소계
	
							Loop
	
							If iCol1 <= 0 Then
								
								cols = cols + 1
								sSQL = "INSERT INTO TMP_CODE_VALUE VALUES ("&cols& ",'LEVEL_C','" & RsCode1("CCLASS") & "')"
								db.execute(sSQL)
								secondLine = secondLine & "<td align='center' class='TDCont' colspan=1 rowspan='2' width='150'>"&RsCode1("CLASSNAME")&"</td>"
								'sCodeList_C = ""
								
							Else
								
								iCol1 = iCol1 + 1
								cols = cols + 1
								sSQL = "INSERT INTO TMP_CODE_VALUE VALUES ("&cols& ",'LEVEL_D','" & sCodeList & "')"
								db.execute(sSQL)
								'iCol = iCol + 1
								threeLine = threeLine & "<td align='center' class='TDCont'  width='150' rowspan='1'>계</td>"
								secondLine = secondLine & "<td align='center' class='TDCont' colspan="&iCol1&">"&RsCode1("CLASSNAME")&"</td>"
								
							End If
							
							iCol1 = 0
							'response.write iCol & "," & vbcrlf
							rsCode1.movenext
							'소계
	
						Loop
	
						'response.write iCol & "," & vbcrlf
					If iCol = 0 Then
						
						cols = cols + 1
						sSQL = "INSERT INTO TMP_CODE_VALUE VALUES ("&cols& ",'LEVEL_B','" & RsCode("BCLASS") & "')"
						db.execute(sSQL)
	
						sCodeList_C = ""
						firstLine = firstLine & "<td align='center' class='TDCont' colspan="&iCol&" rowspan='3' width='150'>"&RsCode("CLASSNAME")&"</td>"
						
					Else
						
						cols = cols + 1
						iCol = iCol + 1
						sSQL = "INSERT INTO TMP_CODE_VALUE VALUES ("&cols& ",'LEVEL_C','" & sCodeList_C & "')"
						db.execute(sSQL)
						secondLine = secondLine & "<td align='center' class='TDCont'  width='150' rowspan='2'>계</td>"
						firstLine = firstLine & "<td align='center' class='TDCont' colspan="&iCol&" >"&RsCode("CLASSNAME")&"</td>"
						iCol = 0
						
					End if
					
					rsCode.movenext
	
					'소계
				Loop
				'총계
				sWidth = cols * 100
				%>
				
				<DIV style="OVERFLOW-X: auto;OVERFLOW-Y: auto; MARGIN: 0px 0px 0px 0px; WIDTH: 1200; HEIGHT=400;">
				<table width="<%=sWidth%>"  border="1" cellpadding="0" cellspacing="0" bgcolor="#EFECE5" align="center" bordercolor="black" bordercolordark="white" bordercolorlight="black">
					
				<%
				firstLine = firstLine & "<td align='center' class='TDCont' rowspan="& rowspan &" width='150'>계</td>"
				firstLine = "<tr bgcolor='#EEF6FF'>" & sSero & " rowspan="& rowspan &">"& sSero1 & firstLine &"</tr>"
				secondLine = "<tr bgcolor='#EEF6FF'>" & secondLine &"</tr>"
				threeLine = "<tr bgcolor='#EEF6FF'>" & threeLine &"</tr>"
				response.write firstLine
				response.write secondLine
				response.write threeLine
	
				''-----세로항목 뿌리기
				sCOLNM = "LEVEL_B"
				sCOLCD = ""
	
				sSQL = " EXEC SP_SUM_BY_HISTORY_BCLASS " & sDepth & ",'" & FromDate & "','" & ToDate & "','TB_LIFECALLHISTORY','"&whereCD2 & "','"& whereCD1 & "','"&sCOLNM&"','Q','','"&sGroupBy&"','"&sSelectCol&"','"&sSelectCol1&"','"&sNullCol&"','"&CHANNELGB&"','"&JEONDOR&"','',''"
	
				'response.write sSQL

				'response.end
	
				set rsSUM = db.execute(sSQL)
				firstLine = ""
	
				Do Until rsSUM.eof
					sBG = "#ffffff"
					firstLine = ""
					'--------------키에 해당하는 값
					For i = 1 To sDepth
						sUser = rsSUM("col_"&i)
						'sCodeName = db_GetUserName(sUser)
	
						If IsNull(rsSUM("col_"&i)) Then
							sBG = "#EEF6FF"
							firstLine = firstLine & "<td align='center' class='TDCont' colspan='"&sDepth-i+1&"' WIDTH='200'>계</td>"
							Exit for
						Else
							If i = 1 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i),"","","","")
							ElseIf i = 2 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-1),rsSUM("col_"&i),"","","")
							ElseIf i = 3 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i),"","")
							ElseIf i = 4 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-3),rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i),"")
							Else
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-4),rsSUM("col_"&i-3),rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i))
							End if
							firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='200'>"&sCodeName&"</td>"
						End If
					Next
					'--------------가로항목의 summary
					For i = sDepth + 1 To rsSUM.Fields.count
	
						firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='200'>"&rsSUM("col_"&i)&"</td>"
	
					Next
					firstLine = "<tr bgcolor='"&sBG&"'>" & firstLine
					response.write firstLine & "</tr>"
					rsSUM.movenext
	
					'소계
				Loop
	
				%>
	
				</table>
				</div>
				
				<%
	
			elseif whereCD9 = "부대1차" then'부대	  - 5depth
	
				rowspan = 1
				sSQL = "	select  ACLASS, BCLASS, CCLASS, CLASSNAME from TB_ARMYINFO "
				sSQL = sSQL & "	where ACLASS < 'O' AND BCLASS IS NULL"
				sSQL = sSQL & "	ORDER BY ACLASS, BCLASS, CCLASS "
				set RsCode = db.execute(sSQL)
	
				Do Until rsCode.eof
	
					sCode = RsCode("ACLASS")
					sCodeName = RsCode("CLASSNAME")
	
					'2DEPTH 값 찾기
					iCol = 0
					'secondLine = ""
					cols = cols + 1
					iCol = iCol + 1
					firstLine = firstLine & "<td align='center' class='TDCont' colspan=1 width='150'>"&RsCode("CLASSNAME")&"</td>"
					
					rsCode.movenext
					'소계
				Loop
				'총계
				sWidth = cols * 150
				%>
				
				<DIV style="OVERFLOW-X: auto;OVERFLOW-Y: auto; MARGIN: 0px 0px 0px 0px; WIDTH: 1200; HEIGHT=400;">
				<table width="<%=sWidth%>"  border="1" cellpadding="0" cellspacing="0" bgcolor="#EFECE5" align="center" bordercolor="black" bordercolordark="white" bordercolorlight="black">
					
				<%
				firstLine = firstLine & "<td align='center' class='TDCont' rowspan="& rowspan &" width='150'>계</td>"
				firstLine = "<tr bgcolor='#EEF6FF'>" & sSero & " rowspan="& rowspan &">"& sSero1 & firstLine &"</tr>"
				response.write firstLine
	
				'---- 소속1차
				sCOLNM = "SOSOKGB_A"
				sCOLCD = ""
	
				sSQL = " EXEC SP_SUM_BY_HISTORY_ACLASS " & sDepth & ",'" & FromDate & "','" & ToDate & "','TB_LIFECALLHISTORY','"&whereCD2 & "','"& whereCD1 & "','"&sCOLNM&"','"&sCOLCD&"','"&sGroupBy&"','"&sSelectCol&"','"&sSelectCol1&"','"&sNullCol&"','"&CHANNELGB&"','"&JEONDOR&"','',''"
	
				'response.write sSQL
	
				set rsSUM = db.execute(sSQL)
				firstLine = ""
	
				Do Until rsSUM.eof
					sBG = "#ffffff"
					firstLine = ""
					For i = 1 To sDepth
						sUser = rsSUM("col_"&i)
						'sCodeName = db_GetUserName(sUser)
	
						If IsNull(rsSUM("col_"&i)) Then
							sBG = "#EEF6FF"
							firstLine = firstLine & "<td align='center' class='TDCont' colspan='"&sDepth-i+1&"' WIDTH='200'>계</td>"
							Exit for
						Else
							If i = 1 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i),"","","","")
							ElseIf i = 2 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-1),rsSUM("col_"&i),"","","")
							ElseIf i = 3 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i),"","")
							ElseIf i = 4 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-3),rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i),"")
							Else
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-4),rsSUM("col_"&i-3),rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i))
							End if
							firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='200'>"&sCodeName&"</td>"
						End If
					Next
	
					For i = sDepth + 1 To rsSUM.Fields.count
	
						firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='200'>"&rsSUM("col_"&i)&"</td>"
	
					Next
					firstLine = "<tr bgcolor='"&sBG&"'>" & firstLine
					response.write firstLine & "</tr>"
					rsSUM.movenext
	
					'소계
				Loop
				%>
	
				</table>
				</div>
				
				<%
	
			elseif whereCD9 = "부대2차" then'부대	  - 5depth
	
				rowspan = 2
				sSQL = "	select  ACLASS, BCLASS, CCLASS, CLASSNAME from TB_ARMYINFO "
				sSQL = sSQL & "	where ACLASS < 'O' AND BCLASS IS NULL AND CCLASS IS NULL AND DCLASS IS NULL"
				sSQL = sSQL & "	ORDER BY ACLASS, BCLASS, CCLASS "
				set RsCode = db.execute(sSQL)
	
				Do Until rsCode.eof
	
					sCode = RsCode("ACLASS")
					sCodeName = RsCode("CLASSNAME")
	
					'2DEPTH 값 찾기
					iCol = 0
					'secondLine = ""
					sCodeList = ""
					sSQL = "	select ACLASS, BCLASS, CCLASS, CLASSNAME from TB_ARMYINFO "
					sSQL = sSQL & "	where ACLASS = '" &sCode&"'  AND BCLASS IS NOT NULL AND CCLASS IS NULL"
					sSQL = sSQL & "	ORDER BY ACLASS, BCLASS, CCLASS "
					set RsCode1 = db.execute(sSQL)
	
					Do Until rsCode1.eof
	
						sCode = RsCode1("BCLASS")
						sCodeName = RsCode1("CLASSNAME")
						iCol = iCol + 1
	
						'3DEPTH 값 찾기
	
						iCol1 = iCol1 + 1
						cols = cols + 1
						sSQL = "INSERT INTO TMP_CODE_VALUE VALUES ("&cols& ",'SOSOKGB_B','" & sCode & "')"
						db.execute(sSQL)
	
						If sCodeList = "" then
							sCodeList = sCode
						Else
							sCodeList = sCodeList & "|" & sCode
						End if
						secondLine = secondLine & "<td align='center' class='TDCont' colspan=1 width='150'>"&RsCode1("CLASSNAME")&"</td>"
						'response.write iCol & "," & vbcrlf
					
						rsCode1.movenext
						'소계
					Loop
	
				'response.write iCol & "," & vbcrlf
					If iCol = 0 Then
						sSQL = "INSERT INTO TMP_CODE_VALUE VALUES ("&cols& ",'SOSOKGB_A','" & sCode & "')"
						db.execute(sSQL)
						cols = cols + 1
						firstLine = firstLine & "<td align='center' class='TDCont' colspan="&iCol&" rowspan='2' width='150'>"&RsCode("CLASSNAME")&"</td>"
					Else
						sSQL = "INSERT INTO TMP_CODE_VALUE VALUES ("&cols& ",'SOSOKGB_B','" & sCodeList & "')"
						db.execute(sSQL)
						cols = cols + 1
						iCol = iCol + 1
						secondLine = secondLine & "<td align='center' class='TDCont'  width='150' rowspan='1'>계</td>"
						firstLine = firstLine & "<td align='center' class='TDCont' colspan="&iCol&" >"&RsCode("CLASSNAME")&"</td>"
						iCol = 0
					End if
					
					rsCode.movenext
	
					'소계
				Loop
				'총계
				sWidth = cols * 150
				%>
				
				<DIV style="OVERFLOW-X: auto;OVERFLOW-Y: auto; MARGIN: 0px 0px 0px 0px; WIDTH: 1200; HEIGHT=400;">
				<table width="<%=sWidth%>"  border="1" cellpadding="0" cellspacing="0" bgcolor="#EFECE5" align="center" bordercolor="black" bordercolordark="white" bordercolorlight="black">
					
				<%
				firstLine = firstLine & "<td align='center' class='TDCont' rowspan="& rowspan &" width='150'>계</td>"
				firstLine = "<tr bgcolor='#EEF6FF'>" & sSero & " rowspan="& rowspan &">"& sSero1 & firstLine &"</tr>"
				secondLine = "<tr bgcolor='#EEF6FF'>" & secondLine &"</tr>"
	
				response.write firstLine
				response.write secondLine
	
				''-----세로항목 뿌리기
				sCOLNM = "SOSOKGB_A"
				sCOLCD = ""
	
				sSQL = " EXEC SP_SUM_BY_HISTORY_BCLASS " & sDepth & ",'" & FromDate & "','" & ToDate & "','TB_LIFECALLHISTORY','"&whereCD2 & "','"& whereCD1 & "','"&sCOLNM&"','Q','','"&sGroupBy&"','"&sSelectCol&"','"&sSelectCol1&"','"&sNullCol&"','"&CHANNELGB&"','"&JEONDOR&"','',''"
	
				'response.write sSQL
				'response.end
				set rsSUM = db.execute(sSQL)
				firstLine = ""
	
				Do Until rsSUM.eof
					sBG = "#ffffff"
					firstLine = ""
					'--------------키에 해당하는 값
					For i = 1 To sDepth
						sUser = rsSUM("col_"&i)
						'sCodeName = db_GetUserName(sUser)
	
						If IsNull(rsSUM("col_"&i)) Then
							sBG = "#EEF6FF"
							firstLine = firstLine & "<td align='center' class='TDCont' colspan='"&sDepth-i+1&"' WIDTH='200'>계</td>"
							Exit for
						Else
							If i = 1 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i),"","","","")
							ElseIf i = 2 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-1),rsSUM("col_"&i),"","","")
							ElseIf i = 3 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i),"","")
							ElseIf i = 4 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-3),rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i),"")
							Else
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-4),rsSUM("col_"&i-3),rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i))
							End if
							firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='200'>"&sCodeName&"</td>"
						End If
					Next
					'--------------가로항목의 summary
					For i = sDepth + 1 To rsSUM.Fields.count
	
						firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='200'>"&rsSUM("col_"&i)&"</td>"
	
					Next
					firstLine = "<tr bgcolor='"&sBG&"'>" & firstLine
					response.write firstLine & "</tr>"
					rsSUM.movenext
	
					'소계
				Loop
				%>
	
				</table>
				</div>
				
				<%
	
			elseif whereCD9 = "부대3차" then'부대	  - 5depth
	
				rowspan = 3
				sSQL = "	select  ACLASS, BCLASS, CCLASS, CLASSNAME from TB_ARMYINFO "
				sSQL = sSQL & "	where ACLASS < 'O' AND BCLASS IS NULL AND CCLASS IS NULL AND DCLASS IS NULL"
				sSQL = sSQL & "	ORDER BY ACLASS, BCLASS, CCLASS "
				set RsCode = db.execute(sSQL)
	
				Do Until rsCode.eof
	
					sCode = RsCode("ACLASS")
					sCodeName = RsCode("CLASSNAME")
	
					'2DEPTH 값 찾기
					iCol = 0
					'secondLine = ""
					sCodeList = ""
					sSQL = "	select ACLASS, BCLASS, CCLASS, CLASSNAME from TB_ARMYINFO "
					sSQL = sSQL & "	where ACLASS = '" &sCode&"'  AND BCLASS IS NOT NULL AND CCLASS IS NULL"
					sSQL = sSQL & "	ORDER BY ACLASS, BCLASS, CCLASS "
					set RsCode1 = db.execute(sSQL)
	
					Do Until rsCode1.eof
	
						sCode = RsCode1("BCLASS")
						sCodeName = RsCode1("CLASSNAME")
						iCol = iCol + 1
	
						'3DEPTH 값 찾기
						iCol1 = 0
	
						'secondLine = ""
						sSQL = "	select ACLASS, BCLASS, CCLASS, CLASSNAME from TB_ARMYINFO "
						sSQL = sSQL & "	where ACLASS = '" &RsCode("ACLASS")&"'  AND BCLASS = '" & RsCode1("BCLASS")&"' AND CCLASS IS NOT NULL"
						sSQL = sSQL & "	ORDER BY ACLASS, BCLASS, CCLASS, DCLASS"
						set RsCode2 = db.execute(sSQL)
						'-------------------------------------------------------------------------------------------------------------------------------------
						
						Do Until rsCode2.eof
							
							iCol1 = iCol1 + 1
							cols = cols + 1
							iCol = iCol + 1
							sSQL = "INSERT INTO TMP_CODE_VALUE VALUES ("&cols& ",'SOSOKGB_C','" & RsCode2("CCLASS") & "')"
							db.execute(sSQL)
							If sCodeList = "" then
								sCodeList = RsCode2("CCLASS")
							Else
								sCodeList = sCodeList & "|" & RsCode2("CCLASS")
							End if
							
							threeLine = threeLine & "<td align='center' class='TDCont'  width='150'>" & RsCode2("CLASSNAME") & "</td>"
							
							rsCode2.movenext
							'소계
						Loop
	
						If iCol1 <= 0 Then
							cols = cols + 1
							sSQL = "INSERT INTO TMP_CODE_VALUE VALUES ("&cols& ",'SOSOKGB_B','" & RsCode1("BCLASS") & "')"
							db.execute(sSQL)
							secondLine = secondLine & "<td align='center' class='TDCont' colspan=1 rowspan='2' width='150'>"&RsCode1("CLASSNAME")&"</td>"
						Else
							iCol1 = iCol1 + 1
							cols = cols + 1
							sSQL = "INSERT INTO TMP_CODE_VALUE VALUES ("&cols& ",'SOSOKGB_B','" & RsCode1("BCLASS") & "')"
							db.execute(sSQL)
							'iCol = iCol + 1
							threeLine = threeLine & "<td align='center' class='TDCont'  width='150' rowspan='1'>계</td>"
							secondLine = secondLine & "<td align='center' class='TDCont' colspan="&iCol1&">"&RsCode1("CLASSNAME")&"</td>"
						End If
						iCol1 = 0
						
						rsCode1.movenext
						'소계
					Loop
	
					If iCol = 0 Then
						cols = cols + 1
						sSQL = "INSERT INTO TMP_CODE_VALUE VALUES ("&cols& ",'SOSOKGB_A','" & RsCode("ACLASS") & "')"
						db.execute(sSQL)
						firstLine = firstLine & "<td align='center' class='TDCont' colspan="&iCol&" rowspan='3' width='150'>"&RsCode("CLASSNAME")&"</td>"
					Else
						cols = cols + 1
						iCol = iCol + 1
						sSQL = "INSERT INTO TMP_CODE_VALUE VALUES ("&cols& ",'SOSOKGB_A','" & RsCode("ACLASS") & "')"
						db.execute(sSQL)
						secondLine = secondLine & "<td align='center' class='TDCont'  width='150' rowspan='2'>계</td>"
						firstLine = firstLine & "<td align='center' class='TDCont' colspan="&iCol&" >"&RsCode("CLASSNAME")&"</td>"
						iCol = 0
					End if
					
					rsCode.movenext
					'소계
				Loop
				'총계
				sWidth = cols * 100
				%>
				
				<DIV style="OVERFLOW-X: auto;OVERFLOW-Y: auto; MARGIN: 0px 0px 0px 0px; WIDTH: 1200; HEIGHT=400;">
				<table width="<%=sWidth%>"  border="1" cellpadding="0" cellspacing="0" bgcolor="#EFECE5" align="center" bordercolor="black" bordercolordark="white" bordercolorlight="black">
					
				<%
				firstLine = firstLine & "<td align='center' class='TDCont' rowspan="& rowspan &" width='150'>계</td>"
				firstLine = "<tr bgcolor='#EEF6FF'>" & sSero & " rowspan="& rowspan &">"& sSero1 & firstLine &"</tr>"
				secondLine = "<tr bgcolor='#EEF6FF'>" & secondLine &"</tr>"
				threeLine = "<tr bgcolor='#EEF6FF'>" & threeLine &"</tr>"
				response.write firstLine
				response.write secondLine
				response.write threeLine
	
				''-----세로항목 뿌리기
				sCOLNM = "SOSOKGB_A"
				sCOLCD = ""
	
				sSQL = " EXEC SP_SUM_BY_HISTORY_BCLASS " & sDepth & ",'" & FromDate & "','" & ToDate & "','TB_LIFECALLHISTORY','"&whereCD2 & "','"& whereCD1 & "','"&sCOLNM&"','Q','','"&sGroupBy&"','"&sSelectCol&"','"&sSelectCol1&"','"&sNullCol&"','"&CHANNELGB&"','"&JEONDOR&"','',''"
	
				'response.write sSQL
	
				set rsSUM = db.execute(sSQL)
				firstLine = ""
	
				Do Until rsSUM.eof
					sBG = "#ffffff"
					firstLine = ""
					'--------------키에 해당하는 값
					For i = 1 To sDepth
						sUser = rsSUM("col_"&i)
						'sCodeName = db_GetUserName(sUser)
	
						If IsNull(rsSUM("col_"&i)) Then
							sBG = "#EEF6FF"
							firstLine = firstLine & "<td align='center' class='TDCont' colspan='"&sDepth-i+1&"' WIDTH='200'>계</td>"
							Exit for
						Else
							If i = 1 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i),"","","","")
							ElseIf i = 2 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-1),rsSUM("col_"&i),"","","")
							ElseIf i = 3 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i),"","")
							ElseIf i = 4 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-3),rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i),"")
							Else
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-4),rsSUM("col_"&i-3),rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i))
							End if
							firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='200'>"&sCodeName&"</td>"
						End If
					Next
					'--------------가로항목의 summary
					For i = sDepth + 1 To rsSUM.Fields.count
	
						firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='200'>"&rsSUM("col_"&i)&"</td>"
	
					Next
					firstLine = "<tr bgcolor='"&sBG&"'>" & firstLine
					response.write firstLine & "</tr>"
					rsSUM.movenext
	
					'소계
				Loop
				%>
	
				</table>
				</div>
				
				<%
	
			elseif whereCD9 = "상담분야" then'상담분야 - 2depth
	
				rowspan = 2
				sSQL = "	select ACLASS, BCLASS, CCLASS, CLASSNAME from TB_ARMYINFO "
				sSQL = sSQL & "	where ACLASS = 'O' AND BCLASS IS NOT NULL AND CCLASS IS NULL"
				sSQL = sSQL & "	ORDER BY ACLASS, BCLASS, CCLASS "
				set RsCode = db.execute(sSQL)
	
				Do Until rsCode.eof
	
					sCode = RsCode("BCLASS")
					sCodeName = RsCode("CLASSNAME")
	
					'2DEPTH 값 찾기
					iCol = 0
					'secondLine = ""
					sCodeList = ""
					sSQL = "	select ACLASS, BCLASS, CCLASS, CLASSNAME from TB_ARMYINFO "
					sSQL = sSQL & "	where ACLASS = 'O' AND BCLASS = '" &sCode&"'  AND CCLASS IS NOT NULL"
					sSQL = sSQL & "	ORDER BY ACLASS, BCLASS, CCLASS "
					set RsCode1 = db.execute(sSQL)
	
					Do Until rsCode1.eof
	
						sCode = RsCode1("CCLASS")
						sCodeName = RsCode1("CLASSNAME")
	
						iCol = iCol + 1
						cols = cols + 1
						sSQL = "INSERT INTO TMP_CODE_VALUE VALUES ("&cols& ",'CALLCLASS_C','" & sCode & "')"
						db.execute(sSQL)
	
						If sCodeList = "" then
							sCodeList = sCode
						Else
							sCodeList = sCodeList & "|" & sCode
						End if
						secondLine = secondLine & "<td align='center' class='TDCont'  width='150'>" & sCodeName & "</td>"
						
						rsCode1.movenext
						'소계
					Loop
	
					If iCol = 0 Then
						sSQL = "INSERT INTO TMP_CODE_VALUE VALUES ("&cols& ",'CALLCLASS_B','" & sCode & "')"
						db.execute(sSQL)
						cols = cols + 1
						firstLine = firstLine & "<td align='center' class='TDCont' colspan="&iCol+1&" rowspan='2' width='150'>"&RsCode("CLASSNAME")&"</td>"
					Else
						sSQL = "INSERT INTO TMP_CODE_VALUE VALUES ("&cols& ",'CALLCLASS_C','" & sCodeList & "')"
						db.execute(sSQL)
						cols = cols + 1
						secondLine = secondLine & "<td align='center' class='TDCont'  width='150'>계</td>"
	
						firstLine = firstLine & "<td align='center' class='TDCont' colspan="&iCol+1&">"&RsCode("CLASSNAME")&"</td>"
					End if
					
					rsCode.movenext
					'소계
				Loop
				'총계
				sWidth = cols * 150
				%>
				
				<DIV style="OVERFLOW-X: auto;OVERFLOW-Y: auto; MARGIN: 0px 0px 0px 0px; WIDTH: 1200; HEIGHT=400;">
				<table width="<%=sWidth%>"  border="1" cellpadding="0" cellspacing="0" bgcolor="#EFECE5" align="center" bordercolor="black" bordercolordark="white" bordercolorlight="black">
					
				<%
				firstLine = firstLine & "<td align='center' class='TDCont' rowspan="& rowspan &" width='150'>계</td>"
				firstLine = "<tr bgcolor='#EEF6FF'>" & sSero & " rowspan="& rowspan &">"& sSero1 & firstLine &"</tr>"
				secondLine = "<tr bgcolor='#EEF6FF'>" & secondLine &"</tr>"
				response.write firstLine
				response.write secondLine
	
				''-----세로항목 뿌리기
				sCOLNM = "CALLCLASS_B"
				sCOLCD = ""
	
				sSQL = " EXEC SP_SUM_BY_HISTORY_BCLASS " & sDepth & ",'" & FromDate & "','" & ToDate & "','TB_LIFECALLHISTORY','"&whereCD2 & "','"& whereCD1 & "','"&sCOLNM&"','Q','','"&sGroupBy&"','"&sSelectCol&"','"&sSelectCol1&"','"&sNullCol&"','"&CHANNELGB&"','"&JEONDOR&"','',''"
	
				'response.write sSQL
	
				set rsSUM = db.execute(sSQL)
				firstLine = ""
	
				Do Until rsSUM.eof
					sBG = "#ffffff"
					firstLine = ""
					'--------------키에 해당하는 값
					For i = 1 To sDepth
						sUser = rsSUM("col_"&i)
						'sCodeName = db_GetUserName(sUser)
	
						If IsNull(rsSUM("col_"&i)) Then
							sBG = "#EEF6FF"
							firstLine = firstLine & "<td align='center' class='TDCont' colspan='"&sDepth-i+1&"' WIDTH='200'>계</td>"
							Exit for
						Else
							If i = 1 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i),"","","","")
							ElseIf i = 2 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-1),rsSUM("col_"&i),"","","")
							ElseIf i = 3 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i),"","")
							ElseIf i = 4 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-3),rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i),"")
							Else
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-4),rsSUM("col_"&i-3),rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i))
							End if
							firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='200'>"&sCodeName&"</td>"
						End If
					Next
					'--------------가로항목의 summary
					For i = sDepth + 1 To rsSUM.Fields.count
	
						firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='200'>"&rsSUM("col_"&i)&"</td>"
	
					Next
					firstLine = "<tr bgcolor='"&sBG&"'>" & firstLine
					response.write firstLine & "</tr>"
					
					rsSUM.movenext
					'소계
				Loop
				%>
	
				</table>
				</div>
				
				<%
	
			elseif whereCD9 = "조치별" then'조치별 - 1depth
	
				rowspan = 2
				sSQL = "	select ACLASS, BCLASS, CCLASS, CLASSNAME from TB_ARMYINFO "
				sSQL = sSQL & "	where ACLASS = 'U' AND BCLASS IS NOT NULL AND CCLASS IS NULL"
				sSQL = sSQL & "	ORDER BY ACLASS, BCLASS, CCLASS "
				set RsCode = db.execute(sSQL)
	
				Do Until rsCode.eof
	
					sCode = RsCode("BCLASS")
					sCodeName = RsCode("CLASSNAME")
	
					'2DEPTH 값 찾기
					iCol = 0
					'secondLine = ""
					sCodeList = ""
					sSQL = "	select ACLASS, BCLASS, CCLASS, CLASSNAME from TB_ARMYINFO "
					sSQL = sSQL & "	where ACLASS = 'U' AND BCLASS = '" &sCode&"'  AND CCLASS IS NOT NULL"
					sSQL = sSQL & "	ORDER BY ACLASS, BCLASS, CCLASS "
					set RsCode1 = db.execute(sSQL)
	
					Do Until rsCode1.eof
	
						sCode = RsCode1("CCLASS")
						sCodeName = RsCode1("CLASSNAME")
	
						iCol = iCol + 1
						cols = cols + 1
						sSQL = "INSERT INTO TMP_CODE_VALUE VALUES ("&cols& ",'PROCESSGB_C','" & sCode & "')"
						db.execute(sSQL)
	
						If sCodeList = "" then
							sCodeList = sCode
						Else
							sCodeList = sCodeList & "|" & sCode
						End if
						secondLine = secondLine & "<td align='center' class='TDCont'  width='150'>" & sCodeName & "</td>"
						
						rsCode1.movenext
						'소계
					Loop
	
					If iCol = 0 Then
						sSQL = "INSERT INTO TMP_CODE_VALUE VALUES ("&cols& ",'PROCESSGB_B','" & sCode & "')"
						db.execute(sSQL)
						cols = cols + 1
						firstLine = firstLine & "<td align='center' class='TDCont' colspan="&iCol+1&" rowspan='2' width='150'>"&RsCode("CLASSNAME")&"</td>"
					Else
						sSQL = "INSERT INTO TMP_CODE_VALUE VALUES ("&cols& ",'PROCESSGB_C','" & sCodeList & "')"
						db.execute(sSQL)
						cols = cols + 1
						secondLine = secondLine & "<td align='center' class='TDCont'  width='150'>계</td>"
	
						firstLine = firstLine & "<td align='center' class='TDCont' colspan="&iCol+1&">"&RsCode("CLASSNAME")&"</td>"
					End if
					
					rsCode.movenext
					'소계
				Loop
				'총계
				sWidth = cols * 150
				%>
				
				<DIV style="OVERFLOW-X: auto;OVERFLOW-Y: auto; MARGIN: 0px 0px 0px 0px; WIDTH: 1200; HEIGHT=400;">
				<table width="<%=sWidth%>"  border="1" cellpadding="0" cellspacing="0" bgcolor="#EFECE5" align="center" bordercolor="black" bordercolordark="white" bordercolorlight="black">
					
				<%
				firstLine = firstLine & "<td align='center' class='TDCont' rowspan="& rowspan &" width='150'>계</td>"
				firstLine = "<tr bgcolor='#EEF6FF'>" & sSero & " rowspan="& rowspan &">"& sSero1 & firstLine &"</tr>"
				secondLine = "<tr bgcolor='#EEF6FF'>" & secondLine &"</tr>"
				response.write firstLine
				response.write secondLine
	
				''-----세로항목 뿌리기
				sCOLNM = "PROCESSGB_B"
				sCOLCD = ""
	
				sSQL = " EXEC SP_SUM_BY_HISTORY_BCLASS " & sDepth & ",'" & FromDate & "','" & ToDate & "','TB_LIFECALLHISTORY','"&whereCD2 & "','"& whereCD1 & "','"&sCOLNM&"','U','','"&sGroupBy&"','"&sSelectCol&"','"&sSelectCol1&"','"&sNullCol&"','"&CHANNELGB&"','"&JEONDOR&"','',''"
	
				'response.write sSQL
	
				set rsSUM = db.execute(sSQL)
				firstLine = ""
	
				Do Until rsSUM.eof
					sBG = "#ffffff"
					firstLine = ""
					'--------------키에 해당하는 값
					For i = 1 To sDepth
						sUser = rsSUM("col_"&i)
						'sCodeName = db_GetUserName(sUser)
	
						If IsNull(rsSUM("col_"&i)) Then
							sBG = "#EEF6FF"
							firstLine = firstLine & "<td align='center' class='TDCont' colspan='"&sDepth-i+1&"' WIDTH='200'>계</td>"
							Exit for
						Else
							If i = 1 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i),"","","","")
							ElseIf i = 2 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-1),rsSUM("col_"&i),"","","")
							ElseIf i = 3 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i),"","")
							ElseIf i = 4 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-3),rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i),"")
							Else
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-4),rsSUM("col_"&i-3),rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i))
							End if
							firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='200'>"&sCodeName&"</td>"
						End If
					Next
					'--------------가로항목의 summary
					For i = sDepth + 1 To rsSUM.Fields.count
	
						firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='200'>"&rsSUM("col_"&i)&"</td>"
	
					Next
					firstLine = "<tr bgcolor='"&sBG&"'>" & firstLine
					response.write firstLine & "</tr>"
					
					rsSUM.movenext
					'소계
				Loop
				%>
	
				</table>
				</div>
				
				<%
	
			elseif whereCD9 = "날씨별" then'날씨 - 1depth
				
				rowspan = 1
				sSQL = "	select CODE, CODENAME from TB_CODE "
				sSQL = sSQL & "	where CODEGROUP = 'C11' AND USEYN = 'Y'"
				sSQL = sSQL & "	ORDER BY CODE "
				set RsCode = db.execute(sSQL)
	
				Do Until rsCode.eof
	
					sCode = RsCode("CODE")
					sCodeName = RsCode("CODENAME")
	
					'2DEPTH 값 찾기
					iCol = iCol + 1
					cols = cols + 1
	
					firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='200'>"&RsCode("CODENAME")&"</td>"
					
					rsCode.movenext
					'소계
				Loop
				cols = cols + 1
				firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='200'>계</td>"
				'총계
				sWidth = cols * 200
				%>
				
				<DIV style="OVERFLOW-X: auto;OVERFLOW-Y: auto; MARGIN: 0px 0px 0px 0px; WIDTH: 1200; HEIGHT='400';">
				<table width="<%=sWidth%>"  border="1" cellpadding="0" cellspacing="0" bgcolor="#EFECE5" align="center" bordercolor="black" bordercolordark="white" bordercolorlight="black">
					
				<%
				firstLine = "<tr bgcolor='#EEF6FF'>" & sSero & " rowspan="& rowspan &">"& sSero1 & firstLine &"</tr>"
				response.write firstLine
	
				sCOLNM = "WEATHER"
				sCOLCD = "C11"
	
				sSQL = " EXEC SP_SUM_BY_CODE " & sDepth & ",'" & FromDate & "','" & ToDate & "','TB_LIFECALLHISTORY','"&whereCD2 & "','"& whereCD1 & "','"&sCOLNM&"','"&sCOLCD&"','"&sGroupBy&"','"&sSelectCol&"','"&sSelectCol1&"','"&sNullCol&"','"&CHANNELGB&"','"&JEONDOR&"','',''"
	
				'response.write sSQL
	
				set rsSUM = db.execute(sSQL)
				firstLine = ""
	
				Do Until rsSUM.eof
					sBG = "#ffffff"
					firstLine = ""
					For i = 1 To sDepth
						sUser = rsSUM("col_"&i)
						If IsNull(rsSUM("col_"&i)) Then
							sBG = "#EEF6FF"
							firstLine = firstLine & "<td align='center' class='TDCont' colspan='"&sDepth-i+1&"' WIDTH='200'>계</td>"
							Exit for
						Else
							If i = 1 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i),"","","","")
							ElseIf i = 2 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-1),rsSUM("col_"&i),"","","")
							ElseIf i = 3 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i),"","")
							ElseIf i = 4 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-3),rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i),"")
							Else
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-4),rsSUM("col_"&i-3),rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i))
							End if
	
							firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='200'>"&sCodeName&"</td>"
						End If
					Next
	
					For i = sDepth + 1 To rsSUM.Fields.count
	
						firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='200'>"&rsSUM("col_"&i)&"</td>"
	
					Next
					firstLine = "<tr bgcolor='"&sBG&"'>" & firstLine
					response.write firstLine & "</tr>"
					
					rsSUM.movenext
					'소계
				Loop
				%>
	
				</table>
				</div>
				
				<%
			elseif whereCD9 = "인지경로" then'인지경로 - 1depth
				
				rowspan = 1
				sSQL = "	select CODE, CODENAME from TB_CODE "
				sSQL = sSQL & "	where CODEGROUP = 'C10' AND USEYN = 'Y'"
				sSQL = sSQL & "	ORDER BY CODE "
				set RsCode = db.execute(sSQL)
	
				Do Until rsCode.eof
	
					sCode = RsCode("CODE")
					sCodeName = RsCode("CODENAME")
	
					'2DEPTH 값 찾기
					iCol = iCol + 1
					cols = cols + 1
	
					firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='150'>"&RsCode("CODENAME")&"</td>"
					
					rsCode.movenext
					'소계
				Loop
				cols = cols + 1
				firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='150'>계</td>"
				'총계
				sWidth = cols * 150
				%>
				
				<DIV style="OVERFLOW-X: auto;OVERFLOW-Y: auto; MARGIN: 0px 0px 0px 0px; WIDTH: 1200; HEIGHT='400';">
				<table width="<%=sWidth%>"  border="1" cellpadding="0" cellspacing="0" bgcolor="#EFECE5" align="center" bordercolor="black" bordercolordark="white" bordercolorlight="black">
					
				<%
				firstLine = "<tr bgcolor='#EEF6FF'>" & sSero & " rowspan="& rowspan &">"& sSero1 & firstLine &"</tr>"
				response.write firstLine
	
				sCOLNM = "CALLFLAG"
				sCOLCD = "C10"
	
				sSQL = " EXEC SP_SUM_BY_CODE " & sDepth & ",'" & FromDate & "','" & ToDate & "','TB_LIFECALLHISTORY','"&whereCD2 & "','"& whereCD1 & "','"&sCOLNM&"','"&sCOLCD&"','"&sGroupBy&"','"&sSelectCol&"','"&sSelectCol1&"','"&sNullCol&"','"&CHANNELGB&"','"&JEONDOR&"','',''"
	
				'response.write sSQL
	
				set rsSUM = db.execute(sSQL)
				firstLine = ""
	
				Do Until rsSUM.eof
					sBG = "#ffffff"
					firstLine = ""
					For i = 1 To sDepth
						sUser = rsSUM("col_"&i)
						'sCodeName = db_GetUserName(sUser)
	
						If IsNull(rsSUM("col_"&i)) Then
							sBG = "#EEF6FF"
							firstLine = firstLine & "<td align='center' class='TDCont' colspan='"&sDepth-i+1&"' WIDTH='200'>계</td>"
							Exit for
						Else
							If i = 1 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i),"","","","")
							ElseIf i = 2 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-1),rsSUM("col_"&i),"","","")
							ElseIf i = 3 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i),"","")
							ElseIf i = 4 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-3),rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i),"")
							Else
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-4),rsSUM("col_"&i-3),rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i))
							End if
							firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='200'>"&sCodeName&"</td>"
						End If
					Next
	
					For i = sDepth + 1 To rsSUM.Fields.count
	
						firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='200'>"&rsSUM("col_"&i)&"</td>"
	
					Next
					firstLine = "<tr bgcolor='"&sBG&"'>" & firstLine
					response.write firstLine & "</tr>"
					
					rsSUM.movenext
					'소계
				Loop
				%>
	
				</table>
				</div>
				
				<%
			elseif whereCD9 = "가족사항" then'가족사항 - 1depth
				
				rowspan = 1
				sSQL = "	select CODE, CODENAME from TB_CODE "
				sSQL = sSQL & "	where CODEGROUP = 'C12' AND USEYN = 'Y'"
				sSQL = sSQL & "	ORDER BY CODE "
				set RsCode = db.execute(sSQL)
	
				Do Until rsCode.eof
	
					sCode = RsCode("CODE")
					sCodeName = RsCode("CODENAME")
	
					'2DEPTH 값 찾기
					iCol = iCol + 1
					cols = cols + 1
	
					firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='150'>"&RsCode("CODENAME")&"</td>"
					
					rsCode.movenext
					'소계
				Loop
				cols = cols + 1
				firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='150'>계</td>"
				'총계
				sWidth = cols * 150
				%>
				
				<DIV style="OVERFLOW-X: auto;OVERFLOW-Y: auto; MARGIN: 0px 0px 0px 0px; WIDTH: 1200; HEIGHT='400';">
				<table width="<%=sWidth%>"  border="1" cellpadding="0" cellspacing="0" bgcolor="#EFECE5" align="center" bordercolor="black" bordercolordark="white" bordercolorlight="black">
					
				<%
				firstLine = "<tr bgcolor='#EEF6FF'>" & sSero & " rowspan="& rowspan &">"& sSero1 & firstLine &"</tr>"
				response.write firstLine
	
				sCOLNM = "FAMILYGB"
				sCOLCD = "C12"
	
				sSQL = " EXEC SP_SUM_BY_CODE " & sDepth & ",'" & FromDate & "','" & ToDate & "','TB_LIFECALLHISTORY','"&whereCD2 & "','"& whereCD1 & "','"&sCOLNM&"','"&sCOLCD&"','"&sGroupBy&"','"&sSelectCol&"','"&sSelectCol1&"','"&sNullCol&"','"&CHANNELGB&"','"&JEONDOR&"','',''"
	
				'response.write sSQL
	
				set rsSUM = db.execute(sSQL)
				firstLine = ""
	
				Do Until rsSUM.eof
					sBG = "#ffffff"
					firstLine = ""
					For i = 1 To sDepth
						sUser = rsSUM("col_"&i)
						'sCodeName = db_GetUserName(sUser)
	
						If IsNull(rsSUM("col_"&i)) Then
							sBG = "#EEF6FF"
							firstLine = firstLine & "<td align='center' class='TDCont' colspan='"&sDepth-i+1&"' WIDTH='200'>계</td>"
							Exit for
						Else
							If i = 1 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i),"","","","")
							ElseIf i = 2 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-1),rsSUM("col_"&i),"","","")
							ElseIf i = 3 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i),"","")
							ElseIf i = 4 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-3),rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i),"")
							Else
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-4),rsSUM("col_"&i-3),rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i))
							End if
							firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='200'>"&sCodeName&"</td>"
						End If
					Next
	
					For i = sDepth + 1 To rsSUM.Fields.count
	
						firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='200'>"&rsSUM("col_"&i)&"</td>"
	
					Next
					firstLine = "<tr bgcolor='"&sBG&"'>" & firstLine
					response.write firstLine & "</tr>"
					
					rsSUM.movenext
					'소계
				Loop
	
				%>
	
				</table>
				</div>
				
				<%
			elseif whereCD9 = "원인제공자" then'원인제공자 - 1depth
				
				rowspan = 1
				sSQL = "	select ACLASS, BCLASS, CCLASS, CLASSNAME from TB_ARMYINFO "
				sSQL = sSQL & "	where ACLASS = 'R' AND BCLASS IS NOT NULL AND CCLASS IS NULL"
				sSQL = sSQL & "	ORDER BY ACLASS, BCLASS, CCLASS "
				set RsCode = db.execute(sSQL)
	
				Do Until rsCode.eof
	
					sCode = RsCode("BCLASS")
					sCodeName = RsCode("CLASSNAME")
	
					cols = cols + 1
	
					firstLine = firstLine & "<td align='center' class='TDCont' colspan=1 width='150'>"&RsCode("CLASSNAME")&"</td>"
	
					rsCode.movenext
					'소계
				Loop
				'총계
				cols = cols + 1
				firstLine = firstLine & "<td align='center' class='TDCont' colspan=1 width='150'>계</td>"
				sWidth = cols * 100
				%>
				
				<DIV style="OVERFLOW-X: auto;OVERFLOW-Y: auto; MARGIN: 0px 0px 0px 0px; WIDTH: 1200; HEIGHT=400;">
				<table width="<%=sWidth%>"  border="1" cellpadding="0" cellspacing="0" bgcolor="#EFECE5" align="center" bordercolor="black" bordercolordark="white" bordercolorlight="black">
					
				<%
	
				firstLine = "<tr bgcolor='#EEF6FF'>" & sSero & " rowspan="& rowspan &">"& sSero1 & firstLine &"</tr>"
				response.write firstLine
	
				'---- 원인제공자
				sCOLNM = "CALLKIND_B"
				sCOLCD = "R"
	
				sSQL = " EXEC SP_SUM_BY_HISTORY_ACLASS " & sDepth & ",'" & FromDate & "','" & ToDate & "','TB_LIFECALLHISTORY','"&whereCD2 & "','"& whereCD1 & "','"&sCOLNM&"','"&sCOLCD&"','"&sGroupBy&"','"&sSelectCol&"','"&sSelectCol1&"','"&sNullCol&"','"&CHANNELGB&"','"&JEONDOR&"','',''"
	
				'response.write sSQL
	
				set rsSUM = db.execute(sSQL)
				firstLine = ""

				Do Until rsSUM.eof
					sBG = "#ffffff"
					firstLine = ""
					For i = 1 To sDepth
						sUser = rsSUM("col_"&i)
						'sCodeName = db_GetUserName(sUser)
	
						If IsNull(rsSUM("col_"&i)) Then
							sBG = "#EEF6FF"
							firstLine = firstLine & "<td align='center' class='TDCont' colspan='"&sDepth-i+1&"' WIDTH='200'>계</td>"
							Exit for
						Else
							If i = 1 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i),"","","","")
							ElseIf i = 2 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-1),rsSUM("col_"&i),"","","")
							ElseIf i = 3 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i),"","")
							ElseIf i = 4 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-3),rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i),"")
							Else
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-4),rsSUM("col_"&i-3),rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i))
							End if
							firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='200'>"&sCodeName&"</td>"
						End If
					Next
	
					For i = sDepth + 1 To rsSUM.Fields.count
	
						firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='200'>"&rsSUM("col_"&i)&"</td>"
	
					Next
					firstLine = "<tr bgcolor='"&sBG&"'>" & firstLine
					response.write firstLine & "</tr>"
					
					rsSUM.movenext
					'소계
				Loop
				%>
	
				</table>
				</div>
				
				<%
			elseif whereCD9 = "상담관" then'상담관 - 1depth
	
				rowspan = 1
				sSQL = "	select distinct INCODE FROM TB_LIFECALLHISTORY"
				sSQL = sSQL & "	where JUBDATE >= '" & FROMDATE &"'"
				sSQL = sSQL & "	AND JUBDATE <= '" & TODATE &"'"
				if len(CHANNELGB1) > 0 or len(CHANNELGB2) > 0 or len(CHANNELGB3) > 0 or len(CHANNELGB4) > 0 then
					sSQL = sSQL & " and CHANNELGB in ('" & CHANNELGB1 & "','" & CHANNELGB2 & "','" & CHANNELGB3 & "','" & CHANNELGB4 & "') "
				end if
				set RsCode1 = db.execute(sSQL)
	
				Do Until rsCode1.eof
	
					sCode = RsCode1("INCODE")
					sCodeName= db_getUserName(RsCode1("INCODE"))
	
					'2DEPTH 값 찾기
					iCol = iCol + 1
					cols = cols + 1
	
					firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='200'>"& sCodeName&"</td>"
					rsCode1.movenext
	
					'소계
				Loop
				cols = cols + 1
				firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='200'>계</td>"
				'총계
				sWidth = cols * 200
				%>
				
				<DIV style="OVERFLOW-X: auto;OVERFLOW-Y: auto; MARGIN: 0px 0px 0px 0px; WIDTH: 1200; HEIGHT='400';">
				<table width="<%=sWidth%>"  border="1" cellpadding="0" cellspacing="0" bgcolor="#EFECE5" align="center" bordercolor="black" bordercolordark="white" bordercolorlight="black">
					
				<%
				firstLine = "<tr bgcolor='#EEF6FF'>" & sSero & " rowspan="& rowspan &">"& sSero1 & firstLine &"</tr>"
				response.write firstLine
	
				sCOLNM = "INCODE"
				sCOLCD = ""
	
				sSQL = " EXEC SP_SUM_BY_INCODE " & sDepth & ",'" & FromDate & "','" & ToDate & "','TB_LIFECALLHISTORY','"&whereCD2 & "','"& whereCD1 & "','"&sCOLNM&"','"&sCOLCD&"','"&sGroupBy&"','"&sSelectCol&"','"&sSelectCol1&"','"&sNullCol&"','"&CHANNELGB&"','"&JEONDOR&"','',''"
				'response.write sSQL
	
				set rsSUM = db.execute(sSQL)
				firstLine = ""
	
				Do Until rsSUM.eof
					sBG = "#ffffff"
					firstLine = ""
					For i = 1 To sDepth
						sUser = rsSUM("col_"&i)
						'sCodeName = db_GetUserName(sUser)
	
						If IsNull(rsSUM("col_"&i)) Then
							sBG = "#EEF6FF"
							firstLine = firstLine & "<td align='center' class='TDCont' colspan='"&sDepth-i+1&"' WIDTH='200'>계</td>"
							Exit for
						Else
							If i = 1 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i),"","","","")
							ElseIf i = 2 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-1),rsSUM("col_"&i),"","","")
							ElseIf i = 3 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i),"","")
							ElseIf i = 4 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-3),rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i),"")
							Else
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-4),rsSUM("col_"&i-3),rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i))
							End if
							firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='200'>"&sCodeName&"</td>"
						End If
					Next
	
					For i = sDepth + 1 To rsSUM.Fields.count
	
						firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='200'>"&rsSUM("col_"&i)&"</td>"
	
					Next
					firstLine = "<tr bgcolor='"&sBG&"'>" & firstLine
	
					firstLine = firstLine & "</tr>"
					response.write firstLine
					
					rsSUM.movenext
					'소계
				Loop
				%>
	
				</table>
				</div>
				
				<%
	
			elseif whereCD9 = "시간" then'시간 - 1depth
	
				rowspan = 1
	
				For i = 0 To 23
	
					sCode = i
					sCodeName  = i & "시"
	
					'2DEPTH 값 찾기
					iCol = iCol + 1
					cols = cols + 1
	
					firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='120'>"& sCodeName&"</td>"
				Next
				cols = cols + 1
				firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='120'>계</td>"
				sWidth = cols * 120
				%>
				
				<DIV style="OVERFLOW-X: auto;OVERFLOW-Y: auto; MARGIN: 0px 0px 0px 0px; WIDTH: 1200; HEIGHT='400';">
				<table width="<%=sWidth%>"  border="1" cellpadding="0" cellspacing="0" bgcolor="#EFECE5" align="center" bordercolor="black" bordercolordark="white" bordercolorlight="black">
					
				<%
				firstLine = "<tr bgcolor='#EEF6FF'>" & sSero & " rowspan="& rowspan &">"& sSero1 & firstLine &"</tr>"
				response.write firstLine
	
				'sGroupBy = "datepart(hour,JUBTIME)"
	
				'---- 소속1차
				sCOLNM = "datepart(hour,JUBTIME)"
				sCOLCD = ""

				sSQL = " EXEC SP_SUM_BY_HISTORY_HOUR " & sDepth & ",'" & FromDate & "','" & ToDate & "','TB_LIFECALLHISTORY','"&whereCD2 & "','"& whereCD1 & "','"&sCOLNM&"','"&sCOLCD&"','"&sGroupBy&"','"&sSelectCol&"','"&sSelectCol1&"','"&sNullCol&"','"&CHANNELGB&"','"&JEONDOR&"','',''"
	
				'response.write sSQL
	
				set rsSUM = db.execute(sSQL)
				firstLine = ""
	
				Do Until rsSUM.eof
					sBG = "#ffffff"
					firstLine = ""
					For i = 1 To sDepth
						sUser = rsSUM("col_"&i)
						'sCodeName = db_GetUserName(sUser)
	
						If IsNull(rsSUM("col_"&i)) Then
							sBG = "#EEF6FF"
							firstLine = firstLine & "<td align='center' class='TDCont' colspan='"&sDepth-i+1&"' WIDTH='120'>계</td>"
							Exit for
						Else
							If i = 1 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i),"","","","")
							ElseIf i = 2 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-1),rsSUM("col_"&i),"","","")
							ElseIf i = 3 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i),"","")
							ElseIf i = 4 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-3),rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i),"")
							Else
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-4),rsSUM("col_"&i-3),rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i))
							End if
							firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='120'>"&sCodeName&"</td>"
						End If
					Next
	
					For i = sDepth + 1 To rsSUM.Fields.count
	
						firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='120'>"&rsSUM("col_"&i)&"</td>"
	
					Next
					firstLine = "<tr bgcolor='"&sBG&"'>" & firstLine
					response.write firstLine & "</tr>"
					
					rsSUM.movenext
					'소계
				Loop
				%>
	
				</table>
				</div>
				
				<%
	
			elseif whereCD9 = "요일" then'요일 - 1depth
				
				rowspan = 1
	
				firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='150'>일</td>"
				firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='150'>월</td>"
				firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='150'>화</td>"
				firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='150'>수</td>"
				firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='150'>목</td>"
				firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='150'>금</td>"
				firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='150'>토</td>"
				cols = cols + 8
				firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='150'>계</td>"
				sWidth = cols * 150
				%>
				
				<DIV style="OVERFLOW-X: auto;OVERFLOW-Y: auto; MARGIN: 0px 0px 0px 0px; WIDTH: 1200; HEIGHT='400';">
				<table width="<%=sWidth%>"  border="1" cellpadding="0" cellspacing="0" bgcolor="#EFECE5" align="center" bordercolor="black" bordercolordark="white" bordercolorlight="black">
					
				<%
				firstLine = "<tr bgcolor='#EEF6FF'>" & sSero & " rowspan="& rowspan &">"& sSero1 & firstLine &"</tr>"
				response.write firstLine
	
				'---- 소속1차
				sCOLNM = "datepart(WEEKDAY,JUBTIME)"
				sCOLCD = ""
	
				sSQL = " EXEC SP_SUM_BY_HISTORY_WEEKDAY " & sDepth & ",'" & FromDate & "','" & ToDate & "','TB_LIFECALLHISTORY','"&whereCD2 & "','"& whereCD1 & "','"&sCOLNM&"','"&sCOLCD&"','"&sGroupBy&"','"&sSelectCol&"','"&sSelectCol1&"','"&sNullCol&"','"&CHANNELGB&"','"&JEONDOR&"','',''"
	
				'response.write sSQL
	
				set rsSUM = db.execute(sSQL)
				firstLine = ""
	
				Do Until rsSUM.eof
					sBG = "#ffffff"
					firstLine = ""
					For i = 1 To sDepth
						sUser = rsSUM("col_"&i)
						'sCodeName = db_GetUserName(sUser)
	
						If IsNull(rsSUM("col_"&i)) Then
							sBG = "#EEF6FF"
							firstLine = firstLine & "<td align='center' class='TDCont' colspan='"&sDepth-i+1&"' WIDTH='200'>계</td>"
							Exit for
						Else
							If i = 1 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i),"","","","")
							ElseIf i = 2 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-1),rsSUM("col_"&i),"","","")
							ElseIf i = 3 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i),"","")
							ElseIf i = 4 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-3),rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i),"")
							Else
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-4),rsSUM("col_"&i-3),rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i))
							End if
							firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='200'>"&sCodeName&"</td>"
						End If
					Next
	
					For i = sDepth + 1 To rsSUM.Fields.count
	
						firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='200'>"&rsSUM("col_"&i)&"</td>"
	
					Next
					firstLine = "<tr bgcolor='"&sBG&"'>" & firstLine
					response.write firstLine & "</tr>"
					
					rsSUM.movenext
					'소계
				Loop
	
				%>
	
				</table>
				</div>
				
				<%
			elseif whereCD9 = "통화시간" then'통화시간 - 1depth
	
				rowspan = 1
				firstLine = firstLine & "<td align='center' class='TDCont' width='150'>1분미만</td>"
				firstLine = firstLine & "<td align='center' class='TDCont' width='150'>1-5분</td>"
				firstLine = firstLine & "<td align='center' class='TDCont' width='150'>6-10분</td>"
				firstLine = firstLine & "<td align='center' class='TDCont' width='150'>11-20분</td>"
				firstLine = firstLine & "<td align='center' class='TDCont' width='150'>21-30분</td>"
				firstLine = firstLine & "<td align='center' class='TDCont' width='150'>31-40분</td>"
				firstLine = firstLine & "<td align='center' class='TDCont' width='150'>41-50분</td>"
				firstLine = firstLine & "<td align='center' class='TDCont' width='150'>51-60분</td>"
				firstLine = firstLine & "<td align='center' class='TDCont' width='150'>60분이상</td>"
				firstLine = firstLine & "<td align='center' class='TDCont' width='150'>계</td>"
				cols = cols + 10
				sWidth = cols * 120
				%>
				
				<DIV style="OVERFLOW-X: auto;OVERFLOW-Y: auto; MARGIN: 0px 0px 0px 0px; WIDTH: 1200; HEIGHT='400';">
				<table width="<%=sWidth%>"  border="1" cellpadding="0" cellspacing="0" bgcolor="#EFECE5" align="center" bordercolor="black" bordercolordark="white" bordercolorlight="black">
					
				<%
				firstLine = "<tr bgcolor='#EEF6FF'>" & sSero & " rowspan="& rowspan &">"& sSero1 & firstLine &"</tr>"
				response.write firstLine
	
				'---- 소속1차
				sCOLNM = "CALLTIME"
				sCOLCD = ""

				sSQL = " EXEC SP_SUM_BY_HISTORY_CALLTIME " & sDepth & ",'" & FromDate & "','" & ToDate & "','TB_LIFECALLHISTORY','"&whereCD2 & "','"& whereCD1 & "','"&sCOLNM&"','"&sCOLCD&"','"&sGroupBy&"','"&sSelectCol&"','"&sSelectCol1&"','"&sNullCol&"','"&CHANNELGB&"','"&JEONDOR&"','',''"
	
				'response.write sSQL
	
				set rsSUM = db.execute(sSQL)
				firstLine = ""
	
				Do Until rsSUM.eof
					sBG = "#ffffff"
					firstLine = ""
					For i = 1 To sDepth
						sUser = rsSUM("col_"&i)
						'sCodeName = db_GetUserName(sUser)
	
						If IsNull(rsSUM("col_"&i)) Then
							sBG = "#EEF6FF"
							firstLine = firstLine & "<td align='center' class='TDCont' colspan='"&sDepth-i+1&"' WIDTH='200'>계</td>"
							Exit for
						Else
							If i = 1 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i),"","","","")
							ElseIf i = 2 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-1),rsSUM("col_"&i),"","","")
							ElseIf i = 3 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i),"","")
							ElseIf i = 4 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-3),rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i),"")
							Else
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-4),rsSUM("col_"&i-3),rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i))
							End if
							firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='200'>"&sCodeName&"</td>"
						End If
					Next
	
					For i = sDepth + 1 To rsSUM.Fields.count
	
						firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='200'>"&rsSUM("col_"&i)&"</td>"
	
					Next
					firstLine = "<tr bgcolor='"&sBG&"'>" & firstLine
					response.write firstLine & "</tr>"
	
					rsSUM.movenext
					'소계
				Loop
				%>
	
				</table>
				</div>
				
				<%
			end if
			%>
			</td>
		</tr>
	</table>
			
<% End if %>
<!-- #include virtual="/Include/Bottom.asp" -->