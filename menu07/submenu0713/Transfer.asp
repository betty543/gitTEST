<!-- #include virtual="/Include/Top.asp" -->
<%
	'1. 파라미터 얻어오기


	SS_Login_Secgroup = SESSION("SS_Login_Secgroup")
	SS_Login_Grade = SESSION("SS_Login_Grade")
	SS_Login_CTIID = SESSION("SS_Login_CTIID")
	SS_Login_EXTNO = SESSION("SS_Login_EXTNO")
	SS_LoginID = SESSION("SS_LoginID")


	curPage = request("curPage")
	'3. 쿼리 실행
	'sql = db_getSqlWithPage(sql_tb, sql_index, sql_field, sql_where, sql_orderby, pageSize, curPage)
	sql = "	select	*	from	TB_TransferNo	order by	[DNIS]"
	set rs = db.execute(sql)

	do until rs.eof

		if rs("DNIS") = 1   Then 	'1차 착신전환
			R_TransferNo1 = rs("TransferNo")
			R_UserId1 = rs("UserId")
			R_OnPhone1 = rs("OnPhone")
			R_UpdateDate1 = rs("UpdateDate")
		elseif rs("DNIS") = 2 then	'2차 착신전환
			R_TransferNo2 = rs("TransferNo")
			R_UserId2 = rs("UserId")
			R_OnPhone2 = rs("OnPhone")
			R_UpdateDate2 = rs("UpdateDate")
		elseif rs("DNIS") = 3 then	'3차 착신전환
			R_TransferNo3 = rs("TransferNo")
			R_UserId3 = rs("UserId")
			R_OnPhone3 = rs("OnPhone")
			R_UpdateDate3 = rs("UpdateDate")
		elseif rs("DNIS") = 4 then	'4차 착신전환
			R_TransferNo4 = rs("TransferNo")
			R_UserId4 = rs("UserId")
			R_OnPhone4 = rs("OnPhone")
			R_UpdateDate4 = rs("UpdateDate")
	
		end if

		rs.movenext
	loop

	'4. Paging HTML 작성

%>

<div id='CalendarLayer' style='display:none;'><iframe name='CalendarFrame' src='/include/Calendar.asp' width='172' height='177' border='0' frameborder='0' scrolling='no'></iframe></div>

<table border="0" width="940" cellpadding="0" cellspacing="3" bgcolor="#EFECE5" align="center">
	<tr bgcolor="#FFFFFF">
		<td>
			<form name="inUpFrm" method="post" action="/Menu07/submenu0713/Transfer_InsUp.asp">
			<input type="hidden" name="jobGb">
			<input type="hidden" name="DNIS">
			<table width="100%" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">
			    <tr>
			        <td width="120" bgcolor="#EFEFEF" class="TDCont" align="center">착신 전환 순서</td>
			        <td width="120" bgcolor="#EFEFEF" class="TDCont" align="center">착신 번호</td>
			        <td width="80" bgcolor="#EFEFEF" class="TDCont" align="center">담당상담관</td>
					<td width="80" bgcolor="#EFEFEF" class="TDCont" align="center">통화중여부</td>
					<td width="80" bgcolor="#EFEFEF" class="TDCont" align="center">초기화</td>
			        <td bgcolor="#EFEFEF" class="TDCont" align="center">비고</td>
				</tr>
			    <tr>
			        <td width="120" bgcolor="#EFEFEF" class="TDCont" align="center">1차 착신전환</td>
			        <td width="120" bgcolor="#FFFFFF" class="TDCont" align="center">
					

					<%
							'======= 착신번호가져오기 ==================================================
							SqlCode = "SELECT code from tb_code"
							SqlCode = SqlCode& " WHERE codegroup =  'A13' and USEYN='Y' "
							
							SqlCode = SqlCode& " ORDER BY code"
							set RsCode = db.execute(SqlCode)
						%>
						<select name="TransferNo1" size="1" class="ComboFFFCE7">
							<option value="">선택</option>
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("code")
										CODENAME = RsCode("code")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &R_TransferNo1& "")%>
							<%
									RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
							%>
						</select>	


					</td>
			        <td width="80" bgcolor="#FFFFFF" class="TDCont" align="center">
					
						<%
							'======= 상담원 가져오기 ==================================================
							SqlCode = "SELECT USERID, USERNAME FROM TB_USERINFO"
							SqlCode = SqlCode& " WHERE USEYN='Y' "
							SqlCode = SqlCode& " AND SECGROUP = 'A'"
							if SS_Login_Grade <> "A" then
								'SqlCode = SqlCode& "	AND GRADE = '"&SS_Login_Grade&"'"
							end if
							if SS_Login_Secgroup = "A" then	'상담관일때는 내것만
								'SqlCode = SqlCode& "	AND USERID = '" &SS_LoginID&"'"
							end if
							
							SqlCode = SqlCode& " ORDER BY USEYN DESC, GRADE ASC, USERNAME ASC"
							set RsCode = db.execute(SqlCode)
						%>
						<select name="UserId1" size="1" class="ComboFFFCE7">
							<option value="">선택</option>
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("USERID")
										CODENAME = RsCode("USERNAME")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &R_UserId1& "")%>
							<%
									RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
							%>
						</select>							
					</td>
					<td width="80" bgcolor="#FFFFFF" class="TDCont" align="center"><%=R_OnPhone1%></td>
					<td width="80" bgcolor="#FFFFFF" class="TDCont" align="center"><% if R_OnPhone1 = "Y" then %><img src="/Images/Btn/BtnRegiAdd_GB9.GIF" style="cursor:hand;" align="absmiddle" onClick="javascript:fn_clear('1');"><% else %>&nbsp;<%end if%></td>
					<td rowspan='7' bgcolor="#FFFFFF"><b>착신번호는 [기초관리->종합코드관리->[A13]착신번호종류] </b><br>메뉴에서 등록/수정할 수 있습니다.</b>
					</td>
				</tr>
			  

			  

				<tr>
			        <td width="120" bgcolor="#EFEFEF" class="TDCont" align="center">2차 착신전환</td>
			        <td width="120" bgcolor="#FFFFFF" class="TDCont" align="center">					<%
							'======= 착신번호가져오기 ==================================================
							SqlCode = "SELECT code from tb_code"
							SqlCode = SqlCode& " WHERE codegroup =  'A13' and USEYN='Y' "
							
							SqlCode = SqlCode& " ORDER BY code"
							set RsCode = db.execute(SqlCode)
						%>
						<select name="TransferNo2" size="1" class="ComboFFFCE7">
							<option value="">선택</option>
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("code")
										CODENAME = RsCode("code")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &R_TransferNo2& "")%>
							<%
									RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
							%>
						</select>	</td>
			        <td width="80" bgcolor="#FFFFFF" class="TDCont" align="center">						<%
							'======= 상담원 가져오기 ==================================================
							SqlCode = "SELECT USERID, USERNAME FROM TB_USERINFO"
							SqlCode = SqlCode& " WHERE USEYN='Y' "
							SqlCode = SqlCode& " AND SECGROUP = 'A'"
							if SS_Login_Grade <> "A" then
								'SqlCode = SqlCode& "	AND GRADE = '"&SS_Login_Grade&"'"
							end if
							if SS_Login_Secgroup = "A" then	'상담관일때는 내것만
								'SqlCode = SqlCode& "	AND USERID = '" &SS_LoginID&"'"
							end if
							
							SqlCode = SqlCode& " ORDER BY USEYN DESC, GRADE ASC, USERNAME ASC"
							set RsCode = db.execute(SqlCode)
						%>
						<select name="UserId2" size="1" class="ComboFFFCE7">
							<option value="">선택</option>
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("USERID")
										CODENAME = RsCode("USERNAME")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &R_UserId2& "")%>
							<%
									RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
							%>
						</select>	</td>
					<td width="80" bgcolor="#FFFFFF" class="TDCont" align="center"><%=R_OnPhone2%></td>
					<td width="80" bgcolor="#FFFFFF" class="TDCont" align="center"><% if R_OnPhone2 = "Y" then %><img src="/Images/Btn/BtnRegiAdd_GB9.GIF" style="cursor:hand;" align="absmiddle" onClick="javascript:fn_clear('2');"><% else %>&nbsp;<%end if%></td>

				</tr>
			    <tr>
			        <td width="120" bgcolor="#EFEFEF" class="TDCont" align="center">3차 착신전환</td>
			        <td width="120" bgcolor="#FFFFFF" class="TDCont" align="center">					<%
							'======= 착신번호가져오기 ==================================================
							SqlCode = "SELECT code from tb_code"
							SqlCode = SqlCode& " WHERE codegroup =  'A13' and USEYN='Y' "
							
							SqlCode = SqlCode& " ORDER BY code"
							set RsCode = db.execute(SqlCode)
						%>
						<select name="TransferNo3" size="1" class="ComboFFFCE7">
							<option value="">선택</option>
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("code")
										CODENAME = RsCode("code")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &R_TransferNo3& "")%>
							<%
									RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
							%>
						</select>	</td>
			        <td width="80" bgcolor="#FFFFFF" class="TDCont" align="center">						<%
							'======= 상담원 가져오기 ==================================================
							SqlCode = "SELECT USERID, USERNAME FROM TB_USERINFO"
							SqlCode = SqlCode& " WHERE USEYN='Y' "
							SqlCode = SqlCode& " AND SECGROUP = 'A'"
							if SS_Login_Grade <> "A" then
								'SqlCode = SqlCode& "	AND GRADE = '"&SS_Login_Grade&"'"
							end if
							if SS_Login_Secgroup = "A" then	'상담관일때는 내것만
								'SqlCode = SqlCode& "	AND USERID = '" &SS_LoginID&"'"
							end if
							
							SqlCode = SqlCode& " ORDER BY USEYN DESC, GRADE ASC, USERNAME ASC"
							set RsCode = db.execute(SqlCode)
						%>
						<select name="UserId3" size="1" class="ComboFFFCE7">
							<option value="">선택</option>
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("USERID")
										CODENAME = RsCode("USERNAME")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &R_UserId3& "")%>
							<%
									RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
							%>
						</select>	</td>
					<td width="80" bgcolor="#FFFFFF" class="TDCont" align="center"><%=R_OnPhone3%></td>
					<td width="80" bgcolor="#FFFFFF" class="TDCont" align="center"><% if R_OnPhone3 = "Y" then %><img src="/Images/Btn/BtnRegiAdd_GB9.GIF" style="cursor:hand;" align="absmiddle" onClick="javascript:fn_clear('3');"><% else %>&nbsp;<%end if%></td>
				</tr>
			    <tr>
			        <td width="120" bgcolor="#EFEFEF" class="TDCont" align="center">4차 착신전환</td>
			        <td width="120" bgcolor="#FFFFFF" class="TDCont" align="center">					<%
							'======= 착신번호가져오기 ==================================================
							SqlCode = "SELECT code from tb_code"
							SqlCode = SqlCode& " WHERE codegroup =  'A13' and USEYN='Y' "
							
							SqlCode = SqlCode& " ORDER BY code"
							set RsCode = db.execute(SqlCode)
						%>
						<select name="TransferNo4" size="1" class="ComboFFFCE7">
							<option value="">선택</option>
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("code")
										CODENAME = RsCode("code")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &R_TransferNo4& "")%>
							<%
									RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
							%>
						</select>	</td>
			        <td width="80" bgcolor="#FFFFFF" class="TDCont" align="center">						<%
							'======= 상담원 가져오기 ==================================================
							SqlCode = "SELECT USERID, USERNAME FROM TB_USERINFO"
							SqlCode = SqlCode& " WHERE USEYN='Y' "
							SqlCode = SqlCode& " AND SECGROUP = 'A'"
							if SS_Login_Grade <> "A" then
								'SqlCode = SqlCode& "	AND GRADE = '"&SS_Login_Grade&"'"
							end if
							if SS_Login_Secgroup = "A" then	'상담관일때는 내것만
								'SqlCode = SqlCode& "	AND USERID = '" &SS_LoginID&"'"
							end if
							
							SqlCode = SqlCode& " ORDER BY USEYN DESC, GRADE ASC, USERNAME ASC"
							set RsCode = db.execute(SqlCode)
						%>
						<select name="UserId4" size="1" class="ComboFFFCE7">
							<option value="">선택</option>
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("USERID")
										CODENAME = RsCode("USERNAME")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &R_UserId4& "")%>
							<%
									RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
							%>
						</select>	</td>
					<td width="80" bgcolor="#FFFFFF" class="TDCont" align="center"><%=R_OnPhone4%></td>
					<td width="80" bgcolor="#FFFFFF" class="TDCont" align="center"><% if R_OnPhone4 = "Y" then %><img src="/Images/Btn/BtnRegiAdd_GB9.GIF" style="cursor:hand;" align="absmiddle" onClick="javascript:fn_clear('4');"><% else %>&nbsp;<%end if%></td>
				</tr>
			    <tr>
			        <td width="120" bgcolor="#EFEFEF" class="TDCont" align="center">5차 착신전환</td>
			        <td width="120" bgcolor="#FFFFFF" class="TDCont" align="center">					<%
							'======= 착신번호가져오기 ==================================================
							SqlCode = "SELECT code from tb_code"
							SqlCode = SqlCode& " WHERE codegroup =  'A13' and USEYN='Y' "
							
							SqlCode = SqlCode& " ORDER BY code"
							set RsCode = db.execute(SqlCode)
						%>
						<select name="TransferNo5" size="1" class="ComboFFFCE7" disabled>
							<option value="">선택</option>
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("code")
										CODENAME = RsCode("code")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &R_TransferNo5& "")%>
							<%
									RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
							%>
						</select>	</td>
			        <td width="80" bgcolor="#FFFFFF" class="TDCont" align="center">						<%
							'======= 상담원 가져오기 ==================================================
							SqlCode = "SELECT USERID, USERNAME FROM TB_USERINFO"
							SqlCode = SqlCode& " WHERE USEYN='Y' "
							SqlCode = SqlCode& " AND SECGROUP = 'A'"
							if SS_Login_Grade <> "A" then
								'SqlCode = SqlCode& "	AND GRADE = '"&SS_Login_Grade&"'"
							end if
							if SS_Login_Secgroup = "A" then	'상담관일때는 내것만
								'SqlCode = SqlCode& "	AND USERID = '" &SS_LoginID&"'"
							end if
							
							SqlCode = SqlCode& " ORDER BY USEYN DESC, GRADE ASC, USERNAME ASC"
							set RsCode = db.execute(SqlCode)
						%>
						<select name="UserId5" size="1" class="ComboFFFCE7" disabled>
							<option value="">선택</option>
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("USERID")
										CODENAME = RsCode("USERNAME")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &R_UserId5& "")%>
							<%
									RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
							%>
						</select>	</td>
					<td width="80" bgcolor="#FFFFFF" class="TDCont" align="center"><%=R_OnPhone5%></td>
					<td width="80" bgcolor="#FFFFFF" class="TDCont" align="center"><% if R_OnPhone5 = "Y" then %><img src="/Images/Btn/BtnRegiAdd_GB9.GIF" style="cursor:hand;" align="absmiddle" onClick="javascript:fn_clear('5');"><% else %>&nbsp;<%end if%></td>
				</tr>
			    <tr>
			        <td width="120" bgcolor="#EFEFEF" class="TDCont" align="center">6차 착신전환</td>
			        <td width="120" bgcolor="#FFFFFF" class="TDCont" align="center">					<%
							'======= 착신번호가져오기 ==================================================
							SqlCode = "SELECT code from tb_code"
							SqlCode = SqlCode& " WHERE codegroup =  'A13' and USEYN='Y' "
							
							SqlCode = SqlCode& " ORDER BY code"
							set RsCode = db.execute(SqlCode)
						%>
						<select name="TransferNo6" size="1" class="ComboFFFCE7" disabled>
							<option value="">선택</option>
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("code")
										CODENAME = RsCode("code")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &R_TransferNo6& "")%>
							<%
									RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
							%>
						</select>	</td>
			        <td width="80" bgcolor="#FFFFFF" class="TDCont" align="center">						<%
							'======= 상담원 가져오기 ==================================================
							SqlCode = "SELECT USERID, USERNAME FROM TB_USERINFO"
							SqlCode = SqlCode& " WHERE USEYN='Y' "
							SqlCode = SqlCode& " AND SECGROUP = 'A'"
							if SS_Login_Grade <> "A" then
								'SqlCode = SqlCode& "	AND GRADE = '"&SS_Login_Grade&"'"
							end if
							if SS_Login_Secgroup = "A" then	'상담관일때는 내것만
								'SqlCode = SqlCode& "	AND USERID = '" &SS_LoginID&"'"
							end if
							
							SqlCode = SqlCode& " ORDER BY USEYN DESC, GRADE ASC, USERNAME ASC"
							set RsCode = db.execute(SqlCode)
						%>
						<select name="UserId6" size="1" class="ComboFFFCE7" disabled>
							<option value="">선택</option>
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("USERID")
										CODENAME = RsCode("USERNAME")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &R_UserId6& "")%>
							<%
									RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
							%>
						</select>	</td>
					<td width="80" bgcolor="#FFFFFF" class="TDCont" align="center"><%=R_OnPhone6%></td>
					<td width="80" bgcolor="#FFFFFF" class="TDCont" align="center"><% if R_OnPhone6 = "Y" then %><img src="/Images/Btn/BtnRegiAdd_GB9.GIF" style="cursor:hand;" align="absmiddle" onClick="javascript:fn_clear('6');"><% else %>&nbsp;<%end if%></td>
				</tr>
			    <tr>
			        <td width="120" bgcolor="#EFEFEF" class="TDCont" align="center">7차 착신전환</td>
			        <td width="120" bgcolor="#FFFFFF" class="TDCont" align="center">					<%
							'======= 착신번호가져오기 ==================================================
							SqlCode = "SELECT code from tb_code"
							SqlCode = SqlCode& " WHERE codegroup =  'A13' and USEYN='Y' "
							
							SqlCode = SqlCode& " ORDER BY code"
							set RsCode = db.execute(SqlCode)
						%>
						<select name="TransferNo7" size="1" class="ComboFFFCE7" disabled>
							<option value="">선택</option>
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("code")
										CODENAME = RsCode("code")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &R_TransferNo7& "")%>
							<%
									RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
							%>
						</select>	</td>
			        <td width="80" bgcolor="#FFFFFF" class="TDCont" align="center">						<%
							'======= 상담원 가져오기 ==================================================
							SqlCode = "SELECT USERID, USERNAME FROM TB_USERINFO"
							SqlCode = SqlCode& " WHERE USEYN='Y' "
							SqlCode = SqlCode& " AND SECGROUP = 'A'"
							if SS_Login_Grade <> "A" then
								'SqlCode = SqlCode& "	AND GRADE = '"&SS_Login_Grade&"'"
							end if
							if SS_Login_Secgroup = "A" then	'상담관일때는 내것만
								'SqlCode = SqlCode& "	AND USERID = '" &SS_LoginID&"'"
							end if
							
							SqlCode = SqlCode& " ORDER BY USEYN DESC, GRADE ASC, USERNAME ASC"
							set RsCode = db.execute(SqlCode)
						%>
						<select name="UserId7" size="1" class="ComboFFFCE7" disabled>
							<option value="">선택</option>
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("USERID")
										CODENAME = RsCode("USERNAME")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &R_UserId7& "")%>
							<%
									RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
							%>
						</select>	</td>
					<td width="80" bgcolor="#FFFFFF" class="TDCont" align="center"><%=R_OnPhone7%></td>
					<td width="80" bgcolor="#FFFFFF" class="TDCont" align="center"><% if R_OnPhone7 = "Y" then %><img src="/Images/Btn/BtnRegiAdd_GB9.GIF" style="cursor:hand;" align="absmiddle" onClick="javascript:fn_clear('4');"><% else %>&nbsp;<%end if%></td>
				</tr>
				</table>

			</form>
			<table border="0" cellspacing="0" width="100%" align="center">
				<tr height="30">
					<td align="center">
						<img src="/Images/Btn/BtnSubmit.gif" style="cursor:hand;" align="absmiddle" onClick="javascript:fn_inup(document.inUpFrm);">
						<img src="/Images/Btn/BtnReset.gif" style="cursor:hand;" align="absmiddle" onClick="javascript:fn_reset();">
					</td>
				</tr>
			</table>	
		</td>
	</tr>
</table>

<table width="940" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
<!-- #include virtual="/Include/Bottom.asp" -->


<script>
function fn_inup(inUpFrm) {

	//필수입력값 체크
	if ( inUpFrm.TransferNo1.value == '' || inUpFrm.TransferNo1.value.length < 4 )
	{
		alert('1차 착신번호를 정확히 입력하십시오!');
		inUpFrm.TransferNo1.focus();
		return;
	}
	if ( inUpFrm.TransferNo2.value == '' || inUpFrm.TransferNo2.value.length < 4 )
	{
		alert('2차 착신번호를 정확히 입력하십시오!');
		inUpFrm.TransferNo2.focus();
		return;
	}
	if ( inUpFrm.TransferNo3.value == '' || inUpFrm.TransferNo3.value.length < 4 )
	{
		alert('3차 착신번호를 정확히 입력하십시오!');
		inUpFrm.TransferNo3.focus();
		return;
	}
	if ( inUpFrm.TransferNo4.value == '' || inUpFrm.TransferNo4.value.length < 4 )
	{
		alert('4차 착신번호를 정확히 입력하십시오!');
		inUpFrm.TransferNo4.focus();
		return;
	}
	
		document.all.jobGb.value = '';
		document.all.DNIS.value = '';

//	if ( inUpFrm.FinishTime8.value == '' || inUpFrm.FinishTime8.value.length != 4 )
//	{
//		alert('법정공휴일의 종료시각을 숫자4자리로 정확히 입력하십시오!');
//		inUpFrm.FinishTime8.focus();
//		return;
//	}


	if(confirm("변경된 값을 저장하시겠습니까?"))
		inUpFrm.submit();
	else
		return;
}

function fn_clear(arg0){
	
	if(confirm(arg0+"차 전화의 통화중 상태를 변경하시겠습니까?"))
	{
		document.all.jobGb.value = 'C';
		document.all.DNIS.value = arg0;
		inUpFrm.submit();
	}
	else
		return;




}

function fn_reset() {

		document.all.jobGb.value = '';
		document.all.DNIS.value = '';

		document.inUpFrm.TransferNo1.value="<%=R_TransferNo1%>";
		document.inUpFrm.UserId1.value="<%=R_UserId1%>";

		document.inUpFrm.TransferNo2.value="<%=R_TransferNo2%>";
		document.inUpFrm.UserId2.value="<%=R_UserId2%>";

		document.inUpFrm.TransferNo3.value="<%=R_TransferNo3%>";
		document.inUpFrm.UserId3.value="<%=R_UserId3%>";

		document.inUpFrm.TransferNo4.value="<%=R_TransferNo4%>";
		document.inUpFrm.UserId4.value="<%=R_UserId4%>";


/*		document.inUpFrm.StartTime2.value="<%=sStartTime2%>";
		document.inUpFrm.FinishTime2.value="<%=sFinishTime2%>";

		document.inUpFrm.StartTime3.value="<%=sStartTime3%>";
		document.inUpFrm.FinishTime3.value="<%=sFinishTime3%>";

		document.inUpFrm.StartTime4.value="<%=sStartTime4%>";
		document.inUpFrm.FinishTime4.value="<%=sFinishTime4%>";

		document.inUpFrm.StartTime5.value="<%=sStartTime5%>";
		document.inUpFrm.FinishTime5.value="<%=sFinishTime5%>";

		document.inUpFrm.StartTime6.value="<%=sStartTime6%>";
		document.inUpFrm.FinishTime6.value="<%=sFinishTime6%>";

		document.inUpFrm.StartTime7.value="<%=sStartTime7%>";
		document.inUpFrm.FinishTime7.value="<%=sFinishTime7%>";

		document.inUpFrm.StartTime8.value="<%=sStartTime8%>";
		document.inUpFrm.FinishTime8.value="<%=sFinishTime8%>";
*/
		return;
}
</script>