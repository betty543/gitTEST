<!-- #include virtual="/Include/Top.asp" -->

<%
'####### �Ķ���� ##################################################################################
QueryYN = request("QueryYN")
FromDate = request("FromDate")
ToDate = request("ToDate")
whereCD3 = Trim(request("whereCD3"))
whereCD2 = Trim(request("whereCD2"))
whereCD7 = Trim(request("whereCD7"))

dim	vtot(100)

If QueryYN = "" Then whereCD3 = "1" End if
if FromDate = "" then FromDate = left(Date(),7) & "-01" end If
if ToDate = "" then ToDate = date() end If

CHANNELGB1 = request("CHANNELGB1")
CHANNELGB2 = request("CHANNELGB2")
CHANNELGB3 = request("CHANNELGB3")
CHANNELGB4 = request("CHANNELGB4")
MAN = request("MAN")
WOMAN = request("WOMAN")

pageWHERE = "QueryYN=" & QueryYN & "&FromDate=" & FromDate & "&ToDate=" & ToDate & "&whereCD2=" & whereCD2 &"&whereCD3=" & whereCD3 & "&whereCD7=" & whereCD7
pageWHERE = pageWHERE & "&channelGb1=" & CHANNELGB1 & "&channelGb2=" & CAHNNELGB2 & "&channelGb3=" & CHANNELGB3 & "&channelGb4=" & CHANNELGB4 & "&MAN="&MAN& "&WOMAN="&WOMAN
%>

<div id='CalendarLayer' style='display:none;'><iframe name='CalendarFrame' src='/include/Calendar.asp' width='172' height='177' border='0' frameborder='0' scrolling='no'></iframe></div>

<script>
	function fn_Search() {
		document.inUpFrm.QueryYN.value = "Y";
		document.inUpFrm.submit();
	}
	
	function fn_Xls() {
		location.href="list01_Xls.asp?<%=pageWHERE%>"
	}
</script>

<table border="0" width="1200" cellpadding="0" cellspacing="3" bgcolor="#EFECE5" align="center">
	<tr bgcolor="#FFFFFF">
		<td>
			
			<form name="inUpFrm" method="post" action="<%=Menu_2nd%>" onsubmit="return fn_Search(this);" style="margin:0">
				<input type="hidden" name="QueryYN" value="<%=QueryYN%>">
				
				<table width="100%" border="0" cellspacing="1" cellpadding="0" style="border:#E1DED6 solid 1px">
					<tr>
						<td class="TDCont" colspan="2">
							��ȭ���� : 
							<input type="checkbox" name="CHANNELGB1" <% if CHANNELGB1 = "0179" then %>checked<% end if %> value="0179" class="none" />080����
							<input type="checkbox" name="CHANNELGB2" <% if CHANNELGB2 = "13031" then %>checked<% end if %> value="13031" class="none"  >1303����
							<input type="checkbox" name="CHANNELGB3" <% if CHANNELGB3 = "13032" then %>checked<% end if %> value="13032" class="none"  >1303������
							<input type="checkbox" name="CHANNELGB4" <% if CHANNELGB4 = "13033" then %>checked<% end if %> value="13033" class="none"  >1303������
							<!--<input type="checkbox" name="CHANNELGB4" <% if CHANNELGB4 = "CYBER" then %>checked<% end if %> value="CYBER" class="none"  >���̹�-->
						</td>
						<td class="TDCont" colspan="2">
							��&nbsp;&nbsp;&nbsp;&nbsp;�� : 
							<input type="checkbox" name="MAN" <% if MAN = "Y" then %>checked<% end if %> value="Y" class="none"  >����
							<input type="checkbox" name="WOMAN" <% if WOMAN = "Y" then %>checked<% end if %> value="Y" class="none"  >����
							<!--<input type="checkbox" name="CHANNELGB4" <% if CHANNELGB4 = "CYBER" then %>checked<% end if %> value="CYBER" class="none"  >���̹�-->
						</td>

					</tr>
					<tr>
						<td class="TDCont">��ȸ�Ⱓ :
							<input value="<%=FromDate%>" name="FromDate" readonly type="text" size="10" onfocus="setFocusColor(this); new CalendarFrame.Calendar(this);" onblur="setOutColor(this);"> ~
							<input value="<%=ToDate%>" name="ToDate" readonly type="text" size="10" onfocus="setFocusColor(this); new CalendarFrame.Calendar(this);" onblur="setOutColor(this);">
						</td>
						<td class="TDCont" COLSPAN = '6'>������� :

							<%
							'======= ó������ �ڵ� �������� ==================================================
							SqlCode = "SELECT BCLASS CODE, CLASSNAME CODENAME FROM TB_ARMYINFO"
							SqlCode = SqlCode& " WHERE ACLASS = 'Q' AND BCLASS IS NOT NULL AND CCLASS IS NULL"
							SqlCode = SqlCode& " ORDER BY ACLASS"
							set RsCode = db.execute(SqlCode)

							Do Until rsCode.eof
								sCode = trim(RsCode("CODE"))
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

<% If QueryYN = "Y" Then %>

<%
chSql  =""
if len(CHANNELGB1) > 0 or len(CHANNELGB2) > 0 or len(CHANNELGB3) > 0 or len(CHANNELGB4) > 0 then
	chSql = " and CHANNELGB in ('" & CHANNELGB1 & "','" & CHANNELGB2 & "','" & CHANNELGB3 & "','" & CHANNELGB4 & "') "
end If
if len(whereCD2) > 0 then
	chSql = chSql & " and CHANNELGB_B in ('" & Replace(whereCD2,", ","','") & "') "
end If
'�����̸鼭.. �����̸�

If MAN = "Y" and WOMAN = "" Then
	chSql = chSql & " and SEXGB = '1' AND ( SOSOKGB_A IN ('A','B','C','D','E') OR LEVEL_B IN ('P01','P02') ) "	
ElseIf MAN = "" and WOMAN = "Y" Then
	chSql = chSql & " and SEXGB = '2'  AND ( SOSOKGB_A IN ('A','B','C','D','E') OR LEVEL_B IN ('P01','P02') ) "	
ElseIf  MAN = "Y" And WOMAN = "Y" Then
	chSql = chSql & " AND ( SOSOKGB_A IN ('A','B','C','D','E') OR LEVEL_B IN ('P01','P02') ) "	
End if





'response.write whereCD2

%>

	<table border="0" cellpadding="0" cellspacing="0" align="center">
		<tr>
			<td>
				<table width="1200"  border="1" cellpadding="0" cellspacing="0" bgcolor="#EFECE5" align="center" bordercolor="black" bordercolordark="white" bordercolorlight="black">
					<tr bgcolor='#EEF6FF'>
						<td align='center' class='TDCont' width='300'>080����</td>
						<td align='center' class='TDCont' width='300'>1303����</td>
						<td align='center' class='TDCont' width='300'>1303������</td>
						<td align='center' class='TDCont' width='300'>1303������</td>
						<td align='center' class='TDCont' width='300'>��</td>
					</tr>

					<%
					'�������
					ttot1 = 0
					ttot2 = 0
					ttot3 = 0
					ttot4 = 0
					ttot5 = 0
					ttot6 = 0
					ttot7 = 0
					ttot7 = 0
					ttot8 = 0
					ttot9 = 0
					ttot10 = 0
					ttot11 = 0

					SQL = "select * from ( "
					SQL = SQL & " 		SELECT	'01' gubun, count(*) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "' "
					SQL = SQL & "			AND	jubdate <= '" & ToDate & "' and CHANNELGB_B in ('Q01','Q03') " & chSql & " and CHANNELGB = '0179'" '���,
					SQL = SQL & "		union all "
					SQL = SQL & " 		SELECT	'02' gubun, count(*) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "' "
					SQL = SQL & "			AND	jubdate <= '" & ToDate & "'  and CHANNELGB_B in ('Q01','Q03') " & chSql & " and CHANNELGB = '13031'" '���,
					SQL = SQL & "		union all "
					SQL = SQL & " 		SELECT	'03' gubun, count(*) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "' "
					SQL = SQL & "			AND	jubdate <= '" & ToDate & "'  and CHANNELGB_B in ('Q01','Q03') " & chSql & " and CHANNELGB = '13032'" '���,
					SQL = SQL & "		union all "
					SQL = SQL & " 		SELECT	'04' gubun, count(*) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "' "
					SQL = SQL & "			AND	jubdate <= '" & ToDate & "'  and CHANNELGB_B in ('Q01','Q03') " & chSql & " and CHANNELGB = '13033'" '���,
					SQL = SQL & "		union all "
					SQL = SQL & " 		SELECT	'05' gubun, count(*) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "' "
					SQL = SQL & "			AND	jubdate <= '" & ToDate & "'  and CHANNELGB_B in ('Q01','Q03') " & chSql & " ) a order by gubun"
					
					'response.write SQL
					set Rs = db.execute(SQL)					

					tot1 = 0
					tot2 = 0
					tot3 = 0
					tot4 = 0
					tot5 = 0
					tot6 = 0
					tot7 = 0
					tot8 = 0
					tot9 = 0
					tot10 = 0
					tot11 = 0
					tot12 = 0
					tot13 = 0

					do until rs.eof
				
							
							if rs("gubun") = "01" then
								tot1 = rs("cnt")
							elseif rs("gubun") = "02" then
								tot2 = rs("cnt")
							elseif rs("gubun") = "03" then
								tot3 = rs("cnt")
							elseif rs("gubun") = "04" then
								tot4 = rs("cnt")
							elseif rs("gubun") = "05" then
								tot5 = rs("cnt")
							end if
				
							rs.movenext
							if rs.eof then
								exit do
							end if
					loop
						%>
						
						<tr bgcolor='#EEF6FF'>
							<td bgcolor='#ffffff' align='center' class='TDCont' width="300"><%=tot1%></td>
							<td bgcolor='#ffffff' align='center' class='TDCont' width="300"><%=tot2%></td>
							<td bgcolor='#ffffff' align='center' class='TDCont' width="300"><%=tot3%></td>
							<td bgcolor='#ffffff' align='center' class='TDCont' width="300"><%=tot4%></td>
							<td bgcolor='#ffffff' align='center' class='TDCont' width="300"><%=tot5%></td>	
	
						</tr>
						
						<%
						ttot1 = ttot1 + tot1
						ttot2 = ttot2 + tot2
						ttot3 = ttot3 + tot3
						ttot4 = ttot4 + tot4
						ttot5 = ttot5 + tot5
						ttot6 = ttot6 + tot6
						ttot7 = ttot7 + tot7
						ttot8 = ttot8 + tot8
						ttot9 = ttot9 + tot9 
						ttot10 = ttot10 + tot10
						ttot11 = ttot11 + tot11
						ttot12 = ttot12 + tot12
						ttot13 = ttot13 + tot13
						tot1 = 0
						tot2 = 0
						tot3 = 0
						tot4 = 0
						tot5 = 0
						tot6 = 0
						tot7 = 0
						tot7 = 0
						tot8 = 0
						tot9 = 0
						tot10 = 0
						tot11 = 0
						tot12 = 0
						tot13 = 0

					%>

				</table>
				
				<table width="100%" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>

				
				<table width="1200"  border="1" cellpadding="0" cellspacing="0" bgcolor="#EFECE5" align="center" bordercolor="black" bordercolordark="white" bordercolorlight="black">
					<tr height="30">
						<td align="left" bgcolor="#FFFFFF" class="TDCont" colspan="18">&nbsp;<img src="/Images/dot_01.gif" >&nbsp;<b><font color="#ff00ff"></font> 1. �������</b></td>
					</tr>
					<tr bgcolor='#EEF6FF'>
						<td align='center' class='TDCont'  width='150' rowspan='2'>����</td>
						<td align='center' class='TDCont' colspan='4'>��ȭ���</td>
						<td align='center' class='TDCont' colspan='4'>���̹����</td>
						<td align='center' class='TDCont' rowspan='2'>������ȭ</td>
						<td align='center' class='TDCont' rowspan='2'>ħ��(Ž��)<br>��ȭ</td>
						<td align='center' class='TDCont' rowspan='2'>���</td>
						<td align='center' class='TDCont' rowspan='2'>��Ÿ</td>
						<td align='center' class='TDCont' rowspan='2'>��</td>
					</tr>
					<tr bgcolor='#EEF6FF'>

						<td align='center' class='TDCont' >����ȭ</td>
						<td align='center' class='TDCont' >�Ϲ���ȭ</td>
						<td align='center' class='TDCont' >�̻�</td>
						<td align='center' class='TDCont' >��</td>

						<td align='center' class='TDCont' >��Ʈ���</td>
						<td align='center' class='TDCont' >���ͳ�</td>
						<td align='center' class='TDCont' >�̻�</td>
						<td align='center' class='TDCont' >��</td>
					</tr>
					
					<%
					'�������
					ttot1 = 0
					ttot2 = 0
					ttot3 = 0
					ttot4 = 0
					ttot5 = 0
					ttot6 = 0
					ttot7 = 0
					ttot7 = 0
					ttot8 = 0
					ttot9 = 0
					ttot10 = 0
					ttot11 = 0
					ttot12 = 0

					SQL = "select * from ( "
					SQL = SQL & " 		SELECT	'01' gubun, incode, count(incode) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "' "
					SQL = SQL & "			AND	jubdate <= '" & ToDate & "' and CHANNELGB_B = 'Q01' " & chSql & " group by incode" '���,
					SQL = SQL & "		union all "
					SQL = SQL & " 		SELECT	'02' gubun, incode, count(incode) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
					SQL = SQL & "			AND jubdate <= '" & ToDate & "' and CHANNELGB_B = 'Q01' AND CHANNELGB_C = 'Q01A' " & chSql & " group by incode" '����,
					SQL = SQL & "		union all "
					SQL = SQL & " 		SELECT	'03' gubun, incode, count(incode) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
					SQL = SQL & "			AND		jubdate <= '" & ToDate & "' and CHANNELGB_B = 'Q01' AND CHANNELGB_C = 'Q01C' " & chSql & " group by incode" '����,
					SQL = SQL & "		union all "
					SQL = SQL & " 		SELECT	'04' gubun, incode, count(incode) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
					SQL = SQL & "			AND		jubdate <= '" & ToDate & "' and CHANNELGB_B = 'Q01' AND CHANNELGB_C NOT IN ('Q01A','Q01C') " & chSql & " group by incode" '����,
					SQL = SQL & "		union all "
					SQL = SQL & " 		SELECT	'05' gubun, incode, count(incode) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
					SQL = SQL & "			AND jubdate <= '" & ToDate & "' and CHANNELGB_B = 'Q03' " & chSql & " group by incode" '���,
					SQL = SQL & "		union all "
					SQL = SQL & " 		SELECT	'06' gubun, incode, count(incode) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
					SQL = SQL & "			AND jubdate <= '" & ToDate & "' and CHANNELGB_B = 'Q03' AND CHANNELGB_C = 'Q03A' " & chSql & " group by incode" '����,
					SQL = SQL & "		union all "
					SQL = SQL & " 		SELECT	'07' gubun, incode, count(incode) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
					SQL = SQL & "			AND jubdate <= '" & ToDate & "' and CHANNELGB_B = 'Q03' AND CHANNELGB_C = 'Q03C' " & chSql & " group by incode" '����,
					SQL = SQL & "		union all "
					SQL = SQL & " 		SELECT	'08' gubun, incode, count(incode) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
					SQL = SQL & "			AND jubdate <= '" & ToDate & "' and CHANNELGB_B = 'Q03' AND CHANNELGB_C NOT IN ('Q03A','Q03C') " & chSql & "  group by incode" '����,
					SQL = SQL & "		union all "
					SQL = SQL & " 		SELECT	'09' gubun, incode, count(incode) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
					SQL = SQL & "			AND jubdate <= '" & ToDate & "' and CHANNELGB_B = 'Q05' " & chSql & " group by incode" '���,
					SQL = SQL & "		union all "
					SQL = SQL & " 		SELECT	'10' gubun, incode, count(incode) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
					SQL = SQL & "			AND jubdate <= '" & ToDate & "' and CHANNELGB_B = 'Q07' " & chSql & " group by incode" '���,
					SQL = SQL & "		union all "
					SQL = SQL & " 		SELECT	'11' gubun, incode, count(incode) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
					SQL = SQL & "			AND jubdate <= '" & ToDate & "' and CHANNELGB_B = 'Q09' " & chSql & " group by incode" '���,
					SQL = SQL & "	 	union all "
					SQL = SQL & "			SELECT	'12' gubun, incode, count(incode) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
					SQL = SQL & "			AND jubdate <= '" & ToDate & "' and CHANNELGB_B = 'Q99' " & chSql & " group by incode"
					SQL = SQL & "		union all "
					SQL = SQL & " 		SELECT	'13' gubun, incode, count(incode) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
					SQL = SQL & "			AND jubdate <= '" & ToDate & "' " & chSql & " group by incode	) a order by incode, gubun"
				
					set Rs = db.execute(SQL)

					tot1 = 0
					tot2 = 0
					tot3 = 0
					tot4 = 0
					tot5 = 0
					tot6 = 0
					tot7 = 0
					tot8 = 0
					tot9 = 0
					tot10 = 0
					tot11 = 0
					tot12 = 0
					tot13 = 0

					do until rs.eof
				
						incode = rs("incode")
						
						do until incode <> rs("incode")
							
							if rs("gubun") = "01" then
								tot1 = rs("cnt")
							elseif rs("gubun") = "02" then
								tot2 = rs("cnt")
							elseif rs("gubun") = "03" then
								tot3 = rs("cnt")
							elseif rs("gubun") = "04" then
								tot4 = rs("cnt")
							elseif rs("gubun") = "05" then
								tot5 = rs("cnt")
							elseif rs("gubun") = "06" then
								tot6 = rs("cnt")
							elseif rs("gubun") = "07" then
								tot7 = rs("cnt")
							elseif rs("gubun") = "08" then
								tot8 = rs("cnt")
							elseif rs("gubun") = "09" then
								tot9 = rs("cnt")
							elseif rs("gubun") = "10" then
								tot10 = rs("cnt")
							elseif rs("gubun") = "11" then
								tot11 = rs("cnt")
							elseif rs("gubun") = "12" then
								tot12 = rs("cnt")
							elseif rs("gubun") = "13" then
								tot13 = rs("cnt")
							end if
				
							rs.movenext
							if rs.eof then
								exit do
							end if
						loop
						%>
						
						<tr bgcolor='#EEF6FF'>
							<td align='center' class='TDCont'  width='150' ><%=db_getUserName(incode)%></td>

							<td bgcolor='#ffffff' align='center' class='TDCont' width="100"><%=tot2%></td>
							<td bgcolor='#ffffff' align='center' class='TDCont' width="100"><%=tot3%></td>
							<td bgcolor='#ffffff' align='center' class='TDCont' width="100"><%=tot4%></td>
							<td bgcolor='#ffffff' align='center' class='TDCont' width="100"><%=tot1%></td>

							<td bgcolor='#ffffff' align='center' class='TDCont' width="100"><%=tot6%></td>
							<td bgcolor='#ffffff' align='center' class='TDCont' width="100"><%=tot7%></td>

							<td bgcolor='#ffffff' align='center' class='TDCont' width="100"><%=tot8%></td>
							<td bgcolor='#ffffff' align='center' class='TDCont' width="100"><%=tot5%></td>
							<td bgcolor='#ffffff' align='center' class='TDCont' width="100"><%=tot9%></td>
							<td bgcolor='#ffffff' align='center' class='TDCont' width="100"><%=tot10%></td>
							<td bgcolor='#ffffff' align='center' class='TDCont' width="100"><%=tot11%></td>
							<td bgcolor='#ffffff' align='center' class='TDCont' width="100"><%=tot12%></td>
							<td bgcolor='#ffffff' align='center' class='TDCont' width="100"><%=tot13%></td>
						</tr>
						
						<%
						ttot1 = ttot1 + tot1
						ttot2 = ttot2 + tot2
						ttot3 = ttot3 + tot3
						ttot4 = ttot4 + tot4
						ttot5 = ttot5 + tot5
						ttot6 = ttot6 + tot6
						ttot7 = ttot7 + tot7
						ttot8 = ttot8 + tot8
						ttot9 = ttot9 + tot9 
						ttot10 = ttot10 + tot10
						ttot11 = ttot11 + tot11
						ttot12 = ttot12 + tot12
						ttot13 = ttot13 + tot13
						tot1 = 0
						tot2 = 0
						tot3 = 0
						tot4 = 0
						tot5 = 0
						tot6 = 0
						tot7 = 0
						tot7 = 0
						tot8 = 0
						tot9 = 0
						tot10 = 0
						tot11 = 0
						tot12 = 0
						tot13 = 0
						
						if rs.eof then
							exit do
						end if
					loop
					%>
					
					<tr bgcolor='#FFEEF9'>
						<td align='center' class='TDCont'  width='150' >�Ѱ�</td>

						<td align='center' class='TDCont' width="100"><%=ttot2%></td>
						<td align='center' class='TDCont' width="100"><%=ttot3%></td>
						<td align='center' class='TDCont' width="100"><%=ttot4%></td>
						<td align='center' class='TDCont' width="100"><%=ttot1%></td>

						<td align='center' class='TDCont' width="100"><%=ttot6%></td>
						<td align='center' class='TDCont' width="100"><%=ttot7%></td>
						<td align='center' class='TDCont' width="100"><%=ttot8%></td>
						<td align='center' class='TDCont' width="100"><%=ttot5%></td>
						<td align='center' class='TDCont' width="100"><%=ttot9%></td>
						<td align='center' class='TDCont' width="100"><%=ttot10%></td>
						<td align='center' class='TDCont' width="100"><%=ttot11%></td>
						<td align='center' class='TDCont' width="100"><%=ttot12%></td>
						<td align='center' class='TDCont' width="100"><%=ttot13%></td>
					</tr>
				</table>
				
				<%
				ttot1 = 0
				ttot2 = 0
				ttot3 = 0
				ttot4 = 0
				ttot5 = 0
				ttot6 = 0
				ttot7 = 0
				ttot7 = 0
				ttot8 = 0
				ttot9 = 0
				ttot10 = 0
				ttot11 = 0
				ttot12 = 0
				ttot13 = 0
				%>
				
				<table width="100%" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
				
				<table width="1200"  border="1" cellpadding="0" cellspacing="0" bgcolor="#EFECE5" align="center" bordercolor="black" bordercolordark="white" bordercolorlight="black">
					
					<%
					'������
					SQL = " SELECT	incode, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
					SQL = SQL & "	AND		jubdate <= '" & ToDate & "' " & chSql & " group by incode order by incode"
				
					set Rs = db.execute(SQL)
					
					tot1 = 0
					tot2 = 0
				
					firstLine = "<tr bgcolor='#EEF6FF'>"
					firstLine = firstLine & "<td align='center' class='TDCont'  width='150'>����</td>"
					secondLine = "<tr bgcolor='#ffffff'>"
					secondLine = secondLine & "<td align='center' class='TDCont'  width='150'>��</td>"
					
					do until rs.eof
	
						incode = rs("incode")
						tot2 = tot2 + 1
						firstLine = firstLine & "<td bgcolor='#EEF6FF' align='center' class='TDCont'  width='150'>"&db_getUserName(incode)&"</td>"
						secondLine = secondLine &"<td bgcolor='#ffffff' align='center' class='TDCont'  width='150'>"&rs("cnt")&"</td>"
						tot1 = tot1 + rs("cnt")
						
						rs.movenext
						if rs.eof then
							exit do
						end if
					loop
					
					firstLine = firstLine & "<td bgcolor='#EEF6FF' align='center' class='TDCont'  width='150'>��</td></tr>"
					secondLine = secondLine &"<td bgcolor='#ffffff'align='center' class='TDCont'  width='150'>"&tot1&"</td></tr>"
					%>
					
					<tr height="30">
						<td align="left" bgcolor="#FFFFFF" class="TDCont" colspan="<%=tot2+2%>">&nbsp;<img src="/Images/dot_01.gif" >&nbsp;<b><font color="#ff00ff"></font> 2. ������</b></td>
					</tr>
					
					<%
					response.write firstLine
					response.write secondLine
					%>
					
				</table>
				
				<!--��޺�-->
				<table width="100%" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
							
				<DIV style="OVERFLOW-X: auto;OVERFLOW-Y: auto; MARGIN: 0px 0px 0px 0px; WIDTH: 1200;">
					<table width="2500"  border="1" cellpadding="0" cellspacing="0" bgcolor="#EFECE5" align="center" bordercolor="black" bordercolordark="white" bordercolorlight="black">					
						<colgroup>
							<col width="150px" />
							<% for i = 1 to 50 %>
								<col width="100px" />
							<% next %>
						</colroup>
						<tr height="30">
							<td align="left" bgcolor="#FFFFFF" class="TDCont" colspan="300">&nbsp;<img src="/Images/dot_01.gif" >&nbsp;<b><font color="#ff00ff"></font> 3. ��޺�</b></td>
						</tr>
						
						<%
						sql = " select BCLASS, CLASSNAME "
						sql = sql & " , case "
						sql = sql & " 	when (select count(*) from TB_ARMYINFO where ACLASS = 'P' and BCLASS = ai.BCLASS and CCLASS is not null and DCLASS is not null) = 0 then "
						sql = sql & " 		(select count(*) from TB_ARMYINFO where ACLASS = 'P' and BCLASS = ai.BCLASS and CCLASS is not null and DCLASS is null) "
						sql = sql & " 	else (select count(*) from TB_ARMYINFO where ACLASS = 'P' and BCLASS = ai.BCLASS and CCLASS is not null and DCLASS is not null) end "
						sql = sql & " , case "
						sql = sql & " 	when (select count(*) from TB_ARMYINFO where ACLASS = 'P' and BCLASS = ai.BCLASS and CCLASS is not null and DCLASS is not null) = 0 then 1 "
						sql = sql & " 	else (select count(*) from TB_ARMYINFO where ACLASS = 'P' and BCLASS = ai.BCLASS and CCLASS is not null and DCLASS is null) end "
						sql = sql & " from TB_ARMYINFO as ai where ACLASS = 'P' and BCLASS is not null and CCLASS is null order by BClass "
						'response.write	sql
						set rs = db.execute(sql)
						if not rs.eof then
							arrRs = rs.getRows
							arrRc = ubound(arrRs,2)
						else
							arrRc = -1
						end if
						rs.close
						set rs = nothing
						
						lineA =	"<tr bgcolor=""#EEF6FF"">"
						lineA = lineA &	"	<td align=""center"" class=""TDCont"" with=""150px"" rowspan=""3"">����</td>"
						
						subSql = " select inCode "
						
						for i = 0 to arrRc
							
							lineA = lineA &	"<td align=""center"" class=""TDCont"" colspan=""" & arrRs(2,i) + arrRs(3,i) & """ "
							if arrRs(2,i) = 0 then
								lineA = lineA &	"rowspan=""3"""
							end if
							lineA = lineA &	">" & arrRs(1,i) & "</td>"
						
							sql = " select BCLASS, CCLASS, CLASSNAME "
							sql = sql & " , case when (select count(*) from TB_ARMYINFO where ACLASS = 'P' and BCLASS = ai.BCLASS and CCLASS = ai.CCLASS and DCLASS is not null) < 2 then 1 "
							sql = sql & " 		else  (select count(*) from TB_ARMYINFO where ACLASS = 'P' and BCLASS = ai.BCLASS and CCLASS = ai.CCLASS and DCLASS is not null) + 1 end "
							sql = sql & " , (select count(*) from TB_ARMYINFO where ACLASS = 'P' and BCLASS = ai.BCLASS and CCLASS = ai.CCLASS and DCLASS is not null) "
							sql = sql & " from TB_ARMYINFO as ai "
							sql = sql & " where BCLASS = '" & arrRs(0,i) & "' and CCLASS is not null and DCLASS is null "
							sql = sql & " order by BCLASS, CCLASS "
							'response.write	sql
							set rs = db.execute(sql)
							if not rs.eof then
								arrRs2 = rs.getRows
								arrRc2 = ubound(arrRs2,2)
							else
								arrRc2 = -1
							end if
							rs.close
							set rs = nothing
							
							if arrRc2 = -1 and arrRc > 0 then
								subSql = subSql & " ,count(case when Level_B = '" & arrRs(0,i) & "' then 1 else null end) col_" & i & " "
							end if
							
							if i = arrRc and arrRc2 = -1 and arrRc > 0 then
								lineA = lineA & "<td align=""center"" class=""TDCont"" rowspan=""3"">��</td>"
								subSql = subSql & " ,count(case when Level_B in ('P01','P02','P09','P13','P15') then 1 else null end) sum_" & i & "  "
							end if
							
							for ii = 0 to arrRc2
								
								lineB = lineB &	"<td align=""center"" class=""TDCont"" colspan=""" & arrRs2(3,ii) & """ "
								if arrRs2(4,ii) = 0 then
									lineB = lineB &	"rowspan=""2"""
								end if
								lineB = lineB &	">" & arrRs2(2,ii) & "</td>"
								
								sql = " select BCLASS, CCLASS, DCLASS, CLASSNAME from TB_ARMYINFO where BCLASS = '" & arrRs2(0,ii) & "' and CCLASS = '" & arrRs2(1,ii) & "' and DCLASS is not null "
								sql = sql & " order by DCLASS "
								'response.write	sql
								set rs = db.execute(sql)
								if not rs.eof then
									arrRs3 = rs.getRows
									arrRc3 = ubound(arrRs3,2)
								else
									arrRc3 = -1
								end if
								rs.close
								set rs = nothing
								
								if arrRc3 = -1 and arrRc2 > 0 then
									subSql = subSql & " ,count(case when Level_B = '" & arrRs2(0,ii) & "' and Level_C = '" & arrRs2(1,ii) & "' then 1 else null end) col_" & i & "_" & ii & " "
								end if
								
								if ii = arrRc2 and arrRc3 = -1 and arrRc2 > 0 then
									lineB = lineB & "<td align=""center"" class=""TDCont"" rowspan=""2"">��</td>"
									subSql = subSql & " ,count(case when Level_B = '" & arrRs2(0,ii) & "' then 1 else null end) sum_" & i & "_" & ii & " "
								end if
								
								for iii = 0 to arrRc3
									
									subSql = subSql & " ,count(case when Level_B = '" & arrRs3(0,iii) & "' and Level_C = '" & arrRs3(1,iii) & "' and Level_D = '" & arrRs3(2,iii) & "' then 1 else null end) col_" & i & "_" & ii & "_" & iii & " "
									
									lineC = lineC &	"<td align=""center"" class=""TDCont"">" & arrRs3(3,iii) & "</td>"
									
									if iii = arrRc3 and arrRc3 > 0 then
										lineC = lineC & "<td align=""center"" class=""TDCont"">��</td>"
										subSql = subSql & " ,count(case when Level_B = '" & arrRs3(0,iii) & "' and Level_C = '" & arrRs3(1,iii) & "' then 1 else null end) sum_" & i & "_" & ii & "_" & iii & " "
									end if
									
								next
								
							next
							
						next
						
						lineA = lineA &	"</tr>"
						
						response.write	lineA
						response.write	"<tr bgcolor=""#EEF6FF"">" & lineB & "</tr>"
						response.write	"<tr bgcolor=""#EEF6FF"">" & lineC & "</tr>"
						
						subSql = subSql & " from tb_lifecallhistory  where jubdate >= '" & FromDate & "' AND		jubdate <= '" & ToDate & "' and CHANNELGB_B in ('Q01','Q03','Q05','Q07','Q09','Q99') " & chSql
						subSql = subSql & " group by inCode order by inCode "
						'response.write	subSql
						set rs = db.execute(subSql)
						if not rs.eof then
							arrRs = rs.getRows
							arrRc2 = ubound(arrRs,2)
							arrRc1 = ubound(arrRs,1)
						else
							arrRc2 = -1
						end if
						rs.close
						set rs = nothing
						
						dim colSum(99)
						
						for i = 0 to arrRc2
							response.write	"<tr bgcolor=""#EEF6FF"">"
							for ii = 0 to arrRc1
								if ii = 0 then
									response.write	"<td align=""center"" class=""TDCont"">" & db_getUserName(arrRs(0,i)) & "</td>"
								else
									response.write	"<td bgcolor=""#FFFFFF"" align=""center"" class=""TDCont"">" & arrRs(ii,i) & "</td>"
									colSum(ii) = colSum(ii) + arrRs(ii,i)
								end if
							next
							response.write	"</tr>"
						next
						
						response.write	"<tr bgcolor=""#FFEEF9"">"
						response.write	"<td align=""center"" class=""TDCont"">�Ѱ�</td>"
						for i = 1 to arrRc1
							response.write	"<td align=""center"" class=""TDCont"">" & colSum(i) & "</td>"
						next
						response.write	"</tr>"
						%>
						
					</table>
				</div>
	
				<!--��޺�-->
				<table width="100%" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
				
				<DIV style="OVERFLOW-X: auto;OVERFLOW-Y: auto; MARGIN: 0px 0px 0px 0px; WIDTH: 1200;">
					<table width="2900"  border="1" cellpadding="0" cellspacing="0" bgcolor="#EFECE5" align="center" bordercolor="black" bordercolordark="white" bordercolorlight="black">
						<tr height="30">
							<td align="left" bgcolor="#FFFFFF" class="TDCont" colspan="300">&nbsp;<img src="/Images/dot_01.gif" >&nbsp;<b><font color="#ff00ff"></font> 4. �δ뺰</b></td>
						</tr>
	
						<%
						SQL = " select * from tb_armyinfo where aclass < 'O' and bclass is null order by aclass "
						set Rs = db.execute(SQL)
					
						firstLine = "<tr bgcolor='#EEF6FF'>"
						firstLine = firstLine & "<td align='center' class='TDCont'  width='150' colspan= '2' rowspan='2'>����</td>"
						secondLine = "<tr bgcolor='#EEF6FF'>"
						execSQL = "select inCode "
					
						iColTot = 0
						
						do until rs.eof
					
							aclass = rs("aclass")
							icol = 0
							subSQL = " select * from tb_armyinfo where aclass = '" & aclass & "' and bclass is not null and Cclass is null order by aclass, bclass "
							set subRs = db.execute(subSQL)
					
							if subRs.eof = false then
								
								inValue = ""
								
								do until subRs.eof
									
									bclass = subRs("bclass")				
									secondLine = secondLine & "<td bgcolor='#EEF6FF' align='center' class='TDCont'  width='50'>"&subRs("classname")&"</td>"
									icol = icol + 1
									iColTot = iColTot + 1
									execSQL = execSQL & " ,case when SOSOKGB_A + SOSOKGB_B = '" & rs("aclass") & subRs("bclass") & "' then 1 else 0 end col" & iColTot & ""
									
									if inValue = "" then 
										inValue =  rs("aclass") & subRs("bclass") 
									else
										inValue =  inValue & "," & rs("aclass") & subRs("bclass") 
									end if
									
									subRs.movenext
								loop
								
								if icol > 1 then
									icol = icol + 1
									secondLine = secondLine & "<td bgcolor='#EEF6FF' align='center' class='TDCont'  width='50'>�Ѱ�</td>"
									iColTot = iColTot + 1
									execSQL = execSQL & " ,case when SOSOKGB_A + SOSOKGB_B in ('" & replace(inValue,",","','") & "') then 1 else 0 end col" & iColTot & ""
								end if
								
								firstLine = firstLine & "<td bgcolor='#EEF6FF' align='center' class='TDCont'  width='" & (50*icol) & "' colspan="&icol&">"&rs("classname")&"</td>"
								
							else
								
								icol = 1
								firstLine = firstLine & "<td bgcolor='#EEF6FF' align='center' class='TDCont'  width='50' rowspan=2>"&rs("classname")&"</td>"
								iColTot = iColTot + 1
								execSQL = execSQL & " ,case when SOSOKGB_A  = '" & rs("Aclass") & "' then 1 else 0 end col" & iColTot & ""
					
								if inValue = "" then 
									inValue =  rs("Aclass") 
								else
									inValue =  inValue & "," & rs("Aclass")
								end if
								
							end if
					
							if inValue <> "" then
								if inTotValue = "" then 
									inTotValue =  inValue
								else
									inTotValue =  inTotValue & "," & inValue
								end if	
								inValue = ""
							end if
							
							rs.movenext
						loop
						
						secondLine = secondLine & "</tr>"
						firstLine = firstLine & "<td bgcolor='#EEF6FF' align='center' class='TDCont'  width='50' rowspan='2'>�Ѱ�</td></tr>"
						response.write firstLine
						response.write secondLine
					
						iColTot = iColTot + 1
						execSQL = execSQL & " 	,case when SOSOKGB_A + SOSOKGB_B in ('" & replace(inTotValue,",","','") & "') then 1 else 0 end col" & iColTot & ""
						execSQL = execSQL & " from tb_lifecallhistory  where jubdate >= '" & FromDate & "' AND		jubdate <= '" & ToDate & "' "
						execSQL = execSQL & chSql
					
						execSQL1 = " select incode"
						for i = 1 to iColTot
							execSQL1 = execSQL1 & ", sum(col"&i &") col"&i	
							vtot(i) = 0
						next
						execSQL1 = execSQL1 & " from (		" & execSQL & " ) b group by incode order by incode"
						
						set Rs = db.execute(execSQL1)
					
						do until rs.eof
							%>
							
							<tr bgcolor='#EEF6FF'>
								<td align='center' class='TDCont'  width='300' colspan='2' ><%=db_getUserName(rs("incode"))%></td>
								
								<%
								for i = 1 to iColTot
									sLine = sLine & "<td bgcolor='#ffffff' align='center' class='TDCont'>" & rs(i) & "</td>"	
									vtot(i) = vtot(i) + rs(i)
								next
								response.write sLine
								sLine = ""
								%>
								
							</tr>
							
							<%
							rs.movenext
						loop
						%>
						
						<tr bgcolor='#FFEEF9'>
							<td align='center' class='TDCont'  width='300' colspan='2'>�Ѱ�</td>
							
							<%
							for i = 1 to iColTot
								sLine = sLine & "<td bgcolor='#FFEEF9' align='center' class='TDCont' >" & vtot(i) & "</td>"	
								vtot(i) = 0
							next
							response.write sLine
							sLine = ""
							%>
							
						</tr>
					</table>
				</div>
				
				<!--���оߺ�-->
				<table width="100%" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
				
				<DIV style="OVERFLOW-X: auto;OVERFLOW-Y: auto; MARGIN: 0px 0px 0px 0px; WIDTH: 1200;">
					<table width="5800"  border="1" cellpadding="0" cellspacing="0" bgcolor="#EFECE5" align="center" bordercolor="black" bordercolordark="white" bordercolorlight="black">
						<tr height="30">
							<td align="left" bgcolor="#FFFFFF" class="TDCont" colspan="300">&nbsp;<img src="/Images/dot_01.gif" >&nbsp;<b><font color="#ff00ff"></font> 5. ���оߺ�</b></td>
						</tr>
						
						<%
						SQL = " select * from tb_armyinfo where aclass = 'O' and bclass is not null and cclass is null order by bclass, cclass "
						set Rs = db.execute(SQL)
					
						firstLine = "<tr bgcolor='#EEF6FF'>"
						firstLine = firstLine & "<td align='center' class='TDCont'  width='150' colspan= '2' rowspan='2'>����</td>"
						secondLine = "<tr bgcolor='#EEF6FF'>"
						execSQL = "select inCode "
					
						iColTot = 0
						do until rs.eof
					
							bclass = rs("bclass")
							icol = 0
							subSQL = " select * from tb_armyinfo where aclass = 'O' and bclass = '" & bclass & "' and cclass is not null order by bclass, cclass "
							set subRs = db.execute(subSQL)
					
							if subRs.eof = false then
								inValue = ""
								do until subRs.eof
									cclass = subRs("cclass")				
									secondLine = secondLine & "<td bgcolor='#EEF6FF' align='center' class='TDCont'  width='50'>"&subRs("classname")&"</td>"
									icol = icol + 1
									iColTot = iColTot + 1
									execSQL = execSQL & " ,case when CALLCLASS_B + CALLCLASS_C = '" & rs("bclass") & subRs("cclass") & "' then 1 else 0 end col" & iColTot & ""
									if inValue = "" then 
										inValue =  rs("bclass") & subRs("cclass") 
									else
										inValue =  inValue & "," & rs("bclass") & subRs("cclass") 
									end if
									subRs.movenext
								loop
								if icol > 1 then
									icol = icol + 1
									secondLine = secondLine & "<td bgcolor='#EEF6FF' align='center' class='TDCont'  width='50'>�Ѱ�</td>"
									iColTot = iColTot + 1
									execSQL = execSQL & " ,case when CALLCLASS_B + CALLCLASS_C in ('" & replace(inValue,",","','") & "') then 1 else 0 end col" & iColTot & ""
								end if
								firstLine = firstLine & "<td bgcolor='#EEF6FF' align='center' class='TDCont'  width='" & (50*icol) & "' colspan="&icol&">"&rs("classname")&"</td>"
							else
								icol = 1
								firstLine = firstLine & "<td bgcolor='#EEF6FF' align='center' class='TDCont'  width='50' rowspan=2>"&rs("classname")&"</td>"
								iColTot = iColTot + 1
								execSQL = execSQL & " ,case when CALLCLASS_B  = '" & rs("bclass") & "' then 1 else 0 end col" & iColTot & ""
					
								if inValue = "" then 
									inValue =  rs("bclass") 
								else
									inValue =  inValue & "," & rs("bclass")
								end if
							end if
					
							if inValue <> "" then
								if inTotValue = "" then 
									inTotValue =  inValue
								else
									inTotValue =  inTotValue & "," & inValue
								end if	
								inValue = ""
							end if
							rs.movenext
						loop
						
						secondLine = secondLine & "</tr>"
						firstLine = firstLine & "<td bgcolor='#EEF6FF' align='center' class='TDCont'  width='50' rowspan='2'>�Ѱ�</td></tr>"
						response.write firstLine
						response.write secondLine
					
						iColTot = iColTot + 1
						execSQL = execSQL & " ,case when CALLCLASS_B + CALLCLASS_C in ('" & replace(inTotValue,",","','") & "') then 1 else 0 end col" & iColTot & ""
						execSQL = execSQL & "	 from tb_lifecallhistory  where jubdate >= '" & FromDate & "' AND		jubdate <= '" & ToDate & "'"
						execSQL = execSQL & chSql
					
						execSQL1 = " select incode"
						for i = 1 to iColTot
							execSQL1 = execSQL1 & ", sum(col"&i &") col"&i	
							vtot(i) = 0
						next
						execSQL1 = execSQL1 & " from (		" & execSQL & "  AND CHANNELGB_B + CHANNELGB_C IN ('Q01Q01A','Q01Q01C','Q03Q03A','Q03Q03C','Q09')) b group by incode order by incode"
						
						set Rs = db.execute(execSQL1)
					
						do until rs.eof
							%>
							
							<tr bgcolor='#EEF6FF'>
								<td align='center' class='TDCont'  width='300' colspan='2' ><%=db_getUserName(rs("incode"))%></td>
								
								<%
								'����������� �Ѹ���
								for i = 1 to iColTot
									sLine = sLine & "<td bgcolor='#ffffff' align='center' class='TDCont'>" & rs(i) & "</td>"	
									vtot(i) = vtot(i) + rs(i)
								next
								response.write sLine
								sLine = ""
								%>
								
							</tr>
							<%
							rs.movenext
						loop
						%>
						
						<tr bgcolor='#FFEEF9'>
							<td align='center' class='TDCont'  width='300' colspan='2'>�Ѱ�</td>
							
							<%
							for i = 1 to iColTot
								sLine = sLine & "<td bgcolor='#FFEEF9' align='center' class='TDCont' >" & vtot(i) & "</td>"	
								vtot(i) = 0
							next
							response.write sLine
							sLine = ""
							%>
							
						</tr>
					</table>
				</div>
		
				<table width="1200" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
				
				<table width="1200"  border="1" cellpadding="0" cellspacing="0" bgcolor="#EFECE5" align="center" bordercolor="black" bordercolordark="white" bordercolorlight="black">
					<tr height="30">
						<td align="left" bgcolor="#FFFFFF" class="TDCont" colspan="9">&nbsp;<img src="/Images/dot_01.gif" >&nbsp;<b><font color="#ff00ff"></font> 6. ���Ϻ�</b></td>
					</tr>
					<tr bgcolor='#EEF6FF'>
						<td align='center' class='TDCont'  width='150'>����</td>
						<td align='center' class='TDCont' width='150'>��</td>
						<td align='center' class='TDCont' width='150'>ȭ</td>
						<td align='center' class='TDCont' width='150'>��</td>
						<td align='center' class='TDCont' width='150'>��</td>
						<td align='center' class='TDCont' width='150'>��</td>
						<td align='center' class='TDCont'width='150' >��</td>
						<td align='center' class='TDCont'width='150' >��</td>
						<td align='center' class='TDCont'  width='150'>�Ѱ�</td>
					</tr>
					
					<%
					'������ �Ѱ�
					SQL = "select * from ( "
					SQL = SQL & " 	SELECT	'1' gubun, incode, count(incode) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
					SQL = SQL & "		AND		jubdate <= '" & ToDate & "' and datepart(weekday,jubdate) = 2 " & chSql & " group by incode "
					SQL = SQL & "	union all "
					SQL = SQL & " 	SELECT	'2' gubun, incode, count(incode) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
					SQL = SQL & "		AND		jubdate <= '" & ToDate & "' and datepart(weekday,jubdate) = 3 " & chSql & " group by incode " 
					SQL = SQL & "	union all "
					SQL = SQL & " 	SELECT	'3' gubun, incode, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
					SQL = SQL & "		AND		jubdate <= '" & ToDate & "' and datepart(weekday,jubdate) = 4 " & chSql & " group by incode" '
					SQL = SQL & "	union all "
					SQL = SQL & " 	SELECT	'4' gubun, incode, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
					SQL = SQL & "		AND		jubdate <= '" & ToDate & "' and datepart(weekday,jubdate) = 5 " & chSql & " group by incode" '
					SQL = SQL & "	union all "
					SQL = SQL & " 	SELECT	'5' gubun, incode, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
					SQL = SQL & "		AND		jubdate <= '" & ToDate & "' and datepart(weekday,jubdate) = 6 " & chSql & " group by incode" '
					SQL = SQL & "	union all "
					SQL = SQL & " 	SELECT	'6' gubun, incode, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
					SQL = SQL & "		AND		jubdate <= '" & ToDate & "' and datepart(weekday,jubdate) = 7 " & chSql & " group by incode" '
					SQL = SQL & "	union all "
					SQL = SQL & " 	SELECT	'7' gubun, incode, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
					SQL = SQL & "		AND		jubdate <= '" & ToDate & "' and datepart(weekday,jubdate) = 1 " & chSql & " group by incode" '
					SQL = SQL & "	union all "
					SQL = SQL & " 	SELECT	'8' gubun, incode, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
					SQL = SQL & "		AND		jubdate <= '" & ToDate & "' " & chSql & " group by incode) a order by incode, gubun" '
				
					set Rs = db.execute(SQL)
	
					tot1 = 0
					tot2 = 0
					tot3 = 0
					tot4 = 0
					tot5 = 0
					tot6 = 0
					tot7 = 0
					tot8 = 0
					tot9 = 0
					
					do until rs.eof
					
						tot1 = 0
						tot2 = 0
						tot3 = 0
						tot4 = 0
						tot5 = 0
						tot6 = 0
						tot7 = 0
						tot8 = 0
						tot9 = 0
		
						incode = rs("incode")
						
						do until incode <> rs("incode")
							if rs("gubun") = "1" then
								tot1 = rs("cnt")
							elseif rs("gubun") = "2" then
								tot2 = rs("cnt")
							elseif rs("gubun") = "3" then
								tot3 = rs("cnt")
							elseif rs("gubun") = "4" then
								tot4 = rs("cnt")
							elseif rs("gubun") = "5" then
								tot5 = rs("cnt")
							elseif rs("gubun") = "6" then
								tot6 = rs("cnt")
							elseif rs("gubun") = "7" then
								tot7 = rs("cnt")
							elseif rs("gubun") = "8" then
								tot8 = rs("cnt")
							end if
				
							rs.movenext
							if rs.eof then
								exit do
							end if
						loop
						%>
						
						<tr bgcolor='#EEF6FF'>
							<td align='center' class='TDCont'  width='150'><%=db_getUserName(incode)%></td>
							<td bgcolor='#ffffff' align='center' class='TDCont'><%=tot1%></td>
							<td bgcolor='#ffffff' align='center' class='TDCont'><%=tot2%></td>
							<td bgcolor='#ffffff' align='center' class='TDCont'><%=tot3%></td>
							<td bgcolor='#ffffff' align='center' class='TDCont'><%=tot4%></td>
							<td bgcolor='#ffffff' align='center' class='TDCont'><%=tot5%></td>
							<td bgcolor='#ffffff' align='center' class='TDCont'><%=tot6%></td>
							<td bgcolor='#ffffff' align='center' class='TDCont'><%=tot7%></td>
							<td bgcolor='#ffffff' align='center' class='TDCont'><%=tot8%></td>
						</tr>
						<%
				
						if rs.eof then
							exit do
						end if
					loop
		
					'������ �Ѱ�
					SQL = "select * from ( SELECT	'1' gubun, count(incode) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
					SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and datepart(weekday,jubdate) = 2 "&chSql
					SQL = SQL & "	union all SELECT	'2' gubun, count(incode) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
					SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and datepart(weekday,jubdate) = 3 "&chSql
					SQL = SQL & "	union all SELECT	'3' gubun, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
					SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and datepart(weekday,jubdate) = 4 "&chSql
					SQL = SQL & "	union all SELECT	'4' gubun, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
					SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and datepart(weekday,jubdate) = 5 "&chSql
					SQL = SQL & "	union all SELECT	'5' gubun, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
					SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and datepart(weekday,jubdate) = 6 "&chSql
					SQL = SQL & "	union all SELECT	'6' gubun, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
					SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and datepart(weekday,jubdate) = 7 "&chSql
					SQL = SQL & "	union all SELECT	'7' gubun, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
					SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and datepart(weekday,jubdate) = 1 "&chSql
					SQL = SQL & "	union all SELECT	'8' gubun, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'  "&chSql
					SQL = SQL & "	AND		jubdate <= '" & ToDate & "') a order by gubun" '
				
					set Rs = db.execute(SQL)
				
					tot1 = 0
					tot2 = 0
					tot3 = 0
					tot4 = 0
					tot5 = 0
					tot6 = 0
					tot7 = 0
					tot8 = 0
					tot9 = 0
					
					do until rs.eof
		
						if rs("gubun") = "1" then
							tot1 = rs("cnt")
						elseif rs("gubun") = "2" then
							tot2 = rs("cnt")
						elseif rs("gubun") = "3" then
							tot3 = rs("cnt")
						elseif rs("gubun") = "4" then
							tot4 = rs("cnt")
						elseif rs("gubun") = "5" then
							tot5 = rs("cnt")
						elseif rs("gubun") = "6" then
							tot6 = rs("cnt")
						elseif rs("gubun") = "7" then
							tot7 = rs("cnt")
						elseif rs("gubun") = "8" then
							tot8 = rs("cnt")
						end if
				
						rs.movenext
						if rs.eof then
							exit do
						end if
					loop
					%>
					
					<tr bgcolor='#FFEEF9'>
						<td align='center' class='TDCont'  width='150'>��</td>
						<td align='center' class='TDCont'><%=tot1%></td>
						<td align='center' class='TDCont'><%=tot2%></td>
						<td align='center' class='TDCont'><%=tot3%></td>
						<td align='center' class='TDCont'><%=tot4%></td>
						<td align='center' class='TDCont'><%=tot5%></td>
						<td align='center' class='TDCont'><%=tot6%></td>
						<td align='center' class='TDCont'><%=tot7%></td>
						<td align='center' class='TDCont'><%=tot8%></td>
					</tr>
				</table>
	
				<table width="1200" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
				
				<table width="1200"  border="1" cellpadding="0" cellspacing="0" bgcolor="#EFECE5" align="center" bordercolor="black" bordercolordark="white" bordercolorlight="black">
					<tr height="30">
						<td align="left" bgcolor="#FFFFFF" class="TDCont" colspan="26">&nbsp;<img src="/Images/dot_01.gif" >&nbsp;<b><font color="#ff00ff"></font> 7. �ð��뺰</b></td>
					</tr>
					<tr bgcolor='#EEF6FF'>
						<td align='center' class='TDCont'  width='200'>����</td>
						
						<%
						for i = 0 to 23
							if i < 10 then
								sHourname = "0" & i & "��"
							else
								sHourname = i & "��"
							end if
							%>
							<td align='center' class='TDCont'  width='140'><%=sHourname%></td>
							<%
						next
						%>
						
						<td align='center' class='TDCont' width='140'>�Ѱ�</td>
					</tr>
					
					<%
					SQL = " SELECT	incode"
					for i = 0 to 23
						SQL = SQL & "			, case when datepart(hour,jubtime) = " & i & " then 1 else 0 end col" & i
					next
					SQL = SQL & "			, 1 col" & i
					SQL = SQL & "	FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'	AND		jubdate <= '" & ToDate & "' " & chSql & " "
					
					sumSQL = " SELECT	incode "
					for i = 0 to 24
						sumSQL = sumSQL & "			,sum(col" & i & ") col" & i
					next
					sumSQL = sumSQL & "	FROM	( " & SQL & ") b group by inCode order by inCode"
					
					set Rs = db.execute(sumSQL)
					
					do until rs.eof
						%>
						
						<tr bgcolor='#EEF6FF'>
							<td align='center' class='TDCont'  width='300' colspan='1' ><%=db_getUserName(rs("incode"))%></td>
							
							<%
							for i = 0 to 24
								sLine = sLine & "<td bgcolor='#ffffff' align='center' class='TDCont'>" & rs(i+1) & "</td>"	
								vtot(i+1) = vtot(i+1) + rs(i+1)
							next
							response.write sLine
							sLine = ""
							%>
							
						</tr>
	
						<%
						rs.movenext
					loop
					%>
					
					<tr bgcolor='#FFEEF9'>
						<td align='center' class='TDCont'  width='300' colspan='1'>�Ѱ�</td>
	
							<%
							for i = 0 to 24
								sLine = sLine & "<td bgcolor='#FFEEF9' align='center' class='TDCont' >" & vtot(i+1) & "</td>"	
								vtot(i+1) = 0
							next
							response.write sLine
							sLine = ""
							%>
					</tr>
				</table>
	
				<table width="1200" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
				
				<table width="1200"  border="1" cellpadding="0" cellspacing="0" bgcolor="#EFECE5" align="center" bordercolor="black" bordercolordark="white" bordercolorlight="black">
					<tr height="30">
						<td align="left" bgcolor="#FFFFFF" class="TDCont" colspan="11">&nbsp;<img src="/Images/dot_01.gif" >&nbsp;<b><font color="#ff00ff"></font> 8. ��ȭ�ð���</b> (���̹�,��ȭ����,ħ��(Ž��)��ȭ,��Ÿ ����)</td>
					</tr>
					<tr bgcolor='#EEF6FF'>
						<td align='center' class='TDCont' width='150'>����</td>
						<td align='center' class='TDCont' width='150'>1�й̸�</td>
						<td align='center' class='TDCont' width='150'>1-5��</td>
						<td align='center' class='TDCont' width='150'>6-10��</td>
						<td align='center' class='TDCont' width='150'>11-20��</td>
						<td align='center' class='TDCont' width='150'>21-30��</td>
						<td align='center' class='TDCont' width='150'>31-40��</td>
						<td align='center' class='TDCont' width='150'>41-50��</td>
						<td align='center' class='TDCont' width='150'>51-60��</td>
						<td align='center' class='TDCont' width='150'>60���̻�</td>
						<td align='center' class='TDCont' width='150'>�Ѱ�</td>
					</tr>
	
					<%
					SQL = "select * from ( "
					SQL = SQL & " 	SELECT	'01' gubun, incode, count(incode) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
					SQL = SQL & "		AND jubdate <= '" & ToDate & "' and CHANNELGB_B not in ('Q03','Q05','Q07') and calltime < 60 " & chSql & " group by incode" '
					SQL = SQL & "	union all "
					SQL = SQL & " 	SELECT	'02' gubun, incode, count(incode) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
					SQL = SQL & "		AND jubdate <= '" & ToDate & "' and CHANNELGB_B not in ('Q03','Q05','Q07') and calltime >=60 and calltime <=300 " & chSql & " group by incode" '
					SQL = SQL & "	union all "
					SQL = SQL & " 	SELECT	'03' gubun, incode, count(incode) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
					SQL = SQL & "		AND		jubdate <= '" & ToDate & "' and CHANNELGB_B not in ('Q03','Q05','Q07') and calltime >=301 and calltime <=600 " & chSql & " group by incode" '
					SQL = SQL & "	union all "
					SQL = SQL & " 	SELECT	'04' gubun, incode, count(incode) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
					SQL = SQL & "		AND		jubdate <= '" & ToDate & "' and CHANNELGB_B not in ('Q03','Q05','Q07') and calltime >=601 and calltime <=1200 " & chSql & " group by incode" '
					SQL = SQL & "	union all "
					SQL = SQL & " 	SELECT	'05' gubun, incode, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
					SQL = SQL & "		AND		jubdate <= '" & ToDate & "' and CHANNELGB_B not in ('Q03','Q05','Q07') and calltime >=1201 and calltime <=1800 " & chSql & " group by incode" '
					SQL = SQL & "	union all "
					SQL = SQL & " 	SELECT	'06' gubun, incode, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
					SQL = SQL & "		AND		jubdate <= '" & ToDate & "' and CHANNELGB_B not in ('Q03','Q05','Q07') and calltime >=1801 and calltime <=2400 " & chSql & " group by incode" '
					SQL = SQL & "	union all "
					SQL = SQL & " 	SELECT	'07' gubun, incode, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
					SQL = SQL & "		AND		jubdate <= '" & ToDate & "' and CHANNELGB_B not in ('Q03','Q05','Q07') and calltime >=2401 and calltime <=3000 " & chSql & " group by incode" '
					SQL = SQL & "	union all "
					SQL = SQL & " 	SELECT	'08' gubun, incode, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
					SQL = SQL & "		AND		jubdate <= '" & ToDate & "' and CHANNELGB_B not in ('Q03','Q05','Q07') and calltime >=3001 and calltime <=3600 " & chSql & " group by incode" '
					SQL = SQL & "	union all "
					SQL = SQL & " 	SELECT	'09' gubun, incode, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
					SQL = SQL & "		AND		jubdate <= '" & ToDate & "' and CHANNELGB_B not in ('Q03','Q05','Q07') and calltime >=3601 " & chSql & " group by incode" '
					SQL = SQL & "	union all "
					SQL = SQL & " 	SELECT	'10' gubun, incode, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
					SQL = SQL & "		AND		jubdate <= '" & ToDate & "' and CHANNELGB_B not in ('Q03','Q05','Q07') " & chSql & " group by incode) a order by incode, gubun" '
				
					set Rs = db.execute(SQL)
				
					do until rs.eof
					
						tot1 = 0
						tot2 = 0
						tot3 = 0
						tot4 = 0
						tot5 = 0
						tot6 = 0
						tot7 = 0
						tot8 = 0
						tot9 = 0
						tot10 = 0
		
						incode = rs("incode")
						do until incode <> rs("incode")
							if rs("gubun") = "01" then
								tot1 = rs("cnt")
							elseif rs("gubun") = "02" then
								tot2 = rs("cnt")
							elseif rs("gubun") = "03" then
								tot3 = rs("cnt")
							elseif rs("gubun") = "04" then
								tot4 = rs("cnt")
							elseif rs("gubun") = "05" then
								tot5 = rs("cnt")
							elseif rs("gubun") = "06" then
								tot6 = rs("cnt")
							elseif rs("gubun") = "07" then
								tot7 = rs("cnt")
							elseif rs("gubun") = "08" then
								tot8 = rs("cnt")
							elseif rs("gubun") = "09" then
								tot9 = rs("cnt")
							elseif rs("gubun") = "10" then
								tot10 = rs("cnt")
							end if
			
							rs.movenext
							if rs.eof then
								exit do
							end if
						loop
						%>
						
						<tr bgcolor='#EEF6FF'>
							<td align='center' class='TDCont'  width='150'><%=db_getUserName(incode)%></td>
							<td bgcolor='#ffffff' align='center' class='TDCont'><%=tot1%></td>
							<td bgcolor='#ffffff' align='center' class='TDCont'><%=tot2%></td>
							<td bgcolor='#ffffff' align='center' class='TDCont'><%=tot3%></td>
							<td bgcolor='#ffffff' align='center' class='TDCont'><%=tot4%></td>
							<td bgcolor='#ffffff' align='center' class='TDCont'><%=tot5%></td>
							<td bgcolor='#ffffff' align='center' class='TDCont'><%=tot6%></td>
							<td bgcolor='#ffffff' align='center' class='TDCont'><%=tot7%></td>
							<td bgcolor='#ffffff' align='center' class='TDCont'><%=tot8%></td>
							<td bgcolor='#ffffff' align='center' class='TDCont'><%=tot9%></td>
							<td bgcolor='#ffffff' align='center' class='TDCont'><%=tot10%></td>
						</tr>
					
						<%
						if rs.eof then
							exit do
						end if
					loop
				
					SQL = "select * from ( SELECT	'01' gubun, incode, count(incode) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
					SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and CHANNELGB_B not in ('Q03','Q05','Q07','Q99') and calltime <60 " & chSql & " group by incode" '
					SQL = SQL & "	union all SELECT	'02' gubun, incode, count(incode) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
					SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and CHANNELGB_B not in ('Q03','Q05','Q07','Q99') and calltime >=60 and calltime <=300 " & chSql & " group by incode" '
					SQL = SQL & "	union all SELECT	'03' gubun, incode, count(incode) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
					SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and CHANNELGB_B not in ('Q03','Q05','Q07','Q99') and calltime >=301 and calltime <=600 " & chSql & " group by incode" '
					SQL = SQL & "	union all SELECT	'04' gubun, incode, count(incode) cnt FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
					SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and CHANNELGB_B not in ('Q03','Q05','Q07','Q99') and calltime >=601 and calltime <=1200 " & chSql & " group by incode" '
					SQL = SQL & "	union all SELECT	'05' gubun, incode, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
					SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and CHANNELGB_B not in ('Q03','Q05','Q07','Q99') and calltime >=1201 and calltime <=1800 " & chSql & " group by incode" '
					SQL = SQL & "	union all SELECT	'06' gubun, incode, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
					SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and CHANNELGB_B not in ('Q03','Q05','Q07','Q99') and calltime >=1801 and calltime <=2400 " & chSql & " group by incode" '
					SQL = SQL & "	union all SELECT	'07' gubun, incode, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
					SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and CHANNELGB_B not in ('Q03','Q05','Q07','Q99') and calltime >=2401 and calltime <=3000 " & chSql & " group by incode" '
					SQL = SQL & "	union all SELECT	'08' gubun, incode, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
					SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and CHANNELGB_B not in ('Q03','Q05','Q07','Q99') and calltime >=3001 and calltime <=3600 " & chSql & " group by incode" '
					SQL = SQL & "	union all SELECT	'09' gubun, incode, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
					SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and CHANNELGB_B not in ('Q03','Q05','Q07','Q99') and calltime >=3601 " & chSql & " group by incode" '
					SQL = SQL & "	union all SELECT	'10' gubun, incode, count(incode) cnt  FROM	TB_lifecallhistory	where jubdate >= '" & FromDate & "'"
					SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and CHANNELGB_B not in ('Q03','Q05','Q07','Q99') " & chSql & " group by incode) a order by incode, gubun" '
				
					set Rs = db.execute(SQL)
				
					tot1 = 0
					tot2 = 0
					tot3 = 0
					tot4 = 0
					tot5 = 0
					tot6 = 0
					tot7 = 0
					tot8 = 0
					tot9 = 0
					tot10 = 0
					do until rs.eof
		
						incode = rs("incode")
						
						do until incode <> rs("incode")
							if rs("gubun") = "01" then
								tot1 = tot1 + rs("cnt")
							elseif rs("gubun") = "02" then
								tot2 = tot2 + rs("cnt")
							elseif rs("gubun") = "03" then
								tot3 = tot3 + rs("cnt")
							elseif rs("gubun") = "04" then
								tot4 = tot4 + rs("cnt")
							elseif rs("gubun") = "05" then
								tot5 = tot5 + rs("cnt")
							elseif rs("gubun") = "06" then
								tot6 = tot6 + rs("cnt")
							elseif rs("gubun") = "07" then
								tot7 = tot7 + rs("cnt")
							elseif rs("gubun") = "08" then
								tot8 = tot8 + rs("cnt")
							elseif rs("gubun") = "09" then
								tot9 = tot9 + rs("cnt")
							elseif rs("gubun") = "10" then
								tot10 = tot10 + rs("cnt")
							end if
				
							rs.movenext
							if rs.eof then
								exit do
							end if
						loop
				
						if rs.eof then
							exit do
						end if
					loop
					%>
			
					<tr bgcolor='#FFEEF9'>
						<td align='center' class='TDCont'  width='150'>�Ѱ�</td>
						<td align='center' class='TDCont'><%=tot1%></td>
						<td align='center' class='TDCont'><%=tot2%></td>
						<td align='center' class='TDCont'><%=tot3%></td>
						<td align='center' class='TDCont'><%=tot4%></td>
						<td align='center' class='TDCont'><%=tot5%></td>
						<td align='center' class='TDCont'><%=tot6%></td>
						<td align='center' class='TDCont'><%=tot7%></td>
						<td align='center' class='TDCont'><%=tot8%></td>
						<td align='center' class='TDCont'><%=tot9%></td>
						<td align='center' class='TDCont'><%=tot10%></td>
					</tr>
				</table>
	
	
				<!--��ġ��-->
				<table width="100%" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
				
					<table width="1200"  border="1" cellpadding="0" cellspacing="0" bgcolor="#EFECE5" align="center" bordercolor="black" bordercolordark="white" bordercolorlight="black">
						<tr height="30">
							<td align="left" bgcolor="#FFFFFF" class="TDCont" colspan="300">&nbsp;<img src="/Images/dot_01.gif" >&nbsp;<b><font color="#ff00ff"></font> 9. ��ġ��</b> (��ȭ����,ħ��(Ž��)��ȭ,��Ÿ ����)</b></td>
						</tr>
						
						<%
						SQL = " select * from tb_armyinfo where aclass = 'U' and bclass is not null and cclass is null order by bclass, cclass "
						set Rs = db.execute(SQL)
					
						firstLine = "<tr bgcolor='#EEF6FF'>"
						firstLine = firstLine & "<td align='center' class='TDCont'  width='150' colspan= '2' rowspan='2'>����</td>"
						secondLine = "<tr bgcolor='#EEF6FF'>"
						execSQL = "select inCode "
					
						iColTot = 0
						do until rs.eof
					
							bclass = rs("bclass")
							icol = 0
							subSQL = " select * from tb_armyinfo where aclass = 'U' and bclass = '" & bclass & "' and cclass is not null order by bclass, cclass "
							set subRs = db.execute(subSQL)
					
							if subRs.eof = false then
								inValue = ""
								do until subRs.eof
									cclass = subRs("cclass")				
									secondLine = secondLine & "<td bgcolor='#EEF6FF' align='center' class='TDCont'  width='90'>"&subRs("classname")&"</td>"
									icol = icol + 1
									iColTot = iColTot + 1
									execSQL = execSQL & " ,case when PROCESSGB_B + PROCESSGB_C = '" & rs("bclass") & subRs("cclass") & "' then 1 else 0 end col" & iColTot & ""
									if inValue = "" then 
										inValue =  rs("bclass") & subRs("cclass") 
									else
										inValue =  inValue & "," & rs("bclass") & subRs("cclass") 
									end if
									subRs.movenext
								loop
								if icol > 1 then
									icol = icol + 1
									secondLine = secondLine & "<td bgcolor='#EEF6FF' align='center' class='TDCont'  width='90'>��</td>"
									iColTot = iColTot + 1
									execSQL = execSQL & " ,case when PROCESSGB_B + PROCESSGB_C in ('" & replace(inValue,",","','") & "') then 1 else 0 end col" & iColTot & ""
								end if
								firstLine = firstLine & "<td bgcolor='#EEF6FF' align='center' class='TDCont'  width='" & (50*icol) & "' colspan="&icol&">"&rs("classname")&"</td>"
							else
								icol = 1
								firstLine = firstLine & "<td bgcolor='#EEF6FF' align='center' class='TDCont'  width='90' rowspan=2>"&rs("classname")&"</td>"
								iColTot = iColTot + 1
								execSQL = execSQL & " ,case when PROCESSGB_B  = '" & rs("bclass") & "' then 1 else 0 end col" & iColTot & ""
					
								if inValue = "" then 
									inValue =  rs("bclass") 
								else
									inValue =  inValue & "," & rs("bclass")
								end if
							end if
					
							if inValue <> "" then
								if inTotValue = "" then 
									inTotValue =  inValue
								else
									inTotValue =  inTotValue & "," & inValue
								end if	
								inValue = ""
							end if
							rs.movenext
						loop
						
						secondLine = secondLine & "</tr>"
						firstLine = firstLine & "<td bgcolor='#EEF6FF' align='center' class='TDCont'  width='50' rowspan='2'>�Ѱ�</td></tr>"
						response.write firstLine
						response.write secondLine
					
						iColTot = iColTot + 1
						execSQL = execSQL & " ,case when PROCESSGB_B + PROCESSGB_C in ('" & replace(inTotValue,",","','") & "') then 1 else 0 end col" & iColTot & ""
						execSQL = execSQL & "	 from tb_lifecallhistory  where jubdate >= '" & FromDate & "' AND		jubdate <= '" & ToDate & "'"
						execSQL = execSQL & chSql
					
						execSQL1 = " select incode"
						for i = 1 to iColTot
							execSQL1 = execSQL1 & ", sum(col"&i &") col"&i	
							vtot(i) = 0
						Next  '
						execSQL1 = execSQL1 & " from (		" & execSQL & "  AND CHANNELGB_B not in ('Q05','Q07','Q99') ) b group by incode order by incode"
						
						set Rs = db.execute(execSQL1)
					
						do until rs.eof
							%>
							
							<tr bgcolor='#EEF6FF'>
								<td align='center' class='TDCont'  width='300' colspan='2' ><%=db_getUserName(rs("incode"))%></td>
								
								<%
								'����������� �Ѹ���
								for i = 1 to iColTot
									sLine = sLine & "<td bgcolor='#ffffff' align='center' class='TDCont'>" & rs(i) & "</td>"	
									vtot(i) = vtot(i) + rs(i)
								next
								response.write sLine
								sLine = ""
								%>
								
							</tr>
							<%
							rs.movenext
						loop
						%>
						
						<tr bgcolor='#FFEEF9'>
							<td align='center' class='TDCont'  width='300' colspan='2'>�Ѱ�</td>
							
							<%
							for i = 1 to iColTot
								sLine = sLine & "<td bgcolor='#FFEEF9' align='center' class='TDCont' >" & vtot(i) & "</td>"	
								vtot(i) = 0
							next
							response.write sLine
							sLine = ""
							%>
							
						</tr>
					</table>


				<table width="1200" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
				
				<table width="1200"  border="1" cellpadding="0" cellspacing="0" bgcolor="#EFECE5" align="center" bordercolor="black" bordercolordark="white" bordercolorlight="black">
					<tr height="30">
						<td align="left" bgcolor="#FFFFFF" class="TDCont" colspan="12">&nbsp;<img src="/Images/dot_01.gif" >&nbsp;<b><font color="#ff00ff"></font> 10. ������</b> </td>
					</tr>
					<tr bgcolor='#EEF6FF'>
						<td align='center' class='TDCont' width='150'>����</td>
						
						<%
						SQL = "		SELECT	*	FROM	TB_CODE WHERE CODEGROUP = 'C11' AND USEYN = 'Y' ORDER BY CODE"
						execSQL = " select incode"
						set Rs = db.execute(SQL)
						iCol = 0
						do until rs.eof
							iCol = iCol + 1
							execSQL = execSQL & ", case when Weather = '" & rs("code") & "' then 1 else 0 end col" & iCol
							%>
							
							<td align='center' class='TDCont' width='150'><%=rs("CodeName")%></td>
							
							<%						
							rs.movenext
						loop
						iCol = iCol + 1
						execSQL = execSQL & ", 1 col" & iCol & " from tb_lifecallhistory where jubdate >= '" & FromDate & "'"
						execSQL = execSQL & "	AND		jubdate <= '" & ToDate & "' " & chSql
						%>
						
						<td align='center' class='TDCont' width='150'>�Ѱ�</td>
					</tr>
					
					<%
					SQL = " select incode"
					for i = 1 to iCol
						SQL = SQL & ", sum(col" & i & ") col" & i
					next
					SQL = SQL & "	from ( " & execSQL & " ) b  group by incode order by incode"
					execSQL = ""
					set Rs = db.execute(SQL)
					do until rs.eof
						%>
						
						<tr bgcolor='#EEF6FF'>
							<td align='center' class='TDCont'  width='300' colspan='1' ><%=db_getUserName(rs("incode"))%></td>
	
							<%
							for i = 1 to iCol
								sLine = sLine & "<td bgcolor='#ffffff' align='center' class='TDCont'>" & rs(i) & "</td>"	
								vtot(i) = vtot(i) + rs(i)
							next
							response.write sLine
							sLine = ""
							%>
						</tr>
	
						<%
						rs.movenext
					loop
					%>
					
					<tr bgcolor='#FFEEF9'>
						<td align='center' class='TDCont'  width='300' colspan='1'>�Ѱ�</td>
	
						<%
						for i = 1 to iCol
							sLine = sLine & "<td bgcolor='#FFEEF9' align='center' class='TDCont' >" & vtot(i) & "</td>"	
							vtot(i) = 0
						next
						response.write sLine
						sLine = ""
						%>
						
					</tr>
	
				</table>

			</DIV>
		</td>
	</tr>
	
<% End if %>
	
</table>

<!-- #include virtual="/Include/Bottom.asp" -->