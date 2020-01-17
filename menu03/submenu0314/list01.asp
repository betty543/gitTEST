<!-- #include virtual="/Include/Top.asp" -->

<%
'####### 파라미터 ##################################################################################
QueryYN = request("QueryYN")
FromDate = request("FromDate")
ToDate = request("ToDate")
whereCD3 = Trim(request("whereCD3"))
whereCD7 = Trim(request("whereCD7"))

dim	vtot(100)

If QueryYN = "" Then
	whereCD3 = "1"
End if

if FromDate = "" then FromDate =left(Date(),7)&"-01" end If
if ToDate = "" then ToDate=date() end If

pageWHERE = "QueryYN="&QueryYN&"&FromDate="&FromDate&"&ToDate="&ToDate&"&whereCD3="&whereCD3&"&whereCD7="&whereCD7
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
						<td class="TDCont">조회기간 :
							<input value="<%=FromDate%>" name="FromDate" readonly type="text" size="10" onfocus="setFocusColor(this); new CalendarFrame.Calendar(this);" onblur="setOutColor(this);"> ~
							<input value="<%=ToDate%>" name="ToDate" readonly type="text" size="10" onfocus="setFocusColor(this); new CalendarFrame.Calendar(this);" onblur="setOutColor(this);">
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

	<table border="0" cellpadding="0" cellspacing="0" align="center">
		<tr>
			<td>

				<!--계급별-->
				<table width="100%" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>

				<div style="overflow-x:auto;overflow-y:hidden;margin:0;width:1200px;">
					<table width="2000" border="1" cellpadding="0" cellspacing="0" bgcolor="#EFECE5" align="center" bordercolor="black" bordercolordark="white" bordercolorlight="black">
					<colgroup>
						<% for i = 0 to 50 %>
							<col width="100px" />
						<% next %>
					</colroup>
						<tr height="30">
							<td align="left" bgcolor="#FFFFFF" class="TDCont" colspan="300">&nbsp;<img src="/Images/dot_01.gif" >&nbsp;<b><font color="#ff00ff"></font> 1. 계급별</b></td>
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
						lineA = lineA &	"	<td align=""center"" class=""TDCont"" with=""150px"" rowspan=""3"">상담관</td>"

						subSql = " select '총계' "

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
								lineA = lineA & "<td align=""center"" class=""TDCont"" rowspan=""3"">계</td>"
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
									lineB = lineB & "<td align=""center"" class=""TDCont"" rowspan=""2"">계</td>"
									subSql = subSql & " ,count(case when Level_B = '" & arrRs2(0,ii) & "' then 1 else null end) sum_" & i & "_" & ii & " "
								end if

								for iii = 0 to arrRc3

									subSql = subSql & " ,count(case when Level_B = '" & arrRs3(0,iii) & "' and Level_C = '" & arrRs3(1,iii) & "' and Level_D = '" & arrRs3(2,iii) & "' then 1 else null end) col_" & i & "_" & ii & "_" & iii & " "

									lineC = lineC &	"<td align=""center"" class=""TDCont"">" & arrRs3(3,iii) & "</td>"

									if iii = arrRc3 and arrRc3 > 0 then
										lineC = lineC & "<td align=""center"" class=""TDCont"">계</td>"
										subSql = subSql & " ,count(case when Level_B = '" & arrRs3(0,iii) & "' and Level_C = '" & arrRs3(1,iii) & "' then 1 else null end) sum_" & i & "_" & ii & "_" & iii & " "
									end if

								next

							next

						next

						lineA = lineA &	"</tr>"

						response.write	lineA
						response.write	"<tr bgcolor=""#EEF6FF"">" & lineB & "</tr>"
						response.write	"<tr bgcolor=""#EEF6FF"">" & lineC & "</tr>"

						subSql = subSql & "	 from tb_lifecallhistory_ob  where jubdate >= '" & FromDate & "' AND		jubdate <= '" & ToDate & "' /*and CHANNELGB_B in ('Q01','Q03','Q05','Q07','Q09','Q99')*/ "
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

						for i = 0 to arrRc2
							response.write	"<tr bgcolor=""#FFEEF9"">"
							for ii = 0 to arrRc1
								response.write	"<td align=""center"" class=""TDCont"">" & arrRs(ii,i) & "</td>"
							next
							response.write	"</tr>"
						next
						%>

						</tr>
					</table>
					<br /><Br />
				</div>

				<!--부대별-->
				<table width="100%" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>

				<div style="overflow-x:auto;overflow-y:hiddne; margin:0;width:1200px;">
					<table width="2200"  border="1" cellpadding="0" cellspacing="0" bgcolor="#EFECE5" align="center" bordercolor="black" bordercolordark="white" bordercolorlight="black">
						<tr height="30">
							<td align="left" bgcolor="#FFFFFF" class="TDCont" colspan="300">&nbsp;<img src="/Images/dot_01.gif" >&nbsp;<b><font color="#ff00ff"></font> 2. 부대별</b></td>
						</tr>

						<%
						SQL = " select * from tb_armyinfo where aclass < 'O' and bclass is null order by aclass "
						set Rs = db.execute(SQL)

						firstLine = "<tr bgcolor='#EEF6FF'>"
						firstLine = firstLine & "<td align='center' class='TDCont'  width='150' colspan= '2' rowspan='2'>상담관</td>"
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
									secondLine = secondLine & "<td bgcolor='#EEF6FF' align='center' class='TDCont'  width='50'>총계</td>"
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
						firstLine = firstLine & "<td bgcolor='#EEF6FF' align='center' class='TDCont'  width='50' rowspan='2'>총계</td></tr>"
						
						response.write firstLine
						response.write secondLine

						iColTot = iColTot + 1
						execSQL = execSQL & " ,case when SOSOKGB_A + SOSOKGB_B in ('" & replace(inTotValue,",","','") & "') then 1 else 0 end col" & iColTot & ""
						execSQL = execSQL & "	 from tb_lifecallhistory_ob  where jubdate >= '" & FromDate & "' AND		jubdate <= '" & ToDate & "'"

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
							<td align='center' class='TDCont'  width='300' colspan='2'>총계</td>
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
					<br /><Br />
				</div>
				
				<!--요일별-->
				<table width="1200" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
				
				<table width="1200"  border="1" cellpadding="0" cellspacing="0" bgcolor="#EFECE5" align="center" bordercolor="black" bordercolordark="white" bordercolorlight="black">
					<tr height="30">
						<td align="left" bgcolor="#FFFFFF" class="TDCont" colspan="9">&nbsp;<img src="/Images/dot_01.gif" >&nbsp;<b><font color="#ff00ff"></font> 3. 요일별</b></td>
					</tr>
					<tr bgcolor='#EEF6FF'>
						<td align='center' class='TDCont' width='150'>구분</td>
						<td align='center' class='TDCont' width='150'>월</td>
						<td align='center' class='TDCont' width='150'>화</td>
						<td align='center' class='TDCont' width='150'>수</td>
						<td align='center' class='TDCont' width='150'>목</td>
						<td align='center' class='TDCont' width='150'>금</td>
						<td align='center' class='TDCont' width='150' >토</td>
						<td align='center' class='TDCont' width='150' >일</td>
						<td align='center' class='TDCont' width='150'>총계</td>
					</tr>
					<%
					'상담관별 총계
					SQL = "select * from ( SELECT	'1' gubun, incode, count(incode) cnt FROM	TB_lifecallhistory_ob where jubdate >= '" & FromDate & "'"
					SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and datepart(weekday,jubdate) = 2 group by incode" '
					SQL = SQL & "	union all SELECT	'2' gubun, incode, count(incode) cnt FROM	TB_lifecallhistory_ob where jubdate >= '" & FromDate & "'"
					SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and datepart(weekday,jubdate) = 3 group by incode" '
					SQL = SQL & "	union all SELECT	'3' gubun, incode, count(incode) cnt  FROM	TB_lifecallhistory_ob where jubdate >= '" & FromDate & "'"
					SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and datepart(weekday,jubdate) = 4 group by incode" '
					SQL = SQL & "	union all SELECT	'4' gubun, incode, count(incode) cnt  FROM	TB_lifecallhistory_ob where jubdate >= '" & FromDate & "'"
					SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and datepart(weekday,jubdate) = 5 group by incode" '
					SQL = SQL & "	union all SELECT	'5' gubun, incode, count(incode) cnt  FROM	TB_lifecallhistory_ob where jubdate >= '" & FromDate & "'"
					SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and datepart(weekday,jubdate) = 6 group by incode" '
					SQL = SQL & "	union all SELECT	'6' gubun, incode, count(incode) cnt  FROM	TB_lifecallhistory_ob where jubdate >= '" & FromDate & "'"
					SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and datepart(weekday,jubdate) = 7 group by incode" '
					SQL = SQL & "	union all SELECT	'7' gubun, incode, count(incode) cnt  FROM	TB_lifecallhistory_ob where jubdate >= '" & FromDate & "'"
					SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and datepart(weekday,jubdate) = 1 group by incode" '
					SQL = SQL & "	union all SELECT	'8' gubun, incode, count(incode) cnt  FROM	TB_lifecallhistory_ob where jubdate >= '" & FromDate & "'"
					SQL = SQL & "	AND		jubdate <= '" & ToDate & "' group by incode) a order by incode, gubun" '
				
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
							<td align='center' class='TDCont' width='150'><%=db_getUserName(incode)%></td>
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

					'상담관별 총계
					SQL = "select * from ( SELECT	'1' gubun, count(incode) cnt FROM	TB_lifecallhistory_ob where jubdate >= '" & FromDate & "'"
					SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and datepart(weekday,jubdate) = 2" '
					SQL = SQL & "	union all SELECT	'2' gubun, count(incode) cnt FROM	TB_lifecallhistory_ob where jubdate >= '" & FromDate & "'"
					SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and datepart(weekday,jubdate) = 3" '
					SQL = SQL & "	union all SELECT	'3' gubun, count(incode) cnt  FROM	TB_lifecallhistory_ob where jubdate >= '" & FromDate & "'"
					SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and datepart(weekday,jubdate) = 4" '
					SQL = SQL & "	union all SELECT	'4' gubun, count(incode) cnt  FROM	TB_lifecallhistory_ob where jubdate >= '" & FromDate & "'"
					SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and datepart(weekday,jubdate) = 5" '
					SQL = SQL & "	union all SELECT	'5' gubun, count(incode) cnt  FROM	TB_lifecallhistory_ob where jubdate >= '" & FromDate & "'"
					SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and datepart(weekday,jubdate) = 6" '
					SQL = SQL & "	union all SELECT	'6' gubun, count(incode) cnt  FROM	TB_lifecallhistory_ob where jubdate >= '" & FromDate & "'"
					SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and datepart(weekday,jubdate) = 7" '
					SQL = SQL & "	union all SELECT	'7' gubun, count(incode) cnt  FROM	TB_lifecallhistory_ob where jubdate >= '" & FromDate & "'"
					SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and datepart(weekday,jubdate) = 1" '
					SQL = SQL & "	union all SELECT	'8' gubun, count(incode) cnt  FROM	TB_lifecallhistory_ob where jubdate >= '" & FromDate & "'"
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
						<td align='center' class='TDCont'  width='150'>계</td>
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
				
				<!--시간대별-->
				<table width="1200" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
				
				<table width="1200"  border="1" cellpadding="0" cellspacing="0" bgcolor="#EFECE5" align="center" bordercolor="black" bordercolordark="white" bordercolorlight="black">
					<tr height="30">
						<td align="left" bgcolor="#FFFFFF" class="TDCont" colspan="26">&nbsp;<img src="/Images/dot_01.gif" >&nbsp;<b><font color="#ff00ff"></font> 4. 시간대별</b></td>
					</tr>
					<tr bgcolor='#EEF6FF'>
						<td align='center' class='TDCont'  width='150'>구분</td>
						<%
						for i = 0 to 23
							if i < 10 then
								sHourname = "0" & i & "시"
							else
								sHourname = i & "시"
							end if
							%>
							<td align='center' class='TDCont'  width='150'><%=sHourname%></td>
							<%
						next
						%>
						<td align='center' class='TDCont' width='150'>총계</td>
					</tr>
					<%
					SQL = " SELECT	incode"
					for i = 0 to 23
						SQL = SQL & "			, case when datepart(hour,jubtime) = " & i & " then 1 else 0 end col" & i
					next
					SQL = SQL & "			, 1 col" & i
					SQL = SQL & "	FROM	TB_lifecallhistory_ob where jubdate >= '" & FromDate & "'	AND		jubdate <= '" & ToDate & "' "
				
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
						<td align='center' class='TDCont'  width='300' colspan='1'>총계</td>
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

				<!--통화시간별-->
				<table width="1200" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
				
				<table width="1200"  border="1" cellpadding="0" cellspacing="0" bgcolor="#EFECE5" align="center" bordercolor="black" bordercolordark="white" bordercolorlight="black">
					<tr height="30">
						<td align="left" bgcolor="#FFFFFF" class="TDCont" colspan="11">&nbsp;<img src="/Images/dot_01.gif" >&nbsp;<b><font color="#ff00ff"></font> 5. 통화시간별</b></td>
					</tr>
					<tr bgcolor='#EEF6FF'>
						<td align='center' class='TDCont' width='150'>구분</td>
						<td align='center' class='TDCont' width='150'>1분미만</td>
						<td align='center' class='TDCont' width='150'>1-5분</td>
						<td align='center' class='TDCont' width='150'>6-10분</td>
						<td align='center' class='TDCont' width='150'>11-20분</td>
						<td align='center' class='TDCont' width='150'>21-30분</td>
						<td align='center' class='TDCont' width='150'>31-40분</td>
						<td align='center' class='TDCont' width='150'>41-50분</td>
						<td align='center' class='TDCont' width='150'>51-60분</td>
						<td align='center' class='TDCont' width='150'>60분이상</td>
						<td align='center' class='TDCont' width='150'>총계</td>
					</tr>
					<%
					SQL = "select * from ( SELECT	'01' gubun, incode, count(incode) cnt FROM	TB_lifecallhistory_ob where jubdate >= '" & FromDate & "'"
					SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and calltime < 60 group by incode" '
				
					SQL = SQL & "	union all SELECT	'02' gubun, incode, count(incode) cnt FROM	TB_lifecallhistory_ob where jubdate >= '" & FromDate & "'"
					SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and  calltime >=60 and calltime <=300 group by incode" '
					SQL = SQL & "	union all SELECT	'03' gubun, incode, count(incode) cnt FROM	TB_lifecallhistory_ob where jubdate >= '" & FromDate & "'"
					SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and  calltime >=301 and calltime <=600 group by incode" '
					SQL = SQL & "	union all SELECT	'04' gubun, incode, count(incode) cnt FROM	TB_lifecallhistory_ob where jubdate >= '" & FromDate & "'"
					SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and  calltime >=601 and calltime <=1200 group by incode" '
				
					SQL = SQL & "	union all SELECT	'05' gubun, incode, count(incode) cnt  FROM	TB_lifecallhistory_ob where jubdate >= '" & FromDate & "'"
					SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and  calltime >=1201 and calltime <=1800 group by incode" '
					SQL = SQL & "	union all SELECT	'06' gubun, incode, count(incode) cnt  FROM	TB_lifecallhistory_ob where jubdate >= '" & FromDate & "'"
					SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and  calltime >=1801 and calltime <=2400 group by incode" '
					SQL = SQL & "	union all SELECT	'07' gubun, incode, count(incode) cnt  FROM	TB_lifecallhistory_ob where jubdate >= '" & FromDate & "'"
					SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and  calltime >=2401 and calltime <=3000 group by incode" '
					SQL = SQL & "	union all SELECT	'08' gubun, incode, count(incode) cnt  FROM	TB_lifecallhistory_ob where jubdate >= '" & FromDate & "'"
					SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and  calltime >=3001 and calltime <=3600 group by incode" '
					SQL = SQL & "	union all SELECT	'09' gubun, incode, count(incode) cnt  FROM	TB_lifecallhistory_ob where jubdate >= '" & FromDate & "'"
					SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and calltime >=3601 group by incode" '
					SQL = SQL & "	union all SELECT	'10' gubun, incode, count(incode) cnt  FROM	TB_lifecallhistory_ob where jubdate >= '" & FromDate & "'"
					SQL = SQL & "	AND		jubdate <= '" & ToDate & "'  group by incode) a order by incode, gubun" '
				
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

					SQL = "select * from ( SELECT	'01' gubun, incode, count(incode) cnt FROM	TB_lifecallhistory_ob where jubdate >= '" & FromDate & "'"
					SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and calltime <60 group by incode" '
					SQL = SQL & "	union all SELECT	'02' gubun, incode, count(incode) cnt FROM	TB_lifecallhistory_ob where jubdate >= '" & FromDate & "'"
					SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and calltime >=60 and calltime <=300 group by incode" '
					SQL = SQL & "	union all SELECT	'03' gubun, incode, count(incode) cnt FROM	TB_lifecallhistory_ob where jubdate >= '" & FromDate & "'"
					SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and calltime >=301 and calltime <=600 group by incode" '
					SQL = SQL & "	union all SELECT	'04' gubun, incode, count(incode) cnt FROM	TB_lifecallhistory_ob where jubdate >= '" & FromDate & "'"
					SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and  calltime >=601 and calltime <=1200 group by incode" '
					SQL = SQL & "	union all SELECT	'05' gubun, incode, count(incode) cnt  FROM	TB_lifecallhistory_ob where jubdate >= '" & FromDate & "'"
					SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and  calltime >=1201 and calltime <=1800 group by incode" '
					SQL = SQL & "	union all SELECT	'06' gubun, incode, count(incode) cnt  FROM	TB_lifecallhistory_ob where jubdate >= '" & FromDate & "'"
					SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and  calltime >=1801 and calltime <=2400 group by incode" '
					SQL = SQL & "	union all SELECT	'07' gubun, incode, count(incode) cnt  FROM	TB_lifecallhistory_ob where jubdate >= '" & FromDate & "'"
					SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and  calltime >=2401 and calltime <=3000 group by incode" '
					SQL = SQL & "	union all SELECT	'08' gubun, incode, count(incode) cnt  FROM	TB_lifecallhistory_ob where jubdate >= '" & FromDate & "'"
					SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and  calltime >=3001 and calltime <=3600 group by incode" '
					SQL = SQL & "	union all SELECT	'09' gubun, incode, count(incode) cnt  FROM	TB_lifecallhistory_ob where jubdate >= '" & FromDate & "'"
					SQL = SQL & "	AND		jubdate <= '" & ToDate & "' and  calltime >=3601 group by incode" '
					SQL = SQL & "	union all SELECT	'10' gubun, incode, count(incode) cnt  FROM	TB_lifecallhistory_ob where jubdate >= '" & FromDate & "'"
					SQL = SQL & "	AND		jubdate <= '" & ToDate & "'  group by incode) a order by incode, gubun" '

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
						<td align='center' class='TDCont'  width='150'>총계</td>
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

				<!--원상담자와의관계-->
				<table width="1200" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
				
				<table width="1200"  border="1" cellpadding="0" cellspacing="0" bgcolor="#EFECE5" align="center" bordercolor="black" bordercolordark="white" bordercolorlight="black">
					<tr height="30">
						<td align="left" bgcolor="#FFFFFF" class="TDCont" colspan="12">&nbsp;<img src="/Images/dot_01.gif" >&nbsp;<b><font color="#ff00ff"></font> 6. 원상담자와의관계</b> </td>
					</tr>
					<tr bgcolor='#EEF6FF'>
						<td align='center' class='TDCont' width='150'>구분</td>
						<%
						SQL = "		SELECT	*	FROM	TB_CODE WHERE CODEGROUP = 'C14' AND USEYN = 'Y' ORDER BY CODE"
						execSQL = " select incode"
						set Rs = db.execute(SQL)
						iCol = 0
						do until rs.eof
							iCol = iCol + 1
							execSQL = execSQL & ", case when REQUESTERGB = '" & rs("code") & "' then 1 else 0 end col" & iCol
							%>
							<td align='center' class='TDCont' width='150'><%=rs("CodeName")%></td>
							<%
							rs.movenext
						loop
						iCol = iCol + 1
						execSQL = execSQL & ", 1 col" & iCol & " from TB_lifecallhistory_ob where jubdate >= '" & FromDate & "'"
						execSQL = execSQL & "	AND		jubdate <= '" & ToDate & "' "
						%>
						<td align='center' class='TDCont' width='150'>총계</td>
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
						<td align='center' class='TDCont'  width='300' colspan='1'>총계</td>
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
	
				<!--후속확인-->
				<table width="1200" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
				
				<table width="1200"  border="1" cellpadding="0" cellspacing="0" bgcolor="#EFECE5" align="center" bordercolor="black" bordercolordark="white" bordercolorlight="black">
					<tr height="30">
						<td align="left" bgcolor="#FFFFFF" class="TDCont" colspan="12">&nbsp;<img src="/Images/dot_01.gif" >&nbsp;<b><font color="#ff00ff"></font> 7. 후속확인</b> </td>
					</tr>
					<tr bgcolor='#EEF6FF'>
						<td align='center' class='TDCont' width='150'>구분</td>
						<%
						SQL = "		SELECT	*	FROM	TB_CODE WHERE CODEGROUP = 'C13' AND USEYN = 'Y' ORDER BY CODE"
						execSQL = " select incode"
						set Rs = db.execute(SQL)
						iCol = 0
						do until rs.eof
							iCol = iCol + 1
							execSQL = execSQL & ", case when CALLKIND_B = '" & rs("code") & "' then 1 else 0 end col" & iCol
							%>
							<td align='center' class='TDCont' width='150'><%=rs("CodeName")%></td>
							<%
							rs.movenext
						loop
						iCol = iCol + 1
						execSQL = execSQL & ", 1 col" & iCol & " from TB_lifecallhistory_ob where jubdate >= '" & FromDate & "'"
						execSQL = execSQL & "	AND		jubdate <= '" & ToDate & "'"
						%>
						<td align='center' class='TDCont' width='150'>총계</td>
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
						<td align='center' class='TDCont'  width='300' colspan='1'>총계</td>
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
</table>

<% End if %>

<!-- #include virtual="/Include/Bottom.asp" -->