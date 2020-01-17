<!-- #include virtual="/Include/Common.asp" -->
<%
	Server.ScriptTimeout = 90000
	Response.ContentType = "application/vnd.ms-excel; name='My_Excel'"
	Call Response.AddHeader("Content-Disposition", "attachment; filename=완료목록_" &Date()& ".xls")	'바로저장하기
	Call Response.AddHeader("Content-Description", "ASP Generated Data")

%>
<%


	SQL = "select top 1 * from TB_Code WHERE	Codegroup = 'B20' AND UseYN = 'Y'"
	set RS = db.Execute(SQL)
	if rs.eof = false then
		sFileLinkURL = rs("Codename")
	else
		sFileLinkURL = "http://16.1.19.160:7001/amcriss/M_investigation/down.jsp?"
	end if

	curPage = request("curPage")
	FromDate = request("FromDate")
	ToDate = request("ToDate")
	whereCD1 = Trim(request("whereCD1"))
	whereCD2 = Trim(request("whereCD2"))
	whereCD3 = Trim(request("whereCD3"))
	whereCD4 = Trim(request("whereCD4"))
	whereCD5 = Trim(request("whereCD5"))
	whereCD6 = Trim(request("whereCD6"))
	whereCD7 = Trim(request("whereCD7"))
	ACLASS =  Trim(request("ACLASS"))
	BCLASS =  Trim(request("BCLASS"))
	CCLASS =  Trim(request("CCLASS"))
	CLASSNAME =  Trim(request("CLASSNAME"))
	if FromDate = "" then
		FromDate = left(date(),4) & "-01-01"
	end if

	'2. 쿼리조건절 셋팅
	pageSize = 10
	pageSector = 10
	if curPage = "" then curPage = 1 end If

	where1 = "FromDate=" & FromDate & "&ToDate=" & ToDate & "&whereCD1=" & whereCD1 & "&whereCD2=" & whereCD2 & "&whereCD3=" & whereCD3 & "&whereCD4=" & whereCD4 &"&QUERYGB="&QUERYGB&"&CLASSNAME="&CLASSNAME&"&whereCD6="&whereCD6&"&whereCD5="&whereCD5&"&whereCD7="&whereCD7
	where2 = "curPage=" & curPage & "&" & where1
	if whereCD7 = "Y" then
		sql_where = " 1=1 and processgb in ( '8') " '완료건
	else
		sql_where = " 1=1 and processgb in ( '100','9') " '완료건
	end if
	if FromDate = "" then
		SQL = "select * from armyinformix.dbo. where sourcedate >= '"& left(date(),7) & "-01' order by sourcedate desc"
		sql_where = sql_where & " and convert(char(10),processdate,121) >= '"& left(date(),7) & "-01'"
		if whereCD2 <> "" then
				sql_where = sql_where & " and receiptfactnum like '%" & whereCD2 & "%'" 
		end if

		if whereCD4 <> "" then
			sql_where = sql_where & " and dutyman = '" & whereCD4 & "'" 
		elseif whereCD5 <> "" then
				sql_where = sql_where & " and dutyman in ( select	id		from armyinformix.dbo.user1 where [name] like '%"&whereCD5&"%' ) "
		end if
		if whereCD3 <> "" then
				sql_where = sql_where & " and receiptfactnum in ( select	factnum		from armyinformix.dbo.factpeople where [name] like '%"&whereCD3&"%' ) "
		end if


		if CLASSNAME <> "" then
				sql_where = sql_where & " and dutyman in ( select	u.id		from armyinformix.dbo.pbudae p, armyinformix.dbo.user1 u  where p.name like '"&CLASSNAME&"%' and p.auth = u.unit ) "			
		end if
		if whereCD6 <> "" then
				sql_where = sql_where & " and nameoffact like '%" & whereCD6 & "%'" 
		end if

	else
		if ToDate = "" then ToDate = date() end if
		SQL = "select * from armyinformix.dbo. where convert(char(10),processdate,121) >= '"& FromDate & "' and convert(char(10),processdate,121) <= '"& ToDate & "'    order by sourcedate desc"
		sql_where = sql_where & " and convert(char(10),processdate,121) >= '"& FromDate & "' and convert(char(10),processdate,121) <= '"& ToDate & "'"
		if whereCD2 <> "" then
				sql_where = sql_where & " and receiptfactnum like '%" & whereCD2 & "%'" 
		end if

		if whereCD4 <> "" then
			sql_where = sql_where & " and dutyman = '" & whereCD4 & "'" 
		elseif whereCD5 <> "" then
				sql_where = sql_where & " and dutyman in ( select	id		from armyinformix.dbo.user1 where [name] like '%"&whereCD5&"%' ) "
		end if
		if whereCD3 <> "" then
				sql_where = sql_where & " and receiptfactnum in ( select	factnum		from armyinformix.dbo.factpeople where [name] like '%"&whereCD3&"%' ) "
		end if


		if CLASSNAME <> "" then
				sql_where = sql_where & " and dutyman in ( select	u.id		from armyinformix.dbo.pbudae p, armyinformix.dbo.user1 u  where p.name like '"&CLASSNAME&"%' and p.auth = u.unit ) "			
		end if
		if whereCD6 <> "" then
				sql_where = sql_where & " and nameoffact like '%" & whereCD6 & "%'" 
		end if
	end if


	sql_tb = "armyinformix.dbo.receiptfact"
	'sql_index = "index_desc(" & sql_tb & " IDX_TB_CALLHISTORY_JUBSEQ)"
	sql_field ="*"
	sql_orderby = "processdate DESC, receiptfactnum desc"


	'3. 쿼리 실행
	sql = db_getSqlWithPage(sql_tb, sql_index, sql_field, sql_where, sql_orderby, pageSize, curPage)
	set RsGBN = db.execute(sql)

	'4. Paging HTML 작성
	totalCount = db_getCount(db, sql_tb, sql_where)
	startRow = totalCount - pageSize * (curPage - 1)
	pageHtml = getPageHtml(pageSector, pageSize, totalCount, curPage, currentURL & "?" & where1)


%>
<table width="940"  border="1" cellpadding="0" cellspacing="0" bgcolor="#EFECE5" align="center" bordercolor="black" bordercolordark="white" bordercolorlight="black">
	<tr height="20" bgcolor="#EEF6FF" align="center">
		<td rowspan=2 class="TDCont">No</td>
		<td rowspan=2 class="TDCont">사건번호</td>
		<td rowspan=2 class="TDCont">사건명</td>
		<td rowspan=2 class="TDCont">모니터링<br>완료일자</td>
		<td colspan=3 class="TDCont">담당수사관</td>
		<td colspan=3 class="TDCont">모니터링실시자</td>
		<td rowspan=2 class="TDCont">평가<br>(점수)</td>


	</tr>

	<tr height="20" bgcolor="#EEF6FF" align="center">
		<td >소속</td>
		<td >계급</td>
		<td >성명</td>
		<td >피의자</td>
		<td >피해자</td>
		<td >지휘관</td>
	</tr>

	<tr><td colspan="15" height="1" bgcolor="#FFFFFF"></td></tr>



	<!--<tr id="cTR1" style="cursor:hand;" bgcolor="#ffffff" >

			<td align="center">1</td>
			<td align="center">9X09-01-0002</td>
			<td align="center">피의자(오후2시)</td>
			<td align="center">군기사고-폭행</td>
			<td align="center"><img src="/Images/Btn/icon_02.gif" style="cursor:hand;" align="absmiddle"></td>
			<td align="center">2009-01-01</td>
			<td align="center">ㅇㅇ군단헌병대</td>
			<td align="center">ㅇㅇ</td>
			<td align="center">ㅇㅇㅇ</td>
			<td align="center"><img src="/Images/Btn/icon_02.gif" style="cursor:hand;" align="absmiddle"></td>
			<td align="center"><img src="/Images/Btn/icon_02.gif" style="cursor:hand;" align="absmiddle"></td>
			<td align="center"><img src="/Images/Btn/icon_02.gif" style="cursor:hand;" align="absmiddle"></td>
			<td align="center">
				<img src="/Images/Btn/BtnIconModify.gif" title="수정" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_update('1111','UP');">
				<img src="/Images/Btn/BtnIconDel.gif" title="삭제" style="cursor:hand;" align="absmiddle"onclick="javascript:fn_del('1111','DEL');">
			</td>
	</tr>-->


<%


	i = 0

	do until RsGBN.eof

		i = i + 1

		'수사관정보
		'SQL = "	select name, class, (select contents from data where restrict = 104 and number = class ) as classname from user1 where id = '" & RsGBN("dutyman") & "'"

		SQL = "	select name, class, (select name from armyinformix.dbo.pbudae where auth = unit) as unitname from armyinformix.dbo.user1 where id = '" & RsGBN("dutyman") & "'"
		SET Rs = DB.execute(SQL)

		if Rs.eof = false then
			sName = rs("name")
			sClassName =  rs("class")
			sBudae = rs("unitname")
		else
			sName = ""
			sClassName = ""
			sBudae = ""
		end if
		Rs.close


		'피의자, 피해자,지휘관 연락처유무
		sYN1 = ""
		sYN2 = ""
		sYN3 = ""

		sql = "	select [name], monitorresult, monitorpoint from armyinformix.dbo.factpeople where factnum = '" & RsGBN("receiptfactnum") & "' and section2 in ('B11','413') "'and monitorresult = '9'"
		SET Rs = DB.execute(SQL)
		do until rs.eof
			if sYN1 = "" then
				if rs("monitorresult") = "9" then
					sYN1 = rs(0)&"<br>("&formatnumber(rs("monitorpoint"))&")"
				else
					sYN1 = rs(0)&"<br>(<font color='#ff0000'>"&db_getCodeName("B10",rs("monitorresult"))&"</font>)"
				end if
			else
				if rs("monitorresult") = "9" then
					sYN1 = sYN1&"<br>"&rs(0)&"<br>("&formatnumber(rs("monitorpoint"))&")"
				else
					sYN1 = sYN1&"<br>"&rs(0)&"<br>(<font color='#ff0000'>"&db_getCodeName("B10",rs("monitorresult"))&"</font>)"
				end if
			end if
			rs.movenext
		loop
		sql = "	select [name], monitorresult, monitorpoint  from armyinformix.dbo.factpeople where factnum = '" & RsGBN("receiptfactnum") & "' and section2 in ('B12','448')"' and monitorresult = '9'"
		SET Rs = DB.execute(SQL)
		do until rs.eof
			if sYN2 = "" then
				if rs("monitorresult") = "9" then
					sYN2 = rs(0)&"<br>("&formatnumber(rs("monitorpoint"))&")"
				else
					sYN2 = rs(0)&"<br>(<font color='#ff0000'>"&db_getCodeName("B10",rs("monitorresult"))&"</font>)"
				end if
			else
				if rs("monitorresult") = "9" then
					sYN2 = sYN2&"<br>"&rs(0)&"<br>("&formatnumber(rs("monitorpoint"))&")"
				else
					sYN2 = sYN2&"<br>"&rs(0)&"<br>(<font color='#ff0000'>"&db_getCodeName("B10",rs("monitorresult"))&"</font>)"
				end if
			end if
			rs.movenext
		loop
		sql = "	select [name], monitorresult, monitorpoint  from armyinformix.dbo.factpeople where factnum = '" & RsGBN("receiptfactnum") & "' and section2 in ('B15','447','449')"' and monitorresult = '9'"
		SET Rs = DB.execute(SQL)
		do until rs.eof
			if sYN3 = "" then
				if rs("monitorresult") = "9" then
					sYN3 = rs(0)&"<br>("&formatnumber(rs("monitorpoint"))&")"
				else
					sYN3 = rs(0)&"<br>(<font color='#ff0000'>"&db_getCodeName("B10",rs("monitorresult"))&"</font>)"
				end if
			else
				if rs("monitorresult") = "9" then
					sYN3 = sYN3&"<br>"&rs(0)&"<br>("&formatnumber(rs("monitorpoint"))&")"
				else
					sYN3 = sYN3&"<br>"&rs(0)&"<br>(<font color='#ff0000'>"&db_getCodeName("B10",rs("monitorresult"))&"</font>)"
				end if
			end if
			rs.movenext
		loop
		set Rs = nothing

		receiptfactnum = RsGBN("receiptfactnum") 
%>

	<tr id="cTR1" style="cursor:hand;" bgcolor="#ffffff" onmouseover="this.style.background='#FFFCE7'" bgcolor="#FFFFFF" onmouseout="this.style.background='#FFFFFF'">
			<td align="center" class="TDCont" width=40><%=startRow%></td>
			<td align="center" class="TDCont" nowrap><a href="javascript:fn_update('<%=receiptfactnum%>','UP');"><%=RsGBN("receiptfactnum")%></a></td>
			<td align="left" class="TDCont" title="<%=RsGBN("nameoffact")%>" nowrap>&nbsp;<a href="javascript:fn_update('<%=receiptfactnum%>','UP');"><%=CutString(RsGBN("nameoffact"), 20, "...")%></a></td>
			<td align="center" class="TDCont"><%=left(RsGBN("processdate"),10)%></td>
			<td align="center" class="TDCont"><%=sBudae%></td>
			<td align="center" class="TDCont"><%=sClassName%></td>
			<td align="center" class="TDCont"><%=sName%></td>
			<td align="center"><%=sYN1%></td>
			<td align="center"><%=sYN2%></td>
			<td align="center"><%=sYN3%></td>
			<% if RsGBN("monitorpoint") <> "" then%>
				<td align="center" class="TDCont"><%=formatNumber(RsGBN("monitorpoint"),2)%></td>
			<% else%>
				<td align="center" class="TDCont">&nbsp;</td>
			<% end if %>
	</tr>

<%
		startRow = startRow - 1

		RsGBN.MoveNext
	LOOP
	RsGBN.CLOSE
	SET RsGBN = Nothing
%>

</table>

<!-- #include virtual="/Include/Bottom.asp" -->