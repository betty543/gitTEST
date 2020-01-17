<!-- #include virtual="/Include/Common.asp" -->
<%
	Server.ScriptTimeout = 90000
	Response.ContentType = "application/vnd.ms-excel; name='My_Excel'"
	Call Response.AddHeader("Content-Disposition", "attachment; filename=대기목록_" &Date()& ".xls")	'바로저장하기
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
	whereCD4 = ""

	CLASSNAME = Trim(request("CLASSNAME"))

	if FromDate = "" then
		FromDate = left(date(),4) & "-01-01"
	end if

	'2. 쿼리조건절 셋팅
	pageSize = 10
	pageSector = 10
	if curPage = "" then curPage = 1 end If

	where1 = "FromDate=" & FromDate & "&ToDate=" & ToDate & "&whereCD1=" & whereCD1 & "&whereCD2=" & whereCD2 & "&whereCD3=" & whereCD3 & "&whereCD4=" & whereCD4 &"&QUERYGB="&QUERYGB&"&CLASSNAME="&CLASSNAME&"&whereCD6="&whereCD6&"&whereCD5="&whereCD5
	where2 = "curPage=" & curPage & "&" & where1
	''sql_where = " 1=1 and handlesection not in ('01','100','0000000000')" '완료건
	sql_where = " 1=1 and filecnt > 0 and ( processgb is null or processgb ='') " '완료건
	if FromDate = "" then		
		SQL = "select * from armyinformix.dbo.receiptfact where inputdate >= '"& left(date(),7) & "-01' order by inputdate desc"
		sql_where = sql_where & " and inputdate >= '"& left(date(),7) & "-01'"
		sql_where = sql_where & " and dutyman is not null and dutyman <> ''"
		if whereCD4 <> "" then
			sql_where = sql_where & " and dutyman = '" & whereCD4 & "'" 
		elseif whereCD5 <> "" then
				sql_where = sql_where & " and dutyman in ( select	id		from armyinformix.dbo.user1 where [name] like '%"&whereCD5&"%' ) "
		end if
		if whereCD3 <> "" then
				sql_where = sql_where & " and receiptfactnum in ( select	factnum		from armyinformix.dbo.factpeople where [name] like '%"&whereCD3&"%' ) "
		end if
		if whereCD2 <> "" then
				sql_where = sql_where & " and receiptfactnum like '%" & whereCD2 & "%'" 
		end if
		if CLASSNAME <> "" then
				sql_where = sql_where & " and dutyman in ( select	u.id		from armyinformix.dbo.pbudae p, armyinformix.dbo.user1 u  where p.name like '"&CLASSNAME&"%' and p.auth = u.unit ) "			
		end if
		if whereCD6 <> "" then
				sql_where = sql_where & " and nameoffact like '%" & whereCD6 & "%'" 
		end if
	else
		if ToDate = "" then ToDate = date() end if
		SQL = "select * from armyinformix.dbo.receiptfact where inputdate >= '"& FromDate & "' and inputdate <= '"& ToDate & "'    order by inputdate desc"
		sql_where = sql_where & " and inputdate >= '"& FromDate & "' and inputdate <= '"& ToDate & "'"
		sql_where = sql_where & " and dutyman is not null and dutyman <> ''"
		if whereCD4 <> "" then
			sql_where = sql_where & " and dutyman = '" & whereCD4 & "'" 
		elseif whereCD5 <> "" then
				sql_where = sql_where & " and dutyman in ( select	id		from armyinformix.dbo.user1 where [name] like '%"&whereCD5&"%' ) "
		end if
		if whereCD3 <> "" then
				sql_where = sql_where & " and receiptfactnum in ( select	factnum		from armyinformix.dbo.factpeople where [name] like '%"&whereCD3&"%' ) "
		end if
		if whereCD2 <> "" then
			sql_where = sql_where & " and receiptfactnum like '%" & whereCD2 & "%'" 
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
	sql_orderby = "inputdate DESC, receiptfactnum desc"

	'3. 쿼리 실행
	sql = db_getSqlWithPage(sql_tb, sql_index, sql_field, sql_where, sql_orderby, pageSize, curPage)
	set RsGBN = db.execute(sql)

	'Response.Write sql

	'4. Paging HTML 작성
	totalCount = db_getCount(db, sql_tb, sql_where)
	startRow = totalCount - pageSize * (curPage - 1)
	pageHtml = getPageHtml(pageSector, pageSize, totalCount, curPage, currentURL & "?" & where1)

%>
<table width="940"  border="1" cellpadding="0" cellspacing="0" bgcolor="#EFECE5" align="center" bordercolor="black" bordercolordark="white" bordercolorlight="black">
	<tr height="20" bgcolor="#EEF6FF" align="center">
		<td rowspan=2 class="TDCont" align='center'>No</td>
		<td rowspan=2 class="TDCont" align='center'>사건번호</td>
		<td rowspan=2 class="TDCont" align='center'>사건명</td>
		<td rowspan=2 class="TDCont" align='center'>첨부파일목록</td>
		<td rowspan=2 class="TDCont" align='center'>출처일자</td>
		<td colspan=3 class="TDCont" align='center'>담당수사관</td>
		<td colspan=3 class="TDCont" align='center'>연락처유무</td>

	</tr>
	<tr height="20" bgcolor="#EEF6FF" align="center">
		<td class="TDCont">소속</td>
		<td class="TDCont">계급</td>
		<td class="TDCont">성명</td>
		<td >피의자</td>
		<td >피해자</td>
		<td >관계자</td>
	</tr>
	<tr><td colspan="16" height="1" bgcolor="#FFFFFF"></td></tr>

<%

	i = 0

	do until RsGBN.eof

		i = i + 1
		'관련파일명


		if mid(RsGBN("receiptfactnum"),6,2) = "10" then		

			'if sFileNum <> "" then
				SQL = "	select top 3 * from armyinformix.dbo.monitorfile where receiptfactnum in ('" & RsGBN("receiptfactnum") & "') order by filenum"
				SET Rs = DB.execute(SQL)
				sfilename = ""
				do until Rs.eof
					if sfilename = "" then
						sfilename = "<a href='http://16.1.150.146:9080/vivid/JspSource/file/fileDownload.jsp?fileNumber="& rs("filenum")&"'>" & CutString(rs("filename"), 10, "...") & "</a>"
					else
						sfilename = sfilename & "<br><a href='http://16.1.150.146:9080/vivid/JspSource/file/fileDownload.jsp?fileNumber="& rs("filenum")&"'>" & CutString(rs("filename"), 10, "...") & "</a>"
					end if

					Rs.movenext
				loop
			'end if

		else

			SQL = "	select top 3 * from armyinformix.dbo.monitorfile where receiptfactnum = '" & RsGBN("receiptfactnum") & "' order by filenum"
			SET Rs = DB.execute(SQL)
			sfilename = ""
			do until Rs.eof
				if sfilename = "" then
					sfilename = "<a href='"&sFileLinkURL&"filename="& rs("filename")&"'>" & CutString(rs("filename"), 16, "...") & "</a>"
				else
					sfilename = sfilename & "<br><a href='"&sFileLinkURL&"filename="& rs("filename")&"'>" & CutString(rs("filename"), 16, "...") & "</a>"
				end if

				Rs.movenext
			loop
		end if
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
		set Rs = nothing

		'피의자, 피해자,지휘관 연락처유무
		sYN1 = ""
		sYN2 = ""
		sYN3 = ""

		sql = "	select count(*) from armyinformix.dbo.factpeople where factnum = '" & RsGBN("receiptfactnum") & "' and section2 in ('B11','413')"
		SET Rs = DB.execute(SQL)
		if rs(0)>0 then
			sYN1 = "Y"
		end if
		sql = "	select count(*) from armyinformix.dbo.factpeople where factnum = '" & RsGBN("receiptfactnum") & "' and section2 in ('B12','448')"
		SET Rs = DB.execute(SQL)
		if rs(0)>0 then
			sYN2 = "Y"
		end if
		sql = "	select count(*) from armyinformix.dbo.factpeople where factnum = '" & RsGBN("receiptfactnum") & "' and section2 in ('B15','447','449','')"
		SET Rs = DB.execute(SQL)
		if rs(0)>0 then
			sYN3 = "Y"
		end if

		receiptfactnum = RsGBN("receiptfactnum") 
		contents = RsGBN("contents") 

%>

	<tr id="cTR1" style="cursor:hand;" bgcolor="#ffffff" onmouseover="this.style.background='#FFFCE7'" bgcolor="#FFFFFF" onmouseout="this.style.background='#FFFFFF'">
			<td align="center" class="TDCont" width=40 nowrap><%=startRow%></td>
			<td align="center" class="TDCont" nowrap title="<%=contents%>"><a href="javascript:fn_new('<%=receiptfactnum%>','UP');"><%=RsGBN("receiptfactnum")%></a></td>
			<td align="left" class="TDCont" title="<%=RsGBN("nameoffact")%>" nowrap>&nbsp;<a href="javascript:fn_new('<%=receiptfactnum%>','UP');"><%=CutString(RsGBN("nameoffact"), 10, "...")%></a></td>
			<td align="left" class="TDCont" nowrap><%=sfilename%></td>
			<td align="center" class="TDCont" nowrap><%=RsGBN("inputdate")%></td>
			<td align="center" class="TDCont" nowrap><%=sBudae%></td>
			<td align="center" class="TDCont" nowrap><%=sClassName%></td>
			<td align="center" class="TDCont" nowrap><%=sName%></td>
			<td align="center" class="TDCont" nowrap><%=sYN1%></td>
			<td align="center"><%=sYN2%></td>
			<td align="center"><%=sYN3%></td>

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