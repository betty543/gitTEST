<!-- #include virtual="/Include/Top.asp" -->
<!-- #include virtual="/Include/PopLayer.asp" -->
<!-- #include virtual="/Include/DBConnection_info.asp" -->

<%


	SQL = "select top 1 * from TB_Code WHERE	Codegroup = 'B20' AND UseYN = 'Y'"
	set RS = db.Execute(SQL)
	if rs.eof = false then
		sFileLinkURL = rs("Codename")
	else
		sFileLinkURL = "http://16.1.19.160:7001/amcriss/M_investigation/down.jsp?"
	end if

	dim CLASSNAME
	curPage = request("curPage")
	FromDate = request("FromDate")
	ToDate = request("ToDate")
	whereCD1 = Trim(request("whereCD1"))
	whereCD2 = Trim(request("whereCD2"))
	whereCD3 = Trim(request("whereCD3"))
	whereCD4 = Trim(request("whereCD4"))
	whereCD6 = Trim(request("whereCD6"))
	QUERYGB =  Trim(request("QUERYGB"))
	ACLASS =  Trim(request("ACLASS"))
	BCLASS =  Trim(request("BCLASS"))
	CCLASS =  Trim(request("CCLASS"))
	CLASSNAME =  Trim(request("CLASSNAME"))


	if QUERYGB = "" then
		QUERYGB = "1"
	end if

	'2. 쿼리조건절 셋팅
	pageSize = 10
	pageSector = 10
	if curPage = "" then curPage = 1 end If

	where1 = "FromDate=" & FromDate & "&ToDate=" & ToDate & "&whereCD1=" & whereCD1 & "&whereCD2=" & whereCD2 & "&whereCD3=" & whereCD3 & "&whereCD4=" & whereCD4 &"&QUERYGB="&QUERYGB&"&CLASSNAME="&CLASSNAME&"&whereCD6="&whereCD6
	where2 = "curPage=" & curPage & "&" & where1
	'sql_where = " 1=1 and handlesection not in ('99')" '진행중
	sql_where = " 1=1 and processgb in ('1','2')" '완료건

	if FromDate = "" then
		
		SQL = "select * from armyinformix.dbo.receiptfact where inputdate >= '"& left(date(),4) & "-01-01' order by inputdate desc"
		sql_where = sql_where & " and inputdate >= '"& left(date(),4) & "-01-01'"
		sql_where = sql_where & " and dutyman is not null and dutyman <> ''"
		if whereCD3 <> "" then
			sql_where = sql_where & " and dutyman = '" & whereCD3 & "'" 
		elseif whereCD4 <> "" then
				sql_where = sql_where & " and dutyman in ( select	id		from armyinformix.dbo.user1 where [name] like '%"&whereCD4&"%' ) "
		end if
		if whereCD2 <> "" then
				sql_where = sql_where & " and receiptfactnum in ( select	factnum		from armyinformix.dbo.factpeople where [name] like '%"&whereCD2&"%' ) "
		end if

		if whereCD1 <> "" then
				sql_where = sql_where & " and receiptfactnum like '%" & whereCD1 & "%'" 
		end if

		if QUERYGB = "1" then
				sql_where = sql_where & " and receiptfactnum in ( select	factnum		from armyinformix.dbo.factpeople where left(ReserveDate,10) = convert(char(10),getdate(),121))"
		end if

		if CLASSNAME <> "" then
				sql_where = sql_where & " and dutyman in ( select	u.id		from armyinformix.dbo.pbudae p, armyinformix.dbo.user1 u  where p.name like '"&CLASSNAME&"%' and p.auth = u.unit ) "			
		end if
		if whereCD6 <> "" then
				sql_where = sql_where & " and nameoffact like '%" & whereCD6 & "%'" 
		end if

	else
		if ToDate = "" then ToDate = date() end if
		SQL = "select * from armyinformix.dbo.receiptfact where sourcedate >= '"& FromDate & "' and sourcedate <= '"& ToDate & "'    order by sourcedate desc"
		sql_where = sql_where & " and inputdate >= '"& FromDate & "' and inputdate <= '"& ToDate & "'"
		sql_where = sql_where & " and dutyman is not null and dutyman <> ''"
		if whereCD3 <> "" then
			sql_where = sql_where & " and dutyman = '" & whereCD3 & "'" 
		elseif whereCD4 <> "" then
				sql_where = sql_where & " and dutyman in ( select	id		from armyinformix.dbo.user1 where [name] like '%"&whereCD4&"%' ) "
		end if
		if whereCD2 <> "" then
				sql_where = sql_where & " and receiptfactnum in ( select	factnum		from armyinformix.dbo.factpeople where [name] like '%"&whereCD2&"%' ) "
		end if
		if whereCD1 <> "" then
			sql_where = sql_where & " and receiptfactnum like '%" & whereCD1 & "%'" 
		end if
		if QUERYGB = "1" then
				sql_where = sql_where & " and receiptfactnum in ( select	factnum		from armyinformix.dbo.factpeople where left(ReserveDate,10) = convert(char(10),getdate(),121))"
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

	'response.write sql

	'4. Paging HTML 작성
	totalCount = db_getCount(db, sql_tb, sql_where)
	startRow = totalCount - pageSize * (curPage - 1)
	pageHtml = getPageHtml(pageSector, pageSize, totalCount, curPage, currentURL & "?" & where1)


%>
<div id='CalendarLayer' style='display:none;'><iframe name='CalendarFrame' src='/include/Calendar.asp' width='172' height='177' border='0' frameborder='0' scrolling='no'></iframe></div>
<table border="0" width="940" cellpadding="0" cellspacing="3" bgcolor="#EFECE5" align="center">
	<tr bgcolor="#FFFFFF">
		<td>
		
			<form method="post" name="inUpFrm" action="<%=Menu_2nd%>"  style="margin:0">
			<input type="hidden" name="QueryYN" value="">
			<table width="100%" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">
						<input type="hidden" name="curPage" value="<%=curPage%>">
			    <tr>
					<td bgcolor="#EFEFEF" class="TDCont" align='center'>구분</td>
					<td bgcolor="#FFFFFF" colspan=3>
						<input type="radio" name="QUERYGB" value="2" class="none" <% if QUERYGB = "2" then%>  checked<%end if%> onClick="fn_Search();">&nbsp;진행목록
						<input type="radio" name="QUERYGB" value="1" class="none" <% if QUERYGB = "1" then%>  checked<%end if%> onClick="fn_Search();">&nbsp;금일예약목록
					</td>
			        <td width="110" bgcolor="#EFEFEF" class="TDCont" align='center'>조회기간</td>
			        <td colspan="3" bgcolor="#FFFFFF" >
						<input value="<%=FromDate%>" name="FromDate" readonly type="text" size="10" onfocus="setFocusColor(this);" onClick="new CalendarFrame.Calendar(this);" onblur="setOutColor(this);">&nbsp;<img src="/Images/Btn/BtnIconDel.gif" title="삭제" style="cursor:hand;" align="absmiddle"onclick="document.all.FromDate.value='';">
				    	~
						<input value="<%=ToDate%>" name="ToDate" readonly type="text" size="10" onfocus="setFocusColor(this);" onClick="new CalendarFrame.Calendar(this);" onblur="setOutColor(this);">&nbsp;<img src="/Images/Btn/BtnIconDel.gif" title="삭제" style="cursor:hand;" align="absmiddle"onclick="document.all.ToDate.value='';">

			        </td>
			        <td colspan='2' rowspan="3" bgcolor="#FFFFFF" align="center">
			        	<img src="/Images/Btn/BtnSearch.gif" style="cursor:hand;" onClick="fn_Search();">
			        	<%IF SS_Login_Secgroup="A" Or SS_Login_Secgroup="B" THEN%><br><br><img src="/Images/Btn/BtnExcel.gif" style="cursor:hand;" onClick="fn_Xls();"><%END IF%>
			        </td>
				</tr>
			    <tr>

			        <td width="100" bgcolor="#EFEFEF" class="TDCont" align='center'>사건번호</td>
			        <td bgcolor="#FFFFFF">
			        	<input value="<%=whereCD1%>" name="whereCD1" type="text" size="16" onfocus="setFocusColor(this);" onblur="setOutColor(this);"  onKeypress="if (event.keyCode==13) {fn_Search();}"></td>

			        <td width="100" bgcolor="#EFEFEF" class="TDCont" align='center'>피의자명</td>
			        <td bgcolor="#FFFFFF">
			        	<input value="<%=whereCD2%>" name="whereCD2" type="text" size="16" onfocus="setFocusColor(this);" onblur="setOutColor(this);"  onKeypress="if (event.keyCode==13) {fn_Search();}"></td>

			        <td width="100" bgcolor="#EFEFEF" class="TDCont" align='center'>담당수사관</td>
			        <td bgcolor="#FFFFFF" colspan="3"><input value="<%=whereCD4%>" name="whereCD4" type="text" size="14" onfocus="setFocusColor(this);" onblur="setOutColor(this);"  onKeypress="if (event.keyCode==13) {fn_Search();}">&nbsp;


			        	<select name="whereCD3" size="1" class="ComboFFFCE7" onchange="pCateSelect('1');">
							<option value="">전체</option>
<%					
							SQL = "	select * from armyinformix.dbo.user1 order by name" '수사관정보
							SET Rs = DB.execute(SQL)
							do until Rs.eof
									CODE = Rs("id")
									if CODE = "" then
										CODENAME = "id없음("&Rs("name")&")"
									else
										CODENAME = Rs("name")
									end if
								%>

									<%=printSelect("" &CODENAME& "","" &CODE& "","" &whereCD3& "")%>
								<%
								Rs.movenext
							loop
%>
						</select>
					</td>
			    </tr>
				<tr>
			        <td width="100" bgcolor="#EFEFEF" class="TDCont" align='center'>사건명</td>
			        <td bgcolor="#FFFFFF"colspan= "3">
			        	<input value="<%=whereCD6%>" name="whereCD6" type="text" size="50" onfocus="setFocusColor(this);" onblur="setOutColor(this);" onKeypress="if (event.keyCode==13) {fn_Search();}"></td>

					</td>

			        <td width="100" bgcolor="#EFEFEF" class="TDCont"  align='center'>소속</td>
					<td bgcolor="#FFFFFF" nowrap colspan= "3">
					<input type="hidden" name="ACLASS" value="" maxlength="15" size="20" onfocus="setFocusColor(this);" onblur="setOutColor(this);" align="absmiddle"><input type="hidden" name="BCLASS" value="" maxlength="15" size="20" onfocus="setFocusColor(this);" onblur="setOutColor(this);" align="absmiddle"><input type="hidden" name="CCLASS" value="" maxlength="15" size="40" onfocus="setFocusColor(this);" onblur="setOutColor(this);" align="absmiddle"><input type="text" size="30" name="CLASSNAME" value="<%=CLASSNAME%>" onfocus="setFocusColor(this);" onblur="setOutColor(this);" align="absmiddle" onKeypress="if (event.keyCode==13) {fn_Search();}">
					</td>
				</tr>

				<!--<tr>
			        

				</tr>-->
			</table>
			</form>
		</td>
	</tr>
</table>

<table width="940" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>

<table width="940" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">
<% if QUERYGB = "1" then %>
	<tr height="20" bgcolor="#EEF6FF" align="center">
		<td rowspan=2 class="TDCont" >No</td>
		<td rowspan=2 class="TDCont" >사건번호</td>
		<td rowspan=2 class="TDCont" >예약정보</td>
		<td rowspan=2 class="TDCont" >사건명</td>
		<td rowspan=2 class="TDCont" >첨부물목록</td>
		<td rowspan=2 class="TDCont" >출처일자</td>
		<td colspan=3 class="TDCont" >담당수사관</td>
		<td colspan=3 class="TDCont" >연락처유무</td>
	</tr>
	<tr height="20" bgcolor="#EEF6FF" align="center">
		<td class="TDCont" >소속</td>
		<td class="TDCont" >계급</td>
		<td class="TDCont" >성명</td>
		<td >피의자</td>
		<td >피해자</td>
		<td >지휘관</td>
	</tr>
	<tr><td colspan="17" height="1" bgcolor="#FFFFFF"></td></tr>

<% else %>

	<tr height="20" bgcolor="#EEF6FF" align="center">
		<td rowspan=2 class="TDCont" >No</td>
		<td rowspan=2 class="TDCont" >사건번호</td>
		<td rowspan=2 class="TDCont" >사건명</td>
		<td rowspan=2 class="TDCont" >첨부물목록</td>
		<td rowspan=2 class="TDCont" >출처일자</td>
		<td colspan=3 class="TDCont" >담당수사관</td>
		<td colspan=3 class="TDCont" >연락처유무</td>
		<td rowspan=2 class="TDCont" >평가점수</td>
		<td rowspan=2 class="TDCont" align='center' width='30'><input type="checkbox" name="chkALL" value="" class="none" onclick="fn_select();"></td>
		<!--<td rowspan=2 class="TDCont" >관리</td>-->
	</tr>

	<tr height="20" bgcolor="#EEF6FF" align="center">
		<td class="TDCont" >소속</td>
		<td class="TDCont" >계급</td>
		<td class="TDCont" >성명</td>
		<td >피의자</td>
		<td >피해자</td>
		<td >지휘관</td>
	</tr>

	<tr><td colspan="16" height="1" bgcolor="#FFFFFF"></td></tr>

<% end if %>

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


		if mid(RsGBN("receiptfactnum"),6,2) >= "10" then		

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

			'SQL = " select filenum from monitorfile where receiptfactnum = '" & RsGBN("receiptfactnum") & "' order by filenum"
			'SET Rs1 = informixDB.execute(SQL)
			'sFileNum = ""
			'do until Rs1.eof
			'	if sFileNum = "" then
			'		sFileNum = Rs1("filenum")
			'	else
			'		sFileNum = sFileNum & "," & Rs1("filenum")
			'	end if

			'	Rs1.movenext
			'loop
			'Rs1.close
			i = i + 1
			'관련파일명
		SQL = "	select top 3 * from armyinformix.dbo.monitorfile where receiptfactnum = '" & RsGBN("receiptfactnum") & "' order by filenum"
			'if sFileNum <> "" then
				'SQL = "	select top 3 * from armyinformix.dbo.monitorfile where filenum in (" & sFileNum & ") order by filenum"


				SET Rs = DB.execute(SQL)
				sfilename = ""
				do until Rs.eof
					if sfilename = "" then
						sfilename = "<a href='"&sFileLinkURL&"filename="& rs("filename")&"'>" & CutString(rs("filename"), 10, "...") & "</a>"
					else
						sfilename = sfilename & "<br><a href='"&sFileLinkURL&"filename="& rs("filename")&"'>" & CutString(rs("filename"), 10, "...") & "</a>"
					end if
					Rs.movenext
				loop
			'end if
		end if
		'수사관정보
		'SQL = "	select name, class, (select contents from data where restrict = 104 and number = class ) as classname from user1 where id = '" & RsGBN("dutyman") & "'"

		SQL = "	select name, class, (select name from armyinformix.dbo.pbudae where auth = unit) as unitname from armyinformix.dbo.user1 where id = '" & RsGBN("dutyman") & "'"
		SET Rs = DB.execute(SQL)

		if Rs.eof = false and RsGBN("dutyman") <> "" then
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

		'sql = "	select count(*) from armyinformix.dbo.factpeople where factnum = '" & RsGBN("receiptfactnum") & "' and section2 in ('B11','413')"
		'SET Rs = DB.execute(SQL)
		'if rs(0)>0 then
		'	sYN1 = "Y"
		'end if
		'sql = "	select count(*) from armyinformix.dbo.factpeople where factnum = '" & RsGBN("receiptfactnum") & "' and section2 in ('B12','448')"
		'SET Rs = DB.execute(SQL)
		'if rs(0)>0 then
		'	sYN2 = "Y"
		'end if
		'sql = "	select count(*) from armyinformix.dbo.factpeople where factnum = '" & RsGBN("receiptfactnum") & "' and section2 in ('B15','447','449')"
		'SET Rs = DB.execute(SQL)
		'if rs(0)>0 then
		'	sYN3 = "Y"
		'end if


		sql = "	select [name], monitorresult, isnull(monitorpoint,0) as monitorpoint from armyinformix.dbo.factpeople where factnum = '" & RsGBN("receiptfactnum") & "' and section2 in ('B11','413') "'and monitorresult = '9'"
		SET Rs = DB.execute(SQL)
		do until rs.eof
			if sYN1 = "" then
				if rs("monitorresult") = "9" then
					sYN1 = rs(0)&"<br>("&formatnumber(rs("monitorpoint"))&")"
				elseif isnull(rs("monitorresult")) then
					sYN1 = rs(0)&"<br>(<font color='#ff0000'>미진행</font>)"
				else
					sYN1 = rs(0)&"<br>(<font color='#ff0000'>"&db_getCodeName("B10",rs("monitorresult"))&"</font>)"
				end if
			else
				if rs("monitorresult") = "9" then
					sYN1 = sYN1&"<br>"&rs(0)&"<br>("&formatnumber(rs("monitorpoint"))&")"
				elseif isnull(rs("monitorresult")) then
					sYN1 = sYN1&"<br>"&rs(0)&"<br>(<font color='#ff0000'>미진행</font>)"
				else
					sYN1 = sYN1&"<br>"&rs(0)&"<br>(<font color='#ff0000'>"&db_getCodeName("B10",rs("monitorresult"))&"</font>)"
				end if
			end if
			rs.movenext
		loop
		sql = "	select [name], monitorresult, isnull(monitorpoint,0) as monitorpoint  from armyinformix.dbo.factpeople where factnum = '" & RsGBN("receiptfactnum") & "' and section2 in ('B12','448')"' and monitorresult = '9'"
		SET Rs = DB.execute(SQL)
		do until rs.eof
			if sYN2 = "" then
				if rs("monitorresult") = "9" then
					sYN2 = rs(0)&"<br>("&formatnumber(rs("monitorpoint"))&")"
				elseif isnull(rs("monitorresult")) then
					sYN2 = rs(0)&"<br>(<font color='#ff0000'>미진행</font>)"
				else
					sYN2 = rs(0)&"<br>(<font color='#ff0000'>"&db_getCodeName("B10",rs("monitorresult"))&"</font>)"
				end if
			else
				if rs("monitorresult") = "9" then
					sYN2 = sYN2&"<br>"&rs(0)&"<br>("&formatnumber(rs("monitorpoint"))&")"
				elseif isnull(rs("monitorresult")) then
					sYN2 = sYN2&"<br>"&rs(0)&"<br>(<font color='#ff0000'>미진행</font>)"
				else
					sYN2 = sYN2&"<br>"&rs(0)&"<br>(<font color='#ff0000'>"&db_getCodeName("B10",rs("monitorresult"))&"</font>)"
				end if
			end if
			rs.movenext
		loop
		sql = "	select [name], monitorresult, isnull(monitorpoint,0) as monitorpoint  from armyinformix.dbo.factpeople where factnum = '" & RsGBN("receiptfactnum") & "' and section2 in ('B15','447','449')"' and monitorresult = '9'"
		SET Rs = DB.execute(SQL)
		do until rs.eof
			if sYN3 = "" then
				if rs("monitorresult") = "9" then
					sYN3 = rs(0)&"<br>("&formatnumber(rs("monitorpoint"))&")"
				elseif isnull(rs("monitorresult")) then
					sYN3 = sYN2&"<br>"&rs(0)&"<br>(<font color='#ff0000'>미진행</font>)"
				else
					sYN3 = rs(0)&"<br>(<font color='#ff0000'>"&db_getCodeName("B10",rs("monitorresult"))&"</font>)"
				end if
			else
				if rs("monitorresult") = "9" then
					sYN3 = sYN3&"<br>"&rs(0)&"<br>("&formatnumber(rs("monitorpoint"))&")"
				elseif isnull(rs("monitorresult")) then
					sYN3 = sYN3&rs(0)&"<br>(<font color='#ff0000'>미진행</font>)"
				else
					sYN3 = sYN3&"<br>"&rs(0)&"<br>(<font color='#ff0000'>"&db_getCodeName("B10",rs("monitorresult"))&"</font>)"
				end if
			end if
			rs.movenext
		loop


		sql = "	select '피의자: '+ReserveDate from armyinformix.dbo.factpeople where factnum = '" & RsGBN("receiptfactnum") & "' and left(ReserveDate,10) = convert(char(10),getdate(),121) and section2 in ('B11','413')"
		SET Rs = DB.execute(SQL)
		
		if Rs.eof = false then
			if sResvervData <> "" then
				sResvervData = sResvervData & "<br>" & Rs(0)''지휘관: '+ReserveDate
			else
				sResvervData = Rs(0)''지휘관: '+ReserveDate
			end if
		end if

		sql = "	select '피해자: '+ReserveDate from armyinformix.dbo.factpeople where factnum = '" & RsGBN("receiptfactnum") & "' and left(ReserveDate,10) = convert(char(10),getdate(),121) and section2 in ('B12','448')"
		SET Rs = DB.execute(SQL)
		
		if Rs.eof = false then
			if sResvervData <> "" then
				sResvervData = sResvervData & "<br>" & Rs(0)''지휘관: '+ReserveDate
			else
				sResvervData = Rs(0)''지휘관: '+ReserveDate
			end if
		end if

		sql = "	select '지휘관: '+ReserveDate from armyinformix.dbo.factpeople where factnum = '" & RsGBN("receiptfactnum") & "' and left(ReserveDate,10) = convert(char(10),getdate(),121) and section2 in ('B15','447','449')"
		SET Rs = DB.execute(SQL)
		
		if Rs.eof = false then
			if sResvervData <> "" then
				sResvervData = sResvervData & "<br>" & Rs(0)''지휘관: '+ReserveDate
			else
				sResvervData = Rs(0)''지휘관: '+ReserveDate
			end if
		end if
		
		set Rs = nothing


		receiptfactnum = RsGBN("receiptfactnum") 
		contents = RsGBN("contents") 

%>

	<tr id="cTR1" style="cursor:hand;" bgcolor="#ffffff" onmouseover="this.style.background='#FFFCE7'" bgcolor="#FFFFFF" onmouseout="this.style.background='#FFFFFF'">
			<td align="center" width=40 class="TDCont" ><%=startRow%></td>
			<td align="center" class="TDCont" nowrap title="<%=contents%>"><a href="javascript:fn_update('<%=receiptfactnum%>','UP');"><%=RsGBN("receiptfactnum")%></a></td>
<% if QUERYGB = "1" then %>
			<td align="center" class="TDCont" ><a href="javascript:fn_update('<%=receiptfactnum%>','UP');"><%=sResvervData%></a></td>
<% end if %>

			<td align="left" class="TDCont" title="<%=RsGBN("nameoffact")%>" nowrap>&nbsp;<a href="javascript:fn_update('<%=receiptfactnum%>','UP');"><%=CutString(RsGBN("nameoffact"), 10, "...")%></a></td>

			<td align="left" class="TDCont" ><%=sfilename%></td>
			<td align="center" class="TDCont" ><%=RsGBN("inputdate")%></td>
			<td align="center" class="TDCont" ><%=sBudae%></td>
			<td align="center" class="TDCont" ><%=sClassName%></td>
			<td align="center" class="TDCont" ><%=sName%></td>
			<td align="center"><%=sYN1%></td>
			<td align="center"><%=sYN2%></td>
			<td align="center"><%=sYN3%></td>

			<td align="center" class="TDCont" ><%=RsGBN("monitorpoint")%></td>
<% if QUERYGB <> "1" then %>
			<td align="center"><input type="checkbox" name="chk" value="<%=receiptfactnum%>" class="none"></td>
<% end if %>
			<!--<td align="center">
				<img src="/Images/Btn/BtnIconModify.gif" title="수정" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_update('<%=RsGBN("receiptfactnum")%>','UP');">
				<img src="/Images/Btn/BtnIconDel.gif" title="삭제" style="cursor:hand;" align="absmiddle"onclick="javascript:fn_del('<%=RsGBN("receiptfactnum")%>','DEL');">
			</td>-->
	</tr>

<%
		startRow = startRow - 1

		RsGBN.MoveNext
	LOOP
	RsGBN.CLOSE
	SET RsGBN = Nothing
%>

</table>
<% if QUERYGB <> "1" then %>
<table border="0" cellpadding="0" cellspacing="0" width="940" align="center">
	<tr><td height="5" colspan='2'></td></tr>
	<tr><td height="1" bgcolor="#D6D6D6" colspan='2'></td></tr>
	<tr height="22" bgcolor="#EEF6FF"><td align="center"><%=pageHtml%></td><td align='right'><img src="/Images/Btn/BtnTabMove1.GIF" style="cursor:hand;" align="absmiddle" onClick="fn_move();"></td></tr>
	<tr><td height="1" bgcolor="#D6D6D6" colspan='2'></td></tr>
</table>
<% else %>
<table border="0" cellpadding="0" cellspacing="0" width="940" align="center">
	<tr><td height="5" colspan='2'></td></tr>
	<tr><td height="1" bgcolor="#D6D6D6" colspan='2'></td></tr>
	<tr height="22" bgcolor="#EEF6FF"><td align="center"><%=pageHtml%></td><td align='right'></td></tr>
	<tr><td height="1" bgcolor="#D6D6D6" colspan='2'></td></tr>
</table>
<% end if %>

<iframe src="about:blank" name="DBFrame" width="0" height="0" frameborder=0 marginheight=0 marginwidth=0 scrolling="no"></iframe>
<script>

	function fn_select() {
		var i, chked=0;

		for(i=0;i<document.all.chk.length;i++) {
				document.all.chk[i].checked=document.all.chkALL.checked;
		}

		if ( document.all.chk[0] == null)
		{
			document.all.chk.checked=document.all.chkALL.checked;
		}
	}

	function fn_move() {
		var i, chked=0;
		var j = 0;
		var codestring;
		for(i=0;i<document.all.chk.length;i++) {
				if ( document.all.chk[i].checked )
				{
					j = j + 1;
					//alert(document.all.chk[i].value);
					if ( j == 1 )
						codestring = "1,"+document.all.chk[i].value;
					else
						codestring = codestring+ ","+document.all.chk[i].value;

				}
		}


		if ( document.all.chk[0] == null && document.all.chk.checked  )
		{
			j = j + 1;
			codestring = "1,"+document.all.chk.value;
		}

		if ( j > 0 )
		{
			DBFrame.location= "/menu01/submenu0101/move.asp?FRM=submenu02&factnum="+codestring;
			alert('['+j+'건]을 대기목록으로 이동 완료하였습니다!');
			document.inUpFrm.submit();
		}

	}
	function fn_Search() {
		document.inUpFrm.curPage.value = "";
		document.inUpFrm.submit();
	}

	function fn_update(arg0,arg1) {	
		location.href="/menu01/submenu0101/monitoring.asp?FRM=submenu02&factnum="+arg0+"&<%=where2%>";
	}

	function fn_Xls() {
		location.href="research02_Xls.asp?<%=where2%>"
	}

	function pCateSelect(cn){

		if ( inUpFrm.whereCD3.value == '' )
			document.inUpFrm.whereCD4.value ='';
		else
			document.inUpFrm.whereCD4.value = document.inUpFrm.whereCD3.options[inUpFrm.whereCD3.selectedIndex].text;
	}
</script>
<%
	'informixDB.close
	Set informixDB=nothing
%>
<!-- #include virtual="/Include/Bottom.asp" -->