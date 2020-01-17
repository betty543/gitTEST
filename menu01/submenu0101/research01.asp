<!-- #include virtual="/Include/Top.asp" -->
<!-- #include virtual="/Include/PopLayer.asp" -->
<!-- #include virtual="/Include/DBConnection_info.asp" -->
<%
'On Error Resume next

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
	whereCD4 = ""

	CLASSNAME = Trim(request("CLASSNAME"))

	if FromDate = "" then
		FromDate = left(date(),4) & "-01-01"
	end if

	'2. 쿼리조건절 셋팅
	pageSize = 10
	pageSector = 10
	if curPage = "" then curPage = 1 end If

	where1 = "FromDate=" & FromDate & "&ToDate=" & ToDate & "&whereCD1=" & whereCD1 & "&whereCD2=" & whereCD2 & "&whereCD3=" & whereCD3 & "&whereCD4=" & whereCD4 &"&QUERYGB="&QUERYGB&"&CLASSNAME="&CLASSNAME&"&whereCD6="&whereCD6&"&whereCD5="&whereCD5&"&whereCD7="&whereCD7
	where2 = "curPage=" & curPage & "&" & where1
	''sql_where = " 1=1 and handlesection not in ('01','100','0000000000')" '완료건
	sql_where = " 1=1 "
	
	if whereCD7 <> "Y" then
		sql_where = sql_where & " and filecnt > 0 and ( processgb is null or processgb ='') " '완료건
	else
		sql_where = sql_where & " and ( processgb is null or processgb ='') " '대기리스트
	end if
	if FromDate = "" then		
		SQL = "select * from armyinformix.dbo.receiptfact where inputdate >= '"& left(date(),7) & "-01' order by inputdate desc"
		sql_where = sql_where & " and inputdate >= '"& left(date(),7) & "-01'"
		'sql_where = sql_where & " and dutyman is not null and dutyman <> ''"
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
		'sql_where = sql_where & " and dutyman is not null and dutyman <> ''"
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
	sql_orderby = "modifydate desc, inputdate DESC, receiptfactnum desc"

	'3. 쿼리 실행
	sql = db_getSqlWithPage(sql_tb, sql_index, sql_field, sql_where, sql_orderby, pageSize, curPage)
	set RsGBN = db.execute(sql)

	'Response.Write sql

	'4. Paging HTML 작성
	totalCount = db_getCount(db, sql_tb, sql_where)
	startRow = totalCount - pageSize * (curPage - 1)
	pageHtml = getPageHtml(pageSector, pageSize, totalCount, curPage, currentURL & "?" & where1)

%>
<div id='CalendarLayer' style='display:none;'><iframe name='CalendarFrame' src='/include/Calendar.asp' width='172' height='177' border='0' frameborder='0' scrolling='no'></iframe></div>

<table border="0" width="940" cellpadding="0" cellspacing="3" bgcolor="#EFECE5" align="center">
	<tr bgcolor="#FFFFFF">
		<td>

			<form method="post" name="inUpFrm" action="<%=Menu_2nd%>" onsubmit="return fn_Search(this);"  style="margin:0">
			<table width="100%" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">
			    <tr>
			        <td width="80" bgcolor="#EFEFEF" class="TDCont" align='center'>조회기간</td>
			        <td colspan="3" bgcolor="#FFFFFF" >
						<input type="hidden" name="curPage" value="<%=curPage%>">
			        	<input value="<%=FromDate%>" name="FromDate" readonly type="text" size="10" onfocus="setFocusColor(this);" onClick="new CalendarFrame.Calendar(this);" onblur="setOutColor(this);">&nbsp;<img src="/Images/Btn/BtnIconDel.gif" title="삭제" style="cursor:hand;" align="absmiddle"onclick="document.all.FromDate.value='';">
				    	~
						<input value="<%=ToDate%>" name="ToDate" readonly type="text" size="10" onfocus="setFocusColor(this);" onClick="new CalendarFrame.Calendar(this);" onblur="setOutColor(this);">&nbsp;<img src="/Images/Btn/BtnIconDel.gif" title="삭제" style="cursor:hand;" align="absmiddle"onclick="document.all.ToDate.value='';">

			        </td>

			        <td width="80" bgcolor="#EFEFEF" class="TDCont"  align='center'>소속</td>
					<td bgcolor="#FFFFFF" nowrap>
					<input type="hidden" name="ACLASS" value="" maxlength="15" size="20" onfocus="setFocusColor(this);" onblur="setOutColor(this);" align="absmiddle"><input type="hidden" name="BCLASS" value="" maxlength="15" size="20" onfocus="setFocusColor(this);" onblur="setOutColor(this);" align="absmiddle"><input type="hidden" name="CCLASS" value="" maxlength="15" size="40" onfocus="setFocusColor(this);" onblur="setOutColor(this);" align="absmiddle"><input type="text" size="30" name="CLASSNAME" value="<%=CLASSNAME%>" onfocus="setFocusColor(this);" onblur="setOutColor(this);" align="absmiddle" onKeypress="if (event.keyCode==13) {fn_Search();}"><!--&nbsp;<img src="/Images/Btn/BtnIconDel.gif" title="삭제" style="cursor:hand;" align="absmiddle"onclick="document.all.ACLASS.value='';document.all.BCLASS.value='';document.all.CCLASS.value='';document.all.CLASSNAME.value='';">&nbsp;<img src="/Images/Comm/IconTip.gif" style="cursor:hand;" align="absmiddle" onClick="pCateSelect('1');" >-->
					</td>

			        <td colspan='2' rowspan="3" bgcolor="#FFFFFF" align="center">
			        	<img src="/Images/Btn/BtnSearch.gif" style="cursor:hand;" onClick="fn_Search();">
			        	<br><br><img src="/Images/Btn/BtnExcel.gif" style="cursor:hand;" onClick="fn_Xls();">
			        </td>
			    </tr>
				<tr>
			        <td bgcolor="#EFEFEF" class="TDCont" align='center'>사건번호</td>
			        <td bgcolor="#FFFFFF">
			        	<input value="<%=whereCD2%>" name="whereCD2" type="text" size="20" onfocus="setFocusColor(this);" onblur="setOutColor(this);" onKeypress="if (event.keyCode==13) {fn_Search();}"></td>
			        <td width="80" bgcolor="#EFEFEF" class="TDCont" align='center'>피의자명</td>
			        <td bgcolor="#FFFFFF">
			        	<input value="<%=whereCD3%>" name="whereCD3" type="text" size="20" onfocus="setFocusColor(this);" onblur="setOutColor(this);" onKeypress="if (event.keyCode==13) {fn_Search();}"></td>

			        <td width="80" bgcolor="#EFEFEF" class="TDCont" align='center'>담당수사관</td>
			        <td bgcolor="#FFFFFF"><input value="<%=whereCD5%>" name="whereCD5" type="text" size="20" onfocus="setFocusColor(this);" onblur="setOutColor(this);" onKeypress="if (event.keyCode==13) {fn_Search();}">&nbsp;

			        	<select name="whereCD4" size="1" class="ComboFFFCE7" onchange="pCateSelect('1');">
						<option value="">선택</option>
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

									<%=printSelect("" &CODENAME& "","" &CODE& "","" &whereCD4& "")%>
								<%
								Rs.movenext
							loop
%>
						</select>
					</td>

				</tr>
				<tr>
			        <td bgcolor="#EFEFEF" class="TDCont" align='center'>사건명</td>
			        <td bgcolor="#FFFFFF"colspan= "3">
			        	<input value="<%=whereCD6%>" name="whereCD6" type="text" size="60" onfocus="setFocusColor(this);" onblur="setOutColor(this);" onKeypress="if (event.keyCode==13) {fn_Search();}"></td>

					</td>
					<td bgcolor="#EFEFEF" class="TDCont" align='center'></td>
					<td bgcolor="#FFFFFF" align="right">
						파일 미첨부 자료 포함<input type="checkbox" name="whereCD7" class="None" value="Y" onClick="fn_Search();" title="파일미첨부자료포함" <% if whereCD7 = "Y" then%>checked<%end if%>>&nbsp;
					</td>
				</tr>
			</table>
			</form>
		</td>
	</tr>
</table>

<table width="940" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>

<table width="940" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">
	<tr height="20" bgcolor="#EEF6FF" align="center">
		<td rowspan=2 class="TDCont" align='center'>No</td>
		<td rowspan=2 class="TDCont" align='center'>사건번호</td>
		<td rowspan=2 class="TDCont" align='center'>사건명</td>
		<td rowspan=2 class="TDCont" align='center'>첨부파일목록</td>
		<td rowspan=2 class="TDCont" align='center'>출처일자</td>
		<td colspan=3 class="TDCont" align='center'>담당수사관</td>
		<td colspan=3 class="TDCont" align='center'>연락처유무</td>
		<td rowspan=2 class="TDCont" align='center' width='30'><input type="checkbox" name="chkALL" value="" class="none" onclick="fn_select();"></td>
	</tr>
	<tr height="20" bgcolor="#EEF6FF" align="center">
		<td class="TDCont">소속</td>
		<td class="TDCont">계급</td>
		<td class="TDCont">성명</td>
		<td >피의자</td>
		<td >피해자</td>
		<td >관계자</td>
	</tr>
	<tr><td colspan="17" height="1" bgcolor="#FFFFFF"></td></tr>

<%

	i = 0


	do until RsGBN.eof

'---------------------------------------------------------------------------------------------------------------
'informixDB
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
			'SET Rs1 = DB.execute(SQL)
			'sFileNum = ""
			'sfilename = ""
			'do until Rs1.eof
		'		if sFileNum = "" then
		'			sFileNum = Rs1("filenum")
		'		else
		'			sFileNum = sFileNum & "," & Rs1("filenum")
		'		end if

		'		Rs1.movenext
		'	loop
		'	Rs1.close
			i = i + 1
			'관련파일명
			SQL = "	select top 3 * from armyinformix.dbo.monitorfile where receiptfactnum = '" & RsGBN("receiptfactnum") & "' order by filenum"
			'if sFileNum <> "" then
			'	SQL = "	select top 3 * from armyinformix.dbo.monitorfile where filenum in (" & sFileNum & ") order by filenum"
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
			<td align="center" class="TDCont" title="<%=RsGBN("modifydate")%>" nowrap><%=RsGBN("inputdate")%></td>
			<td align="center" class="TDCont" nowrap><%=sBudae%></td>
			<td align="center" class="TDCont" nowrap><%=sClassName%></td>
			<td align="center" class="TDCont" nowrap><%=sName%></td>
			<td align="center" class="TDCont" nowrap><% if sYN1 = "Y" then%><img src="/Images/Btn/icon_02.gif" style="cursor:hand;" align="absmiddle"><%end if%></td>
			<td align="center" ><% if sYN2 = "Y" then%><img src="/Images/Btn/icon_02.gif" style="cursor:hand;" align="absmiddle"><%end if%></td>
			<td align="center"><% if sYN3 = "Y" then%><img src="/Images/Btn/icon_02.gif" style="cursor:hand;" align="absmiddle"><%end if%></td>
			<td align="center"><input type="checkbox" name="chk" value="<%=receiptfactnum%>" class="none"></td>

	</tr>

<%
		startRow = startRow - 1

		RsGBN.MoveNext
	LOOP
	RsGBN.CLOSE
	SET RsGBN = Nothing

%>
</table>

<table border="0" cellpadding="0" cellspacing="0" width="940" align="center">
	<tr><td height="5" colspan='2'></td></tr>
	<tr><td height="1" bgcolor="#D6D6D6" colspan='2'></td></tr>
	<tr height="22" bgcolor="#EEF6FF"><td align="center"><%=pageHtml%></td><td align='right'><img src="/Images/Btn/BtnTabMove.GIF" style="cursor:hand;" align="absmiddle" onClick="fn_move();"></td></tr>
	<tr><td height="1" bgcolor="#D6D6D6" colspan='2'></td></tr>
</table>

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
			DBFrame.location= "/menu01/submenu0101/move.asp?FRM=submenu01&factnum="+codestring;
			alert('['+j+'건]을 진행목록으로 이동 완료하였습니다!');
			document.inUpFrm.submit();
		}
	}

	function fn_Search() {
		document.inUpFrm.curPage.value = "";
		document.inUpFrm.submit();
	}

	function fn_new(arg0,arg1) {	
		location.href="/menu01/submenu0101/monitoring.asp?FRM=submenu01&factnum="+arg0+"&<%=where2%>";
	}

	function fn_Xls() {
		location.href="research01_Xls.asp?<%=where2%>"
	}

	function pCateSelect(cn){

		if ( inUpFrm.whereCD4.value == '' )
			document.inUpFrm.whereCD5.value ='';
		else
			document.inUpFrm.whereCD5.value = document.inUpFrm.whereCD4.options[inUpFrm.whereCD4.selectedIndex].text;
	}
</script>

<%
	' informixDB.close
	Set informixDB=nothing
%>
<!-- #include virtual="/Include/Bottom.asp" -->