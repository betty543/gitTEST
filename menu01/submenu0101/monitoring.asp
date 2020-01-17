
<!-- #include virtual="/Include/Top.asp" -->
<!-- #include virtual="/Include/PopLayer.asp" -->
<!-- #include virtual="/Include/DBConnection_info.asp" -->

<%
	'end if

On Error Resume next
	FRM = Trim(request("FRM"))
	factnum = request("factnum")

	'파일경로 찾아오기
	SQL = "select top 1 * from TB_Code WHERE	Codegroup = 'B20' AND UseYN = 'Y'"
	set RS = db.Execute(SQL)
	if rs.eof = false then
		sFileLinkURL = rs("Codename")
	else
		sFileLinkURL = "http://16.1.19.160:7001/amcriss/M_investigation/down.jsp?"
	end if

	if FRM = "submenu01" then	'신규


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
		CLASSNAME = Trim(request("CLASSNAME")) 


		'2. 쿼리조건절 셋팅
		pageSize = 10
		pageSector = 10
		if curPage = "" then curPage = 1 end If

		where1 = "FromDate=" & FromDate & "&ToDate=" & ToDate & "&whereCD1=" & whereCD1 & "&whereCD2=" & whereCD2 & "&whereCD3=" & whereCD3 & "&whereCD4=" & whereCD4 &"&QUERYGB="&QUERYGB&"&CLASSNAME="&CLASSNAME&"&whereCD6="&whereCD6&"&whereCD5="&whereCD5&"&whereCD7="&whereCD7
		where2 = "curPage=" & curPage & "&" & where1

		SQL = "select * from armyinformix.dbo.receiptfact where receiptfactnum = '" & factnum &"'"

		SET RsGBN = DB.execute(SQL)
		if RsGBN.eof = false then

			dutyman = RsGBN("dutyman")
			receiptfactnum = RsGBN("receiptfactnum")
			contents = RsGBN("contents")
			nameoffact = RsGBN("nameoffact")
			occurplace = RsGBN("occurplace")

			inputdate = trim(RsGBN("inputdate"))
			Date2 = trim(RsGBN("Date1"))
			Date3 = trim(RsGBN("Date2"))
			receiptkind = RsGBN("receiptkind")

			processgb = RsGBN("processgb")	'설문조사
			processdate = RsGBN("processdate") '설문일시
			m_monitorpoint = RsGBN("monitorpoint")

			if m_monitorpoint <> "" then
				if m_monitorpoint >= 9.0 then
					m_monitorpoint = formatnumber(m_monitorpoint,2) & " (만족)"
				elseif m_monitorpoint >= 8.0 then
					m_monitorpoint = formatnumber(m_monitorpoint,2) & " (보통)"
				else
					m_monitorpoint = formatnumber(m_monitorpoint,2) & " (<font color='#ff0000'>불만족</font>)"
				end if
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

			if mid(RsGBN("receiptfactnum"),6,2) >= "10" then		

				'if sFileNum <> "" then
					SQL = "	select *  from armyinformix.dbo.monitorfile where receiptfactnum in ('" & RsGBN("receiptfactnum") & "') order by filenum"
					SET Rs = DB.execute(SQL)
					sfilename = ""
					do until Rs.eof
						if sfilename = "" then
							sfilename = "<a href='http://16.1.150.146:9080/vivid/JspSource/file/fileDownload.jsp?fileNumber="& rs("filenum")&"'>" & rs("filename") & "</a>"
						else
							sfilename = sfilename & "<br><a href='http://16.1.150.146:9080/vivid/JspSource/file/fileDownload.jsp?fileNumber="& rs("filenum")&"'>" & rs("filename") & "</a>"
						end if

						Rs.movenext
					loop
				'end if

			else

				'SQL = " select filenum from monitorfile where receiptfactnum = '" & RsGBN("receiptfactnum") & "' order by filenum"
				'SET Rs1 = DB.execute(SQL)
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
				'i = i + 1
				'관련파일명
				SQL = "	select top 3 * from armyinformix.dbo.monitorfile where receiptfactnum = '" & RsGBN("receiptfactnum") & "' order by filenum"
				'if sFileNum <> "" then
					'SQL = "	select * from armyinformix.dbo.monitorfile where filenum in (" & sFileNum & ") order by filenum"

					'SQL = "	select * from armyinformix.dbo.monitorfile where receiptfactnum = '" & RsGBN("receiptfactnum") & "' order by filenum"
					SET Rs = DB.execute(SQL)
					sfilename = ""
					do until Rs.eof
						if sfilename = "" then
							sfilename = "<a href='"&sFileLinkURL&"filename="& rs("filename")&"'>" & rs("filename") & "</a>"
						else
							sfilename = sfilename & "<br><a href='"&sFileLinkURL&"filename="& rs("filename")&"'>" & rs("filename") & "</a>"
						end if

						Rs.movenext
					loop
				'end if

			end if
		end if
	
	else 
	
		if FRM = "submenu02" then

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

		else

			curPage = request("curPage")
			FromDate = request("FromDate")
			ToDate = request("ToDate")
			whereCD1 = Trim(request("whereCD1"))
			whereCD2 = Trim(request("whereCD2"))
			whereCD3 = Trim(request("whereCD3"))
			whereCD4 = Trim(request("whereCD4"))
			whereCD5 = Trim(request("whereCD5"))
			whereCD6 = Trim(request("whereCD6"))

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

			where1 = "FromDate=" & FromDate & "&ToDate=" & ToDate & "&whereCD1=" & whereCD1 & "&whereCD2=" & whereCD2 & "&whereCD3=" & whereCD3 & "&whereCD4=" & whereCD4 &"&QUERYGB="&QUERYGB&"&CLASSNAME="&CLASSNAME&"&whereCD6="&whereCD6&"&whereCD5="&whereCD5
			where2 = "curPage=" & curPage & "&" & where1

		end if
		'기존 DB에서 가져오기

		SQL = "select * from armyinformix.dbo.receiptfact where receiptfactnum = '" & factnum &"'"

		SET RsGBN = DB.execute(SQL)
		if RsGBN.eof = false then

			dutyman = RsGBN("dutyman")
			receiptfactnum = RsGBN("receiptfactnum")
			contents = RsGBN("contents")
			nameoffact = RsGBN("nameoffact")
			occurplace = RsGBN("occurplace")
			Date2 = trim(RsGBN("Date1"))
			Date3 = trim(RsGBN("Date2"))
			receiptkind = RsGBN("receiptkind")
			inputdate = trim(RsGBN("inputdate"))

			processgb = RsGBN("processgb")	'설문조사
			processdate = RsGBN("processdate") '설문일시
			m_monitorpoint = RsGBN("monitorpoint")

			if m_monitorpoint <> "" then
				if m_monitorpoint >= 9.0 then
					m_monitorpoint = formatnumber(m_monitorpoint,2) & " (만족)"
				elseif m_monitorpoint >= 8.0 then
					m_monitorpoint = formatnumber(m_monitorpoint,2) & " (보통)"
				else
					m_monitorpoint = formatnumber(m_monitorpoint,2) & " (<font color='#ff0000'>불만족</font>)"
				end if
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




			if mid(RsGBN("receiptfactnum"),6,2) >= "10" then		

				'if sFileNum <> "" then
					SQL = "	select *  from armyinformix.dbo.monitorfile where receiptfactnum in ('" & RsGBN("receiptfactnum") & "') order by filenum"
					SET Rs = DB.execute(SQL)
					sfilename = ""
					do until Rs.eof
						if sfilename = "" then
							sfilename = "<a href='http://16.1.150.146:9080/vivid/JspSource/file/fileDownload.jsp?fileNumber="& rs("filenum")&"'>" & rs("filename") & "</a>"
						else
							sfilename = sfilename & "<br><a href='http://16.1.150.146:9080/vivid/JspSource/file/fileDownload.jsp?fileNumber="& rs("filenum")&"'>" & rs("filename") & "</a>"
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
				'i = i + 1
				'관련파일명
				SQL = "	select top 3 * from armyinformix.dbo.monitorfile where receiptfactnum = '" & RsGBN("receiptfactnum") & "' order by filenum"
				'if sFileNum <> "" then
					'SQL = "	select * from armyinformix.dbo.monitorfile where filenum in (" & sFileNum & ") order by filenum"

					'SQL = "	select * from armyinformix.dbo.monitorfile where receiptfactnum = '" & RsGBN("receiptfactnum") & "' order by filenum"
					SET Rs = DB.execute(SQL)
					sfilename = ""
					do until Rs.eof
						if sfilename = "" then
							sfilename = "<a href='"&sFileLinkURL&"filename="& rs("filename")&"'>" & rs("filename") & "</a>"
						else
							sfilename = sfilename & "<br><a href='"&sFileLinkURL&"filename="& rs("filename")&"'>" & rs("filename") & "</a>"
						end if

						Rs.movenext
					loop
				'end if
			end if

		end if

	end if




%>
<form method="post" name="inUpFrm" style="margin:0">
<table border="0" width="940" cellpadding="0" cellspacing="1" bgcolor="#EFECE5" align="center">
	<tr bgcolor="#FFFFFF">
		<td>		

			<input type="hidden" name="FRM" value="<%=FRM%>">	
			<table width="100%" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table width="100%" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#ffffff">
			    <tr>
					<td align="left" bgcolor="#FFFFFF" class="TDCont"><b><font color="#0000ff">&nbsp;<img src="/Images/dot_01.gif">&nbsp;담당수사관정보</font></b></td>
				</tr>
			</table>
			<table width="100%" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">
			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>소 속</td>
					<td bgcolor="#FFFFFF" width=150>&nbsp;<%=sBudae%>
					</td>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>계 급</td>
					<td bgcolor="#FFFFFF" width=150>&nbsp;<%=sClassName%>
					</td>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>수사관코드</td>
					<td bgcolor="#FFFFFF" width=150>&nbsp;<%=dutyman%>
					<%
						if dutyman = "" then
					%>
						[<a href="##" onclick="ShowPOPLayer('update_userinfo.asp?receiptfactnum=<%=receiptfactnum%>&dutyman=<%=dutyman%>','320','200');">수사관코드입력 </a>
					<%
						else
					%>
						&nbsp&nbsp;[<a href="##" onclick="ShowPOPLayer('update_userinfo.asp?receiptfactnum=<%=receiptfactnum%>&dutyman=<%=dutyman%>','320','200');">수정 </a>
					<%
						end if
					%>
					]</td>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>성 명</td>
					<td bgcolor="#FFFFFF" width=150>&nbsp;<%=sName%>
					</td>
				</tr>
			</table>

			<table width="100%" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>

			<table width="100%" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#ffffff">
			    <tr>
					<td align="left" bgcolor="#FFFFFF" class="TDCont" ><b><font color="#0000ff">&nbsp;<img src="/Images/dot_01.gif" >&nbsp;사건정보</font></b></td>
				</tr>
			</table>
			<table width="100%" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">
			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>사건번호</td>
					<td bgcolor="#FFFFFF" width=400>&nbsp;<%=receiptfactnum%></td>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>사건장소</td>
					<td bgcolor="#FFFFFF" width=400>&nbsp;<%=occurplace%></td>

				</tr>
			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>사 건 명</td>
					<td bgcolor="#FFFFFF" width=300>&nbsp;<%=nameoffact%>
					</td>

<% if receiptkind = "" then %>					
					<td bgcolor="#FFEEF9" class="TDCont" width=100 align='center'>사건유형</td>
<% else %>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>사건유형</td>
<% end if%>
					<td bgcolor="#FFFFFF" width=300>&nbsp;

<%							'==================================================
							SqlCode = "SELECT CODE, CODENAME FROM TB_CODE"
							SqlCode = SqlCode& " WHERE USEYN='Y' AND SYSYN='Y' AND CODEGROUP='B09'"
							SqlCode = SqlCode& " ORDER BY CODE"
							set RsCode = db.execute(SqlCode)
						%>
						<select name="receiptkind" size="1" class="ComboFFFCE7">
						<Option value =''>사건유형</option>
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("CODE")
										CODENAME = RsCode("CODENAME")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &receiptkind& "")%>
							<%
									RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
							%>
						</select>					</td>
				</tr>
			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>사건개요</td>
					<td bgcolor="#FFFFFF" colspan="3" align='left'><textarea name="contents" style="width:99%; height:40" wrap="soft" class="TextareaInput" readonly><%=contents%></textarea>
					</td>
				</tr>
			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>첨부파일</td>
					<td bgcolor="#FFFFFF">&nbsp;<%=sfilename%></td>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>인지일자</td>
					<td bgcolor="#FFFFFF">&nbsp;<%=inputdate%></td>
				</tr>
			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>조 사 일</td>
					<td bgcolor="#FFFFFF" width=300>&nbsp;<input value="<%=Date2%>" name="Date2" type="text" size="10" onfocus="setFocusColor(this);" >
						&nbsp;<img src="/Images/icon_sche.gif" title="달력" style="cursor:hand;" align="absmiddle" name="Date2_CAR" onClick="window.open('/Include/Calendar_view.asp?goMonth='+document.inUpFrm.Date2.value+'&firstcode=&FlightCode=&StartEndGu=S&ControlName=opener.document.inUpFrm.Date2','','width=800,height=300,top=200,left=300,scrollbars=no,status=no');" onblur="setOutColor(this);">
						&nbsp;<img src="/Images/Btn/BtnIconDel.gif" title="삭제" style="cursor:hand;" align="absmiddle"onclick="document.all.Date2.value='';">

					</td>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>송 치 일</td>
					<td bgcolor="#FFFFFF" width=300>&nbsp;<input value="<%=Date3%>" name="Date3" type="text" size="10" onfocus="setFocusColor(this);" >
						&nbsp;<img src="/Images/icon_sche.gif" title="달력" style="cursor:hand;" align="absmiddle" name="Date3_CAR" onClick="window.open('/Include/Calendar_view.asp?goMonth='+document.inUpFrm.Date3.value+'&firstcode=&FlightCode=&StartEndGu=S&ControlName=opener.document.inUpFrm.Date3','','width=800,height=300,top=200,left=300,scrollbars=no,status=no');" onblur="setOutColor(this);">
						&nbsp;<img src="/Images/Btn/BtnIconDel.gif" title="삭제" style="cursor:hand;" align="absmiddle"onclick="document.all.Date3.value='';">

					</td>
				</tr>

<%


		if ( processgb = "9" ) then	'모니터링 완료	(설문완료자료 생성)
%>
			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>모니터링일시</td>
					<td bgcolor="#FFFFFF">&nbsp;<%=processdate%>
					</td>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>종합평가점수</td>
					<td bgcolor="#FFFFFF">&nbsp;<%=m_monitorpoint%>
					</td>
				</tr>

			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>종합평가</td>
					<td bgcolor="#FFFFFF" colspan='3'>
					<table width="100%" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">


<%

		SQL = "SELECT section2, REMARK from armyinformix.dbo.factpeople where factnum = '" & receiptfactnum & "' and REMARK is not null and datalength(REMARK)>0"
		set RSRemark = db.execute(SQL)
		do until RSRemark.eof	
			s_section2 = RSRemark(0)
			s_REMARK = RSRemark(1)
%>
						<tr><td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>&nbsp;<%=db_getCodeName("B01",s_section2)%></td>
							<td bgcolor="#FFFFFF">&nbsp;<%=s_REMARK%>
							</td>
						</tr>
<%

			RSRemark.movenext
		loop

%>
					</table>

				</tr>

<%
		end if

%>
			</table>
		</td>
	</tr>
</table>

<!--<table width="940" height="10" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>-->
<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
<table width="920" height="30" border="0" cellspacing="0" cellpadding="0" align="center">
	<tr>
		<td colspan= 4 class="TDCont" align='right'>
		<img name="btnsave" src="/Images/Btn/BtnMointorSubmit.GIF" style="cursor:hand;" align="absmiddle" title="설문내용저장" onclick="ListFrame.fn_UpdateData('1','1');"><% if FRM <> "list" then%>&nbsp;<img src="/Images/Btn/BtnList.gif" style="cursor:hand;" align="absmiddle" onClick="fn_list();"><%end if%></td>
	</tr>
</table>
<!--

<table border="0" width="940" cellspacing="0" cellpadding="0" align="center">
	<tr height="22">
		<td colspan="3" background="/Images/AsRegi/TabBG.gif">
			<span id="TabA" style="display:block;">
				<table border="0" cellspacing="0" cellpadding="0">
					<tr>
						<td><img src="/Images/AsRegi/Tab01.jpg"></td>
						<td><img src="/Images/AsRegi/TabA02.jpg" style="cursor:hand;" onClick="TabDisplay('B','TabB');"></td>
						<td><img src="/Images/AsRegi/TabA03.jpg" style="cursor:hand;" onClick="TabDisplay('C','TabC');"></td>
						<td><img src="/Images/AsRegi/TabA04.jpg" style="cursor:hand;" onClick="TabDisplay('D','TabD');"></td>
						<td><img src="/Images/AsRegi/TabA05.jpg" style="cursor:hand;" onClick="TabDisplay('E','TabE');"></td>
						<td><img src="/Images/AsRegi/TabA06.jpg" style="cursor:hand;" onClick="TabDisplay('F','TabF');"></td>
						<td><img src="/Images/AsRegi/TabA07.jpg" style="cursor:hand;" onClick="TabDisplay('G','TabG');"></td>

					</tr>
				</table>
			</span>
			<span id="TabB" style="display:none;">
				<table border="0" cellspacing="0" cellpadding="0">
					<tr>
						<td><img src="/Images/AsRegi/TabA01.jpg" style="cursor:hand;" onClick="TabDisplay('A','TabA');"></td>
						<td><img src="/Images/AsRegi/Tab02.jpg"></td>
						<td><img src="/Images/AsRegi/TabA03.jpg" style="cursor:hand;" onClick="TabDisplay('C','TabC');"></td>
						<td><img src="/Images/AsRegi/TabA04.jpg" style="cursor:hand;" onClick="TabDisplay('D','TabD');"></td>
						<td><img src="/Images/AsRegi/TabA05.jpg" style="cursor:hand;" onClick="TabDisplay('E','TabE');"></td>
						<td><img src="/Images/AsRegi/TabA06.jpg" style="cursor:hand;" onClick="TabDisplay('F','TabF');"></td>
						<td><img src="/Images/AsRegi/TabA07.jpg" style="cursor:hand;" onClick="TabDisplay('G','TabG');"></td>
					</tr>
				</table>
			</span>
			<span id="TabC" style="display:none;">
				<table border="0" cellspacing="0" cellpadding="0">
					<tr>
						<td><img src="/Images/AsRegi/TabA01.jpg" style="cursor:hand;" onClick="TabDisplay('A','TabA');"></td>
						<td><img src="/Images/AsRegi/TabA02.jpg" style="cursor:hand;" onClick="TabDisplay('B','TabB');"></td>
						<td><img src="/Images/AsRegi/TabA03.jpg" style="cursor:hand;" onClick="TabDisplay('C','TabC');"></td>
						<td><img src="/Images/AsRegi/TabA04.jpg" style="cursor:hand;" onClick="TabDisplay('D','TabD');"></td>
						<td><img src="/Images/AsRegi/TabA05.jpg" style="cursor:hand;" onClick="TabDisplay('E','TabE');"></td>
						<td><img src="/Images/AsRegi/TabA06.jpg" style="cursor:hand;" onClick="TabDisplay('F','TabF');"></td>
						<td><img src="/Images/AsRegi/TabA07.jpg" style="cursor:hand;" onClick="TabDisplay('G','TabG');"></td>
					</tr>
				</table>
			</span>
			<span id="TabD" style="display:none;">
				<table border="0" cellspacing="0" cellpadding="0">
					<tr>
						<td><img src="/Images/AsRegi/TabA01.jpg" style="cursor:hand;" onClick="TabDisplay('A','TabA');"></td>
						<td><img src="/Images/AsRegi/TabA02.jpg" style="cursor:hand;" onClick="TabDisplay('B','TabB');"></td>
						<td><img src="/Images/AsRegi/TabA03.jpg" style="cursor:hand;" onClick="TabDisplay('C','TabC');"></td>
						<td><img src="/Images/AsRegi/TabA04.jpg" style="cursor:hand;" onClick="TabDisplay('D','TabD');"></td>
						<td><img src="/Images/AsRegi/TabA05.jpg" style="cursor:hand;" onClick="TabDisplay('E','TabE');"></td>
						<td><img src="/Images/AsRegi/TabA06.jpg" style="cursor:hand;" onClick="TabDisplay('F','TabF');"></td>
						<td><img src="/Images/AsRegi/TabA07.jpg" style="cursor:hand;" onClick="TabDisplay('G','TabG');"></td>
					</tr>
				</table>
			</span>
			<span id="TabE" style="display:none;">
				<table border="0" cellspacing="0" cellpadding="0">
					<tr>
						<td><img src="/Images/AsRegi/TabA01.jpg" style="cursor:hand;" onClick="TabDisplay('A','TabA');"></td>
						<td><img src="/Images/AsRegi/TabA02.jpg" style="cursor:hand;" onClick="TabDisplay('B','TabB');"></td>
						<td><img src="/Images/AsRegi/TabA03.jpg" style="cursor:hand;" onClick="TabDisplay('C','TabC');"></td>
						<td><img src="/Images/AsRegi/TabA04.jpg" style="cursor:hand;" onClick="TabDisplay('D','TabD');"></td>
						<td><img src="/Images/AsRegi/TabA05.jpg" style="cursor:hand;" onClick="TabDisplay('E','TabE');"></td>
						<td><img src="/Images/AsRegi/TabA06.jpg" style="cursor:hand;" onClick="TabDisplay('F','TabF');"></td>
						<td><img src="/Images/AsRegi/TabA07.jpg" style="cursor:hand;" onClick="TabDisplay('G','TabG');"></td>
					</tr>
				</table>
			</span>
			<span id="TabF" style="display:none;">
				<table border="0" cellspacing="0" cellpadding="0">
					<tr>
						<td><img src="/Images/AsRegi/TabA01.jpg" style="cursor:hand;" onClick="TabDisplay('A','TabA');"></td>
						<td><img src="/Images/AsRegi/TabA02.jpg" style="cursor:hand;" onClick="TabDisplay('B','TabB');"></td>
						<td><img src="/Images/AsRegi/TabA03.jpg" style="cursor:hand;" onClick="TabDisplay('C','TabC');"></td>
						<td><img src="/Images/AsRegi/TabA04.jpg" style="cursor:hand;" onClick="TabDisplay('D','TabD');"></td>
						<td><img src="/Images/AsRegi/TabA05.jpg" style="cursor:hand;" onClick="TabDisplay('E','TabE');"></td>
						<td><img src="/Images/AsRegi/TabA06.jpg" style="cursor:hand;" onClick="TabDisplay('F','TabF');"></td>
						<td><img src="/Images/AsRegi/TabA07.jpg" style="cursor:hand;" onClick="TabDisplay('G','TabG');"></td>
					</tr>
				</table>
			</span>
			<span id="TabG" style="display:none;">
				<table border="0" cellspacing="0" cellpadding="0">
					<tr>
						<td><img src="/Images/AsRegi/TabA01.jpg" style="cursor:hand;" onClick="TabDisplay('A','TabA');"></td>
						<td><img src="/Images/AsRegi/TabA02.jpg" style="cursor:hand;" onClick="TabDisplay('B','TabB');"></td>
						<td><img src="/Images/AsRegi/TabA03.jpg" style="cursor:hand;" onClick="TabDisplay('C','TabC');"></td>
						<td><img src="/Images/AsRegi/TabA04.jpg" style="cursor:hand;" onClick="TabDisplay('D','TabD');"></td>
						<td><img src="/Images/AsRegi/TabA05.jpg" style="cursor:hand;" onClick="TabDisplay('E','TabE');"></td>
						<td><img src="/Images/AsRegi/TabA06.jpg" style="cursor:hand;" onClick="TabDisplay('F','TabF');"></td>
						<td><img src="/Images/AsRegi/TabA07.jpg" style="cursor:hand;" onClick="TabDisplay('G','TabG');"></td>
					</tr>
				</table>
			</span>
		</td>
	</tr>

</table>-->


</form>
<table border="0" width="940" cellspacing="0" cellpadding="0" align="center">
	<tr height=580>
		<td width="940" align="left">
			<iframe src="monitoring_input.asp?receiptfactnum=<%=receiptfactnum%>&sGubun=B11&sGubunName=피의자" name="ListFrame" width="940" height="100%" frameborder=0 marginheight=0 marginwidth=0 scrolling="yes"></iframe>
		</td>
	</tr>
</table>

<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
<table width="920" height="30" border="0" cellspacing="0" cellpadding="0" align="center">
	<tr>
		<td colspan= 4 class="TDCont" align='right'>
		<img name="btnsave1" src="/Images/Btn/BtnMointorSubmit.GIF" style="cursor:hand;" align="absmiddle" title="설문내용저장" onclick="ListFrame.fn_UpdateData('1','1');"><% if FRM <> "list" then%>&nbsp;<img src="/Images/Btn/BtnList.gif" style="cursor:hand;" align="absmiddle" onClick="fn_list();"><%end if%></td>
	</tr>
</table>

<script>
	function fn_update(arg0,arg1) {	
		location.href="/menu01/submenu0104/설문대상목록4_1.asp";
	}
	function fn_list(){

		if (document.all.FRM.value == "submenu01" )
		{
			location.href="/menu01/submenu0101/research01.asp?<%=where2%>";
		}
		else if (document.all.FRM.value == "submenu02" )
		{
			location.href="/menu01/submenu0102/research02.asp?<%=where2%>";
		}
		else
		{
			location.href="/menu01/submenu0103/research03.asp?<%=where2%>";
		}
		
	}


</script>
<%
	'informixDB.close
	Set informixDB=nothing
%>
<!-- #include virtual="/Include/Bottom.asp" -->