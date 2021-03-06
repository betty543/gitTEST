<!-- #include virtual="/Include/Top.asp" -->
<%
	'On Error Resume next

	SS_Login_Secgroup = SESSION("SS_Login_Secgroup")
	SS_Login_Grade = SESSION("SS_Login_Grade")
	SS_Login_CTIID = SESSION("SS_Login_CTIID")
	SS_Login_EXTNO = SESSION("SS_Login_EXTNO")
	SS_LoginID = SESSION("SS_LoginID")
	FromDate = Request("FromDate")
	ToDate = Request("ToDate")

	curPage = request("curPage")
	whereCD1 = Trim(request("whereCD1"))
	whereCD2 = Trim(request("whereCD2"))
	whereCD3 = Trim(request("whereCD3"))
	QueryYN = Trim(request("QueryYN"))

	if FromDate = "" then
		FromDate = date()
	end if
	if ToDate = "" then
		ToDate =  date()
	end if

	'2. 쿼리조건절 셋팅
	if QueryYN = "Y" then

		pageSize = 10
		pageSector = 10
		if curPage = "" then curPage = 1 end If

		where1 = "FromDate=" & FromDate & "&ToDate=" & ToDate & "&whereCD1=" & whereCD1 & "&whereCD2=" & whereCD2 & "&whereCD3=" & whereCD3 & "&whereCD4=" & whereCD4 
		where2 = "curPage=" & curPage & "&" & where1




		SQL = "	select sum(cnt) cnt from ( SELECT	count(*) cnt"
		SQL = SQL & "	FROM	I3_IC.dbo.RecordingCall c, I3_IC.dbo.RecordingData d"
		SQL = SQL & "	WHERE	c.RecordingID = d.RecordingID "
		SQL = SQL & "	AND		convert(char(10),dateadd(hour,9,d.recordingdate),121) >= '" & FromDate & "'"
		if ToDate <> "" then
			SQL = SQL & "	AND		convert(char(10),dateadd(hour,9,d.recordingdate),121) <= '" & ToDate & "'"
		end  if
		if whereCD1 <> "" then
			'상담관ID
			SQL = SQL & "	AND		recordedCallidKey in ( select RecordingCallKey from TB_RecordingData  where		convert(char(10),recordstarttime,121) >= '" & FromDate & "'"
			if ToDate <> "" then
				SQL = SQL & "	AND		convert(char(10),recordstarttime,121) <= '" & ToDate & "'"
			end  if
			SQL = SQL & "	AND	 Userid = '" & whereCD1 & "')" 
		end if

		if whereCD2 <> "" then
			'상담관ID
			SQL = SQL & "	AND		c.ani like '%" & whereCD2 & "%'"
		end if

		if whereCD3 = "Y" then
		'Call, user13 recorded on 2009-07-09
		'		SQL = SQL & "	AND		left(d.RecordingTitle,13) = 'Call recorded'"
			SQL = SQL & "	AND 1 = 0"
		else
			if SS_Login_Secgroup = "A" then	'상담관일때는 내것만
	'-----------------------------------------------------------------------------------------------			
				'상담관ID
				SQL = SQL & "	AND		recordedCallidKey in ( select RecordingCallKey from TB_RecordingData  where		Userid = '" & SS_LoginID & "')" 

	'-----------------------------------------------------------------------------------------------
			'elseif SS_Login_Secgroup = "B" then	'관리자일때는 팀원것
	'-----------------------------------------------------------------------------------------------
				'SQL = SQL & "	AND		Substring(d.RecordingTitle,7,6) in ( select ctiid from tb_userinfo where GRADE = '"&SS_Login_Grade&"')"
	'-----------------------------------------------------------------------------------------------
			end if
		end if



		SQL = SQL & "	union select count(*) cnt"
		SQL = SQL & "	FROM	tb_RecordingData d"
		SQL = SQL & "	where		convert(char(10),d.recordstarttime,121) >= '" & FromDate & "'"
		if ToDate <> "" then
			SQL = SQL & "	AND		convert(char(10),d.recordstarttime,121) <= '" & ToDate & "'"
		end  if
		if whereCD1 <> "" then
			'상담관ID
			SQL = SQL & "	AND		d.userid = '" & whereCD1 & "'"
		end if
		if whereCD2 <> "" then
			'상담관ID
			SQL = SQL & "	AND		remoteid2 like '%" & whereCD2 & "%'"
		end if
		SQL = SQL & "	AND RecordFileName <> 'none'"

		if whereCD3 = "Y" then
					SQL = SQL & "	AND 1 = 1"
		'Call, user13 recorded on 2009-07-09
		'		SQL = SQL & "	AND		left(d.RecordingTitle,13) = 'Call recorded'"
		else
			if SS_Login_Secgroup = "A" then	'상담관일때는 내것만
	'-----------------------------------------------------------------------------------------------
				SQL = SQL & "	AND		d.userid = '" & SS_LoginID & "'"
	'-----------------------------------------------------------------------------------------------
			'elseif SS_Login_Secgroup = "B" then	'관리자일때는 팀원것
	'-----------------------------------------------------------------------------------------------
				'SQL = SQL & "	AND		Substring(d.RecordingTitle,7,6) in ( select ctiid from tb_userinfo where GRADE = '"&SS_Login_Grade&"')"
	'-----------------------------------------------------------------------------------------------
			end if
		end if
					SQL = SQL & "	) b"


'response.write SQL
'response.end


				set Rs = db.execute(SQL)

				j = rs("cnt") + 1



'response.write j

		SQL = "	SELECT	convert(char(19),dateadd(hour,9,d.recordingdate),121) as RecDate, c.ANI, c.CallDirection"
		SQL = SQL & ",	e.userid " '(select top 1 userid from TB_RecordingData where RecordingCallKey = c.recordedCallidKey) as userid"
		SQL = SQL & ",	e.dnis  " '(select top 1 dnis from TB_RecordingData where RecordingCallKey = c.recordedCallidKey) as dnis"
		SQL = SQL & ",	d.RecordingFileSize, d.RecordingFileName, c.recordedCallidKey RecordingCallKey"
		SQL = SQL & "	FROM	I3_IC.dbo.RecordingCall c inner join I3_IC.dbo.RecordingData d "
		SQL = SQL & "	on	c.RecordingID = d.RecordingID "
		SQL = SQL & "	left join TB_RecordingData  e"
		SQL = SQL & "	on	c.recordedCallidKey  = e.RecordingCallKey "	
		SQL = SQL & "	where		convert(char(10),dateadd(hour,9,d.recordingdate),121) >= '" & FromDate & "'"
		if ToDate <> "" then
			SQL = SQL & "	AND		convert(char(10),dateadd(hour,9,d.recordingdate),121) <= '" & ToDate & "'"
		end  if
		if whereCD1 <> "" then
			'상담관ID
			SQL = SQL & "	AND		recordedCallidKey in ( select RecordingCallKey from TB_RecordingData  where		convert(char(10),recordstarttime,121) >= '" & FromDate & "'"
			if ToDate <> "" then
				SQL = SQL & "	AND		convert(char(10),recordstarttime,121) <= '" & ToDate & "'"
			end  if
			SQL = SQL & "	AND	 Userid = '" & whereCD1 & "')" 
		end if

		if whereCD2 <> "" then
			'상담관ID
			SQL = SQL & "	AND		c.ani like '%" & whereCD2 & "%'"
		end if

		if whereCD3 = "Y" then
		'Call, user13 recorded on 2009-07-09
		'		SQL = SQL & "	AND		left(d.RecordingTitle,13) = 'Call recorded'"
			SQL = SQL & "	AND 1 = 0"
		else
			if SS_Login_Secgroup = "A" then	'상담관일때는 내것만
	'-----------------------------------------------------------------------------------------------			
				'상담관ID
				SQL = SQL & "	AND (		recordedCallidKey in ( select RecordingCallKey from TB_RecordingData  where		Userid = '" & SS_LoginID & "')" 
				SQL = SQL & "	or		recordedCallidKey not in ( select RecordingCallKey from TB_RecordingData where userid in ( select userid from tb_userinfo) ) )" 

	'-----------------------------------------------------------------------------------------------
			'elseif SS_Login_Secgroup = "B" then	'관리자일때는 팀원것
	'-----------------------------------------------------------------------------------------------
				'SQL = SQL & "	AND		Substring(d.RecordingTitle,7,6) in ( select ctiid from tb_userinfo where GRADE = '"&SS_Login_Grade&"')"
	'-----------------------------------------------------------------------------------------------
			end if
		end if



		SQL = SQL & "	union SELECT	convert(char(19),d.recordstarttime,121) as RecDate, remoteid2 ANI, 'R' CallDirection"
		SQL = SQL & ",	userid, dnis, recordduration RecordingFileSize, RecordFileName RecordingFileName, RecordingCallKey"
		SQL = SQL & "	FROM	tb_RecordingData d"
		SQL = SQL & "	where		convert(char(10),d.recordstarttime,121) >= '" & FromDate & "'"
		if ToDate <> "" then
			SQL = SQL & "	AND		convert(char(10),d.recordstarttime,121) <= '" & ToDate & "'"
		end  if
		if whereCD1 <> "" then
			'상담관ID
			SQL = SQL & "	AND		d.userid = '" & whereCD1 & "'"
		end if
		if whereCD2 <> "" then
			'상담관ID
			SQL = SQL & "	AND		remoteid2 like '%" & whereCD2 & "%'"
		end if
		SQL = SQL & "	AND RecordFileName <> 'none'"

		if whereCD3 = "Y" then
					SQL = SQL & "	AND 1 = 1"
		'Call, user13 recorded on 2009-07-09
		'		SQL = SQL & "	AND		left(d.RecordingTitle,13) = 'Call recorded'"
		else
			if SS_Login_Secgroup = "A" then	'상담관일때는 내것만
	'-----------------------------------------------------------------------------------------------
				SQL = SQL & "	AND		d.userid = '" & SS_LoginID & "'"
	'-----------------------------------------------------------------------------------------------
			'elseif SS_Login_Secgroup = "B" then	'관리자일때는 팀원것
	'-----------------------------------------------------------------------------------------------
				'SQL = SQL & "	AND		Substring(d.RecordingTitle,7,6) in ( select ctiid from tb_userinfo where GRADE = '"&SS_Login_Grade&"')"
	'-----------------------------------------------------------------------------------------------
			end if
		end if

		SQL = SQL & "	order by 1 desc"
		'Response.write SQL

'response.end
		set Rs = db.execute(SQL)


	end if
%>

<!-- #include virtual="/Include/PopLayer.asp" -->

<div id='CalendarLayer' style='display:none;'><iframe name='CalendarFrame' src='/include/Calendar.asp' width='172' height='177' border='0' frameborder='0' scrolling='no'></iframe></div>

<script>
	function fn_Search() {
		document.inUpFrm.QueryYN.value = "Y";
		document.inUpFrm.submit();
	}
	function fn_Search1() {
		document.inUpFrm.QueryYN.value = "Y";
		document.inUpFrm.submit();
	}
	function fn_GetId(arg,arg1) {
		document.inUpFrm.agt_id.value =arg;
		document.inUpFrm.FileName.value =arg1;
	}

</script>
<table border="0" width="1200" cellpadding="0" cellspacing="2" bgcolor="#EFECE5" align="center">
	<tr bgcolor="#FFFFFF">
		<td>
			<form name="inUpFrm" method="post" action="RecordList.asp" onsubmit="return fn_Search(this);" style="margin:0">
			<input type="hidden" name="QueryYN" value="<%=QueryYN%>">	
			<table width="1200" border="0" cellspacing="1" cellpadding="0" style="border:#E1DED6 solid 1px" bgcolor="#CCCCCC">
			    <tr bgcolor="#ffffff" height="30">
			        <td class="TDCont"  width="100" align="center" bgcolor="#EEF6FF">조회기간</td>
			        <td class="TDCont">
			        	<input value="<%=FromDate%>" name="FromDate" readonly type="text" size="10" onfocus="setFocusColor(this); new CalendarFrame.Calendar(this);" onblur="setOutColor(this);">
				    	~
				    	<input value="<%=ToDate%>" name="ToDate" readonly type="text" size="10" onfocus="setFocusColor(this); new CalendarFrame.Calendar(this);" onblur="setOutColor(this);">
					</td>
			        <td class="TDCont"  width="100" align="center" bgcolor="#EEF6FF">상담관</td>
			        <td class="TDCont">
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
							<option value="">선택</option>
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
							%>
						<%
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
			        <td class="TDCont"  width="100" align="center" bgcolor="#EEF6FF">전화번호</td>
			        <td class="TDCont">
						<input value="<%=whereCD2%>" name="whereCD2" type="text" size="14" onfocus="setFocusColor(this);" onblur="setOutColor(this);">
			        </td>
			        <td class="TDCont"  width="100" align="center" bgcolor="#EEF6FF"><input type="checkbox" name="whereCD3" value="Y" class="none" <% if whereCD3="Y" then Response.Write("checked") end if %>>착신통화만</td>
			        <td class="TDCont">
									
						<img src="/Images/Btn/BtnSearch.gif" align="absmiddle" style="cursor:hand;" onClick="fn_Search();">
			        </td>
			    </tr>
			</table>
			</form>
		</td>
	</tr>
</table>

<table border="0" width="1200" cellpadding="0" cellspacing="0" align="center"><tr height="5"><td></td></tr></table>
<table width="1200" border="0" cellpadding="0" cellspacing="0" align="center">
	<tr>
		<td align="center">
			<DIV style="OVERFLOW-Y:auto; OVERFLOW-X:auto; MARGIN: 0px 0px 0px 0px; WIDTH:100%; HEIGHT:600;">
			<table width="100%" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">

				<tr height="20" bgcolor="#FFFFFF">
					<td align='center' bgcolor="#EEF6FF" class="TDCont">No</td>
					<td colspan='2' align='center' bgcolor="#EEF6FF" class="TDCont">녹취일시</td>
					<td align='center' bgcolor="#EEF6FF" class="TDCont">통화시간</td>
					<td align='center' bgcolor="#EEF6FF" class="TDCont">파일명</td>
					<td align='center' bgcolor="#EEF6FF" class="TDCont"></td>
					<td align='center' bgcolor="#EEF6FF" class="TDCont"></td>
					<td colspan='2' align='center' bgcolor="#EEF6FF" class="TDCont">사용자</td>		
					<td align='center' bgcolor="#EEF6FF" class="TDCont"></td>
					<td align='center' bgcolor="#EEF6FF" class="TDCont">IN/OUT</td>	
					<td align='center' bgcolor="#EEF6FF" class="TDCont">회선구분</td>	
					<td align='center' bgcolor="#EEF6FF" class="TDCont">회선구분</td>	
					<td align='center' bgcolor="#EEF6FF" class="TDCont">전화번호</td>						
				</tr>
				<tr><td colspan="14" height="1" bgcolor="#FFFFFF"></td></tr>

<%'####### 실제자료가 들어간다. %>

<%

	if QueryYN = "Y" then


				do until Rs.eof



					j = j - 1
					db_RecDate = Rs("RecDate")
					db_ANI = Rs("ANI")
					db_DNIS = Rs("dnis")
					if Rs("CallDirection") = "R" then
						db_CallDirection = "착신"
						IOFLAG = "1"
						db_RecordingFileSize = Rs("RecordingFileSize")
						db_RecordingFileName = Rs("RecordingFileName")
						db_RecordingFileName = replace(db_RecordingFileName,".i3r",".wav")

						lv_CallTime = db_RecordingFileSize
						sdb_RecordingFileName = "http://1.1.147.31:8081/"&db_RecordingFileName
						sssdb_RecordingFileName = mid(replace(sdb_RecordingFileName,"\","/"),1)

					elseif Rs("CallDirection") = "O" then
						db_CallDirection = "아웃"
						IOFLAG = "2"
						db_RecordingFileSize = Rs("RecordingFileSize")
						db_RecordingFileName = Rs("RecordingFileName")
						db_RecordingFileName = replace(db_RecordingFileName,".i3r",".wav")
						db_RecordingFileSize = db_RecordingFileSize / 10.75

						lv_CallTime = Fix(db_RecordingFileSize / 100)
						sdb_RecordingFileName = "http://1.1.147.31:8080/"&mid(replace(db_RecordingFileName,"\","/"),17)
						sssdb_RecordingFileName = mid(replace(sdb_RecordingFileName,"\","/"),1)

					else
						db_CallDirection = "인"
						IOFLAG = "1"
						db_RecordingFileSize = Rs("RecordingFileSize")
						db_RecordingFileName = Rs("RecordingFileName")
						db_RecordingFileName = replace(db_RecordingFileName,".i3r",".wav")
						db_RecordingFileSize = db_RecordingFileSize / 10.75

						lv_CallTime = Fix(db_RecordingFileSize / 100)
						sdb_RecordingFileName = "http://1.1.147.31:8080/"&mid(replace(db_RecordingFileName,"\","/"),17)
						sssdb_RecordingFileName = mid(replace(sdb_RecordingFileName,"\","/"),1)

					end if
					db_UserId = db_getCTIUserName(Rs("UserId"))

					
					lv_Hour = Fix(lv_CallTime / 3600)
					lv_Min = Fix((lv_CallTime - lv_Hour * 3600) / 60)
					lv_Sec = lv_CallTime - ((lv_Hour * 3600) + (lv_Min * 60))

					if lv_Hour < 10 then
						lv_Hour = "0" & lv_Hour
					end if
					if lv_Min < 10 then
						lv_Min = "0" & lv_Min
					end if
					if lv_Sec < 10 then
						lv_Sec = "0" & lv_Sec
					end if

					i = i + 1
					if ( i mod 2 ) = 1 then
						sBgColor = "#ffffff"
					else
						sBgColor = "#FFFCE7"				
					end if


					if instr(db_DNIS,"sip:5001") > 0 or instr(db_DNIS,"sip:5002") > 0 or instr(db_DNIS,"sip:501") > 0  then
						sGubun = "군전화"	
						LINEKIND = "SIP-Analog"
					else
						sGubun = "일반전화"	
						LINEKIND = "SIP-DigitalE1"
					end if
					if instr(db_ANI,"anonymous") >0 then						
						db_ANI = ""
					elseif isnull(db_ANI) = false then
						db_ANI = replace(db_ANI,"sip:","")
						db_ANI = replace(db_ANI,"@16.1.17.117:5060","")
						db_ANI = replace(db_ANI,"@1.1.147.33:5060","")
						db_ANI = replace(db_ANI,"@1.1.147.104:5060","")
						db_ANI = replace(db_ANI,"@1.1.147.111:5060","")
							db_ANI = replace(db_ANI,"@1.1.147.112:5060","")
						db_ANI = replace(db_ANI,"@1.1.147.113:5060","")
					end if
					IF LEN(db_ANI) = 9 AND LEFT(db_ANI,1) <> "0" THEN
						db_ANI = "0"&db_ANI
					END IF

db_RecordingCallKey = rs("RecordingCallKey")

					'if SS_Login_Grade = "B" then
						URL ="/menu03/submenu0302/lifecallmanage.asp?InType=RECORD&RecDate="&db_RecDate&"&LINEKIND="&mid(db_DNIS,5,4)&"&IOFLAG="&IOFLAG&"&telNo="&db_ANI&"&CALLTIME="&lv_Hour & ":" & lv_Min & ":" & lv_Sec&"&FILENAME="&sdb_RecordingFileName
					'else
					'	URL ="/menu04/submenu0402/callmanage.asp?TELKIND="&SS_Login_Grade&"&InType=RECORD&LINEKIND="&LINEKIND&"&IOFLAG="&IOFLAG&"&telNo="&db_ANI&"&CALLTIME="&lv_Hour & ":" & lv_Min & ":" & lv_Sec&"&FILENAME="&sdb_RecordingFileName&"&IOFLAG="&IOFLAG
					'end if

%>
					<tr height="20" bgcolor="<%=sBgColor%>">
						<td align='center'><%=j%></td>
						<td colspan='2' align='center'><%=db_RecDate%></td>
						
						<td align='center'><%=lv_Hour & ":" & lv_Min & ":" & lv_Sec%></td>
						<td align='left' >&nbsp;<a href="##" onClick="fn_Player('<%=sdb_RecordingFileName%>');"><%=CutString(sssdb_RecordingFileName, 100, "...")%></a></td>
						<td align='center' ><img src="/Images/Comm/IconAlert.gif" align="absmiddle" style="cursor:hand;" onClick="fn_Player('<%=sdb_RecordingFileName%>');" title="녹음내용 청취"></td>
						<td align='center' ><img src="/Images/Comm/IconHome.gif" align="absmiddle" style="cursor:hand;" title="상담접수" border=0 onclick="javascript:goto3002('<%=URL%>');"></td>
						<% if db_UserId = db_GetuserName(db_UserId) then %>

						<td colspan='2' align='center'>&nbsp;</td>	
						<% else %>

						<td colspan='2' align='center'><%=db_GetuserName(db_UserId)%></td>		
						<% end if %>
						<td align='center' ><img src="/Images/Comm/IconWOMAN.gif" align="absmiddle" style="cursor:hand;" onClick="fn_Change('<%=db_RecordingCallKey%>','<%=db_UserId%>');" title="상담관변경"></td>
						<td align='center'><%=db_CallDirection%></td>	

						<td align='center'><%=sGubun%></td>		
						<td align='center'><%=mid(db_DNIS,5,4)%></td>	
						
						<td align='center'><%=db_ANI%></td>						
					</tr>
<%					


					Rs.MoveNext
				loop
		end if
%>
				</table>
			</DIV>
		</td>
	</tr>
</table>

<script>
<!--

	function fn_Player(arg0){
		//파일명
		var x,y;
		x = ( screen.width - 300 )/2;
		y = ( screen.height - 200 )/2;

		ShowPOPLayer("/include/WavePlayer.asp?URL="+arg0,'300','200');	
		//window.open("/include/WavePlayer.asp?URL="+arg0,"Player", "toolbar=no,top=100,left=300,width=300,height=200,resize=no,status=yes, scrollbars=no");
	}

	function goto3002(arg0){
		location.href = arg0;
	}

	function fn_Change(arg0, arg1)
	{
		

		var x,y;
		x = ( screen.width - 300 )/2;
		y = ( screen.height - 200 )/2;

		ShowPOPLayer("RecordUp.asp?RecordingCallKey="+arg0+"&UserId="+arg1,'320','130');	
		//window.open("/include/WavePlayer.asp?URL="+arg0,"Player", "toolbar=no,top=100,left=300,width=300,height=200,resize=no,status=yes, scrollbars=no");

	}
//-->
</script>


<!-- #include virtual="/Include/Bottom.asp" -->