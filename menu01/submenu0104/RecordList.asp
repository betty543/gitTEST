<!-- #include virtual="/Include/Top.asp" -->
<%
	'On Error Resume next

	SS_Login_Secgroup = SESSION("SS_Login_Secgroup")
	SS_Login_CTIID = SESSION("SS_Login_CTIID")
	SS_Login_EXTNO = SESSION("SS_Login_EXTNO")

	FromDate = Request("FromDate")
	ToDate = Request("ToDate")

	curPage = request("curPage")
	whereCD1 = Trim(request("whereCD1"))
	whereCD2 = Trim(request("whereCD2"))
	whereCD3 = Trim(request("whereCD3"))
	QueryYN = Trim(request("QueryYN"))

	if FromDate = "" then
		FromDate = left(date(),7) & "-01"
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

		SQL = "	SELECT	convert(char(19),dateadd(hour,9,d.recordingdate),121) as RecDate, c.ANI, c.CallDirection"
		SQL = SQL & ",	Substring(d.RecordingTitle,7,6) as userid, d.RecordingFileSize, d.RecordingFileName"
		SQL = SQL & "	FROM	I3_IC.dbo.RecordingCall c, I3_IC.dbo.RecordingData d"
		SQL = SQL & "	WHERE	c.RecordingID = d.RecordingID "
		SQL = SQL & "	AND		convert(char(10),dateadd(hour,9,d.recordingdate),121) >= '" & FromDate & "'"
		if ToDate <> "" then
			SQL = SQL & "	AND		convert(char(10),dateadd(hour,9,d.recordingdate),121) <= '" & ToDate & "'"
		end  if
		if whereCD1 <> "" then
			'상담관ID
			SQL = SQL & "	AND		Substring(d.RecordingTitle,7,6) = '" & whereCD1 & "'"
		end if
		if whereCD2 <> "" then
			'상담관ID
			SQL = SQL & "	AND		c.ani like '%" & whereCD2 & "%'"
		end if
		if SS_Login_Secgroup = "A" then	'상담관일때는 내것만
'-----------------------------------------------------------------------------------------------
			SQL = SQL & "	AND		( Substring(d.RecordingTitle,7,6) = '" & SS_Login_CTIID & "' or Substring(d.RecordingTitle,7,6) in ('" & SS_Login_EXTNO &"'))"
'-----------------------------------------------------------------------------------------------
		elseif SS_Login_Secgroup = "B" then	'관리자일때는 팀원것
'-----------------------------------------------------------------------------------------------
			SQL = SQL & "	AND		Substring(d.RecordingTitle,7,6) = '" & whereCD1 & "'"
'-----------------------------------------------------------------------------------------------
		end if
		SQL = SQL & "	order by d.recordingdate desc"
		'Response.write SQL
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
<table border="0" width="940" cellpadding="0" cellspacing="2" bgcolor="#EFECE5" align="center">
	<tr bgcolor="#FFFFFF">
		<td>
			<form name="inUpFrm" method="post" action="RecordList.asp" onsubmit="return fn_Search(this);" style="margin:0">
			<input type="hidden" name="QueryYN" value="<%=QueryYN%>">	
			<table width="940" border="0" cellspacing="1" cellpadding="0" style="border:#E1DED6 solid 1px" bgcolor="#CCCCCC">
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
							SqlCode = "SELECT USERID, USERNAME FROM TB_USERINFO"
							SqlCode = SqlCode& " WHERE USEYN='N'  and	outdate >= '"&DateAdd("d",1,DateAdd("m",-1,Date())) &"'"
							SqlCode = SqlCode& " ORDER BY USEYN DESC, GRADE ASC, USERNAME ASC"
							set RsCode = db.execute(SqlCode)

								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("USERID")
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
						<img src="/Images/Btn/BtnSearch.gif" align="absmiddle" style="cursor:hand;" onClick="fn_Search();">
			        </td>
			    </tr>
			</table>
			</form>
		</td>
	</tr>
</table>

<table border="0" width="940" cellpadding="0" cellspacing="0" align="center"><tr height="5"><td></td></tr></table>
<table width="940" border="0" cellpadding="0" cellspacing="0" align="center">
	<tr>
		<td align="center">
			<DIV style="OVERFLOW-Y:auto; OVERFLOW-X:auto; MARGIN: 0px 0px 0px 0px; WIDTH:100%; HEIGHT:600;">
			<table width="100%" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">

				<tr height="20" bgcolor="#FFFFFF">
					<td colspan='2' align='center' bgcolor="#EEF6FF" class="TDCont">녹취일시</td>
					<td align='center' bgcolor="#EEF6FF" class="TDCont">통화시간</td>
					<td align='center' bgcolor="#EEF6FF" class="TDCont">파일명</td>
					<td colspan='2' align='center' bgcolor="#EEF6FF" class="TDCont">사용자</td>		
					<td align='center' bgcolor="#EEF6FF" class="TDCont">IN/OUT</td>	
					<td align='center' bgcolor="#EEF6FF" class="TDCont">회선구분</td>	
					<td align='center' bgcolor="#EEF6FF" class="TDCont">전화번호</td>						
				</tr>
				<tr><td colspan="8" height="1" bgcolor="#FFFFFF"></td></tr>

<%'####### 실제자료가 들어간다. %>

<%

	if QueryYN = "Y" then

				i = 0
				do until Rs.eof

					db_RecDate = Rs("RecDate")
					db_ANI = Rs("ANI")
					if Rs("CallDirection") = "O" then
						db_CallDirection = "아웃"
					else
						db_CallDirection = "인"
					end if
					db_UserId = Rs("UserId")
					db_RecordingFileSize = Rs("RecordingFileSize")
					db_RecordingFileName = Rs("RecordingFileName")
					db_RecordingFileSize = db_RecordingFileSize / 10.75

					lv_CallTime = Fix(db_RecordingFileSize / 100)
					
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


					if instr(db_ANI,"117") > 0 then
						sGubun = "일반전화"	
					else
						sGubun = "군전화"	
					end if
					db_ANI = replace(db_ANI,"sip:","")
					db_ANI = replace(db_ANI,"@16.1.17.117:5060","")
					db_ANI = replace(db_ANI,"@16.1.153.6:5060","")

					sdb_RecordingFileName = "http://16.1.17.113:8080/"&mid(replace(db_RecordingFileName,"\","/"),27)

					sssdb_RecordingFileName = mid(replace(db_RecordingFileName,"\","/"),27)
%>
					<tr height="20" bgcolor="<%=sBgColor%>" onClick="fn_Player('<%=sdb_RecordingFileName%>');">
						<td colspan='2' align='center'><%=db_RecDate%></td>
						
						<td align='center'><%=lv_Hour & ":" & lv_Min & ":" & lv_Sec%></td>
						<td align='left' >&nbsp;<%=CutString(sssdb_RecordingFileName, 40, "...")%></td>
						<td colspan='2' align='center'><%=db_GetuserName(db_UserId)%></td>		
						<td align='center'><%=db_CallDirection%></td>	

						<td align='center'><%=sGubun%></td>							
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

//-->
</script>


<!-- #include virtual="/Include/Bottom.asp" -->