<!-- #include virtual="/Include/Top.asp" -->
<%
	'On Error Resume next



	DBConnect = "Provider=SQLOLEDB.1;Password=1qaz;Persist Security Info=True;User ID=sa;Initial Catalog=LIFECALLCENTER;Data Source=1.1.147.32"
	'DBConnect = "Provider=SQLOLEDB.1;Password=ctisvr;Persist Security Info=True;User ID=ctisvr;Initial Catalog=ARMYCC;Data Source=210.91.80.141"
	Set db = Server.CreateObject("ADODB.Connection") 
	db.Open DBConnect



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
		FromDate = "2011-12-01"
	end if
	if ToDate = "" then
		ToDate = "2011-12-31"
	end if

	'2. 쿼리조건절 셋팅
	if QueryYN = "Y" then

		pageSize = 10
		pageSector = 10
		if curPage = "" then curPage = 1 end If

		where1 = "FromDate=" & FromDate & "&ToDate=" & ToDate & "&whereCD1=" & whereCD1 & "&whereCD2=" & whereCD2 & "&whereCD3=" & whereCD3 & "&whereCD4=" & whereCD4 
		where2 = "curPage=" & curPage & "&" & where1



		SQL = "	SELECT  count(*) cnt"
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

		SQL = SQL & "	AND		Substring(d.RecordingTitle,7,6) in ( 'user08','user09','user10','user11','user12')"

		if whereCD3 = "Y" then
		'Call, user13 recorded on 2009-07-09
				SQL = SQL & "	AND		left(d.RecordingTitle,13) = 'Call recorded'"
		else
			if SS_Login_Secgroup = "A" then	'상담관일때는 내것만
	'-----------------------------------------------------------------------------------------------
				SQL = SQL & "	AND		( Substring(d.RecordingTitle,7,6) = '" & SS_Login_CTIID & "' or Substring(d.RecordingTitle,7,6) in ('" & SS_Login_EXTNO &"'))"
	'-----------------------------------------------------------------------------------------------
			elseif SS_Login_Secgroup = "B" then	'관리자일때는 팀원것
	'-----------------------------------------------------------------------------------------------
				SQL = SQL & "	AND		Substring(d.RecordingTitle,7,6) in ( select ctiid from tb_userinfo where GRADE = '"&SS_Login_Grade&"')"
	'-----------------------------------------------------------------------------------------------
			end if
		end if

		'Response.write SQL
		set Rs = db.execute(SQL)

		j = rs("cnt") + 1


		SQL = "	SELECT  convert(char(19),dateadd(hour,9,d.recordingdate),121) as RecDate, c.ANI, c.CallDirection"
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

		SQL = SQL & "	AND		Substring(d.RecordingTitle,7,6) in ( 'user08','user09','user10','user11','user12')"

		if whereCD3 = "Y" then
		'Call, user13 recorded on 2009-07-09
				SQL = SQL & "	AND		left(d.RecordingTitle,13) = 'Call recorded'"
		else
			if SS_Login_Secgroup = "A" then	'상담관일때는 내것만
	'-----------------------------------------------------------------------------------------------
				SQL = SQL & "	AND		( Substring(d.RecordingTitle,7,6) = '" & SS_Login_CTIID & "' or Substring(d.RecordingTitle,7,6) in ('" & SS_Login_EXTNO &"'))"
	'-----------------------------------------------------------------------------------------------
			elseif SS_Login_Secgroup = "B" then	'관리자일때는 팀원것
	'-----------------------------------------------------------------------------------------------
				SQL = SQL & "	AND		Substring(d.RecordingTitle,7,6) in ( select ctiid from tb_userinfo where GRADE = '"&SS_Login_Grade&"')"
	'-----------------------------------------------------------------------------------------------
			end if
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
<table border="0" width="1200" cellpadding="0" cellspacing="2" bgcolor="#EFECE5" align="center">
	<tr bgcolor="#FFFFFF">
		<td>
			<form name="inUpFrm" method="post" action="RecordList2011.asp" onsubmit="return fn_Search(this);" style="margin:0">
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
						<select name="whereCD1" size="1" class="ComboFFFCE7">
							<option value="">전체</option>
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
					<td align='center' bgcolor="#EEF6FF" class="TDCont">IN/OUT</td>	
					<td align='center' bgcolor="#EEF6FF" class="TDCont">회선구분</td>	
					<td align='center' bgcolor="#EEF6FF" class="TDCont">전화번호</td>						
				</tr>
				<tr><td colspan="12" height="1" bgcolor="#FFFFFF"></td></tr>

<%'####### 실제자료가 들어간다. %>

<%

	if QueryYN = "Y" then

				i = 0
				do until Rs.eof

				    j = j - 1
					db_RecDate = Rs("RecDate")
					db_ANI = Rs("ANI")
					if Rs("CallDirection") = "O" then
						db_CallDirection = "아웃"
						IOFLAG = "2"
					else
						db_CallDirection = "인"
						IOFLAG = "1"
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
						LINEKIND = "SIP-Analog"
					else
						sGubun = "군전화"	
						LINEKIND = "SIP-DigitalE1"
					end if
					if instr(db_ANI,"anonymous") >0 then						
						db_ANI = ""
					else
						db_ANI = replace(db_ANI,"sip:","")
						db_ANI = replace(db_ANI,"@16.1.17.117:5060","")
						db_ANI = replace(db_ANI,"@16.1.153.6:5060","")
					end if
					IF LEN(db_ANI) = 9 AND LEFT(db_ANI,1) <> "0" THEN
						db_ANI = "0"&db_ANI
					END IF
					sdb_RecordingFileName = "http://16.1.17.113:8080/2011/"&mid(replace(db_RecordingFileName,"\","/"),42)
					sssdb_RecordingFileName = mid(replace(sdb_RecordingFileName,"\","/"),42)

					
					db_RecordingFileSize = Rs("RecordingFileSize")
					db_RecordingFileName = Rs("RecordingFileName")
					db_RecordingFileSize = db_RecordingFileSize / 10.75

					lv_CallTime = Fix(db_RecordingFileSize / 100)
					sdb_RecordingFileName = "http://1.1.147.31:8080/2011/"&mid(replace(db_RecordingFileName,"\","/"),42)
					sssdb_RecordingFileName = mid(replace(sdb_RecordingFileName,"\","/"),1)

%>
					<tr height="20" bgcolor="<%=sBgColor%>">
						<td align='center'><%=j%></td>
						<td colspan='2' align='center'><%=db_RecDate%></td>
						
						<td align='center'><%=lv_Hour & ":" & lv_Min & ":" & lv_Sec%></td>
						<td align='left' >&nbsp;<a href="##" onClick="fn_Player('<%=sdb_RecordingFileName%>');"><%=CutString(sssdb_RecordingFileName, 100, "...")%></a></td>

						<td align='center' ><img src="/Images/Comm/IconAlert.gif" align="absmiddle" style="cursor:hand;" onClick="fn_Player('<%=sdb_RecordingFileName%>');" title="녹음내용 청취"></td>
						<td align='center' ><a href='<%=URL%>'><img src="/Images/Comm/IconHome.gif" align="absmiddle" style="cursor:hand;" title="상담접수" border=0></a></td>
						<td colspan='2' align='center'><%=db_UserId%></td>		
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

//		alert(arg0);
		ShowPOPLayer("/include/WavePlayer.asp?URL="+arg0,'300','200');	
		//window.open("/include/WavePlayer.asp?URL="+arg0,"Player", "toolbar=no,top=100,left=300,width=300,height=200,resize=no,status=yes, scrollbars=no");
	}

//-->
</script>


<!-- #include virtual="/Include/Bottom.asp" -->