<!-- #include virtual="/Include/Top.asp" -->
<!-- #include virtual="/Include/PopLayer.asp" -->

<%
guboon = request("guboon")
JUBSEQ = request("JUBSEQ")
InType = request("InType")

'-------------------------------------------
'콜백리스트에서 호출될 때
'-------------------------------------------

'InType=CALLBACK&Cate1="&ACLASS&"&Channel=A&CUSTNO="&CUSTNO&"&telNo="&CID&"&Pid="&PID&"&CB_SEQ="&SEQ&"&CALLBACKPHONE="&CALLBANKNO
'InType=CALLBACK&LINEKIND="&LINEKIND&"&telNo="&CID&"&CB_SEQ="&SEQ
SS_LoginID = SESSION("SS_LoginID")
SS_Login_Secgroup = SESSION("SS_Login_Secgroup")

if InType = "RECORD" then

	LINEKIND=request("LINEKIND")
	sCID = request("telNo")
	IOFLAG = request("IOFLAG")
	'Filename
	'통화시간
	CallTIME = request("CallTIME")
	RecFileName = request("FILENAME")
	RecDate = request("RecDate")
	JUBTIME = RecDate

	IF CallTIME <> "" THEN
		CALLTIME1 = LEFT(CallTIME,2)
		CALLTIME2 = MID(CallTIME,4,2)
		CALLTIME3 = MID(CallTIME,7,2)
	END IF

	SQL = "select top 1 * from tb_custinfo where ( cellphone = '"&sCID&"' or homephone = '"&sCID&"' or sendphone = '"&sCID&"') order by modate desc"

	set RsCode = db.execute(SQL)
	if RsCode.eof = false then
		CUSTNO = RsCode("CUSTNO") 
		SOSOKGB_A = RsCode("SOSOKGB_A")
		SOSOKGB_B = RsCode("SOSOKGB_B")
		SOSOKGB_C = RsCode("SOSOKGB_C")
		SOSOKGB_D = RsCode("SOSOKGB_D")
		SOSOKGB_E = RsCode("SOSOKGB_E")
		LEVEL_B = RsCode("LEVEL_B")
		LEVEL_C = RsCode("LEVEL_C")	
		LEVEL_D = RsCode("LEVEL_D")	
		CUSTNAME = RsCode("NAME")
		TELNO = RsCode("CELLPHONE")
		TELNO2 = RsCode("HOMEPHONE")
		SEXGB = RsCode("SEX")	
	else
		'------------
		'전화번호로 찾아보기
		'------------
		if len(sCID) = 7 then	'군전화임.
			SQL = " select top 1 * from tb_armyinfo where aclass < 'O' and telno = '" & left(sCID,3) & "'"
			set RsCode = db.execute(SQL)
			if RsCode.eof = false then
				SOSOKGB_A = RsCode("Aclass")
				SOSOKGB_B = RsCode("Bclass")	
				SOSOKGB_C = RsCode("Cclass")	
				SOSOKGB_D = RsCode("Dclass")
				SOSOKGB_E = RsCode("Eclass")					
			end if
		elseif len(sCID) >= 9 then '일반전화임..
			SQL = " select top 1 * from tb_armyinfo where aclass < 'O' and telno2 like '%" & sCID & "%'"
			set RsCode = db.execute(SQL)
			if RsCode.eof = false then
				SOSOKGB_A = RsCode("Aclass")
				SOSOKGB_B = RsCode("Bclass")	
				SOSOKGB_C = RsCode("Cclass")	
				SOSOKGB_D = RsCode("Dclass")
				SOSOKGB_E = RsCode("Eclass")						
			end if
		end if
	end if

	if LINEKIND = "5001" or LINEKIND = "5002" then
		CHANNELGB_B = "Q01"
		CHANNELGB_C = "Q01A"
	else
		'일반전화
		CHANNELGB_B = "Q01"
		CHANNELGB_C = "Q01C"
	end if
	
elseif InType = "CALLBACK" then

	LINEKIND=request("LINEKIND")
	sCID = request("telNo")
	CB_SEQ = request("CB_SEQ")

	SQL = "select top 1 * from tb_custinfo where ( cellphone = '"&sCID&"' or homephone = '"&sCID&"' or sendphone = '"&sCID&"') order by modate desc"
	set RsCode = db.execute(SQL)
	if RsCode.eof = false then
		CUSTNO = RsCode("CUSTNO") 
		SOSOKGB_A = RsCode("SOSOKGB_A")
		SOSOKGB_B = RsCode("SOSOKGB_B")
		SOSOKGB_C = RsCode("SOSOKGB_C")
		SOSOKGB_D = RsCode("SOSOKGB_D")
		SOSOKGB_E = RsCode("SOSOKGB_E")
		LEVEL_B = RsCode("LEVEL_B")
		LEVEL_C = RsCode("LEVEL_C")	
		LEVEL_D = RsCode("LEVEL_D")	
		CUSTNAME = RsCode("NAME")
		TELNO = RsCode("CELLPHONE")
		TELNO2 = RsCode("HOMEPHONE")
		SEXGB = RsCode("SEX")	
	else
		if len(sCID) = 7 then '군전화임.
			SQL = " select top 1 * from tb_armyinfo where aclass < 'O' and telno = '" & left(sCID,3) & "'"
			set RsCode = db.execute(SQL)
			if RsCode.eof = false then
				SOSOKGB_A = RsCode("Aclass")
				SOSOKGB_B = RsCode("Bclass")	
				SOSOKGB_C = RsCode("Cclass")	
				SOSOKGB_D = RsCode("Dclass")
				SOSOKGB_E = RsCode("Eclass")					
			end if
		elseif len(sCID) >= 9 then '일반전화임..
			SQL = " select top 1 * from tb_armyinfo where aclass < 'O' and telno2 like '%" & sCID & "%'"
			set RsCode = db.execute(SQL)
			if RsCode.eof = false then
				SOSOKGB_A = RsCode("Aclass")
				SOSOKGB_B = RsCode("Bclass")	
				SOSOKGB_C = RsCode("Cclass")	
				SOSOKGB_D = RsCode("Dclass")
				SOSOKGB_E = RsCode("Eclass")						
			end if
		end if
	end if

	if LINEKIND = "5001" or LINEKIND = "5002" then
		CHANNELGB_B = "Q01"
		CHANNELGB_C = "Q01A"
	else
		'일반전화
		CHANNELGB_B = "Q01"
		CHANNELGB_C = "Q01C"
	end if

elseif InType = "CALL" then	'인입콜임.

	LINEKIND=request("LINEKIND")
	sCID = request("telNo")
	IOFLAG = "1"
	CHANNELGB = request("inroot")

	if instr(sCID,"anonymous") > 0 then
		sCID = "anonymous"
	end if
	
	'---------------------------------------
	'번호와 일치하는 고객있는지 찾기
	'---------------------------------------

	SQL = "select top 1 * from tb_custinfo where ( cellphone = '"&sCID&"' or homephone = '"&sCID&"' or sendphone = '"&sCID&"') order by modate desc"

	set RsCode = db.execute(SQL)
	if RsCode.eof = false then
		CUSTNO = RsCode("CUSTNO") 
		SOSOKGB_A = RsCode("SOSOKGB_A")
		SOSOKGB_B = RsCode("SOSOKGB_B")
		SOSOKGB_C = RsCode("SOSOKGB_C")
		SOSOKGB_D = RsCode("SOSOKGB_D")
		SOSOKGB_E = RsCode("SOSOKGB_E")
		LEVEL_B = RsCode("LEVEL_B")
		LEVEL_C = RsCode("LEVEL_C")	
		LEVEL_D = RsCode("LEVEL_D")	
		CUSTNAME = RsCode("NAME")
		TELNO = RsCode("CELLPHONE")
		TELNO2 = RsCode("HOMEPHONE")
		SEXGB = RsCode("SEX")	
	else
		'------------
		'전화번호로 찾아보기
		'------------
		if len(sCID) = 7 then '군전화임.
			SQL = " select top 1 * from tb_armyinfo where aclass < 'O' and telno = '" & left(sCID,3) & "'"
			set RsCode = db.execute(SQL)
			if RsCode.eof = false then
				SOSOKGB_A = RsCode("Aclass")
				SOSOKGB_B = RsCode("Bclass")	
				SOSOKGB_C = RsCode("Cclass")	
				SOSOKGB_D = RsCode("Dclass")
				SOSOKGB_E = RsCode("Eclass")					
			end if
		elseif len(sCID) >= 9 then '일반전화임..
			SQL = " select top 1 * from tb_armyinfo where aclass < 'O' and telno2 like '%" & sCID & "%'"
			set RsCode = db.execute(SQL)
			if RsCode.eof = false then
				SOSOKGB_A = RsCode("Aclass")
				SOSOKGB_B = RsCode("Bclass")	
				SOSOKGB_C = RsCode("Cclass")	
				SOSOKGB_D = RsCode("Dclass")
				SOSOKGB_E = RsCode("Eclass")						
			end if
		end if
	end if

	if LINEKIND = "5001" or LINEKIND = "5002" then
		CHANNELGB_B = "Q01"
		CHANNELGB_C = "Q01A"
	else
		'일반전화
		CHANNELGB_B = "Q01"
		CHANNELGB_C = "Q01C"
	end if

	'-----------------------------------------------------------------------------
else

	QueryYN = request("QueryYN")
	FromDate = request("FromDate")
	ToDate = request("ToDate")
	curPage = request("curPage")
	whereCD1 = Trim(request("whereCD1")) '성별
	whereCD2 = Trim(request("whereCD2")) '상담방법
	whereCD3 = Trim(request("whereCD3")) '의뢰인
	whereCD4 = Trim(request("whereCD4")) '상담분야
	whereCD5 = Trim(request("whereCD5")) '소속
	whereCD6 = Trim(request("whereCD6")) '계급구분
	whereCD7 = Trim(request("whereCD7")) '계급구분2
	whereCD8 = Trim(request("whereCD8"))	'성명
	whereCD9 = Trim(request("whereCD9"))	'전화번호
	whereCD10 = Trim(request("whereCD10"))	'소속
	whereCD11 = Trim(request("whereCD11"))	'처리결과
	whereCD12 = Trim(request("whereCD12"))	'처리결과
	whereCD13_A = Trim(request("whereCD13_A"))	'상담분야
	whereCD13_B = Trim(request("whereCD13_B"))	'상담분야
	whereCD5_A = Trim(request("whereCD5_A")) '소속
	whereCD5_B = Trim(request("whereCD5_B")) '소속
	whereCD5_C = Trim(request("whereCD5_C")) '소속
	whereCD5_E = Trim(request("whereCD5_E")) '소속
	whereCD5_F = Trim(request("whereCD5_F")) '소속
	whereCD6_A = Trim(request("whereCD6_A")) '계급구분
	whereCD6_B = Trim(request("whereCD6_B")) '계급구분
	whereCD6_C = Trim(request("whereCD6_C")) '계급구분
	whereGB = Trim(request("whereGB"))
	'CHANNELGB = whereGB
	
	where1 = "FromDate=" & FromDate & "&ToDate=" & ToDate & "&whereCD1=" & whereCD1 & "&whereCD2=" & whereCD2 & "&whereCD3=" & whereCD3 & "&whereCD4=" & whereCD4 & "&whereCD5=" & whereCD5
	where1 = where1 & "&whereCD6=" & whereCD6 & "&whereCD7=" & whereCD7 & "&whereCD8=" & whereCD8 & "&whereCD9=" & whereCD9 & "&whereCD10=" & whereCD10 & "&whereCD11=" & whereCD11
	where1 = where1 & "&whereCD12=" & whereCD12& "&whereCD5_A=" & whereCD5_A& "&whereCD5_B=" & whereCD5_B& "&whereCD5_C=" & whereCD5_C& "&whereCD5_D=" & whereCD5_D& "&whereCD5_E=" & whereCD5_E
	where1 = where1 & "&whereCD6_A=" & whereCD6_A& "&whereCD6_B=" & whereCD6_B& "&whereCD6_C=" & whereCD6_C& "&whereCD13_A=" & whereCD13_A& "&whereCD13_B=" & whereCD13_B & "&whereGB=" & whereGB
	
	where2 = "curPage=" & curPage & "&" & where1

end if


if JUBSEQ = "" then

	guboon = "INS"
	LINEKIND = request("LINEKIND")
	TELNO = request("telNo")
	CID = request("telNo")
	CB_SEQ = request("CB_SEQ")
	if InType = "RECORD" then
	else
		sql = "select convert(varchar(19),getdate(),121)"
		set Rs = db.execute(sql)
		JUBTIME = rs(0)
	end if
	if InType = "CALL" or InType = "RECORD" then
		'IOFLAG = "2"
	else
		IOFLAG = "1"
	end if
	'if LINEKIND = "SIP-DigitalE1" then
	'	CHANNELGB = "A"
	'else
	'	CHANNELGB = "B"
	'end if

	if instr(CID,"anonymous") > 0 then
		CID = "anonymous"
	end if

else

	SQL = "	SELECT *, CONVERT(CHAR(19),JUBTIME,121) AS JUBTIME1 FROM TB_CRIMECALLHISTORY"
	SQL = SQL & "		WHERE	JUBSEQ = '" & JUBSEQ & "'"

	Set Rs = server.createObject("ADODB.Recordset")
	Rs.open SQL,db
	if rs.eof = false then
		
		JUBSEQ = rs("JUBSEQ")
		JUBDATE = rs("JUBDATE")
		JUBTIME = rs("JUBTIME1")
		IOFLAG = rs("IOFLAG")
		CUSTNO = rs("CUSTNO")
		CUSTNAME = rs("CUSTNAME")
		TELNO = rs("TELNO")
		TELNO2 = rs("TELNO2")
		CID = rs("CID")
		SEXGB = TRIM(rs("SEXGB"))
		CHANNELGB = rs("CHANNELGB")
		REQUESTERGB = rs("REQUESTERGB")
		FAMILYGB = rs("FAMILYGB")
		CONSULTGB = rs("CONSULTGB")
		CONSULTETCGB = rs("CONSULTETCGB")
		SOSOKGB_A = rs("SOSOKGB_A")
		SOSOKGB_B = rs("SOSOKGB_B")
		SOSOKGB_C = rs("SOSOKGB_C")
		SOSOKGB_D = rs("SOSOKGB_D")
		SOSOKGB_E = rs("SOSOKGB_E")
		LEVEL_B = rs("LEVEL_B")
		LEVEL_C = rs("LEVEL_C")
		LEVEL_D = rs("LEVEL_D")	
		CALLCLASS_B = rs("CALLCLASS_B")	'상담종류
		CALLCLASS_C = rs("CALLCLASS_C")
		CHANNELGB_B = rs("CHANNELGB_B")
		CHANNELGB_C = rs("CHANNELGB_C")

		CALLFLAG = rs("CALLFLAG")	
		CALLKIND_B = rs("CALLKIND_B")	'가해자
		CALLKIND_C = rs("CALLKIND_C")	'가해자
	
		QUESTION = rs("QUESTION")
		REPLY = rs("REPLY")
		REMARK = rs("REMARK")
		RESULTGB = rs("RESULTGB")
		RESERVEDATE = rs("RESERVEDATE")
		RESERVETIME = rs("RESERVETIME")
		PROCESSGB = rs("PROCESSGB")
		WEATHER = rs("WEATHER")
		CALLID = rs("CALLID")
		RECORDFILE = rs("RECORDFILE")
		INCODE = rs("INCODE")
		EMERYN = rs("EMERYN")
		CB_SEQ = rs("CB_SEQ")
		FILENAME1 = rs("FILENAME")
		REFERJUBSEQ = rs("REFERJUBSEQ")
		REFCNT =  rs("REFCNT")
		CALLTIMEDP = rs("CALLTIMEDP")

		IF CALLTIMEDP <> "" THEN
			CALLTIME1 = LEFT(CALLTIMEDP,2)
			CALLTIME2 = MID(CALLTIMEDP,4,2)
			CALLTIME3 = MID(CALLTIMEDP,7,2)
		END IF
		
		SS_LoginNAME = db_GetUserName(INCODE)

		if RECORDFILE = "" and CALLID <> "" then
			RECORDFILE = db_getRecFileName(CALLID,left(JUBTIME,10))
		end if
	
	end if


end if
'response.write SOSOKGB_A

CUSTNO1 = request("CUSTNO")
if CUSTNO1 <> "" then '고객을 선택한 케이스
	'고객번호가 있다면.. 고객번호를 넣어라
	SQL = "SELECT * FROM TB_CUSTINFO WHERE CUSTNO = '" & CUSTNO1 &"'"
	Set RSCUST = server.createObject("ADODB.Recordset")
	RSCUST.open SQL,db
	if RSCUST.eof = false then
		CUSTNO = CUSTNO1
		SOSOKGB_A = RSCUST("SOSOKGB_A")
		SOSOKGB_B = RSCUST("SOSOKGB_B")
		SOSOKGB_C = RSCUST("SOSOKGB_C")
		SOSOKGB_D = RSCUST("SOSOKGB_D")
		SOSOKGB_E = RSCUST("SOSOKGB_E")
		LEVEL_B = RSCUST("LEVEL_B")
		LEVEL_C = RSCUST("LEVEL_C")	
		LEVEL_D = RSCUST("LEVEL_D")	
		CUSTNAME = RSCUST("NAME")
		TELNO = RSCUST("CELLPHONE")
		TELNO2 = RSCUST("HOMEPHONE")
		SEXGB = RSCUST("SEX")	
	end if
end if

if JUBSEQ = "" then
	if CUSTNO = "" then
		db_TotalREFCNT = 1
	else
		sql = " select count(*) + 1 from tb_crimecallhistory where CUSTNO = '" & CUSTNO & "'"
		set Rs1 = db.execute(sql)
		db_TotalREFCNT = Rs1(0)
	end if
	REFCNT = db_TotalREFCNT
end if

CID1 = request("CID")
if CID1 <> CID THEN
	CID1 = CID
end if

IF len(FILENAME1)>0 THEN
	Filename_Temp = split(FILENAME1,".")
	FileType = FormatFile(Filename_Temp(1))
END IF

if FILENAME1 <> "" then
	FILENAME1_url = "<a href='/Upload/Lifecall/Download.asp?filename="&FILENAME1&"'>"&FILENAME1&"</a>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"&FILENAME1&" 삭제' style='cursor:hand;' align='absmiddle' onClick=FileDel('inUpFrm','"&FILENAME1&"')>&nbsp;"
end if

if REFERJUBSEQ <> "" then
	REFERJUBSEQ_URL = "<a href='##'>"&REFERJUBSEQ&"</a>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"&REFERJUBSEQ&" 삭제' style='cursor:hand;' align='absmiddle' onClick=ReferDel('inUpFrm','"&REFERJUBSEQ&"')>&nbsp;"
end if

IF RecFileName <> "" THEN
	RecFileName_URL = "<a href='##'>"&right(RecFileName,22)&"</a>&nbsp;<img src='/Images/Comm/IconAlert.gif' align='absmiddle' style='cursor:hand;' onClick=fn_Player('"&RecFileName&"'); title='녹음내용 청취'>&nbsp;</a>"	
	RecFile = RecFileName
elseif RECORDFILE <> "" THEN
	RecFileName_URL = "<a href='##'>"&right(RECORDFILE,22)&"</a>&nbsp;<img src='/Images/Comm/IconAlert.gif' align='absmiddle' style='cursor:hand;' onClick=fn_Player('"&RECORDFILE&"'); title='녹음내용 청취'>&nbsp;</a>"	
	RecFile = RECORDFILE
END IF

IF WEEKDAY(JUBTIME)=1 THEN
	JUBDAY="일"
ELSEIF WEEKDAY(JUBTIME)=2 THEN
	JUBDAY="월"
ELSEIF WEEKDAY(JUBTIME)=3 THEN
	JUBDAY="화"
ELSEIF WEEKDAY(JUBTIME)=4 THEN
	JUBDAY="수"
ELSEIF WEEKDAY(JUBTIME)=5 THEN
	JUBDAY="목"
ELSEIF WEEKDAY(JUBTIME)=6 THEN
	JUBDAY="금"
ELSEIF WEEKDAY(JUBTIME)=7 THEN
	JUBDAY="토"
END IF
%>

<table border="0" width="1200" cellpadding="0" cellspacing="3" bgcolor="#EFECE5" align="center">
	<tr bgcolor="#FFFFFF">
		<td>		
			<form name="inUpFrm"  method="post" action="/menu03/submenu0321/lifecallhistory_InsUpDel.asp" onsubmit="return fn_inup(this);" style="margin:0">
			<input type="hidden" name="FromDate" value="<%=FromDate%>">
			<input type="hidden" name="ToDate" value="<%=ToDate%>">
			<input type="hidden" name="curPage" value="<%=curPage%>">
			<input type="hidden" name="QueryYN" value="<%=QueryYN%>">
			<input type="hidden" name="whereCD1" value="<%=whereCD1%>">
			<input type="hidden" name="whereCD2" value="<%=whereCD2%>">
			<input type="hidden" name="whereCD3" value="<%=whereCD3%>">
			<input type="hidden" name="whereCD4" value="<%=whereCD4%>">
			<input type="hidden" name="whereCD5" value="<%=whereCD5%>">
			<input type="hidden" name="whereCD6" value="<%=whereCD6%>">
			<input type="hidden" name="whereCD7" value="<%=whereCD7%>">
			<input type="hidden" name="whereCD8" value="<%=whereCD8%>">
			<input type="hidden" name="whereCD9" value="<%=whereCD9%>">
			<input type="hidden" name="whereCD10" value="<%=whereCD10%>">
			<input type="hidden" name="whereCD11" value="<%=whereCD11%>">
			<input type="hidden" name="whereCD12" value="<%=whereCD12%>">
			<input type="hidden" name="whereCD13_A" value="<%=whereCD13_A%>">
			<input type="hidden" name="whereCD13_B" value="<%=whereCD13_B%>">
			<input type="hidden" name="JUBSEQ" value="<%=JUBSEQ%>">
			<input type="hidden" name="guboon" value="<%=guboon%>">	
			<input type="hidden" name="LEVEL_B" value="<%=LEVEL_B%>">	
			<input type="hidden" name="LEVEL_C" value="<%=LEVEL_C%>">	
			<input type="hidden" name="LEVEL_D" value="<%=LEVEL_D%>">	
			<input type="hidden" name="SOSOKGB_A" value="<%=SOSOKGB_A%>">	
			<input type="hidden" name="SOSOKGB_B" value="<%=SOSOKGB_B%>">	
			<input type="hidden" name="SOSOKGB_C" value="<%=SOSOKGB_C%>">	
			<input type="hidden" name="SOSOKGB_D" value="<%=SOSOKGB_D%>">	
			<input type="hidden" name="SOSOKGB_E" value="<%=SOSOKGB_E%>">	
			<input type="hidden" name="CHANNELGB_B" value="<%=CHANNELGB_B%>">	
			<input type="hidden" name="CHANNELGB_C" value="<%=CHANNELGB_C%>">
			<input type="hidden" name="whereGB" value="<%=whereGB%>">

			<input type="hidden" name="CALLCLASS_B" value="<%=CALLCLASS_B%>">	
			<input type="hidden" name="CALLCLASS_C" value="<%=CALLCLASS_C%>">	
			


			<!--<input type="hidden" name="CALLKIND_B" value="<%=CALLKIND_B%>">	
			<input type="hidden" name="CALLKIND_C" value="<%=CALLKIND_C%>">-->
			<input type="hidden" name="CALLCLASS_B_2" value="<%=CALLCLASS_B_2%>">	
			<input type="hidden" name="CALLCLASS_C_2" value="<%=CALLCLASS_C_2%>">	
			<input type="hidden" name="CALLCLASS_B_3" value="<%=CALLCLASS_B_3%>">	
			<input type="hidden" name="CALLCLASS_C_3" value="<%=CALLCLASS_C_3%>">	
			<input type="hidden" name="CALLCLASS_B_4" value="<%=CALLCLASS_B_4%>">	
			<input type="hidden" name="CALLCLASS_C_4" value="<%=CALLCLASS_C_4%>">	
			<input type="hidden" name="CALLCLASS_B_5" value="<%=CALLCLASS_B_5%>">	
			<input type="hidden" name="CALLCLASS_C_5" value="<%=CALLCLASS_C_5%>">	
			<!--<input type="hidden" name="CALLKIND_B_2" value="<%=CALLKIND_B_2%>">	
			<input type="hidden" name="CALLKIND_C_2" value="<%=CALLKIND_C_2%>">-->


			<input type="hidden" name="CONSULTETCGB" value="<%=CONSULTETCGB%>">
			<input type="hidden" name="CB_SEQ" value="<%=CB_SEQ%>">
			<input type="hidden" name="div_2" value="<%=div_2%>">
			<input type="hidden" name="div_3" value="<%=div_3%>">
			<input type="hidden" name="div_4" value="<%=div_4%>">
			<input type="hidden" name="div_5" value="<%=div_5%>">
			<input type="hidden" name="div_6" value="<%=div_6%>">


			<table width="100%" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#ffffff">
			    <tr height="30">
					<td align="left" bgcolor="#FFFFFF" class="TDCont" colspan="8">&nbsp;<img src="/Images/dot_01.gif" >&nbsp;<b><font color="#ff00ff"></font>상담내용 - <%=SS_LoginNAME%></b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					<!--<% if sGotoURL <> "" then %></td><td align='right'><a href="##" onclick="javascript:fnGoto3012('<%=sGotoURL%>');">사후상담으로 이동</a></td><%else%></td><%end if%>-->
				</tr>
			</table>
			<table width="100%" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">
			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=150 align="center" colspan='2'>상담일시</td>
					<td bgcolor="#FFFFFF" width=230 nowrap><input type="text" name="JUBTIME" value="<%=JUBTIME%>" maxlength="19" size="20" onfocus="setFocusColor(this);" onblur="setOutColor(this);" align="absmiddle">&nbsp;<input type="text" name="JUBDAY" value="<%=JUBDAY%>" maxlength="2" size="2" onfocus="setFocusColor(this);" onblur="setOutColor(this);" align="absmiddle" readonly>&nbsp;<font color="#ff0000"><b><input type="checkbox" name="EMERYN" value="Y" class="none" <% if EMERYN="Y" then Response.Write("checked") end if %>>긴급</b></font>

					</td>

					<td bgcolor="#EEF6FF" class="TDCont" width=150 align="center" colspan='1'>통화시간</td>
					<td bgcolor="#FFFFFF" width=200><input type="text" name="CALLTIME1" value="<%=CALLTIME1%>" maxlength="2" size="2" style="border-width:1px ; border-style:solid; text-align:right" onfocus="setFocusColor(this);" onblur="setOutColor(this);" align="absmiddle">&nbsp;:&nbsp;<input type="text" name="CALLTIME2" value="<%=CALLTIME2%>" maxlength="2" size="2" style="border-width:1px ; border-style:solid; text-align:right"  onfocus="setFocusColor(this);" onblur="setOutColor(this);" align="absmiddle">&nbsp;:&nbsp;<input type="text" name="CALLTIME3" value="<%=CALLTIME3%>" style="border-width:1px ; border-style:solid; text-align:right" maxlength="2" size="2" onfocus="setFocusColor(this);" onblur="setOutColor(this);" align="absmiddle">&nbsp;(시:분:초)
					</td>

					<td bgcolor="#EEF6FF" class="TDCont" width=150 align="center" colspan='1'>인/아웃구분</td>
					<td bgcolor="#FFFFFF" width=200 nowrap><input type="radio" name="IOFLAG" value="1" class="none" <% if IOFLAG = "1" or IOFLAG = "" then %>checked<%end if%> >인
						<input type="radio" name="IOFLAG" value="2" class="none" <% if IOFLAG = "2" then %>checked<%end if%>>아웃
					
					</td>
				</tr>
				<tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=150 align="center" colspan='2'>수신번호</td>
					<td bgcolor="#FFFFFF" width=300><input type="text" name="CID" value="<%=CID%>" maxlength="16" size="16" onfocus="setFocusColor(this);" onblur="setOutColor(this);" align="absmiddle">&nbsp;<img src="/Images/Cti/icon_tel.gif" style="cursor:hand;" onClick="fn_dial('2','0');" align="absmiddle" title="전화걸기">&nbsp<img src="/Images/Cti/icon_tel_1.gif" style="cursor:hand;" onClick="fn_dial_1('2','0');" align="absmiddle" title="일반전화로 전화걸기"><!--&nbsp<img src="/Images/Cti/icon_sche.gif" style="cursor:hand;" onClick="fn_sms('2','2');" align="absmiddle" title="문자전송">-->
					
					</td>	

					<td bgcolor="#EEF6FF" class="TDCont" width=150 align="center">연락처1</td>
					<td bgcolor="#FFFFFF" width=300><input type="text" name="TELNO" value="<%=TELNO%>" maxlength="15" size="15" onfocus="setFocusColor(this);" onblur="setOutColor(this);" align="absmiddle">&nbsp<img src="/Images/Cti/icon_tel.gif" style="cursor:hand;" onClick="fn_dial('2','1');" align="absmiddle" title="군전화로 전화걸기">&nbsp<img src="/Images/Cti/icon_tel_1.gif" style="cursor:hand;" onClick="fn_dial_1('2','1');" align="absmiddle" title="일반전화로 전화걸기"><!--&nbsp<img src="/Images/Cti/icon_sche.gif" style="cursor:hand;" onClick="fn_sms('2','1');" align="absmiddle" title="문자전송">--></td>

					<td bgcolor="#EEF6FF" class="TDCont" width=150 align="center">연락처2</td>
					<td bgcolor="#FFFFFF" width=300><input type="text" name="TELNO2" value="<%=TELNO2%>" maxlength="15" size="15" onfocus="setFocusColor(this);" onblur="setOutColor(this);" align="absmiddle">&nbsp;<img src="/Images/Cti/icon_tel.gif" style="cursor:hand;" onClick="fn_dial('2','2');" align="absmiddle" title="군전화로 전화걸기">&nbsp<img src="/Images/Cti/icon_tel_1.gif" style="cursor:hand;" onClick="fn_dial_1('2','2');" align="absmiddle" title="일반전화로 전화걸기"><!--&nbsp<img src="/Images/Cti/icon_sche.gif" style="cursor:hand;" onClick="fn_sms('2','2');" align="absmiddle" title="문자전송">--></td>

				</tr>
				<tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=150 align="center" colspan='2'>성    명</td>
					<td bgcolor="#FFFFFF" ><input type="text" name="CUSTNAME" value="<%=CUSTNAME%>" maxlength="30" size="30" onfocus="setFocusColor(this);" onblur="setOutColor(this);" align="absmiddle" onKeypress="if (event.keyCode==13) {fn_CustSearch();}"><input type="hidden" name="CUSTNO" value="<%=CUSTNO%>" maxlength="30" size="30" onfocus="setFocusColor(this);" onblur="setOutColor(this);" align="absmiddle">
					</td>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center">성    별</td>
					<td bgcolor="#FFFFFF"><input type="radio" name="SEXGB" value="1" class="none" <% if SEXGB = "1" or SEXGB = "" then %>checked<%end if%> >남
						<input type="radio" name="SEXGB" value="2" class="none" <% if SEXGB = "2" then %>checked<%end if%>>녀
					</td>
					<td bgcolor="#EEF6FF" class="TDCont" width=200 align="center">상담관배치여부</td>
					<td bgcolor="#FFFFFF" width=100 nowrap><input type="text" name="CounselorYN" value="<%=db_getCateNameCounselorYN_(SOSOKGB_A,SOSOKGB_B,SOSOKGB_C,SOSOKGB_D,SOSOKGB_E)%>" style="border-width:0px ; border-color:#EEF6FF ; border-style:solid; text-align:left; font-color:#ff0000;font-weight:bold" readonly>
					</td>
				</tr>

			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=150 align="center"  colspan='2'>관련상담</td>
					<td bgcolor="#FFFFFF" width=200><input type="hidden" name="REFERJUBSEQ" value="<%=REFERJUBSEQ%>" readonly>
								<span id="txtREFERJUBSEQ"><%=REFERJUBSEQ_url%></span><img src="/Images/Btn/BtnRefSrc.GIF" style="cursor:hand;" class="None" align="absmiddle" onClick="ReferUp('A','1');">

					<td bgcolor="#EEF6FF" class="TDCont" width=150 align="center">상담회차</td>
					<td bgcolor="#FFFFFF" width=200><input type="text" name="REFCNT" value="<%=REFCNT%>" readonly size='2' style="border-width:1px ; border-style:solid; text-align:right; text-forecolor=#0000ff">&nbsp;회
					</td>	

					<td bgcolor="#EEF6FF" class="TDCont" width=150 align="center">녹취파일</td>
					<td bgcolor="#FFFFFF" width=200><input type="hidden" name="CALLID" value="<%=CALLID%>"><input type="hidden" name="RECFILE" value="<%=RECFILE%>"><span id="txtRECFILE"><%=RecFileName_URL%></span>					
					</td>	

				</tr>
			    <tr>

					<td bgcolor="#FDE6F3" class="TDCont" width=150 align="center" colspan='2'>군분류</td>
					<td bgcolor="#FFFFFF" colspan='5'><input type="radio" name="CHANNELGB" value="130331" class="none" <% if CHANNELGB = "130331" or CHANNELGB = "" then %>checked<%end if%> >조사본부
						<input type="radio" name="CHANNELGB" value="130332" class="none" <% if CHANNELGB = "130332" then %>checked<%end if%> >육군
						<input type="radio" name="CHANNELGB" value="130333" class="none" <% if CHANNELGB = "130333" then %>checked<%end if%> >해군
						<input type="radio" name="CHANNELGB" value="130334" class="none" <% if CHANNELGB = "130334" then %>checked<%end if%> >공군
						<input type="radio" name="CHANNELGB" value="130335" class="none" <% if CHANNELGB = "130335" then %>checked<%end if%> >해병대
					</td>	

				</tr>
				<tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center" rowspan='2'>상담유형</td>
					<td bgcolor="#EEF6FF" class="TDCont" align="center" width=50>1차</td>
					<td bgcolor="#FFFFFF" nowrap colspan='5'><iframe src="/menu03/submenu0321/frame_CLASS.asp?frame_nm=frame_channelgb_B&CLASSNM=CHANNELGB&CLASSGB=B&ACLASS=Q&BCLASS=<%=CHANNELGB_B%>" scrolling="no" frameborder="0" width=100% height=25 name="frame_channelgb_B"></iframe>		
					</td>
				</tr>
				<tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=50 align="center">2차</td>
					<td bgcolor="#FFFFFF" nowrap colspan='5'><iframe src="/menu03/submenu0321/frame_CLASS.asp?frame_nm=frame_channelgb_C&CLASSNM=CHANNELGB&CLASSGB=C&ACLASS=Q&BCLASS=<%=CHANNELGB_B%>&CCLASS=<%=CHANNELGB_C%>" scrolling="no" frameborder="0" width=100% height=25 name="frame_channelgb_C"></iframe>	
					</td>
				</tr>

				<tr id="divSOSOKGB_A" style="display:block;">
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center" rowspan='5'>소    속<br><img src="/Images/Btn/BtnSearch.gif" style="cursor:hand;" align="absmiddle" onClick="fn_PopCatalog();" title="소속검색"></td>
					<td bgcolor="#EEF6FF" class="TDCont" align="center" width=50>1차</td>
					<td bgcolor="#FFFFFF" nowrap colspan='5'><iframe src="/menu03/submenu0321/frame_CLASS.asp?frame_nm=frame_sosok_A&CLASSNM=SOSOK&CLASSGB=A&ACLASS=<%=SOSOKGB_A%>" scrolling="no" frameborder="0" width=100% height=25 name="frame_sosok_A"></iframe>		

					</td>
				</tr>
				<tr id="divSOSOKGB_B" style="display:block;">
					<td bgcolor="#EEF6FF" class="TDCont" width=50 align="center">2차</td>
					<td bgcolor="#FFFFFF" nowrap colspan='5'><iframe src="/menu03/submenu0321/frame_CLASS.asp?frame_nm=frame_sosok_B&CLASSNM=SOSOK&CLASSGB=B&ACLASS=<%=SOSOKGB_A%>&BCLASS=<%=SOSOKGB_B%>" scrolling="no" frameborder="0" width=100% height=50 name="frame_sosok_B"></iframe>	
					</td>
				</tr>
				<tr id="divSOSOKGB_C" style="display:block;">
					<td bgcolor="#EEF6FF" class="TDCont" width=50 align="center">3차</td>
					<td bgcolor="#FFFFFF" nowrap colspan='5'><iframe src="/menu03/submenu0321/frame_CLASS.asp?frame_nm=frame_sosok_C&CLASSNM=SOSOK&CLASSGB=C&ACLASS=<%=SOSOKGB_A%>&BCLASS=<%=SOSOKGB_B%>&CCLASS=<%=SOSOKGB_C%>" scrolling="no" frameborder="0" width=100% height=50 name="frame_sosok_C"></iframe>	
					</td>
				</tr>
				<tr id="divSOSOKGB_D" style="display:block;">
					<td bgcolor="#EEF6FF" class="TDCont" width=50 align="center">4차</td>
					<td bgcolor="#FFFFFF" nowrap colspan='5'><iframe src="/menu03/submenu0321/frame_CLASS.asp?frame_nm=frame_sosok_D&CLASSNM=SOSOK&CLASSGB=D&ACLASS=<%=SOSOKGB_A%>&BCLASS=<%=SOSOKGB_B%>&CCLASS=<%=SOSOKGB_C%>&DCLASS=<%=SOSOKGB_D%>" scrolling="no" frameborder="0" width=100% height=50 name="frame_sosok_D"></iframe>	
					</td>
				</tr>
				<tr id="divSOSOKGB_E" style="display:block;">
					<td bgcolor="#EEF6FF" class="TDCont" width=50 align="center">5차</td>
					<td bgcolor="#FFFFFF" nowrap colspan='5'><iframe src="/menu03/submenu0321/frame_CLASS.asp?frame_nm=frame_sosok_E&CLASSNM=SOSOK&CLASSGB=E&ACLASS=<%=SOSOKGB_A%>&BCLASS=<%=SOSOKGB_B%>&CCLASS=<%=SOSOKGB_C%>&DCLASS=<%=SOSOKGB_D%>&ECLASS=<%=SOSOKGB_E%>" scrolling="no" frameborder="0" width=100% height=25 name="frame_sosok_E"></iframe>	
					</td>
				</tr>

				<tr id="divLEVEL_B" style="display:block;">
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center" rowspan='3'>계    급</td>
					<td bgcolor="#EEF6FF" class="TDCont" align="center" width=50>1차&nbsp;</td>
					<td bgcolor="#FFFFFF" nowrap colspan='5'><iframe src="/menu03/submenu0321/frame_CLASS.asp?frame_nm=frame_level_B&CLASSNM=LEVEL&CLASSGB=B&ACLASS=P&BCLASS=<%=LEVEL_B%>" scrolling="no" frameborder="0" width=100% height=25 name="frame_level_B"></iframe>		
					</td>
				</tr>
				<tr id="divLEVEL_C" style="display:block;">
					<td bgcolor="#EEF6FF" class="TDCont" width=50 align="center">2차</td>
					<td bgcolor="#FFFFFF" nowrap colspan='5'><iframe src="/menu03/submenu0321/frame_CLASS.asp?frame_nm=frame_level_C&CLASSNM=LEVEL&CLASSGB=C&ACLASS=P&BCLASS=<%=LEVEL_B%>&CCLASS=<%=LEVEL_C%>" scrolling="no" frameborder="0" width=100% height=25 name="frame_level_C"></iframe>	
					</td>
				</tr>
				<tr id="divLEVEL_D" style="display:block;">
					<td bgcolor="#EEF6FF" class="TDCont" width=50 align="center">3차</td>
					<td bgcolor="#FFFFFF" nowrap colspan='5'><iframe src="/menu03/submenu0321/frame_CLASS.asp?frame_nm=frame_level_D&CLASSNM=LEVEL&CLASSGB=D&ACLASS=P&BCLASS=<%=LEVEL_B%>&CCLASS=<%=LEVEL_C%>&DCLASS=<%=LEVEL_D%>" scrolling="no" frameborder="0" width=100% height=25 name="frame_level_D"></iframe>	
					</td>
				</tr>

				<tr id="divCALLFLAG" style="display:block;">
					<td bgcolor="#EEF6FF" class="TDCont" width=150 align="center" colspan='2'>인지경로</td>
					<td bgcolor="#FFFFFF" height=20 nowrap colspan='5'>
<%
						sReplyHtml = ""
						j = 0
							'======= 처리구분 코드 가져오기 ==================================================
							SqlCode = "SELECT CODE, CODENAME FROM TB_CODE"
							SqlCode = SqlCode& " WHERE USEYN='Y' AND CODEGROUP='C10'"
							SqlCode = SqlCode& " ORDER BY CODE"
							set RsCode = db.execute(SqlCode)

							do until RsCode.eof
								j = j + 1
								if CALLFLAG = RsCode("CODE") then
									SelectedValue = "checked"
								else
									SelectedValue = ""
								end if
								if j = 1 then
									sReplyHtml = "<input type='RADIO' value='" & RsCode("CODE") & "' name='CALLFLAG' class='none' " & SelectedValue & " >" & RsCode("CODENAME")	
								else
									sReplyHtml = sReplyHtml & "&nbsp;<input type='RADIO' value='" & RsCode("CODE") & "' name='CALLFLAG' class='none' " & SelectedValue & ">" & RsCode("CODENAME")	
								end if
								RsCode.movenext
							loop
							RsCode.close
							response.write sReplyHtml

							if sReplyHtml <> "" then

								sReplyHtml = "&nbsp;<img src='/Images/Comm/IconDel2.gif' title='선택취소' style='cursor:hand;' align='absmiddle' onclick=""javascript:fn_DEL('CALLFLAG');"">"
								response.write sReplyHtml
								
							end if

						%>	
					</td>					
				</tr>



				<tr id="divCALLCLASS_B_1" style="display:block;">
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center" rowspan='2'>상담분야</td>
					<td bgcolor="#EEF6FF" class="TDCont" align="center" width=50>1차</td>
					<td bgcolor="#FFFFFF" nowrap colspan='5'><iframe src="/menu03/submenu0321/frame_CLASS.asp?frame_nm=frame_callclass_B&CLASSNM=CALLCLASS&CLASSGB=B&ACLASS=S&BCLASS=<%=CALLCLASS_B%>" scrolling="no" frameborder="0" width=100% height=25 name="frame_callclass_B"></iframe>		
					</td>
				</tr>
				<tr id="divCALLCLASS_C_1" style="display:block;">
					<td bgcolor="#EEF6FF" class="TDCont" width=50 align="center">2차</td>
					<td bgcolor="#FFFFFF" nowrap colspan='5'><iframe src="/menu03/submenu0321/frame_CLASS.asp?frame_nm=frame_callclass_C&CLASSNM=CALLCLASS&CLASSGB=C&ACLASS=O&BCLASS=<%=CALLCLASS_B%>&CCLASS=<%=CALLCLASS_C%>" scrolling="no" frameborder="0" width=100% height=25 name="frame_callclass_C"></iframe>	
					</td>
				</tr>


				<tr id="divCALLKIND_1" style="display:block;">
					<td bgcolor="#EEF6FF" class="TDCont" width=150 align="center" colspan='2'>원인제공자</td>
					<td bgcolor="#FFFFFF" nowrap colspan='5'>


<%
						sReplyHtml = ""
						j = 0
							'======= 처리구분 코드 가져오기 ==================================================
							SqlCode = "SELECT bclass CODE, classname CODENAME FROM tb_armyinfo"
							SqlCode = SqlCode& " WHERE aclass = 'T' and Bclass is not null and cclass is null"
							SqlCode = SqlCode& " ORDER BY bclass"
							set RsCode = db.execute(SqlCode)

							do until RsCode.eof
								j = j + 1
								if CALLKIND_B = RsCode("CODE") then
									SelectedValue = "checked"
								else
									SelectedValue = ""
								end if
								if j = 1 then
									sReplyHtml = "<input type='RADIO' value='" & RsCode("CODE") & "' name='CALLKIND_B' class='none' " & SelectedValue & " >" & RsCode("CODENAME")	
								else
									sReplyHtml = sReplyHtml & "&nbsp;<input type='RADIO' value='" & RsCode("CODE") & "' name='CALLKIND_B' class='none' " & SelectedValue & ">" & RsCode("CODENAME")	
								end if
								RsCode.movenext
							loop
							RsCode.close
							response.write sReplyHtml

							if sReplyHtml <> "" then

								sReplyHtml = "&nbsp;<img src='/Images/Comm/IconDel2.gif' title='선택취소' style='cursor:hand;' align='absmiddle' onclick=""javascript:fn_DEL('CALLKIND_B');"">"
								response.write sReplyHtml
								
							end if

						%>	
					</td>
				</tr>
				<!--<tr id="divCALLKIND_2" style="display:block;">
					<td bgcolor="#EEF6FF" class="TDCont" width=50 align="center">2차</td>
					<td bgcolor="#FFFFFF" nowrap colspan='5'><iframe src="/menu03/submenu0301/frame_CLASS.asp?frame_nm=frame_callkind_C&CLASSNM=CALLKIND&CLASSGB=C&ACLASS=R&BCLASS=<%=CALLKIND_B%>&CCLASS=<%=CALLKIND_C%>" scrolling="no" frameborder="0" width=100% height=25 name="frame_callkind_C"></iframe>	
					</td>
				</tr>-->




				<tr id="divREQUESTERGB" style="display:block;">
					<td bgcolor="#EEF6FF" class="TDCont" width=150 align="center" colspan='2'>의뢰인</td>
					<td bgcolor="#FFFFFF" height=20 nowrap colspan='5'>
<%
						sReplyHtml = ""
						j = 0
							'======= 처리구분 코드 가져오기 ==================================================
							SqlCode = "SELECT CODE, CODENAME FROM TB_CODE"
							SqlCode = SqlCode& " WHERE USEYN='Y' AND CODEGROUP='C02'"
							SqlCode = SqlCode& " ORDER BY CODE"
							set RsCode = db.execute(SqlCode)

							do until RsCode.eof
								j = j + 1
								if REQUESTERGB = RsCode("CODE") then
									SelectedValue = "checked"
								else
									SelectedValue = ""
								end if
								if j = 1 then
									sReplyHtml = "<input type='RADIO' value='" & RsCode("CODE") & "' name='REQUESTERGB' class='none' " & SelectedValue & " >" & RsCode("CODENAME")	
								else
									sReplyHtml = sReplyHtml & "&nbsp;<input type='RADIO' value='" & RsCode("CODE") & "' name='REQUESTERGB' class='none' " & SelectedValue & ">" & RsCode("CODENAME")	
								end if
								RsCode.movenext
							loop
							RsCode.close
							response.write sReplyHtml


							if sReplyHtml <> "" then

								sReplyHtml = "&nbsp;<img src='/Images/Comm/IconDel2.gif' title='선택취소' style='cursor:hand;' align='absmiddle' onclick=""javascript:fn_DEL('REQUESTERGB');"">"
								response.write sReplyHtml
								
							end if
						%>	
					</td>					
				</tr>



				<tr id="divPROCESSGB" style="display:block;">
					<td bgcolor="#EEF6FF" class="TDCont" width=150 align="center" colspan='2'>조치결과</td>
					<td bgcolor="#FFFFFF" height=20 nowrap colspan='5'>
<%
						sReplyHtml = ""
						j = 0
							'======= 처리구분 코드 가져오기 ==================================================
							SqlCode = "SELECT CODE, CODENAME FROM TB_CODE"
							SqlCode = SqlCode& " WHERE USEYN='Y' AND CODEGROUP='C21'"
							SqlCode = SqlCode& " ORDER BY CODE"
							set RsCode = db.execute(SqlCode)

							do until RsCode.eof
								j = j + 1
								if PROCESSGB = RsCode("CODE") then
									SelectedValue = "checked"
								else
									SelectedValue = ""
								end if
								if j = 1 then
									sReplyHtml = "<input type='RADIO' value='" & RsCode("CODE") & "' name='PROCESSGB' class='none' " & SelectedValue & " >" & RsCode("CODENAME")	
								else
									sReplyHtml = sReplyHtml & "&nbsp;<input type='RADIO' value='" & RsCode("CODE") & "' name='PROCESSGB' class='none' " & SelectedValue & ">" & RsCode("CODENAME")	
								end if
								RsCode.movenext
							loop
							RsCode.close
							response.write sReplyHtml


							if sReplyHtml <> "" then

								sReplyHtml = "&nbsp;<img src='/Images/Comm/IconDel2.gif' title='선택취소' style='cursor:hand;' align='absmiddle' onclick=""javascript:fn_DEL('PROCESSGB');"">"
								response.write sReplyHtml
								
							end if
						%>	
					</td>					
				</tr>


			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=150 align="center"  colspan='2'>상담내용</td>
					<td bgcolor="#FFFFFF" colspan=5 width=950><textarea name="QUESTION" style="width:100%; height:100" wrap="soft" class="TextareaInput"><%=QUESTION%></textarea>			
					</td>
				</tr>
			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=150 align="center"  colspan='2'>조치내용</td>
					<td bgcolor="#FFFFFF" colspan=5 width=950>	<textarea name="REPLY" style="width:100%; height:100" wrap="soft" class="TextareaInput"><%=REPLY%></textarea>			
					</td>
				</tr>
			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=150 align="center"  colspan='2'>특이사항</td>
					<td bgcolor="#FFFFFF" colspan=5 width=950>	<textarea name="REMARK" style="width:100%; height:50" wrap="soft" class="TextareaInput"><%=REMARK%></textarea>			
					</td>
				</tr>


			    <tr>

					<td bgcolor="#EEF6FF" class="TDCont" width=150 align="center" colspan='2'>첨부파일</td>
					<td bgcolor="#FFFFFF" nowrap colspan='5'><input type="hidden" name="FILENAME1" value="<%=FILENAME1%>" readonly>
						<span id="txtFILENAME1"><%=FILENAME1_url%></span><img src="/Images/Btn/BtnUpload.gif" style="cursor:hand;" align="absmiddle" onClick="FielUp('A','1');">
					</td>
				</tr>


			</table>
			</form>
		</td>
	</tr>
</table>
<table width="1200" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
<table width="1200" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td align='left'><img src="/Images/Btn/BtnList.gif" style="cursor:hand;" align="absmiddle" onClick="fn_list();"></td><td align="right"><img src="/Images/Btn/BtnASRegi.gif" style="cursor:hand;" class="None" align="absmiddle" onClick="fn_inup();">&nbsp;<%if SS_Login_Secgroup <>"A" and JUBSEQ <> "" then%><img src="/Images/Btn/BtnDel.gif" style="cursor:hand;" class="None" align="absmiddle" onClick="fn_listdel('<%=JUBSEQ%>');"><%end if%></td></tr></table>

<!--
<table width="940" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
<table border="0" width="940" cellspacing="1" cellpadding="0" width="100%" bgcolor="#CCCCCC" align="center">
	<tr height="25" bgcolor="#F3F3F3" align="center">
		<td><b>순번</b></td>
		<td><b>상담일시</b></td>
		<td><b>상담방법</b></td>
		<td><b>소속</b></td>
		<td><b>계급</b></td>
		<td><b>성명</b></td>
		<td><b>상담관</b></td>
		<td><b>성별</b></td>
		<td ><b>상담분야</b></td>
		<td><b>관리</b></td>
	</tr>
	<tr height="25" bgcolor="#ffffff" align="center">
		<td align="center" colspan=10>기존 상담이력이 존재하지 않습니다</td>
	</tr>

</table>
-->
<table width="1200" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
<table border="0" width="1200" cellspacing="0" cellpadding="0" width="100%" bgcolor="#CCCCCC" align="center">
	<tr height="25" bgcolor="#F3F3F3" align="center"><td>
<iframe SRC="lifecallhistory_list.asp?CUSTNO=<%=CUSTNO%>" scrolling="yes" frameborder="0" border="0" width="1200" height="200" name="IframeHistory"></iframe>
</td>
</tr>
</table>
<%'======= 화일삭제/매뉴얼등록 =======================================================================================%>
<DIV id="hiddenIframe" style="display:none;">
	<iframe SRC="about:blank" scrolling="auto" frameborder="0" border="0" width="100%" height="50" name="hiddenIframe"></iframe>
</DIV>

<script>

	function fn_AddForm(ty,f,ck){
		if(ty=="ON"){
			eval("divCALLCLASS_B_"+f).style.display = "block";
			eval("divCALLCLASS_C_"+f).style.display = "block";
			//eval("ListForm."+ck).value = "ON";
			eval("inUpFrm.div_"+f).value = "ON";
		}
		else {
			eval("divCALLCLASS_B_"+f).style.display = "none";
			eval("divCALLCLASS_C_"+f).style.display = "none";
			//eval("ListForm."+ck).value = "ON";
			eval("inUpFrm.div_"+f).value = "";
		}
	}

	function fn_sms(arg0,arg1) {

				if ( arg1 == '1' )
					sCellPhone = eval("inUpFrm.TELNO").value;
				else if ( arg1 == '2' )
					sCellPhone = eval("inUpFrm.TELNO2").value;
				else if ( arg1 == '3' )
					sCellPhone = eval("inUpFrm.CID").value;
				//sms = //window.open("/menu05/submenu0502/sms.asp?cellphone="+sCellPhone,"sms","toolbar=no,status=yes,location=no,width=620,height=500,top=0,left=0,scrollbars=yes,resizable=no");
				//sms.focus();


		ShowPOPLayer("/menu05/submenu0502/sms.asp?cellphone="+sCellPhone,'820','430');		
//				sms = window.open("sms.asp","sms","toolbar=no,status=yes,location=no,width=620,height=500,top=0,left=0,scrollbars=yes,resizable=no");
			//	sms.focus();

	}

	function fn_dial(arg0,arg1)
	{
		//전화걸기
		document.all.IOFLAG(1).checked = true;
		if ( arg1 == '0' )
			top.CallStateFrame.document.all.txtCID.value = eval("inUpFrm.CID").value;
		else if ( arg1 == '1' )
			top.CallStateFrame.document.all.txtCID.value = eval("inUpFrm.TELNO").value;
		else
			top.CallStateFrame.document.all.txtCID.value = eval("inUpFrm.TELNO2").value;

		if ( top.CallStateFrame.document.all.txtCID.value == "" )
			alert('전화걸기 실패 : 전화번호가 입력되지 않음');
		else
			top.CallStateFrame.vfn_MakeCall(top.CallStateFrame.document.all.txtCID.value,'');
	}

	function fn_dial_1(arg0,arg1)
	{
		//전화걸기

		document.all.IOFLAG(1).checked = true;

		if ( arg1 == '0' )
			top.CallStateFrame.document.all.txtCID.value = eval("inUpFrm.CID").value;
		else if ( arg1 == '1' )
			top.CallStateFrame.document.all.txtCID.value = eval("inUpFrm.TELNO").value;
		else
			top.CallStateFrame.document.all.txtCID.value = eval("inUpFrm.TELNO2").value;

		if ( top.CallStateFrame.document.all.txtCID.value == "" )
			alert('전화걸기 실패 : 전화번호가 입력되지 않음');
		else
			top.CallStateFrame.vfn_MakeCall(top.CallStateFrame.document.all.txtCID.value,'');

	}

	function fn_CustSearch(){
		//같은 고객이 있는지를 찾는다.
		//이름,cid,전화번호1,2
		ShowPOPLayer("/Include/PopUp/MemSearch.asp?FRM=crime&JUBSEQ=<%=JUBSEQ%>&CB_SEQ=<%=CB_SEQ%>&SENDPHONE="+eval("inUpFrm.CID").value+"&NAME="+eval("inUpFrm.CUSTNAME").value,'800','430');		
	}

	function fn_list(){location.href="/menu03/submenu0321/lifecallhistory.asp?<%=where2%>";}

	function fn_SetLevel2(arg)
	{
		frame_level.location = "/menu03/submenu0321/frame_level.asp?level="+arg+"&level2=";
	}
	function fn_SetSosok2(arg)
	{
		frame_sosok.location = "/menu03/submenu0321/frame_sosok.asp?SOSOKGB="+arg+"&SOSOKETCGB=";
		frame_sosok2.location = "/menu03/submenu0321/frame_sosok_3.asp?SOSOKGB="+arg+"&SOSOKETCGB=";
		document.all.CounselorYN.value = "";
	}
	function fn_SetSosok3(arg,arg1)
	{
		frame_sosok2.location = "/menu03/submenu0321/frame_sosok_3.asp?SOSOKGB="+document.all.SOSOKGB.value+"&SOSOKETCGB=";
		document.all.CounselorYN.value = "";
	}
	function fn_SetConsult2()
	{
		frame_consult.location = "/menu03/submenu0321/frame_consult.asp?CONSULTGB="+document.all.CONSULTGB.value+"&CONSULTETCGB=";
	}
	function fn_inup()
	{


		if ( inUpFrm.JUBTIME.value.length != 19 )
		{
			alert('상담일시를 정확히 입력하십시오!(형식:yyyy-mm-dd hh:nn:ss)');
			inUpFrm.JUBTIME.focus();
			return false;
		}
/*		
		if ( inUpFrm.CHANNELGB.value == '' )
		{
			alert('상담방법을 선택하십시오!');
			inUpFrm.CHANNELGB.focus();
			return false;
		}
		if ( inUpFrm.ACLASS.value == '' )
		{
			alert('상담종류를 선택하십시오!');
			inUpFrm.ACLASS.focus();
			return false;
		}
		if ( inUpFrm.CUSTNAME.value == '' )
		{
			alert('성명을 입력하십시오! 침묵 또는 익명을 원할 경우 [미상]으로 입력하십시오.');
			inUpFrm.CUSTNAME.focus();
			return false;
		}
		//if ( inUpFrm.ACLASS.value != 'C' )
		//{
			//모든 필수항목이 빠진다.
			if ( inUpFrm.SOSOKGB.value == '' )
			{
				alert('소속을 선택하십시오!');
				inUpFrm.SOSOKGB.focus();
				return false;
			}
			//육직(D), 기타(H) 2차소속 입력
			if ( inUpFrm.SOSOKGB.value == 'D' || inUpFrm.SOSOKGB.value == 'H' )
			{
				if ( inUpFrm.SOSOKETCGB.value == '' )
				{
					alert('[육직, 기타]의 소속 2차분류을 선택하십시오!');
					//
					return false;
				}
			}

		//}
		if ( inUpFrm.CONSULTGB.value  == '' )
		{
				alert('상담분야를 선택하십시오!');
				inUpFrm.CONSULTGB.focus();
				return false;			
		}
		if ( inUpFrm.LEVEL1.value  == '' )
		{
				alert('계급을 선택하십시오!');
				inUpFrm.LEVEL1.focus();
				return false;			
		}
		// 계급이 기타, 미상이 아니면 세부 입력 해야함.
		if ( inUpFrm.LEVEL1.value != 'Z' && inUpFrm.LEVEL1.value != 'C' && inUpFrm.LEVEL2.value  == '' )
		{
				alert('세부계급을 선택하십시오!');
				return false;			
		}

		// 상담종류가 (상담,문의,사이버)
		if ( inUpFrm.ACLASS.value == 'A' || inUpFrm.ACLASS.value == 'B' || inUpFrm.ACLASS.value == 'D' )
		{
			if ( inUpFrm.PROCESSGB.value == '' )
			{
					alert('[상담,문의,사이버]일 경우 조치결과는 필수 입력항목입니다!');
					inUpFrm.PROCESSGB.focus();
					return false;			
			}
		}

*/
		//divCALLCLASS_B_2.style.display = "block";
		//divCALLCLASS_C_2.style.display = "block";
		inUpFrm.submit();
	}

	function FielUp(ty,cn){
		//alert("TYPE=" +ty+ ", COUNT=" +cn);
		strTemp = eval("inUpFrm.FILENAME1").value;
		//if (strTemp!=""){fileCNT="2"} else {fileCNT="1"}
		fileCNT="1";
		POPLayerURL = "LifecallFileUpload.asp?fileCNT=" +fileCNT+ "&frmTYPE=" +ty+cn;
		ShowPOPLayer(POPLayerURL,'500','160');
	}


	function ReferUp(ty,cn){
		//alert("TYPE=" +ty+ ", COUNT=" +cn);
		strTemp = eval("inUpFrm.FILENAME1").value;
		//if (strTemp!=""){fileCNT="2"} else {fileCNT="1"}
		fileCNT="1";
		POPLayerURL = "lifecallmanage_refersearch.asp?QueryYN=Y&whereCD8="+document.all.CUSTNAME.value;
		ShowPOPLayer(POPLayerURL,'1000','600');
	}

	function FileDel(ty,f){
		//alert("TYPE=" +ty+ ", COUNT=" +f);
		if(confirm("해당 데이타를 삭제 하시겠습니까?")) {
			hiddenIframe.location.href="LifecallFileUpload_Del.asp?frmTYPE="+ty+"&fn=" +f;
		}
	}

	function ReferDel(arg0,arg1){
		if(confirm("해당 관련상담번호를 삭제 하시겠습니까?")) {
			hiddenIframe.location.href="Refer_Del.asp?JUBSEQ="+arg0+"&REFERJUBSEQ="+arg1;
		}		
	}

	function fn_Player(arg0){
		//파일명
		var x,y;
		x = ( screen.width - 300 )/2;
		y = ( screen.height - 200 )/2;

		ShowPOPLayer("/include/WavePlayer.asp?URL="+arg0,'300','200');	
		//window.open("/include/WavePlayer.asp?URL="+arg0,"Player", "toolbar=no,top=100,left=300,width=300,height=200,resize=no,status=yes, scrollbars=no");
	}

	function fn_DEL(arg0)
	{
		//alert(arg0);
		if ( arg0 == 'WEATHER' || arg0 == 'PROCESSGB' || arg0 == 'REQUESTERGB' || arg0 == 'FAMILYGB' || arg0 == 'CALLFLAG' || arg0 == 'CALLKIND_B' || arg0 == 'CALLKIND_B_2')
		{
			//날씨

			for (var i=0;i<=100;i++)
			{
				if ( eval("inUpFrm."+arg0+"("+i+")") != null )
					eval("inUpFrm."+arg0+"("+i+")").checked = false;
				else
					break;
			}
		}
		else if ( arg0 == 'frame_callkind_B' ) //원인제공자#1-1차
		{
			document.all.CALLKIND_B.value = "";
			document.all.CALLKIND_C.value = "";
			frame_callkind_B.location.href ="/menu03/submenu0321/frame_CLASS.asp?frame_nm=frame_callkind_B&CLASSNM=CALLKIND&CLASSGB=B&ACLASS=R&BCLASS=";
			frame_callkind_C.location.href ="/menu03/submenu0321/frame_CLASS.asp?frame_nm=frame_callclass_C&CLASSNM=CALLKIND&CLASSGB=C&ACLASS=O&BCLASS=&CCLASS=";
		}
		else if ( arg0 == 'frame_callkind_C' ) //원인제공자#1-2차
		{
			
			document.all.CALLKIND_C.value = "";
			frame_callkind_C.location.href ="/menu03/submenu0321/frame_CLASS.asp?frame_nm=frame_callclass_C&CLASSNM=CALLKIND&CLASSGB=C&ACLASS=O&BCLASS="+document.all.CALLKIND_B.value+"&CCLASS=";
		}
		else if ( arg0 == 'frame_callkind_B_2' ) //원인제공자#1-1차
		{
			document.all.CALLKIND_B_2.value = "";
			document.all.CALLKIND_C_2.value = "";
			frame_callkind_B_2.location.href ="/menu03/submenu0321/frame_CLASS.asp?frame_nm=frame_callkind_B_2&CLASSNM=CALLKIND_2&CLASSGB=B&ACLASS=R&BCLASS=";
			frame_callkind_C_2.location.href ="/menu03/submenu0321/frame_CLASS.asp?frame_nm=frame_callkind_C_2&CLASSNM=CALLKIND_2&CLASSGB=C&ACLASS=O&BCLASS=&CCLASS=";
		}
		else if ( arg0 == 'frame_callkind_C_2' ) //원인제공자#1-2차
		{
			document.all.CALLKIND_C_2.value = "";
			frame_callkind_C_2.location.href ="/menu03/submenu0321/frame_CLASS.asp?frame_nm=frame_callkind_C_2&CLASSNM=CALLKIND_2&CLASSGB=C&ACLASS=O&BCLASS="+document.all.CALLKIND_B_2.value+"&CCLASS=";
		}

		else if ( arg0 == 'frame_callclass_B' ) //상담분야#1-1차
		{
			document.all.CALLCLASS_B.value = "";
			document.all.CALLCLASS_C.value = "";
			frame_callclass_B.location.href ="/menu03/submenu0321/frame_CLASS.asp?frame_nm=frame_callclass_B&CLASSNM=CALLCLASS&CLASSGB=B&ACLASS=O&BCLASS=";
			frame_callclass_C.location.href ="/menu03/submenu0321/frame_CLASS.asp?frame_nm=frame_callclass_C&CLASSNM=CALLCLASS&CLASSGB=C&ACLASS=O&BCLASS=&CCLASS=";
		}
		else if ( arg0 == 'frame_callclass_C' ) //상담분야#1-2차
		{
			document.all.CALLCLASS_C.value = "";
			frame_callclass_C.location.href ="/menu03/submenu0321/frame_CLASS.asp?frame_nm=frame_callclass_C&CLASSNM=CALLCLASS&CLASSGB=C&ACLASS=O&BCLASS="+document.all.CALLCLASS_B.value+"&CCLASS=";
		}
		else if ( arg0 == 'frame_callclass_B_2' ) //상담분야#2-1차
		{
			document.all.CALLCLASS_B_2.value = "";
			document.all.CALLCLASS_C_2.value = "";
			frame_callclass_B_2.location.href ="/menu03/submenu0321/frame_CLASS.asp?frame_nm=frame_callclass_B_2&CLASSNM=CALLCLASS_2&CLASSGB=B&ACLASS=O&BCLASS=";
			frame_callclass_C_2.location.href ="/menu03/submenu0321/frame_CLASS.asp?frame_nm=frame_callclass_C_2&CLASSNM=CALLCLASS_2&CLASSGB=C&ACLASS=O&BCLASS=&CCLASS=";
		}
		else if ( arg0 == 'frame_callclass_C_2' ) //상담분야#1-2차
		{
			document.all.CALLCLASS_C_2.value = "";
			frame_callclass_C_2.location.href ="/menu03/submenu0321/frame_CLASS.asp?frame_nm=frame_callclass_C_2&CLASSNM=CALLCLASS_2&CLASSGB=C&ACLASS=O&BCLASS="+document.all.CALLCLASS_B_2.value+"&CCLASS=";
		}
		else if ( arg0 == 'frame_channelgb_B' ) //상담유형
		{
			document.all.CHANNELGB_B.value = "";
			document.all.CHANNELGB_C.value = "";
			frame_channelgb_B.location.href ="/menu03/submenu0321/frame_CLASS.asp?frame_nm=frame_channelgb_B&CLASSNM=CHANNELGB&CLASSGB=B&ACLASS=Q&BCLASS=";
			frame_channelgb_C.location.href ="/menu03/submenu0321/frame_CLASS.asp?frame_nm=frame_channelgb_C&CLASSNM=CHANNELGB&CLASSGB=C&ACLASS=Q&BCLASS=&CCLASS=";
		}
		else if ( arg0 == 'frame_channelgb_C' ) //상담유형
		{
			document.all.CHANNELGB_C.value = "";
			frame_channelgb_C.location.href ="/menu03/submenu0321/frame_CLASS.asp?frame_nm=frame_channelgb_C&CLASSNM=CHANNELGB&CLASSGB=C&ACLASS=Q&BCLASS="+document.all.CHANNELGB_B.value+"&CCLASS=";
		}
		else if ( arg0 == 'frame_level_B' ) //계급1차
		{
			document.all.LEVEL_B.value = "";
			document.all.LEVEL_C.value = "";
			document.all.LEVEL_D.value = "";
			frame_level_B.location.href ="/menu03/submenu0321/frame_CLASS.asp?frame_nm=frame_level_B&CLASSNM=LEVEL&CLASSGB=B&ACLASS=P&BCLASS=";
			frame_level_C.location.href ="/menu03/submenu0321/frame_CLASS.asp?frame_nm=frame_level_C&CLASSNM=LEVEL&CLASSGB=C&ACLASS=P&BCLASS=&CCLASS=&DCLASS=";
			frame_level_D.location.href ="/menu03/submenu0321/frame_CLASS.asp?frame_nm=frame_level_C&CLASSNM=LEVEL&CLASSGB=C&ACLASS=P&BCLASS=&CCLASS=&DCLASS=";
		}
		else if ( arg0 == 'frame_level_C' ) //계급1차
		{
			document.all.LEVEL_C.value = "";
			document.all.LEVEL_D.value = "";
			frame_level_C.location.href ="/menu03/submenu0321/frame_CLASS.asp?frame_nm=frame_level_C&CLASSNM=LEVEL&CLASSGB=C&ACLASS=P&BCLASS="+document.all.LEVEL_B.value+"&CCLASS=";
			frame_level_D.location.href ="/menu03/submenu0321/frame_CLASS.asp?frame_nm=frame_level_D&CLASSNM=LEVEL&CLASSGB=D&ACLASS=P&BCLASS="+document.all.LEVEL_B.value+"&CCLASS=&DCLASS=";
		}
		else if ( arg0 == 'frame_level_D' ) //계급1차
		{
			document.all.LEVEL_D.value = "";
			frame_level_D.location.href ="/menu03/submenu0321/frame_CLASS.asp?frame_nm=frame_level_D&CLASSNM=LEVEL&CLASSGB=D&ACLASS=P&BCLASS="+document.all.LEVEL_B.value+"&CCLASS="+document.all.LEVEL_C.value+"&DCLASS=";
		}
		else if ( arg0 == 'frame_sosok_A' )
		{
			document.all.SOSOKGB_A.value = "";
			document.all.SOSOKGB_B.value = "";
			document.all.SOSOKGB_C.value = "";
			document.all.SOSOKGB_D.value = "";
			document.all.SOSOKGB_E.value = "";

			frame_sosok_A.location.href = "/menu03/submenu0321/frame_CLASS.asp?frame_nm=frame_sosok_A&CLASSNM=SOSOK&CLASSGB=A&ACLASS=";
			frame_sosok_B.location.href = "/menu03/submenu0321/frame_CLASS.asp?frame_nm=frame_sosok_B&CLASSNM=SOSOK&CLASSGB=B&ACLASS=&BCLASS=";
			frame_sosok_C.location.href = "/menu03/submenu0321/frame_CLASS.asp?frame_nm=frame_sosok_C&CLASSNM=SOSOK&CLASSGB=C&ACLASS=&BCLASS=&CCLASS=";
			frame_sosok_D.location.href = "/menu03/submenu0321/frame_CLASS.asp?frame_nm=frame_sosok_D&CLASSNM=SOSOK&CLASSGB=D&ACLASS=&BCLASS=&CCLASS=&DCLASS=";
			frame_sosok_E.location.href = "/menu03/submenu0321/frame_CLASS.asp?frame_nm=frame_sosok_E&CLASSNM=SOSOK&CLASSGB=E&ACLASS=&BCLASS=&CCLASS=&DCLASS=&ECLASS=";
		}
		else if ( arg0 == 'frame_sosok_B' )
		{
			document.all.SOSOKGB_B.value = "";
			document.all.SOSOKGB_C.value = "";
			document.all.SOSOKGB_D.value = "";
			document.all.SOSOKGB_E.value = "";
			frame_sosok_B.location.href = "/menu03/submenu0321/frame_CLASS.asp?frame_nm=frame_sosok_B&CLASSNM=SOSOK&CLASSGB=B&ACLASS="+document.all.SOSOKGB_A.value + "&BCLASS=";
			frame_sosok_C.location.href = "/menu03/submenu0321/frame_CLASS.asp?frame_nm=frame_sosok_C&CLASSNM=SOSOK&CLASSGB=C&ACLASS="+document.all.SOSOKGB_A.value + "&BCLASS=&CCLASS=";
			frame_sosok_D.location.href = "/menu03/submenu0321/frame_CLASS.asp?frame_nm=frame_sosok_D&CLASSNM=SOSOK&CLASSGB=D&ACLASS="+document.all.SOSOKGB_A.value + "&BCLASS=&CCLASS=&DCLASS=";
			frame_sosok_E.location.href = "/menu03/submenu0321/frame_CLASS.asp?frame_nm=frame_sosok_E&CLASSNM=SOSOK&CLASSGB=E&ACLASS="+document.all.SOSOKGB_A.value + "&BCLASS=&CCLASS=&DCLASS=&ECLASS=";
		}
		else if ( arg0 == 'frame_sosok_C' )
		{
			document.all.SOSOKGB_C.value = "";
			document.all.SOSOKGB_D.value = "";
			document.all.SOSOKGB_E.value = "";
			frame_sosok_C.location.href = "/menu03/submenu0321/frame_CLASS.asp?frame_nm=frame_sosok_C&CLASSNM=SOSOK&CLASSGB=C&ACLASS="+document.all.SOSOKGB_A.value + "&BCLASS="+document.all.SOSOKGB_B.value + "&CCLASS=";
			frame_sosok_D.location.href = "/menu03/submenu0321/frame_CLASS.asp?frame_nm=frame_sosok_D&CLASSNM=SOSOK&CLASSGB=D&ACLASS="+document.all.SOSOKGB_A.value + "&BCLASS="+document.all.SOSOKGB_B.value + "&CCLASS=&DCLASS=";
			frame_sosok_E.location.href = "/menu03/submenu0321/frame_CLASS.asp?frame_nm=frame_sosok_E&CLASSNM=SOSOK&CLASSGB=E&ACLASS="+document.all.SOSOKGB_A.value + "&BCLASS="+document.all.SOSOKGB_B.value + "&CCLASS=&DCLASS=&ECLASS=";
		}
		else if ( arg0 == 'frame_sosok_D' )
		{
			document.all.SOSOKGB_D.value = "";
			document.all.SOSOKGB_E.value = "";
			frame_sosok_D.location.href = "/menu03/submenu0321/frame_CLASS.asp?frame_nm=frame_sosok_D&CLASSNM=SOSOK&CLASSGB=D&ACLASS="+document.all.SOSOKGB_A.value + "&BCLASS="+document.all.SOSOKGB_B.value + "&CCLASS="+document.all.SOSOKGB_C.value + "&DCLASS=";
			frame_sosok_E.location.href = "/menu03/submenu0321/frame_CLASS.asp?frame_nm=frame_sosok_E&CLASSNM=SOSOK&CLASSGB=E&ACLASS="+document.all.SOSOKGB_A.value + "&BCLASS="+document.all.SOSOKGB_B.value + "&CCLASS="+document.all.SOSOKGB_C.value + "&DCLASS=&ECLASS=";
		}
		else if ( arg0 == 'frame_sosok_E' )
		{
			document.all.SOSOKGB_E.value = "";
			frame_sosok_E.location.href = "/menu03/submenu0321/frame_CLASS.asp?frame_nm=frame_sosok_E&CLASSNM=SOSOK&CLASSGB=E&ACLASS="+document.all.SOSOKGB_A.value + "&BCLASS="+document.all.SOSOKGB_B.value + "&CCLASS="+document.all.SOSOKGB_C.value + "&DCLASS="+document.all.SOSOKGB_D.value + "&ECLASS=";
		}
	}

	function fn_chkCHANNELGB_B()
	{
		if ( document.all.CHANNELGB_B.value == 'Q05' || document.all.CHANNELGB_B.value == 'Q07' || document.all.CHANNELGB_B.value == 'Q99')
		{
			// 값을 선택할 수 없다.
			eval("divCALLCLASS_B_1").style.display = "none";
			eval("divCALLCLASS_C_1").style.display = "none";

			eval("divCALLKIND_1").style.display = "none";
			//eval("divCALLKIND_2").style.display = "none";
			eval("divREQUESTERGB").style.display = "none";
			eval("divPROCESSGB").style.display = "none";
			
			eval("divSOSOKGB_A").style.display = "none";
			eval("divSOSOKGB_B").style.display = "none";
			eval("divSOSOKGB_C").style.display = "none";
			eval("divSOSOKGB_D").style.display = "none";
			eval("divSOSOKGB_E").style.display = "none";
			eval("divLEVEL_B").style.display = "none";
			eval("divLEVEL_C").style.display = "none";
			eval("divLEVEL_D").style.display = "none";
			eval("divCALLFLAG").style.display = "none";
			eval("divFAMILYGB").style.display = "none";			

		}
		else
		{
			// 값을 선택할 수 없다.
			eval("divCALLCLASS_B_1").style.display = "block";
			eval("divCALLCLASS_C_1").style.display = "block";
			eval("divCALLKIND_1").style.display = "block";


			eval("divREQUESTERGB").style.display = "block";
			eval("divPROCESSGB").style.display = "block";				
			eval("divSOSOKGB_A").style.display = "block";
			eval("divSOSOKGB_B").style.display = "block";
			eval("divSOSOKGB_C").style.display = "block";
			eval("divSOSOKGB_D").style.display = "block";
			eval("divSOSOKGB_E").style.display = "block";
			eval("divLEVEL_B").style.display = "block";
			eval("divLEVEL_C").style.display = "block";
			eval("divLEVEL_D").style.display = "block";
			eval("divCALLFLAG").style.display = "block";
			eval("divFAMILYGB").style.display = "block";
			
		}
	}

	function fn_PopCatalog()
	{
		//파일명
		var x,y;
		x = ( screen.width - 300 )/2;
		y = ( screen.height - 200 )/2;
		ShowPOPLayer("/Popup/pop_post.asp",'420','400');	
		//window.open("/include/WavePlayer.asp?URL="+arg0,"Player", "toolbar=no,top=100,left=300,width=300,height=200,resize=no,status=yes, scrollbars=no");
	}
	function fn_SetHeight()
	{
		frame_sosok_B.style.height = "50";
	}

	function fnGoto3012(url)
	{
		location.href = url;
	}

	function fn_listdel(arg0)
	{

		if ( confirm('자료를 삭제하시겠습니까? 삭제후에는 복구되지 않습니다') )
		{
			location.href="/menu03/submenu0321/lifecallhistory_InsUpDel.asp?guboon=DEL&jubseq="+arg0+"&<%=where2%>";
		}
	}

</script>
<!-- #include virtual="/Include/Bottom.asp" -->