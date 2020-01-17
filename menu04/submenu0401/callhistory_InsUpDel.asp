<!-- #include virtual="/include/CacheNo.asp" -->
<!-- #include virtual="/include/common.asp" -->
<%

guboon = Request("guboon")						'저장/수정/삭제 FLAG

TELKIND = Request("TELKIND")	
JUBSEQ = Request("JUBSEQ")		
JUBDATE = Request("JUBDATE")	
JUBTIME = Request("JUBTIME")
IOFLAG = Request("IOFLAG")
CUSTNO =trim( Request("CUSTNO"))
CUSTNAME = Request("CUSTNAME")
TELNO = Request("TELNO")
TELNO2 = Request("TELNO2")
SEXGB = Request("SEXGB")
CHANNELGB = Request("CHANNELGB")
REQUESTERGB = Request("REQUESTERGB")
CONSULTGB = Request("CONSULTGB")
CONSULTETCGB = Request("CONSULTETCGB")
SOSOKGB = Request("SOSOKGB")
SOSOKETCGB = Request("SOSOKETCGB")
LEVEL1 = Request("LEVEL1")
LEVEL2 = Request("LEVEL2")
ACLASS = Request("ACLASS")
BCLASS = Request("BCLASS")
CCLASS = Request("CCLASS")
CHANNEL = Request("CHANNEL")
CALLFLAG = Request("CALLFLAG")
CALLKIND = Request("CALLKIND")
QUESTION = Request("QUESTION")
CB_SEQ = Request("CB_SEQ")
REPLY = Request("REPLY")
RESULTGB = Request("RESULTGB")
RESERVEDATE = Request("RESERVEDATE")
RESERVETIME = Request("RESERVETIME")
PROCESSGB = Request("PROCESSGB")
CALLID = Request("CALLID")
RECORDFILE = Request("RECORDFILE")
EMERYN = Request("EMERYN")
INCODE = SESSION("SS_LoginID")
if EMERYN = "" then
	EMERYN = "N"
end if
if EMERYN <> "Y" then
	EMERYN = "N"
end if

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

'response.Write "01051850478"

	On Error Resume next

SQL = "SELECT CONVERT(DATETIME,'" & JUBTIME&"')"
db.Execute(SQL)

if db.Errors.count <> 0 then
	Call UrlBack("날자형태가 아닙니다.(yyyy-mm-dd hh:nn:ss).\n\n다시 시도해 주세요")
end if

select case ucase(guboon)

' 등록
case "INS"

	
'response.Write "INS"

	INCODE = SESSION("SS_LoginID")

	If INCODE = "" Then	

		INCODE = Request.Cookies("ASRNC")("WebUserid")
		SQL=" SELECT *"
		SQL = SQL & " FROM TB_USERINFO"
		SQL = SQL & " WHERE USERID = '" & INCODE & "'"

		Set RS = db.Execute(SQL)

		If RS.eof = False Then
		
			SESSION("SS_LoginID") = RS("USERID")
			SESSION("SS_LoginNAME") = RS("UserName")
			SESSION("SS_Login_Secgroup") = RS("SECGROUP")
			SESSION("SS_Login_Grade") = RS("GRADE")
			SESSION("SS_Login_GradeName") = RS("GRADE")' db_getCodeName("Z03",RS("GRADE")) 
			SESSION("SS_Login_CTIYN") = RS("CTIYN")

			SS_LoginID = SESSION("SS_LoginID")
			SS_LoginNAME = SESSION("SS_LoginNAME")
			SS_Login_Secgroup = SESSION("SS_Login_Secgroup")
			SS_Login_Grade = SESSION("SS_Login_Grade")
			SS_Login_GradeName = SESSION("SS_Login_GradeName")
			SS_Login_Agentcode = SESSION("SS_Login_Agentcode")
			SS_Login_CTIYN = SESSION("SS_Login_CTIYN")

		End If

	end if
	
	Set Rs = server.createObject("ADODB.Recordset")
	SQL = "select MAX(JUBSEQ) from TB_CALLHISTORY where LEFT(JUBSEQ,6) = CONVERT(CHAR(6),GETDATE(),112)"
	Rs.open SQL,db
	IF ISNULL(Rs(0)) THEN
		JUBSEQ = LEFT(REPLACE(DATE(),"-",""),6)&"0001"
	ELSEIF cint(right(Rs(0),4)) + 1 < 10 then
		JUBSEQ = left(Rs(0),6) & "000" & cint(right(Rs(0),4)) + 1
	ELSEIF cint(right(Rs(0),4)) + 1 < 100 then
		JUBSEQ = left(Rs(0),6) & "00" & cint(right(Rs(0),4)) + 1
	ELSEIF cint(right(Rs(0),4)) + 1 < 1000 then
		JUBSEQ = left(Rs(0),6) & "0" & cint(right(Rs(0),4)) + 1
	END IF

	On Error Resume next
	db.begintrans	


		'신규내담자라면..
		IF CUSTNO ="" THEN

			Set Rs = server.createObject("ADODB.Recordset")
			SQL = "select MAX(CUSTNO) from TB_CUSTINFO where LEFT(CUSTNO,6) = CONVERT(CHAR(6),GETDATE(),112)"
			Rs.open SQL,db
			IF ISNULL(Rs(0)) THEN
				CUSTNO = LEFT(REPLACE(DATE(),"-",""),6)&"0001"
			ELSEIF cint(right(TRIM(Rs(0)),4)) + 1 < 10 then
				CUSTNO = left(TRIM(Rs(0)),6) & "000" & cint(right(TRIM(Rs(0)),4)) + 1
			ELSEIF cint(right(TRIM(Rs(0)),4)) + 1 < 100 then
				CUSTNO = left(TRIM(Rs(0)),6) & "00" & cint(right(TRIM(Rs(0)),4)) + 1
			ELSEIF cint(right(TRIM(Rs(0)),4)) + 1 < 1000 then
				CUSTNO = left(TRIM(Rs(0)),6) & "0" & cint(right(TRIM(Rs(0)),4)) + 1
			END IF

			'response.write CUSTNO

			SQL = "INSERT INTO TB_CUSTINFO ( CUSTNO"
			SQL = SQL & "		, NAME" 
			SQL = SQL & "		, SEX" 
			SQL = SQL & "		, CELLPHONE" 
			SQL = SQL & "		, HOMEPHONE" 
			SQL = SQL & "		, SENDPHONE" 
			SQL = SQL & "		, ACLASS" 
			SQL = SQL & "		, SOSOKGB" 
			SQL = SQL & "		, SOSOKETCGB" 
			SQL = SQL & "		,	LEVEL1"
			SQL = SQL & "		,	LEVEL2"
			SQL = SQL & "		,	INCODE"
			SQL = SQL & "		,	INDATE"
			SQL = SQL & "		,	MOCODE"
			SQL = SQL & "		,	MODATE)"
			SQL = SQL & "		VALUES ( '" & CUSTNO & "'"
			SQL = SQL & "		,	'" & CUSTNAME & "'"
			SQL = SQL & "		,	'" & SEXGB & "'"
			SQL = SQL & "		,	'" & TELNO & "'"
			SQL = SQL & "		,	'" & TELNO2 & "'"
			SQL = SQL & "		,	'" & CID & "'"
			SQL = SQL & "		,	'B'" '생명의전화
			SQL = SQL & "		,	'" & SOSOKGB & "'"
			SQL = SQL & "		,	'" & SOSOKETCGB & "'"
			SQL = SQL & "		,	'" & LEVEL1 & "'"
			SQL = SQL & "		,	'" & LEVEL2 & "'"
			SQL = SQL & "		,	'" & INCODE  & "'"
			SQL = SQL & "		,	GETDATE()"
			SQL = SQL & "		,	'" & INCODE  & "'"
			SQL = SQL & "		,	GETDATE())"

			db.execute(SQL)

		ELSE

			SQL = "UPDATE TB_CUSTINFO SET"
			SQL = SQL & "		NAME = '" & CUSTNAME & "'"
			SQL = SQL & "		, SEX = '" & SEXGB & "'"
			SQL = SQL & "		, CELLPHONE ='" & TELNO & "'"
			SQL = SQL & "		, HOMEPHONE ='" & TELNO2 & "'" 
			SQL = SQL & "		, SENDPHONE ='" & CID & "'"
			SQL = SQL & "		, ACLASS ='B'" 
			SQL = SQL & "		, SOSOKGB ='" & SOSOKGB & "'" 
			SQL = SQL & "		, SOSOKETCGB='" & SOSOKETCGB & "'"
			SQL = SQL & "		,	LEVEL1='" & LEVEL1 & "'"
			SQL = SQL & "		,	LEVEL2='" & LEVEL2 & "'"

			SQL = SQL & "		,	MOCODE='" & INCODE  & "'"
			SQL = SQL & "		,	MODATE=GETDATE()"
			SQL = SQL & "		WHERE	CUSTNO = '" & CUSTNO &"'"

			db.execute(SQL)


		END IF


		SQL = " INSERT INTO TB_CALLHISTORY ( JUBSEQ, TELKIND"
		SQL = SQL & "		,	JUBDATE"
		SQL = SQL & "		,	JUBTIME"
		SQL = SQL & "		,	IOFLAG"
		SQL = SQL & "		,	CUSTNO"
		SQL = SQL & "		,	CUSTNAME"
		SQL = SQL & "		,	TELNO"
		SQL = SQL & "		,	TELNO2"
		SQL = SQL & "		,	SEXGB"
		SQL = SQL & "		,	CHANNELGB"
		SQL = SQL & "		,	REQUESTERGB"
		SQL = SQL & "		,	CONSULTGB"
		SQL = SQL & "		,	CONSULTETCGB"
		SQL = SQL & "		,	SOSOKGB"
		SQL = SQL & "		,	SOSOKETCGB"
		SQL = SQL & "		,	LEVEL1"
		SQL = SQL & "		,	LEVEL2"
		SQL = SQL & "		,	ACLASS"
		SQL = SQL & "		,	BCLASS"
		SQL = SQL & "		,	CCLASS"
		SQL = SQL & "		,	CHANNEL"
		SQL = SQL & "		,	CALLFLAG"
		SQL = SQL & "		,	CALLKIND"
		SQL = SQL & "		,	QUESTION"
		SQL = SQL & "		,	REPLY"
		SQL = SQL & "		,	RESULTGB"
		SQL = SQL & "		,	RESERVEDATE"
		SQL = SQL & "		,	RESERVETIME"
		SQL = SQL & "		,	PROCESSGB"
		SQL = SQL & "		,	CALLID"
		SQL = SQL & "		,	RECORDFILE"
		SQL = SQL & "		,	CB_SEQ"
		SQL = SQL & "		,	INCODE"
		SQL = SQL & "		,	INDATE"
		SQL = SQL & "		,	MOCODE"
		SQL = SQL & "		,	MODATE)"
		SQL = SQL & "		VALUES ( '" & JUBSEQ & "','"&TELKIND&"'"
		SQL = SQL & "		,	'" & LEFT(JUBTIME,10) & "'"
		SQL = SQL & "		,	'" & JUBTIME & "'"
		SQL = SQL & "		,	'" & IOFLAG & "'"
		SQL = SQL & "		,	'" & CUSTNO & "'"
		SQL = SQL & "		,	'" & CUSTNAME & "'"
		SQL = SQL & "		,	'" & TELNO & "'"
		SQL = SQL & "		,	'" & TELNO2 & "'"
		SQL = SQL & "		,	'" & SEXGB & "'"
		SQL = SQL & "		,	'" & CHANNELGB & "'"
		SQL = SQL & "		,	'" & REQUESTERGB & "'"
		SQL = SQL & "		,	'" & CONSULTGB & "'"
		SQL = SQL & "		,	'" & CONSULTETCGB & "'"
		SQL = SQL & "		,	'" & SOSOKGB & "'"
		SQL = SQL & "		,	'" & SOSOKETCGB & "'"
		SQL = SQL & "		,	'" & LEVEL1 & "'"
		SQL = SQL & "		,	'" & LEVEL2 & "'"
		SQL = SQL & "		,	'" & ACLASS & "'"
		SQL = SQL & "		,	'" & BCLASS & "'"
		SQL = SQL & "		,	'" & CCLASS & "'"
		SQL = SQL & "		,	'" & CHANNEL & "'"
		SQL = SQL & "		,	'" & CALLFLAG  & "'"
		SQL = SQL & "		,	'" & CALLKIND  & "'"
		SQL = SQL & "		,	'" & QUESTION  & "'"
		SQL = SQL & "		,	'" & REPLY & "'"
		SQL = SQL & "		,	'" & RESULTGB & "'"
		SQL = SQL & "		,	'" & RESERVEDATE & "'"
		SQL = SQL & "		,	'" & RESERVETIME & "'"
		SQL = SQL & "		,	'" & PROCESSGB  & "'"
		SQL = SQL & "		,	'" & CALLID  & "'"
		SQL = SQL & "		,	'" & RECORDFILE  & "'"
		SQL = SQL & "		,	'" & CB_SEQ  & "'"
		SQL = SQL & "		,	'" & INCODE  & "'"
		SQL = SQL & "		,	GETDATE()"
		SQL = SQL & "		,	'" & INCODE  & "'"
		SQL = SQL & "		,	GETDATE())"

		'SQL ="INSERT INTO TB_CODE (CODEGROUP,CODE,GROUPNAME,CODENAME,MEMO,USEYN,SYSYN,INCODE) VALUES "
		'SQL = SQL & "('" & CODEGROUP & "','" & CODE & "','" & GROUPNAME & "','" & CODENAME & "','" & MEMO & "','" & USEYN & "','" & SYSYN '& "','" & INCODE & "')"
		
		'response.write SQL
		db.execute(SQL)

		if db.Errors.count = 0 then
			flag = "Y"
			'LogWrite "SQL="&SQL, "Code_InsUpDel.asp", "/Setup/Code/"			
		else
			flag = "N"
		end if
		
		if flag = "Y" then

			if CB_SEQ <> "" then

				SQL = "	UPDATE	TB_CALLBACK SET PROCESSGB = 'C', PROCESSCODE = '" & INCODE  & "', PROCESSTIME = GETDATE()"
				SQL = SQL & "	WHERE	IDX = " & CB_SEQ
				db.execute(SQL)

			end if
			db.CommitTrans
			
			pageUrl = "/menu04/submenu0401/calllist.asp"
			Call MsgGoUrl( "정상적으로 등록되었습니다.",pageUrl)
		else
			db.RollBackTrans
			'LogWrite "ERROR_SQL="&SQL, "Code_InsUpDel.asp", "/Setup/Code/"
			'Call UrlBack("저장중 에러가 발생했습니다.\n\n다시 시도해 주세요")
		end if

	
	Rs.Close
	set Rs = NOTHING


' 수정
case "UP"
	MOCODE = SESSION("SS_LoginID")

	If MOCODE = "" Then	

		MOCODE = Request.Cookies("ASRNC")("WebUserid")
		SQL=" SELECT *"
		SQL = SQL & " FROM TB_USERINFO"
		SQL = SQL & " WHERE USERID = '" & MOCODE & "'"

		Set RS = db.Execute(SQL)

		If RS.eof = False Then
		
			SESSION("SS_LoginID") = RS("USERID")
			SESSION("SS_LoginNAME") = RS("UserName")
			SESSION("SS_Login_Secgroup") = RS("SECGROUP")
			SESSION("SS_Login_Grade") = RS("GRADE")
			SESSION("SS_Login_GradeName") = RS("GRADE")' db_getCodeName("Z03",RS("GRADE")) 
			SESSION("SS_Login_CTIYN") = RS("CTIYN")

			SS_LoginID = SESSION("SS_LoginID")
			SS_LoginNAME = SESSION("SS_LoginNAME")
			SS_Login_Secgroup = SESSION("SS_Login_Secgroup")
			SS_Login_Grade = SESSION("SS_Login_Grade")
			SS_Login_GradeName = SESSION("SS_Login_GradeName")
			SS_Login_Agentcode = SESSION("SS_Login_Agentcode")
			SS_Login_CTIYN = SESSION("SS_Login_CTIYN")

		End If

	end if
	
	On Error Resume next
	db.begintrans
	
	flag = "N"			'UPDATE 성공여부

		SQL = " UPDATE TB_LIFECALLHISTORY SET "
		SQL = SQL & "			IOFLAG = '" & IOFLAG & "'"
		SQL = SQL & "		,	CUSTNO = '" & CUSTNO & "'"
		SQL = SQL & "		,	CUSTNAME = '" & CUSTNAME & "'"
		SQL = SQL & "		,	TELNO = '" & TELNO & "'"
		SQL = SQL & "		,	TELNO2 = '" & TELNO2 & "'"
		SQL = SQL & "		,	SEXGB = '" & SEXGB & "'"
		SQL = SQL & "		,	CHANNELGB = '" & CHANNELGB & "'"
		SQL = SQL & "		,	REQUESTERGB = '" & REQUESTERGB & "'"
		SQL = SQL & "		,	CONSULTGB = '" & CONSULTGB & "'"
		SQL = SQL & "		,	CONSULTETCGB = '" & CONSULTETCGB & "'"
		SQL = SQL & "		,	SOSOKGB = '" & SOSOKGB & "'"
		SQL = SQL & "		,	SOSOKETCGB = '" & SOSOKETCGB & "'"
		SQL = SQL & "		,	LEVEL1 = '" & LEVEL1 & "'"
		SQL = SQL & "		,	LEVEL2 = '" & LEVEL2 & "'"
		SQL = SQL & "		,	ACLASS = '" & ACLASS & "'"
		SQL = SQL & "		,	BCLASS = '" & BCLASS & "'"
		SQL = SQL & "		,	CCLASS = '" & CCLASS & "'"
		SQL = SQL & "		,	CHANNEL = '" & CHANNEL & "'"
		SQL = SQL & "		,	CALLFLAG = '" & CALLFLAG & "'"
		SQL = SQL & "		,	CALLKIND = '" & CALLKIND & "'"
		SQL = SQL & "		,	QUESTION = '" & QUESTION & "'"
		SQL = SQL & "		,	REPLY = '" & REPLY & "'"
		SQL = SQL & "		,	RESULTGB = '" & RESULTGB & "'"
		SQL = SQL & "		,	RESERVEDATE = '" & RESERVEDATE & "'"
		SQL = SQL & "		,	RESERVETIME = '" & RESERVETIME & "'"
		SQL = SQL & "		,	PROCESSGB = '" & PROCESSGB & "'"
		SQL = SQL & "		,	CALLID = '" & CALLID & "'"
		SQL = SQL & "		,	RECORDFILE = '" & RECORDFILE & "'"
		SQL = SQL & "		,	JUBTIME = '"&JUBTIME&"'"	
		SQL = SQL & "		,	JUBDATE = '"&LEFT(JUBTIME,10)&"'"			
		SQL = SQL & "		,	TELKIND = '" & TELKIND & "'"		
		SQL = SQL & "		,	MOCODE= '" & MOCODE & "'"
		SQL = SQL & "		,	MODATE=getdate()"
		SQL = SQL & "		where JUBSEQ = '" & TRIM(JUBSEQ) & "'"

'response.write SQL

	db.execute(SQL)	
	if db.Errors.count = 0 then
		flag = "Y"
		'LogWrite "SQL="&SQL, "Code_InsUpDel.asp", "/Setup/Code/"
	else
		flag = "N"
	end If
	
	if flag = "Y" then
		db.CommitTrans

		where1 = "QueryYN=Y&FromDate=" & FromDate & "&ToDate=" & ToDate & "&whereCD1=" & whereCD1 & "&whereCD2=" & whereCD2 & "&whereCD3=" & whereCD3 & "&whereCD4=" & whereCD4 & "&whereCD5=" & whereCD5 & "&whereCD6=" & whereCD6 & "&whereCD7=" & whereCD7 & "&whereCD8=" & whereCD8 & "&whereCD9=" & whereCD9 & "&whereCD10=" & whereCD10 & "&whereCD11=" & whereCD11 & "&whereCD12=" & whereCD12
		where2 = "curPage=" & curPage & "&" & where1

		pageUrl = "/menu03/submenu0301/lifecallhistory.asp?"&where2
		Call MsgGoUrl( "정상적으로 수정되었습니다.",pageUrl)
	else
		db.RollBackTrans
		'LogWrite "ERROR_SQL="&SQL, "Code_InsUpDel.asp", "/Setup/Code/"
		'Call UrlBack("수정중 에러가 발생했습니다.\n\n다시 시도해 주세요")
	end if

' 삭제
case "DEL"
	flag = "N"				'DELETE 성공여부
	
	SQL = "DELETE FROM TB_LIFECALLHISTORY WHERE JUBSEQ = '" & TRIM(JUBSEQ) & "'"
	
	set Rs = db.execute(SQL)		
		
	if db.Errors.count = 0 then
		flag = "Y"
		'LogWrite "SQL="&SQL, "Code_InsUpDel.asp", "/Setup/Code/"
	else
		flag = "N"
	end if
	
	if flag = "Y" then
		'db.CommitTrans



		where1 = "QueryYN=Y&FromDate=" & FromDate & "&ToDate=" & ToDate & "&whereCD1=" & whereCD1 & "&whereCD2=" & whereCD2 & "&whereCD3=" & whereCD3 & "&whereCD4=" & whereCD4 & "&whereCD5=" & whereCD5 & "&whereCD6=" & whereCD6 & "&whereCD7=" & whereCD7 & "&whereCD8=" & whereCD8 & "&whereCD9=" & whereCD9 & "&whereCD10=" & whereCD10 & "&whereCD11=" & whereCD11 & "&whereCD12=" & whereCD12
		where2 = "curPage=" & curPage & "&" & where1


		pageUrl = "/menu03/submenu0301/lifecallhistory.asp?"&where2
		'response.write pageUrl

		Call MsgGoUrl( "정상적으로 삭제되었습니다.",pageUrl)
	else
		'db.RollBackTrans
		'LogWrite "ERROR_SQL="&SQL, "Code_InsUpDel.asp", "/Setup/Code/"
		Call UrlBack("삭제 중 에러가 발생했습니다.\n\n다시 시도해 주세요")
	end if
		
end Select



%>