<!-- #include virtual="/include/CacheNo.asp" -->
<!-- #include virtual="/include/common.asp" -->


<%

Seq = Request("Seq")									'KEY
class_gb = Request("class_gb")					'클래스 구분(A:1차분류, B:2차분류, C:3차분류, D:4차분류, E:5차분류)
db_flag = Request("db_flag")							'DB 처리 구분(INS:INSERT, UP:UPDATE, DEL:DELETE)

Aclass = Request("Aclass")				'1차분류
Bclass = Request("Bclass")				'2차분류
Cclass = Request("Cclass")				'3차분류
Dclass = Request("Dclass")				'3차분류
Eclass = Request("Eclass")				'3차분류

s_code = ucase(ConvertName(Request("code")))			'폼값(코드)
s_name = ConvertName(Request("code_name"))		'폼값(코드명)
UseYN = TRIM(Request("UseYN"))
COUNSELORYN = TRIM(Request("COUNSELORYN"))
keyword = TRIM(Request("keyword"))
telno = TRIM(Request("telno"))
telno2 = TRIM(Request("telno2"))
'LogWrite "Seq="&Seq&"class_gb="&class_gb&"db_flag="&db_flag&"Aclass="&Aclass&"Bclass="&Bclass&"Cclass="&Cclass&"Dclass="&Dclass&"s_code="&s_code&"s_name="&s_name, "ProType_InsUpDel.asp", "/Setup/ProType/"

select case ucase(db_flag)


' 등록
case "INS"
	INCODE = SESSION("SS_LoginID")
	
	select case ucase(class_gb)
		
	case "A"
		SQL = "SELECT * FROM TB_ARMYINFO WHERE ACLASS = '" & s_code & "' AND BCLASS IS NULL AND CCLASS IS NULL "
		
		INS_SQL = "INSERT INTO TB_ARMYINFO(ACLASS, BCLASS, CCLASS, CLASSNAME, USEYN, INCODE) VALUES "
		INS_SQL = INS_SQL & "('"&s_code&"',null,null,'"&s_name&"','" &UseYN& "','"&INCODE&"')"
		pageUrl = "ProType_1.asp"
		pageFrame = "Pro1fr"
	case "B"
		SQL = "SELECT * FROM TB_ARMYINFO WHERE ACLASS = '" & Aclass & "' AND BCLASS = '" & s_code & "' AND CCLASS IS NULL "
		
		INS_SQL = "INSERT INTO TB_ARMYINFO(ACLASS, BCLASS, CCLASS, CLASSNAME, USEYN, INCODE,COUNSELORYN,KEYWORD,TELNO,TELNO2) VALUES "
		INS_SQL = INS_SQL & "('"&Aclass&"','"&s_code&"',null,'"&s_name&"','" &UseYN& "','"&INCODE&"','"&COUNSELORYN&"','"&keyword&"','"&telno&"','"&telno2&"')"
		pageUrl = "ProType_2.asp?Aclass=" & Aclass
		pageFrame = "Pro2fr"
	case "C"
		SQL = "SELECT * FROM TB_ARMYINFO WHERE ACLASS = '" & Aclass & "' AND BCLASS = '" & Bclass & "' AND CCLASS = '" &s_code& "' "
		
		INS_SQL = "INSERT INTO TB_ARMYINFO(ACLASS, BCLASS, CCLASS, CLASSNAME, USEYN, INCODE, COUNSELORYN,KEYWORD,TELNO,TELNO2) VALUES "
		INS_SQL = INS_SQL & "('"&Aclass&"','"&Bclass&"','"&s_code&"','"&s_name&"','" &UseYN& "','"&INCODE&"','"&COUNSELORYN&"','"&keyword&"','"&telno&"','"&telno2&"')"
		pageUrl = "ProType_3.asp?Aclass=" & Aclass & "&Bclass=" & Bclass
		pageFrame = "Pro3fr"

	case "D"
		SQL = "SELECT * FROM TB_ARMYINFO WHERE ACLASS = '" & Aclass & "' AND BCLASS = '" & Bclass & "' AND CCLASS = '" &Cclass& "' AND DCLASS = '" &s_code& "'"
		
		INS_SQL = "INSERT INTO TB_ARMYINFO(ACLASS, BCLASS, CCLASS,DCLASS, CLASSNAME, USEYN, INCODE, COUNSELORYN,KEYWORD,TELNO,TELNO2) VALUES "
		INS_SQL = INS_SQL & "('"&Aclass&"','"&Bclass&"','"&Cclass&"','"&s_code&"','"&s_name&"','" &UseYN& "','"&INCODE&"','"&COUNSELORYN&"','"&keyword&"','"&telno&"','"&telno2&"')"
		pageUrl = "ProType_4.asp?Aclass=" & Aclass & "&Bclass=" & Bclass& "&Cclass=" & Cclass
		pageFrame = "Pro4fr"

	case "E"
		SQL = "SELECT * FROM TB_ARMYINFO WHERE ACLASS = '" & Aclass & "' AND BCLASS = '" & Bclass & "' AND CCLASS = '" &Cclass& "' AND DCLASS = '" &s_code& "'"
		
		INS_SQL = "INSERT INTO TB_ARMYINFO(ACLASS, BCLASS, CCLASS,DCLASS,ECLASS, CLASSNAME, USEYN, INCODE, COUNSELORYN,KEYWORD,TELNO,TELNO2) VALUES "
		INS_SQL = INS_SQL & "('"&Aclass&"','"&Bclass&"','"&Cclass&"','"&Dclass&"','"&s_code&"','"&s_name&"','" &UseYN& "','"&INCODE&"','"&COUNSELORYN&"','"&keyword&"','"&telno&"','"&telno2&"')"
		pageUrl = "ProType_5.asp?Aclass=" & Aclass & "&Bclass=" & Bclass& "&Cclass=" & Cclass& "&Dclass=" & Dclass
		pageFrame = "Pro5fr"

	end select
	
	'response.write INS_SQL

	Set Rs = db.execute(SQL)
	
	'LogWrite "SQL="&SQL, "ProType_InsUpDel.asp", "/Setup/ProType/"
	
	On Error Resume next
	'db.begintrans
	
	If Rs.Eof Or Rs.bof Then
		flag = "N"			'INSERT 성공여부
		'제품분류 입력
		db.execute(INS_SQL)

		if db.Errors.count = 0 then
			flag = "Y"
			'LogWrite "INS_SQL="&INS_SQL, "ProType_InsUpDel.asp", "/Setup/ProType/"
			
		else
			flag = "N"
			'LogWrite "ERROR_SQL="&INS_SQL, "ProType_InsUpDel.asp", "/Setup/ProType/"
		end if
		
		if flag = "Y" then
			'db.CommitTrans
			Call PFrameMsgGoUrl("정상적으로 등록되었습니다.",pageUrl,pageFrame)
		else
			'db.RollBackTrans
			Call UrlBack("저장중 에러가 발생했습니다.\n\n다시 시도해 주세요")
		end if
	Else
		Call UrlBack("이미 존재하는 코드 입니다..\n\n다른 코드를 입력해 주세요")
	end If
	
	Rs.Close
	set Rs = NOTHING	



' 수정
case "UP"
	MOCODE = SESSION("SS_LoginID")
	
	select case ucase(class_gb)
		
	case "A"
		pageUrl = "ProType_1.asp"
		pageFrame = "Pro1fr"
	case "B"
		pageUrl = "ProType_2.asp?Aclass=" & Aclass
		pageFrame = "Pro2fr"
	case "C"
		pageUrl = "ProType_3.asp?Aclass=" & Aclass & "&Bclass=" & Bclass
		pageFrame = "Pro3fr"
	case "D"
		pageUrl = "ProType_4.asp?Aclass=" & Aclass & "&Bclass=" & Bclass& "&Cclass=" & Cclass
		pageFrame = "Pro4fr"

	case "E"
		pageUrl = "ProType_5.asp?Aclass=" & Aclass & "&Bclass=" & Bclass& "&Cclass=" & Cclass& "&Dclass=" & Dclass
		pageFrame = "Pro5fr"

	end select
	
	On Error Resume next
	db.begintrans
	
	flag = "N"			'UPDATE 성공여부
	'제품분류 1차 수정
	SQL = "UPDATE TB_ARMYINFO SET COUNSELORYN = '" &COUNSELORYN& "', UseYN='" &UseYN& "', CLASSNAME='" &s_name& "',KEYWORD='"&KEYWORD&"',TELNO='"&TELNO&"',TELNO2='"&TELNO2&"', MOCODE='" &MOCODE& "', MODATE = getdate()"
	SQL = SQL& " WHERE SEQ='" &Seq& "'"
	
	db.execute(SQL)
	
	if db.Errors.count = 0 then
		flag = "Y"
		'LogWrite "SQL="&SQL, "ProType_InsUpDel.asp", "/Setup/ProType/"

	else
		flag = "N"
		'LogWrite "ERROR_SQL="&SQL, "ProType_InsUpDel.asp", "/Setup/ProType/"
	end If
	
	if flag = "Y" then
		db.CommitTrans
		Call PFrameMsgGoUrl("정상적으로 수정되었습니다.",pageUrl,pageFrame)
	else
		db.RollBackTrans
		Call UrlBack("수정중 에러가 발생했습니다.\n\n다시 시도해 주세요")
	end if

' 삭제
case "DEL"

	flag = "N"				'DELETE 성공여부
	del_flag = "N"		'삭제 가능 여부
	
	select case ucase(class_gb)
		
	case "A"
		SQL = "SELECT * FROM TB_ARMYINFO WHERE ACLASS = '" &Aclass& "' AND BCLASS IS NOT NULL"
		set Rs = db.execute(SQL)
		
		if Rs.EOF OR RS.BOF then	'제품분류의 하위 제품분류가 없는 경우
			del_flag = "Y"						'삭제 가능 여부
		end if
		
		Rs.Close
		set Rs = NOTHING
		
		pageUrl = "ProType_1.asp"
		pageFrame = "Pro1fr"

	case "B"
		SQL = "SELECT * FROM TB_ARMYINFO WHERE ACLASS = '" & Aclass & "' AND BCLASS = '" & Bclass & "' AND CCLASS IS NOT NULL"
		set Rs = db.execute(SQL)
			
		if Rs.EOF OR RS.BOF then	'제품분류의 하위 제품분류가 없는 경우
			del_flag = "Y"						'삭제 가능 여부
		end if
		
		Rs.Close
		set Rs = NOTHING
		
		pageUrl = "ProType_2.asp?Aclass=" & Aclass
		pageFrame = "Pro2fr"
	case "C"

		del_flag = "Y"						'삭제 가능 여부
		
		pageUrl = "ProType_3.asp?Aclass=" & Aclass & "&Bclass=" & Bclass
		pageFrame = "Pro3fr"

	case "D"

		del_flag = "Y"						'삭제 가능 여부
		
		pageUrl = "ProType_4.asp?Aclass=" & Aclass & "&Bclass=" & Bclass & "&Cclass=" & Cclass
		pageFrame = "Pro4fr"

	case "E"

		del_flag = "Y"						'삭제 가능 여부
		
		pageUrl = "ProType_5.asp?Aclass=" & Aclass & "&Bclass=" & Bclass & "&Cclass=" & Cclass & "&Dclass=" & Dclass
		pageFrame = "Pro5fr"

	end select
	'LogWrite "del_flag="&del_flag, "ProType_InsUpDel.asp", "/Setup/ProType/"
	'LogWrite "SQL="&SQL, "ProType_InsUpDel.asp", "/Setup/ProType/"
	if del_flag = "Y" then	'제품분류의 하위 제품분류가 없는 경우
		On Error Resume next
		db.begintrans
		
		'제품분류 삭제
		SQL = "DELETE FROM TB_ARMYINFO WHERE SEQ = '" & Seq & "'"
		
		db.execute(SQL)
	
		if db.Errors.count = 0 then
			flag = "Y"
			'LogWrite "SQL="&SQL, "ProType_InsUpDel.asp", "/Setup/ProType/"
		else
			flag = "N"
		end if
	
			
		if flag = "Y" then
			db.CommitTrans
			Call PFrameMsgGoUrl("정상적으로 삭제되었습니다.",pageUrl,pageFrame)
		else
			db.RollBackTrans
			'LogWrite "ERROR_SQL="&SQL, "ProType_InsUpDel.asp", "/Setup/ProType/"
			Call UrlBack("삭제 중 에러가 발생했습니다.\n\n다시 시도해 주세요")
		end if
		
	else										'제품분류의 하위 제품분류가 있는 경우
		Call UrlBack("하위 제품분류가 있는 코드 입니다.\n\n하위 제품 삭제 후 삭제하시기 바랍니다.")
	end if
			
end Select



%>