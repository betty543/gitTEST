<!-- #include virtual="/include/CacheNo.asp" -->
<!-- #include virtual="/include/common.asp" -->


<%

guboon = Request("guboon")						'저장/수정/삭제 FLAG
curPage = Request("curPage")
sCode = Request("sCode")							'코드
sCodegroup = Request("sCodegroup")		'구분
sGroupname = Request("sGroupname")		'구분명

CODEGROUP = ConvertName(request("txtCodeGroup"))			'입력/수정폼값(구분)
GROUPNAME = ConvertName(request("txtGroupName"))			'입력/수정폼값(구분명)
CODE = trim(ucase(ConvertName(request("txtCode"))))					'입력/수정폼값(코드)
CODENAME =trim(request("txtCodeName"))	'입력/수정폼값(코드명)

MEMO = ConvertString(Request("txtMemo"))								'입력/수정폼값(메모)
USEYN = Request("optUseYN")													'입력/수정폼값(사용여부)
SYSYN = Request("optSysYN")													'입력/수정폼값(시스템)

select case ucase(guboon)


' 등록
case "INS"
	INCODE = SESSION("SS_LoginID")
	
	Set Rs = server.createObject("ADODB.Recordset")
	SQL = "select * from TB_CODE where CODEGROUP = '" & CODEGROUP & "' and CODE = '" & CODE & "'"
	Rs.open SQL,db
	
	On Error Resume next
	db.begintrans
	
	If Rs.Eof Or Rs.bof Then
		flag = "N"			'INSERT 성공여부
		'코드관리 입력
		SQL ="INSERT INTO TB_CODE (CODEGROUP,CODE,GROUPNAME,CODENAME,MEMO,USEYN,SYSYN,INCODE) VALUES "
		SQL = SQL & "('" & CODEGROUP & "','" & CODE & "','" & GROUPNAME & "','" & CODENAME & "','" & MEMO & "','" & USEYN & "','" & SYSYN & "','" & INCODE & "')"
		
		db.execute(SQL)

		if db.Errors.count = 0 then
			flag = "Y"
			'LogWrite "SQL="&SQL, "Code_InsUpDel.asp", "/Setup/Code/"
			
		else
			flag = "N"
		end if
		
		if flag = "Y" then
			db.CommitTrans
			
			pageUrl = "code_list.asp?curPage=" & curPage & "&sCodegroup=" & sCodegroup & "&sGroupname=" & sGroupname
			Call MsgGoUrl( "정상적으로 등록되었습니다.",pageUrl)
		else
			db.RollBackTrans
			'LogWrite "ERROR_SQL="&SQL, "Code_InsUpDel.asp", "/Setup/Code/"
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
	
	On Error Resume next
	db.begintrans
	
	flag = "N"			'UPDATE 성공여부
	
	SQL ="UPDATE TB_CODE SET CODENAME ='" & CODENAME & "',MEMO ='"& MEMO &"'"
	SQL = SQL & ",USEYN='" & USEYN & "',SYSYN='" & SYSYN & "',MOCODE='" & MOCODE & "',MODATE= getdate()"
	SQL = SQL & " WHERE CODEGROUP = '" & sCodegroup & "' and CODE = '" & sCode & "'" 

	db.execute(SQL)
	
	if db.Errors.count = 0 then
		flag = "Y"
		'LogWrite "SQL="&SQL, "Code_InsUpDel.asp", "/Setup/Code/"

	else
		flag = "N"
	end If
	
	if flag = "Y" then
		db.CommitTrans
		pageUrl = "code_list.asp?curPage=" & curPage & "&sCodegroup=" & sCodegroup & "&sGroupname=" & sGroupname
		Call MsgGoUrl( "정상적으로 수정되었습니다.",pageUrl)
	else
		db.RollBackTrans
		LogWrite "ERROR_SQL="&SQL, "Code_InsUpDel.asp", "/Setup/Code/"
		Call UrlBack("수정중 에러가 발생했습니다.\n\n다시 시도해 주세요")
	end if

' 삭제
case "DEL"
	flag = "N"				'DELETE 성공여부
	
	'해당 코드 사용여부 체크(해당 코드 사용시 삭제불가) - 삭제시 사용여부 테이블 확인 후 처리할 것
	if sCodegroup = "A01" then		'코드 그룹이 A01(제품분류)인 경우
		
		'제품분류 연동 체크(하위 제품분류 존재시 삭제불가)
		SQL = "SELECT * FROM TB_GOODBUNU WHERE ACLASS = '" &sCode& "' AND BCLASS IS NOT NULL"
		
		set Rs = db.execute(SQL)
			
		if Rs.EOF OR RS.BOF then	'제품분류의 하위 제품분류가 없는 경우
			On Error Resume next
			db.begintrans
			
			
			'코드관리 삭제
			SQL = "DELETE FROM TB_CODE WHERE CODEGROUP = '" &sCodegroup& "' AND CODE = '" &sCode& "'"
		
			db.execute(SQL)
			
			if db.Errors.count = 0 then
				flag = "Y"
				'LogWrite "SQL="&SQL, "Code_InsUpDel.asp", "/Setup/Code/"
			else
				flag = "N"
			end if
			
			if flag = "Y" then
				db.CommitTrans
				pageUrl = "code_list.asp?curPage=" & curPage & "&sCodegroup=" & sCodegroup & "&sGroupname=" & sGroupname
				Call MsgGoUrl( "정상적으로 삭제되었습니다.",pageUrl)
			else
				db.RollBackTrans
				'LogWrite "ERROR_SQL="&SQL, "Code_InsUpDel.asp", "/Setup/Code/"
				Call UrlBack("삭제 중 에러가 발생했습니다.\n\n다시 시도해 주세요")
			end if
		else										'제품분류의 하위 제품분류가 있는 경우
			Call UrlBack("하위 제품분류가 있는 코드 입니다.\n\n제품분류 유형관리에서 삭제 처리 하시기 바랍니다.")
		end if
	else											'코드 그룹이 A01(제품분류)가 아닌 경우	
		SQL = "DELETE FROM TB_CODE WHERE CODEGROUP = '" &sCodegroup& "' AND CODE = '" &sCode& "'"
		
		db.execute(SQL)
		
		if db.Errors.count = 0 then
			pageUrl = "code_list.asp?curPage=" & curPage & "&sCodegroup=" & sCodegroup & "&sGroupname=" & sGroupname
			'LogWrite "SQL="&SQL, "Code_InsUpDel.asp", "/Setup/Code/"
			Call MsgGoUrl( "정상적으로 삭제되었습니다.",pageUrl)
		else
			'LogWrite "ERROR_SQL="&SQL, "Code_InsUpDel.asp", "/Setup/Code/"
			Call UrlBack("삭제 중 에러가 발생했습니다.\n\n다시 시도해 주세요")
		end if
	end if
	
		
end Select



%>