<!-- #include virtual="/include/CacheNo.asp" -->
<!-- #include virtual="/include/common.asp" -->


<%
	guboon = Request("guboon")					'저장/수정/삭제 FLAG
	sSec_Group = Request("Sec_group")	'보안그룹 코드
	sPorg_Code = Request("Prog_Code")	'프로그램 코드

select case ucase(guboon)


' 등록
case "INS"
	INCODE = SESSION("SS_LoginID")
	
	SQL = "select * from TB_SECGROUP where SEC_GROUP = '" & sSec_Group & "' and PROG_CODE = '" & sPorg_Code & "'"
	Set Rs = db.execute(SQL)
	
	On Error Resume next
	db.begintrans
	
	If Rs.Eof Or Rs.bof Then
		'코드관리 입력
		SQL ="INSERT INTO TB_SECGROUP (SEC_GROUP,PROG_CODE,INCODE) VALUES "
		SQL = SQL & "('" &sSec_Group& "','" &sPorg_Code& "','" & INCODE & "')"
		
		db.execute(SQL)

		if db.Errors.count = 0 then
			db.CommitTrans
			
			pageUrl = "group_program_list.asp?Sec_group="&sSec_Group
			Response.Write("<script>parent.ifr_Available.location.href = 'group_program_available.asp?sSec_group="&sSec_Group&"';</script>")
			Call MsgGoUrl( "정상적으로 등록되었습니다.",pageUrl)
		else
			db.RollBackTrans
			'LogWrite "ERROR_SQL="&SQL, "Group_Program_InsUpDel.asp", ""
			Call UrlBack("저장중 에러가 발생했습니다.\n\n다시 시도해 주세요")
		end if
	Else
		Call UrlBack("이미 등록된 프로그램 입니다.")
	end If
	
	Rs.Close
	set Rs = NOTHING	

' 삭제
case "DEL"
	
	SQL = "DELETE FROM TB_SECGROUP WHERE SEC_GROUP = '" &sSec_Group& "' AND PROG_CODE = '" &sPorg_Code& "'"
	
	db.execute(SQL)
	
	if db.Errors.count = 0 then
		pageUrl = "group_program_available.asp?sSec_group="&sSec_Group
		Call MsgGoUrl( "정상적으로 삭제되었습니다.",pageUrl)
	else
		'LogWrite "ERROR_SQL="&SQL, "Group_Program_InsUpDel.asp", ""
		Call UrlBack("삭제 중 에러가 발생했습니다.\n\n다시 시도해 주세요")
	end if
		
end Select



%>