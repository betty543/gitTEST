<!-- #include virtual="/include/CacheNo.asp" -->
<!-- #include virtual="/include/common.asp" -->


<%
	guboon = Request("guboon")					'����/����/���� FLAG
	sSec_Group = Request("Sec_group")	'���ȱ׷� �ڵ�
	sPorg_Code = Request("Prog_Code")	'���α׷� �ڵ�

select case ucase(guboon)


' ���
case "INS"
	INCODE = SESSION("SS_LoginID")
	
	SQL = "select * from TB_SECGROUP where SEC_GROUP = '" & sSec_Group & "' and PROG_CODE = '" & sPorg_Code & "'"
	Set Rs = db.execute(SQL)
	
	On Error Resume next
	db.begintrans
	
	If Rs.Eof Or Rs.bof Then
		'�ڵ���� �Է�
		SQL ="INSERT INTO TB_SECGROUP (SEC_GROUP,PROG_CODE,INCODE) VALUES "
		SQL = SQL & "('" &sSec_Group& "','" &sPorg_Code& "','" & INCODE & "')"
		
		db.execute(SQL)

		if db.Errors.count = 0 then
			db.CommitTrans
			
			pageUrl = "group_program_list.asp?Sec_group="&sSec_Group
			Response.Write("<script>parent.ifr_Available.location.href = 'group_program_available.asp?sSec_group="&sSec_Group&"';</script>")
			Call MsgGoUrl( "���������� ��ϵǾ����ϴ�.",pageUrl)
		else
			db.RollBackTrans
			'LogWrite "ERROR_SQL="&SQL, "Group_Program_InsUpDel.asp", ""
			Call UrlBack("������ ������ �߻��߽��ϴ�.\n\n�ٽ� �õ��� �ּ���")
		end if
	Else
		Call UrlBack("�̹� ��ϵ� ���α׷� �Դϴ�.")
	end If
	
	Rs.Close
	set Rs = NOTHING	

' ����
case "DEL"
	
	SQL = "DELETE FROM TB_SECGROUP WHERE SEC_GROUP = '" &sSec_Group& "' AND PROG_CODE = '" &sPorg_Code& "'"
	
	db.execute(SQL)
	
	if db.Errors.count = 0 then
		pageUrl = "group_program_available.asp?sSec_group="&sSec_Group
		Call MsgGoUrl( "���������� �����Ǿ����ϴ�.",pageUrl)
	else
		'LogWrite "ERROR_SQL="&SQL, "Group_Program_InsUpDel.asp", ""
		Call UrlBack("���� �� ������ �߻��߽��ϴ�.\n\n�ٽ� �õ��� �ּ���")
	end if
		
end Select



%>