<!-- #include virtual="/include/CacheNo.asp" -->
<!-- #include virtual="/include/common.asp" -->


<%

guboon = Request("guboon")						'����/����/���� FLAG
curPage = Request("curPage")
sCode = Request("sCode")							'�ڵ�
sCodegroup = Request("sCodegroup")		'����
sGroupname = Request("sGroupname")		'���и�

CODEGROUP = ConvertName(request("txtCodeGroup"))			'�Է�/��������(����)
GROUPNAME = ConvertName(request("txtGroupName"))			'�Է�/��������(���и�)
CODE = trim(ucase(ConvertName(request("txtCode"))))					'�Է�/��������(�ڵ�)
CODENAME =trim(request("txtCodeName"))	'�Է�/��������(�ڵ��)

MEMO = ConvertString(Request("txtMemo"))								'�Է�/��������(�޸�)
USEYN = Request("optUseYN")													'�Է�/��������(��뿩��)
SYSYN = Request("optSysYN")													'�Է�/��������(�ý���)

select case ucase(guboon)


' ���
case "INS"
	INCODE = SESSION("SS_LoginID")
	
	Set Rs = server.createObject("ADODB.Recordset")
	SQL = "select * from TB_CODE where CODEGROUP = '" & CODEGROUP & "' and CODE = '" & CODE & "'"
	Rs.open SQL,db
	
	On Error Resume next
	db.begintrans
	
	If Rs.Eof Or Rs.bof Then
		flag = "N"			'INSERT ��������
		'�ڵ���� �Է�
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
			Call MsgGoUrl( "���������� ��ϵǾ����ϴ�.",pageUrl)
		else
			db.RollBackTrans
			'LogWrite "ERROR_SQL="&SQL, "Code_InsUpDel.asp", "/Setup/Code/"
			Call UrlBack("������ ������ �߻��߽��ϴ�.\n\n�ٽ� �õ��� �ּ���")
		end if
	Else
		Call UrlBack("�̹� �����ϴ� �ڵ� �Դϴ�..\n\n�ٸ� �ڵ带 �Է��� �ּ���")
	end If
	
	Rs.Close
	set Rs = NOTHING


' ����
case "UP"
	MOCODE = SESSION("SS_LoginID")
	
	On Error Resume next
	db.begintrans
	
	flag = "N"			'UPDATE ��������
	
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
		Call MsgGoUrl( "���������� �����Ǿ����ϴ�.",pageUrl)
	else
		db.RollBackTrans
		LogWrite "ERROR_SQL="&SQL, "Code_InsUpDel.asp", "/Setup/Code/"
		Call UrlBack("������ ������ �߻��߽��ϴ�.\n\n�ٽ� �õ��� �ּ���")
	end if

' ����
case "DEL"
	flag = "N"				'DELETE ��������
	
	'�ش� �ڵ� ��뿩�� üũ(�ش� �ڵ� ���� �����Ұ�) - ������ ��뿩�� ���̺� Ȯ�� �� ó���� ��
	if sCodegroup = "A01" then		'�ڵ� �׷��� A01(��ǰ�з�)�� ���
		
		'��ǰ�з� ���� üũ(���� ��ǰ�з� ����� �����Ұ�)
		SQL = "SELECT * FROM TB_GOODBUNU WHERE ACLASS = '" &sCode& "' AND BCLASS IS NOT NULL"
		
		set Rs = db.execute(SQL)
			
		if Rs.EOF OR RS.BOF then	'��ǰ�з��� ���� ��ǰ�з��� ���� ���
			On Error Resume next
			db.begintrans
			
			
			'�ڵ���� ����
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
				Call MsgGoUrl( "���������� �����Ǿ����ϴ�.",pageUrl)
			else
				db.RollBackTrans
				'LogWrite "ERROR_SQL="&SQL, "Code_InsUpDel.asp", "/Setup/Code/"
				Call UrlBack("���� �� ������ �߻��߽��ϴ�.\n\n�ٽ� �õ��� �ּ���")
			end if
		else										'��ǰ�з��� ���� ��ǰ�з��� �ִ� ���
			Call UrlBack("���� ��ǰ�з��� �ִ� �ڵ� �Դϴ�.\n\n��ǰ�з� ������������ ���� ó�� �Ͻñ� �ٶ��ϴ�.")
		end if
	else											'�ڵ� �׷��� A01(��ǰ�з�)�� �ƴ� ���	
		SQL = "DELETE FROM TB_CODE WHERE CODEGROUP = '" &sCodegroup& "' AND CODE = '" &sCode& "'"
		
		db.execute(SQL)
		
		if db.Errors.count = 0 then
			pageUrl = "code_list.asp?curPage=" & curPage & "&sCodegroup=" & sCodegroup & "&sGroupname=" & sGroupname
			'LogWrite "SQL="&SQL, "Code_InsUpDel.asp", "/Setup/Code/"
			Call MsgGoUrl( "���������� �����Ǿ����ϴ�.",pageUrl)
		else
			'LogWrite "ERROR_SQL="&SQL, "Code_InsUpDel.asp", "/Setup/Code/"
			Call UrlBack("���� �� ������ �߻��߽��ϴ�.\n\n�ٽ� �õ��� �ּ���")
		end if
	end if
	
		
end Select



%>