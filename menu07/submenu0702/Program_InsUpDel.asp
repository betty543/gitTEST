<!-- #include virtual="/include/CacheNo.asp" -->
<!-- #include virtual="/include/common.asp" -->
<%
guboon = Request("guboon")								'����/����/���� FLAG
curPage = Request("curPage")							'����������
w_PARENT_ID = Request("w_PARENT_ID")
db_PARENT_ID = Trim(Request("PARENT_ID"))		'�Է�/���� ����(���α׷��ڵ�)
db_PROGRAM_ID = Trim(Request("PROGRAM_ID"))	'�Է�/���� ����(���α׷���)
db_PROGRAM_IDX = Trim(Request("PROGRAM_IDX"))	'�Է�/���� ����(���α׷���)
db_PROGRAM_NM = Trim(Request("PROGRAM_NM"))	'�Է�/���� ����(���α׷���)
db_NAVIGATE_URL = ConvertString(Request("NAVIGATE_URL"))
db_PROGRAM_INSERT = Trim(Request("PROGRAM_INSERT"))	'�Է�/���� ����(���α׷���)
db_PROGRAM_UPDATE = Trim(Request("PROGRAM_UPDATE"))	'�Է�/���� ����(���α׷���)
db_PROGRAM_DELETE = Trim(Request("PROGRAM_DELETE"))	'�Է�/���� ����(���α׷���)
db_PROGRAM_PRINT = Trim(Request("PROGRAM_PRINT"))	'�Է�/���� ����(���α׷���)
db_USE_YN = Trim(Request("USE_YN"))	'�Է�/���� ����(���α׷���)

	
If Request("db_PROGRAM_INSERT") = "���" Then
	INSERTYN = "Y"
else
	INSERTYN = "N"
End if
If Request("db_PROGRAM_UPDATE") = "����" Then
	UPDATEYN = "Y"
else
	UPDATEYN = "N"
End if		
If Request("db_PROGRAM_DELETE") = "����" Then
	DELETEYN = "Y"
else
	DELETEYN = "N"
End if
If Request("db_PROGRAM_PRINT") = "���" Then
	PRINTYN = "Y"
else
	PRINTYN = "N"
End if		
If Request("db_USE_YN") = "���" Then
	USEYN = "Y"
else
	USEYN = "N"
End If

Dim objCmd
Set objCmd = Server.CreateObject("ADODB.Command")

select case ucase(guboon)
case "DEL"
	Dim SQL
	On Error Resume next
	Casamiadb.begintrans



	with objCmd

		.ActiveConnection = Casamiadb
		.CommandText = "USP_PROGRAM_DEL"
		.CommandType = adCmdStoredProc

		.parameters.append(.CreateParameter("@V_COMPANY_ID",advarchar,adParamInput,50,COMPANY_ID))
		.parameters.append(.CreateParameter("@V_PROGRAM_ID",adInteger,adParamInput,,db_PROGRAM_ID))
		.Execute

	End with

	if Casamiadb.Errors.count = 0 then
		'LogWrite "SQL="&SQL, "Program_InsUpDel.asp", "/Setup/Program/"
		Casamiadb.CommitTrans
		  	
	%>	
		<script language="javascript">
			alert("���������� �����Ǿ����ϴ�.");	
			location.href = "Program_detail.asp?guboon=INS&curPage=<%=curPage%>";
			parent.ListFrame.location.href = "Program_List.asp?curPage=<%=curPage%>&w_PARENT_ID=<%=w_PARENT_ID%>";
		</script>	
	<%		
	else
		'LogWrite "ERROR_SQL="&SQL, "Program_InsUpDel.asp", "/Setup/Program/"
		Casamiadb.RollBackTrans
		response.write("<script language=""javascript"">")&vbcr
		response.write("alert(""������ ������ �߻��߽��ϴ�.\n�ٽ� �õ��� �ּ���."");")&vbcr
		response.write("history.back();")&vbcr
		response.write("</script>")&vbcr
	end if

case "INS"

	Casamiadb.begintrans

	with objCmd

		.ActiveConnection = Casamiadb
		.CommandText = "USP_PROGRAM_INS"
		.CommandType = adCmdStoredProc

		.parameters.append(.CreateParameter("@V_COMPANY_ID",advarchar,adParamInput,50,COMPANY_ID))
		.parameters.append(.CreateParameter("@V_PARENT_ID",adInteger,adParamInput,,db_PARENT_ID))
		.parameters.append(.CreateParameter("@V_PROGRAM_IDX",adInteger,adParamInput,,db_PROGRAM_IDX))
		.parameters.append(.CreateParameter("@V_PROGRAM_NM",advarchar,adParamInput,100,db_PROGRAM_NM))
		.parameters.append(.CreateParameter("@V_MENU_NM",advarchar,adParamInput,100,""))
		.parameters.append(.CreateParameter("@V_URL",advarchar,adParamInput,200,""))
		.parameters.append(.CreateParameter("@V_PARAMETER",advarchar,adParamInput,100,""))
		.parameters.append(.CreateParameter("@V_NAVIGATE_URL",advarchar,adParamInput,200,db_NAVIGATE_URL))
		.parameters.append(.CreateParameter("@V_PROGRAM_INSERT",advarchar,adParamInput,1,db_PROGRAM_INSERT))
		.parameters.append(.CreateParameter("@V_PROGRAM_UPDATE",advarchar,adParamInput,1,db_PROGRAM_UPDATE))
		.parameters.append(.CreateParameter("@V_PROGRAM_DELETE",advarchar,adParamInput,1,db_PROGRAM_DELETE))
		.parameters.append(.CreateParameter("@V_PROGRAM_PRINT",advarchar,adParamInput,1,db_PROGRAM_PRINT))
		.parameters.append(.CreateParameter("@V_PROGRAM_EXTEND1",advarchar,adParamInput,100,""))
		.parameters.append(.CreateParameter("@V_PROGRAM_EXTEND2",advarchar,adParamInput,100,""))
		.parameters.append(.CreateParameter("@V_PROGRAM_EXTEND3",advarchar,adParamInput,100,""))
		.parameters.append(.CreateParameter("@V_PROGRAM_EXTEND4",advarchar,adParamInput,100,""))
		.parameters.append(.CreateParameter("@V_PROGRAM_EXTEND5",advarchar,adParamInput,100,""))
		.parameters.append(.CreateParameter("@V_PROGRAM_EXTEND6",advarchar,adParamInput,100,""))
		.parameters.append(.CreateParameter("@V_PROGRAM_EXTEND7",advarchar,adParamInput,100,""))
		.parameters.append(.CreateParameter("@V_PROGRAM_EXTEND8",advarchar,adParamInput,100,""))
		.parameters.append(.CreateParameter("@V_PROGRAM_EXTEND9",advarchar,adParamInput,100,""))
		.parameters.append(.CreateParameter("@V_SYS_YN",advarchar,adParamInput,1,"Y"))
		.parameters.append(.CreateParameter("@V_MENU_YN",advarchar,adParamInput,1,"Y"))
		.parameters.append(.CreateParameter("@V_USE_YN",advarchar,adParamInput,1,db_USE_YN))
		.parameters.append(.CreateParameter("@V_DESCRIPTION",advarchar,adParamInput,1000,""))
		.parameters.append(.CreateParameter("@V_CREATOR_ID",adInteger,adParamInput,,SS_LOGIN_IDX))
		.parameters.append(.CreateParameter("@V_PROGRAM_ID",adInteger,adParamOutput))
		.Execute


	End with
	
	if Casamiadb.Errors.count = 0 then
		'LogWrite "SQL1="&SQL1, "Program_InsUpDel.asp", "/Setup/Program/"
		Casamiadb.CommitTrans
		%>
		<script language="javascript">
			alert("���������� ��ϵǾ����ϴ�.");
			location.href = "Program_detail.asp?guboon=Up&curPage=<%=curPage%>&w_PARENT_ID=<%=w_PARENT_ID%>&PROGRAM_ID=<%=db_PROGRAM_ID%>";
			parent.ListFrame.location.href = "Program_List.asp?curPage=<%=curPage%>&w_PARENT_ID=<%=w_PARENT_ID%>";
		</script>	
		<%	  
	else
		'LogWrite "ERROR_SQL="&SQL, "Program_InsUpDel.asp", "/Setup/Program/"
		Casamiadb.RollBackTrans
		response.write("<script language=""javascript"">")&vbcr
		response.write("alert(""������ ������ �߻��߽��ϴ�.\n�ٽ� �õ��� �ּ���."");")&vbcr
		response.write("history.back();")&vbcr
		response.write("</script>")&vbcr
	end if

	
case "UP"

	Casamiadb.begintrans

	with objCmd

		.ActiveConnection = Casamiadb
		.CommandText = "USP_PROGRAM_UPT"
		.CommandType = adCmdStoredProc

		.parameters.append(.CreateParameter("@V_COMPANY_ID",advarchar,adParamInput,50,COMPANY_ID))
		.parameters.append(.CreateParameter("@V_PROGRAM_ID",adInteger,adParamInput,,db_PROGRAM_ID))
		.parameters.append(.CreateParameter("@V_PROGRAM_IDX",adInteger,adParamInput,,db_PROGRAM_IDX))
		.parameters.append(.CreateParameter("@V_PROGRAM_NM",advarchar,adParamInput,100,db_PROGRAM_NM))
		.parameters.append(.CreateParameter("@V_MENU_NM",advarchar,adParamInput,100,""))
		.parameters.append(.CreateParameter("@V_URL",advarchar,adParamInput,200,""))
		.parameters.append(.CreateParameter("@V_PARAMETER",advarchar,adParamInput,100,""))
		.parameters.append(.CreateParameter("@V_NAVIGATE_URL",advarchar,adParamInput,200,db_NAVIGATE_URL))
		.parameters.append(.CreateParameter("@V_PROGRAM_INSERT",advarchar,adParamInput,1,db_PROGRAM_INSERT))
		.parameters.append(.CreateParameter("@V_PROGRAM_UPDATE",advarchar,adParamInput,1,db_PROGRAM_UPDATE))
		.parameters.append(.CreateParameter("@V_PROGRAM_DELETE",advarchar,adParamInput,1,db_PROGRAM_DELETE))
		.parameters.append(.CreateParameter("@V_PROGRAM_PRINT",advarchar,adParamInput,1,db_PROGRAM_PRINT))
		.parameters.append(.CreateParameter("@V_PROGRAM_EXTEND1",advarchar,adParamInput,100,""))
		.parameters.append(.CreateParameter("@V_PROGRAM_EXTEND2",advarchar,adParamInput,100,""))
		.parameters.append(.CreateParameter("@V_PROGRAM_EXTEND3",advarchar,adParamInput,100,""))
		.parameters.append(.CreateParameter("@V_PROGRAM_EXTEND4",advarchar,adParamInput,100,""))
		.parameters.append(.CreateParameter("@V_PROGRAM_EXTEND5",advarchar,adParamInput,100,""))
		.parameters.append(.CreateParameter("@V_PROGRAM_EXTEND6",advarchar,adParamInput,100,""))
		.parameters.append(.CreateParameter("@V_PROGRAM_EXTEND7",advarchar,adParamInput,100,""))
		.parameters.append(.CreateParameter("@V_PROGRAM_EXTEND8",advarchar,adParamInput,100,""))
		.parameters.append(.CreateParameter("@V_PROGRAM_EXTEND9",advarchar,adParamInput,100,""))
		.parameters.append(.CreateParameter("@V_SYS_YN",advarchar,adParamInput,1,"Y"))
		.parameters.append(.CreateParameter("@V_MENU_YN",advarchar,adParamInput,1,"Y"))
		.parameters.append(.CreateParameter("@V_USE_YN",advarchar,adParamInput,1,db_USE_YN))
		.parameters.append(.CreateParameter("@V_DESCRIPTION",advarchar,adParamInput,1000,""))
		.parameters.append(.CreateParameter("@V_UPDATOR_ID",adInteger,adParamInput,,SS_LOGIN_IDX))

		.Execute

	End with
	
	if Casamiadb.Errors.count = 0 then
		'LogWrite "SQL2="&SQL2, "Program_InsUpDel.asp", "/Setup/Program/"
		Casamiadb.CommitTrans
%>
		<script language="javascript">
			alert("���������� �����Ǿ����ϴ�.");
			location.href = "Program_detail.asp?guboon=Up&curPage=<%=curPage%>&w_PARENT_ID=<%=w_PARENT_ID%>&PROGRAM_ID=<%=PROGRAM_ID%>";
			parent.ListFrame.location.href = "Program_List.asp?curPage=<%=curPage%>&w_PARENT_ID=<%=w_PARENT_ID%>";
		</script>	
<%	
	else
		'LogWrite "ERROR_SQL="&SQL, "Program_InsUpDel.asp", "/Setup/Program/"
		Casamiadb.RollBackTrans
		response.write("<script language=""javascript"">")&vbcr
		response.write("alert(""������ ������ �߻��߽��ϴ�.\n�ٽ� �õ��� �ּ���."");")&vbcr
		response.write("history.back();")&vbcr
		response.write("</script>")&vbcr
	end if
				
end select
%>