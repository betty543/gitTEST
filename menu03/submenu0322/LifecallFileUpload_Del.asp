<!-- #include virtual="/Include/Top_Frame.asp" -->
<%
	'####### �Ķ���� ##################################################################################
	frmTYPE = Trim(Request("frmTYPE"))
	fn = Trim(Request("fn"))


	'####### ����� �ڵ� ###############################################################################
	Response.Write("frmTYPE=" &frmTYPE& "<br>")
	Response.Write("fn=" &fn& "<br>")
	
%>
<%
	'####### ���ϻ��� ##################################################################################
	SET FSO = Server.CreateObject("Scripting.FileSystemObject")
		File1 = Server.MapPath("\Upload\Lifecall")&"\"&fn

	IF FSO.FileExists(File1) THEN FSO.DeleteFile(File1) END IF
%>

<script language="javascript">
<!--
	parent.document.all.FILENAME1.value = "";
	parent.document.getElementById('txtFILENAME1').innerHTML = "";
//-->
</script>


<!-- #include virtual="/Include/Bottom.asp" -->