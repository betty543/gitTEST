<!-- #include virtual="/Include/Top_Frame.asp" -->
<%
	'####### 파라미터 ##################################################################################
	frmTYPE = Trim(Request("frmTYPE"))
	fn = Trim(Request("fn"))


	'####### 디버깅 코드 ###############################################################################
	Response.Write("frmTYPE=" &frmTYPE& "<br>")
	Response.Write("fn=" &fn& "<br>")
	
%>
<%
	'####### 파일삭제 ##################################################################################
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