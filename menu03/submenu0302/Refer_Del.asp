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

%>

<script language="javascript">
<!--
	parent.document.all.REFCNT.value = "1";
	parent.document.all.REFERJUBSEQ.value ="";	//�ڱ��ڽ��̵�.
	parent.document.getElementById('txtREFERJUBSEQ').innerHTML = "";
//-->
</script>


<!-- #include virtual="/Include/Bottom.asp" -->