<!-- #include virtual="/Include/Top_PopUp.asp" -->

<%
	'####### 폼값 받기 #################################################################################
	fileCNT = trim(Request("fileCNT"))
	frmTYPE = trim(Request("frmTYPE"))
	frmTYPE = "txtFILENAME1"
	
	'####### 디버깅 코드 ###############################################################################
	'Response.Write("fileCNT=" &fileCNT& "<br>")
	'Response.Write("frmTYPE=" &frmTYPE& "<br>")
%>

<script>
<!--
	function fn_inup(){
		if(!FieldChk(inUpFrm.aFilename,"첨부파일")) return false;
	}
//-->
</script>

<br>
<table border="0" width="95%" cellpadding="0" cellspacing="3" bgcolor="#EFECE5" align="center">
	<tr bgcolor="#FFFFFF">
		<td style="border:#E1DED6 solid 1px" align="center">
			<table width="100%" cellpadding="0" cellspacing="0" border="0">
				<form name="inUpFrm" method="post" action="LifecallFileUpload_InsUp.asp" onsubmit="return fn_inup(this);" encType="multipart/form-data">
				<input type="hidden" name="isType" value="INS">
				<input type="hidden" name="frmTYPE" value="<%=frmTYPE%>">
				<tr><td height="30" align="center">첨부파일 선택 : <input type="file" size="30" name="aFilename" onblur="this.style.backgroundColor='#FFFFFF'" onfocus="this.style.backgroundColor='#DDE0F1'"></td></tr>
				<tr><td class="TRLine"></td></tr>
				<tr><td bgcolor="#F9F8F4" height="40" align="center"><input type="image" src="/Images/Btn/BtnFileUpload.gif" name="BtnOK" class="None" align="absmiddle"></td></tr>
				</form>
			</table>
		</td>
	</tr>
</table>

<table border="0" cellspacing="0" width="95%" align="center">
	<tr height="30">
		<td align="right">
			<img src="/Images/Btn/BtnClose.gif" style="cursor:hand;" align="absmiddle" onClick="parent.HddnPOPLayer();">
		</td>
	</tr>
</table>

<!-- #include virtual="/Include/Bottom_PopUp.asp" -->