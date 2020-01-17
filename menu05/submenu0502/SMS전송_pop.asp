<!-- #include virtual="/Include/Top_PopUp.asp" -->


<script>
<!--//
	function selectOK(){
		//alert(arg1+","+ arg2);
		parent.HddnPOPLayer();
	}

//-->
</script>
<body leftmargin="0" marginwidth="0" topmargin="0" marginheight="0">
<table width="700" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">
<form method="post" name="inUpFrm" action="<%=currentURL%>">
<input value="" name="JobGb" readonly type="hidden">	
<input value="<%=seq%>" name="Seq" readonly type="hidden">
<input value="<%=curPage%>"  name="curPage" readonly type="hidden">	
<input value="<%=FromDate%>"  name="FromDate" readonly type="hidden">
<input value="<%=ToDate%>"  name="ToDate" readonly type="hidden">
<input value="<%=sProcessYN%>"  name="sProcessYN" readonly type="hidden">
<input value="<%=whereCD2%>"  name="whereCD2" readonly type="hidden">
<input value="<%=cboClassA%>"  name="cboClassA" readonly type="hidden">
	<tr><td bgcolor="#FDE6F3" class="FBlk TDCont">◈ <b>문자전송</b></td></tr>
	<tr>
		<td bgcolor="#FFFFFF">
			<table width="100%" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">
				<tr>
					<td  bgcolor="#FFFFFF" width="100%" class="TDCont"><img src="/Images/문자발송.jpg" style="cursor:hand;" align="absmiddle"></td>
				</tr>
			</table>
		</td>
	</tr>
</form>
</table>


<script>
function fn_inup(form)
{
	form.JobGb.value = "U";
	form.submit();
}
function fn_SetResult(arg){
	if (arg =='1')
	{
		if (document.inUpFrm.Select1.value =='A' || document.inUpFrm.Select1.value =='C')
		{
			document.inUpFrm.Select3.disabled = true;
			document.inUpFrm.Select2.disabled = true;
			document.inUpFrm.sSelect2.value ="";
			document.inUpFrm.sSelect3.value ="";
		}
		if (document.inUpFrm.Select1.value =='B')
		{
			document.inUpFrm.Select3.disabled = true;
			document.inUpFrm.Select2.disabled = false;
			document.inUpFrm.sSelect2.value =document.inUpFrm.Select2.value;
			document.inUpFrm.sSelect3.value ="";
		}
		if (document.inUpFrm.Select1.value =='D')
		{
			document.inUpFrm.Select3.disabled = false;
			document.inUpFrm.Select2.disabled = true;
			document.inUpFrm.sSelect3.value =document.inUpFrm.Select3.value;
			document.inUpFrm.sSelect2.value ="";
		}
	}
	if (arg =='2')
	{
		document.inUpFrm.sSelect2.value =document.inUpFrm.Select2.value;
	}
	if (arg =='3')
	{
		document.inUpFrm.sSelect3.value =document.inUpFrm.Select3.value;
	}
}
</script>

<!-- #include virtual="/Include/Bottom_PopUp.asp" -->