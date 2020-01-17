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
<table width="490" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">
<form method="post" name="inUpFrm" action="<%=currentURL%>">
<input value="" name="JobGb" readonly type="hidden">	
<input value="<%=seq%>" name="Seq" readonly type="hidden">
<input value="<%=curPage%>"  name="curPage" readonly type="hidden">	
<input value="<%=FromDate%>"  name="FromDate" readonly type="hidden">
<input value="<%=ToDate%>"  name="ToDate" readonly type="hidden">
<input value="<%=sProcessYN%>"  name="sProcessYN" readonly type="hidden">
<input value="<%=whereCD2%>"  name="whereCD2" readonly type="hidden">
<input value="<%=cboClassA%>"  name="cboClassA" readonly type="hidden">
	<tr><td bgcolor="#FDE6F3" class="FBlk TDCont">◈ <b>콜백 처리 상태</b></td></tr>
	<tr>
		<td bgcolor="#FFFFFF">
			<table width="100%" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">
				<tr>
					<td  bgcolor="#EEF6FF" width="30%" class="TDCont">처리여부</td>
					<td bgcolor="#FFFFFF">
						<select name="Select1" size="1" class="ComboFFFCE7"  onchange="javascript:fn_SetResult('1');" >
							<option value="">처리여부선택</option>
							<%=db_getTBCodeSelect("Z14", sSelect1, "N")%>
						</select>
					</td>
				</tr>
				<tr>
					<td  bgcolor="#EEF6FF" width="30%" class="TDCont">처리중사유</td>
					<td bgcolor="#FFFFFF">
						<input value="<%=sSelect2%>" name="sSelect2" readonly type="hidden">		
						<select name="Select2" size="1" class="ComboFFFCE7" <%=Select2Disabled%> onchange="javascript:fn_SetResult('2');">
							<option value="">처리중사유선택</option>
							<%=db_getTBCodeSelect("Z16", sSelect2, "N")%>
						</select>
					</td>
				</tr>
				<tr>
					<td bgcolor="#EEF6FF" width="30%" class="TDCont">처리불가사유</td>
					<td bgcolor="#FFFFFF">
						<input value="<%=sSelect3%>" name="sSelect3" readonly type="hidden">	
						<select name="Select3" size="1" class="ComboFFFCE7" <%=Select3Disabled%> onchange="javascript:fn_SetResult('3');">
							<option value="">처리불가사유선택</option>
							<%=db_getTBCodeSelect("Z15", sSelect3, "N")%>
						</select>
					</td>
				</tr>
				<tr>
					<td  bgcolor="#EEF6FF" width="30%" class="TDCont">메모</td>
					<td bgcolor="#FFFFFF"><textarea name="sMEMO" style="width:100%; height:60;" wrap="soft" class="TextareaInput"><%=sMemo%></textarea></td>
				</tr>

			</table>
		</td>
	</tr>
</form>
</table>
<table width="490" border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td align="right" height="35">
			<img src="/Images/Btn/BtnSubmit.gif" style="cursor:hand;" align="absmiddle" onClick="javascript:fn_inup(document.inUpFrm);">
			<img src="/Images/Btn/BtnClose.gif" style="cursor:hand;" align="absmiddle" onClick="javascript:parent.HddnPOPLayer();">
		</td>
	</tr>
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