<!-- #include virtual="/Include/Top_PopUp.asp" -->

<%
	'1. 파라미터
	seq = Request("seq")
	JobGb = Request("JobGb")

	curPage = request("curPage")
	FromDate = Trim(Request("FromDate"))
	ToDate = Trim(Request("ToDate"))
	sProcessYN = Trim(Request("sProcessYN"))
	whereCD2 = Trim(Request("whereCD2"))
	cboClassA = Request("cboClassA")

	IF curPage = "" THEN curPage = 1 END If
	pageWhere = "curPage=" & curPage & "&FromDate=" & FromDate & "&ToDate=" & ToDate
	
	'IF sProcessYN <> "" THEN	'처리여부 sProcessGB
		pageWhere = pageWhere & "&sProcessYN=" & sProcessYN
	'END If
	pageWhere = pageWhere & "&whereCD2=" & whereCD2
	pageWhere = pageWhere & "&cboClassA=" & cboClassA

	

	If JobGb = "U" Then

		INCODE = SESSION("SS_LoginID")
		sSelect1 = Request("Select1")
		sSelect2 = Request("sSelect2")
		sSelect3 = Request("sSelect3")
		sMemo = Request("sMEMO")

		SQL = "	UPDATE	TB_CALLBACK		SET		PROCESSGB	=	'" & sSelect1 & "'"
		If sSelect2 <> "" then
			SQL = SQL & "	,	NONPROCESSGB	=	'" & sSelect2 & "'"
			SQL = SQL & "	,	ProcessTime		=  getdate(),		ProcessCode	=	'" & INCODE & "'"
		ElseIf sSelect3 <> "" then
			SQL = SQL & "	,	NONPROCESSGB	=	'" & sSelect3 & "'"
			SQL = SQL & "	,	ProcessTime		=  getdate(),		ProcessCode	=	'" & INCODE & "'"
		Else
			SQL = SQL & "	,	NONPROCESSGB	=	NULL"
			If sSelect1 = "C" Then
				SQL = SQL & "	,	ProcessTime		= getdate(),		ProcessCode	=	'" & INCODE & "'"			
			Else
				SQL = SQL & "	,	ProcessTime		= NULL		"
			End if
		End If
		SQL = SQL & "		,	Memo				= '" & sMemo & "'"
		SQL = SQL & "	WHERE	idx = " & seq

		db.beginTrans
		db.execute(SQL)	

		if db.Errors.count = 0 then
			db.CommitTrans
			
			url = "CallBack.asp?" & pageWhere
			
			Response.Write ("<script>alert('정상적으로 처리되었습니다!');parent.location.href='"&url&"';parent.HddnPOPLayer();</script>")	
			Response.End
		else
			db.RollBackTrans
			Call UrlBack("저장중 에러가 발생했습니다.\n\n다시 시도해 주세요")
		end if

	End if

	SQL = "		SELECT	PROCESSGB, 	  NONPROCESSGB,	MEMO	FROM	TB_CALLBACK"
	SQL = SQL & "	WHERE	idx = " & seq

	set Rs = db.execute(SQL)

	If Rs.EOF THEN
		Select3Disabled = "Disabled"
		Select2Disabled = "Disabled"
	Else
		If RS("PROCESSGB") = "A" Then
			'접수
			sSelect2 =""
			sSelect3 =""
			Select3Disabled = "Disabled"
			Select2Disabled = "Disabled"
		ElseIf RS("PROCESSGB") = "B" Then
			'처리중
			sSelect3 =""
			sSelect2 =RS("NONPROCESSGB")
			Select2Disabled = ""
			Select3Disabled = "Disabled"
		ElseIf RS("PROCESSGB") = "C" Then
			'처리완료
			sSelect2 =""
			sSelect3 =""
			Select3Disabled = "Disabled"
			Select2Disabled = "Disabled"
		ElseIf RS("PROCESSGB") = "D" Then
			'처리불가
			sSelect2 =""
			sSelect3 =RS("NONPROCESSGB")
			Select3Disabled = ""
			Select2Disabled = "Disabled"
		End If
		sSelect1 =  RS("PROCESSGB")
		sMemo = RS("MEMO")
	End IF

%>

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
							<%=db_getTBCodeSelect("Z15", sSelect2, "N")%>
						</select>
					</td>
				</tr>
				<tr>
					<td bgcolor="#EEF6FF" width="30%" class="TDCont">처리불가사유</td>
					<td bgcolor="#FFFFFF">
						<input value="<%=sSelect3%>" name="sSelect3" readonly type="hidden">	
						<select name="Select3" size="1" class="ComboFFFCE7" <%=Select3Disabled%> onchange="javascript:fn_SetResult('3');">
							<option value="">처리불가사유선택</option>
							<%=db_getTBCodeSelect("Z16", sSelect3, "N")%>
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