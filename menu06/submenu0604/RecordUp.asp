<!-- #include virtual="/Include/Top_PopUp.asp" -->

<%
	'1. 파라미터
	SS_Login_Secgroup = SESSION("SS_Login_Secgroup")
	SS_Login_Grade = SESSION("SS_Login_Grade")
	SS_Login_CTIID = SESSION("SS_Login_CTIID")
	SS_Login_EXTNO = SESSION("SS_Login_EXTNO")
	SS_LoginID = SESSION("SS_LoginID")
	JobGb = Request("JobGb")
	
	If JobGb = "U" Then

		INCODE = SESSION("SS_LoginID")
		UserID = Request("UserID")
		RecordingCallKey = Request("RecordingCallKey")


		SQL = " SELECT * FROM TB_RecordingData "
		SQL = SQL & "	WHERE	RecordingCallKey = '" & RecordingCallKey & "'"
		set RsCode = db.execute(SQL)
		IF RsCode.Eof THEN

'substring(d.recordingfilename,32,254)
			SQL = "	INSERT INTO TB_RecordingData ("
			SQL = SQL & "	RecordingCallKey, CallId, RecordStartTime, RecordEndTime, RecordDuration, RemoteId1, UserId, recordfilename,dnis, RemoteId2 )"

			SQL = SQL & "	SELECT	'"& RecordingCallKey  & "','"& left(RecordingCallKey ,10) & "', convert(char(19),dateadd(hour,9,d.recordingdate),121), convert(char(19),dateadd(hour,9,d.recordingdate),121),"
			SQL = SQL & "  d.RecordingFileSize / 10.75, c.ANI, '" & UserID  & "','none', c.dnis, c.ANI"
		
		SQL = SQL & "	FROM	I3_IC.dbo.RecordingCall c, I3_IC.dbo.RecordingData d"
		SQL = SQL & "	WHERE	c.RecordingID = d.RecordingID "

				SQL = SQL & "	AND		recordedCallidKey = '" & RecordingCallKey & "'"

		ELSE

			SQL = "	UPDATE	TB_RecordingData		SET		UserID	=	'" & UserID & "'"
			SQL = SQL & "	WHERE	RecordingCallKey = '" & RecordingCallKey & "'"
		END IF



		db.beginTrans
		db.execute(SQL)	

		if db.Errors.count = 0 then
			db.CommitTrans	
			
			Response.Write ("<script>alert('정상적으로 처리되었습니다!');parent.fn_Search();parent.HddnPOPLayer();</script>")	
			Response.End
		else
			db.RollBackTrans
			Call UrlBack("저장중 에러가 발생했습니다.\n\n다시 시도해 주세요")
		end if
	else
		RecordingCallKey = Request("RecordingCallKey")
		UserID = Request("UserID")
	End if

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
<table width="300" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">
<form method="post" name="inUpFrm" action="<%=currentURL%>">
<input value="" name="JobGb" readonly type="hidden">	
<input value="<%=RecordingCallKey%>" name="RecordingCallKey" readonly type="hidden">
	<tr><td bgcolor="#FDE6F3" class="FBlk TDCont">◈ <b>녹취 상담관 변경</b></td></tr>
	<tr>
		<td bgcolor="#FFFFFF">
			<table width="100%" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">
				<tr>
					<td  bgcolor="#EEF6FF" width="30%" class="TDCont">처리여부</td>
					<td bgcolor="#FFFFFF">
						<%
							'======= 상담원 가져오기 ==================================================
							SqlCode = "SELECT USERID, USERNAME FROM TB_USERINFO"
							SqlCode = SqlCode& " WHERE USEYN='Y' "
							SqlCode = SqlCode& " AND SECGROUP = 'A'"
							if SS_Login_Grade <> "A" then
								SqlCode = SqlCode& "	AND GRADE = '"&SS_Login_Grade&"'"
							end if
							if SS_Login_Secgroup = "A" then	'상담관일때는 내것만
								SqlCode = SqlCode& "	AND USERID = '" &SS_LoginID&"'"
							end if
							
							SqlCode = SqlCode& " ORDER BY USEYN DESC, GRADE ASC, USERNAME ASC"
							set RsCode = db.execute(SqlCode)
						%>
						<select name="UserID" size="1" class="ComboFFFCE7">
							<option value="">선택</option>
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("USERID")
										CODENAME = RsCode("USERNAME")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &UserID& "")%>
							<%
									RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
							%>
						<%
							'======= 상담원 가져오기 ==================================================
							SqlCode = "SELECT CTIID, USERNAME FROM TB_USERINFO"
							SqlCode = SqlCode& " WHERE USEYN='N'  and	outdate >= '"&DateAdd("d",1,DateAdd("m",-1,Date())) &"'"
							if SS_Login_Grade <> "A" then
								SqlCode = SqlCode& "	AND GRADE = '"&SS_Login_Grade&"'"
							end if
							if SS_Login_Secgroup = "A" then	'상담관일때는 내것만
								SqlCode = SqlCode& "	AND USERID = '" &SS_LoginID&"'"
							end if

							SqlCode = SqlCode& " ORDER BY USEYN DESC, GRADE ASC, USERNAME ASC"
							set RsCode = db.execute(SqlCode)

								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("CTIID")
										CODENAME = "[퇴직]"&RsCode("USERNAME")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &UserID& "")%>
							<%
									RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
							%>
						</select>
					</td>
				</tr>

			</table>
		</td>
	</tr>
</form>
</table>
<table width="300" border="0" cellspacing="0" cellpadding="0">
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
</script>

<!-- #include virtual="/Include/Bottom_PopUp.asp" -->