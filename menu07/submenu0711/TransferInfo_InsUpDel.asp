<!-- #include virtual="/include/CacheNo.asp" -->
<!-- #include virtual="/include/common.asp" -->
<%
guboon = Request("guboon")								'저장/수정/삭제 FLAG
curPage = Request("curPage")							'현재페이지
seqno = Request("seqno")					'프로그램코드

DNIS = Trim(Request("txtDNIS"))		'입력/수정 폼값(프로그램코드)
Transferno = Trim(Request("txtTransferno"))	'입력/수정 폼값(프로그램명)
StartTime = ConvertString(Request("txtStartTime"))
EndTime = ConvertString(Request("txtEndTime"))
userid = ConvertString(Request("userid"))

If Request("chkMon") = "월요일" Then
	Mon = "1"
else
	Mon = "0"
End if	
	
If Request("chkTue") = "화요일" Then
	Tue = "1"
else
	Tue = "0"
End if
If Request("chkWed") = "수요일" Then
	Wed = "1"
else
	Wed = "0"
End if		
If Request("chkThu") = "목요일" Then
	Thu = "1"
else
	Thu = "0"
End if
If Request("chkFri") = "금요일" Then
	Fri = "1"
else
	Fri = "0"
End if	
If Request("chkSta") = "토요일" Then
	Sta = "1"
else
	Sta = "0"
End if		
If Request("chkSun") = "일요일" Then
	Sun = "1"
else
	Sun = "0"
End if		
If Request("chkHoliday") = "휴일" Then
	Holiday = "1"
else
	Holiday = "0"
End if		
If Request("chkUseyn") = "사용" Then
	USEYN = "1"
else
	USEYN = "0"
End if

select case ucase(guboon)
case "DEL"
	Dim SQL
	On Error Resume next
	dbcon.begintrans
	SQL ="DELETE TB_TransferInfo  WHERE seqno = '" & seqno & "'"
	db.execute(SQL)

	if db.Errors.count = 0 then
		'LogWrite "SQL="&SQL, "Program_InsUpDel.asp", "/Setup/Program/"
		db.CommitTrans
		  	
	%>	
		<script language="javascript">
			alert("정상적으로 삭제되었습니다.");	
			location.href = "TransferInfo.asp?curPage=<%=curPage%>";
		</script>	
	<%		
	else
		'LogWrite "ERROR_SQL="&SQL, "Program_InsUpDel.asp", "/Setup/Program/"
		db.RollBackTrans
		response.write("<script language=""javascript"">")&vbcr
		response.write("alert(""삭제중 에러가 발생했습니다.\n다시 시도해 주세요."");")&vbcr
		response.write("history.back();")&vbcr
		response.write("</script>")&vbcr
	end if

case "INS"
	INCODE = SESSION("SS_LoginID")
	
			
		On Error Resume next
		db.begintrans
		SQL1 ="Insert Into TB_TransferInfo (DNIS,StartTime,EndTime,Transferno,mon,tue,wed,thu,fri,sta,sun,holiday,useyn,userid) values "
		SQL1 = SQL1 & "('" & dnis & "','" & starttime & "','" & endtime & "','" & transferno & "','" & mon & "','" & tue & "','" & wed & "','" & thu & "','" & fri & "','" & sta & "','" & sun & "','" & holiday &"','" & USEYN & "','" & userid & "')"
		db.execute(SQL1)

		if db.Errors.count = 0 then
			LogWrite "SQL1="&SQL1, "Transfer_InsUpDel.asp", "/menu07/submenu0711/"
			db.CommitTrans
		%>
			<script language="javascript">
				alert("정상적으로 등록되었습니다.");
				location.href = "TransferInfo.asp?curPage=<%=curPage%>";
			</script>	
		<%	  
		else
			LogWrite "ERROR_SQL1="&SQL1, "Transfer_InsUpDel.asp", "/menu07/submenu0711/"
			db.RollBackTrans
			response.write("<script language=""javascript"">")&vbcr
			response.write("alert(""저장중 에러가 발생했습니다.\n다시 시도해 주세요."");")&vbcr
			response.write("history.back();")&vbcr
			response.write("</script>")&vbcr
		end if

	
case "UP"
	MOCODE = SESSION("SS_LoginID")
	
	On Error Resume next
	db.begintrans
		
	SQL2 = "update TB_TransferInfo Set DNIS ='" & dnis & "',StartTime = '" & StartTime & "',EndTime = '" & EndTime & "'"
	SQL2 = SQL2 & ",Transferno = '" & Transferno & "',mon = '" & mon & "',tue = '" & tue & "'"
	SQL2 = SQL2 & ",wed = '" & wed & "',thu = '" & thu & "',fri = '" & fri & "',sta = '" & sta & "',sun = '" & sun & "',holiday = '" & holiday & "',USEYN = '" & USEYN & "'"
	SQL2 = SQL2 & ",userid = '" & userid &"'"
	SQL2 = SQL2 & "  WHERE seqno = '" & seqno & "'"
	
	db.execute(SQL2)
	if db.Errors.count = 0 then
		LogWrite "SQL2="&SQL2, "Transfer_InsUpDel.asp", "/menu07/submenu0711/"
		db.CommitTrans
%>
		<script language="javascript">
			alert("정상적으로 수정되었습니다.");
			location.href = "TransferInfo.asp?curPage=<%=curPage%>";
		</script>	
<%	
	else
		LogWrite "ERROR_SQL2="&SQL2, "Transfer_InsUpDel.asp", "/menu07/submenu0711/"
		db.RollBackTrans
		response.write("<script language=""javascript"">")&vbcr
		response.write("alert(""저장중 에러가 발생했습니다.\n다시 시도해 주세요."");")&vbcr
		response.write("history.back();")&vbcr
		response.write("</script>")&vbcr
	end if
				
end select
%>