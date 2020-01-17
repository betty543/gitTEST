<!-- #include virtual="/Include/Top2.asp" -->
<%
	On Error Resume next

	sJob = Request("Job")
	iIdx= Request("idx")
	ControlIdx = Request("ControlIdx")
	factPeoplenum = Request("factPeoplenum")
	receiptfactnum = Request("receiptfactnum")
	callid = Request("callid")
	recyn = Request("recyn")

	if sJob = "D" then

		SQL = "DELETE	FROM	armyinformix.dbo.contactlist where idx = " & iIdx
		db.execute(SQL)

		'통화 히스토리 찾기
		sql = " select substring(convert(char(19),contactdate,121),6,11) contacttime, * from	armyinformix.dbo.contactlist where factnum = '" & trim(receiptfactnum) & "' and peoplenum = '" & factPeoplenum & "' order by idx"
		SET Rs1 = db.execute(SQL)
		do until rs1.eof
			if db_History_6 = "" then
				db_History_6 = "<a href='##' a href='##' onClick=HistoryUpdate('"&ControlIdx&"','"&rs1("idx")&"','"&receiptfactnum&"','"&factPeoplenum&"');>"&rs1("contacttime")&"&nbsp;("&rs1("contacttelno")&")&nbsp;"&rs1("remark")&"</a>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"&rs1("contactdate")&" 통화 삭제' style='cursor:hand;' align='absmiddle' onClick=HistoryDel('"&ControlIdx&"','"&rs1("idx")&"','"&receiptfactnum&"','"&factPeoplenum&"');>"

				if rs1("callid") <> "" then
					if rs1("recordyn") = "Y" then

						SQL = "select d.RecordingFileName		FROM	I3_IC.dbo.RecordingCall c, I3_IC.dbo.RecordingData d 		WHERE	c.RecordingID = d.RecordingID 		AND	left(recordedcallidkey ,10) = '" & rs1("callid") & "'"

						SET Rs2 = db.execute(SQL)
						if Rs2.eof = false then
							db_RecordingFileName = rs2("RecordingFileName")

							sdb_RecordingFileName = "http://16.1.17.113:8080/"&mid(replace(db_RecordingFileName,"\","/"),27)
							sssdb_RecordingFileName = mid(replace(sdb_RecordingFileName,"\","/"),27)
							db_History_6 = db_History_6 & "&nbsp;"& CutString(sssdb_RecordingFileName, 40, "...") & "<img src='/Images/Comm/IconAlert.gif' align='absmiddle' style='cursor:hand;' onClick=""fn_Player('"&sdb_RecordingFileName&"');"" title='녹음내용 청취'>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"&rs1("contactdate")&" 녹취 삭제' style='cursor:hand;' align='absmiddle' onClick=RecDel('"&ControlIdx&"','"&rs1("callid")&"','N');>"

						end if
						Rs2.close
					else
						SQL = "select d.RecordingFileName		FROM	I3_IC.dbo.RecordingCall c, I3_IC.dbo.RecordingData d 		WHERE	c.RecordingID = d.RecordingID 		AND	left(recordedcallidkey ,10) = '" & rs1("callid") & "'"

						SET Rs2 = db.execute(SQL)

						if Rs2.eof = false then
							db_RecordingFileName = rs2("RecordingFileName")

							sdb_RecordingFileName = "http://16.1.17.113:8080/"&mid(replace(db_RecordingFileName,"\","/"),27)
							sssdb_RecordingFileName = mid(replace(sdb_RecordingFileName,"\","/"),27)
							db_History_6 = db_History_6 & "&nbsp;<a href='##' onClick=RecDel('"&ControlIdx&"','"&rs1("callid")&"','Y');>녹취포함</a>"
						end if

						Rs2.close
					end if
				end if
			else
				db_History_6 = db_History_6 & "<br>&nbsp;<a href='##' onClick=HistoryUpdate('"&ControlIdx&"','"&rs1("idx")&"','"&receiptfactnum&"','"&factPeoplenum&"');>"&rs1("contacttime")&"&nbsp;("&rs1("contacttelno")&")&nbsp;"&rs1("remark")&"</a>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"&rs1("contactdate")&" 통화 삭제' style='cursor:hand;' align='absmiddle' onClick=HistoryDel('"&ControlIdx&"','"&rs1("idx")&"','"&receiptfactnum&"','"&factPeoplenum&"');>"
				if rs1("callid") <> "" then
					if rs1("recordyn") = "Y" then

						SQL = "select d.RecordingFileName		FROM	I3_IC.dbo.RecordingCall c, I3_IC.dbo.RecordingData d 		WHERE	c.RecordingID = d.RecordingID 		AND	left(recordedcallidkey ,10) = '" & rs1("callid") & "'"

						SET Rs2 = db.execute(SQL)
						if Rs2.eof = false then
							db_RecordingFileName = rs2("RecordingFileName")

							sdb_RecordingFileName = "http://16.1.17.113:8080/"&mid(replace(db_RecordingFileName,"\","/"),27)
							sssdb_RecordingFileName = mid(replace(sdb_RecordingFileName,"\","/"),27)
							db_History_6 = db_History_6 & "&nbsp;"& CutString(sssdb_RecordingFileName, 40, "...") & "<img src='/Images/Comm/IconAlert.gif' align='absmiddle' style='cursor:hand;' onClick=""fn_Player('"&sdb_RecordingFileName&"');"" title='녹음내용 청취'>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"&rs1("contactdate")&" 녹취 삭제' style='cursor:hand;' align='absmiddle' onClick=RecDel('"&ControlIdx&"','"&rs1("callid")&"','N');>"

						end if
						Rs2.close
					else
						SQL = "select d.RecordingFileName		FROM	I3_IC.dbo.RecordingCall c, I3_IC.dbo.RecordingData d 		WHERE	c.RecordingID = d.RecordingID 		AND	left(recordedcallidkey ,10) = '" & rs1("callid") & "'"

						SET Rs2 = db.execute(SQL)

						if Rs2.eof = false then
							db_RecordingFileName = rs2("RecordingFileName")

							sdb_RecordingFileName = "http://16.1.17.113:8080/"&mid(replace(db_RecordingFileName,"\","/"),27)
							sssdb_RecordingFileName = mid(replace(sdb_RecordingFileName,"\","/"),27)
							db_History_6 = db_History_6 & "&nbsp;<a href='##' onClick=RecDel('"&ControlIdx&"','"&rs1("callid")&"','Y');>녹취포함</a>"
						end if

						Rs2.close
					end if
				end if
			end if
			rs1.movenext
		loop
		rs1.close

%>
	
	<script>
		parent.document.getElementById('History_<%=ControlIdx%>').innerHTML = "<%=db_History_6%>";
	</script>

<%

	elseif sJob = "U" then
		SQL = "SELECT *	FROM	armyinformix.dbo.contactlist where idx = " & iIdx
		SET Rs1 = db.execute(SQL)
		IF Rs1.eof = false then
			db_History_6 = rs1("Remark")
			db_RecYN = rs1("RecordYN")

%>
	
	<script>
		parent.document.getElementById('Remark1_<%=ControlIdx%>').value = "<%=db_History_6%>";
		eval("parent.ListForm.idx_<%=ControlIdx%>").value= "<%=iIdx%>";
		if ( "<%=db_RecYN%>" == "Y" )
			eval("parent.ListForm.RecYN_<%=ControlIdx%>").checked = "true";
		else
			eval("parent.ListForm.RecYN_<%=ControlIdx%>").checked = "false";
	</script>

<%
		end if
	elseif sJob = "I" then

		ContactTelNo = request("ContactTelNo")
		SQL = " insert into armyinformix.dbo.contactlist values ( '"&receiptfactnum&"','"&factPeoplenum&"',getdate(),'"&ContactTelNo&"','','"&callid&"','N')"
		db.execute(SQL)


		'통화 히스토리 찾기
		sql = " select substring(convert(char(19),contactdate,121),6,11) contacttime, * from	armyinformix.dbo.contactlist where factnum = '" & trim(receiptfactnum) & "' and peoplenum = '" & factPeoplenum & "' order by idx"
		SET Rs1 = db.execute(SQL)
		do until rs1.eof
			idx = rs1("idx")
			if db_History_6 = "" then
				db_History_6 = "<a href='##' a href='##' onClick=HistoryUpdate('"&ControlIdx&"','"&rs1("idx")&"','"&receiptfactnum&"','"&factPeoplenum&"');>"&rs1("contacttime")&"&nbsp;("&rs1("contacttelno")&")&nbsp;"&rs1("remark")&"</a>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"&rs1("contactdate")&" 통화 삭제' style='cursor:hand;' align='absmiddle' onClick=HistoryDel('"&ControlIdx&"','"&rs1("idx")&"','"&receiptfactnum&"','"&factPeoplenum&"');>"
			else
				db_History_6 = db_History_6 & "<br>&nbsp;<a href='##' onClick=HistoryUpdate('"&ControlIdx&"','"&rs1("idx")&"','"&receiptfactnum&"','"&factPeoplenum&"');>"&rs1("contacttime")&"&nbsp;("&rs1("contacttelno")&")&nbsp;"&rs1("remark")&"</a>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"&rs1("contactdate")&" 통화 삭제' style='cursor:hand;' align='absmiddle' onClick=HistoryDel('"&ControlIdx&"','"&rs1("idx")&"','"&receiptfactnum&"','"&factPeoplenum&"');>"
			end if
			rs1.movenext
		loop
		rs1.close

%>
	
	<script>
		eval("parent.ListForm.idx_<%=ControlIdx%>").value= "<%=idx%>";
	</script>

<%

	elseif sJob = "R" then

		SQL = "update armyinformix.dbo.contactlist set recordyn = '" & recyn & "' where CallId = '" & callid & "'"
		db.execute(SQL)

%>
	
	<script>
		parent.location.reload();
	</script>

<%


	end if

%>
	


<!-- #include virtual="/Include/Bottom.asp" -->