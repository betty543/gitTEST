<!-- #include virtual="/Include/Top_Frame.asp" -->
<!-- #include virtual="/Include/PopLayer.asp" -->
<%

			'ipaddress = Request.ServerVariables("Remote_ADDR")

'response.write ipaddress
	receiptfactnum=Request("receiptfactnum") 
	sGubun = Request("sGubun")
	sGubunName = Request("sGubunName")
	sToDay = date()


	'대상자 count

	sql = "	select sum(cnt1), sum(cnt2), sum(cnt3), sum(cnt4), sum(cnt5), sum(cnt6), sum(cnt7)"
	sql = sql & "	from ( select count(factnum) cnt1, 0 cnt2, 0 cnt3, 0 cnt4, 0 cnt5, 0 cnt6, 0 cnt7 from armyinformix.dbo.factpeople where factnum = '" & receiptfactnum & "' and section2 in ('B11','413')"
	sql = sql & "	union select 0 cnt1, count(factnum) cnt2, 0 cnt3, 0 cnt4, 0 cnt5, 0 cnt6, 0 cnt7 from armyinformix.dbo.factpeople where factnum = '" & receiptfactnum & "' and section2 in ('B12','448')"
	sql = sql & "	union select 0 cnt1, 0 cnt2, count(factnum) cnt3, 0 cnt4, 0 cnt5, 0 cnt6, 0 cnt7 from armyinformix.dbo.factpeople where factnum = '" & receiptfactnum & "' and section2 in ('B13','450')"
	sql = sql & "	union select 0 cnt1, 0 cnt2, 0 cnt3, count(factnum) cnt4, 0 cnt5, 0 cnt6, 0 cnt7 from armyinformix.dbo.factpeople where factnum = '" & receiptfactnum & "' and section2 in ('B14','451')"
	sql = sql & "	union select 0 cnt1, 0 cnt2, 0 cnt3, 0 cnt4, count(factnum) cnt5, 0 cnt6, 0 cnt7 from armyinformix.dbo.factpeople where factnum = '" & receiptfactnum & "' and section2 in ('B15','452')"
	sql = sql & "	union select 0 cnt1, 0 cnt2, 0 cnt3, 0 cnt4, 0 cnt5, count(factnum) cnt6, 0 cnt7 from armyinformix.dbo.factpeople where factnum = '" & receiptfactnum & "' and section2 in ('B16','453')"
	sql = sql & "	union select 0 cnt1, 0 cnt2, 0 cnt3, 0 cnt4, 0 cnt5, 0 cnt6, count(factnum) cnt7 from armyinformix.dbo.factpeople where factnum = '" & receiptfactnum & "' and section2 in ('B17','447','449')) a"	
	SET Rs = DB.execute(SQL)

	if Rs.eof = false then
		cnt1 = Rs(0)
		cnt2 = Rs(1)
		cnt3 = Rs(2)
		cnt4 = Rs(3)
		cnt5 = Rs(4)
		cnt6 = Rs(5)
		cnt7 = Rs(6)
	end if
	Rs.close
	set Rs = nothing


	sql = "	select sum(cnt1), sum(cnt2), sum(cnt3), sum(cnt4), sum(cnt5), sum(cnt6), sum(cnt7)"
	sql = sql & "	from ( select count(factnum) cnt1, 0 cnt2, 0 cnt3, 0 cnt4, 0 cnt5, 0 cnt6, 0 cnt7 from armyinformix.dbo.factpeople where factnum = '" & receiptfactnum & "' and section2 in ('B11','413') and (monitorresult is null or monitorresult = '4' or monitorresult ='') "
	sql = sql & "	union select 0 cnt1, count(factnum) cnt2, 0 cnt3, 0 cnt4, 0 cnt5, 0 cnt6, 0 cnt7 from armyinformix.dbo.factpeople where factnum = '" & receiptfactnum & "' and section2 in ('B12','448') and (monitorresult is null or monitorresult = '4' or monitorresult ='')"
	sql = sql & "	union select 0 cnt1, 0 cnt2, count(factnum) cnt3, 0 cnt4, 0 cnt5, 0 cnt6, 0 cnt7 from armyinformix.dbo.factpeople where factnum = '" & receiptfactnum & "' and section2 in ('B13','450') and (monitorresult is null or monitorresult = '4' or monitorresult ='')"
	sql = sql & "	union select 0 cnt1, 0 cnt2, 0 cnt3, count(factnum) cnt4, 0 cnt5, 0 cnt6, 0 cnt7 from armyinformix.dbo.factpeople where factnum = '" & receiptfactnum & "' and section2 in ('B14','451') and (monitorresult is null or monitorresult = '4' or monitorresult ='')"
	sql = sql & "	union select 0 cnt1, 0 cnt2, 0 cnt3, 0 cnt4, count(factnum) cnt5, 0 cnt6, 0 cnt7 from armyinformix.dbo.factpeople where factnum = '" & receiptfactnum & "' and section2 in ('B15','452') and (monitorresult is null or monitorresult = '4' or monitorresult ='')"
	sql = sql & "	union select 0 cnt1, 0 cnt2, 0 cnt3, 0 cnt4, 0 cnt5, count(factnum) cnt6, 0 cnt7 from armyinformix.dbo.factpeople where factnum = '" & receiptfactnum & "' and section2 in ('B16','453') and (monitorresult is null or monitorresult = '4' or monitorresult ='')"
	sql = sql & "	union select 0 cnt1, 0 cnt2, 0 cnt3, 0 cnt4, 0 cnt5, 0 cnt6, count(factnum) cnt7 from armyinformix.dbo.factpeople where factnum = '" & receiptfactnum & "' and section2 in ('B17','447','449') and (monitorresult is null or monitorresult = '4' or monitorresult ='')) a"	
	SET Rs = DB.execute(SQL)

	if Rs.eof = false then
		cnt11 = Rs(0)
		cnt21 = Rs(1)
		cnt31 = Rs(2)
		cnt41 = Rs(3)
		cnt51 = Rs(4)
		cnt61 = Rs(5)
		cnt71 = Rs(6)
	end if
	Rs.close
	set Rs = nothing


'response.write sGubun
	'관련인 조회하기
	if sGubun = "B11" then '피의자
		sql = "	select *, convert(char(10),MONITORDATE,121) MONITORDATE1 from armyinformix.dbo.factpeople where factnum = '" & receiptfactnum & "' and section2 in ('B11','413')"
	elseif sGubun = "B12" then '피해자
		sql = "	select *, convert(char(10),MONITORDATE,121) MONITORDATE1 from armyinformix.dbo.factpeople where factnum = '" & receiptfactnum & "' and section2 in ('B12','448')"
	elseif sGubun = "B13" then '민원인
		sql = "	select *, convert(char(10),MONITORDATE,121) MONITORDATE1 from armyinformix.dbo.factpeople where factnum = '" & receiptfactnum & "' and section2 in ('B13','450')"
	elseif sGubun = "B14" then '피민원인
		sql = "	select *, convert(char(10),MONITORDATE,121) MONITORDATE1 from armyinformix.dbo.factpeople where factnum = '" & receiptfactnum & "' and section2 in ('B14','451')"
	elseif sGubun = "B15" then '지휘관
		sql = "	select *, convert(char(10),MONITORDATE,121) MONITORDATE1 from armyinformix.dbo.factpeople where factnum = '" & receiptfactnum & "' and section2 in ('B15','452')"
	elseif sGubun = "B16" then '유족
		sql = "	select *, convert(char(10),MONITORDATE,121) MONITORDATE1 from armyinformix.dbo.factpeople where factnum = '" & receiptfactnum & "' and section2 in ('B16','453')"
	elseif sGubun = "B17" then '참고인
		sql = "	select *, convert(char(10),MONITORDATE,121) MONITORDATE1 from armyinformix.dbo.factpeople where factnum = '" & receiptfactnum & "' and section2 in ('B17','447','449')"	
	end if

	if SQL <> "" then

		SET Rs = db.execute(SQL)
		i = 0
		do until Rs.eof	
			i = i + 1
			
			if i = 1 then
				db_factPeoplenum_1 = rs("factPeoplenum")
				db_Name_1 = rs("name")
				db_level_1 = rs("level")
				db_homephone_1 = rs("homephone")		
				db_mobilephone_1 = rs("mobilephone")	
				db_factpeoplenum_1 = rs("factpeoplenum")	
				db_etcphone_1 = rs("etcphone")	
				db_MONITORDATE_1 = rs("MONITORDATE1")
				db_MONITOR_RESULT_1 = rs("MONITORRESULT")
				db_RESERVEDATE_1 = rs("RESERVEDATE")
				db_Remark_1 = rs("Remark")
				db_Remark1_1 = rs("Remark1")	
				if rs("monitorpoint") <> "" then
					db_TOT_1 = formatnumber(rs("monitorpoint"),2)
				end if

				'통화 히스토리 찾기
				sql = " select substring(convert(char(19),contactdate,121),6,11) contacttime, * from	armyinformix.dbo.contactlist where factnum = '" & trim(receiptfactnum) & "' and peoplenum = '" & rs("factPeoplenum") & "' order by idx"
				SET Rs1 = db.execute(SQL)
				do until rs1.eof
					if db_History_1 = "" then
						db_History_1 = "<a href='##' onClick=HistoryUpdate('1','"&rs1("idx")&"','"&trim(receiptfactnum)&"','"&rs("factPeoplenum")&"');>"&rs1("contacttime")&"&nbsp;("&rs1("contacttelno")&")&nbsp;"&rs1("remark")&"</a>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"&rs1("contactdate")&" 통화 삭제' style='cursor:hand;' align='absmiddle' onClick=HistoryDel('1','"&rs1("idx")&"','"&trim(receiptfactnum)&"','"&rs("factPeoplenum")&"');>"
						if rs1("callid") <> "" then
							if rs1("recordyn") = "Y" then

								SQL = "select d.RecordingFileName		FROM	I3_IC.dbo.RecordingCall c, I3_IC.dbo.RecordingData d 		WHERE	c.RecordingID = d.RecordingID 		AND	left(recordedcallidkey ,10) = '" & rs1("callid") & "'"

								SET Rs2 = db.execute(SQL)
								if Rs2.eof = false then
									db_RecordingFileName = rs2("RecordingFileName")

									sdb_RecordingFileName = "http://16.1.17.113:8080/"&mid(replace(db_RecordingFileName,"\","/"),27)
									sssdb_RecordingFileName = mid(replace(sdb_RecordingFileName,"\","/"),27)
									db_History_1 = db_History_1 & "&nbsp;"& CutString(sssdb_RecordingFileName, 40, "...") & "<img src='/Images/Comm/IconAlert.gif' align='absmiddle' style='cursor:hand;' onClick=""fn_Player('"&sdb_RecordingFileName&"');"" title='녹음내용 청취'>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"&rs1("contactdate")&" 녹취 삭제' style='cursor:hand;' align='absmiddle' onClick=RecDel('1','"&rs1("callid")&"','N');>"

								end if
								Rs2.close
							else
								SQL = "select d.RecordingFileName		FROM	I3_IC.dbo.RecordingCall c, I3_IC.dbo.RecordingData d 		WHERE	c.RecordingID = d.RecordingID 		AND	left(recordedcallidkey ,10) = '" & rs1("callid") & "'"

								SET Rs2 = db.execute(SQL)

								if Rs2.eof = false then
									db_RecordingFileName = rs2("RecordingFileName")
									sdb_RecordingFileName = "http://16.1.17.113:8080/"&mid(replace(db_RecordingFileName,"\","/"),27)
									sssdb_RecordingFileName = mid(replace(sdb_RecordingFileName,"\","/"),27)
									db_History_1 = db_History_1 & "&nbsp;<a href='##' onClick=RecDel('1','"&rs1("callid")&"','Y');>녹취첨부</a>"
								end if

								Rs2.close
							end if
						end if
					else
						db_History_1 = db_History_1 & "<br>&nbsp;<a href='##' onClick=HistoryUpdate('1','"&rs1("idx")&"','"&trim(receiptfactnum)&"','"&rs("factPeoplenum")&"');>"&rs1("contacttime")&"&nbsp;("&rs1("contacttelno")&")&nbsp;"&rs1("remark")&"</a>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"&rs1("contactdate")&" 통화 삭제' style='cursor:hand;' align='absmiddle' onClick=HistoryDel('1','"&rs1("idx")&"','"&trim(receiptfactnum)&"','"&rs("factPeoplenum")&"');>"

						if rs1("callid") <> "" then
							if rs1("recordyn") = "Y" then

								SQL = "select d.RecordingFileName		FROM	I3_IC.dbo.RecordingCall c, I3_IC.dbo.RecordingData d 		WHERE	c.RecordingID = d.RecordingID 		AND	left(recordedcallidkey ,10) = '" & rs1("callid") & "'"

								SET Rs2 = db.execute(SQL)
								if Rs2.eof = false then
									db_RecordingFileName = rs2("RecordingFileName")

									sdb_RecordingFileName = "http://16.1.17.113:8080/"&mid(replace(db_RecordingFileName,"\","/"),27)
									sssdb_RecordingFileName = mid(replace(sdb_RecordingFileName,"\","/"),27)
									db_History_1 = db_History_1 & "&nbsp;"& CutString(sssdb_RecordingFileName, 40, "...") & "<img src='/Images/Comm/IconAlert.gif' align='absmiddle' style='cursor:hand;' onClick=""fn_Player('"&sdb_RecordingFileName&"');"" title='녹음내용 청취'>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"&rs1("contactdate")&" 녹취 삭제' style='cursor:hand;' align='absmiddle' onClick=RecDel('1','"&rs1("callid")&"','N');>"

								end if
								Rs2.close
							else
								SQL = "select d.RecordingFileName		FROM	I3_IC.dbo.RecordingCall c, I3_IC.dbo.RecordingData d 		WHERE	c.RecordingID = d.RecordingID 		AND	left(recordedcallidkey ,10) = '" & rs1("callid") & "'"

								SET Rs2 = db.execute(SQL)

								if Rs2.eof = false then
									db_RecordingFileName = rs2("RecordingFileName")

									sdb_RecordingFileName = "http://16.1.17.113:8080/"&mid(replace(db_RecordingFileName,"\","/"),27)
									sssdb_RecordingFileName = mid(replace(sdb_RecordingFileName,"\","/"),27)
									db_History_1 = db_History_1 & "&nbsp;<a href='##' onClick=RecDel('1','"&rs1("callid")&"','Y');>녹취첨부</a>"
								end if

								Rs2.close
							end if
						end if
					end if
					rs1.movenext
				loop
				rs1.close

			elseif i = 2 then
				db_factPeoplenum_2 = rs("factPeoplenum")
				db_Name_2 = rs("name")
				db_level_2 = rs("level")
				db_homephone_2 = rs("homephone")		
				db_mobilephone_2 = rs("mobilephone")	
				db_factpeoplenum_2 = rs("factpeoplenum")
				db_etcphone_2 = rs("etcphone")	
				db_MONITORDATE_2 = rs("MONITORDATE1")
				db_MONITOR_RESULT_2 = rs("MONITORRESULT")
				db_RESERVEDATE_2 = rs("RESERVEDATE")
				db_Remark_2 = rs("Remark")
				db_Remark1_2 = rs("Remark1")
				if rs("monitorpoint") <> "" then
					db_TOT_2 = formatnumber(rs("monitorpoint"),2)
				end if

				'통화 히스토리 찾기
				sql = " select substring(convert(char(19),contactdate,121),6,11) contacttime, * from	armyinformix.dbo.contactlist where factnum = '" & trim(receiptfactnum) & "' and peoplenum = '" & rs("factPeoplenum") & "' order by idx"
				SET Rs1 = db.execute(SQL)
				do until rs1.eof
					if db_History_2 = "" then
						db_History_2 = "<a href='##' onClick=HistoryUpdate('2','"&rs1("idx")&"','"&trim(receiptfactnum)&"','"&rs("factPeoplenum")&"');>"&rs1("contacttime")&"&nbsp;("&rs1("contacttelno")&")&nbsp;"&rs1("remark")&"</a>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"&rs1("contactdate")&" 통화 삭제' style='cursor:hand;' align='absmiddle' onClick=HistoryDel('2','"&rs1("idx")&"','"&trim(receiptfactnum)&"','"&rs("factPeoplenum")&"');>"
						if rs1("callid") <> "" then
							if rs1("recordyn") = "Y" then

								SQL = "select d.RecordingFileName		FROM	I3_IC.dbo.RecordingCall c, I3_IC.dbo.RecordingData d 		WHERE	c.RecordingID = d.RecordingID 		AND	left(recordedcallidkey ,10) = '" & rs1("callid") & "'"

								SET Rs2 = db.execute(SQL)
								if Rs2.eof = false then
									db_RecordingFileName = rs2("RecordingFileName")

									sdb_RecordingFileName = "http://16.1.17.113:8080/"&mid(replace(db_RecordingFileName,"\","/"),27)
									sssdb_RecordingFileName = mid(replace(sdb_RecordingFileName,"\","/"),27)
									db_History_2 = db_History_2 & "&nbsp;"& CutString(sssdb_RecordingFileName, 40, "...") & "<img src='/Images/Comm/IconAlert.gif' align='absmiddle' style='cursor:hand;' onClick=""fn_Player('"&sdb_RecordingFileName&"');"" title='녹음내용 청취'>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"&rs1("contactdate")&" 녹취 삭제' style='cursor:hand;' align='absmiddle' onClick=RecDel('2','"&rs1("callid")&"','N');>"

								end if
								Rs2.close
							else
								SQL = "select d.RecordingFileName		FROM	I3_IC.dbo.RecordingCall c, I3_IC.dbo.RecordingData d 		WHERE	c.RecordingID = d.RecordingID 		AND	left(recordedcallidkey ,10) = '" & rs1("callid") & "'"

								SET Rs2 = db.execute(SQL)

								if Rs2.eof = false then
									db_RecordingFileName = rs2("RecordingFileName")

									sdb_RecordingFileName = "http://16.1.17.113:8080/"&mid(replace(db_RecordingFileName,"\","/"),27)
									sssdb_RecordingFileName = mid(replace(sdb_RecordingFileName,"\","/"),27)
									db_History_2 = db_History_2 & "&nbsp;<a href='##' onClick=RecDel('2','"&rs1("callid")&"','Y');>녹취첨부</a>"
								end if

								Rs2.close
							end if
						end if
					else
						db_History_2 = db_History_2 & "<br>&nbsp;<a href='##' onClick=HistoryUpdate('2','"&rs1("idx")&"','"&trim(receiptfactnum)&"','"&rs("factPeoplenum")&"');>"&rs1("contacttime")&"&nbsp;("&rs1("contacttelno")&")&nbsp;"&rs1("remark")&"</a>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"&rs1("contactdate")&" 통화 삭제' style='cursor:hand;' align='absmiddle' onClick=HistoryDel('2','"&rs1("idx")&"','"&trim(receiptfactnum)&"','"&rs("factPeoplenum")&"');>"

						if rs1("callid") <> "" then
							if rs1("recordyn") = "Y" then

								SQL = "select d.RecordingFileName		FROM	I3_IC.dbo.RecordingCall c, I3_IC.dbo.RecordingData d 		WHERE	c.RecordingID = d.RecordingID 		AND	left(recordedcallidkey ,10) = '" & rs1("callid") & "'"

								SET Rs2 = db.execute(SQL)
								if Rs2.eof = false then
									db_RecordingFileName = rs2("RecordingFileName")

									sdb_RecordingFileName = "http://16.1.17.113:8080/"&mid(replace(db_RecordingFileName,"\","/"),27)
									sssdb_RecordingFileName = mid(replace(sdb_RecordingFileName,"\","/"),27)
									db_History_2 = db_History_2 & "&nbsp;"& CutString(sssdb_RecordingFileName, 40, "...") & "<img src='/Images/Comm/IconAlert.gif' align='absmiddle' style='cursor:hand;' onClick=""fn_Player('"&sdb_RecordingFileName&"');"" title='녹음내용 청취'>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"&rs1("contactdate")&" 녹취 삭제' style='cursor:hand;' align='absmiddle' onClick=RecDel('2','"&rs1("callid")&"','N');>"

								end if
								Rs2.close
							else
								SQL = "select d.RecordingFileName		FROM	I3_IC.dbo.RecordingCall c, I3_IC.dbo.RecordingData d 		WHERE	c.RecordingID = d.RecordingID 		AND	left(recordedcallidkey ,10) = '" & rs1("callid") & "'"

								SET Rs2 = db.execute(SQL)

								if Rs2.eof = false then
									db_RecordingFileName = rs2("RecordingFileName")

									sdb_RecordingFileName = "http://16.1.17.113:8080/"&mid(replace(db_RecordingFileName,"\","/"),27)
									sssdb_RecordingFileName = mid(replace(sdb_RecordingFileName,"\","/"),27)
									db_History_2 = db_History_2 & "&nbsp;<a href='##' onClick=RecDel('2','"&rs1("callid")&"','Y');>녹취첨부</a>"
								end if

								Rs2.close
							end if
						end if
					end if


					rs1.movenext
				loop
				rs1.close
			elseif i = 3 then
				db_factPeoplenum_3 = rs("factPeoplenum")
				db_Name_3 = rs("name")
				db_level_3 = rs("level")
				db_homephone_3 = rs("homephone")		
				db_mobilephone_3 = rs("mobilephone")	
				db_factpeoplenum_3 = rs("factpeoplenum")	
				db_etcphone_3 = rs("etcphone")	
				db_MONITORDATE_3 = rs("MONITORDATE1")
				db_MONITOR_RESULT_3 = rs("MONITORRESULT")
				db_RESERVEDATE_3 = rs("RESERVEDATE")	
				db_Remark_3 = rs("Remark")		
				db_Remark1_3 = rs("Remark1")
				if rs("monitorpoint") <> "" then
					db_TOT_3 = formatnumber(rs("monitorpoint"),2)
				end if		

				'통화 히스토리 찾기
				sql = " select substring(convert(char(19),contactdate,121),6,11) contacttime, * from	armyinformix.dbo.contactlist where factnum = '" & trim(receiptfactnum) & "' and peoplenum = '" & rs("factPeoplenum") & "' order by idx"
				SET Rs1 = db.execute(SQL)
				do until rs1.eof
					if db_History_3 = "" then
						db_History_3 = "<a href='##' onClick=HistoryUpdate('3','"&rs1("idx")&"','"&trim(receiptfactnum)&"','"&rs("factPeoplenum")&"');>"&rs1("contacttime")&"&nbsp;("&rs1("contacttelno")&")&nbsp;"&rs1("remark")&"</a>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"&rs1("contactdate")&" 통화 삭제' style='cursor:hand;' align='absmiddle' onClick=HistoryDel('3','"&rs1("idx")&"','"&trim(receiptfactnum)&"','"&rs("factPeoplenum")&"');>"
						if rs1("callid") <> "" then
							if rs1("recordyn") = "Y" then

								SQL = "select d.RecordingFileName		FROM	I3_IC.dbo.RecordingCall c, I3_IC.dbo.RecordingData d 		WHERE	c.RecordingID = d.RecordingID 		AND	left(recordedcallidkey ,10) = '" & rs1("callid") & "'"

								SET Rs2 = db.execute(SQL)
								if Rs2.eof = false then
									db_RecordingFileName = rs2("RecordingFileName")

									sdb_RecordingFileName = "http://16.1.17.113:8080/"&mid(replace(db_RecordingFileName,"\","/"),27)
									sssdb_RecordingFileName = mid(replace(sdb_RecordingFileName,"\","/"),27)
									db_History_3 = db_History_3 & "&nbsp;"& CutString(sssdb_RecordingFileName, 40, "...") & "<img src='/Images/Comm/IconAlert.gif' align='absmiddle' style='cursor:hand;' onClick=""fn_Player('"&sdb_RecordingFileName&"');"" title='녹음내용 청취'>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"&rs1("contactdate")&" 녹취 삭제' style='cursor:hand;' align='absmiddle' onClick=RecDel('3','"&rs1("callid")&"','N');>"

								end if
								Rs2.close
							else
								SQL = "select d.RecordingFileName		FROM	I3_IC.dbo.RecordingCall c, I3_IC.dbo.RecordingData d 		WHERE	c.RecordingID = d.RecordingID 		AND	left(recordedcallidkey ,10) = '" & rs1("callid") & "'"

								SET Rs2 = db.execute(SQL)

								if Rs2.eof = false then
									db_RecordingFileName = rs2("RecordingFileName")

									sdb_RecordingFileName = "http://16.1.17.113:8080/"&mid(replace(db_RecordingFileName,"\","/"),27)
									sssdb_RecordingFileName = mid(replace(sdb_RecordingFileName,"\","/"),27)
									db_History_3 = db_History_3 & "&nbsp;<a href='##' onClick=RecDel('3','"&rs1("callid")&"','Y');>녹취첨부</a>"
								end if

								Rs2.close
							end if
						end if
					else
						db_History_3 = db_History_3 & "<br>&nbsp;<a href='##' onClick=HistoryUpdate('3','"&rs1("idx")&"','"&trim(receiptfactnum)&"','"&rs("factPeoplenum")&"');>"&rs1("contacttime")&"&nbsp;("&rs1("contacttelno")&")&nbsp;"&rs1("remark")&"</a>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"&rs1("contactdate")&" 통화 삭제' style='cursor:hand;' align='absmiddle' onClick=HistoryDel('3','"&rs1("idx")&"','"&trim(receiptfactnum)&"','"&rs("factPeoplenum")&"');>"

						if rs1("callid") <> "" then
							if rs1("recordyn") = "Y" then

								SQL = "select d.RecordingFileName		FROM	I3_IC.dbo.RecordingCall c, I3_IC.dbo.RecordingData d 		WHERE	c.RecordingID = d.RecordingID 		AND	left(recordedcallidkey ,10) = '" & rs1("callid") & "'"

								SET Rs2 = db.execute(SQL)
								if Rs2.eof = false then
									db_RecordingFileName = rs2("RecordingFileName")

									sdb_RecordingFileName = "http://16.1.17.113:8080/"&mid(replace(db_RecordingFileName,"\","/"),27)
									sssdb_RecordingFileName = mid(replace(sdb_RecordingFileName,"\","/"),27)
									db_History_3 = db_History_3 & "&nbsp;"& CutString(sssdb_RecordingFileName, 40, "...") & "<img src='/Images/Comm/IconAlert.gif' align='absmiddle' style='cursor:hand;' onClick=""fn_Player('"&sdb_RecordingFileName&"');"" title='녹음내용 청취'>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"&rs1("contactdate")&" 녹취 삭제' style='cursor:hand;' align='absmiddle' onClick=RecDel('3','"&rs1("callid")&"','N');>"

								end if
								Rs2.close
							else
								SQL = "select d.RecordingFileName		FROM	I3_IC.dbo.RecordingCall c, I3_IC.dbo.RecordingData d 		WHERE	c.RecordingID = d.RecordingID 		AND	left(recordedcallidkey ,10) = '" & rs1("callid") & "'"

								SET Rs2 = db.execute(SQL)

								if Rs2.eof = false then
									db_RecordingFileName = rs2("RecordingFileName")

									sdb_RecordingFileName = "http://16.1.17.113:8080/"&mid(replace(db_RecordingFileName,"\","/"),27)
									sssdb_RecordingFileName = mid(replace(sdb_RecordingFileName,"\","/"),27)
									db_History_3 = db_History_3 & "&nbsp;<a href='##' onClick=RecDel('3','"&rs1("callid")&"','Y');>녹취첨부</a>"
								end if

								Rs2.close
							end if
						end if
					end if
					rs1.movenext
				loop
				rs1.close

			elseif i = 4 then
				db_factPeoplenum_4 = rs("factPeoplenum")
				db_Name_4 = rs("name")
				db_level_4 = rs("level")
				db_homephone_4 = rs("homephone")		
				db_mobilephone_4 = rs("mobilephone")	
				db_factpeoplenum_4 = rs("factpeoplenum")
				db_etcphone_4 = rs("etcphone")
				db_MONITORDATE_4 = rs("MONITORDATE1")
				db_MONITOR_RESULT_4 = rs("MONITORRESULT")
				db_RESERVEDATE_4 = rs("RESERVEDATE")
				db_Remark_4 = rs("Remark")		
				db_Remark1_4 = rs("Remark1")
				if rs("monitorpoint") <> "" then
					db_TOT_4 = formatnumber(rs("monitorpoint"),2)
				end if			

				'통화 히스토리 찾기
				sql = " select substring(convert(char(19),contactdate,121),6,11) contacttime, * from	armyinformix.dbo.contactlist where factnum = '" & trim(receiptfactnum) & "' and peoplenum = '" & rs("factPeoplenum") & "' order by idx"
				SET Rs1 = db.execute(SQL)
				do until rs1.eof
					if db_History_4 = "" then
						db_History_4 = "<a href='##' onClick=HistoryUpdate('4','"&rs1("idx")&"','"&trim(receiptfactnum)&"','"&rs("factPeoplenum")&"');>"&rs1("contacttime")&"&nbsp;("&rs1("contacttelno")&")&nbsp;"&rs1("remark")&"</a>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"&rs1("contactdate")&" 통화 삭제' style='cursor:hand;' align='absmiddle' onClick=HistoryDel('4','"&rs1("idx")&"','"&trim(receiptfactnum)&"','"&rs("factPeoplenum")&"');>"
						if rs1("callid") <> "" then
							if rs1("recordyn") = "Y" then

								SQL = "select d.RecordingFileName		FROM	I3_IC.dbo.RecordingCall c, I3_IC.dbo.RecordingData d 		WHERE	c.RecordingID = d.RecordingID 		AND	left(recordedcallidkey ,10) = '" & rs1("callid") & "'"

								SET Rs2 = db.execute(SQL)
								if Rs2.eof = false then
									db_RecordingFileName = rs2("RecordingFileName")

									sdb_RecordingFileName = "http://16.1.17.113:8080/"&mid(replace(db_RecordingFileName,"\","/"),27)
									sssdb_RecordingFileName = mid(replace(sdb_RecordingFileName,"\","/"),27)
									db_History_4 = db_History_4 & "&nbsp;"& CutString(sssdb_RecordingFileName, 40, "...") & "<img src='/Images/Comm/IconAlert.gif' align='absmiddle' style='cursor:hand;' onClick=""fn_Player('"&sdb_RecordingFileName&"');"" title='녹음내용 청취'>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"&rs1("contactdate")&" 녹취 삭제' style='cursor:hand;' align='absmiddle' onClick=RecDel('4','"&rs1("callid")&"','N');>"

								end if
								Rs2.close
							else
								SQL = "select d.RecordingFileName		FROM	I3_IC.dbo.RecordingCall c, I3_IC.dbo.RecordingData d 		WHERE	c.RecordingID = d.RecordingID 		AND	left(recordedcallidkey ,10) = '" & rs1("callid") & "'"

								SET Rs2 = db.execute(SQL)

								if Rs2.eof = false then
									db_RecordingFileName = rs2("RecordingFileName")

									sdb_RecordingFileName = "http://16.1.17.113:8080/"&mid(replace(db_RecordingFileName,"\","/"),27)
									sssdb_RecordingFileName = mid(replace(sdb_RecordingFileName,"\","/"),27)
									db_History_4 = db_History_4 & "&nbsp;<a href='##' onClick=RecDel('4','"&rs1("callid")&"','Y');>녹취첨부</a>"
								end if

								Rs2.close
							end if
						end if
					else
						db_History_4 = db_History_4 & "<br>&nbsp;<a href='##' onClick=HistoryUpdate('4','"&rs1("idx")&"','"&trim(receiptfactnum)&"','"&rs("factPeoplenum")&"');>"&rs1("contacttime")&"&nbsp;("&rs1("contacttelno")&")&nbsp;"&rs1("remark")&"</a>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"&rs1("contactdate")&" 통화 삭제' style='cursor:hand;' align='absmiddle' onClick=HistoryDel('4','"&rs1("idx")&"','"&trim(receiptfactnum)&"','"&rs("factPeoplenum")&"');>"

						if rs1("callid") <> "" then
							if rs1("recordyn") = "Y" then

								SQL = "select d.RecordingFileName		FROM	I3_IC.dbo.RecordingCall c, I3_IC.dbo.RecordingData d 		WHERE	c.RecordingID = d.RecordingID 		AND	left(recordedcallidkey ,10) = '" & rs1("callid") & "'"

								SET Rs2 = db.execute(SQL)
								if Rs2.eof = false then
									db_RecordingFileName = rs2("RecordingFileName")

									sdb_RecordingFileName = "http://16.1.17.113:8080/"&mid(replace(db_RecordingFileName,"\","/"),27)
									sssdb_RecordingFileName = mid(replace(sdb_RecordingFileName,"\","/"),27)
									db_History_4 = db_History_4 & "&nbsp;"& CutString(sssdb_RecordingFileName, 40, "...") & "<img src='/Images/Comm/IconAlert.gif' align='absmiddle' style='cursor:hand;' onClick=""fn_Player('"&sdb_RecordingFileName&"');"" title='녹음내용 청취'>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"&rs1("contactdate")&" 녹취 삭제' style='cursor:hand;' align='absmiddle' onClick=RecDel('4','"&rs1("callid")&"','N');>"

								end if
								Rs2.close
							else
								SQL = "select d.RecordingFileName		FROM	I3_IC.dbo.RecordingCall c, I3_IC.dbo.RecordingData d 		WHERE	c.RecordingID = d.RecordingID 		AND	left(recordedcallidkey ,10) = '" & rs1("callid") & "'"

								SET Rs2 = db.execute(SQL)

								if Rs2.eof = false then
									db_RecordingFileName = rs2("RecordingFileName")

									sdb_RecordingFileName = "http://16.1.17.113:8080/"&mid(replace(db_RecordingFileName,"\","/"),27)
									sssdb_RecordingFileName = mid(replace(sdb_RecordingFileName,"\","/"),27)
									db_History_4 = db_History_4 & "&nbsp;<a href='##' onClick=RecDel('4','"&rs1("callid")&"','Y');>녹취첨부</a>"
								end if

								Rs2.close
							end if
						end if

					end if
					rs1.movenext
				loop
				rs1.close
				
			elseif i = 5 then
				db_factPeoplenum_5 = rs("factPeoplenum")
				db_Name_5 = rs("name")
				db_level_5 = rs("level")
				db_homephone_5 = rs("homephone")		
				db_mobilephone_5 = rs("mobilephone")	
				db_factpeoplenum_5 = rs("factpeoplenum")	
				db_etcphone_5 = rs("etcphone")		
				db_MONITORDATE_5 = rs("MONITORDATE1")
				db_MONITOR_RESULT_5 = rs("MONITORRESULT")
				db_RESERVEDATE_5 = rs("RESERVEDATE")	
				db_Remark_5 = rs("Remark")	
				db_Remark1_5 = rs("Remark1")
				if rs("monitorpoint") <> "" then
					db_TOT_5 = formatnumber(rs("monitorpoint"),2)
				end if		

				'통화 히스토리 찾기
				sql = " select substring(convert(char(19),contactdate,121),6,11) contacttime, * from	armyinformix.dbo.contactlist where factnum = '" & trim(receiptfactnum) & "' and peoplenum = '" & rs("factPeoplenum") & "' order by idx"
				SET Rs1 = db.execute(SQL)
				do until rs1.eof
					if db_History_5 = "" then
						db_History_5 = "<a href='##' onClick=HistoryUpdate('5','"&rs1("idx")&"','"&trim(receiptfactnum)&"','"&rs("factPeoplenum")&"');>"&rs1("contacttime")&"&nbsp;("&rs1("contacttelno")&")&nbsp;"&rs1("remark")&"</a>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"&rs1("contactdate")&" 통화 삭제' style='cursor:hand;' align='absmiddle' onClick=HistoryDel('5','"&rs1("idx")&"','"&trim(receiptfactnum)&"','"&rs("factPeoplenum")&"');>"
						if rs1("callid") <> "" then
							if rs1("recordyn") = "Y" then

								SQL = "select d.RecordingFileName		FROM	I3_IC.dbo.RecordingCall c, I3_IC.dbo.RecordingData d 		WHERE	c.RecordingID = d.RecordingID 		AND	left(recordedcallidkey ,10) = '" & rs1("callid") & "'"

								SET Rs2 = db.execute(SQL)
								if Rs2.eof = false then
									db_RecordingFileName = rs2("RecordingFileName")

									sdb_RecordingFileName = "http://16.1.17.113:8080/"&mid(replace(db_RecordingFileName,"\","/"),27)
									sssdb_RecordingFileName = mid(replace(sdb_RecordingFileName,"\","/"),27)
									db_History_5 = db_History_5 & "&nbsp;"& CutString(sssdb_RecordingFileName, 40, "...") & "<img src='/Images/Comm/IconAlert.gif' align='absmiddle' style='cursor:hand;' onClick=""fn_Player('"&sdb_RecordingFileName&"');"" title='녹음내용 청취'>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"&rs1("contactdate")&" 녹취 삭제' style='cursor:hand;' align='absmiddle' onClick=RecDel('5','"&rs1("callid")&"','N');>"

								end if
								Rs2.close
							else
								SQL = "select d.RecordingFileName		FROM	I3_IC.dbo.RecordingCall c, I3_IC.dbo.RecordingData d 		WHERE	c.RecordingID = d.RecordingID 		AND	left(recordedcallidkey ,10) = '" & rs1("callid") & "'"

								SET Rs2 = db.execute(SQL)

								if Rs2.eof = false then
									db_RecordingFileName = rs2("RecordingFileName")

									sdb_RecordingFileName = "http://16.1.17.113:8080/"&mid(replace(db_RecordingFileName,"\","/"),27)
									sssdb_RecordingFileName = mid(replace(sdb_RecordingFileName,"\","/"),27)
									db_History_5 = db_History_5 & "&nbsp;<a href='##' onClick=RecDel('5','"&rs1("callid")&"','Y');>녹취첨부</a>"
								end if

								Rs2.close
							end if
						end if
					else
						db_History_5 = db_History_5 & "<br>&nbsp;<a href='##' onClick=HistoryUpdate('5','"&rs1("idx")&"','"&trim(receiptfactnum)&"','"&rs("factPeoplenum")&"');>"&rs1("contacttime")&"&nbsp;("&rs1("contacttelno")&")&nbsp;"&rs1("remark")&"</a>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"&rs1("contactdate")&" 통화 삭제' style='cursor:hand;' align='absmiddle' onClick=HistoryDel('5','"&rs1("idx")&"','"&trim(receiptfactnum)&"','"&rs("factPeoplenum")&"');>"
						if rs1("callid") <> "" then
							if rs1("recordyn") = "Y" then

								SQL = "select d.RecordingFileName		FROM	I3_IC.dbo.RecordingCall c, I3_IC.dbo.RecordingData d 		WHERE	c.RecordingID = d.RecordingID 		AND	left(recordedcallidkey ,10) = '" & rs1("callid") & "'"

								SET Rs2 = db.execute(SQL)
								if Rs2.eof = false then
									db_RecordingFileName = rs2("RecordingFileName")

									sdb_RecordingFileName = "http://16.1.17.113:8080/"&mid(replace(db_RecordingFileName,"\","/"),27)
									sssdb_RecordingFileName = mid(replace(sdb_RecordingFileName,"\","/"),27)
									db_History_5 = db_History_5 & "&nbsp;"& CutString(sssdb_RecordingFileName, 40, "...") & "<img src='/Images/Comm/IconAlert.gif' align='absmiddle' style='cursor:hand;' onClick=""fn_Player('"&sdb_RecordingFileName&"');"" title='녹음내용 청취'>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"&rs1("contactdate")&" 녹취 삭제' style='cursor:hand;' align='absmiddle' onClick=RecDel('5','"&rs1("callid")&"','N');>"

								end if
								Rs2.close
							else
								SQL = "select d.RecordingFileName		FROM	I3_IC.dbo.RecordingCall c, I3_IC.dbo.RecordingData d 		WHERE	c.RecordingID = d.RecordingID 		AND	left(recordedcallidkey ,10) = '" & rs1("callid") & "'"

								SET Rs2 = db.execute(SQL)

								if Rs2.eof = false then
									db_RecordingFileName = rs2("RecordingFileName")

									sdb_RecordingFileName = "http://16.1.17.113:8080/"&mid(replace(db_RecordingFileName,"\","/"),27)
									sssdb_RecordingFileName = mid(replace(sdb_RecordingFileName,"\","/"),27)
									db_History_5 = db_History_5 & "&nbsp;<a href='##' onClick=RecDel('5','"&rs1("callid")&"','Y');>녹취첨부</a>"
								end if

								Rs2.close
							end if
						end if
					end if
					rs1.movenext
				loop
				rs1.close

			elseif i = 6 then
				db_factPeoplenum_6 = rs("factPeoplenum")
				db_Name_6 = rs("name")
				db_level_6 = rs("level")
				db_homephone_6 = rs("homephone")		
				db_mobilephone_6 = rs("mobilephone")	
				db_factpeoplenum_6 = rs("factpeoplenum")	
				db_etcphone_6 = rs("etcphone")		
				db_MONITORDATE_6 = rs("MONITORDATE1")
				db_MONITOR_RESULT_6 = rs("MONITORRESULT")
				db_RESERVEDATE_6 = rs("RESERVEDATE")	
				db_Remark_6 = rs("Remark")	
				db_Remark1_6 = rs("Remark1")
				if rs("monitorpoint") <> "" then
					db_TOT_6 = formatnumber(rs("monitorpoint"),2)
				end if	

				'통화 히스토리 찾기
				sql = " select substring(convert(char(19),contactdate,121),6,11) contacttime, * from	armyinformix.dbo.contactlist where factnum = '" & trim(receiptfactnum) & "' and peoplenum = '" & rs("factPeoplenum") & "' order by idx"
				SET Rs1 = db.execute(SQL)
				do until rs1.eof
					if db_History_6 = "" then
						db_History_6 = "<a href='##' onClick=HistoryUpdate('6','"&rs1("idx")&"','"&trim(receiptfactnum)&"','"&rs("factPeoplenum")&"');>"&rs1("contacttime")&"&nbsp;("&rs1("contacttelno")&")&nbsp;"&rs1("remark")&"</a>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"&rs1("contactdate")&" 통화 삭제' style='cursor:hand;' align='absmiddle' onClick=HistoryDel('6','"&rs1("idx")&"','"&trim(receiptfactnum)&"','"&rs("factPeoplenum")&"');>"
						if rs1("callid") <> "" then
							if rs1("recordyn") = "Y" then

								SQL = "select d.RecordingFileName		FROM	I3_IC.dbo.RecordingCall c, I3_IC.dbo.RecordingData d 		WHERE	c.RecordingID = d.RecordingID 		AND	left(recordedcallidkey ,10) = '" & rs1("callid") & "'"

								SET Rs2 = db.execute(SQL)
								if Rs2.eof = false then
									db_RecordingFileName = rs2("RecordingFileName")

									sdb_RecordingFileName = "http://16.1.17.113:8080/"&mid(replace(db_RecordingFileName,"\","/"),27)
									sssdb_RecordingFileName = mid(replace(sdb_RecordingFileName,"\","/"),27)
									db_History_6 = db_History_6 & "&nbsp;"& CutString(sssdb_RecordingFileName, 40, "...") & "<img src='/Images/Comm/IconAlert.gif' align='absmiddle' style='cursor:hand;' onClick=""fn_Player('"&sdb_RecordingFileName&"');"" title='녹음내용 청취'>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"&rs1("contactdate")&" 녹취 삭제' style='cursor:hand;' align='absmiddle' onClick=RecDel('6','"&rs1("callid")&"','N');>"

								end if
								Rs2.close
							else
								SQL = "select d.RecordingFileName		FROM	I3_IC.dbo.RecordingCall c, I3_IC.dbo.RecordingData d 		WHERE	c.RecordingID = d.RecordingID 		AND	left(recordedcallidkey ,10) = '" & rs1("callid") & "'"

								SET Rs2 = db.execute(SQL)

								if Rs2.eof = false then
									db_RecordingFileName = rs2("RecordingFileName")

									sdb_RecordingFileName = "http://16.1.17.113:8080/"&mid(replace(db_RecordingFileName,"\","/"),27)
									sssdb_RecordingFileName = mid(replace(sdb_RecordingFileName,"\","/"),27)
									db_History_6 = db_History_6 & "&nbsp;<a href='##' onClick=RecDel('6','"&rs1("callid")&"','Y');>녹취첨부</a>"
								end if

								Rs2.close
							end if
						end if
					else
						db_History_6 = db_History_6 & "<br>&nbsp;<a href='##' onClick=HistoryUpdate('6','"&rs1("idx")&"','"&trim(receiptfactnum)&"','"&rs("factPeoplenum")&"');>"&rs1("contacttime")&"&nbsp;("&rs1("contacttelno")&")&nbsp;"&rs1("remark")&"</a>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"&rs1("contactdate")&" 통화 삭제' style='cursor:hand;' align='absmiddle' onClick=HistoryDel('6','"&rs1("idx")&"','"&trim(receiptfactnum)&"','"&rs("factPeoplenum")&"');>"
						if rs1("callid") <> "" then
							if rs1("recordyn") = "Y" then

								SQL = "select d.RecordingFileName		FROM	I3_IC.dbo.RecordingCall c, I3_IC.dbo.RecordingData d 		WHERE	c.RecordingID = d.RecordingID 		AND	left(recordedcallidkey ,10) = '" & rs1("callid") & "'"

								SET Rs2 = db.execute(SQL)
								if Rs2.eof = false then
									db_RecordingFileName = rs2("RecordingFileName")

									sdb_RecordingFileName = "http://16.1.17.113:8080/"&mid(replace(db_RecordingFileName,"\","/"),27)
									sssdb_RecordingFileName = mid(replace(sdb_RecordingFileName,"\","/"),27)
									db_History_6 = db_History_6 & "&nbsp;"& CutString(sssdb_RecordingFileName, 40, "...") & "<img src='/Images/Comm/IconAlert.gif' align='absmiddle' style='cursor:hand;' onClick=""fn_Player('"&sdb_RecordingFileName&"');"" title='녹음내용 청취'>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"&rs1("contactdate")&" 녹취 삭제' style='cursor:hand;' align='absmiddle' onClick=RecDel('6','"&rs1("callid")&"','N');>"

								end if
								Rs2.close
							else
								SQL = "select d.RecordingFileName		FROM	I3_IC.dbo.RecordingCall c, I3_IC.dbo.RecordingData d 		WHERE	c.RecordingID = d.RecordingID 		AND	left(recordedcallidkey ,10) = '" & rs1("callid") & "'"

								SET Rs2 = db.execute(SQL)

								if Rs2.eof = false then
									db_RecordingFileName = rs2("RecordingFileName")

									sdb_RecordingFileName = "http://16.1.17.113:8080/"&mid(replace(db_RecordingFileName,"\","/"),27)
									sssdb_RecordingFileName = mid(replace(sdb_RecordingFileName,"\","/"),27)
									db_History_6 = db_History_6 & "&nbsp;<a href='##' onClick=RecDel('6','"&rs1("callid")&"','Y');>녹취첨부</a>"
								end if

								Rs2.close
							end if
						end if
					end if
					rs1.movenext
				loop
				rs1.close

			elseif i = 7 then
				db_factPeoplenum_7 = rs("factPeoplenum")
				db_Name_7 = rs("name")
				db_level_7 = rs("level")
				db_homephone_7 = rs("homephone")		
				db_mobilephone_7 = rs("mobilephone")	
				db_factpeoplenum_7 = rs("factpeoplenum")	
				db_etcphone_7 = rs("etcphone")		
				db_MONITORDATE_7 = rs("MONITORDATE1")
				db_MONITOR_RESULT_7 = rs("MONITORRESULT")
				db_RESERVEDATE_7 = rs("RESERVEDATE")	
				db_Remark_7 = rs("Remark")	
				db_Remark1_7 = rs("Remark1")
				if rs("monitorpoint") <> "" then
					db_TOT_7 = formatnumber(rs("monitorpoint"),2)
				end if	
				'통화 히스토리 찾기
				sql = " select substring(convert(char(19),contactdate,121),6,11) contacttime, * from	armyinformix.dbo.contactlist where factnum = '" & trim(receiptfactnum) & "' and peoplenum = '" & rs("factPeoplenum") & "' order by idx"
				SET Rs1 = db.execute(SQL)
				do until rs1.eof
					if db_History_7 = "" then
						db_History_7 = "<a href='##' onClick=HistoryUpdate('7','"&rs1("idx")&"','"&trim(receiptfactnum)&"','"&rs("factPeoplenum")&"');>"&rs1("contacttime")&"&nbsp;("&rs1("contacttelno")&")&nbsp;"&rs1("remark")&"</a>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"&rs1("contactdate")&" 통화 삭제' style='cursor:hand;' align='absmiddle' onClick=HistoryDel('7','"&rs1("idx")&"','"&trim(receiptfactnum)&"','"&rs("factPeoplenum")&"');>"
						if rs1("callid") <> "" then
							if rs1("recordyn") = "Y" then

								SQL = "select d.RecordingFileName		FROM	I3_IC.dbo.RecordingCall c, I3_IC.dbo.RecordingData d 		WHERE	c.RecordingID = d.RecordingID 		AND	left(recordedcallidkey ,10) = '" & rs1("callid") & "'"

								SET Rs2 = db.execute(SQL)
								if Rs2.eof = false then
									db_RecordingFileName = rs2("RecordingFileName")

									sdb_RecordingFileName = "http://16.1.17.113:8080/"&mid(replace(db_RecordingFileName,"\","/"),27)
									sssdb_RecordingFileName = mid(replace(sdb_RecordingFileName,"\","/"),27)
									db_History_7 = db_History_7 & "&nbsp;"& CutString(sssdb_RecordingFileName, 40, "...") & "<img src='/Images/Comm/IconAlert.gif' align='absmiddle' style='cursor:hand;' onClick=""fn_Player('"&sdb_RecordingFileName&"');"" title='녹음내용 청취'>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"&rs1("contactdate")&" 녹취 삭제' style='cursor:hand;' align='absmiddle' onClick=RecDel('7','"&rs1("callid")&"','N');>"

								end if
								Rs2.close
							else
								SQL = "select d.RecordingFileName		FROM	I3_IC.dbo.RecordingCall c, I3_IC.dbo.RecordingData d 		WHERE	c.RecordingID = d.RecordingID 		AND	left(recordedcallidkey ,10) = '" & rs1("callid") & "'"

								SET Rs2 = db.execute(SQL)

								if Rs2.eof = false then
									db_RecordingFileName = rs2("RecordingFileName")

									sdb_RecordingFileName = "http://16.1.17.113:8080/"&mid(replace(db_RecordingFileName,"\","/"),27)
									sssdb_RecordingFileName = mid(replace(sdb_RecordingFileName,"\","/"),27)
									db_History_7 = db_History_7 & "&nbsp;<a href='##' onClick=RecDel('7','"&rs1("callid")&"','Y');>녹취첨부</a>"
								end if

								Rs2.close
							end if
						end if
					else
						db_History_7 = db_History_7 & "<br>&nbsp;<a href='##' onClick=HistoryUpdate('7','"&rs1("idx")&"','"&trim(receiptfactnum)&"','"&rs("factPeoplenum")&"');>"&rs1("contacttime")&"&nbsp;("&rs1("contacttelno")&")&nbsp;"&rs1("remark")&"</a>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"&rs1("contactdate")&" 통화 삭제' style='cursor:hand;' align='absmiddle' onClick=HistoryDel('7','"&rs1("idx")&"','"&trim(receiptfactnum)&"','"&rs("factPeoplenum")&"');>"
						if rs1("callid") <> "" then
							if rs1("recordyn") = "Y" then

								SQL = "select d.RecordingFileName		FROM	I3_IC.dbo.RecordingCall c, I3_IC.dbo.RecordingData d 		WHERE	c.RecordingID = d.RecordingID 		AND	left(recordedcallidkey ,10) = '" & rs1("callid") & "'"

								SET Rs2 = db.execute(SQL)
								if Rs2.eof = false then
									db_RecordingFileName = rs2("RecordingFileName")

									sdb_RecordingFileName = "http://16.1.17.113:8080/"&mid(replace(db_RecordingFileName,"\","/"),27)
									sssdb_RecordingFileName = mid(replace(sdb_RecordingFileName,"\","/"),27)
									db_History_7 = db_History_7 & "&nbsp;"& CutString(sssdb_RecordingFileName, 40, "...") & "<img src='/Images/Comm/IconAlert.gif' align='absmiddle' style='cursor:hand;' onClick=""fn_Player('"&sdb_RecordingFileName&"');"" title='녹음내용 청취'>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"&rs1("contactdate")&" 녹취 삭제' style='cursor:hand;' align='absmiddle' onClick=RecDel('7','"&rs1("callid")&"','N');>"

								end if
								Rs2.close
							else
								SQL = "select d.RecordingFileName		FROM	I3_IC.dbo.RecordingCall c, I3_IC.dbo.RecordingData d 		WHERE	c.RecordingID = d.RecordingID 		AND	left(recordedcallidkey ,10) = '" & rs1("callid") & "'"

								SET Rs2 = db.execute(SQL)

								if Rs2.eof = false then
									db_RecordingFileName = rs2("RecordingFileName")

									sdb_RecordingFileName = "http://16.1.17.113:8080/"&mid(replace(db_RecordingFileName,"\","/"),27)
									sssdb_RecordingFileName = mid(replace(sdb_RecordingFileName,"\","/"),27)
									db_History_7 = db_History_7 & "&nbsp;<a href='##' onClick=RecDel('7','"&rs1("callid")&"','Y');>녹취첨부</a>"
								end if

								Rs2.close
							end if
						end if
					end if
					rs1.movenext
				loop
				rs1.close

			elseif i = 8 then
				db_factPeoplenum_8 = rs("factPeoplenum")
				db_Name_8 = rs("name")
				db_level_8 = rs("level")
				db_homephone_8 = rs("homephone")		
				db_mobilephone_8 = rs("mobilephone")	
				db_factpeoplenum_8 = rs("factpeoplenum")	
				db_etcphone_8 = rs("etcphone")		
				db_MONITORDATE_8 = rs("MONITORDATE1")
				db_MONITOR_RESULT_8 = rs("MONITORRESULT")
				db_RESERVEDATE_8 = rs("RESERVEDATE")	
				db_Remark_8 = rs("Remark")	
				db_Remark1_8 = rs("Remark1")
				if rs("monitorpoint") <> "" then
					db_TOT_8 = formatnumber(rs("monitorpoint"),2)
				end if	
				'통화 히스토리 찾기
				sql = " select substring(convert(char(19),contactdate,121),6,11) contacttime, * from	armyinformix.dbo.contactlist where factnum = '" & trim(receiptfactnum) & "' and peoplenum = '" & rs("factPeoplenum") & "' order by idx"
				SET Rs1 = db.execute(SQL)
				do until rs1.eof
					if db_History_8 = "" then
						db_History_8 = "<a href='##' onClick=HistoryUpdate('8','"&rs1("idx")&"','"&trim(receiptfactnum)&"','"&rs("factPeoplenum")&"');>"&rs1("contacttime")&"&nbsp;("&rs1("contacttelno")&")&nbsp;"&rs1("remark")&"</a>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"&rs1("contactdate")&" 통화 삭제' style='cursor:hand;' align='absmiddle' onClick=HistoryDel('8','"&rs1("idx")&"','"&trim(receiptfactnum)&"','"&rs("factPeoplenum")&"');>"

						if rs1("callid") <> "" then
							if rs1("recordyn") = "Y" then

								SQL = "select d.RecordingFileName		FROM	I3_IC.dbo.RecordingCall c, I3_IC.dbo.RecordingData d 		WHERE	c.RecordingID = d.RecordingID 		AND	left(recordedcallidkey ,10) = '" & rs1("callid") & "'"

								SET Rs2 = db.execute(SQL)
								if Rs2.eof = false then
									db_RecordingFileName = rs2("RecordingFileName")

									sdb_RecordingFileName = "http://16.1.17.113:8080/"&mid(replace(db_RecordingFileName,"\","/"),27)
									sssdb_RecordingFileName = mid(replace(sdb_RecordingFileName,"\","/"),27)
									db_History_8 = db_History_8 & "&nbsp;"& CutString(sssdb_RecordingFileName, 40, "...") & "<img src='/Images/Comm/IconAlert.gif' align='absmiddle' style='cursor:hand;' onClick=""fn_Player('"&sdb_RecordingFileName&"');"" title='녹음내용 청취'>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"&rs1("contactdate")&" 녹취 삭제' style='cursor:hand;' align='absmiddle' onClick=RecDel('8','"&rs1("callid")&"','N');>"

								end if
								Rs2.close
							else
								SQL = "select d.RecordingFileName		FROM	I3_IC.dbo.RecordingCall c, I3_IC.dbo.RecordingData d 		WHERE	c.RecordingID = d.RecordingID 		AND	left(recordedcallidkey ,10) = '" & rs1("callid") & "'"

								SET Rs2 = db.execute(SQL)

								if Rs2.eof = false then
									db_RecordingFileName = rs2("RecordingFileName")

									sdb_RecordingFileName = "http://16.1.17.113:8080/"&mid(replace(db_RecordingFileName,"\","/"),27)
									sssdb_RecordingFileName = mid(replace(sdb_RecordingFileName,"\","/"),27)
									db_History_8 = db_History_8 & "&nbsp;<a href='##' onClick=RecDel('8','"&rs1("callid")&"','Y');>녹취첨부</a>"
								end if

								Rs2.close
							end if
						end if
					else
						db_History_8 = db_History_8 & "<br>&nbsp;<a href='##' onClick=HistoryUpdate('8','"&rs1("idx")&"','"&trim(receiptfactnum)&"','"&rs("factPeoplenum")&"');>"&rs1("contacttime")&"&nbsp;("&rs1("contacttelno")&")&nbsp;"&rs1("remark")&"</a>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"&rs1("contactdate")&" 통화 삭제' style='cursor:hand;' align='absmiddle' onClick=HistoryDel('8','"&rs1("idx")&"','"&trim(receiptfactnum)&"','"&rs("factPeoplenum")&"');>"

						if rs1("callid") <> "" then
							if rs1("recordyn") = "Y" then

								SQL = "select d.RecordingFileName		FROM	I3_IC.dbo.RecordingCall c, I3_IC.dbo.RecordingData d 		WHERE	c.RecordingID = d.RecordingID 		AND	left(recordedcallidkey ,10) = '" & rs1("callid") & "'"

								SET Rs2 = db.execute(SQL)
								if Rs2.eof = false then
									db_RecordingFileName = rs2("RecordingFileName")

									sdb_RecordingFileName = "http://16.1.17.113:8080/"&mid(replace(db_RecordingFileName,"\","/"),27)
									sssdb_RecordingFileName = mid(replace(sdb_RecordingFileName,"\","/"),27)
									db_History_8 = db_History_8 & "&nbsp;"& CutString(sssdb_RecordingFileName, 40, "...") & "<img src='/Images/Comm/IconAlert.gif' align='absmiddle' style='cursor:hand;' onClick=""fn_Player('"&sdb_RecordingFileName&"');"" title='녹음내용 청취'>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"&rs1("contactdate")&" 녹취 삭제' style='cursor:hand;' align='absmiddle' onClick=RecDel('8','"&rs1("callid")&"','N');>"

								end if
								Rs2.close
							else
								SQL = "select d.RecordingFileName		FROM	I3_IC.dbo.RecordingCall c, I3_IC.dbo.RecordingData d 		WHERE	c.RecordingID = d.RecordingID 		AND	left(recordedcallidkey ,10) = '" & rs1("callid") & "'"

								SET Rs2 = db.execute(SQL)

								if Rs2.eof = false then
									db_RecordingFileName = rs2("RecordingFileName")

									sdb_RecordingFileName = "http://16.1.17.113:8080/"&mid(replace(db_RecordingFileName,"\","/"),27)
									sssdb_RecordingFileName = mid(replace(sdb_RecordingFileName,"\","/"),27)
									db_History_8 = db_History_8 & "&nbsp;<a href='##' onClick=RecDel('8','"&rs1("callid")&"','Y');>녹취첨부</a>"
								end if

								Rs2.close
							end if
						end if
					end if
					rs1.movenext
				loop
				rs1.close
				
			elseif i = 9 then
				db_factPeoplenum_9 = rs("factPeoplenum")
				db_Name_9 = rs("name")
				db_level_9 = rs("level")
				db_homephone_9 = rs("homephone")		
				db_mobilephone_9 = rs("mobilephone")	
				db_factpeoplenum_9 = rs("factpeoplenum")	
				db_etcphone_9 = rs("etcphone")		
				db_MONITORDATE_9 = rs("MONITORDATE1")
				db_MONITOR_RESULT_9 = rs("MONITORRESULT")
				db_RESERVEDATE_9 = rs("RESERVEDATE")	
				db_Remark_9 = rs("Remark")	
				db_Remark1_9 = rs("Remark1")
				if rs("monitorpoint") <> "" then
					db_TOT_9 = formatnumber(rs("monitorpoint"),2)
				end if

				'통화 히스토리 찾기
				sql = " select substring(convert(char(19),contactdate,121),6,11) contacttime, * from	armyinformix.dbo.contactlist where factnum = '" & trim(receiptfactnum) & "' and peoplenum = '" & rs("factPeoplenum") & "' order by idx"
				SET Rs1 = db.execute(SQL)
				do until rs1.eof
					if db_History_9 = "" then
						db_History_9 = "<a href='##' onClick=HistoryUpdate('9','"&rs1("idx")&"','"&trim(receiptfactnum)&"','"&rs("factPeoplenum")&"');>"&rs1("contacttime")&"&nbsp;("&rs1("contacttelno")&")&nbsp;"&rs1("remark")&"</a>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"&rs1("contactdate")&" 통화 삭제' style='cursor:hand;' align='absmiddle' onClick=HistoryDel('9','"&rs1("idx")&"','"&trim(receiptfactnum)&"','"&rs("factPeoplenum")&"');>"

						if rs1("callid") <> "" then
							if rs1("recordyn") = "Y" then

								SQL = "select d.RecordingFileName		FROM	I3_IC.dbo.RecordingCall c, I3_IC.dbo.RecordingData d 		WHERE	c.RecordingID = d.RecordingID 		AND	left(recordedcallidkey ,10) = '" & rs1("callid") & "'"

								SET Rs2 = db.execute(SQL)
								if Rs2.eof = false then
									db_RecordingFileName = rs2("RecordingFileName")

									sdb_RecordingFileName = "http://16.1.17.113:8080/"&mid(replace(db_RecordingFileName,"\","/"),27)
									sssdb_RecordingFileName = mid(replace(sdb_RecordingFileName,"\","/"),27)
									db_History_9 = db_History_9 & "&nbsp;"& CutString(sssdb_RecordingFileName, 40, "...") & "<img src='/Images/Comm/IconAlert.gif' align='absmiddle' style='cursor:hand;' onClick=""fn_Player('"&sdb_RecordingFileName&"');"" title='녹음내용 청취'>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"&rs1("contactdate")&" 녹취 삭제' style='cursor:hand;' align='absmiddle' onClick=RecDel('9','"&rs1("callid")&"','N');>"

								end if
								Rs2.close
							else
								SQL = "select d.RecordingFileName		FROM	I3_IC.dbo.RecordingCall c, I3_IC.dbo.RecordingData d 		WHERE	c.RecordingID = d.RecordingID 		AND	left(recordedcallidkey ,10) = '" & rs1("callid") & "'"

								SET Rs2 = db.execute(SQL)

								if Rs2.eof = false then
									db_RecordingFileName = rs2("RecordingFileName")

									sdb_RecordingFileName = "http://16.1.17.113:8080/"&mid(replace(db_RecordingFileName,"\","/"),27)
									sssdb_RecordingFileName = mid(replace(sdb_RecordingFileName,"\","/"),27)
									db_History_9 = db_History_9 & "&nbsp;<a href='##' onClick=RecDel('9','"&rs1("callid")&"','Y');>녹취첨부</a>"
								end if

								Rs2.close
							end if
						end if
					else
						db_History_9 = db_History_9 & "<br>&nbsp;<a href='##' onClick=HistoryUpdate('9','"&rs1("idx")&"','"&trim(receiptfactnum)&"','"&rs("factPeoplenum")&"');>"&rs1("contacttime")&"&nbsp;("&rs1("contacttelno")&")&nbsp;"&rs1("remark")&"</a>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"&rs1("contactdate")&" 통화 삭제' style='cursor:hand;' align='absmiddle' onClick=HistoryDel('9','"&rs1("idx")&"','"&trim(receiptfactnum)&"','"&rs("factPeoplenum")&"');>"
						if rs1("callid") <> "" then
							if rs1("recordyn") = "Y" then

								SQL = "select d.RecordingFileName		FROM	I3_IC.dbo.RecordingCall c, I3_IC.dbo.RecordingData d 		WHERE	c.RecordingID = d.RecordingID 		AND	left(recordedcallidkey ,10) = '" & rs1("callid") & "'"

								SET Rs2 = db.execute(SQL)
								if Rs2.eof = false then
									db_RecordingFileName = rs2("RecordingFileName")

									sdb_RecordingFileName = "http://16.1.17.113:8080/"&mid(replace(db_RecordingFileName,"\","/"),27)
									sssdb_RecordingFileName = mid(replace(sdb_RecordingFileName,"\","/"),27)
									db_History_9 = db_History_9 & "&nbsp;"& CutString(sssdb_RecordingFileName, 40, "...") & "<img src='/Images/Comm/IconAlert.gif' align='absmiddle' style='cursor:hand;' onClick=""fn_Player('"&sdb_RecordingFileName&"');"" title='녹음내용 청취'>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"&rs1("contactdate")&" 녹취 삭제' style='cursor:hand;' align='absmiddle' onClick=RecDel('9','"&rs1("callid")&"','N');>"

								end if
								Rs2.close
							else
								SQL = "select d.RecordingFileName		FROM	I3_IC.dbo.RecordingCall c, I3_IC.dbo.RecordingData d 		WHERE	c.RecordingID = d.RecordingID 		AND	left(recordedcallidkey ,10) = '" & rs1("callid") & "'"

								SET Rs2 = db.execute(SQL)

								if Rs2.eof = false then
									db_RecordingFileName = rs2("RecordingFileName")

									sdb_RecordingFileName = "http://16.1.17.113:8080/"&mid(replace(db_RecordingFileName,"\","/"),27)
									sssdb_RecordingFileName = mid(replace(sdb_RecordingFileName,"\","/"),27)
									db_History_9 = db_History_9 & "&nbsp;<a href='##' onClick=RecDel('9','"&rs1("callid")&"','Y');>녹취첨부</a>"
								end if

								Rs2.close
							end if
						end if
					end if
					rs1.movenext
				loop
				rs1.close
				
			elseif i = 10 then
				db_factPeoplenum_10 = rs("factPeoplenum")
				db_Name_10 = rs("name")
				db_level_10 = rs("level")
				db_homephone_10 = rs("homephone")		
				db_mobilephone_10 = rs("mobilephone")	
				db_factpeoplenum_10 = rs("factpeoplenum")	
				db_etcphone_10 = rs("etcphone")		
				db_MONITORDATE_10 = rs("MONITORDATE1")
				db_MONITOR_RESULT_10 = rs("MONITORRESULT")
				db_RESERVEDATE_10 = rs("RESERVEDATE")	
				db_Remark_10 = rs("Remark")	
				db_Remark1_10 = rs("Remark1")
				if rs("monitorpoint") <> "" then
					db_TOT_10 = formatnumber(rs("monitorpoint"),2)
				end if	

				'통화 히스토리 찾기
				sql = " select substring(convert(char(19),contactdate,121),6,11) contacttime, * from	armyinformix.dbo.contactlist where factnum = '" & trim(receiptfactnum) & "' and peoplenum = '" & rs("factPeoplenum") & "' order by idx"
				SET Rs1 = db.execute(SQL)
				do until rs1.eof
					if db_History_10 = "" then
						db_History_10 = "<a href='##' onClick=HistoryUpdate('10','"&rs1("idx")&"','"&trim(receiptfactnum)&"','"&rs("factPeoplenum")&"');>"&rs1("contacttime")&"&nbsp;("&rs1("contacttelno")&")&nbsp;"&rs1("remark")&"</a>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"&rs1("contactdate")&" 통화 삭제' style='cursor:hand;' align='absmiddle' onClick=HistoryDel('10','"&rs1("idx")&"','"&trim(receiptfactnum)&"','"&rs("factPeoplenum")&"');>"
						if rs1("callid") <> "" then
							if rs1("recordyn") = "Y" then

								SQL = "select d.RecordingFileName		FROM	I3_IC.dbo.RecordingCall c, I3_IC.dbo.RecordingData d 		WHERE	c.RecordingID = d.RecordingID 		AND	left(recordedcallidkey ,10) = '" & rs1("callid") & "'"

								SET Rs2 = db.execute(SQL)
								if Rs2.eof = false then
									db_RecordingFileName = rs2("RecordingFileName")

									sdb_RecordingFileName = "http://16.1.17.113:8080/"&mid(replace(db_RecordingFileName,"\","/"),27)
									sssdb_RecordingFileName = mid(replace(sdb_RecordingFileName,"\","/"),27)
									db_History_10 = db_History_10 & "&nbsp;"& CutString(sssdb_RecordingFileName, 40, "...") & "<img src='/Images/Comm/IconAlert.gif' align='absmiddle' style='cursor:hand;' onClick=""fn_Player('"&sdb_RecordingFileName&"');"" title='녹음내용 청취'>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"&rs1("contactdate")&" 녹취 삭제' style='cursor:hand;' align='absmiddle' onClick=RecDel('10','"&rs1("callid")&"','N');>"

								end if
								Rs2.close
							else
								SQL = "select d.RecordingFileName		FROM	I3_IC.dbo.RecordingCall c, I3_IC.dbo.RecordingData d 		WHERE	c.RecordingID = d.RecordingID 		AND	left(recordedcallidkey ,10) = '" & rs1("callid") & "'"

								SET Rs2 = db.execute(SQL)

								if Rs2.eof = false then
									db_RecordingFileName = rs2("RecordingFileName")

									sdb_RecordingFileName = "http://16.1.17.113:8080/"&mid(replace(db_RecordingFileName,"\","/"),27)
									sssdb_RecordingFileName = mid(replace(sdb_RecordingFileName,"\","/"),27)
									db_History_10 = db_History_10 & "&nbsp;<a href='##' onClick=RecDel('10','"&rs1("callid")&"','Y');>녹취첨부</a>"
								end if

								Rs2.close
							end if
						end if
					else
						db_History_10 = db_History_10 & "<br>&nbsp;<a href='##' onClick=HistoryUpdate('10','"&rs1("idx")&"','"&trim(receiptfactnum)&"','"&rs("factPeoplenum")&"');>"&rs1("contacttime")&"&nbsp;("&rs1("contacttelno")&")&nbsp;"&rs1("remark")&"</a>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"&rs1("contactdate")&" 통화 삭제' style='cursor:hand;' align='absmiddle' onClick=HistoryDel('10','"&rs1("idx")&"','"&trim(receiptfactnum)&"','"&rs("factPeoplenum")&"');>"
						if rs1("callid") <> "" then
							if rs1("recordyn") = "Y" then

								SQL = "select d.RecordingFileName		FROM	I3_IC.dbo.RecordingCall c, I3_IC.dbo.RecordingData d 		WHERE	c.RecordingID = d.RecordingID 		AND	left(recordedcallidkey ,10) = '" & rs1("callid") & "'"

								SET Rs2 = db.execute(SQL)
								if Rs2.eof = false then
									db_RecordingFileName = rs2("RecordingFileName")

									sdb_RecordingFileName = "http://16.1.17.113:8080/"&mid(replace(db_RecordingFileName,"\","/"),27)
									sssdb_RecordingFileName = mid(replace(sdb_RecordingFileName,"\","/"),27)
									db_History_10 = db_History_10 & "&nbsp;"& CutString(sssdb_RecordingFileName, 40, "...") & "<img src='/Images/Comm/IconAlert.gif' align='absmiddle' style='cursor:hand;' onClick=""fn_Player('"&sdb_RecordingFileName&"');"" title='녹음내용 청취'>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"&rs1("contactdate")&" 녹취 삭제' style='cursor:hand;' align='absmiddle' onClick=RecDel('10','"&rs1("callid")&"','N');>"

								end if
								Rs2.close
							else
								SQL = "select d.RecordingFileName		FROM	I3_IC.dbo.RecordingCall c, I3_IC.dbo.RecordingData d 		WHERE	c.RecordingID = d.RecordingID 		AND	left(recordedcallidkey ,10) = '" & rs1("callid") & "'"

								SET Rs2 = db.execute(SQL)

								if Rs2.eof = false then
									db_RecordingFileName = rs2("RecordingFileName")

									sdb_RecordingFileName = "http://16.1.17.113:8080/"&mid(replace(db_RecordingFileName,"\","/"),27)
									sssdb_RecordingFileName = mid(replace(sdb_RecordingFileName,"\","/"),27)
									db_History_10 = db_History_10 & "&nbsp;<a href='##' onClick=RecDel('10','"&rs1("callid")&"','Y');>녹취첨부</a>"
								end if

								Rs2.close
							end if
						end if
					end if
					rs1.movenext
				loop
				rs1.close

			elseif i = 11 then
				db_factPeoplenum_11 = rs("factPeoplenum")
				db_Name_11 = rs("name")
				db_level_11 = rs("level")
				db_homephone_11 = rs("homephone")		
				db_mobilephone_11 = rs("mobilephone")	
				db_factpeoplenum_11 = rs("factpeoplenum")	
				db_etcphone_11 = rs("etcphone")		
				db_MONITORDATE_11 = rs("MONITORDATE1")
				db_MONITOR_RESULT_11 = rs("MONITORRESULT")
				db_RESERVEDATE_11 = rs("RESERVEDATE")	
				db_Remark_11 = rs("Remark")	
				db_Remark1_11 = rs("Remark1")
				if rs("monitorpoint") <> "" then
					db_TOT_11 = formatnumber(rs("monitorpoint"),2)
				end if	

				'통화 히스토리 찾기
				sql = " select substring(convert(char(19),contactdate,121),6,11) contacttime, * from	armyinformix.dbo.contactlist where factnum = '" & trim(receiptfactnum) & "' and peoplenum = '" & rs("factPeoplenum") & "' order by idx"
				SET Rs1 = db.execute(SQL)
				do until rs1.eof
					if db_History_11 = "" then
						db_History_11 = "<a href='##' onClick=HistoryUpdate('11','"&rs1("idx")&"','"&trim(receiptfactnum)&"','"&rs("factPeoplenum")&"');>"&rs1("contacttime")&"&nbsp;("&rs1("contacttelno")&")&nbsp;"&rs1("remark")&"</a>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"&rs1("contactdate")&" 통화 삭제' style='cursor:hand;' align='absmiddle' onClick=HistoryDel('11','"&rs1("idx")&"','"&trim(receiptfactnum)&"','"&rs("factPeoplenum")&"');>"

						if rs1("callid") <> "" then
							if rs1("recordyn") = "Y" then

								SQL = "select d.RecordingFileName		FROM	I3_IC.dbo.RecordingCall c, I3_IC.dbo.RecordingData d 		WHERE	c.RecordingID = d.RecordingID 		AND	left(recordedcallidkey ,10) = '" & rs1("callid") & "'"

								SET Rs2 = db.execute(SQL)
								if Rs2.eof = false then
									db_RecordingFileName = rs2("RecordingFileName")

									sdb_RecordingFileName = "http://16.1.17.113:8080/"&mid(replace(db_RecordingFileName,"\","/"),27)
									sssdb_RecordingFileName = mid(replace(sdb_RecordingFileName,"\","/"),27)
									db_History_11 = db_History_11 & "&nbsp;"& CutString(sssdb_RecordingFileName, 40, "...") & "<img src='/Images/Comm/IconAlert.gif' align='absmiddle' style='cursor:hand;' onClick=""fn_Player('"&sdb_RecordingFileName&"');"" title='녹음내용 청취'>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"&rs1("contactdate")&" 녹취 삭제' style='cursor:hand;' align='absmiddle' onClick=RecDel('11','"&rs1("callid")&"','N');>"

								end if
								Rs2.close
							else
								SQL = "select d.RecordingFileName		FROM	I3_IC.dbo.RecordingCall c, I3_IC.dbo.RecordingData d 		WHERE	c.RecordingID = d.RecordingID 		AND	left(recordedcallidkey ,10) = '" & rs1("callid") & "'"

								SET Rs2 = db.execute(SQL)

								if Rs2.eof = false then
									db_RecordingFileName = rs2("RecordingFileName")

									sdb_RecordingFileName = "http://16.1.17.113:8080/"&mid(replace(db_RecordingFileName,"\","/"),27)
									sssdb_RecordingFileName = mid(replace(sdb_RecordingFileName,"\","/"),27)
									db_History_11 = db_History_11 & "&nbsp;<a href='##' onClick=RecDel('11','"&rs1("callid")&"','Y');>녹취첨부</a>"
								end if

								Rs2.close
							end if
						end if
					else
						db_History_11 = db_History_11 & "<br>&nbsp;<a href='##' onClick=HistoryUpdate('11','"&rs1("idx")&"','"&trim(receiptfactnum)&"','"&rs("factPeoplenum")&"');>"&rs1("contacttime")&"&nbsp;("&rs1("contacttelno")&")&nbsp;"&rs1("remark")&"</a>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"&rs1("contactdate")&" 통화 삭제' style='cursor:hand;' align='absmiddle' onClick=HistoryDel('11','"&rs1("idx")&"','"&trim(receiptfactnum)&"','"&rs("factPeoplenum")&"');>"
						if rs1("callid") <> "" then
							if rs1("recordyn") = "Y" then

								SQL = "select d.RecordingFileName		FROM	I3_IC.dbo.RecordingCall c, I3_IC.dbo.RecordingData d 		WHERE	c.RecordingID = d.RecordingID 		AND	left(recordedcallidkey ,10) = '" & rs1("callid") & "'"

								SET Rs2 = db.execute(SQL)
								if Rs2.eof = false then
									db_RecordingFileName = rs2("RecordingFileName")

									sdb_RecordingFileName = "http://16.1.17.113:8080/"&mid(replace(db_RecordingFileName,"\","/"),27)
									sssdb_RecordingFileName = mid(replace(sdb_RecordingFileName,"\","/"),27)
									db_History_11 = db_History_11 & "&nbsp;"& CutString(sssdb_RecordingFileName, 40, "...") & "<img src='/Images/Comm/IconAlert.gif' align='absmiddle' style='cursor:hand;' onClick=""fn_Player('"&sdb_RecordingFileName&"');"" title='녹음내용 청취'>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"&rs1("contactdate")&" 녹취 삭제' style='cursor:hand;' align='absmiddle' onClick=RecDel('11','"&rs1("callid")&"','N');>"

								end if
								Rs2.close
							else
								SQL = "select d.RecordingFileName		FROM	I3_IC.dbo.RecordingCall c, I3_IC.dbo.RecordingData d 		WHERE	c.RecordingID = d.RecordingID 		AND	left(recordedcallidkey ,10) = '" & rs1("callid") & "'"

								SET Rs2 = db.execute(SQL)

								if Rs2.eof = false then
									db_RecordingFileName = rs2("RecordingFileName")

									sdb_RecordingFileName = "http://16.1.17.113:8080/"&mid(replace(db_RecordingFileName,"\","/"),27)
									sssdb_RecordingFileName = mid(replace(sdb_RecordingFileName,"\","/"),27)
									db_History_11 = db_History_11 & "&nbsp;<a href='##' onClick=RecDel('11','"&rs1("callid")&"','Y');>녹취첨부</a>"
								end if

								Rs2.close
							end if
						end if
					end if
					rs1.movenext
				loop
				rs1.close

			elseif i = 12 then
				db_factPeoplenum_12 = rs("factPeoplenum")
				db_Name_12 = rs("name")
				db_level_12 = rs("level")
				db_homephone_12 = rs("homephone")		
				db_mobilephone_12 = rs("mobilephone")	
				db_factpeoplenum_12 = rs("factpeoplenum")	
				db_etcphone_12 = rs("etcphone")		
				db_MONITORDATE_12 = rs("MONITORDATE1")
				db_MONITOR_RESULT_12 = rs("MONITORRESULT")
				db_RESERVEDATE_12 = rs("RESERVEDATE")	
				db_Remark_12 = rs("Remark")	
				db_Remark1_12 = rs("Remark1")
				if rs("monitorpoint") <> "" then
					db_TOT_12 = formatnumber(rs("monitorpoint"),2)
				end if	

				'통화 히스토리 찾기
				sql = " select substring(convert(char(19),contactdate,121),6,11) contacttime, * from	armyinformix.dbo.contactlist where factnum = '" & trim(receiptfactnum) & "' and peoplenum = '" & rs("factPeoplenum") & "' order by idx"
				SET Rs1 = db.execute(SQL)
				do until rs1.eof
					if db_History_12 = "" then
						db_History_12 = "<a href='##' onClick=HistoryUpdate('12','"&rs1("idx")&"','"&trim(receiptfactnum)&"','"&rs("factPeoplenum")&"');>"&rs1("contacttime")&"&nbsp;("&rs1("contacttelno")&")&nbsp;"&rs1("remark")&"</a>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"&rs1("contactdate")&" 통화 삭제' style='cursor:hand;' align='absmiddle' onClick=HistoryDel('12','"&rs1("idx")&"','"&trim(receiptfactnum)&"','"&rs("factPeoplenum")&"');>"
						if rs1("callid") <> "" then
							if rs1("recordyn") = "Y" then

								SQL = "select d.RecordingFileName		FROM	I3_IC.dbo.RecordingCall c, I3_IC.dbo.RecordingData d 		WHERE	c.RecordingID = d.RecordingID 		AND	left(recordedcallidkey ,10) = '" & rs1("callid") & "'"

								SET Rs2 = db.execute(SQL)
								if Rs2.eof = false then
									db_RecordingFileName = rs2("RecordingFileName")

									sdb_RecordingFileName = "http://16.1.17.113:8080/"&mid(replace(db_RecordingFileName,"\","/"),27)
									sssdb_RecordingFileName = mid(replace(sdb_RecordingFileName,"\","/"),27)
									db_History_12 = db_History_12 & "&nbsp;"& CutString(sssdb_RecordingFileName, 40, "...") & "<img src='/Images/Comm/IconAlert.gif' align='absmiddle' style='cursor:hand;' onClick=""fn_Player('"&sdb_RecordingFileName&"');"" title='녹음내용 청취'>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"&rs1("contactdate")&" 녹취 삭제' style='cursor:hand;' align='absmiddle' onClick=RecDel('12','"&rs1("callid")&"','N');>"

								end if
								Rs2.close
							else
								SQL = "select d.RecordingFileName		FROM	I3_IC.dbo.RecordingCall c, I3_IC.dbo.RecordingData d 		WHERE	c.RecordingID = d.RecordingID 		AND	left(recordedcallidkey ,10) = '" & rs1("callid") & "'"

								SET Rs2 = db.execute(SQL)

								if Rs2.eof = false then
									db_RecordingFileName = rs2("RecordingFileName")

									sdb_RecordingFileName = "http://16.1.17.113:8080/"&mid(replace(db_RecordingFileName,"\","/"),27)
									sssdb_RecordingFileName = mid(replace(sdb_RecordingFileName,"\","/"),27)
									db_History_12 = db_History_12 & "&nbsp;<a href='##' onClick=RecDel('12','"&rs1("callid")&"','Y');>녹취첨부</a>"
								end if

								Rs2.close
							end if
						end if
					else
						db_History_12 = db_History_12 & "<br>&nbsp;<a href='##' onClick=HistoryUpdate('12','"&rs1("idx")&"','"&trim(receiptfactnum)&"','"&rs("factPeoplenum")&"');>"&rs1("contacttime")&"&nbsp;("&rs1("contacttelno")&")&nbsp;"&rs1("remark")&"</a>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"&rs1("contactdate")&" 통화 삭제' style='cursor:hand;' align='absmiddle' onClick=HistoryDel('12','"&rs1("idx")&"','"&trim(receiptfactnum)&"','"&rs("factPeoplenum")&"');>"
						if rs1("callid") <> "" then
							if rs1("recordyn") = "Y" then

								SQL = "select d.RecordingFileName		FROM	I3_IC.dbo.RecordingCall c, I3_IC.dbo.RecordingData d 		WHERE	c.RecordingID = d.RecordingID 		AND	left(recordedcallidkey ,10) = '" & rs1("callid") & "'"

								SET Rs2 = db.execute(SQL)
								if Rs2.eof = false then
									db_RecordingFileName = rs2("RecordingFileName")

									sdb_RecordingFileName = "http://16.1.17.113:8080/"&mid(replace(db_RecordingFileName,"\","/"),27)
									sssdb_RecordingFileName = mid(replace(sdb_RecordingFileName,"\","/"),27)
									db_History_12 = db_History_12 & "&nbsp;"& CutString(sssdb_RecordingFileName, 40, "...") & "<img src='/Images/Comm/IconAlert.gif' align='absmiddle' style='cursor:hand;' onClick=""fn_Player('"&sdb_RecordingFileName&"');"" title='녹음내용 청취'>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"&rs1("contactdate")&" 녹취 삭제' style='cursor:hand;' align='absmiddle' onClick=RecDel('12','"&rs1("callid")&"','N');>"

								end if
								Rs2.close
							else
								SQL = "select d.RecordingFileName		FROM	I3_IC.dbo.RecordingCall c, I3_IC.dbo.RecordingData d 		WHERE	c.RecordingID = d.RecordingID 		AND	left(recordedcallidkey ,10) = '" & rs1("callid") & "'"

								SET Rs2 = db.execute(SQL)

								if Rs2.eof = false then
									db_RecordingFileName = rs2("RecordingFileName")

									sdb_RecordingFileName = "http://16.1.17.113:8080/"&mid(replace(db_RecordingFileName,"\","/"),27)
									sssdb_RecordingFileName = mid(replace(sdb_RecordingFileName,"\","/"),27)
									db_History_12 = db_History_12 & "&nbsp;<a href='##' onClick=RecDel('12','"&rs1("callid")&"','Y');>녹취첨부</a>"
								end if

								Rs2.close
							end if
						end if
					end if
					rs1.movenext
				loop
				rs1.close

			elseif i = 13 then
				db_factPeoplenum_13 = rs("factPeoplenum")
				db_Name_13 = rs("name")
				db_level_13 = rs("level")
				db_homephone_13 = rs("homephone")		
				db_mobilephone_13 = rs("mobilephone")	
				db_factpeoplenum_13 = rs("factpeoplenum")	
				db_etcphone_13 = rs("etcphone")		
				db_MONITORDATE_13 = rs("MONITORDATE1")
				db_MONITOR_RESULT_13 = rs("MONITORRESULT")
				db_RESERVEDATE_13 = rs("RESERVEDATE")	
				db_Remark_13 = rs("Remark")	
				db_Remark1_13 = rs("Remark1")
				if rs("monitorpoint") <> "" then
					db_TOT_13 = formatnumber(rs("monitorpoint"),2)
				end if	

				'통화 히스토리 찾기
				sql = " select substring(convert(char(19),contactdate,121),6,11) contacttime, * from	armyinformix.dbo.contactlist where factnum = '" & trim(receiptfactnum) & "' and peoplenum = '" & rs("factPeoplenum") & "' order by idx"
				SET Rs1 = db.execute(SQL)
				do until rs1.eof
					if db_History_13 = "" then
						db_History_13 = "<a href='##' onClick=HistoryUpdate('13','"&rs1("idx")&"','"&trim(receiptfactnum)&"','"&rs("factPeoplenum")&"');>"&rs1("contacttime")&"&nbsp;("&rs1("contacttelno")&")&nbsp;"&rs1("remark")&"</a>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"&rs1("contactdate")&" 통화 삭제' style='cursor:hand;' align='absmiddle' onClick=HistoryDel('13','"&rs1("idx")&"','"&trim(receiptfactnum)&"','"&rs("factPeoplenum")&"');>"
						if rs1("callid") <> "" then
							if rs1("recordyn") = "Y" then

								SQL = "select d.RecordingFileName		FROM	I3_IC.dbo.RecordingCall c, I3_IC.dbo.RecordingData d 		WHERE	c.RecordingID = d.RecordingID 		AND	left(recordedcallidkey ,10) = '" & rs1("callid") & "'"

								SET Rs2 = db.execute(SQL)
								if Rs2.eof = false then
									db_RecordingFileName = rs2("RecordingFileName")

									sdb_RecordingFileName = "http://16.1.17.113:8080/"&mid(replace(db_RecordingFileName,"\","/"),27)
									sssdb_RecordingFileName = mid(replace(sdb_RecordingFileName,"\","/"),27)
									db_History_13 = db_History_13 & "&nbsp;"& CutString(sssdb_RecordingFileName, 40, "...") & "<img src='/Images/Comm/IconAlert.gif' align='absmiddle' style='cursor:hand;' onClick=""fn_Player('"&sdb_RecordingFileName&"');"" title='녹음내용 청취'>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"&rs1("contactdate")&" 녹취 삭제' style='cursor:hand;' align='absmiddle' onClick=RecDel('13','"&rs1("callid")&"','N');>"

								end if
								Rs2.close
							else
								SQL = "select d.RecordingFileName		FROM	I3_IC.dbo.RecordingCall c, I3_IC.dbo.RecordingData d 		WHERE	c.RecordingID = d.RecordingID 		AND	left(recordedcallidkey ,10) = '" & rs1("callid") & "'"

								SET Rs2 = db.execute(SQL)

								if Rs2.eof = false then
									db_RecordingFileName = rs2("RecordingFileName")

									sdb_RecordingFileName = "http://16.1.17.113:8080/"&mid(replace(db_RecordingFileName,"\","/"),27)
									sssdb_RecordingFileName = mid(replace(sdb_RecordingFileName,"\","/"),27)
									db_History_13 = db_History_13 & "&nbsp;<a href='##' onClick=RecDel('13','"&rs1("callid")&"','Y');>녹취첨부</a>"
								end if

								Rs2.close
							end if
						end if
					else
						db_History_13 = db_History_13 & "<br>&nbsp;<a href='##' onClick=HistoryUpdate('13','"&rs1("idx")&"','"&trim(receiptfactnum)&"','"&rs("factPeoplenum")&"');>"&rs1("contacttime")&"&nbsp;("&rs1("contacttelno")&")&nbsp;"&rs1("remark")&"</a>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"&rs1("contactdate")&" 통화 삭제' style='cursor:hand;' align='absmiddle' onClick=HistoryDel('13','"&rs1("idx")&"','"&trim(receiptfactnum)&"','"&rs("factPeoplenum")&"');>"
						if rs1("callid") <> "" then
							if rs1("recordyn") = "Y" then

								SQL = "select d.RecordingFileName		FROM	I3_IC.dbo.RecordingCall c, I3_IC.dbo.RecordingData d 		WHERE	c.RecordingID = d.RecordingID 		AND	left(recordedcallidkey ,10) = '" & rs1("callid") & "'"

								SET Rs2 = db.execute(SQL)
								if Rs2.eof = false then
									db_RecordingFileName = rs2("RecordingFileName")

									sdb_RecordingFileName = "http://16.1.17.113:8080/"&mid(replace(db_RecordingFileName,"\","/"),27)
									sssdb_RecordingFileName = mid(replace(sdb_RecordingFileName,"\","/"),27)
									db_History_13 = db_History_13 & "&nbsp;"& CutString(sssdb_RecordingFileName, 40, "...") & "<img src='/Images/Comm/IconAlert.gif' align='absmiddle' style='cursor:hand;' onClick=""fn_Player('"&sdb_RecordingFileName&"');"" title='녹음내용 청취'>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"&rs1("contactdate")&" 녹취 삭제' style='cursor:hand;' align='absmiddle' onClick=RecDel('13','"&rs1("callid")&"','N');>"

								end if
								Rs2.close
							else
								SQL = "select d.RecordingFileName		FROM	I3_IC.dbo.RecordingCall c, I3_IC.dbo.RecordingData d 		WHERE	c.RecordingID = d.RecordingID 		AND	left(recordedcallidkey ,10) = '" & rs1("callid") & "'"

								SET Rs2 = db.execute(SQL)

								if Rs2.eof = false then
									db_RecordingFileName = rs2("RecordingFileName")

									sdb_RecordingFileName = "http://16.1.17.113:8080/"&mid(replace(db_RecordingFileName,"\","/"),27)
									sssdb_RecordingFileName = mid(replace(sdb_RecordingFileName,"\","/"),27)
									db_History_13 = db_History_13 & "&nbsp;<a href='##' onClick=RecDel('13','"&rs1("callid")&"','Y');>녹취첨부</a>"
								end if

								Rs2.close
							end if
						end if
					end if
					rs1.movenext
				loop
				rs1.close
			elseif i = 14 then
				db_factPeoplenum_14 = rs("factPeoplenum")
				db_Name_14 = rs("name")
				db_level_14 = rs("level")
				db_homephone_14 = rs("homephone")		
				db_mobilephone_14 = rs("mobilephone")	
				db_factpeoplenum_14 = rs("factpeoplenum")	
				db_etcphone_14 = rs("etcphone")		
				db_MONITORDATE_14 = rs("MONITORDATE1")
				db_MONITOR_RESULT_14 = rs("MONITORRESULT")
				db_RESERVEDATE_14 = rs("RESERVEDATE")	
				db_Remark_14 = rs("Remark")	
				db_Remark1_14 = rs("Remark1")
				if rs("monitorpoint") <> "" then
					db_TOT_14 = formatnumber(rs("monitorpoint"),2)
				end if	

				'통화 히스토리 찾기
				sql = " select substring(convert(char(19),contactdate,121),6,11) contacttime, * from	armyinformix.dbo.contactlist where factnum = '" & trim(receiptfactnum) & "' and peoplenum = '" & rs("factPeoplenum") & "' order by idx"
				SET Rs1 = db.execute(SQL)
				do until rs1.eof
					if db_History_14 = "" then
						db_History_14 = "<a href='##' onClick=HistoryUpdate('14','"&rs1("idx")&"','"&trim(receiptfactnum)&"','"&rs("factPeoplenum")&"');>"&rs1("contacttime")&"&nbsp;("&rs1("contacttelno")&")&nbsp;"&rs1("remark")&"</a>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"&rs1("contactdate")&" 통화 삭제' style='cursor:hand;' align='absmiddle' onClick=HistoryDel('14','"&rs1("idx")&"','"&trim(receiptfactnum)&"','"&rs("factPeoplenum")&"');>"
						if rs1("callid") <> "" then
							if rs1("recordyn") = "Y" then

								SQL = "select d.RecordingFileName		FROM	I3_IC.dbo.RecordingCall c, I3_IC.dbo.RecordingData d 		WHERE	c.RecordingID = d.RecordingID 		AND	left(recordedcallidkey ,10) = '" & rs1("callid") & "'"

								SET Rs2 = db.execute(SQL)
								if Rs2.eof = false then
									db_RecordingFileName = rs2("RecordingFileName")

									sdb_RecordingFileName = "http://16.1.17.113:8080/"&mid(replace(db_RecordingFileName,"\","/"),27)
									sssdb_RecordingFileName = mid(replace(sdb_RecordingFileName,"\","/"),27)
									db_History_14 = db_History_14 & "&nbsp;"& CutString(sssdb_RecordingFileName, 40, "...") & "<img src='/Images/Comm/IconAlert.gif' align='absmiddle' style='cursor:hand;' onClick=""fn_Player('"&sdb_RecordingFileName&"');"" title='녹음내용 청취'>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"&rs1("contactdate")&" 녹취 삭제' style='cursor:hand;' align='absmiddle' onClick=RecDel('14','"&rs1("callid")&"','N');>"

								end if
								Rs2.close
							else
								SQL = "select d.RecordingFileName		FROM	I3_IC.dbo.RecordingCall c, I3_IC.dbo.RecordingData d 		WHERE	c.RecordingID = d.RecordingID 		AND	left(recordedcallidkey ,10) = '" & rs1("callid") & "'"

								SET Rs2 = db.execute(SQL)

								if Rs2.eof = false then
									db_RecordingFileName = rs2("RecordingFileName")

									sdb_RecordingFileName = "http://16.1.17.113:8080/"&mid(replace(db_RecordingFileName,"\","/"),27)
									sssdb_RecordingFileName = mid(replace(sdb_RecordingFileName,"\","/"),27)
									db_History_14 = db_History_14 & "&nbsp;<a href='##' onClick=RecDel('14','"&rs1("callid")&"','Y');>녹취첨부</a>"
								end if

								Rs2.close
							end if
						end if
					else
						db_History_14 = db_History_14 & "<br>&nbsp;<a href='##' onClick=HistoryUpdate('14','"&rs1("idx")&"','"&trim(receiptfactnum)&"','"&rs("factPeoplenum")&"');>"&rs1("contacttime")&"&nbsp;("&rs1("contacttelno")&")&nbsp;"&rs1("remark")&"</a>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"&rs1("contactdate")&" 통화 삭제' style='cursor:hand;' align='absmiddle' onClick=HistoryDel('14','"&rs1("idx")&"','"&trim(receiptfactnum)&"','"&rs("factPeoplenum")&"');>"
						if rs1("callid") <> "" then
							if rs1("recordyn") = "Y" then

								SQL = "select d.RecordingFileName		FROM	I3_IC.dbo.RecordingCall c, I3_IC.dbo.RecordingData d 		WHERE	c.RecordingID = d.RecordingID 		AND	left(recordedcallidkey ,10) = '" & rs1("callid") & "'"

								SET Rs2 = db.execute(SQL)
								if Rs2.eof = false then
									db_RecordingFileName = rs2("RecordingFileName")

									sdb_RecordingFileName = "http://16.1.17.113:8080/"&mid(replace(db_RecordingFileName,"\","/"),27)
									sssdb_RecordingFileName = mid(replace(sdb_RecordingFileName,"\","/"),27)
									db_History_14 = db_History_14 & "&nbsp;"& CutString(sssdb_RecordingFileName, 40, "...") & "<img src='/Images/Comm/IconAlert.gif' align='absmiddle' style='cursor:hand;' onClick=""fn_Player('"&sdb_RecordingFileName&"');"" title='녹음내용 청취'>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"&rs1("contactdate")&" 녹취 삭제' style='cursor:hand;' align='absmiddle' onClick=RecDel('14','"&rs1("callid")&"','N');>"

								end if
								Rs2.close
							else
								SQL = "select d.RecordingFileName		FROM	I3_IC.dbo.RecordingCall c, I3_IC.dbo.RecordingData d 		WHERE	c.RecordingID = d.RecordingID 		AND	left(recordedcallidkey ,10) = '" & rs1("callid") & "'"

								SET Rs2 = db.execute(SQL)

								if Rs2.eof = false then
									db_RecordingFileName = rs2("RecordingFileName")

									sdb_RecordingFileName = "http://16.1.17.113:8080/"&mid(replace(db_RecordingFileName,"\","/"),27)
									sssdb_RecordingFileName = mid(replace(sdb_RecordingFileName,"\","/"),27)
									db_History_14 = db_History_14 & "&nbsp;<a href='##' onClick=RecDel('14','"&rs1("callid")&"','Y');>녹취첨부</a>"
								end if

								Rs2.close
							end if
						end if
					end if
					rs1.movenext
				loop
				rs1.close

			elseif i = 15 then
				db_factPeoplenum_15 = rs("factPeoplenum")
				db_Name_15 = rs("name")
				db_level_15 = rs("level")
				db_homephone_15 = rs("homephone")		
				db_mobilephone_15 = rs("mobilephone")	
				db_factpeoplenum_15 = rs("factpeoplenum")	
				db_etcphone_15 = rs("etcphone")		
				db_MONITORDATE_15 = rs("MONITORDATE1")
				db_MONITOR_RESULT_15 = rs("MONITORRESULT")
				db_RESERVEDATE_15 = rs("RESERVEDATE")	
				db_Remark_15 = rs("Remark")	
				db_Remark1_15 = rs("Remark1")
				if rs("monitorpoint") <> "" then
					db_TOT_15 = formatnumber(rs("monitorpoint"),2)
				end if	
				
				'통화 히스토리 찾기
				sql = " select substring(convert(char(19),contactdate,121),6,11) contacttime, * from	armyinformix.dbo.contactlist where factnum = '" & trim(receiptfactnum) & "' and peoplenum = '" & rs("factPeoplenum") & "' order by idx"
				SET Rs1 = db.execute(SQL)
				do until rs1.eof
					if db_History_15 = "" then
						db_History_15 = "<a href='##' onClick=HistoryUpdate('15','"&rs1("idx")&"','"&trim(receiptfactnum)&"','"&rs("factPeoplenum")&"');>"&rs1("contacttime")&"&nbsp;("&rs1("contacttelno")&")&nbsp;"&rs1("remark")&"</a>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"&rs1("contactdate")&" 통화 삭제' style='cursor:hand;' align='absmiddle' onClick=HistoryDel('15','"&rs1("idx")&"','"&trim(receiptfactnum)&"','"&rs("factPeoplenum")&"');>"
						if rs1("callid") <> "" then
							if rs1("recordyn") = "Y" then

								SQL = "select d.RecordingFileName		FROM	I3_IC.dbo.RecordingCall c, I3_IC.dbo.RecordingData d 		WHERE	c.RecordingID = d.RecordingID 		AND	left(recordedcallidkey ,10) = '" & rs1("callid") & "'"

								SET Rs2 = db.execute(SQL)
								if Rs2.eof = false then
									db_RecordingFileName = rs2("RecordingFileName")

									sdb_RecordingFileName = "http://16.1.17.113:8080/"&mid(replace(db_RecordingFileName,"\","/"),27)
									sssdb_RecordingFileName = mid(replace(sdb_RecordingFileName,"\","/"),27)
									db_History_15 = db_History_15 & "&nbsp;"& CutString(sssdb_RecordingFileName, 40, "...") & "<img src='/Images/Comm/IconAlert.gif' align='absmiddle' style='cursor:hand;' onClick=""fn_Player('"&sdb_RecordingFileName&"');"" title='녹음내용 청취'>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"&rs1("contactdate")&" 녹취 삭제' style='cursor:hand;' align='absmiddle' onClick=RecDel('15','"&rs1("callid")&"','N');>"

								end if
								Rs2.close
							else
								SQL = "select d.RecordingFileName		FROM	I3_IC.dbo.RecordingCall c, I3_IC.dbo.RecordingData d 		WHERE	c.RecordingID = d.RecordingID 		AND	left(recordedcallidkey ,10) = '" & rs1("callid") & "'"

								SET Rs2 = db.execute(SQL)

								if Rs2.eof = false then
									db_RecordingFileName = rs2("RecordingFileName")

									sdb_RecordingFileName = "http://16.1.17.113:8080/"&mid(replace(db_RecordingFileName,"\","/"),27)
									sssdb_RecordingFileName = mid(replace(sdb_RecordingFileName,"\","/"),27)
									db_History_15 = db_History_15 & "&nbsp;<a href='##' onClick=RecDel('15','"&rs1("callid")&"','Y');>녹취첨부</a>"
								end if

								Rs2.close
							end if
						end if
					else
						db_History_15 = db_History_15 & "<br>&nbsp;<a href='##' onClick=HistoryUpdate('15','"&rs1("idx")&"','"&trim(receiptfactnum)&"','"&rs("factPeoplenum")&"');>"&rs1("contacttime")&"&nbsp;("&rs1("contacttelno")&")&nbsp;"&rs1("remark")&"</a>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"&rs1("contactdate")&" 통화 삭제' style='cursor:hand;' align='absmiddle' onClick=HistoryDel('15','"&rs1("idx")&"','"&trim(receiptfactnum)&"','"&rs("factPeoplenum")&"');>"
						if rs1("callid") <> "" then
							if rs1("recordyn") = "Y" then

								SQL = "select d.RecordingFileName		FROM	I3_IC.dbo.RecordingCall c, I3_IC.dbo.RecordingData d 		WHERE	c.RecordingID = d.RecordingID 		AND	left(recordedcallidkey ,10) = '" & rs1("callid") & "'"

								SET Rs2 = db.execute(SQL)
								if Rs2.eof = false then
									db_RecordingFileName = rs2("RecordingFileName")

									sdb_RecordingFileName = "http://16.1.17.113:8080/"&mid(replace(db_RecordingFileName,"\","/"),27)
									sssdb_RecordingFileName = mid(replace(sdb_RecordingFileName,"\","/"),27)
									db_History_15 = db_History_15 & "&nbsp;"& CutString(sssdb_RecordingFileName, 40, "...") & "<img src='/Images/Comm/IconAlert.gif' align='absmiddle' style='cursor:hand;' onClick=""fn_Player('"&sdb_RecordingFileName&"');"" title='녹음내용 청취'>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"&rs1("contactdate")&" 녹취 삭제' style='cursor:hand;' align='absmiddle' onClick=RecDel('15','"&rs1("callid")&"','N');>"

								end if
								Rs2.close
							else
								SQL = "select d.RecordingFileName		FROM	I3_IC.dbo.RecordingCall c, I3_IC.dbo.RecordingData d 		WHERE	c.RecordingID = d.RecordingID 		AND	left(recordedcallidkey ,10) = '" & rs1("callid") & "'"

								SET Rs2 = db.execute(SQL)

								if Rs2.eof = false then
									db_RecordingFileName = rs2("RecordingFileName")

									sdb_RecordingFileName = "http://16.1.17.113:8080/"&mid(replace(db_RecordingFileName,"\","/"),27)
									sssdb_RecordingFileName = mid(replace(sdb_RecordingFileName,"\","/"),27)
									db_History_15 = db_History_15 & "&nbsp;<a href='##' onClick=RecDel('15','"&rs1("callid")&"','Y');>녹취첨부</a>"
								end if

								Rs2.close
							end if
						end if
					end if
					rs1.movenext
				loop
				rs1.close

			end if			

			Rs.movenext
		loop

	end if
%>			

<div id='CalendarLayer' style='display:none;'><iframe name='CalendarFrame' src='/include/Calendar.asp' width='172' height='177' border='0' frameborder='0' scrolling='no'></iframe></div>
<form name="ListForm" method="post" style="margin:0" action="/menu01/submenu0101/monitoring_input_InsUpDel.asp">
<input type="hidden" name="IsType" value="INS">
<input type="hidden" name="RECEIPTFACTNUM" value="<%=receiptfactnum%>">
<input type="hidden" name="factPeoplenum_1" value="<%=db_factPeoplenum_1%>">
<input type="hidden" name="FRM1" value="ON">
<input type="hidden" name="factPeoplenum_2" value="<%=db_factPeoplenum_2%>">
<% if db_factPeoplenum_2 <> "" then%>
	<input type="hidden" name="FRM2" value="ON">
<% else %>
	<input type="hidden" name="FRM2" value="">
<% end if%>
<input type="hidden" name="factPeoplenum_3" value="<%=db_factPeoplenum_3%>">
<% if db_factPeoplenum_3 <> "" then%>
	<input type="hidden" name="FRM3" value="ON">
<% else %>
	<input type="hidden" name="FRM3" value="">
<% end if%>
<input type="hidden" name="factPeoplenum_4" value="<%=db_factPeoplenum_4%>">
<% if db_factPeoplenum_4 <> "" then%>
	<input type="hidden" name="FRM4" value="ON">
<% else %>
	<input type="hidden" name="FRM4" value="">
<% end if%>
<input type="hidden" name="factPeoplenum_5" value="<%=db_factPeoplenum_5%>">
<% if db_factPeoplenum_5 <> "" then%>
	<input type="hidden" name="FRM5" value="ON">
<% else %>
	<input type="hidden" name="FRM5" value="">
<% end if%>
<input type="hidden" name="factPeoplenum_6" value="<%=db_factPeoplenum_6%>">
<% if db_factPeoplenum_6 <> "" then%>
	<input type="hidden" name="FRM6" value="ON">
<% else %>
	<input type="hidden" name="FRM6" value="">
<% end if%>
<input type="hidden" name="factPeoplenum_7" value="<%=db_factPeoplenum_7%>">
<% if db_factPeoplenum_7 <> "" then%>
	<input type="hidden" name="FRM7" value="ON">
<% else %>
	<input type="hidden" name="FRM7" value="">
<% end if%>
<input type="hidden" name="factPeoplenum_8" value="<%=db_factPeoplenum_8%>">
<% if db_factPeoplenum_8 <> "" then%>
	<input type="hidden" name="FRM8" value="ON">
<% else %>
	<input type="hidden" name="FRM8" value="">
<% end if%>
<input type="hidden" name="factPeoplenum_9" value="<%=db_factPeoplenum_9%>">
<% if db_factPeoplenum_9 <> "" then%>
	<input type="hidden" name="FRM9" value="ON">
<% else %>
	<input type="hidden" name="FRM9" value="">
<% end if%>
<input type="hidden" name="factPeoplenum_10" value="<%=db_factPeoplenum_10%>">
<% if db_factPeoplenum_10 <> "" then%>
	<input type="hidden" name="FRM10" value="ON">
<% else %>
	<input type="hidden" name="FRM10" value="">
<% end if%>
<input type="hidden" name="factPeoplenum_11" value="<%=db_factPeoplenum_11%>">
<% if db_factPeoplenum_11 <> "" then%>
	<input type="hidden" name="FRM11" value="ON">
<% else %>
	<input type="hidden" name="FRM11" value="">
<% end if%>
<input type="hidden" name="factPeoplenum_12" value="<%=db_factPeoplenum_12%>">
<% if db_factPeoplenum_12 <> "" then%>
	<input type="hidden" name="FRM12" value="ON">
<% else %>
	<input type="hidden" name="FRM12" value="">
<% end if%>
<input type="hidden" name="factPeoplenum_13" value="<%=db_factPeoplenum_13%>">
<% if db_factPeoplenum_13 <> "" then%>
	<input type="hidden" name="FRM13" value="ON">
<% else %>
	<input type="hidden" name="FRM13" value="">
<% end if%>
<input type="hidden" name="factPeoplenum_14" value="<%=db_factPeoplenum_14%>">
<% if db_factPeoplenum_14 <> "" then%>
	<input type="hidden" name="FRM14" value="ON">
<% else %>
	<input type="hidden" name="FRM14" value="">
<% end if%>
<input type="hidden" name="factPeoplenum_15" value="<%=db_factPeoplenum_15%>">
<% if db_factPeoplenum_15 <> "" then%>
	<input type="hidden" name="FRM15" value="ON">
<% else %>
	<input type="hidden" name="FRM15" value="">
<% end if%>
<input type="hidden" name="Date2">
<input type="hidden" name="Date3">
<input type="hidden" name="receiptkind">
<input type="hidden" name="idx_1" value="">
<input type="hidden" name="idx_2" value="">
<input type="hidden" name="idx_3" value="">
<input type="hidden" name="idx_4" value="">
<input type="hidden" name="idx_5" value="">
<input type="hidden" name="idx_6" value="">
<input type="hidden" name="idx_7" value="">
<input type="hidden" name="idx_8" value="">
<input type="hidden" name="idx_9" value="">
<input type="hidden" name="idx_10" value="">
<input type="hidden" name="idx_11" value="">
<input type="hidden" name="idx_12" value="">
<input type="hidden" name="idx_13" value="">
<input type="hidden" name="idx_14" value="">
<input type="hidden" name="idx_15" value="">

<%'====== 상담접수 폼 #1 시작 =======================================================================================%>
<span id="divFORM1" style="display:block;">

<table border="0" width="940" cellspacing="1" cellpadding="1" bgcolor="#EFECE5" align="center">
	<tr>
		<td background="/Images/AsRegi/TabBG.gif" bgcolor="#ffffff" align='right'>
		
			<input type="button" name="btn1" value="피의자(<%=cnt1%>-<%=cnt11%>)" style="width:100; height:20%;" <%if cnt11 <>"0" then%> class="Btn3"<%else%>class="Btn14"<%end if%> onClick="javascript:fn_Tabclick('1')">
			<input type="button" name="btn2" value="피해자(<%=cnt2%>-<%=cnt21%>)" style="width:100; height:20%;" <%if cnt21 >0 then%> class="Btn3"<%else%>class="Btn14"<%end if%> onClick="javascript:fn_Tabclick('2')">
			<input type="button" name="btn3" value="민원인(<%=cnt3%>-<%=cnt31%>)" style="width:100; height:20%;" <%if cnt31 >0 then%> class="Btn3"<%else%>class="Btn14"<%end if%> onClick="javascript:fn_Tabclick('3')">
			<input type="button" name="btn4" value="피민원인(<%=cnt4%>-<%=cnt41%>)" style="width:100; height:20%;" <%if cnt41 >0 then%> class="Btn3"<%else%>class="Btn14"<%end if%> onClick="javascript:fn_Tabclick('4')">
			<input type="button" name="btn5" value="지휘관(<%=cnt5%>-<%=cnt51%>)" style="width:100; height:20%;" <%if cnt51 >0 then%> class="Btn3"<%else%>class="Btn14"<%end if%> onClick="javascript:fn_Tabclick('5')">
			<input type="button" name="btn6" value=" 유 족 (<%=cnt6%>-<%=cnt61%>)" style="width:100; height:20%;" <%if cnt61 >0 then%> class="Btn3"<%else%>class="Btn14"<%end if%> onClick="javascript:fn_Tabclick('6')">
			<input type="button" name="btn7" value="참고인(<%=cnt7%>-<%=cnt71%>)" style="width:100; height:20%;" <%if cnt71 >0 then%> class="Btn3"<%else%>class="Btn14"<%end if%> onClick="javascript:fn_Tabclick('7')">
		</td>
	</tr>
</table>


<table border="0" width="920" cellpadding="0" cellspacing="1" bgcolor="#EFECE5" align="center">
	<tr bgcolor="#FFFFFF">
		<td>
			<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table width="920" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#ffffff">
			    <tr>
					<td align="left" bgcolor="#FFFFFF" class="TDCont">&nbsp;<img src="/Images/dot_01.gif" >&nbsp;<b><font color="#ff00ff">(<%=sGubunName%>)</font> 설문지#1</b></td>
				</tr>
			</table>
			<table width="920" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">
			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>관 계</td>
					<td bgcolor="#FFFFFF">						<%
							'======= 처리구분 코드 가져오기 ==================================================
							SqlCode = "SELECT CODE, CODENAME FROM TB_CODE"
							SqlCode = SqlCode& " WHERE USEYN='Y' AND SYSYN='Y' AND CODEGROUP='B01'"
							SqlCode = SqlCode& " ORDER BY CODE"
							set RsCode = db.execute(SqlCode)
						%>
						&nbsp;<select name="SECTION2_1" size="1" class="ComboFFFCE7" onChange="fn_UpdateData('1','SECTION2_');">
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("CODE")
										CODENAME = RsCode("CODENAME")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &sGubun& "")%>
							<%
									RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
							%>
						</select>
					&nbsp;&nbsp;&nbsp;&nbsp;<font color="#0000ff"><%=db_factPeoplenum_1%></font></td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>계급(신분)</td>
					<td bgcolor="#FFFFFF"><%
							'======= 처리구분 코드 가져오기 ==================================================
							SqlCode = "SELECT *	FROM armyinformix.dbo.data"
							SqlCode = SqlCode& " where [restrict] = '104' order by [restrict]"
							set RsCode = db.execute(SqlCode)
						%>
						&nbsp;<select name="LEVEL_1" size="1" class="ComboFFFCE7">
							<option value="">계급선택</option>
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("number")
										CODENAME = RsCode("contents")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &db_level_1& "")%>
							<%
									RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
							%>
						</select></td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>성  명</td>
					<td bgcolor="#FFFFFF">&nbsp;<input type="text" name="NAME_1" value="<%=db_NAME_1%>" size="20" maxlength="20" style="border-width:1px ; border-color:#cccccc ; border-style:solid" ></td>
				</tr>
			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>연락처 1</td>
					<td bgcolor="#FFFFFF">&nbsp;<input type="text" name="HOMEPHONE_1" value="<%=db_HOMEPHONE_1%>" size="13" maxlength="13" style="border-width:1px ; border-color:#cccccc ; border-style:solid">
					<img src="/Images/Cti/icon_tel.gif" style="cursor:hand;" onClick="fn_dial('1','1');" align="absmiddle" title="군전화로 전화걸기">&nbsp<img src="/Images/Cti/icon_tel_1.gif" style="cursor:hand;" onClick="fn_dial_1('1','1');" align="absmiddle" title="일반전화로 전화걸기">&nbsp<img src="/Images/Cti/icon_sche.gif" style="cursor:hand;" onClick="fn_sms('1','1');" align="absmiddle" title="문자전송"></td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>연락처 2</td>
					<td bgcolor="#FFFFFF">&nbsp;<input type="text" name="MOBILEPHONE_1" value="<%=db_MOBILEPHONE_1%>" size="13" maxlength="13" style="border-width:1px ; border-color:#cccccc ; border-style:solid">&nbsp;<img src="/Images/Cti/icon_tel.gif" style="cursor:hand;" onClick="fn_dial('2','1');" align="absmiddle" title="군전화로 전화걸기">&nbsp<img src="/Images/Cti/icon_tel_1.gif" style="cursor:hand;" onClick="fn_dial_1('2','1');" align="absmiddle" title="일반전화로 전화걸기">&nbsp<img src="/Images/Cti/icon_sche.gif" style="cursor:hand;" onClick="fn_sms('2','1');" align="absmiddle" title="문자전송"></td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>연락처비고</td>
					<td bgcolor="#FFFFFF">&nbsp;<input type="text" name="ETCPHONE_1" value="<%=db_ETCPHONE_1%>" size="20" maxlength="20" style="border-width:1px ; border-color:#cccccc ; border-style:solid">
					</td>						
				</tr>
			    <tr>

					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>설문일시</td>
					<td bgcolor="#FFFFFF" width=200>&nbsp;<input type="text" name="MONITORDATE_1" value="<%=db_MONITORDATE_1%>" size="10" maxlength="10" style="border-width:1px ; border-color:#cccccc ; border-style:solid" >&nbsp;<img src="/Images/icon_sche.gif" title="달력" style="cursor:hand;" align="absmiddle" name="Date2_CAR" onClick="window.open('/Include/Calendar_view.asp?goMonth='+document.all.MONITORDATE_1.value+'&firstcode=&FlightCode=&StartEndGu=S&ControlName=opener.document.all.MONITORDATE_1','','width=800,height=300,top=200,left=300,scrollbars=no,status=no');" onblur="setOutColor(this);">
						&nbsp;<img src="/Images/Btn/BtnIconDel.gif" title="삭제" style="cursor:hand;" align="absmiddle"onclick="document.all.MONITORDATE_1.value='';">
					</td>

					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>설문결과</td>
					<td bgcolor="#FFFFFF" width=200>&nbsp;<%=db_getCodeName("B10",db_MONITOR_RESULT_1)%>
					</td>

					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>예약일시</td>
					<td bgcolor="#FFFFFF" width=200>&nbsp;<%=db_RESERVEDATE_1%>
					</td>
				</tr>
				<tr>
					<td bgcolor="#FFEEF9" class="TDCont" align='center' width=100>통화내역</td>
					<td bgcolor="#FFFFFF" colspan=6>&nbsp;<span id="HISTORY_1"><%=db_History_1%></span>
					</td>	
				</tr>

			</table>
			<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table width="920" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">
			    <tr height="25">
					<td bgcolor="#EEF6FF" width=530 rowspan=2 class="TDCont"  colspan='2' align='center'>질문사항</td>
					<td bgcolor="#EEF6FF" colspan='3' class="TDCont" align='center' width=210>답변결과</td>
					<td bgcolor="#EEF6FF" rowspan=2 class="TDCont" align='center' width=70>가점<br>(+1)</td>
					<td bgcolor="#EEF6FF" rowspan=2 class="TDCont" align='center' width=70>점수</td>
					<td bgcolor="#EEF6FF" rowspan=2 class="TDCont" align='center' width=70>초기화<br><img src="/Images/Btn/BtnIconDel.gif" title="점수초기화" style="cursor:hand;" align="absmiddle"onclick="javascript:fn_ALLDEL('1','<%=i%>');"></td>
				</tr>
			    <tr height="20">
					<td bgcolor="#EEF6FF" class="TDCont" align='center' width=70>만족<br>(9)</td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center' width=70>보통<br>(8)</td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center' width=70>불만족<br>(7)</td>
				</tr>
<%

			SQL = "SELECT * FROM TB_CODE where CODEGROUP = '" & sGubun &"' AND USEYN = 'Y' ORDER BY CODE"
			SET Rs = DB.execute(SQL)			

			i = 0
			do until rs.eof
				i = i + 1
				if ( i mod 2 ) = 1 then
					sBgColor = "#ffffff"
				else
					sBgColor = "#FFFCE7"				
				end if

				'값 불러오기
				SQL1 = "select *"
				SQL1 = SQL1 & " from armyinformix.dbo.monitor where factnum = '" & receiptfactnum & "' and factpeoplenum='"& db_factPeoplenum_1 & "' and  seqno = " & i
				
				SET Rs1 = DB.execute(SQL1)	
				if Rs1.eof = false then
					point9 = Rs1("point9")
					point8 = Rs1("point8")
					point7 = Rs1("point7")
					pointplus = Rs1("pointplus")
					totpoint = Rs1("totpoint")
				else
					point9 = ""
					point8 = ""
					point7 = ""
					pointplus = ""
					totpoint = ""
				end if
%>				
			    <tr>
					<td bgcolor="<%=sBgColor%>" width=530  class="TDCont"  colspan='2'>&nbsp;<%=rs("codename")%></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="radio" name="QUESTION_1<%=i%>" value="9" class="none" onClick="fn_YES('1','<%=i%>','9');" <%if point9 = "1" then%>checked<%end if%> ></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="radio" name="QUESTION_1<%=i%>" value="8" class="none" onClick="fn_YES('1','<%=i%>','8');" <%if point8 = "1" then%>checked<%end if%> ></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="radio" name="QUESTION_1<%=i%>" value="7" class="none" onClick="fn_YES('1','<%=i%>','7');" <%if point7 = "1" then%>checked<%end if%> ></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="checkbox" name="QUESTIONP_1<%=i%>" value="1" class="none" onClick="fn_YES('1','<%=i%>','1');"<% if pointplus="1" then Response.Write("checked") end if %>>					
					</td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="text" name="POINT_1<%=i%>" value="<%=totpoint%>" size="2" maxlength="2" style="border-width:0px ; border-color:#cccccc ; border-style:solid; text-align:right; background-color:<%=sBgColor%>" readonly></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><img src="/Images/Btn/BtnIconDel.gif" title="삭제" style="cursor:hand;" align="absmiddle"onclick="javascript:fn_DEL('1','<%=i%>');"></td>

				</tr>
<%
				rs.movenext
			loop
%>
			    <tr ><td bgcolor="#ffffff" class="TDCont"  colspan='5' valign="top" height="100">
					<table width="100%" height="100%" border="0" cellspacing="1" cellpadding="0" align="center" bgcolor="#CCCCCC">
						<tr height="30">
							<td bgcolor="#FFEEF9" class="TDCont" width=100 align='center'>설문결과</td>
							<td bgcolor="#FFFFFF" width="120">
								<%
									'======= 처리구분 코드 가져오기 ==================================================
									SqlCode = "SELECT CODE, CODENAME FROM TB_CODE"
									SqlCode = SqlCode& " WHERE USEYN='Y' AND SYSYN='Y' AND CODEGROUP='B10'"
									SqlCode = SqlCode& " ORDER BY CODE"
									set RsCode = db.execute(SqlCode)
								%>
								&nbsp;<select name="MONITORRESULT_1" size="1" class="ComboFFFCE7" onChange="fn_ResultSet('1');">
									<option value="">설문결과선택</option>
									<%
										IF NOT(RsCode.Eof OR RsCode.bof) THEN
											DO until RsCode.EOF
												CODE = RsCode("CODE")
												CODENAME = RsCode("CODENAME")
									%>
									<%=printSelect("" &CODENAME& "","" &CODE& "","" &db_MONITOR_RESULT_1& "")%>
									<%
											RsCode.MoveNext
											LOOP
										END IF
										RsCode.Close
										set RsCode = NOTHING
									%>
								</select>
							
							</td>


							<td bgcolor="#FFEEF9" class="TDCont" align='center' width="100">상담예약일시</td>
							<td bgcolor="#FFFFFF">&nbsp;<input value="<%=RESERVEDATE_1%>" name="RESERVEDATE_1" readonly type="text" size="10" onfocus="setFocusColor(this);" onClick="new CalendarFrame.Calendar(this);" onblur="setOutColor(this);" onchange="fn_settime('1')">&nbsp;<img src="/Images/icon_sche.gif" title="달력" style="cursor:hand;" align="absmiddle" name="RESERVE_CAR_1" onClick="window.open('/Include/Calendar_view.asp?goMonth='+document.ListForm.RESERVEDATE_1.value+'&firstcode=&FlightCode=&StartEndGu=S&ControlName=opener.document.ListForm.RESERVEDATE_1','','width=800,height=300,top=200,left=300,scrollbars=no,status=no');" onblur="setOutColor(this);" >&nbsp;<img src="/Images/Btn/BtnIconDel.gif" title="삭제" name="RESERVE_DEL_1" style="cursor:hand;" align="absmiddle"onclick="document.all.RESERVEDATE_1.value='';">&nbsp;&nbsp;<input value="<%=RESERVEHOUR_1%>" name="RESERVEHOUR_1" type="text" size="2" maxlength="2" onfocus="setFocusColor(this);" >시&nbsp;<input value="<%=RESERVEMIN_1%>" name="RESERVEMIN_1" type="text" size="2" maxlength="2" onfocus="setFocusColor(this);">분&nbsp;&nbsp;&nbsp;&nbsp;&nbsp<select name="RESERVETIME_1" size="1" class="ComboFFFCE7" onchange="fn_settime('1')">
									<Option value ='' selected>시간선택</option>
									<Option value ='1' >10분후</option>
									<Option value ='2' >30분후</option>
									<Option value ='3' >1시간후</option>
									<Option value ='4' >2시간후</option>
									<Option value ='08' >오전 7시</option>
									<Option value ='08' >오전 8시</option>
									<Option value ='09' >오전 9시</option>
									<Option value ='10' >오전10시</option>
									<Option value ='11' >오전11시</option>
									<Option value ='12' >오후12시</option>
									<Option value ='13' >오후13시</option>
									<Option value ='14' >오후14시</option>
									<Option value ='15' >오후15시</option>
									<Option value ='16' >오후16시</option>
									<Option value ='17' >오후17시</option>
									<Option value ='18' >오후18시</option>
									<Option value ='19' >오후19시</option>
									<Option value ='20' >오후20시</option>
									<Option value ='21' >오후21시</option>
									<Option value ='22' >오후22시</option>

								</select>
							</td>
						</tr>
						<tr height="40">
							<td bgcolor="#FFEEF9" class="TDCont" align='center' width=100>종합평가</td>
							<td bgcolor="#FFFFFF" colspan=6>&nbsp;<textarea name="Remark_1" style="width:99%; height:100%" wrap="soft" class="TextareaInput"><%=db_Remark_1%></textarea>
							</td>	
						</tr>
						<tr height="40">
							<td bgcolor="#FFEEF9" class="TDCont" align='center' width=100>비고</td>
							<td bgcolor="#FFFFFF" colspan=6>&nbsp;<textarea name="Remark1_1" style="width:99%; height:100%" wrap="soft" class="TextareaInput"><%=db_Remark1_1%></textarea>
							</td>
						</tr>
					</table>
					</td>
					<td bgcolor="#EEF6FF" width=70 class="TDCont" align='center'>점수(평균):</td>
					<td bgcolor="#FFFFFF" width=140 class="TDCont" align='center' colspan=2><input type="text" name="TOT_1" value="<%=db_TOT_1%>" size="5" maxlength="5" style="border-width:0px ; border-color:#EEF6FF ; border-style:solid; text-align:right; font-color:#ff0000;font-size:15px;font-weight:bold" readonly ></td>
				</tr>
			</table>

<%if db_factPeoplenum_2 = "" then%>
			<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table width="920" height="30" border="0" cellspacing="0" cellpadding="0" align="center">
				<tr>
					<td align="left"><img src="/Images/Btn/BtnRegiAdd_P<%=right(sGubun,1)%>.GIF" style="cursor:hand;" align="absmiddle" onClick="fn_AddForm('ON','divFORM2','FRM2');">&nbsp;<img src="/Images/Btn/BtnRegiDel_P<%=right(sGubun,1)%>.GIF" style="cursor:hand;" align="absmiddle" onClick="fn_AddForm('OFF','divFORM1','FRM1');"></td>
					<td align='left' width="50%"><input type="checkbox" name="RecYN_1" value="Y" class="none"><font color='#0000ff'><b>녹취첨부</b></font></td>
				</tr>
			</table>
<%else%>
			<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table width="920" height="30" border="0" cellspacing="0" cellpadding="0" align="center">
				<tr>
					<td align="left"><img src="/Images/Btn/BtnRegiDel_P<%=right(sGubun,1)%>.GIF" style="cursor:hand;" align="absmiddle" onClick="fn_AddForm('OFF','divFORM1','FRM1');"></td>
					<td align='left' width="50%"><input type="checkbox" name="RecYN_1" value="Y" class="none"><font color='#0000ff'><b>녹취첨부</b></font></td>

				</tr>
			</table>
<%end if%>
		</td>
	</tr>
</table>
</span>


<%'====== 상담접수 폼 #2 시작 =======================================================================================%>
<%if db_factPeoplenum_2 = "" then%>
	<span id="divFORM2" style="display:none;">
<% else %>
	<span id="divFORM2" style="display:block;">
<%end if%>
<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
<table border="0" width="920" cellpadding="0" cellspacing="1" bgcolor="#EFECE5" align="center">
	<tr bgcolor="#FFFFFF">
		<td>
			<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table width="920" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#ffffff">
			    <tr>
					<td align="left" bgcolor="#FFFFFF" class="TDCont">&nbsp;<img src="/Images/dot_01.gif" >&nbsp;<b><font color="#ff00ff">(<%=sGubunName%>)</font> 설문지#2</b></td>
				</tr>
			</table>
			<table width="920" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">
			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>관 계</td>
					<td bgcolor="#FFFFFF">						<%
							'======= 처리구분 코드 가져오기 ==================================================
							SqlCode = "SELECT CODE, CODENAME FROM TB_CODE"
							SqlCode = SqlCode& " WHERE USEYN='Y' AND SYSYN='Y' AND CODEGROUP='B01'"
							SqlCode = SqlCode& " ORDER BY CODE"
							set RsCode = db.execute(SqlCode)
						%>
						&nbsp;<select name="SECTION2_2" size="1" class="ComboFFFCE7" onChange="fn_UpdateData('2','SECTION2_');">
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("CODE")
										CODENAME = RsCode("CODENAME")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &sGubun& "")%>
							<%
									RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
							%>
						</select>
					&nbsp;&nbsp;&nbsp;&nbsp;<font color="#0000ff"><%=db_factPeoplenum_2%></font></td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>계급(신분)</td>
					<td bgcolor="#FFFFFF"><%
							'======= 처리구분 코드 가져오기 ==================================================
							SqlCode = "SELECT *	FROM armyinformix.dbo.data"
							SqlCode = SqlCode& " where [restrict] = '104' order by [restrict]"
							set RsCode = db.execute(SqlCode)
						%>
						&nbsp;<select name="LEVEL_2" size="1" class="ComboFFFCE7">
							<option value="">계급선택</option>
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("number")
										CODENAME = RsCode("contents")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &db_level_2& "")%>
							<%
									RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
							%>
						</select></td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>성  명</td>
					<td bgcolor="#FFFFFF">&nbsp;<input type="text" name="NAME_2" value="<%=db_NAME_2%>" size="20" maxlength="20" style="border-width:1px ; border-color:#cccccc ; border-style:solid" ></td>
				</tr>
			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>연락처 1</td>
					<td bgcolor="#FFFFFF">&nbsp;<input type="text" name="HOMEPHONE_2" value="<%=db_HOMEPHONE_2%>" size="13" maxlength="13" style="border-width:1px ; border-color:#cccccc ; border-style:solid">
					<img src="/Images/Cti/icon_tel.gif" style="cursor:hand;" onClick="fn_dial('1','2');" align="absmiddle" title="군전화로 전화걸기">&nbsp<img src="/Images/Cti/icon_tel_1.gif" style="cursor:hand;" onClick="fn_dial_1('1','2');" align="absmiddle" title="일반전화로 전화걸기">&nbsp<img src="/Images/Cti/icon_sche.gif" style="cursor:hand;" onClick="fn_sms('1','2');" align="absmiddle" title="문자전송"></td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>연락처 2</td>
					<td bgcolor="#FFFFFF">&nbsp;<input type="text" name="MOBILEPHONE_2" value="<%=db_MOBILEPHONE_2%>" size="13" maxlength="13" style="border-width:1px ; border-color:#cccccc ; border-style:solid">&nbsp;<img src="/Images/Cti/icon_tel.gif" style="cursor:hand;" onClick="fn_dial('2','2');" align="absmiddle" title="군전화로 전화걸기">&nbsp<img src="/Images/Cti/icon_tel_1.gif" style="cursor:hand;" onClick="fn_dial_1('2','2');" align="absmiddle" title="일반전화로 전화걸기">&nbsp<img src="/Images/Cti/icon_sche.gif" style="cursor:hand;" onClick="fn_sms('2','2');" align="absmiddle" title="문자전송"></td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>연락처비고</td>
					<td bgcolor="#FFFFFF">&nbsp;<input type="text" name="ETCPHONE_2" value="<%=db_ETCPHONE_2%>" size="20" maxlength="20" style="border-width:1px ; border-color:#cccccc ; border-style:solid">
					</td>						
				</tr>
			    <tr>

					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>설문일시</td>
					<td bgcolor="#FFFFFF" width=200>&nbsp;<input type="text" name="MONITORDATE_2" value="<%=db_MONITORDATE_2%>" size="10" maxlength="10" style="border-width:1px ; border-color:#cccccc ; border-style:solid" >&nbsp;<img src="/Images/icon_sche.gif" title="달력" style="cursor:hand;" align="absmiddle" name="Date2_CAR" onClick="window.open('/Include/Calendar_view.asp?goMonth='+document.all.MONITORDATE_2.value+'&firstcode=&FlightCode=&StartEndGu=S&ControlName=opener.document.all.MONITORDATE_2','','width=800,height=300,top=200,left=300,scrollbars=no,status=no');" onblur="setOutColor(this);">
						&nbsp;<img src="/Images/Btn/BtnIconDel.gif" title="삭제" style="cursor:hand;" align="absmiddle"onclick="document.all.MONITORDATE_2.value='';">
					</td>

					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>설문결과</td>
					<td bgcolor="#FFFFFF" width=200>&nbsp;<%=db_getCodeName("B10",db_MONITOR_RESULT_2)%>
					</td>

					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>예약일시</td>
					<td bgcolor="#FFFFFF" width=200>&nbsp;<%=db_RESERVEDATE_2%>
					</td>
				</tr>
				<tr>
					<td bgcolor="#FFEEF9" class="TDCont" align='center' width=100>통화내역</td>
					<td bgcolor="#FFFFFF" colspan=6>&nbsp;<span id="HISTORY_2"><%=db_History_2%></span>
					</td>	
				</tr>

			</table>
			<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table width="920" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">
			    <tr height="25">
					<td bgcolor="#EEF6FF" width=530 rowspan=2 class="TDCont"  colspan='2' align='center'>질문사항</td>
					<td bgcolor="#EEF6FF" colspan='3' class="TDCont" align='center' width=210>답변결과</td>
					<td bgcolor="#EEF6FF" rowspan=2 class="TDCont" align='center' width=70>가점<br>(+1)</td>
					<td bgcolor="#EEF6FF" rowspan=2 class="TDCont" align='center' width=70>점수</td>
					<td bgcolor="#EEF6FF" rowspan=2 class="TDCont" align='center' width=70>초기화<br><img src="/Images/Btn/BtnIconDel.gif" title="점수초기화" style="cursor:hand;" align="absmiddle"onclick="javascript:fn_ALLDEL('2','<%=i%>');"></td>
				</tr>
			    <tr height="20">
					<td bgcolor="#EEF6FF" class="TDCont" align='center' width=70>만족<br>(9)</td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center' width=70>보통<br>(8)</td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center' width=70>불만족<br>(7)</td>
				</tr>
<%

			SQL = "SELECT * FROM TB_CODE where CODEGROUP = '" & sGubun &"' AND USEYN = 'Y' ORDER BY CODE"
			SET Rs = DB.execute(SQL)			

			i = 0
			do until rs.eof
				i = i + 1
				if ( i mod 2 ) = 1 then
					sBgColor = "#ffffff"
				else
					sBgColor = "#FFFCE7"				
				end if

				'값 불러오기
				SQL1 = "select *"
				SQL1 = SQL1 & " from armyinformix.dbo.monitor where factnum = '" & receiptfactnum & "' and factpeoplenum='"& db_factPeoplenum_2 & "' and  seqno = " & i
				
				SET Rs1 = DB.execute(SQL1)	
				if Rs1.eof = false then
					point9 = Rs1("point9")
					point8 = Rs1("point8")
					point7 = Rs1("point7")
					pointplus = Rs1("pointplus")
					totpoint = Rs1("totpoint")
				else
					point9 = ""
					point8 = ""
					point7 = ""
					pointplus = ""
					totpoint = ""
				end if
%>				
			    <tr>
					<td bgcolor="<%=sBgColor%>" width=530  class="TDCont"  colspan='2'>&nbsp;<%=rs("codename")%></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="radio" name="QUESTION_2<%=i%>" value="9" class="none" onClick="fn_YES('2','<%=i%>','9');" <%if point9 = "1" then%>checked<%end if%> ></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="radio" name="QUESTION_2<%=i%>" value="8" class="none" onClick="fn_YES('2','<%=i%>','8');" <%if point8 = "1" then%>checked<%end if%> ></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="radio" name="QUESTION_2<%=i%>" value="7" class="none" onClick="fn_YES('2','<%=i%>','7');" <%if point7 = "1" then%>checked<%end if%> ></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="checkbox" name="QUESTIONP_2<%=i%>" value="1" class="none" onClick="fn_YES('2','<%=i%>','1');"<% if pointplus="1" then Response.Write("checked") end if %>>					
					</td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="text" name="POINT_2<%=i%>" value="<%=totpoint%>" size="2" maxlength="2" style="border-width:0px ; border-color:#cccccc ; border-style:solid; text-align:right; background-color:<%=sBgColor%>" readonly></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><img src="/Images/Btn/BtnIconDel.gif" title="삭제" style="cursor:hand;" align="absmiddle"onclick="javascript:fn_DEL('2','<%=i%>');"></td>

				</tr>
<%
				rs.movenext
			loop
%>
			    <tr ><td bgcolor="#ffffff" class="TDCont"  colspan='5' valign="top" height="100">
					<table width="100%" height="100%" border="0" cellspacing="1" cellpadding="0" align="center" bgcolor="#CCCCCC">
						<tr height="30">
							<td bgcolor="#FFEEF9" class="TDCont" width=100 align='center'>설문결과</td>
							<td bgcolor="#FFFFFF" width="120">
								<%
									'======= 처리구분 코드 가져오기 ==================================================
									SqlCode = "SELECT CODE, CODENAME FROM TB_CODE"
									SqlCode = SqlCode& " WHERE USEYN='Y' AND SYSYN='Y' AND CODEGROUP='B10'"
									SqlCode = SqlCode& " ORDER BY CODE"
									set RsCode = db.execute(SqlCode)
								%>
								&nbsp;<select name="MONITORRESULT_2" size="1" class="ComboFFFCE7" onChange="fn_ResultSet('2');">
									<option value="">설문결과선택</option>
									<%
										IF NOT(RsCode.Eof OR RsCode.bof) THEN
											DO until RsCode.EOF
												CODE = RsCode("CODE")
												CODENAME = RsCode("CODENAME")
									%>
									<%=printSelect("" &CODENAME& "","" &CODE& "","" &db_MONITOR_RESULT_2& "")%>
									<%
											RsCode.MoveNext
											LOOP
										END IF
										RsCode.Close
										set RsCode = NOTHING
									%>
								</select>
							
							</td>


							<td bgcolor="#FFEEF9" class="TDCont" align='center' width="100">상담예약일시</td>
							<td bgcolor="#FFFFFF">&nbsp;<input value="<%=RESERVEDATE_2%>" name="RESERVEDATE_2" readonly type="text" size="10" onfocus="setFocusColor(this);" onClick="new CalendarFrame.Calendar(this);" onblur="setOutColor(this);" onchange="fn_settime('2')">&nbsp;<img src="/Images/icon_sche.gif" title="달력" style="cursor:hand;" align="absmiddle" name="RESERVE_CAR_2" onClick="window.open('/Include/Calendar_view.asp?goMonth='+document.ListForm.RESERVEDATE_2.value+'&firstcode=&FlightCode=&StartEndGu=S&ControlName=opener.document.ListForm.RESERVEDATE_2','','width=800,height=300,top=200,left=300,scrollbars=no,status=no');" onblur="setOutColor(this);" >&nbsp;<img src="/Images/Btn/BtnIconDel.gif" title="삭제" name="RESERVE_DEL_2" style="cursor:hand;" align="absmiddle"onclick="document.all.RESERVEDATE_2.value='';">&nbsp;&nbsp;<input value="<%=RESERVEHOUR_2%>" name="RESERVEHOUR_2" type="text" size="2" maxlength="2" onfocus="setFocusColor(this);" >시&nbsp;<input value="<%=RESERVEMIN_2%>" name="RESERVEMIN_2" type="text" size="2" maxlength="2" onfocus="setFocusColor(this);">분&nbsp;&nbsp;&nbsp;&nbsp;&nbsp<select name="RESERVETIME_2" size="1" class="ComboFFFCE7" onchange="fn_settime('2')">
									<Option value ='' selected>시간선택</option>
									<Option value ='1' >10분후</option>
									<Option value ='2' >30분후</option>
									<Option value ='3' >1시간후</option>
									<Option value ='4' >2시간후</option>
									<Option value ='08' >오전 7시</option>
									<Option value ='08' >오전 8시</option>
									<Option value ='09' >오전 9시</option>
									<Option value ='10' >오전10시</option>
									<Option value ='11' >오전11시</option>
									<Option value ='12' >오후12시</option>
									<Option value ='13' >오후13시</option>
									<Option value ='14' >오후14시</option>
									<Option value ='15' >오후15시</option>
									<Option value ='16' >오후16시</option>
									<Option value ='17' >오후17시</option>
									<Option value ='18' >오후18시</option>
									<Option value ='19' >오후19시</option>
									<Option value ='20' >오후20시</option>
									<Option value ='21' >오후21시</option>
									<Option value ='22' >오후22시</option>

								</select>
							</td>
						</tr>
						<tr height="40">
							<td bgcolor="#FFEEF9" class="TDCont" align='center' width=100>종합평가</td>
							<td bgcolor="#FFFFFF" colspan=6>&nbsp;<textarea name="Remark_2" style="width:99%; height:100%" wrap="soft" class="TextareaInput"><%=db_Remark_2%></textarea>
							</td>	
						</tr>
						<tr height="40">
							<td bgcolor="#FFEEF9" class="TDCont" align='center' width=100>비고</td>
							<td bgcolor="#FFFFFF" colspan=6>&nbsp;<textarea name="Remark1_2" style="width:99%; height:100%" wrap="soft" class="TextareaInput"><%=db_Remark1_2%></textarea>
							</td>	
						</tr>
					</table>
					</td>
					<td bgcolor="#EEF6FF" width=70 class="TDCont" align='center'>점수(평균):</td>
					<td bgcolor="#FFFFFF" width=140 class="TDCont" align='center' colspan=2><input type="text" name="TOT_2" value="<%=db_TOT_2%>" size="5" maxlength="5" style="border-width:0px ; border-color:#EEF6FF ; border-style:solid; text-align:right; font-color:#ff0000;font-size:15px;font-weight:bold" readonly ></td>
				</tr>
			</table>

<%if db_factPeoplenum_3 = "" then%>
			<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table width="920" height="30" border="0" cellspacing="0" cellpadding="0" align="center">
				<tr>
					<td align="left"><img src="/Images/Btn/BtnRegiAdd_P<%=right(sGubun,1)%>.GIF" style="cursor:hand;" align="absmiddle" onClick="fn_AddForm('ON','divFORM3','FRM3');">&nbsp;<img src="/Images/Btn/BtnRegiDel_P<%=right(sGubun,1)%>.GIF" style="cursor:hand;" align="absmiddle" onClick="fn_AddForm('OFF','divFORM2','FRM2');"></td>
					<td align='left' width="50%"><input type="checkbox" name="RecYN_2" value="Y" class="none"><font color='#0000ff'><b>녹취첨부</b></font></td>
				</tr>
			</table>
<%else%>
			<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table width="920" height="30" border="0" cellspacing="0" cellpadding="0" align="center">
				<tr>
					<td align="left"><img src="/Images/Btn/BtnRegiDel_P<%=right(sGubun,1)%>.GIF" style="cursor:hand;" align="absmiddle" onClick="fn_AddForm('OFF','divFORM2','FRM2');"></td>
					<td align='left' width="50%"><input type="checkbox" name="RecYN_2" value="Y" class="none"><font color='#0000ff'><b>녹취첨부</b></font></td>
				</tr>
			</table>
<%end if%>
		</td>
	</tr>
</table>
</span>


<%'====== 상담접수 폼 #3 시작 =======================================================================================%>
<%if db_factPeoplenum_3 = "" then%>
	<span id="divFORM3" style="display:none;">
<% else %>
	<span id="divFORM3" style="display:block;">
<%end if%>
<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
<table border="0" width="920" cellpadding="0" cellspacing="1" bgcolor="#EFECE5" align="center">
	<tr bgcolor="#FFFFFF">
		<td>
			<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table width="920" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#ffffff">
			    <tr>
					<td align="left" bgcolor="#FFFFFF" class="TDCont">&nbsp;<img src="/Images/dot_01.gif" >&nbsp;<b><font color="#ff00ff">(<%=sGubunName%>)</font> 설문지#3</b></td>
				</tr>
			</table>
			<table width="920" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">
			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>관 계</td>
					<td bgcolor="#FFFFFF">						<%
							'======= 처리구분 코드 가져오기 ==================================================
							SqlCode = "SELECT CODE, CODENAME FROM TB_CODE"
							SqlCode = SqlCode& " WHERE USEYN='Y' AND SYSYN='Y' AND CODEGROUP='B01'"
							SqlCode = SqlCode& " ORDER BY CODE"
							set RsCode = db.execute(SqlCode)
						%>
						&nbsp;<select name="SECTION2_3" size="1" class="ComboFFFCE7" onChange="fn_UpdateData('3','SECTION2_');">
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("CODE")
										CODENAME = RsCode("CODENAME")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &sGubun& "")%>
							<%
									RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
							%>
						</select>
					&nbsp;&nbsp;&nbsp;&nbsp;<font color="#0000ff"><%=db_factPeoplenum_3%></font></td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>계급(신분)</td>
					<td bgcolor="#FFFFFF"><%
							'======= 처리구분 코드 가져오기 ==================================================
							SqlCode = "SELECT *	FROM armyinformix.dbo.data"
							SqlCode = SqlCode& " where [restrict] = '104' order by [restrict]"
							set RsCode = db.execute(SqlCode)
						%>
						&nbsp;<select name="LEVEL_3" size="1" class="ComboFFFCE7">
							<option value="">계급선택</option>
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("number")
										CODENAME = RsCode("contents")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &db_level_3& "")%>
							<%
									RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
							%>
						</select></td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>성  명</td>
					<td bgcolor="#FFFFFF">&nbsp;<input type="text" name="NAME_3" value="<%=db_NAME_3%>" size="20" maxlength="20" style="border-width:1px ; border-color:#cccccc ; border-style:solid" ></td>
				</tr>
			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>연락처 1</td>
					<td bgcolor="#FFFFFF">&nbsp;<input type="text" name="HOMEPHONE_3" value="<%=db_HOMEPHONE_3%>" size="13" maxlength="13" style="border-width:1px ; border-color:#cccccc ; border-style:solid">
					<img src="/Images/Cti/icon_tel.gif" style="cursor:hand;" onClick="fn_dial('1','3');" align="absmiddle" title="군전화로 전화걸기">&nbsp<img src="/Images/Cti/icon_tel_1.gif" style="cursor:hand;" onClick="fn_dial_1('1','3');" align="absmiddle" title="일반전화로 전화걸기">&nbsp<img src="/Images/Cti/icon_sche.gif" style="cursor:hand;" onClick="fn_sms('1','3');" align="absmiddle" title="문자전송"></td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>연락처 2</td>
					<td bgcolor="#FFFFFF">&nbsp;<input type="text" name="MOBILEPHONE_3" value="<%=db_MOBILEPHONE_3%>" size="13" maxlength="13" style="border-width:1px ; border-color:#cccccc ; border-style:solid">&nbsp;<img src="/Images/Cti/icon_tel.gif" style="cursor:hand;" onClick="fn_dial('2','3');" align="absmiddle" title="군전화로 전화걸기">&nbsp<img src="/Images/Cti/icon_tel_1.gif" style="cursor:hand;" onClick="fn_dial_1('2','3');" align="absmiddle" title="일반전화로 전화걸기">&nbsp<img src="/Images/Cti/icon_sche.gif" style="cursor:hand;" onClick="fn_sms('2','3');" align="absmiddle" title="문자전송"></td>

					<td bgcolor="#EEF6FF" class="TDCont" align='center'>연락처비고</td>
					<td bgcolor="#FFFFFF">&nbsp;<input type="text" name="ETCPHONE_3" value="<%=db_ETCPHONE_3%>" size="20" maxlength="20" style="border-width:1px ; border-color:#cccccc ; border-style:solid">
					</td>
						
				</tr>
			    <tr>

					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>설문일시</td>
					<td bgcolor="#FFFFFF" width=200>&nbsp;<input type="text" name="MONITORDATE_3" value="<%=db_MONITORDATE_3%>" size="10" maxlength="10" style="border-width:1px ; border-color:#cccccc ; border-style:solid" >&nbsp;<img src="/Images/icon_sche.gif" title="달력" style="cursor:hand;" align="absmiddle" name="Date2_CAR" onClick="window.open('/Include/Calendar_view.asp?goMonth='+document.all.MONITORDATE_3.value+'&firstcode=&FlightCode=&StartEndGu=S&ControlName=opener.document.all.MONITORDATE_3','','width=800,height=300,top=200,left=300,scrollbars=no,status=no');" onblur="setOutColor(this);">
						&nbsp;<img src="/Images/Btn/BtnIconDel.gif" title="삭제" style="cursor:hand;" align="absmiddle"onclick="document.all.MONITORDATE_3.value='';">
					</td>

					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>설문결과</td>
					<td bgcolor="#FFFFFF" width=200>&nbsp;<%=db_getCodeName("B10",db_MONITOR_RESULT_3)%>
					</td>

					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>예약일시</td>
					<td bgcolor="#FFFFFF" width=200>&nbsp;<%=db_RESERVEDATE_3%>
					</td>
				</tr>
				<tr>
					<td bgcolor="#FFEEF9" class="TDCont" align='center' width=100>통화내역</td>
					<td bgcolor="#FFFFFF" colspan=6>&nbsp;<span id="HISTORY_3"><%=db_History_3%></span>
					</td>	
				</tr>

			</table>
			<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table width="920" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">
			    <tr height="25">
					<td bgcolor="#EEF6FF" width=530 rowspan=2 class="TDCont"  colspan='2' align='center'>질문사항</td>
					<td bgcolor="#EEF6FF" colspan='3' class="TDCont" align='center' width=210>답변결과</td>
					<td bgcolor="#EEF6FF" rowspan=2 class="TDCont" align='center' width=70>가점<br>(+1)</td>
					<td bgcolor="#EEF6FF" rowspan=2 class="TDCont" align='center' width=70>점수</td>
					<td bgcolor="#EEF6FF" rowspan=2 class="TDCont" align='center' width=70>초기화<br><img src="/Images/Btn/BtnIconDel.gif" title="점수초기화" style="cursor:hand;" align="absmiddle"onclick="javascript:fn_ALLDEL('3','<%=i%>');"></td>
				</tr>
			    <tr height="20">
					<td bgcolor="#EEF6FF" class="TDCont" align='center' width=70>만족<br>(9)</td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center' width=70>보통<br>(8)</td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center' width=70>불만족<br>(7)</td>
				</tr>
<%

			SQL = "SELECT * FROM TB_CODE where CODEGROUP = '" & sGubun &"' AND USEYN = 'Y' ORDER BY CODE"
			SET Rs = DB.execute(SQL)			

			i = 0
			do until rs.eof
				i = i + 1
				if ( i mod 2 ) = 1 then
					sBgColor = "#ffffff"
				else
					sBgColor = "#FFFCE7"				
				end if

				'값 불러오기
				SQL1 = "select *"
				SQL1 = SQL1 & " from armyinformix.dbo.monitor where factnum = '" & receiptfactnum & "' and factpeoplenum='"& db_factPeoplenum_3 & "' and  seqno = " & i
				
				SET Rs1 = DB.execute(SQL1)	
				if Rs1.eof = false then
					point9 = Rs1("point9")
					point8 = Rs1("point8")
					point7 = Rs1("point7")
					pointplus = Rs1("pointplus")
					totpoint = Rs1("totpoint")
				else
					point9 = ""
					point8 = ""
					point7 = ""
					pointplus = ""
					totpoint = ""
				end if
%>				
			    <tr>
					<td bgcolor="<%=sBgColor%>" width=530  class="TDCont"  colspan='2'>&nbsp;<%=rs("codename")%></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="radio" name="QUESTION_3<%=i%>" value="9" class="none" onClick="fn_YES('3','<%=i%>','9');" <%if point9 = "1" then%>checked<%end if%> ></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="radio" name="QUESTION_3<%=i%>" value="8" class="none" onClick="fn_YES('3','<%=i%>','8');" <%if point8 = "1" then%>checked<%end if%> ></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="radio" name="QUESTION_3<%=i%>" value="7" class="none" onClick="fn_YES('3','<%=i%>','7');" <%if point7 = "1" then%>checked<%end if%> ></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="checkbox" name="QUESTIONP_3<%=i%>" value="1" class="none" onClick="fn_YES('3','<%=i%>','1');"<% if pointplus="1" then Response.Write("checked") end if %>>					
					</td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="text" name="POINT_3<%=i%>" value="<%=totpoint%>" size="2" maxlength="2" style="border-width:0px ; border-color:#cccccc ; border-style:solid; text-align:right; background-color:<%=sBgColor%>" readonly></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><img src="/Images/Btn/BtnIconDel.gif" title="삭제" style="cursor:hand;" align="absmiddle"onclick="javascript:fn_DEL('3','<%=i%>');"></td>

				</tr>
<%
				rs.movenext
			loop
%>
			    <tr ><td bgcolor="#ffffff" class="TDCont"  colspan='5' valign="top" height="100">
					<table width="100%" height="100%" border="0" cellspacing="1" cellpadding="0" align="center" bgcolor="#CCCCCC">
						<tr height="30">
							<td bgcolor="#FFEEF9" class="TDCont" width=100 align='center'>설문결과</td>
							<td bgcolor="#FFFFFF" width="120">
								<%
									'======= 처리구분 코드 가져오기 ==================================================
									SqlCode = "SELECT CODE, CODENAME FROM TB_CODE"
									SqlCode = SqlCode& " WHERE USEYN='Y' AND SYSYN='Y' AND CODEGROUP='B10'"
									SqlCode = SqlCode& " ORDER BY CODE"
									set RsCode = db.execute(SqlCode)
								%>
								&nbsp;<select name="MONITORRESULT_3" size="1" class="ComboFFFCE7" onChange="fn_ResultSet('3');">
									<option value="">설문결과선택</option>
									<%
										IF NOT(RsCode.Eof OR RsCode.bof) THEN
											DO until RsCode.EOF
												CODE = RsCode("CODE")
												CODENAME = RsCode("CODENAME")
									%>
									<%=printSelect("" &CODENAME& "","" &CODE& "","" &db_MONITOR_RESULT_3& "")%>
									<%
											RsCode.MoveNext
											LOOP
										END IF
										RsCode.Close
										set RsCode = NOTHING
									%>
								</select>
							
							</td>


							<td bgcolor="#FFEEF9" class="TDCont" align='center' width="100">상담예약일시</td>
							<td bgcolor="#FFFFFF">&nbsp;<input value="<%=RESERVEDATE_3%>" name="RESERVEDATE_3" readonly type="text" size="10" onfocus="setFocusColor(this);" onClick="new CalendarFrame.Calendar(this);" onblur="setOutColor(this);" onchange="fn_settime('3')">&nbsp;<img src="/Images/icon_sche.gif" title="달력" style="cursor:hand;" align="absmiddle" name="RESERVE_CAR_3" onClick="window.open('/Include/Calendar_view.asp?goMonth='+document.ListForm.RESERVEDATE_3.value+'&firstcode=&FlightCode=&StartEndGu=S&ControlName=opener.document.ListForm.RESERVEDATE_3','','width=800,height=300,top=200,left=300,scrollbars=no,status=no');" onblur="setOutColor(this);" >&nbsp;<img src="/Images/Btn/BtnIconDel.gif" title="삭제" name="RESERVE_DEL_3" style="cursor:hand;" align="absmiddle"onclick="document.all.RESERVEDATE_3.value='';">&nbsp;&nbsp;<input value="<%=RESERVEHOUR_3%>" name="RESERVEHOUR_3" type="text" size="2" maxlength="2" onfocus="setFocusColor(this);" >시&nbsp;<input value="<%=RESERVEMIN_3%>" name="RESERVEMIN_3" type="text" size="2" maxlength="2" onfocus="setFocusColor(this);">분&nbsp;&nbsp;&nbsp;&nbsp;&nbsp<select name="RESERVETIME_3" size="1" class="ComboFFFCE7" onchange="fn_settime('3')">
									<Option value ='' selected>시간선택</option>
									<Option value ='1' >10분후</option>
									<Option value ='2' >30분후</option>
									<Option value ='3' >1시간후</option>
									<Option value ='4' >2시간후</option>
									<Option value ='08' >오전 7시</option>
									<Option value ='08' >오전 8시</option>
									<Option value ='09' >오전 9시</option>
									<Option value ='10' >오전10시</option>
									<Option value ='11' >오전11시</option>
									<Option value ='12' >오후12시</option>
									<Option value ='13' >오후13시</option>
									<Option value ='14' >오후14시</option>
									<Option value ='15' >오후15시</option>
									<Option value ='16' >오후16시</option>
									<Option value ='17' >오후17시</option>
									<Option value ='18' >오후18시</option>
									<Option value ='19' >오후19시</option>
									<Option value ='20' >오후20시</option>
									<Option value ='21' >오후21시</option>
									<Option value ='22' >오후22시</option>

								</select>
							</td>
						</tr>
						<tr height="40">
							<td bgcolor="#FFEEF9" class="TDCont" align='center' width=100>종합평가</td>
							<td bgcolor="#FFFFFF" colspan=6>&nbsp;<textarea name="Remark_3" style="width:99%; height:100%" wrap="soft" class="TextareaInput"><%=db_Remark_3%></textarea>
							</td>	
						</tr>
						<tr height="40">
							<td bgcolor="#FFEEF9" class="TDCont" align='center' width=100>비고</td>
							<td bgcolor="#FFFFFF" colspan=6>&nbsp;<textarea name="Remark1_3" style="width:99%; height:100%" wrap="soft" class="TextareaInput"><%=db_Remark1_3%></textarea>
							</td>	
						</tr>
					</table>
					</td>
					<td bgcolor="#EEF6FF" width=70 class="TDCont" align='center'>점수(평균):</td>
					<td bgcolor="#FFFFFF" width=140 class="TDCont" align='center' colspan=2><input type="text" name="TOT_3" value="<%=db_TOT_3%>" size="5" maxlength="5" style="border-width:0px ; border-color:#EEF6FF ; border-style:solid; text-align:right; font-color:#ff0000;font-size:15px;font-weight:bold" readonly ></td>
				</tr>
			</table>

<%if db_factPeoplenum_4 = "" then%>
			<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table width="920" height="30" border="0" cellspacing="0" cellpadding="0" align="center">
				<tr>
					<td align="left"><img src="/Images/Btn/BtnRegiAdd_P<%=right(sGubun,1)%>.GIF" style="cursor:hand;" align="absmiddle" onClick="fn_AddForm('ON','divFORM4','FRM4');">&nbsp;<img src="/Images/Btn/BtnRegiDel_P<%=right(sGubun,1)%>.GIF" style="cursor:hand;" align="absmiddle" onClick="fn_AddForm('OFF','divFORM3','FRM3');"></td>
					<td align='left' width="50%"><input type="checkbox" name="RecYN_3" value="Y" class="none"><font color='#0000ff'><b>녹취첨부</b></font></td>
				</tr>
			</table>
<%else%>
			<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table width="920" height="30" border="0" cellspacing="0" cellpadding="0" align="center">
				<tr>
					<td align="left"><img src="/Images/Btn/BtnRegiDel_P<%=right(sGubun,1)%>.GIF" style="cursor:hand;" align="absmiddle" onClick="fn_AddForm('OFF','divFORM3','FRM3');"></td>
					<td align='left' width="50%"><input type="checkbox" name="RecYN_3" value="Y" class="none"><font color='#0000ff'><b>녹취첨부</b></font></td>
				</tr>
			</table>
<%end if%>
		</td>
	</tr>
</table>
</span>


<%'====== 상담접수 폼 #4 시작 =======================================================================================%>
<%if db_factPeoplenum_4 = "" then%>
	<span id="divFORM4" style="display:none;">
<% else %>
	<span id="divFORM4" style="display:block;">
<%end if%>
<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
<table border="0" width="920" cellpadding="0" cellspacing="1" bgcolor="#EFECE5" align="center">
	<tr bgcolor="#FFFFFF">
		<td>
			<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table width="920" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#ffffff">
			    <tr>
					<td align="left" bgcolor="#FFFFFF" class="TDCont">&nbsp;<img src="/Images/dot_01.gif" >&nbsp;<b><font color="#ff00ff">(<%=sGubunName%>)</font> 설문지#4</b></td>
				</tr>
			</table>
			<table width="920" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">
			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>관 계</td>
					<td bgcolor="#FFFFFF">						<%
							'======= 처리구분 코드 가져오기 ==================================================
							SqlCode = "SELECT CODE, CODENAME FROM TB_CODE"
							SqlCode = SqlCode& " WHERE USEYN='Y' AND SYSYN='Y' AND CODEGROUP='B01'"
							SqlCode = SqlCode& " ORDER BY CODE"
							set RsCode = db.execute(SqlCode)
						%>
						&nbsp;<select name="SECTION2_4" size="1" class="ComboFFFCE7" onChange="fn_UpdateData('4','SECTION2_');">
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("CODE")
										CODENAME = RsCode("CODENAME")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &sGubun& "")%>
							<%
									RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
							%>
						</select>
					&nbsp;&nbsp;&nbsp;&nbsp;<font color="#0000ff"><%=db_factPeoplenum_4%></font></td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>계급(신분)</td>
					<td bgcolor="#FFFFFF"><%
							'======= 처리구분 코드 가져오기 ==================================================
							SqlCode = "SELECT *	FROM armyinformix.dbo.data"
							SqlCode = SqlCode& " where [restrict] = '104' order by [restrict]"
							set RsCode = db.execute(SqlCode)
						%>
						&nbsp;<select name="LEVEL_4" size="1" class="ComboFFFCE7">
							<option value="">계급선택</option>
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("number")
										CODENAME = RsCode("contents")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &db_level_4& "")%>
							<%
									RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
							%>
						</select></td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>성  명</td>
					<td bgcolor="#FFFFFF">&nbsp;<input type="text" name="NAME_4" value="<%=db_NAME_4%>" size="20" maxlength="20" style="border-width:1px ; border-color:#cccccc ; border-style:solid" ></td>
				</tr>
			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>연락처 1</td>
					<td bgcolor="#FFFFFF">&nbsp;<input type="text" name="HOMEPHONE_4" value="<%=db_HOMEPHONE_4%>" size="13" maxlength="13" style="border-width:1px ; border-color:#cccccc ; border-style:solid">
					<img src="/Images/Cti/icon_tel.gif" style="cursor:hand;" onClick="fn_dial('1','4');" align="absmiddle" title="군전화로 전화걸기">&nbsp<img src="/Images/Cti/icon_tel_1.gif" style="cursor:hand;" onClick="fn_dial_1('1','4');" align="absmiddle" title="일반전화로 전화걸기">&nbsp<img src="/Images/Cti/icon_sche.gif" style="cursor:hand;" onClick="fn_sms('1','4');" align="absmiddle" title="문자전송"></td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>연락처 2</td>
					<td bgcolor="#FFFFFF">&nbsp;<input type="text" name="MOBILEPHONE_4" value="<%=db_MOBILEPHONE_4%>" size="13" maxlength="13" style="border-width:1px ; border-color:#cccccc ; border-style:solid">&nbsp;<img src="/Images/Cti/icon_tel.gif" style="cursor:hand;" onClick="fn_dial('2','4');" align="absmiddle" title="군전화로 전화걸기">&nbsp<img src="/Images/Cti/icon_tel_1.gif" style="cursor:hand;" onClick="fn_dial_1('2','4');" align="absmiddle" title="일반전화로 전화걸기">&nbsp<img src="/Images/Cti/icon_sche.gif" style="cursor:hand;" onClick="fn_sms('2','4');" align="absmiddle" title="문자전송"></td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>연락처비고</td>
					<td bgcolor="#FFFFFF">&nbsp;<input type="text" name="ETCPHONE_4" value="<%=db_ETCPHONE_4%>" size="20" maxlength="20" style="border-width:1px ; border-color:#cccccc ; border-style:solid">
					</td>
						
				</tr>
			    <tr>

					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>설문일시</td>
					<td bgcolor="#FFFFFF" width=200>&nbsp;<input type="text" name="MONITORDATE_4" value="<%=db_MONITORDATE_4%>" size="10" maxlength="10" style="border-width:1px ; border-color:#cccccc ; border-style:solid" >&nbsp;<img src="/Images/icon_sche.gif" title="달력" style="cursor:hand;" align="absmiddle" name="Date2_CAR" onClick="window.open('/Include/Calendar_view.asp?goMonth='+document.all.MONITORDATE_4.value+'&firstcode=&FlightCode=&StartEndGu=S&ControlName=opener.document.all.MONITORDATE_4','','width=800,height=300,top=200,left=300,scrollbars=no,status=no');" onblur="setOutColor(this);">
						&nbsp;<img src="/Images/Btn/BtnIconDel.gif" title="삭제" style="cursor:hand;" align="absmiddle"onclick="document.all.MONITORDATE_4.value='';">
					</td>

					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>설문결과</td>
					<td bgcolor="#FFFFFF" width=200>&nbsp;<%=db_getCodeName("B10",db_MONITOR_RESULT_4)%>
					</td>

					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>예약일시</td>
					<td bgcolor="#FFFFFF" width=200>&nbsp;<%=db_RESERVEDATE_4%>
					</td>
				</tr>
				<tr>
					<td bgcolor="#FFEEF9" class="TDCont" align='center' width=100>통화내역</td>
					<td bgcolor="#FFFFFF" colspan=6>&nbsp;<span id="HISTORY_4"><%=db_History_4%></span>
					</td>	
				</tr>

			</table>
			<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table width="920" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">
			    <tr height="25">
					<td bgcolor="#EEF6FF" width=530 rowspan=2 class="TDCont"  colspan='2' align='center'>질문사항</td>
					<td bgcolor="#EEF6FF" colspan='3' class="TDCont" align='center' width=210>답변결과</td>
					<td bgcolor="#EEF6FF" rowspan=2 class="TDCont" align='center' width=70>가점<br>(+1)</td>
					<td bgcolor="#EEF6FF" rowspan=2 class="TDCont" align='center' width=70>점수</td>
					<td bgcolor="#EEF6FF" rowspan=2 class="TDCont" align='center' width=70>초기화<br><img src="/Images/Btn/BtnIconDel.gif" title="점수초기화" style="cursor:hand;" align="absmiddle"onclick="javascript:fn_ALLDEL('4','<%=i%>');"></td>
				</tr>
			    <tr height="20">
					<td bgcolor="#EEF6FF" class="TDCont" align='center' width=70>만족<br>(9)</td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center' width=70>보통<br>(8)</td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center' width=70>불만족<br>(7)</td>
				</tr>
<%

			SQL = "SELECT * FROM TB_CODE where CODEGROUP = '" & sGubun &"' AND USEYN = 'Y' ORDER BY CODE"
			SET Rs = DB.execute(SQL)			

			i = 0
			do until rs.eof
				i = i + 1
				if ( i mod 2 ) = 1 then
					sBgColor = "#ffffff"
				else
					sBgColor = "#FFFCE7"				
				end if

				'값 불러오기
				SQL1 = "select *"
				SQL1 = SQL1 & " from armyinformix.dbo.monitor where factnum = '" & receiptfactnum & "' and factpeoplenum='"& db_factPeoplenum_4 & "' and  seqno = " & i
				
				SET Rs1 = DB.execute(SQL1)	
				if Rs1.eof = false then
					point9 = Rs1("point9")
					point8 = Rs1("point8")
					point7 = Rs1("point7")
					pointplus = Rs1("pointplus")
					totpoint = Rs1("totpoint")
				else
					point9 = ""
					point8 = ""
					point7 = ""
					pointplus = ""
					totpoint = ""
				end if
%>				
			    <tr>
					<td bgcolor="<%=sBgColor%>" width=530  class="TDCont"  colspan='2'>&nbsp;<%=rs("codename")%></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="radio" name="QUESTION_4<%=i%>" value="9" class="none" onClick="fn_YES('4','<%=i%>','9');" <%if point9 = "1" then%>checked<%end if%> ></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="radio" name="QUESTION_4<%=i%>" value="8" class="none" onClick="fn_YES('4','<%=i%>','8');" <%if point8 = "1" then%>checked<%end if%> ></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="radio" name="QUESTION_4<%=i%>" value="7" class="none" onClick="fn_YES('4','<%=i%>','7');" <%if point7 = "1" then%>checked<%end if%> ></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="checkbox" name="QUESTIONP_4<%=i%>" value="1" class="none" onClick="fn_YES('4','<%=i%>','1');"<% if pointplus="1" then Response.Write("checked") end if %>>					
					</td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="text" name="POINT_4<%=i%>" value="<%=totpoint%>" size="2" maxlength="2" style="border-width:0px ; border-color:#cccccc ; border-style:solid; text-align:right; background-color:<%=sBgColor%>" readonly></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><img src="/Images/Btn/BtnIconDel.gif" title="삭제" style="cursor:hand;" align="absmiddle"onclick="javascript:fn_DEL('4','<%=i%>');"></td>

				</tr>
<%
				rs.movenext
			loop
%>
			    <tr ><td bgcolor="#ffffff" class="TDCont"  colspan='5' valign="top" height="100">
					<table width="100%" height="100%" border="0" cellspacing="1" cellpadding="0" align="center" bgcolor="#CCCCCC">
						<tr height="30">
							<td bgcolor="#FFEEF9" class="TDCont" width=100 align='center'>설문결과</td>
							<td bgcolor="#FFFFFF" width="120">
								<%
									'======= 처리구분 코드 가져오기 ==================================================
									SqlCode = "SELECT CODE, CODENAME FROM TB_CODE"
									SqlCode = SqlCode& " WHERE USEYN='Y' AND SYSYN='Y' AND CODEGROUP='B10'"
									SqlCode = SqlCode& " ORDER BY CODE"
									set RsCode = db.execute(SqlCode)
								%>
								&nbsp;<select name="MONITORRESULT_4" size="1" class="ComboFFFCE7" onChange="fn_ResultSet('4');">
									<option value="">설문결과선택</option>
									<%
										IF NOT(RsCode.Eof OR RsCode.bof) THEN
											DO until RsCode.EOF
												CODE = RsCode("CODE")
												CODENAME = RsCode("CODENAME")
									%>
									<%=printSelect("" &CODENAME& "","" &CODE& "","" &db_MONITOR_RESULT_4& "")%>
									<%
											RsCode.MoveNext
											LOOP
										END IF
										RsCode.Close
										set RsCode = NOTHING
									%>
								</select>
							
							</td>


							<td bgcolor="#FFEEF9" class="TDCont" align='center' width="100">상담예약일시</td>
							<td bgcolor="#FFFFFF">&nbsp;<input value="<%=RESERVEDATE_4%>" name="RESERVEDATE_4" readonly type="text" size="10" onfocus="setFocusColor(this);" onClick="new CalendarFrame.Calendar(this);" onblur="setOutColor(this);" onchange="fn_settime('4')">&nbsp;<img src="/Images/icon_sche.gif" title="달력" style="cursor:hand;" align="absmiddle" name="RESERVE_CAR_4" onClick="window.open('/Include/Calendar_view.asp?goMonth='+document.ListForm.RESERVEDATE_4.value+'&firstcode=&FlightCode=&StartEndGu=S&ControlName=opener.document.ListForm.RESERVEDATE_4','','width=800,height=300,top=200,left=300,scrollbars=no,status=no');" onblur="setOutColor(this);" >&nbsp;<img src="/Images/Btn/BtnIconDel.gif" title="삭제" name="RESERVE_DEL_4" style="cursor:hand;" align="absmiddle"onclick="document.all.RESERVEDATE_4.value='';">&nbsp;&nbsp;<input value="<%=RESERVEHOUR_4%>" name="RESERVEHOUR_4" type="text" size="2" maxlength="2" onfocus="setFocusColor(this);" >시&nbsp;<input value="<%=RESERVEMIN_4%>" name="RESERVEMIN_4" type="text" size="2" maxlength="2" onfocus="setFocusColor(this);">분&nbsp;&nbsp;&nbsp;&nbsp;&nbsp<select name="RESERVETIME_4" size="1" class="ComboFFFCE7" onchange="fn_settime('4')">
									<Option value ='' selected>시간선택</option>
									<Option value ='1' >10분후</option>
									<Option value ='2' >30분후</option>
									<Option value ='3' >1시간후</option>
									<Option value ='4' >2시간후</option>
									<Option value ='08' >오전 7시</option>
									<Option value ='08' >오전 8시</option>
									<Option value ='09' >오전 9시</option>
									<Option value ='10' >오전10시</option>
									<Option value ='11' >오전11시</option>
									<Option value ='12' >오후12시</option>
									<Option value ='13' >오후13시</option>
									<Option value ='14' >오후14시</option>
									<Option value ='15' >오후15시</option>
									<Option value ='16' >오후16시</option>
									<Option value ='17' >오후17시</option>
									<Option value ='18' >오후18시</option>
									<Option value ='19' >오후19시</option>
									<Option value ='20' >오후20시</option>
									<Option value ='21' >오후21시</option>
									<Option value ='22' >오후22시</option>

								</select>
							</td>
						</tr>
						<tr height="40">
							<td bgcolor="#FFEEF9" class="TDCont" align='center' width=100>종합평가</td>
							<td bgcolor="#FFFFFF" colspan=6>&nbsp;<textarea name="Remark_4" style="width:99%; height:100%" wrap="soft" class="TextareaInput"><%=db_Remark_4%></textarea>
							</td>	
						</tr>
						<tr height="40">
							<td bgcolor="#FFEEF9" class="TDCont" align='center' width=100>비고</td>
							<td bgcolor="#FFFFFF" colspan=6>&nbsp;<textarea name="Remark1_4" style="width:99%; height:100%" wrap="soft" class="TextareaInput"><%=db_Remark1_4%></textarea>
							</td>	
						</tr>
					</table>
					</td>
					<td bgcolor="#EEF6FF" width=70 class="TDCont" align='center'>점수(평균):</td>
					<td bgcolor="#FFFFFF" width=140 class="TDCont" align='center' colspan=2><input type="text" name="TOT_4" value="<%=db_TOT_4%>" size="5" maxlength="5" style="border-width:0px ; border-color:#EEF6FF ; border-style:solid; text-align:right; font-color:#ff0000;font-size:15px;font-weight:bold" readonly ></td>
				</tr>
			</table>

<%if db_factPeoplenum_5 = "" then%>
			<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table width="920" height="30" border="0" cellspacing="0" cellpadding="0" align="center">
				<tr>
					<td align="left"><img src="/Images/Btn/BtnRegiAdd_P<%=right(sGubun,1)%>.GIF" style="cursor:hand;" align="absmiddle" onClick="fn_AddForm('ON','divFORM5','FRM5');">&nbsp;<img src="/Images/Btn/BtnRegiDel_P<%=right(sGubun,1)%>.GIF" style="cursor:hand;" align="absmiddle" onClick="fn_AddForm('OFF','divFORM4','FRM4');"></td>
					<td align='left' width="50%"><input type="checkbox" name="RecYN_4" value="Y" class="none"><font color='#0000ff'><b>녹취첨부</b></font></td>
				</tr>
			</table>
<%else%>
			<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table width="920" height="30" border="0" cellspacing="0" cellpadding="0" align="center">
				<tr>
					<td align="left"><img src="/Images/Btn/BtnRegiDel_P<%=right(sGubun,1)%>.GIF" style="cursor:hand;" align="absmiddle" onClick="fn_AddForm('OFF','divFORM4','FRM4');"></td>
					<td align='left' width="50%"><input type="checkbox" name="RecYN_4" value="Y" class="none"><font color='#0000ff'><b>녹취첨부</b></font></td>
				</tr>
			</table>
<%end if%>
		</td>
	</tr>
</table>
</span>

<%'====== 상담접수 폼 #5 시작 =======================================================================================%>
<%if db_factPeoplenum_5 = "" then%>
	<span id="divFORM5" style="display:none;">
<% else %>
	<span id="divFORM5" style="display:block;">
<%end if%>
<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
<table border="0" width="920" cellpadding="0" cellspacing="1" bgcolor="#EFECE5" align="center">
	<tr bgcolor="#FFFFFF">
		<td>
			<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table width="920" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#ffffff">
			    <tr>
					<td align="left" bgcolor="#FFFFFF" class="TDCont">&nbsp;<img src="/Images/dot_01.gif" >&nbsp;<b><font color="#ff00ff">(<%=sGubunName%>)</font> 설문지#5</b></td>
				</tr>
			</table>
			<table width="920" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">
			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>관 계</td>
					<td bgcolor="#FFFFFF">						<%
							'======= 처리구분 코드 가져오기 ==================================================
							SqlCode = "SELECT CODE, CODENAME FROM TB_CODE"
							SqlCode = SqlCode& " WHERE USEYN='Y' AND SYSYN='Y' AND CODEGROUP='B01'"
							SqlCode = SqlCode& " ORDER BY CODE"
							set RsCode = db.execute(SqlCode)
						%>
						&nbsp;<select name="SECTION2_5" size="1" class="ComboFFFCE7" onChange="fn_UpdateData('5','SECTION2_');">
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("CODE")
										CODENAME = RsCode("CODENAME")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &sGubun& "")%>
							<%
									RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
							%>
						</select>
					&nbsp;&nbsp;&nbsp;&nbsp;<font color="#0000ff"><%=db_factPeoplenum_5%></font></td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>계급(신분)</td>
					<td bgcolor="#FFFFFF"><%
							'======= 처리구분 코드 가져오기 ==================================================
							SqlCode = "SELECT *	FROM armyinformix.dbo.data"
							SqlCode = SqlCode& " where [restrict] = '104' order by [restrict]"
							set RsCode = db.execute(SqlCode)
						%>
						&nbsp;<select name="LEVEL_5" size="1" class="ComboFFFCE7">
							<option value="">계급선택</option>
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("number")
										CODENAME = RsCode("contents")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &db_level_5& "")%>
							<%
									RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
							%>
						</select></td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>성  명</td>
					<td bgcolor="#FFFFFF">&nbsp;<input type="text" name="NAME_5" value="<%=db_NAME_5%>" size="20" maxlength="20" style="border-width:1px ; border-color:#cccccc ; border-style:solid" ></td>
				</tr>
			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>연락처 1</td>
					<td bgcolor="#FFFFFF">&nbsp;<input type="text" name="HOMEPHONE_5" value="<%=db_HOMEPHONE_5%>" size="13" maxlength="13" style="border-width:1px ; border-color:#cccccc ; border-style:solid">
					<img src="/Images/Cti/icon_tel.gif" style="cursor:hand;" onClick="fn_dial('1','5');" align="absmiddle" title="군전화로 전화걸기">&nbsp<img src="/Images/Cti/icon_tel_1.gif" style="cursor:hand;" onClick="fn_dial_1('1','5');" align="absmiddle" title="일반전화로 전화걸기">&nbsp<img src="/Images/Cti/icon_sche.gif" style="cursor:hand;" onClick="fn_sms('1','5');" align="absmiddle" title="문자전송"></td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>연락처 2</td>
					<td bgcolor="#FFFFFF">&nbsp;<input type="text" name="MOBILEPHONE_5" value="<%=db_MOBILEPHONE_5%>" size="13" maxlength="13" style="border-width:1px ; border-color:#cccccc ; border-style:solid">&nbsp;<img src="/Images/Cti/icon_tel.gif" style="cursor:hand;" onClick="fn_dial('2','5');" align="absmiddle" title="군전화로 전화걸기">&nbsp<img src="/Images/Cti/icon_tel_1.gif" style="cursor:hand;" onClick="fn_dial_1('2','5');" align="absmiddle" title="일반전화로 전화걸기">&nbsp<img src="/Images/Cti/icon_sche.gif" style="cursor:hand;" onClick="fn_sms('2','5');" align="absmiddle" title="문자전송"></td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>연락처비고</td>
					<td bgcolor="#FFFFFF">&nbsp;<input type="text" name="ETCPHONE_5" value="<%=db_ETCPHONE_5%>" size="20" maxlength="20" style="border-width:1px ; border-color:#cccccc ; border-style:solid">
					</td>
						
				</tr>
			    <tr>

					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>설문일시</td>
					<td bgcolor="#FFFFFF" width=200>&nbsp;<input type="text" name="MONITORDATE_5" value="<%=db_MONITORDATE_5%>" size="10" maxlength="10" style="border-width:1px ; border-color:#cccccc ; border-style:solid" >&nbsp;<img src="/Images/icon_sche.gif" title="달력" style="cursor:hand;" align="absmiddle" name="Date2_CAR" onClick="window.open('/Include/Calendar_view.asp?goMonth='+document.all.MONITORDATE_5.value+'&firstcode=&FlightCode=&StartEndGu=S&ControlName=opener.document.all.MONITORDATE_5','','width=800,height=300,top=200,left=300,scrollbars=no,status=no');" onblur="setOutColor(this);">
						&nbsp;<img src="/Images/Btn/BtnIconDel.gif" title="삭제" style="cursor:hand;" align="absmiddle"onclick="document.all.MONITORDATE_5.value='';">
					</td>

					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>설문결과</td>
					<td bgcolor="#FFFFFF" width=200>&nbsp;<font color="#0000ff"><%=db_getCodeName("B10",db_MONITOR_RESULT_5)%></font>
					</td>

					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>예약일시</td>
					<td bgcolor="#FFFFFF" width=200>&nbsp;<%=db_RESERVEDATE_5%>
					</td>
				</tr>
				<tr>
					<td bgcolor="#FFEEF9" class="TDCont" align='center' width=100>통화내역</td>
					<td bgcolor="#FFFFFF" colspan=6>&nbsp;<span id="HISTORY_5"><%=db_History_5%></span>
					</td>	
				</tr>

			</table>
			<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table width="920" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">
			    <tr height="25">
					<td bgcolor="#EEF6FF" width=530 rowspan=2 class="TDCont"  colspan='2' align='center'>질문사항</td>
					<td bgcolor="#EEF6FF" colspan='3' class="TDCont" align='center' width=210>답변결과</td>
					<td bgcolor="#EEF6FF" rowspan=2 class="TDCont" align='center' width=70>가점<br>(+1)</td>
					<td bgcolor="#EEF6FF" rowspan=2 class="TDCont" align='center' width=70>점수</td>
					<td bgcolor="#EEF6FF" rowspan=2 class="TDCont" align='center' width=70>초기화<br><img src="/Images/Btn/BtnIconDel.gif" title="점수초기화" style="cursor:hand;" align="absmiddle"onclick="javascript:fn_ALLDEL('5','<%=i%>');"></td>
				</tr>
			    <tr height="20">
					<td bgcolor="#EEF6FF" class="TDCont" align='center' width=70>만족<br>(9)</td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center' width=70>보통<br>(8)</td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center' width=70>불만족<br>(7)</td>
				</tr>
<%

			SQL = "SELECT * FROM TB_CODE where CODEGROUP = '" & sGubun &"' AND USEYN = 'Y' ORDER BY CODE"
			SET Rs = DB.execute(SQL)			

			i = 0
			do until rs.eof
				i = i + 1
				if ( i mod 2 ) = 1 then
					sBgColor = "#ffffff"
				else
					sBgColor = "#FFFCE7"				
				end if

				'값 불러오기
				SQL1 = "select *"
				SQL1 = SQL1 & " from armyinformix.dbo.monitor where factnum = '" & receiptfactnum & "' and factpeoplenum='"& db_factPeoplenum_5 & "' and  seqno = " & i
				
				SET Rs1 = DB.execute(SQL1)	
				if Rs1.eof = false then
					point9 = Rs1("point9")
					point8 = Rs1("point8")
					point7 = Rs1("point7")
					pointplus = Rs1("pointplus")
					totpoint = Rs1("totpoint")
				else
					point9 = ""
					point8 = ""
					point7 = ""
					pointplus = ""
					totpoint = ""
				end if
%>				
			    <tr>
					<td bgcolor="<%=sBgColor%>" width=530  class="TDCont"  colspan='2'>&nbsp;<%=rs("codename")%></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="radio" name="QUESTION_5<%=i%>" value="9" class="none" onClick="fn_YES('5','<%=i%>','9');" <%if point9 = "1" then%>checked<%end if%> ></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="radio" name="QUESTION_5<%=i%>" value="8" class="none" onClick="fn_YES('5','<%=i%>','8');" <%if point8 = "1" then%>checked<%end if%> ></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="radio" name="QUESTION_5<%=i%>" value="7" class="none" onClick="fn_YES('5','<%=i%>','7');" <%if point7 = "1" then%>checked<%end if%> ></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="checkbox" name="QUESTIONP_5<%=i%>" value="1" class="none" onClick="fn_YES('5','<%=i%>','1');"<% if pointplus="1" then Response.Write("checked") end if %>>					
					</td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="text" name="POINT_5<%=i%>" value="<%=totpoint%>" size="2" maxlength="2" style="border-width:0px ; border-color:#cccccc ; border-style:solid; text-align:right; background-color:<%=sBgColor%>" readonly></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><img src="/Images/Btn/BtnIconDel.gif" title="삭제" style="cursor:hand;" align="absmiddle"onclick="javascript:fn_DEL('5','<%=i%>');"></td>

				</tr>
<%
				rs.movenext
			loop
%>
			    <tr ><td bgcolor="#ffffff" class="TDCont"  colspan='5' valign="top" height="100">
					<table width="100%" height="100%" border="0" cellspacing="1" cellpadding="0" align="center" bgcolor="#CCCCCC">
						<tr height="30">
							<td bgcolor="#FFEEF9" class="TDCont" width=100 align='center'>설문결과</td>
							<td bgcolor="#FFFFFF" width="120">
								<%
									'======= 처리구분 코드 가져오기 ==================================================
									SqlCode = "SELECT CODE, CODENAME FROM TB_CODE"
									SqlCode = SqlCode& " WHERE USEYN='Y' AND SYSYN='Y' AND CODEGROUP='B10'"
									SqlCode = SqlCode& " ORDER BY CODE"
									set RsCode = db.execute(SqlCode)
								%>
								&nbsp;<select name="MONITORRESULT_5" size="1" class="ComboFFFCE7" onChange="fn_ResultSet('5');">
									<option value="">설문결과선택</option>
									<%
										IF NOT(RsCode.Eof OR RsCode.bof) THEN
											DO until RsCode.EOF
												CODE = RsCode("CODE")
												CODENAME = RsCode("CODENAME")
									%>
									<%=printSelect("" &CODENAME& "","" &CODE& "","" &db_MONITOR_RESULT_5& "")%>
									<%
											RsCode.MoveNext
											LOOP
										END IF
										RsCode.Close
										set RsCode = NOTHING
									%>
								</select>
							
							</td>


							<td bgcolor="#FFEEF9" class="TDCont" align='center' width="100">상담예약일시</td>
							<td bgcolor="#FFFFFF">&nbsp;<input value="<%=RESERVEDATE_5%>" name="RESERVEDATE_5" readonly type="text" size="10" onfocus="setFocusColor(this);" onClick="new CalendarFrame.Calendar(this);" onblur="setOutColor(this);" onchange="fn_settime('5')">&nbsp;<img src="/Images/icon_sche.gif" title="달력" style="cursor:hand;" align="absmiddle" name="RESERVE_CAR_5" onClick="window.open('/Include/Calendar_view.asp?goMonth='+document.ListForm.RESERVEDATE_5.value+'&firstcode=&FlightCode=&StartEndGu=S&ControlName=opener.document.ListForm.RESERVEDATE_5','','width=800,height=300,top=200,left=300,scrollbars=no,status=no');" onblur="setOutColor(this);" >&nbsp;<img src="/Images/Btn/BtnIconDel.gif" title="삭제" name="RESERVE_DEL_5" style="cursor:hand;" align="absmiddle"onclick="document.all.RESERVEDATE_5.value='';">&nbsp;&nbsp;<input value="<%=RESERVEHOUR_5%>" name="RESERVEHOUR_5" type="text" size="2" maxlength="2" onfocus="setFocusColor(this);" >시&nbsp;<input value="<%=RESERVEMIN_5%>" name="RESERVEMIN_5" type="text" size="2" maxlength="2" onfocus="setFocusColor(this);">분&nbsp;&nbsp;&nbsp;&nbsp;&nbsp<select name="RESERVETIME_5" size="1" class="ComboFFFCE7" onchange="fn_settime('5')">
									<Option value ='' selected>시간선택</option>
									<Option value ='1' >10분후</option>
									<Option value ='2' >30분후</option>
									<Option value ='3' >1시간후</option>
									<Option value ='4' >2시간후</option>
									<Option value ='08' >오전 7시</option>
									<Option value ='08' >오전 8시</option>
									<Option value ='09' >오전 9시</option>
									<Option value ='10' >오전10시</option>
									<Option value ='11' >오전11시</option>
									<Option value ='12' >오후12시</option>
									<Option value ='13' >오후13시</option>
									<Option value ='14' >오후14시</option>
									<Option value ='15' >오후15시</option>
									<Option value ='16' >오후16시</option>
									<Option value ='17' >오후17시</option>
									<Option value ='18' >오후18시</option>
									<Option value ='19' >오후19시</option>
									<Option value ='20' >오후20시</option>
									<Option value ='21' >오후21시</option>
									<Option value ='22' >오후22시</option>

								</select>
							</td>
						</tr>
						<tr height="40">
							<td bgcolor="#FFEEF9" class="TDCont" align='center' width=100>종합평가</td>
							<td bgcolor="#FFFFFF" colspan=6>&nbsp;<textarea name="Remark_5" style="width:99%; height:100%" wrap="soft" class="TextareaInput"><%=db_Remark_5%></textarea>
							</td>	
						</tr>
						<tr height="40">
							<td bgcolor="#FFEEF9" class="TDCont" align='center' width=100>비고</td>
							<td bgcolor="#FFFFFF" colspan=6>&nbsp;<textarea name="Remark1_5" style="width:99%; height:100%" wrap="soft" class="TextareaInput"><%=db_Remark1_5%></textarea>
							</td>	
						</tr>
					</table>
					</td>
					<td bgcolor="#EEF6FF" width=70 class="TDCont" align='center'>점수(평균):</td>
					<td bgcolor="#FFFFFF" width=140 class="TDCont" align='center' colspan=2><input type="text" name="TOT_5" value="<%=db_TOT_5%>" size="5" maxlength="5" style="border-width:0px ; border-color:#EEF6FF ; border-style:solid; text-align:right; font-color:#ff0000;font-size:15px;font-weight:bold" readonly ></td>
				</tr>
			</table>

<%if db_factPeoplenum_6 = "" then%>
			<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table width="920" height="30" border="0" cellspacing="0" cellpadding="0" align="center">
				<tr>
					<td align="left"><img src="/Images/Btn/BtnRegiAdd_P<%=right(sGubun,1)%>.GIF" style="cursor:hand;" align="absmiddle" onClick="fn_AddForm('ON','divFORM6','FRM6');">&nbsp;<img src="/Images/Btn/BtnRegiDel_P<%=right(sGubun,1)%>.GIF" style="cursor:hand;" align="absmiddle" onClick="fn_AddForm('OFF','divFORM5','FRM5');"></td>
					<td align='left' width="50%"><input type="checkbox" name="RecYN_5" value="Y" class="none"><font color='#0000ff'><b>녹취첨부</b></font></td>
				</tr>
			</table>
<%else%>
			<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table width="920" height="30" border="0" cellspacing="0" cellpadding="0" align="center">
				<tr>
					<td align="left"><img src="/Images/Btn/BtnRegiDel_P<%=right(sGubun,1)%>.GIF" style="cursor:hand;" align="absmiddle" onClick="fn_AddForm('OFF','divFORM5','FRM5');"></td>
					<td align='left' width="50%"><input type="checkbox" name="RecYN_5" value="Y" class="none"><font color='#0000ff'><b>녹취첨부</b></font></td>
				</tr>
			</table>
<%end if%>

		</td>
	</tr>
</table>
</span>

<%'====== 상담접수 폼 #6 시작 =======================================================================================%>
<%if db_factPeoplenum_6 = "" then%>
	<span id="divFORM6" style="display:none;">
<% else %>
	<span id="divFORM6" style="display:block;">
<%end if%>
<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
<table border="0" width="920" cellpadding="0" cellspacing="1" bgcolor="#EFECE5" align="center">
	<tr bgcolor="#FFFFFF">
		<td>
			<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table width="920" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#ffffff">
			    <tr>
					<td align="left" bgcolor="#FFFFFF" class="TDCont">&nbsp;<img src="/Images/dot_01.gif" >&nbsp;<b><font color="#ff00ff">(<%=sGubunName%>)</font> 설문지#6</b></td>
				</tr>
			</table>
			<table width="920" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">
			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>관 계</td>
					<td bgcolor="#FFFFFF">						<%
							'======= 처리구분 코드 가져오기 ==================================================
							SqlCode = "SELECT CODE, CODENAME FROM TB_CODE"
							SqlCode = SqlCode& " WHERE USEYN='Y' AND SYSYN='Y' AND CODEGROUP='B01'"
							SqlCode = SqlCode& " ORDER BY CODE"
							set RsCode = db.execute(SqlCode)
						%>
						&nbsp;<select name="SECTION2_6" size="1" class="ComboFFFCE7" onChange="fn_UpdateData('6','SECTION2_');">
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("CODE")
										CODENAME = RsCode("CODENAME")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &sGubun& "")%>
							<%
									RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
							%>
						</select>
					&nbsp;&nbsp;&nbsp;&nbsp;<font color="#0000ff"><%=db_factPeoplenum_6%></font></td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>계급(신분)</td>
					<td bgcolor="#FFFFFF"><%
							'======= 처리구분 코드 가져오기 ==================================================
							SqlCode = "SELECT *	FROM armyinformix.dbo.data"
							SqlCode = SqlCode& " where [restrict] = '104' order by [restrict]"
							set RsCode = db.execute(SqlCode)
						%>
						&nbsp;<select name="LEVEL_6" size="1" class="ComboFFFCE7">
							<option value="">계급선택</option>
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("number")
										CODENAME = RsCode("contents")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &db_level_6& "")%>
							<%
									RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
							%>
						</select></td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>성  명</td>
					<td bgcolor="#FFFFFF">&nbsp;<input type="text" name="NAME_6" value="<%=db_NAME_6%>" size="20" maxlength="20" style="border-width:1px ; border-color:#cccccc ; border-style:solid" ></td>
				</tr>
			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>연락처 1</td>
					<td bgcolor="#FFFFFF">&nbsp;<input type="text" name="HOMEPHONE_6" value="<%=db_HOMEPHONE_6%>" size="13" maxlength="13" style="border-width:1px ; border-color:#cccccc ; border-style:solid">
					<img src="/Images/Cti/icon_tel.gif" style="cursor:hand;" onClick="fn_dial('1','6');" align="absmiddle" title="군전화로 전화걸기">&nbsp<img src="/Images/Cti/icon_tel_1.gif" style="cursor:hand;" onClick="fn_dial_1('1','6');" align="absmiddle" title="일반전화로 전화걸기">&nbsp<img src="/Images/Cti/icon_sche.gif" style="cursor:hand;" onClick="fn_sms('1','6');" align="absmiddle" title="문자전송"></td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>연락처 2</td>
					<td bgcolor="#FFFFFF">&nbsp;<input type="text" name="MOBILEPHONE_6" value="<%=db_MOBILEPHONE_6%>" size="13" maxlength="13" style="border-width:1px ; border-color:#cccccc ; border-style:solid">&nbsp;<img src="/Images/Cti/icon_tel.gif" style="cursor:hand;" onClick="fn_dial('2','6');" align="absmiddle" title="군전화로 전화걸기">&nbsp<img src="/Images/Cti/icon_tel_1.gif" style="cursor:hand;" onClick="fn_dial_1('2','6');" align="absmiddle" title="일반전화로 전화걸기">&nbsp<img src="/Images/Cti/icon_sche.gif" style="cursor:hand;" onClick="fn_sms('2','6');" align="absmiddle" title="문자전송"></td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>연락처비고</td>
					<td bgcolor="#FFFFFF">&nbsp;<input type="text" name="ETCPHONE_6" value="<%=db_ETCPHONE_6%>" size="20" maxlength="20" style="border-width:1px ; border-color:#cccccc ; border-style:solid">
					</td>
						
				</tr>
			    <tr>

					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>설문일시</td>
					<td bgcolor="#FFFFFF" width=200>&nbsp;<input type="text" name="MONITORDATE_6" value="<%=db_MONITORDATE_6%>" size="10" maxlength="10" style="border-width:1px ; border-color:#cccccc ; border-style:solid" >&nbsp;<img src="/Images/icon_sche.gif" title="달력" style="cursor:hand;" align="absmiddle" name="Date2_CAR" onClick="window.open('/Include/Calendar_view.asp?goMonth='+document.all.MONITORDATE_6.value+'&firstcode=&FlightCode=&StartEndGu=S&ControlName=opener.document.all.MONITORDATE_6','','width=800,height=300,top=200,left=300,scrollbars=no,status=no');" onblur="setOutColor(this);">
						&nbsp;<img src="/Images/Btn/BtnIconDel.gif" title="삭제" style="cursor:hand;" align="absmiddle"onclick="document.all.MONITORDATE_6.value='';">
					</td>

					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>설문결과</td>
					<td bgcolor="#FFFFFF" width=200>&nbsp;<font color="#0000ff"><%=db_getCodeName("B10",db_MONITOR_RESULT_6)%></font>
					</td>

					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>예약일시</td>
					<td bgcolor="#FFFFFF" width=200>&nbsp;<%=db_RESERVEDATE_6%>
					</td>
				</tr>
				<tr>
					<td bgcolor="#FFEEF9" class="TDCont" align='center' width=100>통화내역</td>
					<td bgcolor="#FFFFFF" colspan=6>&nbsp;<span id="HISTORY_6"><%=db_History_6%></span>
					</td>	
				</tr>

			</table>
			<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table width="920" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">
			    <tr height="25">
					<td bgcolor="#EEF6FF" width=530 rowspan=2 class="TDCont"  colspan='2' align='center'>질문사항</td>
					<td bgcolor="#EEF6FF" colspan='3' class="TDCont" align='center' width=210>답변결과</td>
					<td bgcolor="#EEF6FF" rowspan=2 class="TDCont" align='center' width=70>가점<br>(+1)</td>
					<td bgcolor="#EEF6FF" rowspan=2 class="TDCont" align='center' width=70>점수</td>
					<td bgcolor="#EEF6FF" rowspan=2 class="TDCont" align='center' width=70>초기화<br><img src="/Images/Btn/BtnIconDel.gif" title="점수초기화" style="cursor:hand;" align="absmiddle"onclick="javascript:fn_ALLDEL('6','<%=i%>');"></td>
				</tr>
			    <tr height="20">
					<td bgcolor="#EEF6FF" class="TDCont" align='center' width=70>만족<br>(9)</td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center' width=70>보통<br>(8)</td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center' width=70>불만족<br>(7)</td>
				</tr>
<%

			SQL = "SELECT * FROM TB_CODE where CODEGROUP = '" & sGubun &"' AND USEYN = 'Y' ORDER BY CODE"
			SET Rs = DB.execute(SQL)			

			i = 0
			do until rs.eof
				i = i + 1
				if ( i mod 2 ) = 1 then
					sBgColor = "#ffffff"
				else
					sBgColor = "#FFFCE7"				
				end if

				'값 불러오기
				SQL1 = "select *"
				SQL1 = SQL1 & " from armyinformix.dbo.monitor where factnum = '" & receiptfactnum & "' and factpeoplenum='"& db_factPeoplenum_6 & "' and  seqno = " & i
				
				SET Rs1 = DB.execute(SQL1)	
				if Rs1.eof = false then
					point9 = Rs1("point9")
					point8 = Rs1("point8")
					point7 = Rs1("point7")
					pointplus = Rs1("pointplus")
					totpoint = Rs1("totpoint")
				else
					point9 = ""
					point8 = ""
					point7 = ""
					pointplus = ""
					totpoint = ""
				end if
%>				
			    <tr>
					<td bgcolor="<%=sBgColor%>" width=530  class="TDCont"  colspan='2'>&nbsp;<%=rs("codename")%></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="radio" name="QUESTION_6<%=i%>" value="9" class="none" onClick="fn_YES('6','<%=i%>','9');" <%if point9 = "1" then%>checked<%end if%> ></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="radio" name="QUESTION_6<%=i%>" value="8" class="none" onClick="fn_YES('6','<%=i%>','8');" <%if point8 = "1" then%>checked<%end if%> ></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="radio" name="QUESTION_6<%=i%>" value="7" class="none" onClick="fn_YES('6','<%=i%>','7');" <%if point7 = "1" then%>checked<%end if%> ></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="checkbox" name="QUESTIONP_6<%=i%>" value="1" class="none" onClick="fn_YES('6','<%=i%>','1');"<% if pointplus="1" then Response.Write("checked") end if %>>					
					</td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="text" name="POINT_6<%=i%>" value="<%=totpoint%>" size="2" maxlength="2" style="border-width:0px ; border-color:#cccccc ; border-style:solid; text-align:right; background-color:<%=sBgColor%>" readonly></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><img src="/Images/Btn/BtnIconDel.gif" title="삭제" style="cursor:hand;" align="absmiddle"onclick="javascript:fn_DEL('6','<%=i%>');"></td>

				</tr>
<%
				rs.movenext
			loop
%>
			    <tr ><td bgcolor="#ffffff" class="TDCont"  colspan='5' valign="top" height="100">
					<table width="100%" height="100%" border="0" cellspacing="1" cellpadding="0" align="center" bgcolor="#CCCCCC">
						<tr height="30">
							<td bgcolor="#FFEEF9" class="TDCont" width=100 align='center'>설문결과</td>
							<td bgcolor="#FFFFFF" width="120">
								<%
									'======= 처리구분 코드 가져오기 ==================================================
									SqlCode = "SELECT CODE, CODENAME FROM TB_CODE"
									SqlCode = SqlCode& " WHERE USEYN='Y' AND SYSYN='Y' AND CODEGROUP='B10'"
									SqlCode = SqlCode& " ORDER BY CODE"
									set RsCode = db.execute(SqlCode)
								%>
								&nbsp;<select name="MONITORRESULT_6" size="1" class="ComboFFFCE7" onChange="fn_ResultSet('6');">
									<option value="">설문결과선택</option>
									<%
										IF NOT(RsCode.Eof OR RsCode.bof) THEN
											DO until RsCode.EOF
												CODE = RsCode("CODE")
												CODENAME = RsCode("CODENAME")
									%>
									<%=printSelect("" &CODENAME& "","" &CODE& "","" &db_MONITOR_RESULT_6& "")%>
									<%
											RsCode.MoveNext
											LOOP
										END IF
										RsCode.Close
										set RsCode = NOTHING
									%>
								</select>
							
							</td>


							<td bgcolor="#FFEEF9" class="TDCont" align='center' width="100">상담예약일시</td>
							<td bgcolor="#FFFFFF">&nbsp;<input value="<%=RESERVEDATE_5%>" name="RESERVEDATE_6" readonly type="text" size="10" onfocus="setFocusColor(this);" onClick="new CalendarFrame.Calendar(this);" onblur="setOutColor(this);" onchange="fn_settime('6')">&nbsp;<img src="/Images/icon_sche.gif" title="달력" style="cursor:hand;" align="absmiddle" name="RESERVE_CAR_6" onClick="window.open('/Include/Calendar_view.asp?goMonth='+document.ListForm.RESERVEDATE_6.value+'&firstcode=&FlightCode=&StartEndGu=S&ControlName=opener.document.ListForm.RESERVEDATE_6','','width=800,height=300,top=200,left=300,scrollbars=no,status=no');" onblur="setOutColor(this);" >&nbsp;<img src="/Images/Btn/BtnIconDel.gif" title="삭제" name="RESERVE_DEL_6" style="cursor:hand;" align="absmiddle"onclick="document.all.RESERVEDATE_6.value='';">&nbsp;&nbsp;<input value="<%=RESERVEHOUR_6%>" name="RESERVEHOUR_6" type="text" size="2" maxlength="2" onfocus="setFocusColor(this);" >시&nbsp;<input value="<%=RESERVEMIN_6%>" name="RESERVEMIN_6" type="text" size="2" maxlength="2" onfocus="setFocusColor(this);">분&nbsp;&nbsp;&nbsp;&nbsp;&nbsp<select name="RESERVETIME_6" size="1" class="ComboFFFCE7" onchange="fn_settime('6')">
									<Option value ='' selected>시간선택</option>
									<Option value ='1' >10분후</option>
									<Option value ='2' >30분후</option>
									<Option value ='3' >1시간후</option>
									<Option value ='4' >2시간후</option>
									<Option value ='08' >오전 7시</option>
									<Option value ='08' >오전 8시</option>
									<Option value ='09' >오전 9시</option>
									<Option value ='10' >오전10시</option>
									<Option value ='11' >오전11시</option>
									<Option value ='12' >오후12시</option>
									<Option value ='13' >오후13시</option>
									<Option value ='14' >오후14시</option>
									<Option value ='15' >오후15시</option>
									<Option value ='16' >오후16시</option>
									<Option value ='17' >오후17시</option>
									<Option value ='18' >오후18시</option>
									<Option value ='19' >오후19시</option>
									<Option value ='20' >오후20시</option>
									<Option value ='21' >오후21시</option>
									<Option value ='22' >오후22시</option>

								</select>
							</td>
						</tr>
						<tr height="40">
							<td bgcolor="#FFEEF9" class="TDCont" align='center' width=100>종합평가</td>
							<td bgcolor="#FFFFFF" colspan=6>&nbsp;<textarea name="Remark_6" style="width:99%; height:100%" wrap="soft" class="TextareaInput"><%=db_Remark_6%></textarea>
							</td>	
						</tr>
						<tr height="40">
							<td bgcolor="#FFEEF9" class="TDCont" align='center' width=100>비고</td>
							<td bgcolor="#FFFFFF" colspan=6>&nbsp;<textarea name="Remark1_6" style="width:99%; height:100%" wrap="soft" class="TextareaInput"><%=db_Remark1_6%></textarea>
							</td>	
						</tr>
					</table>
					</td>
					<td bgcolor="#EEF6FF" width=70 class="TDCont" align='center'>점수(평균):</td>
					<td bgcolor="#FFFFFF" width=140 class="TDCont" align='center' colspan=2><input type="text" name="TOT_6" value="<%=db_TOT_6%>" size="5" maxlength="5" style="border-width:0px ; border-color:#EEF6FF ; border-style:solid; text-align:right; font-color:#ff0000;font-size:15px;font-weight:bold" readonly ></td>
				</tr>
			</table>

<%if db_factPeoplenum_7 = "" then%>
			<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table width="920" height="30" border="0" cellspacing="0" cellpadding="0" align="center">
				<tr>
					<td align="left"><img src="/Images/Btn/BtnRegiAdd_P<%=right(sGubun,1)%>.GIF" style="cursor:hand;" align="absmiddle" onClick="fn_AddForm('ON','divFORM7','FRM7');">&nbsp;<img src="/Images/Btn/BtnRegiDel_P<%=right(sGubun,1)%>.GIF" style="cursor:hand;" align="absmiddle" onClick="fn_AddForm('OFF','divFORM6','FRM6');"></td>
					<td align='left' width="50%"><input type="checkbox" name="RecYN_6" value="Y" class="none"><font color='#0000ff'><b>녹취첨부</b></font></td>
				</tr>
			</table>
<%else%>
			<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table width="920" height="30" border="0" cellspacing="0" cellpadding="0" align="center">
				<tr>
					<td align="left"><img src="/Images/Btn/BtnRegiDel_P<%=right(sGubun,1)%>.GIF" style="cursor:hand;" align="absmiddle" onClick="fn_AddForm('OFF','divFORM6','FRM6');"></td>
					<td align='left' width="50%"><input type="checkbox" name="RecYN_6" value="Y" class="none"><font color='#0000ff'><b>녹취첨부</b></font></td>
				</tr>
			</table>
<%end if%>

		</td>
	</tr>
</table>
</span>

<%'====== 상담접수 폼 #7 시작 =======================================================================================%>
<%if db_factPeoplenum_7 = "" then%>
	<span id="divFORM7" style="display:none;">
<% else %>
	<span id="divFORM7" style="display:block;">
<%end if%>
<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
<table border="0" width="920" cellpadding="0" cellspacing="1" bgcolor="#EFECE5" align="center">
	<tr bgcolor="#FFFFFF">
		<td>
			<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table width="920" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#ffffff">
			    <tr>
					<td align="left" bgcolor="#FFFFFF" class="TDCont">&nbsp;<img src="/Images/dot_01.gif" >&nbsp;<b><font color="#ff00ff">(<%=sGubunName%>)</font> 설문지#7</b></td>
				</tr>
			</table>
			<table width="920" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">
			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>관 계</td>
					<td bgcolor="#FFFFFF">						<%
							'======= 처리구분 코드 가져오기 ==================================================
							SqlCode = "SELECT CODE, CODENAME FROM TB_CODE"
							SqlCode = SqlCode& " WHERE USEYN='Y' AND SYSYN='Y' AND CODEGROUP='B01'"
							SqlCode = SqlCode& " ORDER BY CODE"
							set RsCode = db.execute(SqlCode)
						%>
						&nbsp;<select name="SECTION2_7" size="1" class="ComboFFFCE7" onChange="fn_UpdateData('7','SECTION2_');">
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("CODE")
										CODENAME = RsCode("CODENAME")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &sGubun& "")%>
							<%
									RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
							%>
						</select>
					&nbsp;&nbsp;&nbsp;&nbsp;<font color="#0000ff"><%=db_factPeoplenum_7%></font></td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>계급(신분)</td>
					<td bgcolor="#FFFFFF"><%
							'======= 처리구분 코드 가져오기 ==================================================
							SqlCode = "SELECT *	FROM armyinformix.dbo.data"
							SqlCode = SqlCode& " where [restrict] = '104' order by [restrict]"
							set RsCode = db.execute(SqlCode)
						%>
						&nbsp;<select name="LEVEL_7" size="1" class="ComboFFFCE7">
							<option value="">계급선택</option>
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("number")
										CODENAME = RsCode("contents")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &db_level_7& "")%>
							<%
									RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
							%>
						</select></td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>성  명</td>
					<td bgcolor="#FFFFFF">&nbsp;<input type="text" name="NAME_7" value="<%=db_NAME_7%>" size="20" maxlength="20" style="border-width:1px ; border-color:#cccccc ; border-style:solid" ></td>
				</tr>
			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>연락처 1</td>
					<td bgcolor="#FFFFFF">&nbsp;<input type="text" name="HOMEPHONE_7" value="<%=db_HOMEPHONE_7%>" size="13" maxlength="13" style="border-width:1px ; border-color:#cccccc ; border-style:solid">
					<img src="/Images/Cti/icon_tel.gif" style="cursor:hand;" onClick="fn_dial('1','7');" align="absmiddle" title="군전화로 전화걸기">&nbsp<img src="/Images/Cti/icon_tel_1.gif" style="cursor:hand;" onClick="fn_dial_1('1','7');" align="absmiddle" title="일반전화로 전화걸기">&nbsp<img src="/Images/Cti/icon_sche.gif" style="cursor:hand;" onClick="fn_sms('1','7');" align="absmiddle" title="문자전송"></td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>연락처 2</td>
					<td bgcolor="#FFFFFF">&nbsp;<input type="text" name="MOBILEPHONE_7" value="<%=db_MOBILEPHONE_7%>" size="13" maxlength="13" style="border-width:1px ; border-color:#cccccc ; border-style:solid">&nbsp;<img src="/Images/Cti/icon_tel.gif" style="cursor:hand;" onClick="fn_dial('2','7');" align="absmiddle" title="군전화로 전화걸기">&nbsp<img src="/Images/Cti/icon_tel_1.gif" style="cursor:hand;" onClick="fn_dial_1('2','7');" align="absmiddle" title="일반전화로 전화걸기">&nbsp<img src="/Images/Cti/icon_sche.gif" style="cursor:hand;" onClick="fn_sms('2','7');" align="absmiddle" title="문자전송"></td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>연락처비고</td>
					<td bgcolor="#FFFFFF">&nbsp;<input type="text" name="ETCPHONE_7" value="<%=db_ETCPHONE_7%>" size="20" maxlength="20" style="border-width:1px ; border-color:#cccccc ; border-style:solid">
					</td>
						
				</tr>
			    <tr>

					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>설문일시</td>
					<td bgcolor="#FFFFFF" width=200>&nbsp;<input type="text" name="MONITORDATE_7" value="<%=db_MONITORDATE_7%>" size="10" maxlength="10" style="border-width:1px ; border-color:#cccccc ; border-style:solid" >&nbsp;<img src="/Images/icon_sche.gif" title="달력" style="cursor:hand;" align="absmiddle" name="Date2_CAR" onClick="window.open('/Include/Calendar_view.asp?goMonth='+document.all.MONITORDATE_7.value+'&firstcode=&FlightCode=&StartEndGu=S&ControlName=opener.document.all.MONITORDATE_7','','width=800,height=300,top=200,left=300,scrollbars=no,status=no');" onblur="setOutColor(this);">
						&nbsp;<img src="/Images/Btn/BtnIconDel.gif" title="삭제" style="cursor:hand;" align="absmiddle"onclick="document.all.MONITORDATE_7.value='';">
					</td>

					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>설문결과</td>
					<td bgcolor="#FFFFFF" width=200>&nbsp;<font color="#0000ff"><%=db_getCodeName("B10",db_MONITOR_RESULT_7)%></font>
					</td>

					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>예약일시</td>
					<td bgcolor="#FFFFFF" width=200>&nbsp;<%=db_RESERVEDATE_7%>
					</td>
				</tr>
				<tr>
					<td bgcolor="#FFEEF9" class="TDCont" align='center' width=100>통화내역</td>
					<td bgcolor="#FFFFFF" colspan=6>&nbsp;<span id="HISTORY_7"><%=db_History_7%></span>
					</td>	
				</tr>

			</table>
			<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table width="920" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">
			    <tr height="25">
					<td bgcolor="#EEF6FF" width=530 rowspan=2 class="TDCont"  colspan='2' align='center'>질문사항</td>
					<td bgcolor="#EEF6FF" colspan='3' class="TDCont" align='center' width=210>답변결과</td>
					<td bgcolor="#EEF6FF" rowspan=2 class="TDCont" align='center' width=70>가점<br>(+1)</td>
					<td bgcolor="#EEF6FF" rowspan=2 class="TDCont" align='center' width=70>점수</td>
					<td bgcolor="#EEF6FF" rowspan=2 class="TDCont" align='center' width=70>초기화<br><img src="/Images/Btn/BtnIconDel.gif" title="점수초기화" style="cursor:hand;" align="absmiddle"onclick="javascript:fn_ALLDEL('7','<%=i%>');"></td>
				</tr>
			    <tr height="20">
					<td bgcolor="#EEF6FF" class="TDCont" align='center' width=70>만족<br>(9)</td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center' width=70>보통<br>(8)</td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center' width=70>불만족<br>(7)</td>
				</tr>
<%

			SQL = "SELECT * FROM TB_CODE where CODEGROUP = '" & sGubun &"' AND USEYN = 'Y' ORDER BY CODE"
			SET Rs = DB.execute(SQL)			

			i = 0
			do until rs.eof
				i = i + 1
				if ( i mod 2 ) = 1 then
					sBgColor = "#ffffff"
				else
					sBgColor = "#FFFCE7"				
				end if

				'값 불러오기
				SQL1 = "select *"
				SQL1 = SQL1 & " from armyinformix.dbo.monitor where factnum = '" & receiptfactnum & "' and factpeoplenum='"& db_factPeoplenum_7 & "' and  seqno = " & i
				
				SET Rs1 = DB.execute(SQL1)	
				if Rs1.eof = false then
					point9 = Rs1("point9")
					point8 = Rs1("point8")
					point7 = Rs1("point7")
					pointplus = Rs1("pointplus")
					totpoint = Rs1("totpoint")
				else
					point9 = ""
					point8 = ""
					point7 = ""
					pointplus = ""
					totpoint = ""
				end if
%>				
			    <tr>
					<td bgcolor="<%=sBgColor%>" width=530  class="TDCont"  colspan='2'>&nbsp;<%=rs("codename")%></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="radio" name="QUESTION_7<%=i%>" value="9" class="none" onClick="fn_YES('7','<%=i%>','9');" <%if point9 = "1" then%>checked<%end if%> ></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="radio" name="QUESTION_7<%=i%>" value="8" class="none" onClick="fn_YES('7','<%=i%>','8');" <%if point8 = "1" then%>checked<%end if%> ></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="radio" name="QUESTION_7<%=i%>" value="7" class="none" onClick="fn_YES('7','<%=i%>','7');" <%if point7 = "1" then%>checked<%end if%> ></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="checkbox" name="QUESTIONP_7<%=i%>" value="1" class="none" onClick="fn_YES('7','<%=i%>','1');"<% if pointplus="1" then Response.Write("checked") end if %>>					
					</td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="text" name="POINT_7<%=i%>" value="<%=totpoint%>" size="2" maxlength="2" style="border-width:0px ; border-color:#cccccc ; border-style:solid; text-align:right; background-color:<%=sBgColor%>" readonly></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><img src="/Images/Btn/BtnIconDel.gif" title="삭제" style="cursor:hand;" align="absmiddle"onclick="javascript:fn_DEL('7','<%=i%>');"></td>

				</tr>
<%
				rs.movenext
			loop
%>
			    <tr ><td bgcolor="#ffffff" class="TDCont"  colspan='5' valign="top" height="100">
					<table width="100%" height="100%" border="0" cellspacing="1" cellpadding="0" align="center" bgcolor="#CCCCCC">
						<tr height="30">
							<td bgcolor="#FFEEF9" class="TDCont" width=100 align='center'>설문결과</td>
							<td bgcolor="#FFFFFF" width="120">
								<%
									'======= 처리구분 코드 가져오기 ==================================================
									SqlCode = "SELECT CODE, CODENAME FROM TB_CODE"
									SqlCode = SqlCode& " WHERE USEYN='Y' AND SYSYN='Y' AND CODEGROUP='B10'"
									SqlCode = SqlCode& " ORDER BY CODE"
									set RsCode = db.execute(SqlCode)
								%>
								&nbsp;<select name="MONITORRESULT_7" size="1" class="ComboFFFCE7" onChange="fn_ResultSet('7');">
									<option value="">설문결과선택</option>
									<%
										IF NOT(RsCode.Eof OR RsCode.bof) THEN
											DO until RsCode.EOF
												CODE = RsCode("CODE")
												CODENAME = RsCode("CODENAME")
									%>
									<%=printSelect("" &CODENAME& "","" &CODE& "","" &db_MONITOR_RESULT_7& "")%>
									<%
											RsCode.MoveNext
											LOOP
										END IF
										RsCode.Close
										set RsCode = NOTHING
									%>
								</select>
							
							</td>


							<td bgcolor="#FFEEF9" class="TDCont" align='center' width="100">상담예약일시</td>
							<td bgcolor="#FFFFFF">&nbsp;<input value="<%=RESERVEDATE_7%>" name="RESERVEDATE_7" readonly type="text" size="10" onfocus="setFocusColor(this);" onClick="new CalendarFrame.Calendar(this);" onblur="setOutColor(this);" onchange="fn_settime('7')">&nbsp;<img src="/Images/icon_sche.gif" title="달력" style="cursor:hand;" align="absmiddle" name="RESERVE_CAR_7" onClick="window.open('/Include/Calendar_view.asp?goMonth='+document.ListForm.RESERVEDATE_7.value+'&firstcode=&FlightCode=&StartEndGu=S&ControlName=opener.document.ListForm.RESERVEDATE_7','','width=800,height=300,top=200,left=300,scrollbars=no,status=no');" onblur="setOutColor(this);" >&nbsp;<img src="/Images/Btn/BtnIconDel.gif" title="삭제" name="RESERVE_DEL_7" style="cursor:hand;" align="absmiddle"onclick="document.all.RESERVEDATE_7.value='';">&nbsp;&nbsp;<input value="<%=RESERVEHOUR_7%>" name="RESERVEHOUR_7" type="text" size="2" maxlength="2" onfocus="setFocusColor(this);" >시&nbsp;<input value="<%=RESERVEMIN_7%>" name="RESERVEMIN_7" type="text" size="2" maxlength="2" onfocus="setFocusColor(this);">분&nbsp;&nbsp;&nbsp;&nbsp;&nbsp<select name="RESERVETIME_7" size="1" class="ComboFFFCE7" onchange="fn_settime('7')">
									<Option value ='' selected>시간선택</option>
									<Option value ='1' >10분후</option>
									<Option value ='2' >30분후</option>
									<Option value ='3' >1시간후</option>
									<Option value ='4' >2시간후</option>
									<Option value ='08' >오전 7시</option>
									<Option value ='08' >오전 8시</option>
									<Option value ='09' >오전 9시</option>
									<Option value ='10' >오전10시</option>
									<Option value ='11' >오전11시</option>
									<Option value ='12' >오후12시</option>
									<Option value ='13' >오후13시</option>
									<Option value ='14' >오후14시</option>
									<Option value ='15' >오후15시</option>
									<Option value ='16' >오후16시</option>
									<Option value ='17' >오후17시</option>
									<Option value ='18' >오후18시</option>
									<Option value ='19' >오후19시</option>
									<Option value ='20' >오후20시</option>
									<Option value ='21' >오후21시</option>
									<Option value ='22' >오후22시</option>

								</select>
							</td>
						</tr>
						<tr height="40">
							<td bgcolor="#FFEEF9" class="TDCont" align='center' width=100>종합평가</td>
							<td bgcolor="#FFFFFF" colspan=6>&nbsp;<textarea name="Remark_7" style="width:99%; height:100%" wrap="soft" class="TextareaInput"><%=db_Remark_7%></textarea>
							</td>	
						</tr>
						<tr height="40">
							<td bgcolor="#FFEEF9" class="TDCont" align='center' width=100>비고</td>
							<td bgcolor="#FFFFFF" colspan=6>&nbsp;<textarea name="Remark1_7" style="width:99%; height:100%" wrap="soft" class="TextareaInput"><%=db_Remark1_7%></textarea>
							</td>	
						</tr>
					</table>
					</td>
					<td bgcolor="#EEF6FF" width=70 class="TDCont" align='center'>점수(평균):</td>
					<td bgcolor="#FFFFFF" width=140 class="TDCont" align='center' colspan=2><input type="text" name="TOT_7" value="<%=db_TOT_7%>" size="5" maxlength="5" style="border-width:0px ; border-color:#EEF6FF ; border-style:solid; text-align:right; font-color:#ff0000;font-size:15px;font-weight:bold" readonly ></td>
				</tr>
			</table>

<%if db_factPeoplenum_8 = "" then%>
			<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table width="920" height="30" border="0" cellspacing="0" cellpadding="0" align="center">
				<tr>
					<td align="left"><img src="/Images/Btn/BtnRegiAdd_P<%=right(sGubun,1)%>.GIF" style="cursor:hand;" align="absmiddle" onClick="fn_AddForm('ON','divFORM8','FRM8');">&nbsp;<img src="/Images/Btn/BtnRegiDel_P<%=right(sGubun,1)%>.GIF" style="cursor:hand;" align="absmiddle" onClick="fn_AddForm('OFF','divFORM7','FRM7');"></td>
					<td align='left' width="50%"><input type="checkbox" name="RecYN_7" value="Y" class="none"><font color='#0000ff'><b>녹취첨부</b></font></td>
				</tr>
			</table>
<%else%>
			<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table width="920" height="30" border="0" cellspacing="0" cellpadding="0" align="center">
				<tr>
					<td align="left"><img src="/Images/Btn/BtnRegiDel_P<%=right(sGubun,1)%>.GIF" style="cursor:hand;" align="absmiddle" onClick="fn_AddForm('OFF','divFORM7','FRM7');"></td>
					<td align='left' width="50%"><input type="checkbox" name="RecYN_7" value="Y" class="none"><font color='#0000ff'><b>녹취첨부</b></font></td>
				</tr>
			</table>
<%end if%>


		</td>
	</tr>
</table>
</span>


<%'====== 상담접수 폼 #8 시작 =======================================================================================%>
<%if db_factPeoplenum_8 = "" then%>
	<span id="divFORM8" style="display:none;">
<% else %>
	<span id="divFORM8" style="display:block;">
<%end if%>
<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
<table border="0" width="920" cellpadding="0" cellspacing="1" bgcolor="#EFECE5" align="center">
	<tr bgcolor="#FFFFFF">
		<td>
			<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table width="920" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#ffffff">
			    <tr>
					<td align="left" bgcolor="#FFFFFF" class="TDCont">&nbsp;<img src="/Images/dot_01.gif" >&nbsp;<b><font color="#ff00ff">(<%=sGubunName%>)</font> 설문지#8</b></td>
				</tr>
			</table>
			<table width="920" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">
			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>관 계</td>
					<td bgcolor="#FFFFFF">						<%
							'======= 처리구분 코드 가져오기 ==================================================
							SqlCode = "SELECT CODE, CODENAME FROM TB_CODE"
							SqlCode = SqlCode& " WHERE USEYN='Y' AND SYSYN='Y' AND CODEGROUP='B01'"
							SqlCode = SqlCode& " ORDER BY CODE"
							set RsCode = db.execute(SqlCode)
						%>
						&nbsp;<select name="SECTION2_8" size="1" class="ComboFFFCE7" onChange="fn_UpdateData('8','SECTION2_');">
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("CODE")
										CODENAME = RsCode("CODENAME")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &sGubun& "")%>
							<%
									RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
							%>
						</select>
					&nbsp;&nbsp;&nbsp;&nbsp;<font color="#0000ff"><%=db_factPeoplenum_8%></font></td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>계급(신분)</td>
					<td bgcolor="#FFFFFF"><%
							'======= 처리구분 코드 가져오기 ==================================================
							SqlCode = "SELECT *	FROM armyinformix.dbo.data"
							SqlCode = SqlCode& " where [restrict] = '104' order by [restrict]"
							set RsCode = db.execute(SqlCode)
						%>
						&nbsp;<select name="LEVEL_8" size="1" class="ComboFFFCE7">
							<option value="">계급선택</option>
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("number")
										CODENAME = RsCode("contents")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &db_level_8& "")%>
							<%
									RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
							%>
						</select></td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>성  명</td>
					<td bgcolor="#FFFFFF">&nbsp;<input type="text" name="NAME_8" value="<%=db_NAME_8%>" size="20" maxlength="20" style="border-width:1px ; border-color:#cccccc ; border-style:solid" ></td>
				</tr>
			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>연락처 1</td>
					<td bgcolor="#FFFFFF">&nbsp;<input type="text" name="HOMEPHONE_8" value="<%=db_HOMEPHONE_8%>" size="13" maxlength="13" style="border-width:1px ; border-color:#cccccc ; border-style:solid">
					<img src="/Images/Cti/icon_tel.gif" style="cursor:hand;" onClick="fn_dial('1','8');" align="absmiddle" title="군전화로 전화걸기">&nbsp<img src="/Images/Cti/icon_tel_1.gif" style="cursor:hand;" onClick="fn_dial_1('1','8');" align="absmiddle" title="일반전화로 전화걸기">&nbsp<img src="/Images/Cti/icon_sche.gif" style="cursor:hand;" onClick="fn_sms('1','8');" align="absmiddle" title="문자전송"></td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>연락처 2</td>
					<td bgcolor="#FFFFFF">&nbsp;<input type="text" name="MOBILEPHONE_8" value="<%=db_MOBILEPHONE_8%>" size="13" maxlength="13" style="border-width:1px ; border-color:#cccccc ; border-style:solid">&nbsp;<img src="/Images/Cti/icon_tel.gif" style="cursor:hand;" onClick="fn_dial('2','8');" align="absmiddle" title="군전화로 전화걸기">&nbsp<img src="/Images/Cti/icon_tel_1.gif" style="cursor:hand;" onClick="fn_dial_1('2','8');" align="absmiddle" title="일반전화로 전화걸기">&nbsp<img src="/Images/Cti/icon_sche.gif" style="cursor:hand;" onClick="fn_sms('2','8');" align="absmiddle" title="문자전송"></td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>연락처비고</td>
					<td bgcolor="#FFFFFF">&nbsp;<input type="text" name="ETCPHONE_8" value="<%=db_ETCPHONE_8%>" size="20" maxlength="20" style="border-width:1px ; border-color:#cccccc ; border-style:solid">
					</td>
						
				</tr>
			    <tr>

					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>설문일시</td>
					<td bgcolor="#FFFFFF" width=200>&nbsp;<input type="text" name="MONITORDATE_8" value="<%=db_MONITORDATE_8%>" size="10" maxlength="10" style="border-width:1px ; border-color:#cccccc ; border-style:solid" >&nbsp;<img src="/Images/icon_sche.gif" title="달력" style="cursor:hand;" align="absmiddle" name="Date2_CAR" onClick="window.open('/Include/Calendar_view.asp?goMonth='+document.all.MONITORDATE_8.value+'&firstcode=&FlightCode=&StartEndGu=S&ControlName=opener.document.all.MONITORDATE_8','','width=800,height=300,top=200,left=300,scrollbars=no,status=no');" onblur="setOutColor(this);">
						&nbsp;<img src="/Images/Btn/BtnIconDel.gif" title="삭제" style="cursor:hand;" align="absmiddle"onclick="document.all.MONITORDATE_8.value='';">
					</td>

					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>설문결과</td>
					<td bgcolor="#FFFFFF" width=200>&nbsp;<font color="#0000ff"><%=db_getCodeName("B10",db_MONITOR_RESULT_8)%></font>
					</td>

					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>예약일시</td>
					<td bgcolor="#FFFFFF" width=200>&nbsp;<%=db_RESERVEDATE_8%>
					</td>
				</tr>
				<tr>
					<td bgcolor="#FFEEF9" class="TDCont" align='center' width=100>통화내역</td>
					<td bgcolor="#FFFFFF" colspan=6>&nbsp;<span id="HISTORY_8"><%=db_History_8%></span>
					</td>	
				</tr>

			</table>
			<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table width="920" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">
			    <tr height="25">
					<td bgcolor="#EEF6FF" width=530 rowspan=2 class="TDCont"  colspan='2' align='center'>질문사항</td>
					<td bgcolor="#EEF6FF" colspan='3' class="TDCont" align='center' width=210>답변결과</td>
					<td bgcolor="#EEF6FF" rowspan=2 class="TDCont" align='center' width=70>가점<br>(+1)</td>
					<td bgcolor="#EEF6FF" rowspan=2 class="TDCont" align='center' width=70>점수</td>
					<td bgcolor="#EEF6FF" rowspan=2 class="TDCont" align='center' width=70>초기화<br><img src="/Images/Btn/BtnIconDel.gif" title="점수초기화" style="cursor:hand;" align="absmiddle"onclick="javascript:fn_ALLDEL('8','<%=i%>');"></td>
				</tr>
			    <tr height="20">
					<td bgcolor="#EEF6FF" class="TDCont" align='center' width=70>만족<br>(9)</td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center' width=70>보통<br>(8)</td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center' width=70>불만족<br>(7)</td>
				</tr>
<%

			SQL = "SELECT * FROM TB_CODE where CODEGROUP = '" & sGubun &"' AND USEYN = 'Y' ORDER BY CODE"
			SET Rs = DB.execute(SQL)			

			i = 0
			do until rs.eof
				i = i + 1
				if ( i mod 2 ) = 1 then
					sBgColor = "#ffffff"
				else
					sBgColor = "#FFFCE7"				
				end if

				'값 불러오기
				SQL1 = "select *"
				SQL1 = SQL1 & " from armyinformix.dbo.monitor where factnum = '" & receiptfactnum & "' and factpeoplenum='"& db_factPeoplenum_8 & "' and  seqno = " & i
				
				SET Rs1 = DB.execute(SQL1)	
				if Rs1.eof = false then
					point9 = Rs1("point9")
					point8 = Rs1("point8")
					point7 = Rs1("point7")
					pointplus = Rs1("pointplus")
					totpoint = Rs1("totpoint")
				else
					point9 = ""
					point8 = ""
					point7 = ""
					pointplus = ""
					totpoint = ""
				end if
%>				
			    <tr>
					<td bgcolor="<%=sBgColor%>" width=530  class="TDCont"  colspan='2'>&nbsp;<%=rs("codename")%></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="radio" name="QUESTION_8<%=i%>" value="9" class="none" onClick="fn_YES('8','<%=i%>','9');" <%if point9 = "1" then%>checked<%end if%> ></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="radio" name="QUESTION_8<%=i%>" value="8" class="none" onClick="fn_YES('8','<%=i%>','8');" <%if point8 = "1" then%>checked<%end if%> ></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="radio" name="QUESTION_8<%=i%>" value="7" class="none" onClick="fn_YES('8','<%=i%>','7');" <%if point7 = "1" then%>checked<%end if%> ></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="checkbox" name="QUESTIONP_8<%=i%>" value="1" class="none" onClick="fn_YES('8','<%=i%>','1');"<% if pointplus="1" then Response.Write("checked") end if %>>					
					</td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="text" name="POINT_8<%=i%>" value="<%=totpoint%>" size="2" maxlength="2" style="border-width:0px ; border-color:#cccccc ; border-style:solid; text-align:right; background-color:<%=sBgColor%>" readonly></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><img src="/Images/Btn/BtnIconDel.gif" title="삭제" style="cursor:hand;" align="absmiddle"onclick="javascript:fn_DEL('8','<%=i%>');"></td>

				</tr>
<%
				rs.movenext
			loop
%>
			    <tr ><td bgcolor="#ffffff" class="TDCont"  colspan='5' valign="top" height="100">
					<table width="100%" height="100%" border="0" cellspacing="1" cellpadding="0" align="center" bgcolor="#CCCCCC">
						<tr height="30">
							<td bgcolor="#FFEEF9" class="TDCont" width=100 align='center'>설문결과</td>
							<td bgcolor="#FFFFFF" width="120">
								<%
									'======= 처리구분 코드 가져오기 ==================================================
									SqlCode = "SELECT CODE, CODENAME FROM TB_CODE"
									SqlCode = SqlCode& " WHERE USEYN='Y' AND SYSYN='Y' AND CODEGROUP='B10'"
									SqlCode = SqlCode& " ORDER BY CODE"
									set RsCode = db.execute(SqlCode)
								%>
								&nbsp;<select name="MONITORRESULT_8" size="1" class="ComboFFFCE7" onChange="fn_ResultSet('8');">
									<option value="">설문결과선택</option>
									<%
										IF NOT(RsCode.Eof OR RsCode.bof) THEN
											DO until RsCode.EOF
												CODE = RsCode("CODE")
												CODENAME = RsCode("CODENAME")
									%>
									<%=printSelect("" &CODENAME& "","" &CODE& "","" &db_MONITOR_RESULT_8& "")%>
									<%
											RsCode.MoveNext
											LOOP
										END IF
										RsCode.Close
										set RsCode = NOTHING
									%>
								</select>
							
							</td>


							<td bgcolor="#FFEEF9" class="TDCont" align='center' width="100">상담예약일시</td>
							<td bgcolor="#FFFFFF">&nbsp;<input value="<%=RESERVEDATE_8%>" name="RESERVEDATE_8" readonly type="text" size="10" onfocus="setFocusColor(this);" onClick="new CalendarFrame.Calendar(this);" onblur="setOutColor(this);" onchange="fn_settime('8')">&nbsp;<img src="/Images/icon_sche.gif" title="달력" style="cursor:hand;" align="absmiddle" name="RESERVE_CAR_8" onClick="window.open('/Include/Calendar_view.asp?goMonth='+document.ListForm.RESERVEDATE_8.value+'&firstcode=&FlightCode=&StartEndGu=S&ControlName=opener.document.ListForm.RESERVEDATE_8','','width=800,height=300,top=200,left=300,scrollbars=no,status=no');" onblur="setOutColor(this);" >&nbsp;<img src="/Images/Btn/BtnIconDel.gif" title="삭제" name="RESERVE_DEL_8" style="cursor:hand;" align="absmiddle"onclick="document.all.RESERVEDATE_8.value='';">&nbsp;&nbsp;<input value="<%=RESERVEHOUR_8%>" name="RESERVEHOUR_8" type="text" size="2" maxlength="2" onfocus="setFocusColor(this);" >시&nbsp;<input value="<%=RESERVEMIN_8%>" name="RESERVEMIN_8" type="text" size="2" maxlength="2" onfocus="setFocusColor(this);">분&nbsp;&nbsp;&nbsp;&nbsp;&nbsp<select name="RESERVETIME_8" size="1" class="ComboFFFCE7" onchange="fn_settime('8')">
									<Option value ='' selected>시간선택</option>
									<Option value ='1' >10분후</option>
									<Option value ='2' >30분후</option>
									<Option value ='3' >1시간후</option>
									<Option value ='4' >2시간후</option>
									<Option value ='08' >오전 7시</option>
									<Option value ='08' >오전 8시</option>
									<Option value ='09' >오전 9시</option>
									<Option value ='10' >오전10시</option>
									<Option value ='11' >오전11시</option>
									<Option value ='12' >오후12시</option>
									<Option value ='13' >오후13시</option>
									<Option value ='14' >오후14시</option>
									<Option value ='15' >오후15시</option>
									<Option value ='16' >오후16시</option>
									<Option value ='17' >오후17시</option>
									<Option value ='18' >오후18시</option>
									<Option value ='19' >오후19시</option>
									<Option value ='20' >오후20시</option>
									<Option value ='21' >오후21시</option>
									<Option value ='22' >오후22시</option>

								</select>
							</td>
						</tr>
						<tr height="40">
							<td bgcolor="#FFEEF9" class="TDCont" align='center' width=100>종합평가</td>
							<td bgcolor="#FFFFFF" colspan=6>&nbsp;<textarea name="Remark_8" style="width:99%; height:100%" wrap="soft" class="TextareaInput"><%=db_Remark_8%></textarea>
							</td>	
						</tr>
						<tr height="40">
							<td bgcolor="#FFEEF9" class="TDCont" align='center' width=100>비고</td>
							<td bgcolor="#FFFFFF" colspan=6>&nbsp;<textarea name="Remark1_8" style="width:99%; height:100%" wrap="soft" class="TextareaInput"><%=db_Remark1_8%></textarea>
							</td>	
						</tr>
					</table>
					</td>
					<td bgcolor="#EEF6FF" width=70 class="TDCont" align='center'>점수(평균):</td>
					<td bgcolor="#FFFFFF" width=140 class="TDCont" align='center' colspan=2><input type="text" name="TOT_8" value="<%=db_TOT_8%>" size="5" maxlength="5" style="border-width:0px ; border-color:#EEF6FF ; border-style:solid; text-align:right; font-color:#ff0000;font-size:15px;font-weight:bold" readonly ></td>
				</tr>
			</table>

<%if db_factPeoplenum_9 = "" then%>
			<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table width="920" height="30" border="0" cellspacing="0" cellpadding="0" align="center">
				<tr>
					<td align="left"><img src="/Images/Btn/BtnRegiAdd_P<%=right(sGubun,1)%>.GIF" style="cursor:hand;" align="absmiddle" onClick="fn_AddForm('ON','divFORM9','FRM9');">&nbsp;<img src="/Images/Btn/BtnRegiDel_P<%=right(sGubun,1)%>.GIF" style="cursor:hand;" align="absmiddle" onClick="fn_AddForm('OFF','divFORM8','FRM8');"></td>
					<td align='left' width="50%"><input type="checkbox" name="RecYN_8" value="Y" class="none"><font color='#0000ff'><b>녹취첨부</b></font></td>
				</tr>
			</table>
<%else%>
			<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table width="920" height="30" border="0" cellspacing="0" cellpadding="0" align="center">
				<tr>
					<td align="left"><img src="/Images/Btn/BtnRegiDel_P<%=right(sGubun,1)%>.GIF" style="cursor:hand;" align="absmiddle" onClick="fn_AddForm('OFF','divFORM8','FRM8');"></td>
					<td align='left' width="50%"><input type="checkbox" name="RecYN_8" value="Y" class="none"><font color='#0000ff'><b>녹취첨부</b></font></td>
				</tr>
			</table>
<%end if%>

		</td>
	</tr>
</table>
</span>



<%'====== 상담접수 폼 #9 시작 =======================================================================================%>
<%if db_factPeoplenum_9 = "" then%>
	<span id="divFORM9" style="display:none;">
<% else %>
	<span id="divFORM9" style="display:block;">
<%end if%>
<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
<table border="0" width="920" cellpadding="0" cellspacing="1" bgcolor="#EFECE5" align="center">
	<tr bgcolor="#FFFFFF">
		<td>
			<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table width="920" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#ffffff">
			    <tr>
					<td align="left" bgcolor="#FFFFFF" class="TDCont">&nbsp;<img src="/Images/dot_01.gif" >&nbsp;<b><font color="#ff00ff">(<%=sGubunName%>)</font> 설문지#9</b></td>
				</tr>
			</table>
			<table width="920" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">
			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>관 계</td>
					<td bgcolor="#FFFFFF">						<%
							'======= 처리구분 코드 가져오기 ==================================================
							SqlCode = "SELECT CODE, CODENAME FROM TB_CODE"
							SqlCode = SqlCode& " WHERE USEYN='Y' AND SYSYN='Y' AND CODEGROUP='B01'"
							SqlCode = SqlCode& " ORDER BY CODE"
							set RsCode = db.execute(SqlCode)
						%>
						&nbsp;<select name="SECTION2_9" size="1" class="ComboFFFCE7" onChange="fn_UpdateData('9','SECTION2_');">
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("CODE")
										CODENAME = RsCode("CODENAME")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &sGubun& "")%>
							<%
									RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
							%>
						</select>
					&nbsp;&nbsp;&nbsp;&nbsp;<font color="#0000ff"><%=db_factPeoplenum_9%></font></td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>계급(신분)</td>
					<td bgcolor="#FFFFFF"><%
							'======= 처리구분 코드 가져오기 ==================================================
							SqlCode = "SELECT *	FROM armyinformix.dbo.data"
							SqlCode = SqlCode& " where [restrict] = '104' order by [restrict]"
							set RsCode = db.execute(SqlCode)
						%>
						&nbsp;<select name="LEVEL_9" size="1" class="ComboFFFCE7">
							<option value="">계급선택</option>
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("number")
										CODENAME = RsCode("contents")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &db_level_9& "")%>
							<%
									RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
							%>
						</select></td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>성  명</td>
					<td bgcolor="#FFFFFF">&nbsp;<input type="text" name="NAME_9" value="<%=db_NAME_9%>" size="20" maxlength="20" style="border-width:1px ; border-color:#cccccc ; border-style:solid" ></td>
				</tr>
			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>연락처 1</td>
					<td bgcolor="#FFFFFF">&nbsp;<input type="text" name="HOMEPHONE_9" value="<%=db_HOMEPHONE_9%>" size="13" maxlength="13" style="border-width:1px ; border-color:#cccccc ; border-style:solid">
					<img src="/Images/Cti/icon_tel.gif" style="cursor:hand;" onClick="fn_dial('1','9');" align="absmiddle" title="군전화로 전화걸기">&nbsp<img src="/Images/Cti/icon_tel_1.gif" style="cursor:hand;" onClick="fn_dial_1('1','9');" align="absmiddle" title="일반전화로 전화걸기">&nbsp<img src="/Images/Cti/icon_sche.gif" style="cursor:hand;" onClick="fn_sms('1','9');" align="absmiddle" title="문자전송"></td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>연락처 2</td>
					<td bgcolor="#FFFFFF">&nbsp;<input type="text" name="MOBILEPHONE_9" value="<%=db_MOBILEPHONE_9%>" size="13" maxlength="13" style="border-width:1px ; border-color:#cccccc ; border-style:solid">&nbsp;<img src="/Images/Cti/icon_tel.gif" style="cursor:hand;" onClick="fn_dial('2','9');" align="absmiddle" title="군전화로 전화걸기">&nbsp<img src="/Images/Cti/icon_tel_1.gif" style="cursor:hand;" onClick="fn_dial_1('2','9');" align="absmiddle" title="일반전화로 전화걸기">&nbsp<img src="/Images/Cti/icon_sche.gif" style="cursor:hand;" onClick="fn_sms('2','9');" align="absmiddle" title="문자전송"></td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>연락처비고</td>
					<td bgcolor="#FFFFFF">&nbsp;<input type="text" name="ETCPHONE_9" value="<%=db_ETCPHONE_9%>" size="20" maxlength="20" style="border-width:1px ; border-color:#cccccc ; border-style:solid">
					</td>
						
				</tr>
			    <tr>

					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>설문일시</td>
					<td bgcolor="#FFFFFF" width=200>&nbsp;<input type="text" name="MONITORDATE_9" value="<%=db_MONITORDATE_9%>" size="10" maxlength="10" style="border-width:1px ; border-color:#cccccc ; border-style:solid" >&nbsp;<img src="/Images/icon_sche.gif" title="달력" style="cursor:hand;" align="absmiddle" name="Date2_CAR" onClick="window.open('/Include/Calendar_view.asp?goMonth='+document.all.MONITORDATE_9.value+'&firstcode=&FlightCode=&StartEndGu=S&ControlName=opener.document.all.MONITORDATE_9','','width=800,height=300,top=200,left=300,scrollbars=no,status=no');" onblur="setOutColor(this);">
						&nbsp;<img src="/Images/Btn/BtnIconDel.gif" title="삭제" style="cursor:hand;" align="absmiddle"onclick="document.all.MONITORDATE_9.value='';">
					</td>

					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>설문결과</td>
					<td bgcolor="#FFFFFF" width=200>&nbsp;<font color="#0000ff"><%=db_getCodeName("B10",db_MONITOR_RESULT_9)%></font>
					</td>

					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>예약일시</td>
					<td bgcolor="#FFFFFF" width=200>&nbsp;<%=db_RESERVEDATE_9%>
					</td>
				</tr>
				<tr>
					<td bgcolor="#FFEEF9" class="TDCont" align='center' width=100>통화내역</td>
					<td bgcolor="#FFFFFF" colspan=6>&nbsp;<span id="HISTORY_9"><%=db_History_9%></span>
					</td>	
				</tr>

			</table>
			<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table width="920" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">
			    <tr height="25">
					<td bgcolor="#EEF6FF" width=530 rowspan=2 class="TDCont"  colspan='2' align='center'>질문사항</td>
					<td bgcolor="#EEF6FF" colspan='3' class="TDCont" align='center' width=210>답변결과</td>
					<td bgcolor="#EEF6FF" rowspan=2 class="TDCont" align='center' width=70>가점<br>(+1)</td>
					<td bgcolor="#EEF6FF" rowspan=2 class="TDCont" align='center' width=70>점수</td>
					<td bgcolor="#EEF6FF" rowspan=2 class="TDCont" align='center' width=70>초기화<br><img src="/Images/Btn/BtnIconDel.gif" title="점수초기화" style="cursor:hand;" align="absmiddle"onclick="javascript:fn_ALLDEL('9','<%=i%>');"></td>
				</tr>
			    <tr height="20">
					<td bgcolor="#EEF6FF" class="TDCont" align='center' width=70>만족<br>(9)</td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center' width=70>보통<br>(8)</td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center' width=70>불만족<br>(7)</td>
				</tr>
<%

			SQL = "SELECT * FROM TB_CODE where CODEGROUP = '" & sGubun &"' AND USEYN = 'Y' ORDER BY CODE"
			SET Rs = DB.execute(SQL)			

			i = 0
			do until rs.eof
				i = i + 1
				if ( i mod 2 ) = 1 then
					sBgColor = "#ffffff"
				else
					sBgColor = "#FFFCE7"				
				end if

				'값 불러오기
				SQL1 = "select *"
				SQL1 = SQL1 & " from armyinformix.dbo.monitor where factnum = '" & receiptfactnum & "' and factpeoplenum='"& db_factPeoplenum_9 & "' and  seqno = " & i
				
				SET Rs1 = DB.execute(SQL1)	
				if Rs1.eof = false then
					point9 = Rs1("point9")
					point8 = Rs1("point8")
					point7 = Rs1("point7")
					pointplus = Rs1("pointplus")
					totpoint = Rs1("totpoint")
				else
					point9 = ""
					point8 = ""
					point7 = ""
					pointplus = ""
					totpoint = ""
				end if
%>				
			    <tr>
					<td bgcolor="<%=sBgColor%>" width=530  class="TDCont"  colspan='2'>&nbsp;<%=rs("codename")%></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="radio" name="QUESTION_9<%=i%>" value="9" class="none" onClick="fn_YES('9','<%=i%>','9');" <%if point9 = "1" then%>checked<%end if%> ></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="radio" name="QUESTION_9<%=i%>" value="8" class="none" onClick="fn_YES('9','<%=i%>','8');" <%if point8 = "1" then%>checked<%end if%> ></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="radio" name="QUESTION_9<%=i%>" value="7" class="none" onClick="fn_YES('9','<%=i%>','7');" <%if point7 = "1" then%>checked<%end if%> ></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="checkbox" name="QUESTIONP_9<%=i%>" value="1" class="none" onClick="fn_YES('9','<%=i%>','1');"<% if pointplus="1" then Response.Write("checked") end if %>>					
					</td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="text" name="POINT_9<%=i%>" value="<%=totpoint%>" size="2" maxlength="2" style="border-width:0px ; border-color:#cccccc ; border-style:solid; text-align:right; background-color:<%=sBgColor%>" readonly></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><img src="/Images/Btn/BtnIconDel.gif" title="삭제" style="cursor:hand;" align="absmiddle"onclick="javascript:fn_DEL('9','<%=i%>');"></td>

				</tr>
<%
				rs.movenext
			loop
%>
			    <tr ><td bgcolor="#ffffff" class="TDCont"  colspan='5' valign="top" height="100">
					<table width="100%" height="100%" border="0" cellspacing="1" cellpadding="0" align="center" bgcolor="#CCCCCC">
						<tr height="30">
							<td bgcolor="#FFEEF9" class="TDCont" width=100 align='center'>설문결과</td>
							<td bgcolor="#FFFFFF" width="120">
								<%
									'======= 처리구분 코드 가져오기 ==================================================
									SqlCode = "SELECT CODE, CODENAME FROM TB_CODE"
									SqlCode = SqlCode& " WHERE USEYN='Y' AND SYSYN='Y' AND CODEGROUP='B10'"
									SqlCode = SqlCode& " ORDER BY CODE"
									set RsCode = db.execute(SqlCode)
								%>
								&nbsp;<select name="MONITORRESULT_9" size="1" class="ComboFFFCE7" onChange="fn_ResultSet('9');">
									<option value="">설문결과선택</option>
									<%
										IF NOT(RsCode.Eof OR RsCode.bof) THEN
											DO until RsCode.EOF
												CODE = RsCode("CODE")
												CODENAME = RsCode("CODENAME")
									%>
									<%=printSelect("" &CODENAME& "","" &CODE& "","" &db_MONITOR_RESULT_9& "")%>
									<%
											RsCode.MoveNext
											LOOP
										END IF
										RsCode.Close
										set RsCode = NOTHING
									%>
								</select>
							
							</td>


							<td bgcolor="#FFEEF9" class="TDCont" align='center' width="100">상담예약일시</td>
							<td bgcolor="#FFFFFF">&nbsp;<input value="<%=RESERVEDATE_9%>" name="RESERVEDATE_9" readonly type="text" size="10" onfocus="setFocusColor(this);" onClick="new CalendarFrame.Calendar(this);" onblur="setOutColor(this);" onchange="fn_settime('9')">&nbsp;<img src="/Images/icon_sche.gif" title="달력" style="cursor:hand;" align="absmiddle" name="RESERVE_CAR_9" onClick="window.open('/Include/Calendar_view.asp?goMonth='+document.ListForm.RESERVEDATE_9.value+'&firstcode=&FlightCode=&StartEndGu=S&ControlName=opener.document.ListForm.RESERVEDATE_9','','width=800,height=300,top=200,left=300,scrollbars=no,status=no');" onblur="setOutColor(this);" >&nbsp;<img src="/Images/Btn/BtnIconDel.gif" title="삭제" name="RESERVE_DEL_9" style="cursor:hand;" align="absmiddle"onclick="document.all.RESERVEDATE_9.value='';">&nbsp;&nbsp;<input value="<%=RESERVEHOUR_9%>" name="RESERVEHOUR_9" type="text" size="2" maxlength="2" onfocus="setFocusColor(this);" >시&nbsp;<input value="<%=RESERVEMIN_9%>" name="RESERVEMIN_9" type="text" size="2" maxlength="2" onfocus="setFocusColor(this);">분&nbsp;&nbsp;&nbsp;&nbsp;&nbsp<select name="RESERVETIME_9" size="1" class="ComboFFFCE7" onchange="fn_settime('9')">
									<Option value ='' selected>시간선택</option>
									<Option value ='1' >10분후</option>
									<Option value ='2' >30분후</option>
									<Option value ='3' >1시간후</option>
									<Option value ='4' >2시간후</option>
									<Option value ='08' >오전 7시</option>
									<Option value ='08' >오전 8시</option>
									<Option value ='09' >오전 9시</option>
									<Option value ='10' >오전10시</option>
									<Option value ='11' >오전11시</option>
									<Option value ='12' >오후12시</option>
									<Option value ='13' >오후13시</option>
									<Option value ='14' >오후14시</option>
									<Option value ='15' >오후15시</option>
									<Option value ='16' >오후16시</option>
									<Option value ='17' >오후17시</option>
									<Option value ='18' >오후18시</option>
									<Option value ='19' >오후19시</option>
									<Option value ='20' >오후20시</option>
									<Option value ='21' >오후21시</option>
									<Option value ='22' >오후22시</option>

								</select>
							</td>
						</tr>
						<tr height="40">
							<td bgcolor="#FFEEF9" class="TDCont" align='center' width=100>종합평가</td>
							<td bgcolor="#FFFFFF" colspan=6>&nbsp;<textarea name="Remark_9" style="width:99%; height:100%" wrap="soft" class="TextareaInput"><%=db_Remark_9%></textarea>
							</td>	
						</tr>
						<tr height="40">
							<td bgcolor="#FFEEF9" class="TDCont" align='center' width=100>비고</td>
							<td bgcolor="#FFFFFF" colspan=6>&nbsp;<textarea name="Remark1_9" style="width:99%; height:100%" wrap="soft" class="TextareaInput"><%=db_Remark1_9%></textarea>
							</td>	
						</tr>
					</table>
					</td>
					<td bgcolor="#EEF6FF" width=70 class="TDCont" align='center'>점수(평균):</td>
					<td bgcolor="#FFFFFF" width=140 class="TDCont" align='center' colspan=2><input type="text" name="TOT_9" value="<%=db_TOT_9%>" size="5" maxlength="5" style="border-width:0px ; border-color:#EEF6FF ; border-style:solid; text-align:right; font-color:#ff0000;font-size:15px;font-weight:bold" readonly ></td>
				</tr>
			</table>

<%if db_factPeoplenum_10 = "" then%>
			<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table width="920" height="30" border="0" cellspacing="0" cellpadding="0" align="center">
				<tr>
					<td align="left"><img src="/Images/Btn/BtnRegiAdd_P<%=right(sGubun,1)%>.GIF" style="cursor:hand;" align="absmiddle" onClick="fn_AddForm('ON','divFORM10','FRM10');">&nbsp;<img src="/Images/Btn/BtnRegiDel_P<%=right(sGubun,1)%>.GIF" style="cursor:hand;" align="absmiddle" onClick="fn_AddForm('OFF','divFORM9','FRM9');"></td>
					<td align='left' width="50%"><input type="checkbox" name="RecYN_9" value="Y" class="none"><font color='#0000ff'><b>녹취첨부</b></font></td>
				</tr>
			</table>
<%else%>
			<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table width="920" height="30" border="0" cellspacing="0" cellpadding="0" align="center">
				<tr>
					<td align="left"><img src="/Images/Btn/BtnRegiDel_P<%=right(sGubun,1)%>.GIF" style="cursor:hand;" align="absmiddle" onClick="fn_AddForm('OFF','divFORM9','FRM9');"></td>
					<td align='left' width="50%"><input type="checkbox" name="RecYN_9" value="Y" class="none"><font color='#0000ff'><b>녹취첨부</b></font></td>

				</tr>
			</table>
<%end if%>

		</td>
	</tr>
</table>
</span>



<%'====== 상담접수 폼 #10 시작 =======================================================================================%>
<%if db_factPeoplenum_10 = "" then%>
	<span id="divFORM10" style="display:none;">
<% else %>
	<span id="divFORM10" style="display:block;">
<%end if%>
<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
<table border="0" width="920" cellpadding="0" cellspacing="1" bgcolor="#EFECE5" align="center">
	<tr bgcolor="#FFFFFF">
		<td>
			<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table width="920" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#ffffff">
			    <tr>
					<td align="left" bgcolor="#FFFFFF" class="TDCont">&nbsp;<img src="/Images/dot_01.gif" >&nbsp;<b><font color="#ff00ff">(<%=sGubunName%>)</font> 설문지#10</b></td>
				</tr>
			</table>
			<table width="920" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">
			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>관 계</td>
					<td bgcolor="#FFFFFF">						<%
							'======= 처리구분 코드 가져오기 ==================================================
							SqlCode = "SELECT CODE, CODENAME FROM TB_CODE"
							SqlCode = SqlCode& " WHERE USEYN='Y' AND SYSYN='Y' AND CODEGROUP='B01'"
							SqlCode = SqlCode& " ORDER BY CODE"
							set RsCode = db.execute(SqlCode)
						%>
						&nbsp;<select name="SECTION2_10" size="1" class="ComboFFFCE7" onChange="fn_UpdateData('10','SECTION2_');">
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("CODE")
										CODENAME = RsCode("CODENAME")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &sGubun& "")%>
							<%
									RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
							%>
						</select>
					&nbsp;&nbsp;&nbsp;&nbsp;<font color="#0000ff"><%=db_factPeoplenum_10%></font></td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>계급(신분)</td>
					<td bgcolor="#FFFFFF"><%
							'======= 처리구분 코드 가져오기 ==================================================
							SqlCode = "SELECT *	FROM armyinformix.dbo.data"
							SqlCode = SqlCode& " where [restrict] = '104' order by [restrict]"
							set RsCode = db.execute(SqlCode)
						%>
						&nbsp;<select name="LEVEL_10" size="1" class="ComboFFFCE7">
							<option value="">계급선택</option>
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("number")
										CODENAME = RsCode("contents")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &db_level_10& "")%>
							<%
									RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
							%>
						</select></td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>성  명</td>
					<td bgcolor="#FFFFFF">&nbsp;<input type="text" name="NAME_10" value="<%=db_NAME_10%>" size="20" maxlength="20" style="border-width:1px ; border-color:#cccccc ; border-style:solid" ></td>
				</tr>
			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>연락처 1</td>
					<td bgcolor="#FFFFFF">&nbsp;<input type="text" name="HOMEPHONE_10" value="<%=db_HOMEPHONE_10%>" size="13" maxlength="13" style="border-width:1px ; border-color:#cccccc ; border-style:solid">
					<img src="/Images/Cti/icon_tel.gif" style="cursor:hand;" onClick="fn_dial('1','10');" align="absmiddle" title="군전화로 전화걸기">&nbsp<img src="/Images/Cti/icon_tel_1.gif" style="cursor:hand;" onClick="fn_dial_1('1','10');" align="absmiddle" title="일반전화로 전화걸기">&nbsp<img src="/Images/Cti/icon_sche.gif" style="cursor:hand;" onClick="fn_sms('1','10');" align="absmiddle" title="문자전송"></td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>연락처 2</td>
					<td bgcolor="#FFFFFF">&nbsp;<input type="text" name="MOBILEPHONE_10" value="<%=db_MOBILEPHONE_10%>" size="13" maxlength="13" style="border-width:1px ; border-color:#cccccc ; border-style:solid">&nbsp;<img src="/Images/Cti/icon_tel.gif" style="cursor:hand;" onClick="fn_dial('2','10');" align="absmiddle" title="군전화로 전화걸기">&nbsp<img src="/Images/Cti/icon_tel_1.gif" style="cursor:hand;" onClick="fn_dial_1('2','10');" align="absmiddle" title="일반전화로 전화걸기">&nbsp<img src="/Images/Cti/icon_sche.gif" style="cursor:hand;" onClick="fn_sms('2','10');" align="absmiddle" title="문자전송"></td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>연락처비고</td>
					<td bgcolor="#FFFFFF">&nbsp;<input type="text" name="ETCPHONE_10" value="<%=db_ETCPHONE_10%>" size="20" maxlength="20" style="border-width:1px ; border-color:#cccccc ; border-style:solid">
					</td>
						
				</tr>
			    <tr>

					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>설문일시</td>
					<td bgcolor="#FFFFFF" width=200>&nbsp;<input type="text" name="MONITORDATE_10" value="<%=db_MONITORDATE_10%>" size="10" maxlength="10" style="border-width:1px ; border-color:#cccccc ; border-style:solid" >&nbsp;<img src="/Images/icon_sche.gif" title="달력" style="cursor:hand;" align="absmiddle" name="Date2_CAR" onClick="window.open('/Include/Calendar_view.asp?goMonth='+document.all.MONITORDATE_10.value+'&firstcode=&FlightCode=&StartEndGu=S&ControlName=opener.document.all.MONITORDATE_10','','width=800,height=300,top=200,left=300,scrollbars=no,status=no');" onblur="setOutColor(this);">
						&nbsp;<img src="/Images/Btn/BtnIconDel.gif" title="삭제" style="cursor:hand;" align="absmiddle"onclick="document.all.MONITORDATE_10.value='';">
					</td>

					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>설문결과</td>
					<td bgcolor="#FFFFFF" width=200>&nbsp;<font color="#0000ff"><%=db_getCodeName("B10",db_MONITOR_RESULT_10)%></font>
					</td>

					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>예약일시</td>
					<td bgcolor="#FFFFFF" width=200>&nbsp;<%=db_RESERVEDATE_10%>
					</td>
				</tr>
				<tr>
					<td bgcolor="#FFEEF9" class="TDCont" align='center' width=100>통화내역</td>
					<td bgcolor="#FFFFFF" colspan=6>&nbsp;<span id="HISTORY_10"><%=db_History_10%></span>
					</td>	
				</tr>

			</table>
			<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table width="920" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">
			    <tr height="25">
					<td bgcolor="#EEF6FF" width=530 rowspan=2 class="TDCont"  colspan='2' align='center'>질문사항</td>
					<td bgcolor="#EEF6FF" colspan='3' class="TDCont" align='center' width=210>답변결과</td>
					<td bgcolor="#EEF6FF" rowspan=2 class="TDCont" align='center' width=70>가점<br>(+1)</td>
					<td bgcolor="#EEF6FF" rowspan=2 class="TDCont" align='center' width=70>점수</td>
					<td bgcolor="#EEF6FF" rowspan=2 class="TDCont" align='center' width=70>초기화<br><img src="/Images/Btn/BtnIconDel.gif" title="점수초기화" style="cursor:hand;" align="absmiddle"onclick="javascript:fn_ALLDEL('10','<%=i%>');"></td>
				</tr>
			    <tr height="20">
					<td bgcolor="#EEF6FF" class="TDCont" align='center' width=70>만족<br>(9)</td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center' width=70>보통<br>(8)</td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center' width=70>불만족<br>(7)</td>
				</tr>
<%

			SQL = "SELECT * FROM TB_CODE where CODEGROUP = '" & sGubun &"' AND USEYN = 'Y' ORDER BY CODE"
			SET Rs = DB.execute(SQL)			

			i = 0
			do until rs.eof
				i = i + 1
				if ( i mod 2 ) = 1 then
					sBgColor = "#ffffff"
				else
					sBgColor = "#FFFCE7"				
				end if

				'값 불러오기
				SQL1 = "select *"
				SQL1 = SQL1 & " from armyinformix.dbo.monitor where factnum = '" & receiptfactnum & "' and factpeoplenum='"& db_factPeoplenum_10 & "' and  seqno = " & i
				
				SET Rs1 = DB.execute(SQL1)	
				if Rs1.eof = false then
					point9 = Rs1("point9")
					point8 = Rs1("point8")
					point7 = Rs1("point7")
					pointplus = Rs1("pointplus")
					totpoint = Rs1("totpoint")
				else
					point9 = ""
					point8 = ""
					point7 = ""
					pointplus = ""
					totpoint = ""
				end if
%>				
			    <tr>
					<td bgcolor="<%=sBgColor%>" width=530  class="TDCont"  colspan='2'>&nbsp;<%=rs("codename")%></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="radio" name="QUESTION_10<%=i%>" value="9" class="none" onClick="fn_YES('10','<%=i%>','9');" <%if point9 = "1" then%>checked<%end if%> ></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="radio" name="QUESTION_10<%=i%>" value="8" class="none" onClick="fn_YES('10','<%=i%>','8');" <%if point8 = "1" then%>checked<%end if%> ></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="radio" name="QUESTION_10<%=i%>" value="7" class="none" onClick="fn_YES('10','<%=i%>','7');" <%if point7 = "1" then%>checked<%end if%> ></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="checkbox" name="QUESTIONP_10<%=i%>" value="1" class="none" onClick="fn_YES('10','<%=i%>','1');"<% if pointplus="1" then Response.Write("checked") end if %>>					
					</td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="text" name="POINT_10<%=i%>" value="<%=totpoint%>" size="2" maxlength="2" style="border-width:0px ; border-color:#cccccc ; border-style:solid; text-align:right; background-color:<%=sBgColor%>" readonly></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><img src="/Images/Btn/BtnIconDel.gif" title="삭제" style="cursor:hand;" align="absmiddle"onclick="javascript:fn_DEL('10','<%=i%>');"></td>

				</tr>
<%
				rs.movenext
			loop
%>
			    <tr ><td bgcolor="#ffffff" class="TDCont"  colspan='5' valign="top" height="100">
					<table width="100%" height="100%" border="0" cellspacing="1" cellpadding="0" align="center" bgcolor="#CCCCCC">
						<tr height="30">
							<td bgcolor="#FFEEF9" class="TDCont" width=100 align='center'>설문결과</td>
							<td bgcolor="#FFFFFF" width="120">
								<%
									'======= 처리구분 코드 가져오기 ==================================================
									SqlCode = "SELECT CODE, CODENAME FROM TB_CODE"
									SqlCode = SqlCode& " WHERE USEYN='Y' AND SYSYN='Y' AND CODEGROUP='B10'"
									SqlCode = SqlCode& " ORDER BY CODE"
									set RsCode = db.execute(SqlCode)
								%>
								&nbsp;<select name="MONITORRESULT_10" size="1" class="ComboFFFCE7" onChange="fn_ResultSet('10');">
									<option value="">설문결과선택</option>
									<%
										IF NOT(RsCode.Eof OR RsCode.bof) THEN
											DO until RsCode.EOF
												CODE = RsCode("CODE")
												CODENAME = RsCode("CODENAME")
									%>
									<%=printSelect("" &CODENAME& "","" &CODE& "","" &db_MONITOR_RESULT_10& "")%>
									<%
											RsCode.MoveNext
											LOOP
										END IF
										RsCode.Close
										set RsCode = NOTHING
									%>
								</select>
							
							</td>


							<td bgcolor="#FFEEF9" class="TDCont" align='center' width="100">상담예약일시</td>
							<td bgcolor="#FFFFFF">&nbsp;<input value="<%=RESERVEDATE_10%>" name="RESERVEDATE_10" readonly type="text" size="10" onfocus="setFocusColor(this);" onClick="new CalendarFrame.Calendar(this);" onblur="setOutColor(this);" onchange="fn_settime('10')">&nbsp;<img src="/Images/icon_sche.gif" title="달력" style="cursor:hand;" align="absmiddle" name="RESERVE_CAR_10" onClick="window.open('/Include/Calendar_view.asp?goMonth='+document.ListForm.RESERVEDATE_10.value+'&firstcode=&FlightCode=&StartEndGu=S&ControlName=opener.document.ListForm.RESERVEDATE_10','','width=800,height=300,top=200,left=300,scrollbars=no,status=no');" onblur="setOutColor(this);" >&nbsp;<img src="/Images/Btn/BtnIconDel.gif" title="삭제" name="RESERVE_DEL_10" style="cursor:hand;" align="absmiddle"onclick="document.all.RESERVEDATE_10.value='';">&nbsp;&nbsp;<input value="<%=RESERVEHOUR_10%>" name="RESERVEHOUR_10" type="text" size="2" maxlength="2" onfocus="setFocusColor(this);" >시&nbsp;<input value="<%=RESERVEMIN_10%>" name="RESERVEMIN_10" type="text" size="2" maxlength="2" onfocus="setFocusColor(this);">분&nbsp;&nbsp;&nbsp;&nbsp;&nbsp<select name="RESERVETIME_10" size="1" class="ComboFFFCE7" onchange="fn_settime('10')">
									<Option value ='' selected>시간선택</option>
									<Option value ='1' >10분후</option>
									<Option value ='2' >30분후</option>
									<Option value ='3' >1시간후</option>
									<Option value ='4' >2시간후</option>
									<Option value ='08' >오전 7시</option>
									<Option value ='08' >오전 8시</option>
									<Option value ='09' >오전 9시</option>
									<Option value ='10' >오전10시</option>
									<Option value ='11' >오전11시</option>
									<Option value ='12' >오후12시</option>
									<Option value ='13' >오후13시</option>
									<Option value ='14' >오후14시</option>
									<Option value ='15' >오후15시</option>
									<Option value ='16' >오후16시</option>
									<Option value ='17' >오후17시</option>
									<Option value ='18' >오후18시</option>
									<Option value ='19' >오후19시</option>
									<Option value ='20' >오후20시</option>
									<Option value ='21' >오후21시</option>
									<Option value ='22' >오후22시</option>

								</select>
							</td>
						</tr>
						<tr height="40">
							<td bgcolor="#FFEEF9" class="TDCont" align='center' width=100>종합평가</td>
							<td bgcolor="#FFFFFF" colspan=6>&nbsp;<textarea name="Remark_10" style="width:99%; height:100%" wrap="soft" class="TextareaInput"><%=db_Remark_10%></textarea>
							</td>	
						</tr>
						<tr height="40">
							<td bgcolor="#FFEEF9" class="TDCont" align='center' width=100>비고</td>
							<td bgcolor="#FFFFFF" colspan=6>&nbsp;<textarea name="Remark1_10" style="width:99%; height:100%" wrap="soft" class="TextareaInput"><%=db_Remark1_10%></textarea>
							</td>	
						</tr>
					</table>
					</td>
					<td bgcolor="#EEF6FF" width=70 class="TDCont" align='center'>점수(평균):</td>
					<td bgcolor="#FFFFFF" width=140 class="TDCont" align='center' colspan=2><input type="text" name="TOT_10" value="<%=db_TOT_10%>" size="5" maxlength="5" style="border-width:0px ; border-color:#EEF6FF ; border-style:solid; text-align:right; font-color:#ff0000;font-size:15px;font-weight:bold" readonly ></td>
				</tr>
			</table>

<%if db_factPeoplenum_11 = "" then%>
			<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table width="920" height="30" border="0" cellspacing="0" cellpadding="0" align="center">
				<tr>
					<td align="left"><img src="/Images/Btn/BtnRegiAdd_P<%=right(sGubun,1)%>.GIF" style="cursor:hand;" align="absmiddle" onClick="fn_AddForm('ON','divFORM11','FRM11');">&nbsp;<img src="/Images/Btn/BtnRegiDel_P<%=right(sGubun,1)%>.GIF" style="cursor:hand;" align="absmiddle" onClick="fn_AddForm('OFF','divFORM10','FRM10');"></td>
					<td align='left' width="50%"><input type="checkbox" name="RecYN_10" value="Y" class="none"><font color='#0000ff'><b>녹취첨부</b></font></td>
				</tr>
			</table>
<%else%>
			<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table width="920" height="30" border="0" cellspacing="0" cellpadding="0" align="center">
				<tr>
					<td align="left"><img src="/Images/Btn/BtnRegiDel_P<%=right(sGubun,1)%>.GIF" style="cursor:hand;" align="absmiddle" onClick="fn_AddForm('OFF','divFORM10','FRM10');"></td>
					<td align='left' width="50%"><input type="checkbox" name="RecYN_10" value="Y" class="none"><font color='#0000ff'><b>녹취첨부</b></font></td>
				</tr>
			</table>
<%end if%>

		</td>
	</tr>
</table>
</span>



<%'====== 상담접수 폼 #11 시작 =======================================================================================%>
<%if db_factPeoplenum_11 = "" then%>
	<span id="divFORM11" style="display:none;">
<% else %>
	<span id="divFORM11" style="display:block;">
<%end if%>
<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
<table border="0" width="920" cellpadding="0" cellspacing="1" bgcolor="#EFECE5" align="center">
	<tr bgcolor="#FFFFFF">
		<td>
			<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table width="920" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#ffffff">
			    <tr>
					<td align="left" bgcolor="#FFFFFF" class="TDCont">&nbsp;<img src="/Images/dot_01.gif" >&nbsp;<b><font color="#ff00ff">(<%=sGubunName%>)</font> 설문지#11</b></td>
				</tr>
			</table>
			<table width="920" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">
			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>관 계</td>
					<td bgcolor="#FFFFFF">						<%
							'======= 처리구분 코드 가져오기 ==================================================
							SqlCode = "SELECT CODE, CODENAME FROM TB_CODE"
							SqlCode = SqlCode& " WHERE USEYN='Y' AND SYSYN='Y' AND CODEGROUP='B01'"
							SqlCode = SqlCode& " ORDER BY CODE"
							set RsCode = db.execute(SqlCode)
						%>
						&nbsp;<select name="SECTION2_11" size="1" class="ComboFFFCE7" onChange="fn_UpdateData('11','SECTION2_');">
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("CODE")
										CODENAME = RsCode("CODENAME")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &sGubun& "")%>
							<%
									RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
							%>
						</select>
					&nbsp;&nbsp;&nbsp;&nbsp;<font color="#0000ff"><%=db_factPeoplenum_11%></font></td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>계급(신분)</td>
					<td bgcolor="#FFFFFF"><%
							'======= 처리구분 코드 가져오기 ==================================================
							SqlCode = "SELECT *	FROM armyinformix.dbo.data"
							SqlCode = SqlCode& " where [restrict] = '104' order by [restrict]"
							set RsCode = db.execute(SqlCode)
						%>
						&nbsp;<select name="LEVEL_11" size="1" class="ComboFFFCE7">
							<option value="">계급선택</option>
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("number")
										CODENAME = RsCode("contents")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &db_level_11& "")%>
							<%
									RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
							%>
						</select></td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>성  명</td>
					<td bgcolor="#FFFFFF">&nbsp;<input type="text" name="NAME_11" value="<%=db_NAME_11%>" size="20" maxlength="20" style="border-width:1px ; border-color:#cccccc ; border-style:solid" ></td>
				</tr>
			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>연락처 1</td>
					<td bgcolor="#FFFFFF">&nbsp;<input type="text" name="HOMEPHONE_11" value="<%=db_HOMEPHONE_11%>" size="13" maxlength="13" style="border-width:1px ; border-color:#cccccc ; border-style:solid">
					<img src="/Images/Cti/icon_tel.gif" style="cursor:hand;" onClick="fn_dial('1','11');" align="absmiddle" title="군전화로 전화걸기">&nbsp<img src="/Images/Cti/icon_tel_1.gif" style="cursor:hand;" onClick="fn_dial_1('1','11');" align="absmiddle" title="일반전화로 전화걸기">&nbsp<img src="/Images/Cti/icon_sche.gif" style="cursor:hand;" onClick="fn_sms('1','11');" align="absmiddle" title="문자전송"></td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>연락처 2</td>
					<td bgcolor="#FFFFFF">&nbsp;<input type="text" name="MOBILEPHONE_11" value="<%=db_MOBILEPHONE_11%>" size="13" maxlength="13" style="border-width:1px ; border-color:#cccccc ; border-style:solid">&nbsp;<img src="/Images/Cti/icon_tel.gif" style="cursor:hand;" onClick="fn_dial('2','11');" align="absmiddle" title="군전화로 전화걸기">&nbsp<img src="/Images/Cti/icon_tel_1.gif" style="cursor:hand;" onClick="fn_dial_1('2','11');" align="absmiddle" title="일반전화로 전화걸기">&nbsp<img src="/Images/Cti/icon_sche.gif" style="cursor:hand;" onClick="fn_sms('2','11');" align="absmiddle" title="문자전송"></td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>연락처비고</td>
					<td bgcolor="#FFFFFF">&nbsp;<input type="text" name="ETCPHONE_11" value="<%=db_ETCPHONE_11%>" size="20" maxlength="20" style="border-width:1px ; border-color:#cccccc ; border-style:solid">
					</td>
						
				</tr>
			    <tr>

					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>설문일시</td>
					<td bgcolor="#FFFFFF" width=200>&nbsp;<input type="text" name="MONITORDATE_11" value="<%=db_MONITORDATE_11%>" size="25" maxlength="25" style="border-width:0px ; border-color:#cccccc ; border-style:solid" readonly>
					</td>

					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>설문결과</td>
					<td bgcolor="#FFFFFF" width=200>&nbsp;<font color="#0000ff"><%=db_getCodeName("B10",db_MONITOR_RESULT_11)%></font>
					</td>

					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>예약일시</td>
					<td bgcolor="#FFFFFF" width=200>&nbsp;<%=db_RESERVEDATE_11%>
					</td>
				</tr>
				<tr>
					<td bgcolor="#FFEEF9" class="TDCont" align='center' width=100>통화내역</td>
					<td bgcolor="#FFFFFF" colspan=6>&nbsp;<span id="HISTORY_11"><%=db_History_11%></span>
					</td>	
				</tr>

			</table>
			<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table width="920" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">
			    <tr height="25">
					<td bgcolor="#EEF6FF" width=530 rowspan=2 class="TDCont"  colspan='2' align='center'>질문사항</td>
					<td bgcolor="#EEF6FF" colspan='3' class="TDCont" align='center' width=210>답변결과</td>
					<td bgcolor="#EEF6FF" rowspan=2 class="TDCont" align='center' width=70>가점<br>(+1)</td>
					<td bgcolor="#EEF6FF" rowspan=2 class="TDCont" align='center' width=70>점수</td>
					<td bgcolor="#EEF6FF" rowspan=2 class="TDCont" align='center' width=70>초기화<br><img src="/Images/Btn/BtnIconDel.gif" title="점수초기화" style="cursor:hand;" align="absmiddle"onclick="javascript:fn_ALLDEL('11','<%=i%>');"></td>
				</tr>
			    <tr height="20">
					<td bgcolor="#EEF6FF" class="TDCont" align='center' width=70>만족<br>(9)</td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center' width=70>보통<br>(8)</td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center' width=70>불만족<br>(7)</td>
				</tr>
<%

			SQL = "SELECT * FROM TB_CODE where CODEGROUP = '" & sGubun &"' AND USEYN = 'Y' ORDER BY CODE"
			SET Rs = DB.execute(SQL)			

			i = 0
			do until rs.eof
				i = i + 1
				if ( i mod 2 ) = 1 then
					sBgColor = "#ffffff"
				else
					sBgColor = "#FFFCE7"				
				end if

				'값 불러오기
				SQL1 = "select *"
				SQL1 = SQL1 & " from armyinformix.dbo.monitor where factnum = '" & receiptfactnum & "' and factpeoplenum='"& db_factPeoplenum_11 & "' and  seqno = " & i
				
				SET Rs1 = DB.execute(SQL1)	
				if Rs1.eof = false then
					point9 = Rs1("point9")
					point8 = Rs1("point8")
					point7 = Rs1("point7")
					pointplus = Rs1("pointplus")
					totpoint = Rs1("totpoint")
				else
					point9 = ""
					point8 = ""
					point7 = ""
					pointplus = ""
					totpoint = ""
				end if
%>				
			    <tr>
					<td bgcolor="<%=sBgColor%>" width=530  class="TDCont"  colspan='2'>&nbsp;<%=rs("codename")%></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="radio" name="QUESTION_11<%=i%>" value="9" class="none" onClick="fn_YES('11','<%=i%>','9');" <%if point9 = "1" then%>checked<%end if%> ></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="radio" name="QUESTION_11<%=i%>" value="8" class="none" onClick="fn_YES('11','<%=i%>','8');" <%if point8 = "1" then%>checked<%end if%> ></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="radio" name="QUESTION_11<%=i%>" value="7" class="none" onClick="fn_YES('11','<%=i%>','7');" <%if point7 = "1" then%>checked<%end if%> ></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="checkbox" name="QUESTIONP_11<%=i%>" value="1" class="none" onClick="fn_YES('11','<%=i%>','1');"<% if pointplus="1" then Response.Write("checked") end if %>>					
					</td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="text" name="POINT_11<%=i%>" value="<%=totpoint%>" size="2" maxlength="2" style="border-width:0px ; border-color:#cccccc ; border-style:solid; text-align:right; background-color:<%=sBgColor%>" readonly></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><img src="/Images/Btn/BtnIconDel.gif" title="삭제" style="cursor:hand;" align="absmiddle"onclick="javascript:fn_DEL('11','<%=i%>');"></td>

				</tr>
<%
				rs.movenext
			loop
%>
			    <tr ><td bgcolor="#ffffff" class="TDCont"  colspan='5' valign="top" height="100">
					<table width="100%" height="100%" border="0" cellspacing="1" cellpadding="0" align="center" bgcolor="#CCCCCC">
						<tr height="30">
							<td bgcolor="#FFEEF9" class="TDCont" width=100 align='center'>설문결과</td>
							<td bgcolor="#FFFFFF" width="120">
								<%
									'======= 처리구분 코드 가져오기 ==================================================
									SqlCode = "SELECT CODE, CODENAME FROM TB_CODE"
									SqlCode = SqlCode& " WHERE USEYN='Y' AND SYSYN='Y' AND CODEGROUP='B10'"
									SqlCode = SqlCode& " ORDER BY CODE"
									set RsCode = db.execute(SqlCode)
								%>
								&nbsp;<select name="MONITORRESULT_11" size="1" class="ComboFFFCE7" onChange="fn_ResultSet('11');">
									<option value="">설문결과선택</option>
									<%
										IF NOT(RsCode.Eof OR RsCode.bof) THEN
											DO until RsCode.EOF
												CODE = RsCode("CODE")
												CODENAME = RsCode("CODENAME")
									%>
									<%=printSelect("" &CODENAME& "","" &CODE& "","" &db_MONITOR_RESULT_11& "")%>
									<%
											RsCode.MoveNext
											LOOP
										END IF
										RsCode.Close
										set RsCode = NOTHING
									%>
								</select>
							
							</td>


							<td bgcolor="#FFEEF9" class="TDCont" align='center' width="100">상담예약일시</td>
							<td bgcolor="#FFFFFF">&nbsp;<input value="<%=RESERVEDATE_11%>" name="RESERVEDATE_11" readonly type="text" size="10" onfocus="setFocusColor(this);" onClick="new CalendarFrame.Calendar(this);" onblur="setOutColor(this);" onchange="fn_settime('11')">&nbsp;<img src="/Images/icon_sche.gif" title="달력" style="cursor:hand;" align="absmiddle" name="RESERVE_CAR_11" onClick="window.open('/Include/Calendar_view.asp?goMonth='+document.ListForm.RESERVEDATE_11.value+'&firstcode=&FlightCode=&StartEndGu=S&ControlName=opener.document.ListForm.RESERVEDATE_11','','width=800,height=300,top=200,left=300,scrollbars=no,status=no');" onblur="setOutColor(this);" >&nbsp;<img src="/Images/Btn/BtnIconDel.gif" title="삭제" name="RESERVE_DEL_11" style="cursor:hand;" align="absmiddle"onclick="document.all.RESERVEDATE_11.value='';">&nbsp;&nbsp;<input value="<%=RESERVEHOUR_11%>" name="RESERVEHOUR_11" type="text" size="2" maxlength="2" onfocus="setFocusColor(this);" >시&nbsp;<input value="<%=RESERVEMIN_11%>" name="RESERVEMIN_11" type="text" size="2" maxlength="2" onfocus="setFocusColor(this);">분&nbsp;&nbsp;&nbsp;&nbsp;&nbsp<select name="RESERVETIME_11" size="1" class="ComboFFFCE7" onchange="fn_settime('11')">
									<Option value ='' selected>시간선택</option>
									<Option value ='1' >10분후</option>
									<Option value ='2' >30분후</option>
									<Option value ='3' >1시간후</option>
									<Option value ='4' >2시간후</option>
									<Option value ='08' >오전 7시</option>
									<Option value ='08' >오전 8시</option>
									<Option value ='09' >오전 9시</option>
									<Option value ='10' >오전10시</option>
									<Option value ='11' >오전11시</option>
									<Option value ='12' >오후12시</option>
									<Option value ='13' >오후13시</option>
									<Option value ='14' >오후14시</option>
									<Option value ='15' >오후15시</option>
									<Option value ='16' >오후16시</option>
									<Option value ='17' >오후17시</option>
									<Option value ='18' >오후18시</option>
									<Option value ='19' >오후19시</option>
									<Option value ='20' >오후20시</option>
									<Option value ='21' >오후21시</option>
									<Option value ='22' >오후22시</option>

								</select>
							</td>
						</tr>
						<tr height="40">
							<td bgcolor="#FFEEF9" class="TDCont" align='center' width=100>종합평가</td>
							<td bgcolor="#FFFFFF" colspan=6>&nbsp;<textarea name="Remark_11" style="width:99%; height:100%" wrap="soft" class="TextareaInput"><%=db_Remark_11%></textarea>
							</td>	
						</tr>
						<tr height="40">
							<td bgcolor="#FFEEF9" class="TDCont" align='center' width=100>비고</td>
							<td bgcolor="#FFFFFF" colspan=6>&nbsp;<textarea name="Remark1_11" style="width:99%; height:100%" wrap="soft" class="TextareaInput"><%=db_Remark1_11%></textarea>
							</td>	
						</tr>
					</table>
					</td>
					<td bgcolor="#EEF6FF" width=70 class="TDCont" align='center'>점수(평균):</td>
					<td bgcolor="#FFFFFF" width=140 class="TDCont" align='center' colspan=2><input type="text" name="TOT_11" value="<%=db_TOT_11%>" size="5" maxlength="5" style="border-width:0px ; border-color:#EEF6FF ; border-style:solid; text-align:right; font-color:#ff0000;font-size:15px;font-weight:bold" readonly ></td>
				</tr>
			</table>

<%if db_factPeoplenum_12 = "" then%>
			<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table width="920" height="30" border="0" cellspacing="0" cellpadding="0" align="center">
				<tr>
					<td align="left"><img src="/Images/Btn/BtnRegiAdd_P<%=right(sGubun,1)%>.GIF" style="cursor:hand;" align="absmiddle" onClick="fn_AddForm('ON','divFORM12','FRM12');">&nbsp;<img src="/Images/Btn/BtnRegiDel_P<%=right(sGubun,1)%>.GIF" style="cursor:hand;" align="absmiddle" onClick="fn_AddForm('OFF','divFORM11','FRM11');"></td>
					<td align='left' width="50%"><input type="checkbox" name="RecYN_11" value="Y" class="none"><font color='#0000ff'><b>녹취첨부</b></font></td>
				</tr>
			</table>
<%else%>
			<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table width="920" height="30" border="0" cellspacing="0" cellpadding="0" align="center">
				<tr>
					<td align="left"><img src="/Images/Btn/BtnRegiDel_P<%=right(sGubun,1)%>.GIF" style="cursor:hand;" align="absmiddle" onClick="fn_AddForm('OFF','divFORM11','FRM11');"></td>
					<td align='left' width="50%"><input type="checkbox" name="RecYN_11" value="Y" class="none"><font color='#0000ff'><b>녹취첨부</b></font></td>
				</tr>
			</table>
<%end if%>

		</td>
	</tr>
</table>
</span>



<%'====== 상담접수 폼 #12 시작 =======================================================================================%>
<%if db_factPeoplenum_12 = "" then%>
	<span id="divFORM12" style="display:none;">
<% else %>
	<span id="divFORM12" style="display:block;">
<%end if%>
<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
<table border="0" width="920" cellpadding="0" cellspacing="1" bgcolor="#EFECE5" align="center">
	<tr bgcolor="#FFFFFF">
		<td>
			<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table width="920" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#ffffff">
			    <tr>
					<td align="left" bgcolor="#FFFFFF" class="TDCont">&nbsp;<img src="/Images/dot_01.gif" >&nbsp;<b><font color="#ff00ff">(<%=sGubunName%>)</font> 설문지#12</b></td>
				</tr>
			</table>
			<table width="920" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">
			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>관 계</td>
					<td bgcolor="#FFFFFF">						<%
							'======= 처리구분 코드 가져오기 ==================================================
							SqlCode = "SELECT CODE, CODENAME FROM TB_CODE"
							SqlCode = SqlCode& " WHERE USEYN='Y' AND SYSYN='Y' AND CODEGROUP='B01'"
							SqlCode = SqlCode& " ORDER BY CODE"
							set RsCode = db.execute(SqlCode)
						%>
						&nbsp;<select name="SECTION2_12" size="1" class="ComboFFFCE7" onChange="fn_UpdateData('12','SECTION2_');">
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("CODE")
										CODENAME = RsCode("CODENAME")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &sGubun& "")%>
							<%
									RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
							%>
						</select>
					&nbsp;&nbsp;&nbsp;&nbsp;<font color="#0000ff"><%=db_factPeoplenum_12%></font></td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>계급(신분)</td>
					<td bgcolor="#FFFFFF"><%
							'======= 처리구분 코드 가져오기 ==================================================
							SqlCode = "SELECT *	FROM armyinformix.dbo.data"
							SqlCode = SqlCode& " where [restrict] = '104' order by [restrict]"
							set RsCode = db.execute(SqlCode)
						%>
						&nbsp;<select name="LEVEL_12" size="1" class="ComboFFFCE7">
							<option value="">계급선택</option>
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("number")
										CODENAME = RsCode("contents")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &db_level_12& "")%>
							<%
									RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
							%>
						</select></td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>성  명</td>
					<td bgcolor="#FFFFFF">&nbsp;<input type="text" name="NAME_12" value="<%=db_NAME_12%>" size="20" maxlength="20" style="border-width:1px ; border-color:#cccccc ; border-style:solid" ></td>
				</tr>
			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>연락처 1</td>
					<td bgcolor="#FFFFFF">&nbsp;<input type="text" name="HOMEPHONE_12" value="<%=db_HOMEPHONE_12%>" size="13" maxlength="13" style="border-width:1px ; border-color:#cccccc ; border-style:solid">
					<img src="/Images/Cti/icon_tel.gif" style="cursor:hand;" onClick="fn_dial('1','12');" align="absmiddle" title="군전화로 전화걸기">&nbsp<img src="/Images/Cti/icon_tel_1.gif" style="cursor:hand;" onClick="fn_dial_1('1','12');" align="absmiddle" title="일반전화로 전화걸기">&nbsp<img src="/Images/Cti/icon_sche.gif" style="cursor:hand;" onClick="fn_sms('1','12');" align="absmiddle" title="문자전송"></td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>연락처 2</td>
					<td bgcolor="#FFFFFF">&nbsp;<input type="text" name="MOBILEPHONE_12" value="<%=db_MOBILEPHONE_12%>" size="13" maxlength="13" style="border-width:1px ; border-color:#cccccc ; border-style:solid">&nbsp;<img src="/Images/Cti/icon_tel.gif" style="cursor:hand;" onClick="fn_dial('2','12');" align="absmiddle" title="군전화로 전화걸기">&nbsp<img src="/Images/Cti/icon_tel_1.gif" style="cursor:hand;" onClick="fn_dial_1('2','12');" align="absmiddle" title="일반전화로 전화걸기">&nbsp<img src="/Images/Cti/icon_sche.gif" style="cursor:hand;" onClick="fn_sms('2','12');" align="absmiddle" title="문자전송"></td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>연락처비고</td>
					<td bgcolor="#FFFFFF">&nbsp;<input type="text" name="ETCPHONE_12" value="<%=db_ETCPHONE_12%>" size="20" maxlength="20" style="border-width:1px ; border-color:#cccccc ; border-style:solid">
					</td>
						
				</tr>
			    <tr>

					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>설문일시</td>
					<td bgcolor="#FFFFFF" width=200>&nbsp;<input type="text" name="MONITORDATE_12" value="<%=db_MONITORDATE_12%>" size="25" maxlength="25" style="border-width:0px ; border-color:#cccccc ; border-style:solid" readonly>
					</td>

					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>설문결과</td>
					<td bgcolor="#FFFFFF" width=200>&nbsp;<font color="#0000ff"><%=db_getCodeName("B10",db_MONITOR_RESULT_12)%></font>
					</td>

					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>예약일시</td>
					<td bgcolor="#FFFFFF" width=200>&nbsp;<%=db_RESERVEDATE_12%>
					</td>
				</tr>
				<tr>
					<td bgcolor="#FFEEF9" class="TDCont" align='center' width=100>통화내역</td>
					<td bgcolor="#FFFFFF" colspan=6>&nbsp;<span id="HISTORY_12"><%=db_History_12%></span>
					</td>	
				</tr>

			</table>
			<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table width="920" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">
			    <tr height="25">
					<td bgcolor="#EEF6FF" width=530 rowspan=2 class="TDCont"  colspan='2' align='center'>질문사항</td>
					<td bgcolor="#EEF6FF" colspan='3' class="TDCont" align='center' width=210>답변결과</td>
					<td bgcolor="#EEF6FF" rowspan=2 class="TDCont" align='center' width=70>가점<br>(+1)</td>
					<td bgcolor="#EEF6FF" rowspan=2 class="TDCont" align='center' width=70>점수</td>
					<td bgcolor="#EEF6FF" rowspan=2 class="TDCont" align='center' width=70>초기화<br><img src="/Images/Btn/BtnIconDel.gif" title="점수초기화" style="cursor:hand;" align="absmiddle"onclick="javascript:fn_ALLDEL('12','<%=i%>');"></td>
				</tr>
			    <tr height="20">
					<td bgcolor="#EEF6FF" class="TDCont" align='center' width=70>만족<br>(9)</td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center' width=70>보통<br>(8)</td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center' width=70>불만족<br>(7)</td>
				</tr>
<%

			SQL = "SELECT * FROM TB_CODE where CODEGROUP = '" & sGubun &"' AND USEYN = 'Y' ORDER BY CODE"
			SET Rs = DB.execute(SQL)			

			i = 0
			do until rs.eof
				i = i + 1
				if ( i mod 2 ) = 1 then
					sBgColor = "#ffffff"
				else
					sBgColor = "#FFFCE7"				
				end if

				'값 불러오기
				SQL1 = "select *"
				SQL1 = SQL1 & " from armyinformix.dbo.monitor where factnum = '" & receiptfactnum & "' and factpeoplenum='"& db_factPeoplenum_12 & "' and  seqno = " & i
				
				SET Rs1 = DB.execute(SQL1)	
				if Rs1.eof = false then
					point9 = Rs1("point9")
					point8 = Rs1("point8")
					point7 = Rs1("point7")
					pointplus = Rs1("pointplus")
					totpoint = Rs1("totpoint")
				else
					point9 = ""
					point8 = ""
					point7 = ""
					pointplus = ""
					totpoint = ""
				end if
%>				
			    <tr>
					<td bgcolor="<%=sBgColor%>" width=530  class="TDCont"  colspan='2'>&nbsp;<%=rs("codename")%></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="radio" name="QUESTION_12<%=i%>" value="9" class="none" onClick="fn_YES('12','<%=i%>','9');" <%if point9 = "1" then%>checked<%end if%> ></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="radio" name="QUESTION_12<%=i%>" value="8" class="none" onClick="fn_YES('12','<%=i%>','8');" <%if point8 = "1" then%>checked<%end if%> ></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="radio" name="QUESTION_12<%=i%>" value="7" class="none" onClick="fn_YES('12','<%=i%>','7');" <%if point7 = "1" then%>checked<%end if%> ></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="checkbox" name="QUESTIONP_12<%=i%>" value="1" class="none" onClick="fn_YES('12','<%=i%>','1');"<% if pointplus="1" then Response.Write("checked") end if %>>					
					</td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="text" name="POINT_12<%=i%>" value="<%=totpoint%>" size="2" maxlength="2" style="border-width:0px ; border-color:#cccccc ; border-style:solid; text-align:right; background-color:<%=sBgColor%>" readonly></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><img src="/Images/Btn/BtnIconDel.gif" title="삭제" style="cursor:hand;" align="absmiddle"onclick="javascript:fn_DEL('12','<%=i%>');"></td>

				</tr>
<%
				rs.movenext
			loop
%>
			    <tr ><td bgcolor="#ffffff" class="TDCont"  colspan='5' valign="top" height="100">
					<table width="100%" height="100%" border="0" cellspacing="1" cellpadding="0" align="center" bgcolor="#CCCCCC">
						<tr height="30">
							<td bgcolor="#FFEEF9" class="TDCont" width=100 align='center'>설문결과</td>
							<td bgcolor="#FFFFFF" width="120">
								<%
									'======= 처리구분 코드 가져오기 ==================================================
									SqlCode = "SELECT CODE, CODENAME FROM TB_CODE"
									SqlCode = SqlCode& " WHERE USEYN='Y' AND SYSYN='Y' AND CODEGROUP='B10'"
									SqlCode = SqlCode& " ORDER BY CODE"
									set RsCode = db.execute(SqlCode)
								%>
								&nbsp;<select name="MONITORRESULT_12" size="1" class="ComboFFFCE7" onChange="fn_ResultSet('12');">
									<option value="">설문결과선택</option>
									<%
										IF NOT(RsCode.Eof OR RsCode.bof) THEN
											DO until RsCode.EOF
												CODE = RsCode("CODE")
												CODENAME = RsCode("CODENAME")
									%>
									<%=printSelect("" &CODENAME& "","" &CODE& "","" &db_MONITOR_RESULT_12& "")%>
									<%
											RsCode.MoveNext
											LOOP
										END IF
										RsCode.Close
										set RsCode = NOTHING
									%>
								</select>
							
							</td>


							<td bgcolor="#FFEEF9" class="TDCont" align='center' width="100">상담예약일시</td>
							<td bgcolor="#FFFFFF">&nbsp;<input value="<%=RESERVEDATE_12%>" name="RESERVEDATE_12" readonly type="text" size="10" onfocus="setFocusColor(this);" onClick="new CalendarFrame.Calendar(this);" onblur="setOutColor(this);" onchange="fn_settime('12')">&nbsp;<img src="/Images/icon_sche.gif" title="달력" style="cursor:hand;" align="absmiddle" name="RESERVE_CAR_12" onClick="window.open('/Include/Calendar_view.asp?goMonth='+document.ListForm.RESERVEDATE_12.value+'&firstcode=&FlightCode=&StartEndGu=S&ControlName=opener.document.ListForm.RESERVEDATE_12','','width=800,height=300,top=200,left=300,scrollbars=no,status=no');" onblur="setOutColor(this);" >&nbsp;<img src="/Images/Btn/BtnIconDel.gif" title="삭제" name="RESERVE_DEL_12" style="cursor:hand;" align="absmiddle"onclick="document.all.RESERVEDATE_12.value='';">&nbsp;&nbsp;<input value="<%=RESERVEHOUR_12%>" name="RESERVEHOUR_12" type="text" size="2" maxlength="2" onfocus="setFocusColor(this);" >시&nbsp;<input value="<%=RESERVEMIN_12%>" name="RESERVEMIN_12" type="text" size="2" maxlength="2" onfocus="setFocusColor(this);">분&nbsp;&nbsp;&nbsp;&nbsp;&nbsp<select name="RESERVETIME_12" size="1" class="ComboFFFCE7" onchange="fn_settime('12')">
									<Option value ='' selected>시간선택</option>
									<Option value ='1' >10분후</option>
									<Option value ='2' >30분후</option>
									<Option value ='3' >1시간후</option>
									<Option value ='4' >2시간후</option>
									<Option value ='08' >오전 7시</option>
									<Option value ='08' >오전 8시</option>
									<Option value ='09' >오전 9시</option>
									<Option value ='10' >오전10시</option>
									<Option value ='11' >오전11시</option>
									<Option value ='12' >오후12시</option>
									<Option value ='13' >오후13시</option>
									<Option value ='14' >오후14시</option>
									<Option value ='15' >오후15시</option>
									<Option value ='16' >오후16시</option>
									<Option value ='17' >오후17시</option>
									<Option value ='18' >오후18시</option>
									<Option value ='19' >오후19시</option>
									<Option value ='20' >오후20시</option>
									<Option value ='21' >오후21시</option>
									<Option value ='22' >오후22시</option>

								</select>
							</td>
						</tr>
						<tr height="40">
							<td bgcolor="#FFEEF9" class="TDCont" align='center' width=100>종합평가</td>
							<td bgcolor="#FFFFFF" colspan=6>&nbsp;<textarea name="Remark_12" style="width:99%; height:100%" wrap="soft" class="TextareaInput"><%=db_Remark_12%></textarea>
							</td>	
						</tr>
						<tr height="40">
							<td bgcolor="#FFEEF9" class="TDCont" align='center' width=100>비고</td>
							<td bgcolor="#FFFFFF" colspan=6>&nbsp;<textarea name="Remark1_12" style="width:99%; height:100%" wrap="soft" class="TextareaInput"><%=db_Remark1_12%></textarea>
							</td>	
						</tr>
					</table>
					</td>
					<td bgcolor="#EEF6FF" width=70 class="TDCont" align='center'>점수(평균):</td>
					<td bgcolor="#FFFFFF" width=140 class="TDCont" align='center' colspan=2><input type="text" name="TOT_12" value="<%=db_TOT_12%>" size="5" maxlength="5" style="border-width:0px ; border-color:#EEF6FF ; border-style:solid; text-align:right; font-color:#ff0000;font-size:15px;font-weight:bold" readonly ></td>
				</tr>
			</table>

<%if db_factPeoplenum_13 = "" then%>
			<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table width="920" height="30" border="0" cellspacing="0" cellpadding="0" align="center">
				<tr>
					<td align="left"><img src="/Images/Btn/BtnRegiAdd_P<%=right(sGubun,1)%>.GIF" style="cursor:hand;" align="absmiddle" onClick="fn_AddForm('ON','divFORM13','FRM13');">&nbsp;<img src="/Images/Btn/BtnRegiDel_P<%=right(sGubun,1)%>.GIF" style="cursor:hand;" align="absmiddle" onClick="fn_AddForm('OFF','divFORM12','FRM12');"></td>
					<td align='left' width="50%"><input type="checkbox" name="RecYN_12" value="Y" class="none"><font color='#0000ff'><b>녹취첨부</b></font></td>
				</tr>
			</table>
<%else%>
			<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table width="920" height="30" border="0" cellspacing="0" cellpadding="0" align="center">
				<tr>
					<td align="left"><img src="/Images/Btn/BtnRegiDel_P<%=right(sGubun,1)%>.GIF" style="cursor:hand;" align="absmiddle" onClick="fn_AddForm('OFF','divFORM12','FRM12');"></td>
					<td align='left' width="50%"><input type="checkbox" name="RecYN_12" value="Y" class="none"><font color='#0000ff'><b>녹취첨부</b></font></td>
				</tr>
			</table>
<%end if%>

		</td>
	</tr>
</table>
</span>



<%'====== 상담접수 폼 #13 시작 =======================================================================================%>
<%if db_factPeoplenum_13 = "" then%>
	<span id="divFORM13" style="display:none;">
<% else %>
	<span id="divFORM13" style="display:block;">
<%end if%>
<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
<table border="0" width="920" cellpadding="0" cellspacing="1" bgcolor="#EFECE5" align="center">
	<tr bgcolor="#FFFFFF">
		<td>
			<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table width="920" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#ffffff">
			    <tr>
					<td align="left" bgcolor="#FFFFFF" class="TDCont">&nbsp;<img src="/Images/dot_01.gif" >&nbsp;<b><font color="#ff00ff">(<%=sGubunName%>)</font> 설문지#13</b></td>
				</tr>
			</table>
			<table width="920" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">
			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>관 계</td>
					<td bgcolor="#FFFFFF">						<%
							'======= 처리구분 코드 가져오기 ==================================================
							SqlCode = "SELECT CODE, CODENAME FROM TB_CODE"
							SqlCode = SqlCode& " WHERE USEYN='Y' AND SYSYN='Y' AND CODEGROUP='B01'"
							SqlCode = SqlCode& " ORDER BY CODE"
							set RsCode = db.execute(SqlCode)
						%>
						&nbsp;<select name="SECTION2_13" size="1" class="ComboFFFCE7" onChange="fn_UpdateData('13','SECTION2_');">
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("CODE")
										CODENAME = RsCode("CODENAME")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &sGubun& "")%>
							<%
									RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
							%>
						</select>
					&nbsp;&nbsp;&nbsp;&nbsp;<font color="#0000ff"><%=db_factPeoplenum_13%></font></td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>계급(신분)</td>
					<td bgcolor="#FFFFFF"><%
							'======= 처리구분 코드 가져오기 ==================================================
							SqlCode = "SELECT *	FROM armyinformix.dbo.data"
							SqlCode = SqlCode& " where [restrict] = '104' order by [restrict]"
							set RsCode = db.execute(SqlCode)
						%>
						&nbsp;<select name="LEVEL_13" size="1" class="ComboFFFCE7">
							<option value="">계급선택</option>
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("number")
										CODENAME = RsCode("contents")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &db_level_13& "")%>
							<%
									RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
							%>
						</select></td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>성  명</td>
					<td bgcolor="#FFFFFF">&nbsp;<input type="text" name="NAME_13" value="<%=db_NAME_13%>" size="20" maxlength="20" style="border-width:1px ; border-color:#cccccc ; border-style:solid" ></td>
				</tr>
			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>연락처 1</td>
					<td bgcolor="#FFFFFF">&nbsp;<input type="text" name="HOMEPHONE_13" value="<%=db_HOMEPHONE_13%>" size="13" maxlength="13" style="border-width:1px ; border-color:#cccccc ; border-style:solid">
					<img src="/Images/Cti/icon_tel.gif" style="cursor:hand;" onClick="fn_dial('1','13');" align="absmiddle" title="군전화로 전화걸기">&nbsp<img src="/Images/Cti/icon_tel_1.gif" style="cursor:hand;" onClick="fn_dial_1('1','13');" align="absmiddle" title="일반전화로 전화걸기">&nbsp<img src="/Images/Cti/icon_sche.gif" style="cursor:hand;" onClick="fn_sms('1','13');" align="absmiddle" title="문자전송"></td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>연락처 2</td>
					<td bgcolor="#FFFFFF">&nbsp;<input type="text" name="MOBILEPHONE_13" value="<%=db_MOBILEPHONE_10%>" size="13" maxlength="13" style="border-width:1px ; border-color:#cccccc ; border-style:solid">&nbsp;<img src="/Images/Cti/icon_tel.gif" style="cursor:hand;" onClick="fn_dial('2','13');" align="absmiddle" title="군전화로 전화걸기">&nbsp<img src="/Images/Cti/icon_tel_1.gif" style="cursor:hand;" onClick="fn_dial_1('2','13');" align="absmiddle" title="일반전화로 전화걸기">&nbsp<img src="/Images/Cti/icon_sche.gif" style="cursor:hand;" onClick="fn_sms('2','13');" align="absmiddle" title="문자전송"></td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>연락처비고</td>
					<td bgcolor="#FFFFFF">&nbsp;<input type="text" name="ETCPHONE_13" value="<%=db_ETCPHONE_13%>" size="20" maxlength="20" style="border-width:1px ; border-color:#cccccc ; border-style:solid">
					</td>
						
				</tr>
			    <tr>

					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>설문일시</td>
					<td bgcolor="#FFFFFF" width=200>&nbsp;<input type="text" name="MONITORDATE_13" value="<%=db_MONITORDATE_13%>" size="25" maxlength="25" style="border-width:0px ; border-color:#cccccc ; border-style:solid" readonly>
					</td>

					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>설문결과</td>
					<td bgcolor="#FFFFFF" width=200>&nbsp;<font color="#0000ff"><%=db_getCodeName("B10",db_MONITOR_RESULT_13)%></font>
					</td>

					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>예약일시</td>
					<td bgcolor="#FFFFFF" width=200>&nbsp;<%=db_RESERVEDATE_13%>
					</td>
				</tr>
				<tr>
					<td bgcolor="#FFEEF9" class="TDCont" align='center' width=100>통화내역</td>
					<td bgcolor="#FFFFFF" colspan=6>&nbsp;<span id="HISTORY_3"><%=db_History_3%></span>
					</td>	
				</tr>

			</table>
			<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table width="920" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">
			    <tr height="25">
					<td bgcolor="#EEF6FF" width=530 rowspan=2 class="TDCont"  colspan='2' align='center'>질문사항</td>
					<td bgcolor="#EEF6FF" colspan='3' class="TDCont" align='center' width=210>답변결과</td>
					<td bgcolor="#EEF6FF" rowspan=2 class="TDCont" align='center' width=70>가점<br>(+1)</td>
					<td bgcolor="#EEF6FF" rowspan=2 class="TDCont" align='center' width=70>점수</td>
					<td bgcolor="#EEF6FF" rowspan=2 class="TDCont" align='center' width=70>초기화<br><img src="/Images/Btn/BtnIconDel.gif" title="점수초기화" style="cursor:hand;" align="absmiddle"onclick="javascript:fn_ALLDEL('13','<%=i%>');"></td>
				</tr>
			    <tr height="20">
					<td bgcolor="#EEF6FF" class="TDCont" align='center' width=70>만족<br>(9)</td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center' width=70>보통<br>(8)</td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center' width=70>불만족<br>(7)</td>
				</tr>
<%

			SQL = "SELECT * FROM TB_CODE where CODEGROUP = '" & sGubun &"' AND USEYN = 'Y' ORDER BY CODE"
			SET Rs = DB.execute(SQL)			

			i = 0
			do until rs.eof
				i = i + 1
				if ( i mod 2 ) = 1 then
					sBgColor = "#ffffff"
				else
					sBgColor = "#FFFCE7"				
				end if

				'값 불러오기
				SQL1 = "select *"
				SQL1 = SQL1 & " from armyinformix.dbo.monitor where factnum = '" & receiptfactnum & "' and factpeoplenum='"& db_factPeoplenum_13 & "' and  seqno = " & i
				
				SET Rs1 = DB.execute(SQL1)	
				if Rs1.eof = false then
					point9 = Rs1("point9")
					point8 = Rs1("point8")
					point7 = Rs1("point7")
					pointplus = Rs1("pointplus")
					totpoint = Rs1("totpoint")
				else
					point9 = ""
					point8 = ""
					point7 = ""
					pointplus = ""
					totpoint = ""
				end if
%>				
			    <tr>
					<td bgcolor="<%=sBgColor%>" width=530  class="TDCont"  colspan='2'>&nbsp;<%=rs("codename")%></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="radio" name="QUESTION_13<%=i%>" value="9" class="none" onClick="fn_YES('13','<%=i%>','9');" <%if point9 = "1" then%>checked<%end if%> ></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="radio" name="QUESTION_13<%=i%>" value="8" class="none" onClick="fn_YES('13','<%=i%>','8');" <%if point8 = "1" then%>checked<%end if%> ></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="radio" name="QUESTION_13<%=i%>" value="7" class="none" onClick="fn_YES('13','<%=i%>','7');" <%if point7 = "1" then%>checked<%end if%> ></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="checkbox" name="QUESTIONP_13<%=i%>" value="1" class="none" onClick="fn_YES('13','<%=i%>','1');"<% if pointplus="1" then Response.Write("checked") end if %>>					
					</td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="text" name="POINT_13<%=i%>" value="<%=totpoint%>" size="2" maxlength="2" style="border-width:0px ; border-color:#cccccc ; border-style:solid; text-align:right; background-color:<%=sBgColor%>" readonly></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><img src="/Images/Btn/BtnIconDel.gif" title="삭제" style="cursor:hand;" align="absmiddle"onclick="javascript:fn_DEL('13','<%=i%>');"></td>

				</tr>
<%
				rs.movenext
			loop
%>
			    <tr ><td bgcolor="#ffffff" class="TDCont"  colspan='5' valign="top" height="100">
					<table width="100%" height="100%" border="0" cellspacing="1" cellpadding="0" align="center" bgcolor="#CCCCCC">
						<tr height="30">
							<td bgcolor="#FFEEF9" class="TDCont" width=100 align='center'>설문결과</td>
							<td bgcolor="#FFFFFF" width="120">
								<%
									'======= 처리구분 코드 가져오기 ==================================================
									SqlCode = "SELECT CODE, CODENAME FROM TB_CODE"
									SqlCode = SqlCode& " WHERE USEYN='Y' AND SYSYN='Y' AND CODEGROUP='B10'"
									SqlCode = SqlCode& " ORDER BY CODE"
									set RsCode = db.execute(SqlCode)
								%>
								&nbsp;<select name="MONITORRESULT_13" size="1" class="ComboFFFCE7" onChange="fn_ResultSet('13');">
									<option value="">설문결과선택</option>
									<%
										IF NOT(RsCode.Eof OR RsCode.bof) THEN
											DO until RsCode.EOF
												CODE = RsCode("CODE")
												CODENAME = RsCode("CODENAME")
									%>
									<%=printSelect("" &CODENAME& "","" &CODE& "","" &db_MONITOR_RESULT_13& "")%>
									<%
											RsCode.MoveNext
											LOOP
										END IF
										RsCode.Close
										set RsCode = NOTHING
									%>
								</select>
							
							</td>


							<td bgcolor="#FFEEF9" class="TDCont" align='center' width="100">상담예약일시</td>
							<td bgcolor="#FFFFFF">&nbsp;<input value="<%=RESERVEDATE_13%>" name="RESERVEDATE_13" readonly type="text" size="10" onfocus="setFocusColor(this);" onClick="new CalendarFrame.Calendar(this);" onblur="setOutColor(this);" onchange="fn_settime('13')">&nbsp;<img src="/Images/icon_sche.gif" title="달력" style="cursor:hand;" align="absmiddle" name="RESERVE_CAR_13" onClick="window.open('/Include/Calendar_view.asp?goMonth='+document.ListForm.RESERVEDATE_13.value+'&firstcode=&FlightCode=&StartEndGu=S&ControlName=opener.document.ListForm.RESERVEDATE_13','','width=800,height=300,top=200,left=300,scrollbars=no,status=no');" onblur="setOutColor(this);" >&nbsp;<img src="/Images/Btn/BtnIconDel.gif" title="삭제" name="RESERVE_DEL_13" style="cursor:hand;" align="absmiddle"onclick="document.all.RESERVEDATE_13.value='';">&nbsp;&nbsp;<input value="<%=RESERVEHOUR_13%>" name="RESERVEHOUR_13" type="text" size="2" maxlength="2" onfocus="setFocusColor(this);" >시&nbsp;<input value="<%=RESERVEMIN_13%>" name="RESERVEMIN_13" type="text" size="2" maxlength="2" onfocus="setFocusColor(this);">분&nbsp;&nbsp;&nbsp;&nbsp;&nbsp<select name="RESERVETIME_13" size="1" class="ComboFFFCE7" onchange="fn_settime('13')">
									<Option value ='' selected>시간선택</option>
									<Option value ='1' >10분후</option>
									<Option value ='2' >30분후</option>
									<Option value ='3' >1시간후</option>
									<Option value ='4' >2시간후</option>
									<Option value ='08' >오전 7시</option>
									<Option value ='08' >오전 8시</option>
									<Option value ='09' >오전 9시</option>
									<Option value ='10' >오전10시</option>
									<Option value ='11' >오전11시</option>
									<Option value ='12' >오후12시</option>
									<Option value ='13' >오후13시</option>
									<Option value ='14' >오후14시</option>
									<Option value ='15' >오후15시</option>
									<Option value ='16' >오후16시</option>
									<Option value ='17' >오후17시</option>
									<Option value ='18' >오후18시</option>
									<Option value ='19' >오후19시</option>
									<Option value ='20' >오후20시</option>
									<Option value ='21' >오후21시</option>
									<Option value ='22' >오후22시</option>

								</select>
							</td>
						</tr>
						<tr height="40">
							<td bgcolor="#FFEEF9" class="TDCont" align='center' width=100>종합평가</td>
							<td bgcolor="#FFFFFF" colspan=6>&nbsp;<textarea name="Remark_13" style="width:99%; height:100%" wrap="soft" class="TextareaInput"><%=db_Remark_13%></textarea>
							</td>	
						</tr>
						<tr height="40">
							<td bgcolor="#FFEEF9" class="TDCont" align='center' width=100>비고</td>
							<td bgcolor="#FFFFFF" colspan=6>&nbsp;<textarea name="Remark1_13" style="width:99%; height:100%" wrap="soft" class="TextareaInput"><%=db_Remark1_13%></textarea>
							</td>	
						</tr>
					</table>
					</td>
					<td bgcolor="#EEF6FF" width=70 class="TDCont" align='center'>점수(평균):</td>
					<td bgcolor="#FFFFFF" width=140 class="TDCont" align='center' colspan=2><input type="text" name="TOT_13" value="<%=db_TOT_13%>" size="5" maxlength="5" style="border-width:0px ; border-color:#EEF6FF ; border-style:solid; text-align:right; font-color:#ff0000;font-size:15px;font-weight:bold" readonly ></td>
				</tr>
			</table>

<%if db_factPeoplenum_14 = "" then%>
			<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table width="920" height="30" border="0" cellspacing="0" cellpadding="0" align="center">
				<tr>
					<td align="left"><img src="/Images/Btn/BtnRegiAdd_P<%=right(sGubun,1)%>.GIF" style="cursor:hand;" align="absmiddle" onClick="fn_AddForm('ON','divFORM14','FRM14');">&nbsp;<img src="/Images/Btn/BtnRegiDel_P<%=right(sGubun,1)%>.GIF" style="cursor:hand;" align="absmiddle" onClick="fn_AddForm('OFF','divFORM13','FRM13');"></td>
					<td align='left' width="50%"><input type="checkbox" name="RecYN_13" value="Y" class="none"><font color='#0000ff'><b>녹취첨부</b></font></td>
				</tr>
			</table>
<%else%>
			<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table width="920" height="30" border="0" cellspacing="0" cellpadding="0" align="center">
				<tr>
					<td align="left"><img src="/Images/Btn/BtnRegiDel_P<%=right(sGubun,1)%>.GIF" style="cursor:hand;" align="absmiddle" onClick="fn_AddForm('OFF','divFORM13','FRM13');"></td>
					<td align='left' width="50%"><input type="checkbox" name="RecYN_13" value="Y" class="none"><font color='#0000ff'><b>녹취첨부</b></font></td>
				</tr>
			</table>
<%end if%>

		</td>
	</tr>
</table>
</span>



<%'====== 상담접수 폼 #14 시작 =======================================================================================%>
<%if db_factPeoplenum_14 = "" then%>
	<span id="divFORM14" style="display:none;">
<% else %>
	<span id="divFORM14" style="display:block;">
<%end if%>
<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
<table border="0" width="920" cellpadding="0" cellspacing="1" bgcolor="#EFECE5" align="center">
	<tr bgcolor="#FFFFFF">
		<td>
			<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table width="920" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#ffffff">
			    <tr>
					<td align="left" bgcolor="#FFFFFF" class="TDCont">&nbsp;<img src="/Images/dot_01.gif" >&nbsp;<b><font color="#ff00ff">(<%=sGubunName%>)</font> 설문지#14</b></td>
				</tr>
			</table>
			<table width="920" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">
			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>관 계</td>
					<td bgcolor="#FFFFFF">						<%
							'======= 처리구분 코드 가져오기 ==================================================
							SqlCode = "SELECT CODE, CODENAME FROM TB_CODE"
							SqlCode = SqlCode& " WHERE USEYN='Y' AND SYSYN='Y' AND CODEGROUP='B01'"
							SqlCode = SqlCode& " ORDER BY CODE"
							set RsCode = db.execute(SqlCode)
						%>
						&nbsp;<select name="SECTION2_14" size="1" class="ComboFFFCE7" onChange="fn_UpdateData('14','SECTION2_');">
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("CODE")
										CODENAME = RsCode("CODENAME")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &sGubun& "")%>
							<%
									RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
							%>
						</select>
					&nbsp;&nbsp;&nbsp;&nbsp;<font color="#0000ff"><%=db_factPeoplenum_14%></font></td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>계급(신분)</td>
					<td bgcolor="#FFFFFF"><%
							'======= 처리구분 코드 가져오기 ==================================================
							SqlCode = "SELECT *	FROM armyinformix.dbo.data"
							SqlCode = SqlCode& " where [restrict] = '104' order by [restrict]"
							set RsCode = db.execute(SqlCode)
						%>
						&nbsp;<select name="LEVEL_14" size="1" class="ComboFFFCE7">
							<option value="">계급선택</option>
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("number")
										CODENAME = RsCode("contents")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &db_level_14& "")%>
							<%
									RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
							%>
						</select></td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>성  명</td>
					<td bgcolor="#FFFFFF">&nbsp;<input type="text" name="NAME_14" value="<%=db_NAME_14%>" size="20" maxlength="20" style="border-width:1px ; border-color:#cccccc ; border-style:solid" ></td>
				</tr>
			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>연락처 1</td>
					<td bgcolor="#FFFFFF">&nbsp;<input type="text" name="HOMEPHONE_14" value="<%=db_HOMEPHONE_14%>" size="13" maxlength="13" style="border-width:1px ; border-color:#cccccc ; border-style:solid">
					<img src="/Images/Cti/icon_tel.gif" style="cursor:hand;" onClick="fn_dial('1','14');" align="absmiddle" title="군전화로 전화걸기">&nbsp<img src="/Images/Cti/icon_tel_1.gif" style="cursor:hand;" onClick="fn_dial_1('1','14');" align="absmiddle" title="일반전화로 전화걸기">&nbsp<img src="/Images/Cti/icon_sche.gif" style="cursor:hand;" onClick="fn_sms('1','14');" align="absmiddle" title="문자전송"></td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>연락처 2</td>
					<td bgcolor="#FFFFFF">&nbsp;<input type="text" name="MOBILEPHONE_14" value="<%=db_MOBILEPHONE_14%>" size="13" maxlength="13" style="border-width:1px ; border-color:#cccccc ; border-style:solid">&nbsp;<img src="/Images/Cti/icon_tel.gif" style="cursor:hand;" onClick="fn_dial('2','14');" align="absmiddle" title="군전화로 전화걸기">&nbsp<img src="/Images/Cti/icon_tel_1.gif" style="cursor:hand;" onClick="fn_dial_1('2','14');" align="absmiddle" title="일반전화로 전화걸기">&nbsp<img src="/Images/Cti/icon_sche.gif" style="cursor:hand;" onClick="fn_sms('2','14');" align="absmiddle" title="문자전송"></td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>연락처비고</td>
					<td bgcolor="#FFFFFF">&nbsp;<input type="text" name="ETCPHONE_14" value="<%=db_ETCPHONE_14%>" size="20" maxlength="20" style="border-width:1px ; border-color:#cccccc ; border-style:solid">
					</td>
						
				</tr>
			    <tr>

					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>설문일시</td>
					<td bgcolor="#FFFFFF" width=200>&nbsp;<input type="text" name="MONITORDATE_14" value="<%=db_MONITORDATE_14%>" size="25" maxlength="25" style="border-width:0px ; border-color:#cccccc ; border-style:solid" readonly>
					</td>

					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>설문결과</td>
					<td bgcolor="#FFFFFF" width=200>&nbsp;<font color="#0000ff"><%=db_getCodeName("B10",db_MONITOR_RESULT_14)%></font>
					</td>

					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>예약일시</td>
					<td bgcolor="#FFFFFF" width=200>&nbsp;<%=db_RESERVEDATE_14%>
					</td>
				</tr>
				<tr>
					<td bgcolor="#FFEEF9" class="TDCont" align='center' width=100>통화내역</td>
					<td bgcolor="#FFFFFF" colspan=6>&nbsp;<span id="HISTORY_14"><%=db_History_14%></span>
					</td>	
				</tr>

			</table>
			<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table width="920" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">
			    <tr height="25">
					<td bgcolor="#EEF6FF" width=530 rowspan=2 class="TDCont"  colspan='2' align='center'>질문사항</td>
					<td bgcolor="#EEF6FF" colspan='3' class="TDCont" align='center' width=210>답변결과</td>
					<td bgcolor="#EEF6FF" rowspan=2 class="TDCont" align='center' width=70>가점<br>(+1)</td>
					<td bgcolor="#EEF6FF" rowspan=2 class="TDCont" align='center' width=70>점수</td>
					<td bgcolor="#EEF6FF" rowspan=2 class="TDCont" align='center' width=70>초기화<br><img src="/Images/Btn/BtnIconDel.gif" title="점수초기화" style="cursor:hand;" align="absmiddle"onclick="javascript:fn_ALLDEL('14','<%=i%>');"></td>
				</tr>
			    <tr height="20">
					<td bgcolor="#EEF6FF" class="TDCont" align='center' width=70>만족<br>(9)</td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center' width=70>보통<br>(8)</td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center' width=70>불만족<br>(7)</td>
				</tr>
<%

			SQL = "SELECT * FROM TB_CODE where CODEGROUP = '" & sGubun &"' AND USEYN = 'Y' ORDER BY CODE"
			SET Rs = DB.execute(SQL)			

			i = 0
			do until rs.eof
				i = i + 1
				if ( i mod 2 ) = 1 then
					sBgColor = "#ffffff"
				else
					sBgColor = "#FFFCE7"				
				end if

				'값 불러오기
				SQL1 = "select *"
				SQL1 = SQL1 & " from armyinformix.dbo.monitor where factnum = '" & receiptfactnum & "' and factpeoplenum='"& db_factPeoplenum_14 & "' and  seqno = " & i
				
				SET Rs1 = DB.execute(SQL1)	
				if Rs1.eof = false then
					point9 = Rs1("point9")
					point8 = Rs1("point8")
					point7 = Rs1("point7")
					pointplus = Rs1("pointplus")
					totpoint = Rs1("totpoint")
				else
					point9 = ""
					point8 = ""
					point7 = ""
					pointplus = ""
					totpoint = ""
				end if
%>				
			    <tr>
					<td bgcolor="<%=sBgColor%>" width=530  class="TDCont"  colspan='2'>&nbsp;<%=rs("codename")%></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="radio" name="QUESTION_14<%=i%>" value="9" class="none" onClick="fn_YES('14','<%=i%>','9');" <%if point9 = "1" then%>checked<%end if%> ></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="radio" name="QUESTION_14<%=i%>" value="8" class="none" onClick="fn_YES('14','<%=i%>','8');" <%if point8 = "1" then%>checked<%end if%> ></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="radio" name="QUESTION_14<%=i%>" value="7" class="none" onClick="fn_YES('14','<%=i%>','7');" <%if point7 = "1" then%>checked<%end if%> ></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="checkbox" name="QUESTIONP_14<%=i%>" value="1" class="none" onClick="fn_YES('14','<%=i%>','1');"<% if pointplus="1" then Response.Write("checked") end if %>>					
					</td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="text" name="POINT_14<%=i%>" value="<%=totpoint%>" size="2" maxlength="2" style="border-width:0px ; border-color:#cccccc ; border-style:solid; text-align:right; background-color:<%=sBgColor%>" readonly></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><img src="/Images/Btn/BtnIconDel.gif" title="삭제" style="cursor:hand;" align="absmiddle"onclick="javascript:fn_DEL('14','<%=i%>');"></td>

				</tr>
<%
				rs.movenext
			loop
%>
			    <tr ><td bgcolor="#ffffff" class="TDCont"  colspan='5' valign="top" height="100">
					<table width="100%" height="100%" border="0" cellspacing="1" cellpadding="0" align="center" bgcolor="#CCCCCC">
						<tr height="30">
							<td bgcolor="#FFEEF9" class="TDCont" width=100 align='center'>설문결과</td>
							<td bgcolor="#FFFFFF" width="120">
								<%
									'======= 처리구분 코드 가져오기 ==================================================
									SqlCode = "SELECT CODE, CODENAME FROM TB_CODE"
									SqlCode = SqlCode& " WHERE USEYN='Y' AND SYSYN='Y' AND CODEGROUP='B10'"
									SqlCode = SqlCode& " ORDER BY CODE"
									set RsCode = db.execute(SqlCode)
								%>
								&nbsp;<select name="MONITORRESULT_14" size="1" class="ComboFFFCE7" onChange="fn_ResultSet('14');">
									<option value="">설문결과선택</option>
									<%
										IF NOT(RsCode.Eof OR RsCode.bof) THEN
											DO until RsCode.EOF
												CODE = RsCode("CODE")
												CODENAME = RsCode("CODENAME")
									%>
									<%=printSelect("" &CODENAME& "","" &CODE& "","" &db_MONITOR_RESULT_14& "")%>
									<%
											RsCode.MoveNext
											LOOP
										END IF
										RsCode.Close
										set RsCode = NOTHING
									%>
								</select>
							
							</td>


							<td bgcolor="#FFEEF9" class="TDCont" align='center' width="100">상담예약일시</td>
							<td bgcolor="#FFFFFF">&nbsp;<input value="<%=RESERVEDATE_14%>" name="RESERVEDATE_14" readonly type="text" size="10" onfocus="setFocusColor(this);" onClick="new CalendarFrame.Calendar(this);" onblur="setOutColor(this);" onchange="fn_settime('14')">&nbsp;<img src="/Images/icon_sche.gif" title="달력" style="cursor:hand;" align="absmiddle" name="RESERVE_CAR_14" onClick="window.open('/Include/Calendar_view.asp?goMonth='+document.ListForm.RESERVEDATE_14.value+'&firstcode=&FlightCode=&StartEndGu=S&ControlName=opener.document.ListForm.RESERVEDATE_14','','width=800,height=300,top=200,left=300,scrollbars=no,status=no');" onblur="setOutColor(this);" >&nbsp;<img src="/Images/Btn/BtnIconDel.gif" title="삭제" name="RESERVE_DEL_14" style="cursor:hand;" align="absmiddle"onclick="document.all.RESERVEDATE_14.value='';">&nbsp;&nbsp;<input value="<%=RESERVEHOUR_14%>" name="RESERVEHOUR_14" type="text" size="2" maxlength="2" onfocus="setFocusColor(this);" >시&nbsp;<input value="<%=RESERVEMIN_14%>" name="RESERVEMIN_14" type="text" size="2" maxlength="2" onfocus="setFocusColor(this);">분&nbsp;&nbsp;&nbsp;&nbsp;&nbsp<select name="RESERVETIME_14" size="1" class="ComboFFFCE7" onchange="fn_settime('14')">
									<Option value ='' selected>시간선택</option>
									<Option value ='1' >10분후</option>
									<Option value ='2' >30분후</option>
									<Option value ='3' >1시간후</option>
									<Option value ='4' >2시간후</option>
									<Option value ='08' >오전 7시</option>
									<Option value ='08' >오전 8시</option>
									<Option value ='09' >오전 9시</option>
									<Option value ='10' >오전10시</option>
									<Option value ='11' >오전11시</option>
									<Option value ='12' >오후12시</option>
									<Option value ='13' >오후13시</option>
									<Option value ='14' >오후14시</option>
									<Option value ='15' >오후15시</option>
									<Option value ='16' >오후16시</option>
									<Option value ='17' >오후17시</option>
									<Option value ='18' >오후18시</option>
									<Option value ='19' >오후19시</option>
									<Option value ='20' >오후20시</option>
									<Option value ='21' >오후21시</option>
									<Option value ='22' >오후22시</option>

								</select>
							</td>
						</tr>
						<tr height="40">
							<td bgcolor="#FFEEF9" class="TDCont" align='center' width=100>종합평가</td>
							<td bgcolor="#FFFFFF" colspan=6>&nbsp;<textarea name="Remark_14" style="width:99%; height:100%" wrap="soft" class="TextareaInput"><%=db_Remark_14%></textarea>
							</td>	
						</tr>
						<tr height="40">
							<td bgcolor="#FFEEF9" class="TDCont" align='center' width=100>비고</td>
							<td bgcolor="#FFFFFF" colspan=6>&nbsp;<textarea name="Remark1_14" style="width:99%; height:100%" wrap="soft" class="TextareaInput"><%=db_Remark1_14%></textarea>
							</td>	
						</tr>
					</table>
					</td>
					<td bgcolor="#EEF6FF" width=70 class="TDCont" align='center'>점수(평균):</td>
					<td bgcolor="#FFFFFF" width=140 class="TDCont" align='center' colspan=2><input type="text" name="TOT_14" value="<%=db_TOT_14%>" size="5" maxlength="5" style="border-width:0px ; border-color:#EEF6FF ; border-style:solid; text-align:right; font-color:#ff0000;font-size:15px;font-weight:bold" readonly ></td>
				</tr>
			</table>

<%if db_factPeoplenum_15 = "" then%>
			<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table width="920" height="30" border="0" cellspacing="0" cellpadding="0" align="center">
				<tr>
					<td align="left"><img src="/Images/Btn/BtnRegiAdd_P<%=right(sGubun,1)%>.GIF" style="cursor:hand;" align="absmiddle" onClick="fn_AddForm('ON','divFORM15','FRM15');">&nbsp;<img src="/Images/Btn/BtnRegiDel_P<%=right(sGubun,1)%>.GIF" style="cursor:hand;" align="absmiddle" onClick="fn_AddForm('OFF','divFORM14','FRM14');"></td>
					<td align='left' width="50%"><input type="checkbox" name="RecYN_14" value="Y" class="none"><font color='#0000ff'><b>녹취첨부</b></font></td>
				</tr>
			</table>
<%else%>
			<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table width="920" height="30" border="0" cellspacing="0" cellpadding="0" align="center">
				<tr>
					<td align="left"><img src="/Images/Btn/BtnRegiDel_P<%=right(sGubun,1)%>.GIF" style="cursor:hand;" align="absmiddle" onClick="fn_AddForm('OFF','divFORM14','FRM14');"></td>
					<td align='left' width="50%"><input type="checkbox" name="RecYN_14" value="Y" class="none"><font color='#0000ff'><b>녹취첨부</b></font></td>
				</tr>
			</table>
<%end if%>

		</td>
	</tr>
</table>
</span>



<%'====== 상담접수 폼 #15 시작 =======================================================================================%>
<%if db_factPeoplenum_15 = "" then%>
	<span id="divFORM15" style="display:none;">
<% else %>
	<span id="divFORM15" style="display:block;">
<%end if%>
<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
<table border="0" width="920" cellpadding="0" cellspacing="1" bgcolor="#EFECE5" align="center">
	<tr bgcolor="#FFFFFF">
		<td>
			<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table width="920" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#ffffff">
			    <tr>
					<td align="left" bgcolor="#FFFFFF" class="TDCont">&nbsp;<img src="/Images/dot_01.gif" >&nbsp;<b><font color="#ff00ff">(<%=sGubunName%>)</font> 설문지#15</b></td>
				</tr>
			</table>
			<table width="920" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">
			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>관 계</td>
					<td bgcolor="#FFFFFF">						<%
							'======= 처리구분 코드 가져오기 ==================================================
							SqlCode = "SELECT CODE, CODENAME FROM TB_CODE"
							SqlCode = SqlCode& " WHERE USEYN='Y' AND SYSYN='Y' AND CODEGROUP='B01'"
							SqlCode = SqlCode& " ORDER BY CODE"
							set RsCode = db.execute(SqlCode)
						%>
						&nbsp;<select name="SECTION2_15" size="1" class="ComboFFFCE7" onChange="fn_UpdateData('15','SECTION2_');">
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("CODE")
										CODENAME = RsCode("CODENAME")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &sGubun& "")%>
							<%
									RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
							%>
						</select>
					&nbsp;&nbsp;&nbsp;&nbsp;<font color="#0000ff"><%=db_factPeoplenum_15%></font></td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>계급(신분)</td>
					<td bgcolor="#FFFFFF"><%
							'======= 처리구분 코드 가져오기 ==================================================
							SqlCode = "SELECT *	FROM armyinformix.dbo.data"
							SqlCode = SqlCode& " where [restrict] = '104' order by [restrict]"
							set RsCode = db.execute(SqlCode)
						%>
						&nbsp;<select name="LEVEL_15" size="1" class="ComboFFFCE7">
							<option value="">계급선택</option>
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("number")
										CODENAME = RsCode("contents")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &db_level_15& "")%>
							<%
									RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
							%>
						</select></td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>성  명</td>
					<td bgcolor="#FFFFFF">&nbsp;<input type="text" name="NAME_15" value="<%=db_NAME_15%>" size="20" maxlength="20" style="border-width:1px ; border-color:#cccccc ; border-style:solid" ></td>
				</tr>
			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>연락처 1</td>
					<td bgcolor="#FFFFFF">&nbsp;<input type="text" name="HOMEPHONE_15" value="<%=db_HOMEPHONE_15%>" size="13" maxlength="13" style="border-width:1px ; border-color:#cccccc ; border-style:solid">
					<img src="/Images/Cti/icon_tel.gif" style="cursor:hand;" onClick="fn_dial('1','15');" align="absmiddle" title="군전화로 전화걸기">&nbsp<img src="/Images/Cti/icon_tel_1.gif" style="cursor:hand;" onClick="fn_dial_1('1','15');" align="absmiddle" title="일반전화로 전화걸기">&nbsp<img src="/Images/Cti/icon_sche.gif" style="cursor:hand;" onClick="fn_sms('1','15');" align="absmiddle" title="문자전송"></td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>연락처 2</td>
					<td bgcolor="#FFFFFF">&nbsp;<input type="text" name="MOBILEPHONE_15" value="<%=db_MOBILEPHONE_15%>" size="13" maxlength="13" style="border-width:1px ; border-color:#cccccc ; border-style:solid">&nbsp;<img src="/Images/Cti/icon_tel.gif" style="cursor:hand;" onClick="fn_dial('2','15');" align="absmiddle" title="군전화로 전화걸기">&nbsp<img src="/Images/Cti/icon_tel_1.gif" style="cursor:hand;" onClick="fn_dial_1('2','15');" align="absmiddle" title="일반전화로 전화걸기">&nbsp<img src="/Images/Cti/icon_sche.gif" style="cursor:hand;" onClick="fn_sms('2','15');" align="absmiddle" title="문자전송"></td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center'>연락처비고</td>
					<td bgcolor="#FFFFFF">&nbsp;<input type="text" name="ETCPHONE_15" value="<%=db_ETCPHONE_15%>" size="20" maxlength="20" style="border-width:1px ; border-color:#cccccc ; border-style:solid">
					</td>
						
				</tr>
			    <tr>

					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>설문일시</td>
					<td bgcolor="#FFFFFF" width=200>&nbsp;<input type="text" name="MONITORDATE_15" value="<%=db_MONITORDATE_15%>" size="25" maxlength="25" style="border-width:0px ; border-color:#cccccc ; border-style:solid" readonly>
					</td>

					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>설문결과</td>
					<td bgcolor="#FFFFFF" width=200>&nbsp;<font color="#0000ff"><%=db_getCodeName("B10",db_MONITOR_RESULT_15)%></font>
					</td>

					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>예약일시</td>
					<td bgcolor="#FFFFFF" width=200>&nbsp;<%=db_RESERVEDATE_15%>
					</td>
				</tr>
				<tr>
					<td bgcolor="#FFEEF9" class="TDCont" align='center' width=100>통화내역</td>
					<td bgcolor="#FFFFFF" colspan=6>&nbsp;<span id="HISTORY_15"><%=db_History_15%></span>
					</td>	
				</tr>

			</table>
			<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table width="920" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">
			    <tr height="25">
					<td bgcolor="#EEF6FF" width=530 rowspan=2 class="TDCont"  colspan='2' align='center'>질문사항</td>
					<td bgcolor="#EEF6FF" colspan='3' class="TDCont" align='center' width=210>답변결과</td>
					<td bgcolor="#EEF6FF" rowspan=2 class="TDCont" align='center' width=70>가점<br>(+1)</td>
					<td bgcolor="#EEF6FF" rowspan=2 class="TDCont" align='center' width=70>점수</td>
					<td bgcolor="#EEF6FF" rowspan=2 class="TDCont" align='center' width=70>초기화<br><img src="/Images/Btn/BtnIconDel.gif" title="점수초기화" style="cursor:hand;" align="absmiddle"onclick="javascript:fn_ALLDEL('15','<%=i%>');"></td>
				</tr>
			    <tr height="20">
					<td bgcolor="#EEF6FF" class="TDCont" align='center' width=70>만족<br>(9)</td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center' width=70>보통<br>(8)</td>
					<td bgcolor="#EEF6FF" class="TDCont" align='center' width=70>불만족<br>(7)</td>
				</tr>
<%

			SQL = "SELECT * FROM TB_CODE where CODEGROUP = '" & sGubun &"' AND USEYN = 'Y' ORDER BY CODE"
			SET Rs = DB.execute(SQL)			

			i = 0
			do until rs.eof
				i = i + 1
				if ( i mod 2 ) = 1 then
					sBgColor = "#ffffff"
				else
					sBgColor = "#FFFCE7"				
				end if

				'값 불러오기
				SQL1 = "select *"
				SQL1 = SQL1 & " from armyinformix.dbo.monitor where factnum = '" & receiptfactnum & "' and factpeoplenum='"& db_factPeoplenum_15 & "' and  seqno = " & i
				
				SET Rs1 = DB.execute(SQL1)	
				if Rs1.eof = false then
					point9 = Rs1("point9")
					point8 = Rs1("point8")
					point7 = Rs1("point7")
					pointplus = Rs1("pointplus")
					totpoint = Rs1("totpoint")
				else
					point9 = ""
					point8 = ""
					point7 = ""
					pointplus = ""
					totpoint = ""
				end if
%>				
			    <tr>
					<td bgcolor="<%=sBgColor%>" width=530  class="TDCont"  colspan='2'>&nbsp;<%=rs("codename")%></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="radio" name="QUESTION_15<%=i%>" value="9" class="none" onClick="fn_YES('15','<%=i%>','9');" <%if point9 = "1" then%>checked<%end if%> ></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="radio" name="QUESTION_15<%=i%>" value="8" class="none" onClick="fn_YES('15','<%=i%>','8');" <%if point8 = "1" then%>checked<%end if%> ></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="radio" name="QUESTION_15<%=i%>" value="7" class="none" onClick="fn_YES('15','<%=i%>','7');" <%if point7 = "1" then%>checked<%end if%> ></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="checkbox" name="QUESTIONP_15<%=i%>" value="1" class="none" onClick="fn_YES('15','<%=i%>','1');"<% if pointplus="1" then Response.Write("checked") end if %>>					
					</td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><input type="text" name="POINT_15<%=i%>" value="<%=totpoint%>" size="2" maxlength="2" style="border-width:0px ; border-color:#cccccc ; border-style:solid; text-align:right; background-color:<%=sBgColor%>" readonly></td>
					<td bgcolor="<%=sBgColor%>" class="TDCont" align='center' width=70><img src="/Images/Btn/BtnIconDel.gif" title="삭제" style="cursor:hand;" align="absmiddle"onclick="javascript:fn_DEL('15','<%=i%>');"></td>

				</tr>
<%
				rs.movenext
			loop
%>
			    <tr ><td bgcolor="#ffffff" class="TDCont"  colspan='5' valign="top" height="100">
					<table width="100%" height="100%" border="0" cellspacing="1" cellpadding="0" align="center" bgcolor="#CCCCCC">
						<tr height="30">
							<td bgcolor="#FFEEF9" class="TDCont" width=100 align='center'>설문결과</td>
							<td bgcolor="#FFFFFF" width="120">
								<%
									'======= 처리구분 코드 가져오기 ==================================================
									SqlCode = "SELECT CODE, CODENAME FROM TB_CODE"
									SqlCode = SqlCode& " WHERE USEYN='Y' AND SYSYN='Y' AND CODEGROUP='B10'"
									SqlCode = SqlCode& " ORDER BY CODE"
									set RsCode = db.execute(SqlCode)
								%>
								&nbsp;<select name="MONITORRESULT_15" size="1" class="ComboFFFCE7" onChange="fn_ResultSet('15');">
									<option value="">설문결과선택</option>
									<%
										IF NOT(RsCode.Eof OR RsCode.bof) THEN
											DO until RsCode.EOF
												CODE = RsCode("CODE")
												CODENAME = RsCode("CODENAME")
									%>
									<%=printSelect("" &CODENAME& "","" &CODE& "","" &db_MONITOR_RESULT_15& "")%>
									<%
											RsCode.MoveNext
											LOOP
										END IF
										RsCode.Close
										set RsCode = NOTHING
									%>
								</select>
							
							</td>


							<td bgcolor="#FFEEF9" class="TDCont" align='center' width="100">상담예약일시</td>
							<td bgcolor="#FFFFFF">&nbsp;<input value="<%=RESERVEDATE_15%>" name="RESERVEDATE_15" readonly type="text" size="10" onfocus="setFocusColor(this);" onClick="new CalendarFrame.Calendar(this);" onblur="setOutColor(this);" onchange="fn_settime('15')">&nbsp;<img src="/Images/icon_sche.gif" title="달력" style="cursor:hand;" align="absmiddle" name="RESERVE_CAR_15" onClick="window.open('/Include/Calendar_view.asp?goMonth='+document.ListForm.RESERVEDATE_15.value+'&firstcode=&FlightCode=&StartEndGu=S&ControlName=opener.document.ListForm.RESERVEDATE_15','','width=800,height=300,top=200,left=300,scrollbars=no,status=no');" onblur="setOutColor(this);" >&nbsp;<img src="/Images/Btn/BtnIconDel.gif" title="삭제" name="RESERVE_DEL_15" style="cursor:hand;" align="absmiddle"onclick="document.all.RESERVEDATE_15.value='';">&nbsp;&nbsp;<input value="<%=RESERVEHOUR_15%>" name="RESERVEHOUR_15" type="text" size="2" maxlength="2" onfocus="setFocusColor(this);" >시&nbsp;<input value="<%=RESERVEMIN_15%>" name="RESERVEMIN_15" type="text" size="2" maxlength="2" onfocus="setFocusColor(this);">분&nbsp;&nbsp;&nbsp;&nbsp;&nbsp<select name="RESERVETIME_15" size="1" class="ComboFFFCE7" onchange="fn_settime('15')">
									<Option value ='' selected>시간선택</option>
									<Option value ='1' >10분후</option>
									<Option value ='2' >30분후</option>
									<Option value ='3' >1시간후</option>
									<Option value ='4' >2시간후</option>
									<Option value ='08' >오전 7시</option>
									<Option value ='08' >오전 8시</option>
									<Option value ='09' >오전 9시</option>
									<Option value ='10' >오전10시</option>
									<Option value ='11' >오전11시</option>
									<Option value ='12' >오후12시</option>
									<Option value ='13' >오후13시</option>
									<Option value ='14' >오후14시</option>
									<Option value ='15' >오후15시</option>
									<Option value ='16' >오후16시</option>
									<Option value ='17' >오후17시</option>
									<Option value ='18' >오후18시</option>
									<Option value ='19' >오후19시</option>
									<Option value ='20' >오후20시</option>
									<Option value ='21' >오후21시</option>
									<Option value ='22' >오후22시</option>

								</select>
							</td>
						</tr>
						<tr height="40">
							<td bgcolor="#FFEEF9" class="TDCont" align='center' width=100>종합평가</td>
							<td bgcolor="#FFFFFF" colspan=6>&nbsp;<textarea name="Remark_15" style="width:99%; height:100%" wrap="soft" class="TextareaInput"><%=db_Remark_15%></textarea>
							</td>	
						</tr>
						<tr height="40">
							<td bgcolor="#FFEEF9" class="TDCont" align='center' width=100>비고</td>
							<td bgcolor="#FFFFFF" colspan=6>&nbsp;<textarea name="Remark1_15" style="width:99%; height:100%" wrap="soft" class="TextareaInput"><%=db_Remark1_15%></textarea>
							</td>	
						</tr>
					</table>
					</td>
					<td bgcolor="#EEF6FF" width=70 class="TDCont" align='center'>점수(평균):</td>
					<td bgcolor="#FFFFFF" width=140 class="TDCont" align='center' colspan=2><input type="text" name="TOT_15" value="<%=db_TOT_15%>" size="5" maxlength="5" style="border-width:0px ; border-color:#EEF6FF ; border-style:solid; text-align:right; font-color:#ff0000;font-size:15px;font-weight:bold" readonly ></td>
				</tr>
			</table>

			<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table width="920" height="30" border="0" cellspacing="0" cellpadding="0" align="center">
				<tr>
					<td align="left"><img src="/Images/Btn/BtnRegiDel_P<%=right(sGubun,1)%>.GIF" style="cursor:hand;" align="absmiddle" onClick="fn_AddForm('OFF','divFORM15','FRM15');"></td>
					<td align='left' width="50%"><input type="checkbox" name="RecYN_15" value="Y" class="none"><font color='#0000ff'><b>녹취첨부</b></font></td>
				</tr>
			</table>


		</td>
	</tr>
</table>
</span>

<span id="buttonsavepan" style="display:none;">
	<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
	<table width="920" height="30" border="0" cellspacing="0" cellpadding="0" align="center">
		<tr>
			<td colspan= 4 class="TDCont" align='right'>
			<img src="/Images/Btn/BtnMointorSubmit.GIF" style="cursor:hand;" align="absmiddle" title="(<%=sGubunName%>)설문내용저장" onclick="fn_UpdateData('1','1');">&nbsp;<img src="/Images/Btn/BtnList.gif" style="cursor:hand;" align="absmiddle" onClick="parent.fn_list();"></td>
		</tr>
	</table>
</span>


</form>
<A name="divFORMLink"></a>
<iframe src="about:blank" name="DBFrame" width="0" height="0" frameborder=1 marginheight=0 marginwidth=0 scrolling="no"></iframe>


<!-- #include virtual="/Include/Bottom.asp" -->

<script>
<!--

	parent.document.all.btnsave.title = "(<%=sGubunName%>)설문내용저장";
	parent.document.all.btnsave1.title = "(<%=sGubunName%>)설문내용저장";
	function fn_Tabclick(arg0)
	{

		if ( arg0 == '1' )
			location.href='monitoring_input.asp?receiptfactnum=<%=receiptfactnum%>&sGubun=B11&sGubunName=피의자';
		else if ( arg0 == '2' )
			location.href='monitoring_input.asp?receiptfactnum=<%=receiptfactnum%>&sGubun=B12&sGubunName=피해자';
		else if ( arg0 == '3' )
			location.href='monitoring_input.asp?receiptfactnum=<%=receiptfactnum%>&sGubun=B13&sGubunName=민원인';
		else if ( arg0 == '4' )
			location.href='monitoring_input.asp?receiptfactnum=<%=receiptfactnum%>&sGubun=B14&sGubunName=피민원인';
		else if ( arg0 == '5' )
			location.href='monitoring_input.asp?receiptfactnum=<%=receiptfactnum%>&sGubun=B15&sGubunName=지휘관';
		else if ( arg0 == '6' )
			location.href='monitoring_input.asp?receiptfactnum=<%=receiptfactnum%>&sGubun=B16&sGubunName=유족';
		else if ( arg0 == '7' )
			location.href='monitoring_input.asp?receiptfactnum=<%=receiptfactnum%>&sGubun=B17&sGubunName=참고인';

	}
	function fn_settime(arg0)
	{

		if ( eval("ListForm.RESERVETIME_"+arg0).value == '1' || eval("ListForm.RESERVETIME_"+arg0).value == '2' || eval("ListForm.RESERVETIME_"+arg0).value == '3' || eval("ListForm.RESERVETIME_"+arg0).value == '4' )
		{
			DBFrame.location= "/menu01/submenu0101/time_calculation.asp?DateControlName=parent.ListForm.RESERVEDATE_"+arg0+"&HourControlName=parent.ListForm.RESERVEHOUR_"+arg0+"&MinControlName=parent.ListForm.RESERVEMIN_"+arg0+"&RESERVETIME="+eval("ListForm.RESERVETIME_"+arg0).value;
		}
		else
		{
			eval("ListForm.RESERVEHOUR_"+arg0).value = eval("ListForm.RESERVETIME_"+arg0).value;
			eval("ListForm.RESERVEMIN_"+arg0).value = "00";
		}
		eval("ListForm.MONITORRESULT_"+arg0).value= "4";

	}
	function fn_UpdateData(arg0,arg1)
	{
		if ( arg1 == 'SECTION2_' )
		{
			if ( confirm('사건과의 관계를 수정하시겠습니까?') )
			{
				// 자료 업데이트 arg0 = 폼번호, arg1
				ListForm.Date2.value = parent.document.all.Date2.value;
				ListForm.Date3.value = parent.document.all.Date3.value;
				ListForm.receiptkind.value = parent.document.all.receiptkind.value;
				ListForm.submit();
			}
			else
			{
				eval("document.ListForm.SECTION2_"+arg0).value ="<%=sGubun%>";
			}
		}
		else
		{
			if ( parent.document.all.receiptkind.value == '' )
			{
				alert('사건유형을 선택해 주세요!');
				return false;
			}
			// 자료 업데이트 arg0 = 폼번호, arg1
			ListForm.Date2.value = parent.document.all.Date2.value;
			ListForm.Date3.value = parent.document.all.Date3.value;
			ListForm.receiptkind.value = parent.document.all.receiptkind.value;
			ListForm.submit();
		}

	}
	function fn_AddForm(ty,f,ck){
		if(ty=="ON"){
			eval(f).style.display = "block";
			eval("ListForm."+ck).value = "ON";
		} else {
			if(confirm("자료를 삭제 하시겠습니까?")) {
				if(f=="divFORM1"){
					//eval(f).style.display = "none";
					eval("ListForm."+ck).value = "";

					ListForm.Date2.value = parent.document.all.Date2.value;
					ListForm.Date3.value = parent.document.all.Date3.value;
					ListForm.receiptkind.value = parent.document.all.receiptkind.value;
					ListForm.submit();		
					
				}
				else
				{
					eval("ListForm."+ck).value = "";
					if(f=="divFORM2" && ListForm.factPeoplenum_2.value !='')
					{

						ListForm.Date2.value = parent.document.all.Date2.value;
						ListForm.Date3.value = parent.document.all.Date3.value;
						ListForm.receiptkind.value = parent.document.all.receiptkind.value;
						ListForm.submit();		
					}
					else if(f=="divFORM3" && ListForm.factPeoplenum_3.value !='')
					{
						ListForm.Date2.value = parent.document.all.Date2.value;
						ListForm.Date3.value = parent.document.all.Date3.value;
						ListForm.receiptkind.value = parent.document.all.receiptkind.value;
						ListForm.submit();		
					}
					else if(f=="divFORM4" && ListForm.factPeoplenum_4.value !='')
					{
						ListForm.Date2.value = parent.document.all.Date2.value;
						ListForm.Date3.value = parent.document.all.Date3.value;
						ListForm.receiptkind.value = parent.document.all.receiptkind.value;
						ListForm.submit();		
					}
					else if(f=="divFORM5" && ListForm.factPeoplenum_5.value !='')
					{
						ListForm.Date2.value = parent.document.all.Date2.value;
						ListForm.Date3.value = parent.document.all.Date3.value;
						ListForm.receiptkind.value = parent.document.all.receiptkind.value;
						ListForm.submit();		
					}
					else
					{
						eval(f).style.display = "none";
					}
				}
			}
		}
	}

	function fn_ResultSet(arg0)
	{
		if ( eval("ListForm.MONITORRESULT_"+arg0).value == "4" )
		{
			eval("ListForm.RESERVEDATE_"+arg0).disabled = false;
			eval("ListForm.RESERVETIME_"+arg0).disabled = false;
			eval("ListForm.RESERVE_CAR_"+arg0).disabled = false;
			eval("ListForm.RESERVE_DEL_"+arg0).disabled = false;
			eval("ListForm.RESERVEHOUR_"+arg0).disabled = false;
			eval("ListForm.RESERVEMIN_"+arg0).disabled = false;
			if ( eval("ListForm.RESERVEDATE_"+arg0).value == "" )
			{
				eval("ListForm.RESERVEDATE_"+arg0).value = "<%=sToday%>";
			}
			eval("ListForm.RESERVEDATE_"+arg0).focus();

		}
		else
		{
			eval("ListForm.RESERVEDATE_"+arg0).disabled = true;
			eval("ListForm.RESERVETIME_"+arg0).disabled = true;
			eval("ListForm.RESERVE_CAR_"+arg0).disabled = true;
			eval("ListForm.RESERVE_DEL_"+arg0).disabled = true;
			eval("ListForm.RESERVEHOUR_"+arg0).disabled = true;
			eval("ListForm.RESERVEMIN_"+arg0).disabled = true;
		}
	}

	function fn_dial(arg0,arg1)
	{
		//전화걸기

		if ( arg0 == '1' )
			top.CallStateFrame.document.all.txtCID.value = eval("ListForm.HOMEPHONE_"+arg1).value;
		else
			top.CallStateFrame.document.all.txtCID.value = eval("ListForm.MOBILEPHONE_"+arg1).value;

		if ( top.CallStateFrame.document.all.txtCID.value == "" )
			alert('전화걸기 실패 : 전화번호가 입력되지 않음');
		else
		{
			// 전화걸기
			//DBFrame.location= "/menu01/submenu0101/contactlist.asp?Job=I&idx=0&ControlIdx="+arg1+"&receiptfactnum="+eval("ListForm.RECEIPTFACTNUM").value+"&factPeoplenum="+eval("ListForm.factPeoplenum_"+arg1).value+"&ContactTelNo="+top.CallStateFrame.document.all.txtCID.value;

			top.CallStateFrame.vfn_MakeCall(top.CallStateFrame.document.all.txtCID.value,'');

			top.CallStateFrame.document.all.contactgb1.value = "0";	//군전화,일반전화
			top.CallStateFrame.document.all.contactgb2.value = arg0; //arg0
			top.CallStateFrame.document.all.contactgb3.value = arg1; //arg1
		}
	}

	function fn_contact(arg0,arg1)
	{
		//alert('탄다123');
			DBFrame.location= "/menu01/submenu0101/contactlist.asp?Job=I&idx=0&ControlIdx="+arg1+"&receiptfactnum="+eval("ListForm.RECEIPTFACTNUM").value+"&factPeoplenum="+eval("ListForm.factPeoplenum_"+arg1).value+"&ContactTelNo="+top.CallStateFrame.document.all.txtCID.value+"&CallId="+top.CallStateFrame.document.all.txtCallId.value;

	}

	function fn_dial_1(arg0,arg1)
	{
		//전화걸기

		if ( arg0 == '1' )
			top.CallStateFrame.document.all.txtCID.value = "9"+eval("ListForm.HOMEPHONE_"+arg1).value;
		else
			top.CallStateFrame.document.all.txtCID.value = "9"+eval("ListForm.MOBILEPHONE_"+arg1).value;

		if ( top.CallStateFrame.document.all.txtCID.value == "" )
			alert('전화걸기 실패 : 전화번호가 입력되지 않음');
		else
		{
			//DBFrame.location= "/menu01/submenu0101/contactlist.asp?Job=I&idx=0&ControlIdx="+arg1+"&receiptfactnum="+eval("ListForm.RECEIPTFACTNUM").value+"&factPeoplenum="+eval("ListForm.factPeoplenum_"+arg1).value+"&ContactTelNo="+top.CallStateFrame.document.all.txtCID.value;
			top.CallStateFrame.vfn_MakeCall(top.CallStateFrame.document.all.txtCID.value,'');
			top.CallStateFrame.document.all.contactgb1.value = "1";	//군전화,일반전화
			top.CallStateFrame.document.all.contactgb2.value = arg0; //arg0
			top.CallStateFrame.document.all.contactgb3.value = arg1; //arg1
		}

	}

	function fn_contact_1(arg0,arg1)
	{
		//alert('탄다1234');
			DBFrame.location= "/menu01/submenu0101/contactlist.asp?Job=I&idx=0&ControlIdx="+arg1+"&receiptfactnum="+eval("ListForm.RECEIPTFACTNUM").value+"&factPeoplenum="+eval("ListForm.factPeoplenum_"+arg1).value+"&ContactTelNo="+top.CallStateFrame.document.all.txtCID.value+"&CallId="+top.CallStateFrame.document.all.txtCallId.value;
	}

	function fn_sms(arg0,arg1) {

			if ( arg0 == '1' )
			{	// 1차전화번호
				//sms = window.open("/menu05/submenu0502/sms.asp?cellphone="+eval("ListForm.HOMEPHONE_"+arg0).value,"sms","toolbar=no,status=yes,location=no,width=620,height=500,top=0,left=0,scrollbars=yes,resizable=no");
				//sms.focus();

				ShowPOPLayer("/menu05/submenu0502/sms.asp?cellphone="+eval("ListForm.HOMEPHONE_"+arg0).value,'620','430');	
			}
			else
			{
				//sms = window.open("/menu05/submenu0502/sms.asp?cellphone="+eval("ListForm.MOBILEPHONE_"+arg0).value,"sms","toolbar=no,status=yes,location=no,width=620,height=500,top=0,left=0,scrollbars=yes,resizable=no");
				//sms.focus();

				ShowPOPLayer("/menu05/submenu0502/sms.asp?cellphone="+eval("ListForm.MOBILEPHONE_"+arg0).value,'620','430');		

			}

	}

	function fn_DEL(arg0,arg1)
	{
		eval("ListForm.QUESTIONP_"+arg0+arg1).checked=false;
		eval("ListForm.QUESTION_"+arg0+arg1+"(0)").checked=false;
		eval("ListForm.QUESTION_"+arg0+arg1+"(1)").checked=false;
		eval("ListForm.QUESTION_"+arg0+arg1+"(2)").checked=false;
		eval("ListForm.POINT_"+arg0+arg1).value = "";
		//합구하기

		//eval("ListForm.TOT_"+arg0).value = "";
		//for (i=1; i<10; i++)
		//{
		//	if ( eval("ListForm.POINT_"+arg0+i) != null )
		//		eval("ListForm.TOT_"+arg0).value = Number(eval("ListForm.TOT_"+arg0).value) + Number(eval("ListForm.POINT_"+arg0+i).value);
		//}

		var iTot = 0;
		var iCnt = 0;
		for (i=1; i<10; i++)
		{
			if ( eval("ListForm.POINT_"+arg0+i) != null )
			{
				if ( Number(eval("ListForm.POINT_"+arg0+i).value) > 0 )
				{
					iCnt = iCnt + 1;
					iTot = iTot + Number(eval("ListForm.POINT_"+arg0+i).value);
				}
			}
		}
		
		//var sValue = eval("ListForm.TOT_"+arg0).value;
		var ControlName = "parent.ListForm.TOT_"+arg0;

		DBFrame.location= "/menu01/submenu0101/point_calculation.asp?Tot="+iTot+"&Cnt="+iCnt+"&ControlName="+ControlName+"&Value=";

		if ( eval("ListForm.TOT_"+arg0).value != "" || eval("ListForm.TOT_"+arg0).value != "0" )
			eval("ListForm.MONITORRESULT_"+arg0).value= "9";		
		else
			eval("ListForm.MONITORRESULT_"+arg0).value= "";
			
		fn_ResultSet(arg0);
	
	}


	function fn_list(){

		if (parent.document.all.FRM.value == "submenu01" )
		{
			parent.location.href="/menu01/submenu0101/research01.asp";
		}
		else if (parent.document.all.FRM.value == "submenu02" )
		{
			parent.location.href="/menu01/submenu0102/research02.asp";
		}
		else
		{
			parent.location.href="/menu01/submenu0103/research03.asp";
		}
		
	}

	function fn_ALLDEL(arg0,arg1)
	{

		//합구하기

		//eval("ListForm.TOT_"+arg0).value = "";
		//for (i=1; i<10; i++)
		//{
		//	if ( eval("ListForm.POINT_"+arg0+i) != null )
		//		eval("ListForm.TOT_"+arg0).value = Number(eval("ListForm.TOT_"+arg0).value) + Number(eval("ListForm.POINT_"+arg0+i).value);
		//}



		var iTot = 0;
		var iCnt = 0;
		for (i=1; i<10; i++)
		{
			if ( eval("ListForm.POINT_"+arg0+i) != null )
			{
				eval("ListForm.QUESTIONP_"+arg0+i).checked=false;
				eval("ListForm.QUESTION_"+arg0+i+"(0)").checked=false;
				eval("ListForm.QUESTION_"+arg0+i+"(1)").checked=false;
				eval("ListForm.QUESTION_"+arg0+i+"(2)").checked=false;
				eval("ListForm.POINT_"+arg0+i).value = "";
			}

		}
		
		eval("ListForm.TOT_"+arg0).value = "";
		eval("ListForm.MONITORRESULT_"+arg0).value= "";	
		fn_ResultSet(arg0);
		
	}

	function fn_YES(arg0,arg1,arg2)
	{
		//폼번호
		//항목번호
		//점수
		if ( arg2 == "7" )
		{
			if ( eval("ListForm.QUESTIONP_"+arg0+arg1).checked  )
				eval("ListForm.POINT_"+arg0+arg1).value = "8";
			else
				eval("ListForm.POINT_"+arg0+arg1).value = "7";
		}
		else if ( arg2 == "8" )
		{
			if ( eval("ListForm.QUESTIONP_"+arg0+arg1).checked )
				eval("ListForm.POINT_"+arg0+arg1).value = "9";
			else
				eval("ListForm.POINT_"+arg0+arg1).value = "8";
		}
		else if ( arg2 == "9" )
		{
			if ( eval("ListForm.QUESTIONP_"+arg0+arg1).checked  )
				eval("ListForm.POINT_"+arg0+arg1).value = "10";
			else
				eval("ListForm.POINT_"+arg0+arg1).value = "9";
		}
		else if ( arg2 == "1" )
		{
			if ( eval("ListForm.QUESTION_"+arg0+arg1+"(0)").checked && eval("ListForm.QUESTIONP_"+arg0+arg1).checked )
				eval("ListForm.POINT_"+arg0+arg1).value = "10";
			else if ( eval("ListForm.QUESTION_"+arg0+arg1+"(1)").checked && eval("ListForm.QUESTIONP_"+arg0+arg1).checked )
				eval("ListForm.POINT_"+arg0+arg1).value = "9";
			else if ( eval("ListForm.QUESTION_"+arg0+arg1+"(2)").checked && eval("ListForm.QUESTIONP_"+arg0+arg1).checked )
				eval("ListForm.POINT_"+arg0+arg1).value = "8";
			else if ( eval("ListForm.QUESTIONP_"+arg0+arg1).checked )
				eval("ListForm.POINT_"+arg0+arg1).value = "1";
			else
				eval("ListForm.POINT_"+arg0+arg1).value = "0";
		}
		//합구하기

		//eval("ListForm.TOT_"+arg0).value = "";
		var iTot = 0;
		var iCnt = 0;
		for (i=1; i<10; i++)
		{
			if ( eval("ListForm.POINT_"+arg0+i) != null )
			{
				if ( Number(eval("ListForm.POINT_"+arg0+i).value) > 0 )
				{
					iCnt = iCnt + 1;
					iTot = iTot + Number(eval("ListForm.POINT_"+arg0+i).value);
				}
			}
		}
		
		//var sValue = eval("ListForm.TOT_"+arg0).value;
		var ControlName = "parent.ListForm.TOT_"+arg0;
		DBFrame.location= "/menu01/submenu0101/point_calculation.asp?Tot="+iTot+"&Cnt="+iCnt+"&ControlName="+ControlName+"&Value=";
		eval("ListForm.MONITORRESULT_"+arg0).value= "9";
		fn_ResultSet(arg0);
	}

	function RecDel(arg0,arg1,arg2)
	{

		if ( arg2 == 'N' )
		{

			if (confirm("녹취자료를 삭제하시겠습니까?")) 
			{
				DBFrame.location= "/menu01/submenu0101/contactlist.asp?Job=R&callid="+arg1+"&RecYN="+arg2;
			}
			else
			{
				alert('녹취자료 삭제취소');
			}
		}
		else
		{

				DBFrame.location= "/menu01/submenu0101/contactlist.asp?Job=R&callid="+arg1+"&RecYN="+arg2;
		}
	}

	function HistoryDel(arg0,arg1,arg2,arg3)
	{

		if (confirm("통화내역을 삭제하시겠습니까?")) 
		{
			DBFrame.location= "/menu01/submenu0101/contactlist.asp?Job=D&idx="+arg1+"&ControlIdx="+arg0+"&receiptfactnum="+arg2+"&factPeoplenum="+arg3;
		}
		else
		{
			alert('삭제취소');
		}
	}

	function HistoryUpdate(arg0,arg1,arg2,arg3)
	{
		//수정하기 위해서 ContactHistory Index를 넣는다.
		eval("ListForm.idx_"+arg0).value= arg1;

		DBFrame.location= "/menu01/submenu0101/contactlist.asp?Job=U&idx="+arg1+"&ControlIdx="+arg0+"&receiptfactnum="+arg2+"&factPeoplenum="+arg3;
		//eval("ListForm.idx_"+arg0).focus();
	}


//-->
</script>

<script>
<!--

	function fn_Player(arg0){
		//파일명
		var x,y;
		x = ( screen.width - 300 )/2;
		y = ( screen.height - 200 )/2;

		ShowPOPLayer("/include/WavePlayer.asp?URL="+arg0,'300','200');	
		//window.open("/include/WavePlayer.asp?URL="+arg0,"Player", "toolbar=no,top=100,left=300,width=300,height=200,resize=no,status=yes, scrollbars=no");
	}

//-->
</script>
