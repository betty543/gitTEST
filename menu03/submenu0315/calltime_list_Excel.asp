<!-- #include virtual="/Include/Common.asp" -->

<%
	QueryYN = request("QueryYN")
	FromDate = request("FromDate")
	ToDate = request("ToDate")
	curPage = request("curPage")
	whereCD1 = Trim(request("whereCD1")) '����
	whereCD2 = Trim(request("whereCD2")) '�����
	whereCD3 = Trim(request("whereCD3")) '�Ƿ���
	whereCD4 = Trim(request("whereCD4")) '���о�
	whereCD5 = Trim(request("whereCD5")) '�Ҽ�
	whereCD6 = Trim(request("whereCD6")) '��ޱ���
	whereCD7 = Trim(request("whereCD7")) '��ޱ���2
	whereCD8 = Trim(request("whereCD8"))	'����
	whereCD9 = Trim(request("whereCD9"))	'��ȭ��ȣ
	whereCD10 = Trim(request("whereCD10"))	'�Ҽ�
	whereCD11 = Trim(request("whereCD11"))	'ó�����
	whereCD12 = Trim(request("whereCD12"))	'ó�����
	whereCD2 = Replace(whereCD2," ","")
	if FromDate = "" then
		FromDate =  date()
	end if
	if ToDate = "" then
		ToDate = date()
	end If


	Server.ScriptTimeout = 90000
	Response.ContentType = "application/vnd.ms-excel; name='My_Excel'"
	Call Response.AddHeader("Content-Disposition", "attachment; filename=�Ⱓ�����ð���Ȳ.xls")	'�ٷ������ϱ�
	Call Response.AddHeader("Content-Description", "ASP Generated Data")

%>
<%


	'Set Rs = server.createObject("ADODB.Recordset")
	'Rs.open SQL,db


	'3. ���� ����
	sql = " select JUBDATE, INCODE, SUM(CALLTIME) CALLTIME, count(*) CALLCNT from TB_LIFECALLHISTORY "
	sql = sql & "	where   JUBDATE >= '" & FromDate & "'"
	sql = sql & "	AND     JUBDATE <= '" & ToDate & "'"
	If whereCD10 <> "" Then
		sql = sql & "	AND     INCODE = '" & whereCD10 & "'"
	End if
	If whereCD2 <> "" Then
		sql = sql & "	AND     CHANNELGB_B IN ('" & REPLACE(whereCD2,",","','") & "')"
	End if
	sql = sql & "	group by JUBDATE, INCODE"
	sql = sql & "	ORDER by INCODE, JUBDATE "

	'response.write sql

	set Rs = db.execute(sql)




%>


<table width="600" height="10" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
<table width="600" cellspacing="0" align="center" border="1" bordercolor="black" bordercolordark="white" bordercolorlight="black">
	<tr height="25" bgcolor="#EEF6FF" align="center">
		<td class="TDCont" width="20%" align='center'><b>����</b></td>
		<td class="TDCont" width="15%" align='center'><b>��ȣ</b></td>
		<td class="TDCont" width="15%" align='center'><b>����</b></td>
		<td class="TDCont" width="15%" align='center'>����</td>
		<td class="TDCont" width="15%" align='center'>���Ǽ�</td>
		<td class="TDCont" width="20%" align='center'><b>���ð�</b></td>

	</tr>

<%

	DO UNTIL RS.EOF


		db_INCODE	= RS("INCODE")

		i = 0
		db_TotCALLTIME = 0
		db_CALLCNT = 0

		Do Until db_INCODE <> RS("INCODE")
		i = i + 1
		db_JUBDATE	= RS("JUBDATE")

		db_CALLTIME	= RS("CALLTIME")

		db_INCODE	= RS("INCODE")



		IF WEEKDAY(db_JUBDATE)=1 THEN
			JUBDAY="��"
		ELSEIF WEEKDAY(db_JUBDATE)=2 THEN
			JUBDAY="��"
		ELSEIF WEEKDAY(db_JUBDATE)=3 THEN
			JUBDAY="ȭ"
		ELSEIF WEEKDAY(db_JUBDATE)=4 THEN
			JUBDAY="��"
		ELSEIF WEEKDAY(db_JUBDATE)=5 THEN
			JUBDAY="��"
		ELSEIF WEEKDAY(db_JUBDATE)=6 THEN
			JUBDAY="��"
		ELSEIF WEEKDAY(db_JUBDATE)=7 THEN
			JUBDAY="��"
		END If

		lv_Cnt = RS("CALLCNT")
		db_TotCALLTIME = db_TotCALLTIME + db_CALLTIME		
		db_CALLCNT = db_CALLCNT + lv_Cnt

		lv_Hour = Fix(db_CALLTIME / 3600)
		lv_Min = Fix((db_CALLTIME - lv_Hour * 3600) / 60)
		lv_Sec = db_CALLTIME - ((lv_Hour * 3600) + (lv_Min * 60))

		if lv_Hour < 10 then
			lv_Hour = "0" & lv_Hour
		end if
		if lv_Min < 10 then
			lv_Min = "0" & lv_Min
		end if
		if lv_Sec < 10 then
			lv_Sec = "0" & lv_Sec
		end if

%>

		<tr bgcolor="#FFFFFF">
			<td align="center"><%=db_getUserName(db_INCODE)%></td>
			<td align="center"><%=i%></td>
			<td align="center"><%=db_JUBDATE%></td>
			<td align="center"><%=JUBDAY%></td>
			<td align="center"><%=lv_Cnt%></td>
			<td align="center"><%=lv_Hour & ":" & lv_Min & ":" & lv_Sec%></td>

		</tr>
<%
		startRow = startRow - 1
		RS.MOVENEXT
		If rs.eof Then
			Exit do
		End If
		Loop
		
'������



		lv_Hour = Fix(db_TotCALLTIME / 3600)
		lv_Min = Fix((db_TotCALLTIME - lv_Hour * 3600) / 60)
		lv_Sec = db_TotCALLTIME - ((lv_Hour * 3600) + (lv_Min * 60))

		if lv_Hour < 10 then
			lv_Hour = "0" & lv_Hour
		end if
		if lv_Min < 10 then
			lv_Min = "0" & lv_Min
		end if
		if lv_Sec < 10 then
			lv_Sec = "0" & lv_Sec
		end if

%>

		<tr bgcolor="#EEF6FF">
			<td align="center" colspan= '2'>������</td>
			<td align="center"></td>
			<td align="center"></td>
			<td align="center"><%=db_CALLCNT%></td>
			<td align="center"><%=lv_Hour & ":" & lv_Min & ":" & lv_Sec%></td>

		</tr>
<%

	LOOP


%>

</table>

