<!-- #include virtual="/Include/Top.asp" -->
<!-- #include virtual="/Include/PopLayer.asp" -->
<%
	guboon = request("guboon")
	JUBSEQ = request("JUBSEQ")
	sToday = date()
	SS_Login_Grade = SESSION("SS_Login_Grade")
	InType = request("InType")

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


	where1 = "FromDate=" & FromDate & "&ToDate=" & ToDate & "&whereCD1=" & whereCD1 & "&whereCD2=" & whereCD2 & "&whereCD3=" & whereCD3 & "&whereCD4=" & whereCD4 & "&whereCD5=" & whereCD5 & "&whereCD6=" & whereCD6 & "&whereCD7=" & whereCD7 & "&whereCD8=" & whereCD8 & "&whereCD9=" & whereCD9 & "&whereCD10=" & whereCD10 & "&whereCD11=" & whereCD11 & "&whereCD12=" & whereCD12
	where2 = "curPage=" & curPage & "&" & where1

	




	sql = "select convert(varchar(19),getdate(),121)"
	set Rs = db.execute(sql)
	JUBTIME = rs(0)

	if JUBSEQ = "" then

		guboon = "INS"
		 
	else

		SQL = "	SELECT *, CONVERT(CHAR(19),JUBTIME,121) AS JUBTIME1 FROM TB_CALLHISTORY"
		SQL = SQL & "		WHERE	JUBSEQ = '" & JUBSEQ & "'"

		Set Rs = server.createObject("ADODB.Recordset")
		Rs.open SQL,db
		if rs.eof = false then

			JUBSEQ = rs("JUBSEQ")
			JUBDATE = rs("JUBDATE")
			JUBTIME = rs("JUBTIME1")
			IOFLAG = rs("IOFLAG")
			CUSTNO = rs("CUSTNO")
			TELKIND = rs("TELKIND")
			CUSTNAME = rs("CUSTNAME")
			TELNO = rs("TELNO")
			TELNO2 = rs("TELNO2")
			SEXGB = rs("SEXGB")
			CHANNELGB = rs("CHANNELGB")
			REQUESTERGB = rs("REQUESTERGB")
			CONSULTGB = rs("CONSULTGB")
			CONSULTETCGB = rs("CONSULTETCGB")
			SOSOKGB = rs("SOSOKGB")
			SOSOKETCGB = rs("SOSOKETCGB")
			LEVEL1 = rs("LEVEL1")
			LEVEL2 = rs("LEVEL2")
			ACLASS = rs("ACLASS")	'�������
			BCLASS = rs("BCLASS")
			CCLASS = rs("CCLASS")
			CHANNEL = rs("CHANNEL")
			CALLFLAG = rs("CALLFLAG")	
			CALLKIND = rs("CALLKIND")	'������
			QUESTION = rs("QUESTION")
			REPLY = rs("REPLY")
			RESULTGB = rs("RESULTGB")
			RESERVEDATE = rs("RESERVEDATE")
			RESERVETIME = rs("RESERVETIME")
			PROCESSGB = rs("PROCESSGB")
			CALLID = rs("CALLID")
			RECORDFILE = rs("RECORDFILE")
			INCODE = rs("INCODE")

			SS_LoginNAME = db_GetUserName(INCODE)
		end if
	end if


	if JUBSEQ = "" then

		guboon = "INS"

		LINEKIND = request("LINEKIND")
		TELNO = request("telNo")
		CID = request("telNo")
		CB_SEQ = request("CB_SEQ")
		IOFLAG = "2"
		if LINEKIND = "SIP-DigitalE1" then
			CHANNELGB = "A"
		else
			CHANNELGB = "B"
		end if
		if CB_SEQ <> "" then
			TELKIND = request("TELKIND")
		end if
	end if


	CUSTNO1 = request("CUSTNO")
	if CUSTNO1 <> "" then '���� ������ ���̽�
		'����ȣ�� �ִٸ�.. ����ȣ�� �־��
		SQL = "SELECT * FROM TB_CUSTINFO WHERE CUSTNO = '" & CUSTNO1 &"'"

		Set RSCUST = server.createObject("ADODB.Recordset")
		RSCUST.open SQL,db

		if RSCUST.eof = false then
			CUSTNO = CUSTNO1
			SOSOKGB = RSCUST("SOSOKGB")
			SOSOKETCGB = RSCUST("SOSOKETCGB")
			LEVEL1 = RSCUST("LEVEL1")
			LEVEL2 = RSCUST("LEVEL2")			
			CUSTNAME = RSCUST("NAME")
			TELNO = RSCUST("CELLPHONE")
			TELNO2 = RSCUST("HOMEPHONE")
			SEXGB = RSCUST("SEX")	
		end if
	end if
	CID1 = request("CID")
	if CID1 <> CID THEN
		CID1 = CID
	end if



	if InType = "CALL" then	'��������.

		'response.write TELKIND

		TELKIND=request("DNIS")
		LINEKIND=request("LINEKIND")
		sCID = request("telNo")

		IOFLAG = "1"
		'---------------------------------------
		'��ȣ�� ��ġ�ϴ� ���ִ��� ã��
		'---------------------------------------

		SQL = "select top 1 * from tb_custinfo where ( cellphone = '"&sCID&"' or homephone = '"&sCID&"' or sendphone = '"&sCID&"')"

		set RsCode = db.execute(SQL)
		if RsCode.eof = false then
			CUSTNO = RsCode("CUSTNO")
			SOSOKGB = RsCode("SOSOKGB")
			SOSOKETCGB = RsCode("SOSOKETCGB")
			LEVEL1 = RsCode("LEVEL1")
			LEVEL2 = RsCode("LEVEL2")			
			CUSTNAME = RsCode("NAME")
			TELNO = RsCode("CELLPHONE")
			TELNO2 = RsCode("HOMEPHONE")
			SEXGB = RsCode("SEX")	
		end if

	end if

%>
<table border="0" width="940" cellpadding="0" cellspacing="3" bgcolor="#EFECE5" align="center">
	<tr bgcolor="#FFFFFF">
		<td>		
			<form name="inUpFrm"  method="post" action="/menu04/submenu0401/callhistory_InsUpDel.asp" onsubmit="return fn_inup(this);" style="margin:0">
			<input type="hidden" name="FromDate" value="<%=FromDate%>">
			<input type="hidden" name="ToDate" value="<%=ToDate%>">
			<input type="hidden" name="curPage" value="<%=curPage%>">
			<input type="hidden" name="QueryYN" value="<%=QueryYN%>">
			<input type="hidden" name="whereCD1" value="<%=whereCD1%>">
			<input type="hidden" name="whereCD2" value="<%=whereCD2%>">
			<input type="hidden" name="whereCD3" value="<%=whereCD3%>">
			<input type="hidden" name="whereCD4" value="<%=whereCD4%>">
			<input type="hidden" name="whereCD5" value="<%=whereCD5%>">
			<input type="hidden" name="whereCD6" value="<%=whereCD6%>">
			<input type="hidden" name="whereCD7" value="<%=whereCD7%>">
			<input type="hidden" name="whereCD8" value="<%=whereCD8%>">
			<input type="hidden" name="whereCD9" value="<%=whereCD9%>">
			<input type="hidden" name="whereCD10" value="<%=whereCD10%>">
			<input type="hidden" name="whereCD11" value="<%=whereCD11%>">
			<input type="hidden" name="whereCD12" value="<%=whereCD12%>">
			<input type="hidden" name="JUBSEQ" value="<%=JUBSEQ%>">
			<input type="hidden" name="guboon" value="<%=guboon%>">	
			<input type="hidden" name="LEVEL2" value="<%=LEVEL2%>">	
			<input type="hidden" name="SOSOKETCGB" value="<%=SOSOKETCGB%>">	
			<input type="hidden" name="CONSULTETCGB" value="<%=CONSULTETCGB%>">	
			<input type="hidden" name="CB_SEQ" value="<%=CB_SEQ%>">
			<table width="100%" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#ffffff">
			    <tr height="30">
					<td align="left" bgcolor="#FFFFFF" class="TDCont" colspan="8">&nbsp;<img src="/Images/dot_01.gif" >&nbsp;<b><font color="#ff00ff"></font>����Ͻ�: <input type="text" name="JUBTIME" value="<%=JUBTIME%>" maxlength="20" size="20" onfocus="setFocusColor(this);" onblur="setOutColor(this);" align="absmiddle"> - <%=SS_LoginNAME%></b></td>
				</tr>
			</table>
			<table width="100%" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">
			    <tr>

					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center">��ȭ����</td>
					<td bgcolor="#FFFFFF" width=200 colspan=1><%
							'======= ó������ �ڵ� �������� ==================================================
							SqlCode = "SELECT CODE, CODENAME FROM TB_CODE"
							SqlCode = SqlCode& " WHERE USEYN='Y' AND SYSYN='Y' AND CODEGROUP='Z04'"
							if SS_Login_Grade <> "A" and SS_Login_Grade <> "C" then
								SqlCode = SqlCode & " and CODE = '" & SS_Login_Grade &"'"
							end if
							SqlCode = SqlCode& " ORDER BY CODE"

							set RsCode = db.execute(SqlCode)
						%>
						<select name="TELKIND" size="1" class="ComboFFFCE7">
						<option value="">����</option>
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("CODE")
										CODENAME = RsCode("CODENAME")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &TELKIND& "")%>
							<%
									RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
							%>	</select>				
					</td>

					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center">�����</td>
					<td bgcolor="#FFFFFF" width=200 nowrap>
<%
							'======= ó������ �ڵ� �������� ==================================================
							SqlCode = "SELECT CODE, CODENAME FROM TB_CODE"
							SqlCode = SqlCode& " WHERE USEYN='Y' AND SYSYN='Y' AND CODEGROUP='C01'"
							SqlCode = SqlCode& " ORDER BY CODE"
							set RsCode = db.execute(SqlCode)
						%>
						<select name="CHANNELGB" size="1" class="ComboFFFCE7">
						<option value="">���������</option>
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("CODE")
										CODENAME = RsCode("CODENAME")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &CHANNELGB& "")%>
							<%
									RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
							%>		</select>				
					</td>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center">��ȭ����</td>
					<td bgcolor="#FFFFFF"><input type="radio" name="IOFLAG" value="1" class="none" <% if IOFLAG = "1" or IOFLAG = "" then %>checked<%end if%> >��
						<input type="radio" name="IOFLAG" value="2" class="none" <% if IOFLAG = "2" then %>checked<%end if%>>�ƿ�						<input type="radio" name="IOFLAG" value="" class="none" <% if IOFLAG = "" then %>checked<%end if%>>���þ���
					</td>

				</tr>
				<tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center">��    ��</td>
					<td bgcolor="#FFFFFF"><input type="radio" name="SEXGB" value="1" class="none" <% if SEXGB = "1" or SEXGB = "" then %>checked<%end if%> >��
						<input type="radio" name="SEXGB" value="2" class="none" <% if SEXGB = "2" then %>checked<%end if%>>��
					</td>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center">��    ��</td>
					<td bgcolor="#FFFFFF" ><input type="text" name="CUSTNAME" value="<%=CUSTNAME%>" maxlength="15" size="15" onfocus="setFocusColor(this);" onblur="setOutColor(this);" align="absmiddle" onKeypress="if (event.keyCode==13) {fn_CustSearch();}"><input type="hidden" name="CUSTNO" value="<%=CUSTNO%>" maxlength="16" size="16" onfocus="setFocusColor(this);" onblur="setOutColor(this);" align="absmiddle">
					</td>


					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center">��ȭ�ð�</td>
					<td bgcolor="#FFFFFF" width=200><input type="text" name="CALLTIME1" value="<%=CALLTIME1%>" maxlength="2" size="2" style="border-width:1px ; border-style:solid; text-align:right" onfocus="setFocusColor(this);" onblur="setOutColor(this);" align="absmiddle">&nbsp;:&nbsp;<input type="text" name="CALLTIME2" value="<%=CALLTIME2%>" maxlength="2" size="2" style="border-width:1px ; border-style:solid; text-align:right"  onfocus="setFocusColor(this);" onblur="setOutColor(this);" align="absmiddle">&nbsp;:&nbsp;<input type="text" name="CALLTIME3" value="<%=CALLTIME3%>" style="border-width:1px ; border-style:solid; text-align:right" maxlength="2" size="2" onfocus="setFocusColor(this);" onblur="setOutColor(this);" align="absmiddle">
					</td>
				</tr>

			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center">��    ��</td>
					<td bgcolor="#FFFFFF" nowrap><%
							'======= ó������ �ڵ� �������� ==================================================
							SqlCode = "SELECT CODE, CODENAME FROM TB_CODE"
							SqlCode = SqlCode& " WHERE USEYN='Y' AND SYSYN='Y' AND CODEGROUP='C04'"
							SqlCode = SqlCode& " ORDER BY CODE"
							set RsCode = db.execute(SqlCode)
						%>
						<select name="SOSOKGB" size="1" class="ComboFFFCE7" onChange="fn_SetSosok2();">
						<option value="">����</option>
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("CODE")
										CODENAME = RsCode("CODENAME")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &SOSOKGB& "")%>
							<%
									RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
							%>	</select><iframe src="/menu03/submenu0301/frame_sosok.asp?SOSOKGB=<%=SOSOKGB%>&SOSOKETCGB=<%=SOSOKETCGB%>" scrolling="no" frameborder="0" width=100% height=32 name="frame_sosok"></iframe>
					</td>

					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center">��    ��</td>
					<td bgcolor="#FFFFFF" height=20 nowrap>

<%
							'======= ó������ �ڵ� �������� ==================================================
							SqlCode = "SELECT CODE, CODENAME FROM TB_CODE"
							SqlCode = SqlCode& " WHERE USEYN='Y' AND SYSYN='Y' AND CODEGROUP='C05'"
							SqlCode = SqlCode& " ORDER BY CODE"
							set RsCode = db.execute(SqlCode)
						%>
						<select name="LEVEL1" size="1" class="ComboFFFCE7" onChange="fn_SetLevel2();">
							<Option value ='' selected>��ޱ���</option>
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("CODE")
										CODENAME = RsCode("CODENAME")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &LEVEL1& "")%>
							<%
									RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
							%>
						</select><iframe src="/menu03/submenu0301/frame_level.asp?level=<%=LEVEL1%>&level2=<%=LEVEL2%>" scrolling="no" frameborder="0" width=100% height=32 name="frame_level"></iframe>
					</td>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center">��������</td>
					<td bgcolor="#FFFFFF" width=200><input type="hidden" name="CALLID" value="<%=CALLID%>"><input type="hidden" name="RECFILE" value="<%=RECFILE%>">					
					</td>
				</tr>
				<tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center">����ó1</td>
					<td bgcolor="#FFFFFF" width=200><input type="text" name="TELNO" value="<%=TELNO%>" maxlength="15" size="15" onfocus="setFocusColor(this);" onblur="setOutColor(this);" align="absmiddle">&nbsp<img src="/Images/Cti/icon_tel.gif" style="cursor:hand;" onClick="fn_dial('2','1');" align="absmiddle" title="��ȭ�ɱ�">&nbsp<img src="/Images/Cti/icon_sche.gif" style="cursor:hand;" onClick="fn_sms('2','1');" align="absmiddle" title="��������"></td>

					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center">����ó2</td>
					<td bgcolor="#FFFFFF" width=200><input type="text" name="TELNO2" value="<%=TELNO2%>" maxlength="15" size="15" onfocus="setFocusColor(this);" onblur="setOutColor(this);" align="absmiddle">&nbsp;<img src="/Images/Cti/icon_tel.gif" style="cursor:hand;" onClick="fn_dial('2','2');" align="absmiddle" title="��ȭ�ɱ�">&nbsp<img src="/Images/Cti/icon_sche.gif" style="cursor:hand;" onClick="fn_sms('2','2');" align="absmiddle" title="��������"></td>
						
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center">�߽Ź�ȣ</td>
					<td bgcolor="#FFFFFF" width=200><input type="text" name="CID" value="<%=CID%>" maxlength="16" size="16" onfocus="setFocusColor(this);" onblur="setOutColor(this);" align="absmiddle" readonly>&nbsp;<img src="/Images/Cti/icon_tel.gif" style="cursor:hand;" onClick="fn_dial('2','2');" align="absmiddle" title="��ȭ�ɱ�">&nbsp<img src="/Images/Cti/icon_sche.gif" style="cursor:hand;" onClick="fn_sms('2','2');" align="absmiddle" title="��������"></td>
				</tr>
				<tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center">�������</td>
					<td bgcolor="#FFFFFF">
<%
							'======= ó������ �ڵ� �������� ==================================================
							SqlCode = "SELECT CODE, CODENAME FROM TB_CODE"
							SqlCode = SqlCode& " WHERE USEYN='Y' AND SYSYN='Y' AND CODEGROUP='C00'"
							SqlCode = SqlCode& " ORDER BY CODE"
							set RsCode = db.execute(SqlCode)
						%>
						<select name="ACLASS" size="1" class="ComboFFFCE7">
						<option value="">�����������</option>
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("CODE")
										CODENAME = RsCode("CODENAME")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &ACLASS& "")%>
							<%
									RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
							%>		</select>
					</td>

					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center">�� �� ��</td>
					<td bgcolor="#FFFFFF">						<%
							'======= ó������ �ڵ� �������� ==================================================
							SqlCode = "SELECT CODE, CODENAME FROM TB_CODE"
							SqlCode = SqlCode& " WHERE USEYN='Y' AND SYSYN='Y' AND CODEGROUP='C02'"
							SqlCode = SqlCode& " ORDER BY CODE"
							set RsCode = db.execute(SqlCode)
						%>
						<select name="REQUESTERGB" size="1" class="ComboFFFCE7">
						<option value="">����</option>
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("CODE")
										CODENAME = RsCode("CODENAME")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &REQUESTERGB& "")%>
							<%
									RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
							%>		</select>		
					</td>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center">�������</td>
					<td bgcolor="#FFFFFF" width=200><%
							'======= ó������ �ڵ� �������� ==================================================
							SqlCode = "SELECT CODE, CODENAME FROM TB_CODE"
							SqlCode = SqlCode& " WHERE USEYN='Y' AND SYSYN='Y' AND CODEGROUP='C10'"
							SqlCode = SqlCode& " ORDER BY CODE"
							set RsCode = db.execute(SqlCode)
						%>
						<select name="CALLFLAG" size="1" class="ComboFFFCE7">
						<option value="">����</option>
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("CODE")
										CODENAME = RsCode("CODENAME")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &CALLFLAG& "")%>
							<%
									RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
							%>	</select>
					</td>

				</tr>
			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center">�� �� ��</td>
					<td bgcolor="#FFFFFF" width=200>						<%
							'======= ó������ �ڵ� �������� ==================================================
							SqlCode = "SELECT CODE, CODENAME FROM TB_CODE"
							SqlCode = SqlCode& " WHERE USEYN='Y' AND SYSYN='Y' AND CODEGROUP='C08'"
							SqlCode = SqlCode& " ORDER BY CODE"
							set RsCode = db.execute(SqlCode)
						%>
						<select name="CALLKIND" size="1" class="ComboFFFCE7">
						<option value="">����</option>
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("CODE")
										CODENAME = RsCode("CODENAME")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &CALLKIND& "")%>
							<%
									RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
							%>	</select>
					</td>

					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center">ó�����</td>
					<td bgcolor="#FFFFFF" width=200><%
							'======= ó������ �ڵ� �������� ==================================================
							SqlCode = "SELECT CODE, CODENAME FROM TB_CODE"
							SqlCode = SqlCode& " WHERE USEYN='Y' AND SYSYN='Y' AND CODEGROUP='A02'"
							SqlCode = SqlCode& " ORDER BY CODE"
							set RsCode = db.execute(SqlCode)
						%>
						<select name="PROCESSGB" size="1" class="ComboFFFCE7" onChange="fn_ResultSet('5');">
						<option value="">����</option>
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("CODE")
										CODENAME = RsCode("CODENAME")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &PROCESSGB& "")%>
							<%
									RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
							%>	</select>				
					</td>

							<td bgcolor="#EEF6FF" class="TDCont" align='center' width="100">��㿹���Ͻ�</td>
							<td bgcolor="#FFFFFF"><input readonly value="<%=RESERVEDATE%>" name="RESERVEDATE" type="text" size="10" onfocus="setFocusColor(this);" onchange="fn_settime('5')" onClick="window.open('/Include/Calendar_view.asp?goMonth='+document.ListForm.RESERVEDATE.value+'&firstcode=&FlightCode=&StartEndGu=S&ControlName=opener.document.ListForm.RESERVEDATE','','width=800,height=300,top=200,left=300,scrollbars=no,status=no');" onblur="setOutColor(this);">&nbsp;<img src="/Images/Btn/BtnIconDel.gif" title="����" name="RESERVE_DEL" style="cursor:hand;" align="absmiddle"onclick="document.all.RESERVEDATE.value='';">&nbsp;&nbsp;<input value="<%=RESERVEHOUR%>" name="RESERVEHOUR" type="text" size="2" maxlength="2" onfocus="setFocusColor(this);" >��&nbsp;<input value="<%=RESERVEMIN%>" name="RESERVEMIN" type="text" size="2" maxlength="2" onfocus="setFocusColor(this);">��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp<select name="RESERVETIME" size="1" class="ComboFFFCE7" onchange="fn_settime('5')">
									<Option value ='' selected>�ð�����</option>
									<Option value ='1' >10����</option>
									<Option value ='2' >30����</option>
									<Option value ='3' >1�ð���</option>
									<Option value ='4' >2�ð���</option>
									<Option value ='08' >���� 7��</option>
									<Option value ='08' >���� 8��</option>
									<Option value ='09' >���� 9��</option>
									<Option value ='10' >����10��</option>
									<Option value ='11' >����11��</option>
									<Option value ='12' >����12��</option>
									<Option value ='13' >����13��</option>
									<Option value ='14' >����14��</option>
									<Option value ='15' >����15��</option>
									<Option value ='16' >����16��</option>
									<Option value ='17' >����17��</option>
									<Option value ='18' >����18��</option>
									<Option value ='19' >����19��</option>
									<Option value ='20' >����20��</option>
									<Option value ='21' >����21��</option>
									<Option value ='22' >����22��</option>

								</select>
							</td>


				</tr>
			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center">��㳻��</td>
					<td bgcolor="#FFFFFF" colspan=5 width=850><textarea name="QUESTION" style="width:100%; height:80" wrap="soft" class="TextareaInput"><%=QUESTION%></textarea>			
					</td>
				</tr>
			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center">��ġ����</td>
					<td bgcolor="#FFFFFF" colspan=5 width=850>	<textarea name="REPLY" style="width:100%; height:80" wrap="soft" class="TextareaInput"><%=REPLY%></textarea>			
					</td>
				</tr>
			</table>
			</form>
		</td>
	</tr>
</table>
<table width="940" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
<table width="940" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td align='left'><img src="/Images/Btn/BtnList.gif" style="cursor:hand;" align="absmiddle" onClick="fn_list();"></td><td align="right"><img src="/Images/Btn/BtnASRegi.gif" style="cursor:hand;" class="None" align="absmiddle" onClick="fn_inup();"></td></tr></table>


<table width="940" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
<table border="0" width="940" cellspacing="0" cellpadding="0" width="100%" bgcolor="#CCCCCC" align="center">
	<tr height="25" bgcolor="#F3F3F3" align="center"><td>
<iframe SRC="callhistory_list.asp?CUSTNO=<%=CUSTNO%>" scrolling="yes" frameborder="0" border="0" width="940" height="200" name="IframeHistory"></iframe>
</td>
</tr>
</table>


<%'======= ȭ�ϻ���/�Ŵ����� =======================================================================================%>
<DIV id="hiddenIframe" style="display:none;">
	<iframe SRC="about:blank" scrolling="auto" frameborder="0" border="0" width="100%" height="50" name="hiddenIframe"></iframe>
	<iframe src="about:blank" name="DBFrame" width="0" height="0" frameborder=1 marginheight=0 marginwidth=0 scrolling="no"></iframe>
</DIV>



<script>


	function fn_settime(arg0)
	{

		if ( eval("inUpFrm.RESERVETIME").value == '1' || eval("inUpFrm.RESERVETIME").value == '2' || eval("inUpFrm.RESERVETIME").value == '3' || eval("inUpFrm.RESERVETIME").value == '4' )
		{
			DBFrame.location= "/menu01/submenu0101/time_calculation.asp?DateControlName=parent.inUpFrm.RESERVEDATE&HourControlName=parent.inUpFrm.RESERVEHOUR&MinControlName=parent.inUpFrm.RESERVEMIN&RESERVETIME="+eval("inUpFrm.RESERVETIME").value;
		}
		else
		{
			eval("inUpFrm.RESERVEHOUR").value = eval("inUpFrm.RESERVETIME").value;
			eval("inUpFrm.RESERVEMIN").value = "00";
		}

	}


	function fn_ResultSet(arg0)
	{
		if ( eval("inUpFrm.PROCESSGB").value == "C" )
		{
			eval("inUpFrm.RESERVEDATE").disabled = false;
			eval("inUpFrm.RESERVETIME").disabled = false;
			eval("inUpFrm.RESERVE_DEL").disabled = false;
			eval("inUpFrm.RESERVEHOUR").disabled = false;
			eval("inUpFrm.RESERVEMIN").disabled = false;
			if ( eval("inUpFrm.RESERVEDATE").value == "" )
			{
				eval("inUpFrm.RESERVEDATE").value = "<%=sToday%>";
			}
			eval("inUpFrm.RESERVEDATE").focus();

		}
		else
		{
			eval("inUpFrm.RESERVEDATE").disabled = true;
			eval("inUpFrm.RESERVETIME").disabled = true;
			eval("inUpFrm.RESERVE_DEL").disabled = true;
			eval("inUpFrm.RESERVEHOUR").disabled = true;
			eval("inUpFrm.RESERVEMIN").disabled = true;
		}
	}

	function fn_dial(arg0,arg1)
	{
		//��ȭ�ɱ�
		if ( arg1 == '1' )
			top.CallStateFrame.document.all.txtCID.value = eval("inUpFrm.TELNO").value;
		else
			top.CallStateFrame.document.all.txtCID.value = eval("inUpFrm.TELNO2").value;

		if ( top.CallStateFrame.document.all.txtCID.value == "" )
			alert('��ȭ�ɱ� ���� : ��ȭ��ȣ�� �Էµ��� ����');
		else
			top.CallStateFrame.vfn_MakeCall(top.CallStateFrame.document.all.txtCID.value,'');
	}


	function fn_CustSearch(){
		//���� ���� �ִ����� ã�´�.

		ShowPOPLayer("/Include/PopUp/MemSearch.asp?FRM=etc&JUBSEQ=<%=JUBSEQ%>&CB_SEQ=<%=CB_SEQ%>&SENDPHONE="+eval("inUpFrm.CID").value+"&NAME="+eval("inUpFrm.CUSTNAME").value,'800','430');	

	}

	function fn_list(){location.href="/menu04/submenu0401/calllist.asp?<%=where2%>";}

	function fn_SetLevel2()
	{
		frame_level.location = "/menu04/submenu0401/frame_level.asp?level="+document.all.LEVEL1.value+"&level2=";
	}
	function fn_SetSosok2()
	{
		frame_sosok.location = "/menu04/submenu0401/frame_sosok.asp?SOSOKGB="+document.all.SOSOKGB.value+"&SOSOKETCGB=";
	}
	function fn_SetConsult2()
	{
		frame_consult.location = "/menu04/submenu0401/frame_consult.asp?CONSULTGB="+document.all.CONSULTGB.value+"&CONSULTETCGB=";
	}
	function fn_inup()
	{
		if ( inUpFrm.CHANNELGB.value == '' )
		{
			alert('������� �����Ͻʽÿ�!');
			inUpFrm.CHANNELGB.focus();
			return false;
		}
		if ( inUpFrm.ACLASS.value != 'C' )
		{
			//��� �ʼ��׸��� ������.
			if ( inUpFrm.SOSOKGB.value == '' )
			{
				alert('�Ҽ��� �����Ͻʽÿ�!');
				inUpFrm.SOSOKGB.focus();
				return false;
			}

		}

		inUpFrm.submit();
	}

	function fn_sms(arg0,arg1) {

				if ( arg1 == '1' )
					sCellPhone = eval("inUpFrm.TELNO").value;
				else if ( arg1 == '2' )
					sCellPhone = eval("inUpFrm.TELNO2").value;
				else if ( arg1 == '3' )
					sCellPhone = eval("inUpFrm.CID").value;

				ShowPOPLayer("/menu05/submenu0502/sms.asp?cellphone="+sCellPhone,'620','430');		

	}


</script>
<!-- #include virtual="/Include/Bottom.asp" -->