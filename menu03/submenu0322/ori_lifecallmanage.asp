<!-- #include virtual="/Include/Top.asp" -->
<!-- #include virtual="/Include/PopLayer.asp" -->
<%
	guboon = request("guboon")
	JUBSEQ = request("JUBSEQ")
	InType = request("InType")
	'-------------------------------------------
	'�ݹ鸮��Ʈ���� ȣ��� ��
	'-------------------------------------------

	'InType=CALLBACK&Cate1="&ACLASS&"&Channel=A&CUSTNO="&CUSTNO&"&telNo="&CID&"&Pid="&PID&"&CB_SEQ="&SEQ&"&CALLBACKPHONE="&CALLBANKNO
	'InType=CALLBACK&LINEKIND="&LINEKIND&"&telNo="&CID&"&CB_SEQ="&SEQ


	if InType = "RECORD" then

		LINEKIND=request("LINEKIND")
		sCID = request("telNo")
		IOFLAG = request("IOFLAG")
		'Filename
		'��ȭ�ð�
		CallTIME = request("CallTIME")
		RecFileName = request("FileName")

		IF CallTIME <> "" THEN
			CALLTIME1 = LEFT(CallTIME,2)
			CALLTIME2 = MID(CallTIME,4,2)
			CALLTIME3 = MID(CallTIME,7,2)
		END IF

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



	elseif InType = "CALLBACK" then

		LINEKIND=request("LINEKIND")
		sCID = request("telNo")
		CB_SEQ = request("CB_SEQ")

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

	elseif InType = "CALL" then	'��������.

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


	else
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


		
	end if

	sql = "select convert(varchar(19),getdate(),121)"
	set Rs = db.execute(sql)
	JUBTIME = rs(0)

	if JUBSEQ = "" then

		guboon = "INS"
		LINEKIND = request("LINEKIND")
		TELNO = request("telNo")
		CID = request("telNo")
		CB_SEQ = request("CB_SEQ")
		if InType = "CALL" or InType = "RECORD" then
			'IOFLAG = "2"
		else
			IOFLAG = "2"
		end if
		if LINEKIND = "SIP-DigitalE1" then
			CHANNELGB = "A"
		else
			CHANNELGB = "B"
		end if
		 
	else

		SQL = "	SELECT *, CONVERT(CHAR(19),JUBTIME,121) AS JUBTIME1 FROM TB_LIFECALLHISTORY"
		SQL = SQL & "		WHERE	JUBSEQ = '" & JUBSEQ & "'"

		Set Rs = server.createObject("ADODB.Recordset")
		Rs.open SQL,db
		if rs.eof = false then

			JUBSEQ = rs("JUBSEQ")
			JUBDATE = rs("JUBDATE")
			JUBTIME = rs("JUBTIME1")
			IOFLAG = rs("IOFLAG")
			CUSTNO = rs("CUSTNO")
			CUSTNAME = rs("CUSTNAME")
			TELNO = rs("TELNO")
			TELNO2 = rs("TELNO2")
			SEXGB = TRIM(rs("SEXGB"))
			CHANNELGB = rs("CHANNELGB")
			REQUESTERGB = rs("REQUESTERGB")
			CONSULTGB = rs("CONSULTGB")
			CONSULTETCGB = rs("CONSULTETCGB")
			SOSOKGB = rs("SOSOKGB")
			SOSOKETCGB = rs("SOSOKETCGB")
			SOSOKETCGB2 = rs("SOSOKETCGB2")
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
			REMARK = rs("REMARK")
			RESULTGB = rs("RESULTGB")
			RESERVEDATE = rs("RESERVEDATE")
			RESERVETIME = rs("RESERVETIME")
			PROCESSGB = rs("PROCESSGB")
			CALLID = rs("CALLID")
			RECORDFILE = rs("RECORDFILE")
			INCODE = rs("INCODE")
			EMERYN = rs("EMERYN")
			CB_SEQ = rs("CB_SEQ")
			FILENAME1 = rs("FILENAME")
			REFERJUBSEQ = rs("REFERJUBSEQ")
			REFCNT =  rs("REFCNT")
			CALLTIMEDP = rs("CALLTIMEDP")
			IF CALLTIMEDP <> "" THEN
				CALLTIME1 = LEFT(CALLTIMEDP,2)
				CALLTIME2 = MID(CALLTIMEDP,4,2)
				CALLTIME3 = MID(CALLTIMEDP,7,2)
			END IF
			SS_LoginNAME = db_GetUserName(INCODE)
			
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

	IF len(FILENAME1)>0 THEN
		Filename_Temp = split(FILENAME1,".")
		FileType = FormatFile(Filename_Temp(1))
	END IF

	if FILENAME1 <> "" then
		FILENAME1_url = "<a href='/Upload/Lifecall/Download.asp?filename="&FILENAME1&"'>"&FILENAME1&"</a>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"&FILENAME1&" ����' style='cursor:hand;' align='absmiddle' onClick=FileDel('inUpFrm','"&FILENAME1&"')>&nbsp;"

	end if

	if REFERJUBSEQ <> "" then
		REFERJUBSEQ_URL = "<a href='##'>"&REFERJUBSEQ&"</a>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='"&REFERJUBSEQ&" ����' style='cursor:hand;' align='absmiddle' onClick=ReferDel('inUpFrm','"&REFERJUBSEQ&"')>&nbsp;"
	end if


	IF RecFileName <> "" THEN
		RecFileName_URL = "<a href='##'>"&right(RecFileName,22)&"</a>&nbsp;<img src='/Images/Comm/IconAlert.gif' align='absmiddle' style='cursor:hand;' onClick=fn_Player('"&RecFileName&"'); title='�������� û��'>&nbsp;</a>"		
	END IF

%>
<table border="0" width="940" cellpadding="0" cellspacing="3" bgcolor="#EFECE5" align="center">
	<tr bgcolor="#FFFFFF">
		<td>		
			<form name="inUpFrm"  method="post" action="/menu03/submenu0301/lifecallhistory_InsUpDel.asp" onsubmit="return fn_inup(this);" style="margin:0">
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
			<input type="hidden" name="SOSOKETCGB2" value="<%=SOSOKETCGB2%>">	
			<input type="hidden" name="CONSULTETCGB" value="<%=CONSULTETCGB%>">
			<input type="hidden" name="CB_SEQ" value="<%=CB_SEQ%>">
			<table width="100%" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#ffffff">
			    <tr height="30">
					<td align="left" bgcolor="#FFFFFF" class="TDCont" colspan="8">&nbsp;<img src="/Images/dot_01.gif" >&nbsp;<b><font color="#ff00ff"></font>����Ͻ�: <input type="text" name="JUBTIME" value="<%=JUBTIME%>" maxlength="20" size="20" onfocus="setFocusColor(this);" onblur="setOutColor(this);" align="absmiddle"> - <%=SS_LoginNAME%></b></td>
				</tr>
			</table>
			<table width="100%" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">
			    <tr>
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
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center">�߽Ź�ȣ</td>
					<td bgcolor="#FFFFFF" width=200><input type="text" name="CID" value="<%=CID%>" maxlength="16" size="16" onfocus="setFocusColor(this);" onblur="setOutColor(this);" align="absmiddle" readonly>&nbsp;<img src="/Images/Cti/icon_tel.gif" style="cursor:hand;" onClick="fn_dial('2','2');" align="absmiddle" title="��ȭ�ɱ�">&nbsp<img src="/Images/Cti/icon_tel_1.gif" style="cursor:hand;" onClick="fn_dial_1('2','2');" align="absmiddle" title="�Ϲ���ȭ�� ��ȭ�ɱ�">&nbsp<img src="/Images/Cti/icon_sche.gif" style="cursor:hand;" onClick="fn_sms('2','2');" align="absmiddle" title="��������"></td>
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
					<td bgcolor="#FFFFFF" width=200><input type="text" name="CALLTIME1" value="<%=CALLTIME1%>" maxlength="2" size="2" style="border-width:1px ; border-style:solid; text-align:right" onfocus="setFocusColor(this);" onblur="setOutColor(this);" align="absmiddle">&nbsp;:&nbsp;<input type="text" name="CALLTIME2" value="<%=CALLTIME2%>" maxlength="2" size="2" style="border-width:1px ; border-style:solid; text-align:right"  onfocus="setFocusColor(this);" onblur="setOutColor(this);" align="absmiddle">&nbsp;:&nbsp;<input type="text" name="CALLTIME3" value="<%=CALLTIME3%>" style="border-width:1px ; border-style:solid; text-align:right" maxlength="2" size="2" onfocus="setFocusColor(this);" onblur="setOutColor(this);" align="absmiddle">&nbsp;(��:��:��)
					</td>
				</tr>

			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center">��    ��</td>
					<td bgcolor="#FFFFFF" nowrap><%
							'======= ó������ �ڵ� �������� ==================================================
							SqlCode = "select * from tb_armyinfo where bclass is null and cclass is null order by aclass"
							set RsCode = db.execute(SqlCode)
						%>
						<select name="SOSOKGB" size="1" class="ComboFFFCE7" onChange="fn_SetSosok2();">
						<option value="">����</option>
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("aclass")
										CODENAME = RsCode("classname")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &SOSOKGB& "")%>
							<%
									RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
							%>	</select><iframe src="/menu03/submenu0301/frame_sosok.asp?SOSOKGB=<%=SOSOKGB%>&SOSOKETCGB=<%=SOSOKETCGB%>" scrolling="no" frameborder="0" width=100% height=32 name="frame_sosok"></iframe><iframe src="/menu03/submenu0301/frame_sosok_3.asp?SOSOKGB=<%=SOSOKGB%>&SOSOKETCGB=<%=SOSOKETCGB%>&SOSOKETCGB2=<%=SOSOKETCGB2%>" scrolling="no" frameborder="0" width=100% height=32 name="frame_sosok2"></iframe>
					</td>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center">������ġ����</td>
					<td bgcolor="#FFFFFF" width=200 nowrap><input type="text" name="CounselorYN" value="<%=db_getCateNameCounselorYN_(SOSOKGB,SOSOKETCGB,SOSOKETCGB2)%>" style="border-width:0px ; border-color:#EEF6FF ; border-style:solid; text-align:left; font-color:#ff0000;font-weight:bold" readonly>
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

				</tr>
				<tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center">����ó1</td>
					<td bgcolor="#FFFFFF" width=200><input type="text" name="TELNO" value="<%=TELNO%>" maxlength="15" size="15" onfocus="setFocusColor(this);" onblur="setOutColor(this);" align="absmiddle">&nbsp<img src="/Images/Cti/icon_tel.gif" style="cursor:hand;" onClick="fn_dial('2','1');" align="absmiddle" title="����ȭ�� ��ȭ�ɱ�">&nbsp<img src="/Images/Cti/icon_tel_1.gif" style="cursor:hand;" onClick="fn_dial_1('2','1');" align="absmiddle" title="�Ϲ���ȭ�� ��ȭ�ɱ�">&nbsp<img src="/Images/Cti/icon_sche.gif" style="cursor:hand;" onClick="fn_sms('2','1');" align="absmiddle" title="��������"></td>

					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center">����ó2</td>
					<td bgcolor="#FFFFFF" width=200><input type="text" name="TELNO2" value="<%=TELNO2%>" maxlength="15" size="15" onfocus="setFocusColor(this);" onblur="setOutColor(this);" align="absmiddle">&nbsp;<img src="/Images/Cti/icon_tel.gif" style="cursor:hand;" onClick="fn_dial('2','2');" align="absmiddle" title="����ȭ�� ��ȭ�ɱ�">&nbsp<img src="/Images/Cti/icon_tel_1.gif" style="cursor:hand;" onClick="fn_dial_1('2','2');" align="absmiddle" title="�Ϲ���ȭ�� ��ȭ�ɱ�">&nbsp<img src="/Images/Cti/icon_sche.gif" style="cursor:hand;" onClick="fn_sms('2','2');" align="absmiddle" title="��������"></td>
					
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center">���û��</td>
					<td bgcolor="#FFFFFF" width=200><input type="hidden" name="REFERJUBSEQ" value="<%=REFERJUBSEQ%>" readonly>
								<span id="txtREFERJUBSEQ"><%=REFERJUBSEQ_url%></span><img src="/Images/Btn/BtnRefSrc.GIF" style="cursor:hand;" class="None" align="absmiddle" onClick="ReferUp('A','1');">
					</td>		

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
							%>		</select>&nbsp;&nbsp;&nbsp;&nbsp;<input type="checkbox" name="EMERYN" value="Y" class="none" <% if EMERYN="Y" then Response.Write("checked") end if %>><font color="#ff0000"><b>���</b></font>
					</td>

					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center">���о�</td>
					<td bgcolor="#FFFFFF" nowrap valign="top">						
<%
							'======= ó������ �ڵ� �������� ==================================================
							SqlCode = "SELECT CODE, CODENAME FROM TB_CODE"
							SqlCode = SqlCode& " WHERE USEYN='Y' AND SYSYN='Y' AND CODEGROUP='C03'"
							SqlCode = SqlCode& " ORDER BY CODE"
							set RsCode = db.execute(SqlCode)
						%>
						<select name="CONSULTGB" size="1" class="ComboFFFCE7" onChange="fn_SetConsult2();">
						<option value="">����</option>
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("CODE")
										CODENAME = RsCode("CODENAME")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &CONSULTGB& "")%>
							<%
									RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
							%>		</select><iframe src="/menu03/submenu0301/frame_consult.asp?CONSULTGB=<%=CONSULTGB%>&CONSULTETCGB=<%=CONSULTETCGB%>" scrolling="no" frameborder="0" width=100% height=100% name="frame_consult"></iframe>				
					</td>

					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center">���ȸ��</td>
					<td bgcolor="#FFFFFF" width=200><input type="text" name="REFCNT" value="<%=REFCNT%>" readonly size='2' style="border-width:1px ; border-style:solid; text-align:right; text-forecolor=#0000ff">&nbsp;ȸ
					</td>		

				</tr>
			    <tr>
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
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center">��������</td>
					<td bgcolor="#FFFFFF" width=200><input type="hidden" name="CALLID" value="<%=CALLID%>"><input type="hidden" name="RECFILE" value="<%=RECFILE%>"><span id="txtRECFILE"><%=RecFileName_URL%></span>					
					</td>					

				</tr>
			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center">�� �� ��</td>
					<td bgcolor="#FFFFFF" width=200><%
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

					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center">��ġ���</td>
					<td bgcolor="#FFFFFF" width=200><%
							'======= ó������ �ڵ� �������� ==================================================
							SqlCode = "SELECT CODE, CODENAME FROM TB_CODE"
							SqlCode = SqlCode& " WHERE USEYN='Y' AND SYSYN='Y' AND CODEGROUP='C09'"
							SqlCode = SqlCode& " ORDER BY CODE"
							set RsCode = db.execute(SqlCode)
						%>
						<select name="PROCESSGB" size="1" class="ComboFFFCE7">
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
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center">÷������</td>
					<td bgcolor="#FFFFFF" width=200 nowrap><input type="hidden" name="FILENAME1" value="<%=FILENAME1%>" readonly>
						<span id="txtFILENAME1"><%=FILENAME1_url%></span><img src="/Images/Btn/BtnUpload.gif" style="cursor:hand;" align="absmiddle" onClick="FielUp('A','1');">
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
			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center">Ư�̻���</td>
					<td bgcolor="#FFFFFF" colspan=5 width=850>	<textarea name="REMARK" style="width:100%; height:50" wrap="soft" class="TextareaInput"><%=REMARK%></textarea>			
					</td>
				</tr>
			</table>
			</form>
		</td>
	</tr>
</table>
<table width="940" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
<table width="940" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td align��='left'><img src="/Images/Btn/BtnList.gif" style="cursor:hand;" align="absmiddle" onClick="fn_list();"></td><td align="right"><img src="/Images/Btn/BtnASRegi.gif" style="cursor:hand;" class="None" align="absmiddle" onClick="fn_inup();"></td></tr></table>

<!--
<table width="940" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
<table border="0" width="940" cellspacing="1" cellpadding="0" width="100%" bgcolor="#CCCCCC" align="center">
	<tr height="25" bgcolor="#F3F3F3" align="center">
		<td><b>����</b></td>
		<td><b>����Ͻ�</b></td>
		<td><b>�����</b></td>
		<td><b>�Ҽ�</b></td>
		<td><b>���</b></td>
		<td><b>����</b></td>
		<td><b>����</b></td>
		<td><b>����</b></td>
		<td ><b>���о�</b></td>
		<td><b>����</b></td>
	</tr>
	<tr height="25" bgcolor="#ffffff" align="center">
		<td align="center" colspan=10>���� ����̷��� �������� �ʽ��ϴ�</td>
	</tr>

</table>
-->
<table width="940" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
<table border="0" width="940" cellspacing="0" cellpadding="0" width="100%" bgcolor="#CCCCCC" align="center">
	<tr height="25" bgcolor="#F3F3F3" align="center"><td>
<iframe SRC="lifecallhistory_list.asp?CUSTNO=<%=CUSTNO%>" scrolling="yes" frameborder="0" border="0" width="940" height="200" name="IframeHistory"></iframe>
</td>
</tr>
</table>
<%'======= ȭ�ϻ���/�Ŵ����� =======================================================================================%>
<DIV id="hiddenIframe" style="display:none;">
	<iframe SRC="about:blank" scrolling="auto" frameborder="0" border="0" width="100%" height="50" name="hiddenIframe"></iframe>
</DIV>


<script>


	function fn_sms(arg0,arg1) {

				if ( arg1 == '1' )
					sCellPhone = eval("inUpFrm.TELNO").value;
				else if ( arg1 == '2' )
					sCellPhone = eval("inUpFrm.TELNO2").value;
				else if ( arg1 == '3' )
					sCellPhone = eval("inUpFrm.CID").value;
				//sms = //window.open("/menu05/submenu0502/sms.asp?cellphone="+sCellPhone,"sms","toolbar=no,status=yes,location=no,width=620,height=500,top=0,left=0,scrollbars=yes,resizable=no");
				//sms.focus();


		ShowPOPLayer("/menu05/submenu0502/sms.asp?cellphone="+sCellPhone,'820','430');		
//				sms = window.open("sms.asp","sms","toolbar=no,status=yes,location=no,width=620,height=500,top=0,left=0,scrollbars=yes,resizable=no");
			//	sms.focus();

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

	function fn_dial_1(arg0,arg1)
	{
		//��ȭ�ɱ�

		if ( arg0 == '1' )
			top.CallStateFrame.document.all.txtCID.value = "9"+eval("inUpFrm.TELNO").value;
		else
			top.CallStateFrame.document.all.txtCID.value = "9"+eval("inUpFrm.TELNO2").value;

		if ( top.CallStateFrame.document.all.txtCID.value == "" )
			alert('��ȭ�ɱ� ���� : ��ȭ��ȣ�� �Էµ��� ����');
		else
			top.CallStateFrame.vfn_MakeCall(top.CallStateFrame.document.all.txtCID.value,'');

	}

	function fn_CustSearch(){
		//���� ���� �ִ����� ã�´�.
		//�̸�,cid,��ȭ��ȣ1,2
		ShowPOPLayer("/Include/PopUp/MemSearch.asp?FRM=life&JUBSEQ=<%=JUBSEQ%>&CB_SEQ=<%=CB_SEQ%>&SENDPHONE="+eval("inUpFrm.CID").value+"&NAME="+eval("inUpFrm.CUSTNAME").value,'800','430');		
	}

	function fn_list(){location.href="/menu03/submenu0301/lifecallhistory.asp?<%=where2%>";}

	function fn_SetLevel2()
	{
		frame_level.location = "/menu03/submenu0301/frame_level.asp?level="+document.all.LEVEL1.value+"&level2=";
	}
	function fn_SetSosok2()
	{
		frame_sosok.location = "/menu03/submenu0301/frame_sosok.asp?SOSOKGB="+document.all.SOSOKGB.value+"&SOSOKETCGB=";
		frame_sosok2.location = "/menu03/submenu0301/frame_sosok_3.asp?SOSOKGB="+document.all.SOSOKGB.value+"&SOSOKETCGB=";
		document.all.CounselorYN.value = "";
	}
	function fn_SetSosok3()
	{
		frame_sosok2.location = "/menu03/submenu0301/frame_sosok_3.asp?SOSOKGB="+document.all.SOSOKGB.value+"&SOSOKETCGB=";
		document.all.CounselorYN.value = "";
	}
	function fn_SetConsult2()
	{
		frame_consult.location = "/menu03/submenu0301/frame_consult.asp?CONSULTGB="+document.all.CONSULTGB.value+"&CONSULTETCGB=";
	}
	function fn_inup()
	{
		if ( inUpFrm.CHANNELGB.value == '' )
		{
			alert('������� �����Ͻʽÿ�!');
			inUpFrm.CHANNELGB.focus();
			return false;
		}
		if ( inUpFrm.ACLASS.value == '' )
		{
			alert('��������� �����Ͻʽÿ�!');
			inUpFrm.ACLASS.focus();
			return false;
		}
		if ( inUpFrm.CUSTNAME.value == '' )
		{
			alert('������ �Է��Ͻʽÿ�! ħ�� �Ǵ� �͸��� ���� ��� [�̻�]���� �Է��Ͻʽÿ�.');
			inUpFrm.CUSTNAME.focus();
			return false;
		}
		//if ( inUpFrm.ACLASS.value != 'C' )
		//{
			//��� �ʼ��׸��� ������.
			if ( inUpFrm.SOSOKGB.value == '' )
			{
				alert('�Ҽ��� �����Ͻʽÿ�!');
				inUpFrm.SOSOKGB.focus();
				return false;
			}
			//����(D), ��Ÿ(H) 2���Ҽ� �Է�
			if ( inUpFrm.SOSOKGB.value == 'D' || inUpFrm.SOSOKGB.value == 'H' )
			{
				if ( inUpFrm.SOSOKETCGB.value == '' )
				{
					alert('[����, ��Ÿ]�� �Ҽ� 2���з��� �����Ͻʽÿ�!');
					//
					return false;
				}
			}

		//}
		if ( inUpFrm.CONSULTGB.value  == '' )
		{
				alert('���о߸� �����Ͻʽÿ�!');
				inUpFrm.CONSULTGB.focus();
				return false;			
		}
		if ( inUpFrm.LEVEL1.value  == '' )
		{
				alert('����� �����Ͻʽÿ�!');
				inUpFrm.LEVEL1.focus();
				return false;			
		}
		// ����� ��Ÿ, �̻��� �ƴϸ� ���� �Է� �ؾ���.
		if ( inUpFrm.LEVEL1.value != 'Z' && inUpFrm.LEVEL1.value != 'C' && inUpFrm.LEVEL2.value  == '' )
		{
				alert('���ΰ���� �����Ͻʽÿ�!');
				return false;			
		}

		// ��������� (���,����,���̹�)
		if ( inUpFrm.ACLASS.value == 'A' || inUpFrm.ACLASS.value == 'B' || inUpFrm.ACLASS.value == 'D' )
		{
			if ( inUpFrm.PROCESSGB.value == '' )
			{
					alert('[���,����,���̹�]�� ��� ��ġ����� �ʼ� �Է��׸��Դϴ�!');
					inUpFrm.PROCESSGB.focus();
					return false;			
			}
		}

		inUpFrm.submit();
	}

	function FielUp(ty,cn){
		//alert("TYPE=" +ty+ ", COUNT=" +cn);
		strTemp = eval("inUpFrm.FILENAME1").value;
		//if (strTemp!=""){fileCNT="2"} else {fileCNT="1"}
		fileCNT="1";
		POPLayerURL = "LifecallFileUpload.asp?fileCNT=" +fileCNT+ "&frmTYPE=" +ty+cn;
		ShowPOPLayer(POPLayerURL,'500','160');
	}


	function ReferUp(ty,cn){
		//alert("TYPE=" +ty+ ", COUNT=" +cn);
		strTemp = eval("inUpFrm.FILENAME1").value;
		//if (strTemp!=""){fileCNT="2"} else {fileCNT="1"}
		fileCNT="1";
		POPLayerURL = "lifecallmanage_refersearch.asp?fileCNT=" +fileCNT+ "&frmTYPE=" +ty+cn;
		ShowPOPLayer(POPLayerURL,'1000','500');
	}

	function FileDel(ty,f){
		//alert("TYPE=" +ty+ ", COUNT=" +f);
		if(confirm("�ش� ����Ÿ�� ���� �Ͻðڽ��ϱ�?")) {
			hiddenIframe.location.href="LifecallFileUpload_Del.asp?frmTYPE="+ty+"&fn=" +f;
		}
	}

	function ReferDel(arg0,arg1){
		if(confirm("�ش� ���û���ȣ�� ���� �Ͻðڽ��ϱ�?")) {
			hiddenIframe.location.href="Refer_Del.asp?JUBSEQ="+arg0+"&REFERJUBSEQ="+arg1;
		}		
	}

	function fn_Player(arg0){
		//���ϸ�
		var x,y;
		x = ( screen.width - 300 )/2;
		y = ( screen.height - 200 )/2;

		ShowPOPLayer("/include/WavePlayer.asp?URL="+arg0,'300','200');	
		//window.open("/include/WavePlayer.asp?URL="+arg0,"Player", "toolbar=no,top=100,left=300,width=300,height=200,resize=no,status=yes, scrollbars=no");
	}

</script>
<!-- #include virtual="/Include/Bottom.asp" -->