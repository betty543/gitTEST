<!-- #include virtual="/Include/Top.asp" -->
<%

	SS_LoginID = SESSION("SS_LoginID")
	SS_Login_Secgroup = SESSION("SS_Login_Secgroup")

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

	
	sWhere = "whereCD10="&whereCD10&"&FromDate="&FromDate & "&ToDate="&ToDate&"&whereCD2="&whereCD2


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


<div id='CalendarLayer' style='display:none;'><iframe name='CalendarFrame' src='/include/Calendar.asp' width='172' height='177' border='0' frameborder='0' scrolling='no'></iframe></div>

<!-- #include virtual="/Include/PopLayer.asp" -->

<table border="0" width="1200" cellpadding="0" cellspacing="3" bgcolor="#EFECE5" align="center">
		
	<form method="post" name="inUpFrm" action="<%=Menu_2nd%>" style="margin:0">
	<tr bgcolor="#FFFFFF">
		<td>

			<input type="hidden" name="QueryYN" value="">
			<input type="hidden" name="whereCD7" value="<%=whereCD7%>">
			<table width="100%" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">

			    <tr>
			        <td width="80" bgcolor="#EFEFEF" class="TDCont" align='center'>��ȸ�Ⱓ</td>
			        <td colspan="1" bgcolor="#FFFFFF" >&nbsp;
			        	<input value="<%=FromDate%>" name="FromDate" readonly type="text" size="10" onfocus="setFocusColor(this); new CalendarFrame.Calendar(this);" onblur="setOutColor(this);" />
				    	~
				    		<input value="<%=ToDate%>" name="ToDate" readonly type="text" size="10" onfocus="setFocusColor(this); new CalendarFrame.Calendar(this);" onblur="setOutColor(this);">
			        </td>


					<td bgcolor="#EFEFEF" class="TDCont" width=80 align='center'>����</td>
					<td bgcolor="#FFFFFF"  nowrap>
<%
							'======= ó������ �ڵ� �������� ==================================================
							SqlCode = "SELECT USERID, USERNAME FROM TB_USERINFO"
							SqlCode = SqlCode& " WHERE USEYN='Y'"
							SqlCode = SqlCode& " AND	GRADE='B'" '��������ȭ �׷�
							SqlCode = SqlCode& " AND	SECGROUP = 'A'" '��������ȭ �׷�
							if SS_Login_Secgroup = "A" then
								'���͸�
								'SqlCode = SqlCode& " AND	USERID = '"&SS_LoginID&"'"
							end if
							SqlCode = SqlCode& " ORDER BY USERID"
							set RsCode = db.execute(SqlCode)
						%>
						&nbsp;<select name="whereCD10" size="1" class="ComboFFFCE7">
							<Option value ='' selected>��������</option>
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("USERID")
										CODENAME = RsCode("USERNAME")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &whereCD10& "")%>
							<%
									RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
							%>
						</select>
					</td>

			        <td class="TDCont" bgcolor="#EFEFEF">������� :

<%

'select * from TB_ARMYINFO where ACLASS = 'Q' AND BCLASS IS NOT NULL AND CCLASS IS NULL
							'======= ó������ �ڵ� �������� ==================================================
							SqlCode = "SELECT BCLASS CODE, CLASSNAME CODENAME FROM TB_ARMYINFO"
							SqlCode = SqlCode& " WHERE ACLASS = 'Q' AND BCLASS IS NOT NULL AND CCLASS IS NULL"
							SqlCode = SqlCode& " ORDER BY ACLASS"
							set RsCode = db.execute(SqlCode)

							Do Until rsCode.eof 
								sCode = RsCode("CODE")
								sCodeName = RsCode("CODENAME")
								If InStr(whereCD2,sCode) > 0 then
									sChecked = "checked"
								ElseIf whereCD2 = "" Then
									sChecked = "checked"
								else
									sChecked = ""
								End if
							%>
								<input type="checkbox" name="whereCD2" value="<%=sCode%>" class="none" <%=sChecked%>><%=sCodeName%>
							<%
								rsCode.movenext
							loop
						%>
						&nbsp;				        	
			        </td>

			        <td colspan='2' rowspan="3" bgcolor="#FFFFFF" align="center">
			        	<img src="/Images/Btn/BtnSearch.gif" style="cursor:hand;" onClick="fn_Search();">
			        	<img src="/Images/Btn/BtnExcel.gif" style="cursor:hand;" onClick="fn_save();">
			        </td>
				</tr>

			</table>
			</form>
		</td>
	</tr>
</table>


<table width="1200" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
<table border="0" width="1200" cellspacing="1" cellpadding="0" width="100%" bgcolor="#CCCCCC" align="center">
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




<script>

	function fn_Search()
	{

		inUpFrm.submit();
	}

	function fn_save() {	

		location.href="/menu03/submenu0315/calltime_list_Excel.asp?<%=sWhere%>";
		
	}

</script>



<!-- #include virtual="/Include/Bottom.asp" -->