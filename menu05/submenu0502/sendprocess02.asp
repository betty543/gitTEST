<!-- #include virtual="/Include/Top_PopUp.asp" -->
<%

group_name		= Request.Form("group_name")			' �޴»��(�׷�)
group_name		= Replace(group_name,chr(32),"")	'�޴�������Ʈ�� �迭�� ��´�
group_name		= Replace(group_name,chr(13),"")
group_name		= Replace(group_name,chr(10),"+")  
group_name		= group_name&"+isend+"
group_name_arr	= split((group_name),"+") 
group_name_cnt  = Request.form("group_name_cnt")	'�׷����� ���� ���� |=|

to_message		= Trim(Request.Form("to_message"))		' �޼��� ����
'to_message		= Replace(to_message, "'", "''")
to_message		= Replace(to_message, vbcrlf, "")

to_message2		= Trim(Request.Form("to_message2"))		' �޼��� ����
to_message2		= Replace(to_message2, "'", "''")
to_message2		= Replace(to_message2, vbcrlf, "")
to_message3		= Trim(Request.Form("to_message3"))		' �޼��� ����
to_message3		= Replace(to_message3, "'", "''")
to_message3		= Replace(to_message3, vbcrlf, "")
to_message4		= Trim(Request.Form("to_message4"))		' �޼��� ����
to_message4		= Replace(to_message4, "'", "''")
to_message4		= Replace(to_message4, vbcrlf, "")
to_message5		= Trim(Request.Form("to_message5"))		' �޼��� ����
to_message5		= Replace(to_message5, "'", "''")
to_message5		= Replace(to_message5, vbcrlf, "")

to_message6		= Trim(Request.Form("to_message6"))		' �޼��� ����
to_message6		= Replace(to_message6, "'", "''")
to_message6		= Replace(to_message6, vbcrlf, "")
to_message7		= Trim(Request.Form("to_message7"))		' �޼��� ����
to_message7		= Replace(to_message7, "'", "''")
to_message7		= Replace(to_message7, vbcrlf, "")
to_message8		= Trim(Request.Form("to_message8"))		' �޼��� ����
to_message8		= Replace(to_message8, "'", "''")
to_message8		= Replace(to_message8, vbcrlf, "")

from_num		= Request.Form("from_num")				' �������

sendType	= Request.Form("sendType")				' ���౸��/ 0: ������� / 6 :  ���� /     �������� 
yy			= Request.Form("yy")					' ���೯¥ �⵵
mm			= Request.Form("mm")					' ���೯¥ ��
dd			= Request.Form("dd")					' ���೯¥ ��
h			= Request.Form("h")						' ����ð� ��
m			= Request.Form("m")						' ����ð� ��
sendtype2	= Request.Form("sendtype2")				' �������� üũ�ڽ�(true/false)
stt			= Request.Form("stt")					' �������� �ð�
smm			= Request.Form("smm")					' �������� ��
snum		= Request.Form("snum")					' �������� �Ǽ�


if cint(mm) < 10 then
	mm = "0" & CInt(mm)
end If
if cint(dd) < 10 then
	dd = "0" & CInt(dd)
end If
if cint(h) < 10 then
	h = "0" & CInt(h)
end If
if cint(m) < 10 then
	m = "0" & CInt(m)
end if


if sendType = "0" then '�������
	sYear = Year(now)
	sMonth = Month(now)
	sDay   = Day(now)
	if cint(sMonth) < 10 then
		sMonth = "0" & sMonth
	end if
	if cint(sDay) < 10 then
		sDay = "0" & sDay
	end if
	sHour = hour(Now)
	if cint(sHour) < 10 then
		sHour = "0" & sHour
	end if
	sMin = minute(Now)
	if cint(sMin) < 10 then
		sMin = "0" & sMin
	end if
	sDate = sYear &"-"& sMonth &"-"&sDay &" "& sHour &":"& sMin &":00"
	stime = sHour & sMin &"00"
	send_flag = "0"
	sDate1 = sYear &"-"& sMonth &"-"& sDay &" "& sHour &":"& sMin &":00"
else								' �������� : 6
	sYear			= yy			' ���೯¥ �⵵
	sMonth			= mm			' ���೯¥ ��
	sDay			= dd			' ���೯¥ ��
	sHour			= h				' ����ð� ��
	sMin			= m				' ����ð� ��
	sDate = sYear &"-"& sMonth &"-"&sDay &" "& sHour &":"& sMin &":00"
	stime = sHour & sMin & "00"
	send_flag = "2"
	sDate1 = sYear &"-"& sMonth &"-"& sDay &" "& sHour &":"& sMin &":00"
end if

if sendtype2 = "on" then		' ��������
	stt			= stt			' �������� �ð�
	smm			= smm			' �������� ��
	if not(stt = 0) then
		stt = stt * 3600
	end if
	if not(smm = 0) then
		smm = smm * 60
	end if
	vge = stt + smm
	snum		= clng(snum)	' �������� �Ǽ�
	ssnum		= snum
	send_flag = "2"				' ���������� �������۰� ����
	sendType  = "6"				' ���������� ������������ �ٲ�(�αױ�Ͻ�..���)
end if

	vDate = sYear &""& sMonth &""& sDay &""& sHour &""& sMin &"00"


	mem_count = 10000

	j=0
	int totcnt = 0


'----------------------------------------------------------------------
'SMS ����
'----------------------------------------------------------------------

		'--------------------------------
		'80����Ʈ�� ¥����
		'--------------------------------
		j=0
		startpoint = 1
		for i = 2 to len(to_message)		

			strsql = " select datalength('" & mid(to_message,startpoint,i-startpoint+1) & "')"
			Set rs = DB.Execute(strsql)			

			if rs(0) >= 81 then
				j=j+1
				if j = 1 then
					to_message2 = mid(to_message,startpoint,i-1-startpoint+1)
					startpoint = i
				elseif j = 2 then
					to_message3 = mid(to_message,startpoint,i-1-startpoint+1)
					startpoint = i
				elseif j = 3 then
					to_message4 = mid(to_message,startpoint,i-1-startpoint+1)
					startpoint = i
				elseif j = 4 then
					to_message5 = mid(to_message,startpoint,i-1-startpoint+1)
					startpoint = i
				elseif j = 5 then
					to_message6 = mid(to_message,startpoint,i-1-startpoint+1)
					startpoint = i
				elseif j = 6 then
					to_message7 = mid(to_message,startpoint,i-1-startpoint+1)
					startpoint = i
				elseif j = 7 then
					to_message8 = mid(to_message,startpoint,i-1-startpoint+1)
					startpoint = i
				end if
			elseif i = len(to_message) then
				j=j+1
				if j = 1 then
					to_message2 = mid(to_message,startpoint,80)
					startpoint = i
				elseif j = 2 then
					to_message3 = mid(to_message,startpoint,80)
					startpoint = i
				elseif j = 3 then
					to_message4 = mid(to_message,startpoint,80)
					startpoint = i
				elseif j = 4 then
					to_message5 = mid(to_message,startpoint,80)
					startpoint = i
				elseif j = 5 then
					to_message6 = mid(to_message,startpoint,80)
					startpoint = i
				elseif j = 6 then
					to_message7 = mid(to_message,startpoint,80)
					startpoint = i
				elseif j = 7 then
					to_message8 = mid(to_message,startpoint,80)
					startpoint = i
				end if
			end if
			rs.close

		next 

j = 0
INCODE = SESSION("SS_LoginID")
SS_Login_Grade = SESSION("SS_Login_Grade")


					strSQL = " SELECT * FROM temp_conference WHERE userid = '"& INCODE & "' and datagb = '2'"

					Set rs = DB.Execute(strSQL)

					If Not(rs.eof Or rs.bof) then

						Do While Not RS.EOF 

								send_date = Mid(sdate,1,10)
								send_time = Mid(sdate,12,5)	
								cellphone = rs("cellphone")
								if sendType = "6" then

									SQLt = "INSERT INTO sms.dbo.SMS_Reserve (Sm_InDate, Sm_SdMbNo, Sm_RvMbNo, Sm_Msg, Sm_Code1, Sm_Code2) Values ('"&sDate1&"', '"&cellphone&"','"&from_num&"','"&to_message2&"','"&INCODE&"','"&SS_Login_Grade&"')"


									'response.write "1 --> " & sqlt
									DB.Execute(SQLt)

									if to_message3 <> "" then

												SQLt = "INSERT INTO sms.dbo.SMS_Reserve (Sm_InDate, Sm_SdMbNo, Sm_RvMbNo, Sm_Msg, Sm_Code1, Sm_Code2) Values ('"&sDate1&"', '"&cellphone&"','"&from_num&"','"&to_message3&"','"&INCODE&"','"&SS_Login_Grade&"')"


										'response.write "1 --> " & sqlt
										DB.Execute(SQLt)
									end if


								else

									SQLt = "INSERT INTO sms.dbo.SMS (Sm_InDate, Sm_SdMbNo, Sm_RvMbNo, Sm_Msg, Sm_Code1, Sm_Code2) Values ('"&sDate1&"', '"&cellphone&"','"&from_num&"','"&to_message2&"','"&INCODE&"','"&SS_Login_Grade&"')"


									'response.write "1 --> " & sqlt
									DB.Execute(SQLt)

									if to_message3 <> "" then

												SQLt = "INSERT INTO sms.dbo.SMS (Sm_InDate, Sm_SdMbNo, Sm_RvMbNo, Sm_Msg, Sm_Code1, Sm_Code2) Values ('"&sDate1&"', '"&cellphone&"','"&from_num&"','"&to_message3&"','"&INCODE&"','"&SS_Login_Grade&"')"


										'response.write "1 --> " & sqlt
										DB.Execute(SQLt)
									end if

								end if

								sqlInsert = "Delete temp_conference where idx = "&rs("idx")&""		
								db.Execute sqlInsert
							rs.MoveNext 
							j = j + 1
						Loop
								'DB.Execute(SQL)
								if sendType = "6" then
									response.write"<script>"
									response.write"alert('���� �߼� �Ϸ�');"
									response.write"self.location.href='sms.asp'"
									response.write"</script>"
								else
									response.write"<script>"
									response.write"parent.location.href='smssend.asp';parent.HddnPOPLayer();"
									response.write"alert('�߼��� �Ϸ�Ǿ����ϴ�');"
									response.write"</script>"

								end if
							response.End
					Else
								response.write"<script>"
								response.write"alert('������ ȣ���ȣ�� �����ϴ�.');"
								response.write"self.location.href='sms.asp';"
								response.write"</script>"
								response.end	
					End If 
					rs.close
					set rs = nothing
%>



<!-- #include virtual="/Include/Bottom_PopUp.asp" -->