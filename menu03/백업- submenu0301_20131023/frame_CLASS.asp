<!-- #include virtual="/Include/Top_Frame.asp" -->

<script>
<!--

// iframe 사이즈 적용
function fn_SetSosok3(CLASSNM,CLASSGB,ACLASS,BCLASS,CCLASS,DCLASS,ECLASS,CLASSNAME,CounselorYN)
{

	if ( CLASSNM == 'SOSOK' )
	{
		if ( CLASSGB == 'A' )
		{
			parent.document.all.CounselorYN.value = CounselorYN;
			parent.document.all.SOSOKGB_A.value = ACLASS;
			parent.document.all.SOSOKGB_B.value = "";
			parent.document.all.SOSOKGB_C.value = "";
			parent.document.all.SOSOKGB_D.value = "";
			parent.document.all.SOSOKGB_E.value = "";
			parent.frame_sosok_B.location = "/menu03/submenu0301/frame_CLASS.asp?frame_nm=frame_sosok_B&CLASSNM=SOSOK&CLASSGB=B&ACLASS="+ACLASS+"&BCLASS="+BCLASS+"&CCLASS="+CCLASS+"&DCLASS="+DCLASS+"&ECLASS="+ECLASS;
			parent.frame_sosok_C.location = "/menu03/submenu0301/frame_CLASS.asp?frame_nm=frame_sosok_C&CLASSNM=SOSOK&CLASSGB=C&ACLASS="+ACLASS+"&BCLASS="+BCLASS+"&CCLASS="+CCLASS+"&DCLASS="+DCLASS+"&ECLASS="+ECLASS;
			parent.frame_sosok_D.location = "/menu03/submenu0301/frame_CLASS.asp?frame_nm=frame_sosok_D&CLASSNM=SOSOK&CLASSGB=D&ACLASS="+ACLASS+"&BCLASS="+BCLASS+"&CCLASS="+CCLASS+"&DCLASS="+DCLASS+"&ECLASS="+ECLASS;
			parent.frame_sosok_E.location = "/menu03/submenu0301/frame_CLASS.asp?frame_nm=frame_sosok_E&CLASSNM=SOSOK&CLASSGB=E&ACLASS="+ACLASS+"&BCLASS="+BCLASS+"&CCLASS="+CCLASS+"&DCLASS="+DCLASS+"&ECLASS="+ECLASS;
		}
		else if ( CLASSGB == 'B' )
		{
			parent.document.all.CounselorYN.value = CounselorYN;
			parent.document.all.SOSOKGB_B.value = BCLASS;
			parent.document.all.SOSOKGB_C.value = "";
			parent.document.all.SOSOKGB_D.value = "";
			parent.document.all.SOSOKGB_E.value = "";
			parent.frame_sosok_C.location = "/menu03/submenu0301/frame_CLASS.asp?frame_nm=frame_sosok_C&CLASSNM=SOSOK&CLASSGB=C&ACLASS="+ACLASS+"&BCLASS="+BCLASS+"&CCLASS="+CCLASS+"&DCLASS="+DCLASS+"&ECLASS="+ECLASS;
			parent.frame_sosok_D.location = "/menu03/submenu0301/frame_CLASS.asp?frame_nm=frame_sosok_D&CLASSNM=SOSOK&CLASSGB=D&ACLASS="+ACLASS+"&BCLASS="+BCLASS+"&CCLASS="+CCLASS+"&DCLASS="+DCLASS+"&ECLASS="+ECLASS;
			parent.frame_sosok_E.location = "/menu03/submenu0301/frame_CLASS.asp?frame_nm=frame_sosok_E&CLASSNM=SOSOK&CLASSGB=E&ACLASS="+ACLASS+"&BCLASS="+BCLASS+"&CCLASS="+CCLASS+"&DCLASS="+DCLASS+"&ECLASS="+ECLASS;
		}
		else if ( CLASSGB == 'C' )
		{
			parent.document.all.CounselorYN.value = CounselorYN;
			parent.document.all.SOSOKGB_C.value = CCLASS;
			parent.document.all.SOSOKGB_D.value = "";
			parent.document.all.SOSOKGB_E.value = "";
			parent.frame_sosok_D.location = "/menu03/submenu0301/frame_CLASS.asp?frame_nm=frame_sosok_D&CLASSNM=SOSOK&CLASSGB=D&ACLASS="+ACLASS+"&BCLASS="+BCLASS+"&CCLASS="+CCLASS+"&DCLASS="+DCLASS+"&ECLASS="+ECLASS;	
			parent.frame_sosok_E.location = "/menu03/submenu0301/frame_CLASS.asp?frame_nm=frame_sosok_E&CLASSNM=SOSOK&CLASSGB=E&ACLASS="+ACLASS+"&BCLASS="+BCLASS+"&CCLASS="+CCLASS+"&DCLASS="+DCLASS+"&ECLASS="+ECLASS;
		}
		else if ( CLASSGB == 'D' )
		{
			parent.document.all.CounselorYN.value = CounselorYN;
			parent.document.all.SOSOKGB_D.value = DCLASS;
			parent.document.all.SOSOKGB_E.value = "";
			parent.frame_sosok_E.location = "/menu03/submenu0301/frame_CLASS.asp?frame_nm=frame_sosok_E&CLASSNM=SOSOK&CLASSGB=E&ACLASS="+ACLASS+"&BCLASS="+BCLASS+"&CCLASS="+CCLASS+"&DCLASS="+DCLASS+"&ECLASS="+ECLASS;		
		}
		else if ( CLASSGB == 'E' )
		{
			parent.document.all.CounselorYN.value = CounselorYN;
			parent.document.all.SOSOKGB_E.value = ECLASS;
	
		}
	}
	if ( CLASSNM == 'LEVEL' )
	{
		if ( CLASSGB == 'A' )
		{
			parent.frame_level_B.location = "/menu03/submenu0301/frame_CLASS.asp?frame_nm=frame_level_B&CLASSNM=LEVEL&CLASSGB=B&ACLASS="+ACLASS+"&BCLASS="+BCLASS+"&CCLASS="+CCLASS+"&DCLASS="+DCLASS+"&ECLASS="+ECLASS;
		}
		if ( CLASSGB == 'B' )
		{
		
			parent.document.all.LEVEL_B.value = BCLASS;
			parent.document.all.LEVEL_C.value = "";
			parent.document.all.LEVEL_D.value = "";
			parent.frame_level_C.location = "/menu03/submenu0301/frame_CLASS.asp?frame_nm=frame_level_C&CLASSNM=LEVEL&CLASSGB=C&ACLASS="+ACLASS+"&BCLASS="+BCLASS+"&CCLASS="+CCLASS+"&DCLASS="+DCLASS+"&ECLASS="+ECLASS;
		}
		if ( CLASSGB == 'C' )
		{

			parent.document.all.LEVEL_C.value = CCLASS;
			parent.document.all.LEVEL_D.value = "";
			parent.frame_level_D.location = "/menu03/submenu0301/frame_CLASS.asp?frame_nm=frame_level_D&CLASSNM=LEVEL&CLASSGB=D&ACLASS="+ACLASS+"&BCLASS="+BCLASS+"&CCLASS="+CCLASS+"&DCLASS="+DCLASS+"&ECLASS="+ECLASS;
		}
		if ( CLASSGB == 'D' )
		{

			parent.document.all.LEVEL_D.value = DCLASS;
		}
	}

	if ( CLASSNM == 'CHANNELGB' )
	{
		if ( CLASSGB == 'A' )
		{
			parent.frame_channelgb_B.location = "/menu03/submenu0301/frame_CLASS.asp?frame_nm=frame_channelgb_B&CLASSNM=CHANNELGB&CLASSGB=B&ACLASS="+ACLASS+"&BCLASS="+BCLASS+"&CCLASS="+CCLASS+"&DCLASS="+DCLASS+"&ECLASS="+ECLASS;
		}
		if ( CLASSGB == 'B' )
		{
			parent.document.all.CHANNELGB_B.value = BCLASS;
			parent.document.all.CHANNELGB_C.value = "";
			parent.frame_channelgb_C.location = "/menu03/submenu0301/frame_CLASS.asp?frame_nm=frame_channelgb_C&CLASSNM=CHANNELGB&CLASSGB=C&ACLASS="+ACLASS+"&BCLASS="+BCLASS+"&CCLASS="+CCLASS+"&DCLASS="+DCLASS+"&ECLASS="+ECLASS;
		}
		if ( CLASSGB == 'C' )
		{
			parent.document.all.CHANNELGB_C.value = CCLASS;
		}

		parent.fn_chkCHANNELGB_B();

	}
	if ( CLASSNM == 'CALLCLASS' )
	{

		if ( CLASSGB == 'A' )
		{
			parent.frame_callclass_B.location = "/menu03/submenu0301/frame_CLASS.asp?frame_nm=frame_callclass_B&CLASSNM=CALLCLASS&CLASSGB=B&ACLASS="+ACLASS+"&BCLASS="+BCLASS+"&CCLASS="+CCLASS+"&DCLASS="+DCLASS+"&ECLASS="+ECLASS;
		}
		if ( CLASSGB == 'B' )
		{
			parent.document.all.CALLCLASS_B.value = BCLASS;
			parent.document.all.CALLCLASS_C.value = "";
			parent.frame_callclass_C.location = "/menu03/submenu0301/frame_CLASS.asp?frame_nm=frame_callclass_C&CLASSNM=CALLCLASS&CLASSGB=C&ACLASS="+ACLASS+"&BCLASS="+BCLASS+"&CCLASS="+CCLASS+"&DCLASS="+DCLASS+"&ECLASS="+ECLASS;
		}
		if ( CLASSGB == 'C' )
		{
			parent.document.all.CALLCLASS_C.value = CCLASS;
		}
	}
	if ( CLASSNM == 'CALLCLASS_2' )
	{

		if ( CLASSGB == 'A' )
		{
			parent.frame_callclass_B_2.location = "/menu03/submenu0301/frame_CLASS.asp?frame_nm=frame_callclass_B_2&CLASSNM=CALLCLASS&CLASSGB=B&ACLASS="+ACLASS+"&BCLASS="+BCLASS+"&CCLASS="+CCLASS+"&DCLASS="+DCLASS+"&ECLASS="+ECLASS;
		}
		if ( CLASSGB == 'B' )
		{
			parent.document.all.CALLCLASS_B_2.value = BCLASS;
			parent.document.all.CALLCLASS_C_2.value = "";
			parent.frame_callclass_C_2.location = "/menu03/submenu0301/frame_CLASS.asp?frame_nm=frame_callclass_C_2&CLASSNM=CALLCLASS&CLASSGB=C&ACLASS="+ACLASS+"&BCLASS="+BCLASS+"&CCLASS="+CCLASS+"&DCLASS="+DCLASS+"&ECLASS="+ECLASS;
		}
		if ( CLASSGB == 'C' )
		{
			parent.document.all.CALLCLASS_C_2.value = CCLASS;
		}
	}
	if ( CLASSNM == 'CALLCLASS_3' )
	{

		if ( CLASSGB == 'A' )
		{
			parent.frame_callclass_B_3.location = "/menu03/submenu0301/frame_CLASS.asp?frame_nm=frame_callclass_B_3&CLASSNM=CALLCLASS&CLASSGB=B&ACLASS="+ACLASS+"&BCLASS="+BCLASS+"&CCLASS="+CCLASS+"&DCLASS="+DCLASS+"&ECLASS="+ECLASS;
		}
		if ( CLASSGB == 'B' )
		{
			parent.document.all.CALLCLASS_B_3.value = BCLASS;
			parent.document.all.CALLCLASS_C_3.value = "";
			parent.frame_callclass_C_3.location = "/menu03/submenu0301/frame_CLASS.asp?frame_nm=frame_callclass_C_3&CLASSNM=CALLCLASS&CLASSGB=C&ACLASS="+ACLASS+"&BCLASS="+BCLASS+"&CCLASS="+CCLASS+"&DCLASS="+DCLASS+"&ECLASS="+ECLASS;
		}
		if ( CLASSGB == 'C' )
		{
			parent.document.all.CALLCLASS_C_3.value = CCLASS;
		}
	}

	if ( CLASSNM == 'CALLCLASS_4' )
	{

		if ( CLASSGB == 'A' )
		{
			parent.frame_callclass_B_4.location = "/menu03/submenu0301/frame_CLASS.asp?frame_nm=frame_callclass_B_4&CLASSNM=CALLCLASS&CLASSGB=B&ACLASS="+ACLASS+"&BCLASS="+BCLASS+"&CCLASS="+CCLASS+"&DCLASS="+DCLASS+"&ECLASS="+ECLASS;
		}
		if ( CLASSGB == 'B' )
		{
			parent.document.all.CALLCLASS_B_4.value = BCLASS;
			parent.document.all.CALLCLASS_C_4.value = "";
			parent.frame_callclass_C_4.location = "/menu03/submenu0301/frame_CLASS.asp?frame_nm=frame_callclass_C_4&CLASSNM=CALLCLASS&CLASSGB=C&ACLASS="+ACLASS+"&BCLASS="+BCLASS+"&CCLASS="+CCLASS+"&DCLASS="+DCLASS+"&ECLASS="+ECLASS;
		}
		if ( CLASSGB == 'C' )
		{
			parent.document.all.CALLCLASS_C_4.value = CCLASS;
		}
	}



	if ( CLASSNM == 'CALLCLASS_5' )
	{

		if ( CLASSGB == 'A' )
		{
			parent.frame_callclass_B_5.location = "/menu03/submenu0301/frame_CLASS.asp?frame_nm=frame_callclass_B_5&CLASSNM=CALLCLASS&CLASSGB=B&ACLASS="+ACLASS+"&BCLASS="+BCLASS+"&CCLASS="+CCLASS+"&DCLASS="+DCLASS+"&ECLASS="+ECLASS;
		}
		if ( CLASSGB == 'B' )
		{
			parent.document.all.CALLCLASS_B_5.value = BCLASS;
			parent.document.all.CALLCLASS_C_5.value = "";
			parent.frame_callclass_C_5.location = "/menu03/submenu0301/frame_CLASS.asp?frame_nm=frame_callclass_C_5&CLASSNM=CALLCLASS&CLASSGB=C&ACLASS="+ACLASS+"&BCLASS="+BCLASS+"&CCLASS="+CCLASS+"&DCLASS="+DCLASS+"&ECLASS="+ECLASS;
		}
		if ( CLASSGB == 'C' )
		{
			parent.document.all.CALLCLASS_C_5.value = CCLASS;
		}
	}


	if ( CLASSNM == 'CALLKIND' )
	{

		if ( CLASSGB == 'A' )
		{
			parent.frame_callkind_B.location = "/menu03/submenu0301/frame_CLASS.asp?frame_nm=frame_callkind_B&CLASSNM=CALLKIND&CLASSGB=B&ACLASS="+ACLASS+"&BCLASS="+BCLASS+"&CCLASS="+CCLASS+"&DCLASS="+DCLASS+"&ECLASS="+ECLASS;
		}
		if ( CLASSGB == 'B' )
		{
			parent.document.all.CALLKIND_B.value = BCLASS;
			parent.document.all.CALLKIND_C.value = "";
			parent.frame_callkind_C.location = "/menu03/submenu0301/frame_CLASS.asp?frame_nm=frame_callkind_C&CLASSNM=CALLKIND&CLASSGB=C&ACLASS="+ACLASS+"&BCLASS="+BCLASS+"&CCLASS="+CCLASS+"&DCLASS="+DCLASS+"&ECLASS="+ECLASS;
		}
		if ( CLASSGB == 'C' )
		{
			parent.document.all.CALLKIND_C.value = CCLASS;
		}
	}
	if ( CLASSNM == 'CALLKIND_2' )
	{

		if ( CLASSGB == 'A' )
		{
			parent.frame_callkind_B_2.location = "/menu03/submenu0301/frame_CLASS.asp?frame_nm=frame_callkind_B_2&CLASSNM=CALLKIND_2&CLASSGB=B&ACLASS="+ACLASS+"&BCLASS="+BCLASS+"&CCLASS="+CCLASS+"&DCLASS="+DCLASS+"&ECLASS="+ECLASS;

		}
		if ( CLASSGB == 'B' )
		{
			parent.document.all.CALLKIND_B_2.value = BCLASS;
			parent.document.all.CALLKIND_C_2.value = "";
			parent.frame_callkind_C_2.location = "/menu03/submenu0301/frame_CLASS.asp?frame_nm=frame_callkind_C_2&CLASSNM=CALLKIND_2&CLASSGB=C&ACLASS="+ACLASS+"&BCLASS="+BCLASS+"&CCLASS="+CCLASS+"&DCLASS="+DCLASS+"&ECLASS="+ECLASS;

		}
		if ( CLASSGB == 'C' )
		{
			parent.document.all.CALLKIND_C_2.value = CCLASS;
		}

	}
}
function fn_putetc2()
{
	try{
		//eval("parent.document.all.whereCD7").value = document.all.level2.value;
		eval("parent.document.all.SOSOKETCGB").value = document.all.level2.value;
		parent.frame_sosok2.location = "/menu03/submenu0301/frame_sosok_3.asp?frame_nm=frame_sosok2&SOSOKGB="+parent.document.all.SOSOKGB.value+"&SOSOKETCGB="+document.all.level2.value;
	}
	catch(e){}
}
-->
</script>

<!-- 프레임1 시작 -->
<form name="frmCode" style="margin:0">

<%

CLASSNM = REQUEST("CLASSNM")
CLASSGB = REQUEST("CLASSGB")
ACLASS = REQUEST("ACLASS")
BCLASS = REQUEST("BCLASS")
CCLASS = REQUEST("CCLASS")
DCLASS = REQUEST("DCLASS")
ECLASS = REQUEST("ECLASS")

frame_nm = REQUEST("frame_nm")
bgcolor = REQUEST("bgcolor")
if frame_nm = "frame_callclass_B_2" or frame_nm = "frame_callclass_C_2" or frame_nm = "frame_callclass_B_4" or frame_nm = "frame_callclass_C_4" or frame_nm = "frame_callkind_B_2" or frame_nm = "frame_callkind_C_2" then
	bgcolor = "#FDE6F3"
else
	bgcolor = "#ffffff"
end if

%>
<body bgcolor="#FFFFFF" topmargin=0 leftmargin=0>
<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0" align=left  bgcolor="#000000">
   <tr>
    <td bgcolor="<%=bgcolor%>" align="left" height="29" valign="center">

		<%
			'if SOSOKGB = "" then
			'	sReplyHtml = "소속(대)를 선택하시면 중분류가 표시됩니다."
			'	response.write sReplyHtml
			'else
				'======= 처리구분 코드 가져오기 ==================================================

				IF CLASSNM = "SOSOK" THEN
					IF CLASSGB = "A" THEN
						SqlCode = "select * from tb_armyinfo where aclass < 'O' and bclass is NULL order by aclass"
					ELSEIF CLASSGB = "B" THEN
						SqlCode = "select * from tb_armyinfo where aclass = '"&ACLASS&"' and bclass is not null and cclass is null order by bclass"
					ELSEIF CLASSGB = "C" THEN
						SqlCode = "select * from tb_armyinfo where aclass = '"&ACLASS&"' and bclass = '"&BCLASS&"' and cclass is not null and dclass is null order by dclass"
					ELSEIF CLASSGB = "D" THEN
						SqlCode = "select * from tb_armyinfo where aclass = '"&ACLASS&"' and bclass = '"&BCLASS&"' and cclass = '"&CCLASS&"' and dclass is not null and eclass is null order by dclass"
					ELSEIF CLASSGB = "E" THEN
						SqlCode = "select * from tb_armyinfo where aclass = '"&ACLASS&"' and bclass = '"&BCLASS&"' and cclass = '"&CCLASS&"' and dclass = '"&DCLASS&"' and eclass is not null order by eclass"
					END IF

				ELSE
					IF CLASSGB = "A" THEN
						SqlCode = "select * from tb_armyinfo where aclass = '"&ACLASS&"' and bclass is NULL order by bclass"
					ELSEIF CLASSGB = "B" THEN
						SqlCode = "select * from tb_armyinfo where aclass = '"&ACLASS&"' and bclass is not null and cclass is null order by bclass"
					ELSEIF CLASSGB = "C" THEN
						SqlCode = "select * from tb_armyinfo where aclass = '"&ACLASS&"' and bclass = '"&BCLASS&"' and cclass is not null and dclass is null order by dclass"
					ELSEIF CLASSGB = "D" THEN
						SqlCode = "select * from tb_armyinfo where aclass = '"&ACLASS&"' and bclass = '"&BCLASS&"' and cclass = '"&CCLASS&"' and dclass is not null and eclass is null order by dclass"
					ELSEIF CLASSGB = "E" THEN
						SqlCode = "select * from tb_armyinfo where aclass = '"&ACLASS&"' and bclass = '"&BCLASS&"' and cclass = '"&CCLASS&"' and dclass = '"&DCLASS&"' and eclass is not null order by eclass"
					END IF
				END IF
				set RsCode = db.execute(SqlCode)

				do until RsCode.eof
					j = j + 1
					SelectedValue = ""

					IF CLASSGB = "A" THEN

						sValue = RsCode("aclass")
						sValueName = RsCode("classname")

						if ACLASS = sValue then
							SelectedValue = "checked"
						else
							SelectedValue = ""
						end if
						if RsCode("CounselorYN") = "Y" then
							CounselorYN = "배치"
						else
							CounselorYN = "배치안됨"
						end if

						if j = 1 then
							sReplyHtml = "<input type='RADIO' value='" & sValue & "' name='"&CLASSNM&CLASSGB &"' class='none' " & SelectedValue & " onClick=""fn_SetSosok3('"&CLASSNM&"','"&CLASSGB&"','"&sValue&"','','','','','"&sValueName&"','"&CounselorYN&"');"">" & sValueName	
						elseif j = 11 or j = 21  then
							sReplyHtml = sReplyHtml & "<br><input type='RADIO' value='" & sValue & "' name='"&CLASSNM&CLASSGB &"' class='none' " & SelectedValue & " onClick=""fn_SetSosok3('"&CLASSNM&"','"&CLASSGB&"','"&sValue&"','','','','','"&sValueName&"','"&CounselorYN&"');"">" & sValueName
						else
							sReplyHtml = sReplyHtml & "&nbsp;<input type='RADIO' value='" & sValue & "' name='"&CLASSNM&CLASSGB &"' class='none' " & SelectedValue & " onClick=""fn_SetSosok3('"&CLASSNM&"','"&CLASSGB&"','"&sValue&"','','','','','"&sValueName&"','"&CounselorYN&"');"">" & sValueName
						end if

					ELSEIF CLASSGB = "B" AND CLASSNM = "SOSOK" THEN

						sValue = RsCode("bclass")
						sValueName = RsCode("classname")

						if BCLASS = sValue then
							SelectedValue = "checked"
						else
							SelectedValue = ""
						end if

						if RsCode("CounselorYN") = "Y" then
							CounselorYN = "배치"
						else
							CounselorYN = "배치안됨"
						end if

						if j = 1 then
							sReplyHtml = "<input type='RADIO' value='" & sValue & "' name='"&CLASSNM&CLASSGB &"' class='none' " & SelectedValue & " onClick=""fn_SetSosok3('"&CLASSNM&"','"&CLASSGB&"','"&ACLASS&"','"&sValue&"','','','','"&sValueName&"','"&CounselorYN&"');"">" & sValueName	
						elseif j = 11 or j = 21  then
							sReplyHtml = sReplyHtml & "<br><input type='RADIO' value='" & sValue & "' name='"&CLASSNM&CLASSGB &"' class='none' " & SelectedValue & " onClick=""fn_SetSosok3('"&CLASSNM&"','"&CLASSGB&"','"&ACLASS&"','"&sValue&"','','','','"&sValueName&"','"&CounselorYN&"');"">" & sValueName
						else
							sReplyHtml = sReplyHtml & "&nbsp;<input type='RADIO' value='" & sValue & "' name='"&CLASSNM&CLASSGB &"' class='none' " & SelectedValue & " onClick=""fn_SetSosok3('"&CLASSNM&"','"&CLASSGB&"','"&ACLASS&"','"&sValue&"','','','','"&sValueName&"','"&CounselorYN&"');"">" & sValueName
						end if

					ELSEIF CLASSGB = "B" THEN

						sValue = RsCode("bclass")
						sValueName = RsCode("classname")

						if BCLASS = sValue then
							SelectedValue = "checked"
						else
							SelectedValue = ""
						end if

						if RsCode("CounselorYN") = "Y" then
							CounselorYN = "배치"
						else
							CounselorYN = "배치안됨"
						end if

						if j = 1 then
							sReplyHtml = "<input type='RADIO' value='" & sValue & "' name='"&CLASSNM&CLASSGB &"' class='none' " & SelectedValue & " onClick=""fn_SetSosok3('"&CLASSNM&"','"&CLASSGB&"','"&ACLASS&"','"&sValue&"','','','','"&sValueName&"','"&CounselorYN&"');"">" & sValueName	
						elseif j = 21  then
							sReplyHtml = sReplyHtml & "<br><input type='RADIO' value='" & sValue & "' name='"&CLASSNM&CLASSGB &"' class='none' " & SelectedValue & " onClick=""fn_SetSosok3('"&CLASSNM&"','"&CLASSGB&"','"&ACLASS&"','"&sValue&"','','','','"&sValueName&"','"&CounselorYN&"');"">" & sValueName
						else
							sReplyHtml = sReplyHtml & "&nbsp;<input type='RADIO' value='" & sValue & "' name='"&CLASSNM&CLASSGB &"' class='none' " & SelectedValue & " onClick=""fn_SetSosok3('"&CLASSNM&"','"&CLASSGB&"','"&ACLASS&"','"&sValue&"','','','','"&sValueName&"','"&CounselorYN&"');"">" & sValueName
						end if

					ELSEIF CLASSGB = "C" THEN

						sValue = RsCode("cclass")
						sValueName = RsCode("classname")

						if CCLASS = sValue then
							SelectedValue = "checked"
						else
							SelectedValue = ""
						end if

						if RsCode("CounselorYN") = "Y" then
							CounselorYN = "배치"
						else
							CounselorYN = "배치안됨"
						end if

						if j = 1 then
							sReplyHtml = "<input type='RADIO' value='" & sValue & "' name='"&CLASSNM&CLASSGB &"' class='none' " & SelectedValue & " onClick=""fn_SetSosok3('"&CLASSNM&"','"&CLASSGB&"','"&ACLASS&"','"&BCLASS&"','"&sValue&"','','','"&sValueName&"','"&CounselorYN&"');"">" & sValueName	
						elseif j = 13  then
							sReplyHtml = sReplyHtml & "<br><input type='RADIO' value='" & sValue & "' name='"&CLASSNM&CLASSGB &"' class='none' " & SelectedValue & " onClick=""fn_SetSosok3('"&CLASSNM&"','"&CLASSGB&"','"&ACLASS&"','"&BCLASS&"','"&sValue&"','','','"&sValueName&"','"&CounselorYN&"');"">" & sValueName	
						else
							sReplyHtml = sReplyHtml & "&nbsp;<input type='RADIO' value='" & sValue & "' name='"&CLASSNM&CLASSGB &"' class='none' " & SelectedValue & " onClick=""fn_SetSosok3('"&CLASSNM&"','"&CLASSGB&"','"&ACLASS&"','"&BCLASS&"','"&sValue&"','','','"&sValueName&"','"&CounselorYN&"');"">" & sValueName	
						end if


					ELSEIF CLASSGB = "D" THEN
						sValue = RsCode("dclass")
						sValueName = RsCode("classname")
						if DCLASS = sValue then
							SelectedValue = "checked"
						else
							SelectedValue = ""
						end if

						if RsCode("CounselorYN") = "Y" then
							CounselorYN = "배치"
						else
							CounselorYN = "배치안됨"
						end if
						if j = 1 then
							sReplyHtml = "<input type='RADIO' value='" & sValue & "' name='"&CLASSNM&CLASSGB &"' class='none' " & SelectedValue & " onClick=""fn_SetSosok3('"&CLASSNM&"','"&CLASSGB&"','"&ACLASS&"','"&BCLASS&"','"&CCLASS&"','"&sValue&"','','"&sValueName&"','"&CounselorYN&"');"">" & sValueName	
						elseif j = 11 or j = 21  then
							sReplyHtml = sReplyHtml & "<br><input type='RADIO' value='" & sValue & "' name='"&CLASSNM&CLASSGB &"' class='none' " & SelectedValue & " onClick=""fn_SetSosok3('"&CLASSNM&"','"&CLASSGB&"','"&ACLASS&"','"&BCLASS&"','"&CCLASS&"','"&sValue&"','','"&sValueName&"','"&CounselorYN&"');"">" & sValueName	
						else
							sReplyHtml = sReplyHtml & "&nbsp;<input type='RADIO' value='" & sValue & "' name='"&CLASSNM&CLASSGB &"' class='none' " & SelectedValue & " onClick=""fn_SetSosok3('"&CLASSNM&"','"&CLASSGB&"','"&ACLASS&"','"&BCLASS&"','"&CCLASS&"','"&sValue&"','','"&sValueName&"','"&CounselorYN&"');"">" & sValueName	
						end if

					ELSEIF CLASSGB = "E" THEN
						sValue = RsCode("eclass")
						sValueName = RsCode("classname")

						if ECLASS = sValue then
							SelectedValue = "checked"
						else
							SelectedValue = ""
						end if

						if RsCode("CounselorYN") = "Y" then
							CounselorYN = "배치"
						else
							CounselorYN = "배치안됨"
						end if
						if j = 1 then
							sReplyHtml = "<input type='RADIO' value='" & sValue & "' name='"&CLASSNM&CLASSGB &"' class='none' " & SelectedValue & " onClick=""fn_SetSosok3('"&CLASSNM&"','"&CLASSGB&"','"&ACLASS&"','"&BCLASS&"','"&CCLASS&"','"&DCLASS&"','"&sValue&"','"&sValueName&"','"&CounselorYN&"');"">" & sValueName	
						elseif j = 11 or j = 21  then
							sReplyHtml = sReplyHtml & "<br><input type='RADIO' value='" & sValue & "' name='"&CLASSNM&CLASSGB &"' class='none' " & SelectedValue & " onClick=""fn_SetSosok3('"&CLASSNM&"','"&CLASSGB&"','"&ACLASS&"','"&BCLASS&"','"&CCLASS&"','"&DCLASS&"',"&sValue&",'"&sValueName&"','"&CounselorYN&"');"">" & sValueName
						else
							sReplyHtml = sReplyHtml & "&nbsp;<input type='RADIO' value='" & sValue & "' name='"&CLASSNM&CLASSGB &"' class='none' " & SelectedValue & " onClick=""fn_SetSosok3('"&CLASSNM&"','"&CLASSGB&"','"&ACLASS&"','"&BCLASS&"','"&CCLASS&"','"&DCLASS&"',"&sValue&",'"&sValueName&"','"&CounselorYN&"');"">" & sValueName
						end if


					END IF


					RsCode.movenext
				loop
				RsCode.close
				response.write sReplyHtml

			'end if

		if sReplyHtml <> "" then

			sReplyHtml = "&nbsp;<img src='/Images/Comm/IconDel2.gif' title='선택취소' style='cursor:hand;' align='absmiddle' onclick=""javascript:parent.fn_DEL('" & frame_nm & "');"">"
			response.write sReplyHtml
			
		end if

		IF CLASSNM = "SOSOK" and j >= 9 THEN
			%>
			<script>
				//parent.fn_SetHeight();
			</script>
			<%
		end if
		%>						

	</td>
</tr>
</table>
</form>
<!-- #include virtual="/Include/Bottom.asp" -->
