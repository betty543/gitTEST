<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!-- #include virtual="/include/CacheNo.asp" -->
<!-- #include virtual="/include/common.asp" -->
<%

guboon = Request("guboon")						'저장/수정/삭제 FLAG


db_Date2 = Request("Date2")
db_Date3 = Request("Date3")
db_receiptkind = Request("receiptkind")

db_FRM1 = Request("FRM1")
db_FRM2 = Request("FRM2")
db_FRM3 = Request("FRM3")
db_FRM4 = Request("FRM4")
db_FRM5 = Request("FRM5")
db_FRM6 = Request("FRM6")
db_FRM7 = Request("FRM7")
db_FRM8 = Request("FRM8")
db_FRM9 = Request("FRM9")
db_FRM10 = Request("FRM10")
db_FRM11 = Request("FRM11")
db_FRM12 = Request("FRM12")
db_FRM13 = Request("FRM13")
db_FRM14 = Request("FRM14")
db_FRM15 = Request("FRM15")

db_RECYN1 = Request("RecYN_1")
if db_RECYN1 = "" then 	db_RECYN1 = "N" end if
db_RECYN2 = Request("RecYN_2")
if db_RECYN2 = "" then 	db_RECYN2 = "N" end if
db_RECYN3 = Request("RecYN_3")
if db_RECYN3 = "" then 	db_RECYN3 = "N" end if
db_RECYN4 = Request("RecYN_4")
if db_RECYN4 = "" then 	db_RECYN4 = "N" end if
db_RECYN5 = Request("RecYN_5")
if db_RECYN5 = "" then 	db_RECYN5 = "N" end if
db_RECYN6 = Request("RecYN_6")
if db_RECYN6 = "" then 	db_RECYN6 = "N" end if
db_RECYN7 = Request("RecYN_7")
if db_RECYN7 = "" then 	db_RECYN7 = "N" end if
db_RECYN8 = Request("RecYN_8")
if db_RECYN8 = "" then 	db_RECYN8 = "N" end if
db_RECYN9 = Request("RecYN_9")
if db_RECYN9 = "" then 	db_RECYN9 = "N" end if
db_RECYN10 = Request("RecYN_10")
if db_RECYN10 = "" then 	db_RECYN10 = "N" end if
db_RECYN11 = Request("RecYN_11")
if db_RECYN11 = "" then 	db_RECYN11 = "N" end if
db_RECYN12 = Request("RecYN_12")
if db_RECYN12 = "" then 	db_RECYN12 = "N" end if
db_RECYN13 = Request("RecYN_13")
if db_RECYN13 = "" then 	db_RECYN13 = "N" end if
db_RECYN14 = Request("RecYN_14")
if db_RECYN14 = "" then 	db_RECYN14 = "N" end if
db_RECYN15 = Request("RecYN_15")
if db_RECYN15 = "" then 	db_RECYN15 = "N" end if

db_IDX1 = Request("IDX_1")
db_IDX2 = Request("IDX_2")
db_IDX3 = Request("IDX_3")
db_IDX4 = Request("IDX_4")
db_IDX5 = Request("IDX_5")
db_IDX6 = Request("IDX_6")
db_IDX7 = Request("IDX_7")
db_IDX8 = Request("IDX_8")
db_IDX9 = Request("IDX_9")
db_IDX10 = Request("IDX_10")
db_IDX11 = Request("IDX_11")
db_IDX12 = Request("IDX_12")
db_IDX13 = Request("IDX_13")
db_IDX14 = Request("IDX_14")
db_IDX15 = Request("IDX_15")

db_RECEIPTFACTNUM = Request("RECEIPTFACTNUM")
'##1번설문지
db_factPeoplenum_1 = Request("factPeoplenum_1")
db_SECTION2_1 = Request("SECTION2_1")		
db_NAME_1 = Request("NAME_1")	
db_HOMEPHONE_1 = Request("HOMEPHONE_1")
db_MOBILEPHONE_1 = Request("MOBILEPHONE_1")
db_ETCPHONE_1 = Request("ETCPHONE_1")
db_MONITORDATE_1 = Request("MONITORDATE_1")
db_QUESTIONP_1 = Request("QUESTIONP_1")
db_LEVEL_1 = Request("LEVEL_1")
db_RESERVEHOUR_1 = Request("RESERVEHOUR_1")
db_RESERVEMIN_1 = Request("RESERVEMIN_1")
db_MONITORRESULT_1 = Request("MONITORRESULT_1")
db_RESERVEDATE_1 = Request("RESERVEDATE_1")
db_Remark_1 = Request("Remark_1")
db_Remark1_1 = Request("Remark1_1")
db_TOT_1 = Request("TOT_1")
if db_TOT_1 = "" then
	db_TOT_1 = "0.00"
end if

'##2번설문지
db_factPeoplenum_2 = Request("factPeoplenum_2")
db_SECTION2_2 = Request("SECTION2_2")		
db_NAME_2 = Request("NAME_2")	
db_HOMEPHONE_2 = Request("HOMEPHONE_2")
db_MOBILEPHONE_2 = Request("MOBILEPHONE_2")
db_ETCPHONE_2 = Request("ETCPHONE_2")
db_MONITORDATE_2 = Request("MONITORDATE_2")
db_QUESTIONP_2 = Request("QUESTIONP_2")
db_LEVEL_2 = Request("LEVEL_2")
db_RESERVEHOUR_2 = Request("RESERVEHOUR_2")
db_RESERVEMIN_2 = Request("RESERVEMIN_2")
db_MONITORRESULT_2 = Request("MONITORRESULT_2")
db_RESERVEDATE_2 = Request("RESERVEDATE_2")
db_Remark_2 = Request("Remark_2")
db_Remark1_2 = Request("Remark1_2")
db_TOT_2 = Request("TOT_2")
if db_TOT_2 = "" then
	db_TOT_2 = "0.00"
end if


'##3번설문지
db_factPeoplenum_3 = Request("factPeoplenum_3")
db_SECTION2_3 = Request("SECTION2_3")		
db_NAME_3 = Request("NAME_3")	
db_HOMEPHONE_3 = Request("HOMEPHONE_3")
db_MOBILEPHONE_3 = Request("MOBILEPHONE_3")
db_ETCPHONE_3 = Request("ETCPHONE_3")
db_MONITORDATE_3 = Request("MONITORDATE_3")
db_QUESTIONP_3 = Request("QUESTIONP_3")
db_LEVEL_3 = Request("LEVEL_3")
db_RESERVEHOUR_3 = Request("RESERVEHOUR_3")
db_RESERVEMIN_3 = Request("RESERVEMIN_3")
db_MONITORRESULT_3 = Request("MONITORRESULT_3")
db_RESERVEDATE_3 = Request("RESERVEDATE_3")
db_Remark_3 = Request("Remark_3")
db_Remark1_3 = Request("Remark1_3")
db_TOT_3 = Request("TOT_3")
if db_TOT_3 = "" then
	db_TOT_3 = "0.00"
end if


'##4번설문지
db_factPeoplenum_4 = Request("factPeoplenum_4")
db_SECTION2_4 = Request("SECTION2_4")		
db_NAME_4 = Request("NAME_4")	
db_HOMEPHONE_4 = Request("HOMEPHONE_4")
db_MOBILEPHONE_4 = Request("MOBILEPHONE_4")
db_ETCPHONE_4 = Request("ETCPHONE_4")
db_MONITORDATE_4 = Request("MONITORDATE_4")
db_QUESTIONP_4 = Request("QUESTIONP_4")
db_LEVEL_4 = Request("LEVEL_4")
db_RESERVEHOUR_4 = Request("RESERVEHOUR_4")
db_RESERVEMIN_4 = Request("RESERVEMIN_4")
db_MONITORRESULT_4 = Request("MONITORRESULT_4")
db_RESERVEDATE_4 = Request("RESERVEDATE_4")
db_Remark_4 = Request("Remark_4")
db_Remark1_4 = Request("Remark1_4")
db_TOT_4 = Request("TOT_4")
if db_TOT_4 = "" then
	db_TOT_4 = "0.00"
end if
'response.write db_Remark1_4 & "vbcrlf"


'##5번설문지
db_factPeoplenum_5 = Request("factPeoplenum_5")
db_SECTION2_5 = Request("SECTION2_5")		
db_NAME_5 = Request("NAME_5")	
db_HOMEPHONE_5 = Request("HOMEPHONE_5")
db_MOBILEPHONE_5 = Request("MOBILEPHONE_5")
db_ETCPHONE_5 = Request("ETCPHONE_5")
db_MONITORDATE_5 = Request("MONITORDATE_5")
db_QUESTIONP_5 = Request("QUESTIONP_5")
db_LEVEL_5 = Request("LEVEL_5")
db_RESERVEHOUR_5 = Request("RESERVEHOUR_5")
db_RESERVEMIN_5 = Request("RESERVEMIN_5")
db_MONITORRESULT_5 = Request("MONITORRESULT_5")
db_RESERVEDATE_5 = Request("RESERVEDATE_5")
db_Remark_5 = Request("Remark_5")
db_Remark1_5 = Request("Remark1_5")
db_TOT_5 = Request("TOT_5")
if db_TOT_5 = "" then
	db_TOT_5 = "0.00"
end if



'##6번설문지
db_factPeoplenum_6 = Request("factPeoplenum_6")
db_SECTION2_6 = Request("SECTION2_6")		
db_NAME_6 = Request("NAME_6")	
db_HOMEPHONE_6 = Request("HOMEPHONE_6")
db_MOBILEPHONE_6 = Request("MOBILEPHONE_6")
db_ETCPHONE_6 = Request("ETCPHONE_6")
db_MONITORDATE_6 = Request("MONITORDATE_6")
db_QUESTIONP_6 = Request("QUESTIONP_6")
db_LEVEL_6 = Request("LEVEL_6")
db_RESERVEHOUR_6 = Request("RESERVEHOUR_6")
db_RESERVEMIN_6 = Request("RESERVEMIN_6")
db_MONITORRESULT_6 = Request("MONITORRESULT_6")
db_RESERVEDATE_6 = Request("RESERVEDATE_6")
db_Remark_6 = Request("Remark_6")
db_Remark1_6 = Request("Remark1_6")
db_TOT_6 = Request("TOT_6")
if db_TOT_6 = "" then
	db_TOT_6 = "0.00"
end if


'##7번설문지
db_factPeoplenum_7 = Request("factPeoplenum_7")
db_SECTION2_7 = Request("SECTION2_7")		
db_NAME_7 = Request("NAME_7")	
db_HOMEPHONE_7 = Request("HOMEPHONE_7")
db_MOBILEPHONE_7 = Request("MOBILEPHONE_7")
db_ETCPHONE_7 = Request("ETCPHONE_7")
db_MONITORDATE_7 = Request("MONITORDATE_7")
db_QUESTIONP_7 = Request("QUESTIONP_7")
db_LEVEL_7 = Request("LEVEL_7")
db_RESERVEHOUR_7 = Request("RESERVEHOUR_7")
db_RESERVEMIN_7 = Request("RESERVEMIN_7")
db_MONITORRESULT_7 = Request("MONITORRESULT_7")
db_RESERVEDATE_7 = Request("RESERVEDATE_7")
db_Remark_7 = Request("Remark_7")
db_Remark1_7 = Request("Remark1_7")
db_TOT_7 = Request("TOT_7")
if db_TOT_7 = "" then
	db_TOT_7 = "0.00"
end if


'##8번설문지
db_factPeoplenum_8 = Request("factPeoplenum_8")
db_SECTION2_8 = Request("SECTION2_8")		
db_NAME_8 = Request("NAME_8")	
db_HOMEPHONE_8 = Request("HOMEPHONE_8")
db_MOBILEPHONE_8 = Request("MOBILEPHONE_8")
db_ETCPHONE_8 = Request("ETCPHONE_8")
db_MONITORDATE_8 = Request("MONITORDATE_8")
db_QUESTIONP_8 = Request("QUESTIONP_8")
db_LEVEL_8 = Request("LEVEL_8")
db_RESERVEHOUR_8 = Request("RESERVEHOUR_8")
db_RESERVEMIN_8 = Request("RESERVEMIN_8")
db_MONITORRESULT_8 = Request("MONITORRESULT_8")
db_RESERVEDATE_8 = Request("RESERVEDATE_8")
db_Remark_8 = Request("Remark_8")
db_Remark1_8 = Request("Remark1_8")
db_TOT_8 = Request("TOT_8")
if db_TOT_8 = "" then
	db_TOT_8 = "0.00"
end if

'##9번설문지
db_factPeoplenum_9 = Request("factPeoplenum_9")
db_SECTION2_9 = Request("SECTION2_9")		
db_NAME_9 = Request("NAME_9")	
db_HOMEPHONE_9 = Request("HOMEPHONE_9")
db_MOBILEPHONE_9 = Request("MOBILEPHONE_9")
db_ETCPHONE_9 = Request("ETCPHONE_9")
db_MONITORDATE_9 = Request("MONITORDATE_9")
db_QUESTIONP_9 = Request("QUESTIONP_9")
db_LEVEL_9 = Request("LEVEL_9")
db_RESERVEHOUR_9 = Request("RESERVEHOUR_9")
db_RESERVEMIN_9 = Request("RESERVEMIN_9")
db_MONITORRESULT_9 = Request("MONITORRESULT_9")
db_RESERVEDATE_9 = Request("RESERVEDATE_9")
db_Remark_9 = Request("Remark_9")
db_Remark1_9 = Request("Remark1_9")
db_TOT_9 = Request("TOT_9")
if db_TOT_9 = "" then
	db_TOT_9 = "0.00"
end if


'##10번설문지
db_factPeoplenum_10 = Request("factPeoplenum_10")
db_SECTION2_10 = Request("SECTION2_10")		
db_NAME_10 = Request("NAME_10")	
db_HOMEPHONE_10 = Request("HOMEPHONE_10")
db_MOBILEPHONE_10 = Request("MOBILEPHONE_10")
db_ETCPHONE_10 = Request("ETCPHONE_10")
db_MONITORDATE_10 = Request("MONITORDATE_10")
db_QUESTIONP_10 = Request("QUESTIONP_10")
db_LEVEL_10 = Request("LEVEL_10")
db_RESERVEHOUR_10 = Request("RESERVEHOUR_10")
db_RESERVEMIN_10 = Request("RESERVEMIN_10")
db_MONITORRESULT_10 = Request("MONITORRESULT_10")
db_RESERVEDATE_10 = Request("RESERVEDATE_10")
db_Remark_10 = Request("Remark_10")
db_Remark1_10 = Request("Remark1_10")
db_TOT_10 = Request("TOT_10")
if db_TOT_10 = "" then
	db_TOT_10 = "0.00"
end if


'##11번설문지
db_factPeoplenum_11 = Request("factPeoplenum_11")
db_SECTION2_11 = Request("SECTION2_11")		
db_NAME_11 = Request("NAME_11")	
db_HOMEPHONE_11 = Request("HOMEPHONE_11")
db_MOBILEPHONE_11 = Request("MOBILEPHONE_11")
db_ETCPHONE_11 = Request("ETCPHONE_11")
db_MONITORDATE_11 = Request("MONITORDATE_11")
db_QUESTIONP_11 = Request("QUESTIONP_11")
db_LEVEL_11 = Request("LEVEL_11")
db_RESERVEHOUR_11 = Request("RESERVEHOUR_11")
db_RESERVEMIN_11 = Request("RESERVEMIN_11")
db_MONITORRESULT_11 = Request("MONITORRESULT_11")
db_RESERVEDATE_11 = Request("RESERVEDATE_11")
db_Remark_11 = Request("Remark_11")
db_Remark1_11 = Request("Remark1_11")
db_TOT_11 = Request("TOT_11")
if db_TOT_11 = "" then
	db_TOT_11 = "0.00"
end if


'##12번설문지
db_factPeoplenum_12 = Request("factPeoplenum_12")
db_SECTION2_12 = Request("SECTION2_12")		
db_NAME_12 = Request("NAME_12")	
db_HOMEPHONE_12 = Request("HOMEPHONE_12")
db_MOBILEPHONE_12 = Request("MOBILEPHONE_12")
db_ETCPHONE_12 = Request("ETCPHONE_12")
db_MONITORDATE_12 = Request("MONITORDATE_12")
db_QUESTIONP_12 = Request("QUESTIONP_12")
db_LEVEL_12 = Request("LEVEL_12")
db_RESERVEHOUR_12 = Request("RESERVEHOUR_12")
db_RESERVEMIN_12 = Request("RESERVEMIN_12")
db_MONITORRESULT_12 = Request("MONITORRESULT_12")
db_RESERVEDATE_12 = Request("RESERVEDATE_12")
db_Remark_12 = Request("Remark_12")
db_Remark1_12 = Request("Remark1_12")
db_TOT_12 = Request("TOT_12")
if db_TOT_12 = "" then
	db_TOT_12 = "0.00"
end if


'##13번설문지
db_factPeoplenum_13 = Request("factPeoplenum_13")
db_SECTION2_13 = Request("SECTION2_13")		
db_NAME_13 = Request("NAME_13")	
db_HOMEPHONE_13 = Request("HOMEPHONE_13")
db_MOBILEPHONE_13 = Request("MOBILEPHONE_13")
db_ETCPHONE_13 = Request("ETCPHONE_13")
db_MONITORDATE_13 = Request("MONITORDATE_13")
db_QUESTIONP_13 = Request("QUESTIONP_13")
db_LEVEL_13 = Request("LEVEL_13")
db_RESERVEHOUR_13 = Request("RESERVEHOUR_13")
db_RESERVEMIN_13 = Request("RESERVEMIN_13")
db_MONITORRESULT_13 = Request("MONITORRESULT_13")
db_RESERVEDATE_13 = Request("RESERVEDATE_13")
db_Remark_13 = Request("Remark_13")
db_Remark1_13 = Request("Remark1_13")
db_TOT_13 = Request("TOT_13")
if db_TOT_13 = "" then
	db_TOT_13 = "0.00"
end if


'##14번설문지
db_factPeoplenum_14 = Request("factPeoplenum_14")
db_SECTION2_14 = Request("SECTION2_14")		
db_NAME_14 = Request("NAME_14")	
db_HOMEPHONE_14 = Request("HOMEPHONE_14")
db_MOBILEPHONE_14 = Request("MOBILEPHONE_14")
db_ETCPHONE_14 = Request("ETCPHONE_14")
db_MONITORDATE_14 = Request("MONITORDATE_14")
db_QUESTIONP_14 = Request("QUESTIONP_14")
db_LEVEL_14 = Request("LEVEL_14")
db_RESERVEHOUR_14 = Request("RESERVEHOUR_14")
db_RESERVEMIN_14 = Request("RESERVEMIN_14")
db_MONITORRESULT_14 = Request("MONITORRESULT_14")
db_RESERVEDATE_14 = Request("RESERVEDATE_14")
db_Remark_14 = Request("Remark_14")
db_Remark1_14 = Request("Remark1_14")
db_TOT_14 = Request("TOT_14")
if db_TOT_14 = "" then
	db_TOT_14 = "0.00"
end if


'##15번설문지
db_factPeoplenum_15 = Request("factPeoplenum_15")
db_SECTION2_15 = Request("SECTION2_15")		
db_NAME_15 = Request("NAME_15")	
db_HOMEPHONE_15 = Request("HOMEPHONE_15")
db_MOBILEPHONE_15 = Request("MOBILEPHONE_15")
db_ETCPHONE_15 = Request("ETCPHONE_15")
db_MONITORDATE_15 = Request("MONITORDATE_15")
db_QUESTIONP_15 = Request("QUESTIONP_15")
db_LEVEL_15 = Request("LEVEL_15")
db_RESERVEHOUR_15 = Request("RESERVEHOUR_15")
db_RESERVEMIN_15 = Request("RESERVEMIN_15")
db_MONITORRESULT_15 = Request("MONITORRESULT_15")
db_RESERVEDATE_15 = Request("RESERVEDATE_15")
db_Remark_15 = Request("Remark_15")
db_Remark1_15 = Request("Remark1_15")
db_TOT_15 = Request("TOT_15")
if db_TOT_15 = "" then
	db_TOT_15 = "0.00"
end if

db_Remark_1 = replace(db_Remark_1,"'","''")
db_Remark1_1 = replace(db_Remark1_1,"'","''")

db_Remark_2 = replace(db_Remark_2,"'","''")
db_Remark1_2 = replace(db_Remark1_2,"'","''")

db_Remark_3 = replace(db_Remark_3,"'","''")
db_Remark1_3 = replace(db_Remark1_3,"'","''")

db_Remark_4 = replace(db_Remark_4,"'","''")
db_Remark1_4 = replace(db_Remark1_4,"'","''")

db_Remark_5 = replace(db_Remark_5,"'","''")
db_Remark1_5 = replace(db_Remark1_5,"'","''")

db_Remark_6 = replace(db_Remark_6,"'","''")
db_Remark1_6 = replace(db_Remark1_6,"'","''")

db_Remark_7 = replace(db_Remark_7,"'","''")
db_Remark1_7 = replace(db_Remark1_7,"'","''")

db_Remark_8 = replace(db_Remark_8,"'","''")
db_Remark1_8 = replace(db_Remark1_8,"'","''")


db_Remark_9 = replace(db_Remark_9,"'","''")
db_Remark1_9 = replace(db_Remark1_9,"'","''")

db_Remark_10 = replace(db_Remark_10,"'","''")
db_Remark1_10 = replace(db_Remark1_10,"'","''")

db_Remark_11 = replace(db_Remark_11,"'","''")
db_Remark1_11 = replace(db_Remark1_11,"'","''")

db_Remark_12 = replace(db_Remark_12,"'","''")
db_Remark1_12 = replace(db_Remark1_12,"'","''")

db_Remark_13 = replace(db_Remark_13,"'","''")
db_Remark1_13 = replace(db_Remark1_13,"'","''")

db_Remark_14 = replace(db_Remark_14,"'","''")
db_Remark1_14 = replace(db_Remark1_14,"'","''")

db_Remark_15 = replace(db_Remark_15,"'","''")
db_Remark1_15 = replace(db_Remark1_15,"'","''")


	INCODE = SESSION("SS_LoginID")

	If INCODE = "" Then	

		INCODE = Request.Cookies("ASRNC")("WebUserid")
		SQL=" SELECT *"
		SQL = SQL & " FROM TB_USERINFO"
		SQL = SQL & " WHERE USERID = '" & INCODE & "'"

		Set RS = db.Execute(SQL)

		If RS.eof = False Then
		
			SESSION("SS_LoginID") = RS("USERID")
			SESSION("SS_LoginNAME") = RS("UserName")
			SESSION("SS_Login_Secgroup") = RS("SECGROUP")
			SESSION("SS_Login_Grade") = RS("GRADE")
			SESSION("SS_Login_GradeName") = RS("GRADE")' db_getCodeName("Z03",RS("GRADE")) 
			SESSION("SS_Login_CTIYN") = RS("CTIYN")

			SS_LoginID = SESSION("SS_LoginID")
			SS_LoginNAME = SESSION("SS_LoginNAME")
			SS_Login_Secgroup = SESSION("SS_Login_Secgroup")
			SS_Login_Grade = SESSION("SS_Login_Grade")
			SS_Login_GradeName = SESSION("SS_Login_GradeName")
			SS_Login_Agentcode = SESSION("SS_Login_Agentcode")
			SS_Login_CTIYN = SESSION("SS_Login_CTIYN")

		End If

	end if

	SQL = "select count(*) from tb_code where codegroup = '" & db_SECTION2_1 & "'"

	set rs = db.execute(SQL)

	icnt = rs(0)

	if db_FRM1 = "ON" then	'수정또는 Insert
		
		if db_factPeoplenum_1 = "" then

			SQL = "select convert(varchar,convert(int,isnull(rtrim(max(convert(int,substring(factPeoplenum,2,10)))),'0'))+1) from armyinformix.dbo.factpeople where left(factPeoplenum,1) = 'M'"
	
			set RS = db.execute(SQL)
			if isnull(rs(0)) then
				db_factPeoplenum_1 = "M1"
			else
				db_factPeoplenum_1 = "M"&rs(0)
			end if


			'신규로 임의 입력건
			SQL = "INSERT INTO armyinformix.dbo.factpeople ( factPeoplenum, factnum"
			SQL = SQL & "		, SECTION2, NAME, level"
			SQL = SQL & "		,homephone, mobilephone, etcphone"
			SQL = SQL & "		,MONITORDATE, MONITORRESULT"
			SQL = SQL & "		,RESERVEDATE, Remark, monitorpoint, Remark1 )"
			SQL = SQL & " VALUES ('" & db_factPeoplenum_1 & "', '" & db_RECEIPTFACTNUM & "'"
			SQL = SQL & "		,'" & db_SECTION2_1 & "','" & db_name_1 & "','" & db_level_1 & "'"
			SQL = SQL & "		,'" & db_homephone_1 & "','" & db_mobilephone_1 & "','" & db_etcphone_1 & "'"
			if db_MONITORDATE_1 = "" then
				SQL = SQL & "		,	getdate()"
			else
				SQL = SQL & "		,	'" & db_MONITORDATE_1 & "'"
			end if
			SQL = SQL & "		,'" & db_MONITORRESULT_1 & "'"

			if db_MONITORRESULT_1 = "4" then	'예약이라면
				if db_RESERVETIME_1 = "1" then
					SQL = SQL & "		,	CONVERT(CHAR(19),DATEADD(n,10,GETDATE()),121)"
				elseif db_RESERVETIME_1 = "2" then
					SQL = SQL & "		,	CONVERT(CHAR(19),DATEADD(n,30,GETDATE()),121)"
				elseif db_RESERVETIME_1 = "3" then
					SQL = SQL & "		,	CONVERT(CHAR(19),DATEADD(n,60,GETDATE()),121)"
				elseif db_RESERVETIME_1 = "4" then
					SQL = SQL & "		,	CONVERT(CHAR(19),DATEADD(n,120,GETDATE()),121)"
				elseif db_RESERVETIME_1 <> "" then
					SQL = SQL & "		,	'" & db_RESERVEDATE_1 & " " & db_RESERVETIME_1 & ":00:00'"
				else
					SQL = SQL & "		,	'" & db_RESERVEDATE_1 & " " & db_RESERVEHOUR_1 & ":"&  db_RESERVEMIN_1 &":00'"
				end if
			else
				SQL = SQL & "		,	''"
			end if
			SQL = SQL & "		,	'" & db_Remark_1 & "'"
			SQL = SQL & "		,	'" & db_TOT_1 & "'"
			SQL = SQL & "		,	'" & db_Remark1_1 & "')"

			db.execute(SQL)
	
		else

			SQL = " update armyinformix.dbo.factpeople set SECTION2 = '" & db_SECTION2_1 & "'"
			SQL = SQL & "		,	name = '" & db_name_1 & "',	level = '" & db_level_1 & "', homephone = '" & db_homephone_1 & "'"
			SQL = SQL & "		,	mobilephone = '" & db_mobilephone_1 & "', etcphone = '" & db_etcphone_1 & "'"

			if db_MONITORDATE_1 = "" then
				SQL = SQL & "		,	MONITORDATE = getdate()"
			else
				SQL = SQL & "		,	MONITORDATE = convert(datetime,'"&db_MONITORDATE_1&"')"
			end if
			SQL = SQL & "		,	MONITORRESULT = '" & db_MONITORRESULT_1 & "'"

			if db_MONITORRESULT_1 = "4" then	'예약이라면
				if db_RESERVETIME_1 = "1" then
					SQL = SQL & "		,	RESERVEDATE = CONVERT(CHAR(19),DATEADD(n,10,GETDATE()),121)"
				elseif db_RESERVETIME_1 = "2" then
					SQL = SQL & "		,	RESERVEDATE = CONVERT(CHAR(19),DATEADD(n,30,GETDATE()),121)"
				elseif db_RESERVETIME_1 = "3" then
					SQL = SQL & "		,	RESERVEDATE = CONVERT(CHAR(19),DATEADD(n,60,GETDATE()),121)"
				elseif db_RESERVETIME_1 = "4" then
					SQL = SQL & "		,	RESERVEDATE = CONVERT(CHAR(19),DATEADD(n,120,GETDATE()),121)"
				elseif db_RESERVETIME_1 <> "" then
					SQL = SQL & "		,	RESERVEDATE = '" & db_RESERVEDATE_1 & " " & db_RESERVETIME_1 & ":00:00'"
				else
					SQL = SQL & "		,	RESERVEDATE = '" & db_RESERVEDATE_1 & " " & db_RESERVEHOUR_1 & ":"&  db_RESERVEMIN_1 &":00'"
				end if
			else
				SQL = SQL & "		,	RESERVEDATE = ''"
			end if
			SQL = SQL & "		,	Remark = '" & db_Remark_1 & "'"
			SQL = SQL & "		,	Remark1 = '" & db_Remark1_1 & "'"
			SQL = SQL & "		,	monitorpoint = '" & db_TOT_1 & "'"
			SQL = SQL & "	where factPeoplenum = '" & db_factPeoplenum_1 & "' and factnum = '" & db_RECEIPTFACTNUM & "'"

'response.write SQL

			db.execute(SQL)
		end if

		'## 1번
		SQL = "	delete from armyinformix.dbo.monitor where factnum = '" & db_RECEIPTFACTNUM & "' and factPeoplenum = '" & db_factPeoplenum_1 & "'"
		db.execute(SQL)
		for i = 1 to icnt
			point9 =  Request("QUESTION_1"&i)
			if point9 = "9" then
				ipoint9 = "1"
			else
				ipoint9 = "0"
			end if
			if point9 = "8" then
				ipoint8 = "1"
			else
				ipoint8 = "0"
			end if
			if point9 = "7" then
				ipoint7 = "1"
			else
				ipoint7 = "0"
			end if
			pointplus =  Request("QUESTIONP_1"&i)
			if pointplus = "" then
				pointplus = "0"
			end if
			totpoint =  Request("POINT_1"&i)
			if totpoint = "" then
				totpoint = "0"
			end if

			SQL = "insert into armyinformix.dbo.monitor ( factnum, factPeoplenum, seqno, point9, point8, point7, pointplus, totpoint, monitordate, monitoruser)"
			SQL = SQL & " values ( '" & db_RECEIPTFACTNUM & "', '" & db_factPeoplenum_1 & "', " & i & ", " & ipoint9& ", " & ipoint8& ", " & ipoint7
			SQL = SQL & " , " & pointplus & ","&totpoint&",getdate(),'" &INCODE&"')"
			db.execute(SQL)
			'Response.write SQL
		next
	else

		SQL = "delete from armyinformix.dbo.factpeople"
		SQL = SQL & "	where factPeoplenum = '" & db_factPeoplenum_1 & "' and factnum = '" & db_RECEIPTFACTNUM & "'"
		db.execute(SQL)
		'## 1번
		SQL = "	delete from armyinformix.dbo.monitor where factnum = '" & db_RECEIPTFACTNUM & "' and factPeoplenum = '" & db_factPeoplenum_1 & "'"
		db.execute(SQL)
	end if

	'##2번

	if db_FRM2 = "ON" then	'수정또는 Insert
		
		if db_factPeoplenum_2 = "" then

			SQL = "select convert(varchar,convert(int,isnull(rtrim(max(convert(int,substring(factPeoplenum,2,10)))),'0'))+1) from armyinformix.dbo.factpeople where left(factPeoplenum,1) = 'M'"
	
			set RS = db.execute(SQL)
			if isnull(rs(0)) then
				db_factPeoplenum_2 = "M1"
			else
				db_factPeoplenum_2 = "M"&rs(0)
			end if
			'신규로 임의 입력건
			SQL = "INSERT INTO armyinformix.dbo.factpeople ( factPeoplenum, factnum"
			SQL = SQL & "		, SECTION2, NAME, level"
			SQL = SQL & "		,homephone, mobilephone, etcphone"
			SQL = SQL & "		,MONITORDATE, MONITORRESULT"
			SQL = SQL & "		,RESERVEDATE, Remark, Remark1, monitorpoint )"
			SQL = SQL & " VALUES ('" & db_factPeoplenum_2 & "', '" & db_RECEIPTFACTNUM & "'"
			SQL = SQL & "		,'" & db_SECTION2_2 & "','" & db_name_2 & "','" & db_level_2 & "'"
			SQL = SQL & "		,'" & db_homephone_2 & "','" & db_mobilephone_2 & "','" & db_etcphone_2 & "'"
			if db_MONITORDATE_2 = "" then
				SQL = SQL & "		,	getdate()"
			else
				SQL = SQL & "		,	'" & db_MONITORDATE_2 & "'"
			end if
			SQL = SQL & "		,'" & db_MONITORRESULT_2 & "'"

			if db_MONITORRESULT_2 = "4" then	'예약이라면
				if db_RESERVETIME_2 = "1" then
					SQL = SQL & "		,	CONVERT(CHAR(19),DATEADD(n,10,GETDATE()),121)"
				elseif db_RESERVETIME_2 = "2" then
					SQL = SQL & "		,	CONVERT(CHAR(19),DATEADD(n,30,GETDATE()),121)"
				elseif db_RESERVETIME_2 = "3" then
					SQL = SQL & "		,	CONVERT(CHAR(19),DATEADD(n,60,GETDATE()),121)"
				elseif db_RESERVETIME_2 = "4" then
					SQL = SQL & "		,	CONVERT(CHAR(19),DATEADD(n,120,GETDATE()),121)"
				elseif db_RESERVETIME_2 <> "" then
					SQL = SQL & "		,	'" & db_RESERVEDATE_2 & " " & db_RESERVETIME_2 & ":00:00'"
				else
					SQL = SQL & "		,	'" & db_RESERVEDATE_2 & " " & db_RESERVEHOUR_2 & ":"&  db_RESERVEMIN_2 &":00'"
				end if
			else
				SQL = SQL & "		,	''"
			end if
			SQL = SQL & "		,	'" & db_Remark_2 & "'"
			SQL = SQL & "		,	'" & db_Remark1_2 & "'"
			SQL = SQL & "		,	'" & db_TOT_2 & "')"

			db.execute(SQL)
	
		else

			SQL = " update armyinformix.dbo.factpeople set SECTION2 = '" & db_SECTION2_2 & "'"
			SQL = SQL & "		,	name = '" & db_name_2 & "',	level = '" & db_level_2 & "', homephone = '" & db_homephone_2 & "'"
			SQL = SQL & "		,	mobilephone = '" & db_mobilephone_2 & "', etcphone = '" & db_etcphone_2 & "'"

			if db_MONITORDATE_2 = "" then
				SQL = SQL & "		,	MONITORDATE = getdate()"
			else
				SQL = SQL & "		,	MONITORDATE = convert(datetime,'"&db_MONITORDATE_2&"')"
			end if
			SQL = SQL & "		,	MONITORRESULT = '" & db_MONITORRESULT_2 & "'"

			if db_MONITORRESULT_2 = "4" then	'예약이라면
				if db_RESERVETIME_2 = "1" then
					SQL = SQL & "		,	RESERVEDATE = CONVERT(CHAR(19),DATEADD(n,10,GETDATE()),121)"
				elseif db_RESERVETIME_2 = "2" then
					SQL = SQL & "		,	RESERVEDATE = CONVERT(CHAR(19),DATEADD(n,30,GETDATE()),121)"
				elseif db_RESERVETIME_2 = "3" then
					SQL = SQL & "		,	RESERVEDATE = CONVERT(CHAR(19),DATEADD(n,60,GETDATE()),121)"
				elseif db_RESERVETIME_2 = "4" then
					SQL = SQL & "		,	RESERVEDATE = CONVERT(CHAR(19),DATEADD(n,120,GETDATE()),121)"
				elseif db_RESERVETIME_2 <> "" then
					SQL = SQL & "		,	RESERVEDATE = '" & db_RESERVEDATE_2 & " " & db_RESERVETIME_2 & ":00:00'"
				else
					SQL = SQL & "		,	RESERVEDATE = '" & db_RESERVEDATE_2 & " " & db_RESERVEHOUR_2 & ":"&  db_RESERVEMIN_2 &":00'"
				end if
			else
				SQL = SQL & "		,	RESERVEDATE = ''"
			end if
			SQL = SQL & "		,	Remark = '" & db_Remark_2 & "'"
			SQL = SQL & "		,	Remark1 = '" & db_Remark1_2 & "'"
			SQL = SQL & "		,	monitorpoint = '" & db_TOT_2 & "'"
			SQL = SQL & "	where factPeoplenum = '" & db_factPeoplenum_2 & "' and factnum = '" & db_RECEIPTFACTNUM & "'"

			db.execute(SQL)
		end if

		'## 2번
		SQL = "	delete from armyinformix.dbo.monitor where factnum = '" & db_RECEIPTFACTNUM & "' and factPeoplenum = '" & db_factPeoplenum_2 & "'"
		db.execute(SQL)
		for i = 1 to icnt
			point9 =  Request("QUESTION_2"&i)
			if point9 = "9" then
				ipoint9 = "1"
			else
				ipoint9 = "0"
			end if
			if point9 = "8" then
				ipoint8 = "1"
			else
				ipoint8 = "0"
			end if
			if point9 = "7" then
				ipoint7 = "1"
			else
				ipoint7 = "0"
			end if
			pointplus =  Request("QUESTIONP_2"&i)
			if pointplus = "" then
				pointplus = "0"
			end if
			totpoint =  Request("POINT_2"&i)
			if totpoint = "" then
				totpoint = "0"
			end if

			SQL = "insert into armyinformix.dbo.monitor ( factnum, factPeoplenum, seqno, point9, point8, point7, pointplus, totpoint, monitordate, monitoruser)"
			SQL = SQL & " values ( '" & db_RECEIPTFACTNUM & "', '" & db_factPeoplenum_2 & "', " & i & ", " & ipoint9& ", " & ipoint8& ", " & ipoint7
			SQL = SQL & " , " & pointplus & ","&totpoint&",getdate(),'" &INCODE&"')"
			db.execute(SQL)
			'Response.write SQL
		next
	else
		'## 2번
		SQL = "delete from armyinformix.dbo.factpeople"
		SQL = SQL & "	where factPeoplenum = '" & db_factPeoplenum_2 & "' and factnum = '" & db_RECEIPTFACTNUM & "'"
		db.execute(SQL)

		SQL = "	delete from armyinformix.dbo.monitor where factnum = '" & db_RECEIPTFACTNUM & "' and factPeoplenum = '" & db_factPeoplenum_2 & "'"
		db.execute(SQL)
	end if



	if db_FRM3 = "ON" then	'수정또는 Insert
		
		if db_factPeoplenum_3 = "" then

			SQL = "select convert(varchar,convert(int,isnull(rtrim(max(convert(int,substring(factPeoplenum,2,10)))),'0'))+1) from armyinformix.dbo.factpeople where left(factPeoplenum,1) = 'M'"
	
			set RS = db.execute(SQL)
			if isnull(rs(0)) then
				db_factPeoplenum_3 = "M1"
			else
				db_factPeoplenum_3 = "M"&rs(0)
			end if
			'신규로 임의 입력건
			SQL = "INSERT INTO armyinformix.dbo.factpeople ( factPeoplenum, factnum"
			SQL = SQL & "		, SECTION2, NAME, level"
			SQL = SQL & "		,homephone, mobilephone, etcphone"
			SQL = SQL & "		,MONITORDATE, MONITORRESULT"
			SQL = SQL & "		,RESERVEDATE, Remark, Remark1, monitorpoint )"
			SQL = SQL & " VALUES ('" & db_factPeoplenum_3 & "' , '" & db_RECEIPTFACTNUM & "'"
			SQL = SQL & "		,'" & db_SECTION2_3 & "','" & db_name_3 & "','" & db_level_3 & "'"
			SQL = SQL & "		,'" & db_homephone_3 & "','" & db_mobilephone_3 & "','" & db_etcphone_3 & "'"
			if db_MONITORDATE_3 = "" then
				SQL = SQL & "		,	getdate()"
			else
				SQL = SQL & "		,	'" & db_MONITORDATE_3 & "'"
			end if
			SQL = SQL & "		,'" & db_MONITORRESULT_3 & "'"

			if db_MONITORRESULT_3 = "4" then	'예약이라면
				if db_RESERVETIME_3 = "1" then
					SQL = SQL & "		,	CONVERT(CHAR(19),DATEADD(n,10,GETDATE()),121)"
				elseif db_RESERVETIME_3 = "2" then
					SQL = SQL & "		,	CONVERT(CHAR(19),DATEADD(n,30,GETDATE()),121)"
				elseif db_RESERVETIME_3 = "3" then
					SQL = SQL & "		,	CONVERT(CHAR(19),DATEADD(n,60,GETDATE()),121)"
				elseif db_RESERVETIME_3 = "4" then
					SQL = SQL & "		,	CONVERT(CHAR(19),DATEADD(n,120,GETDATE()),121)"
				elseif db_RESERVETIME_3 <> "" then
					SQL = SQL & "		,	'" & db_RESERVEDATE_3 & " " & db_RESERVETIME_3 & ":00:00'"
				else
					SQL = SQL & "		,	'" & db_RESERVEDATE_3 & " " & db_RESERVEHOUR_3 & ":"&  db_RESERVEMIN_3 &":00'"
				end if
			else
				SQL = SQL & "		,	''"
			end if
			SQL = SQL & "		,	'" & db_Remark_3 & "'"
			SQL = SQL & "		,	'" & db_Remark1_3 & "'"
			SQL = SQL & "		,	'" & db_TOT_3 & "')"

			db.execute(SQL)
	
		else

			SQL = " update armyinformix.dbo.factpeople set SECTION2 = '" & db_SECTION2_3 & "'"
			SQL = SQL & "		,	name = '" & db_name_3 & "',	level = '" & db_level_3 & "', homephone = '" & db_homephone_3 & "'"
			SQL = SQL & "		,	mobilephone = '" & db_mobilephone_3 & "', etcphone = '" & db_etcphone_3 & "'"

			if db_MONITORDATE_3 = "" then
				SQL = SQL & "		,	MONITORDATE = getdate()"
			else
				SQL = SQL & "		,	MONITORDATE = convert(datetime,'"&db_MONITORDATE_3&"')"
			end if
			SQL = SQL & "		,	MONITORRESULT = '" & db_MONITORRESULT_3 & "'"

			if db_MONITORRESULT_3 = "4" then	'예약이라면
				if db_RESERVETIME_3 = "1" then
					SQL = SQL & "		,	RESERVEDATE = CONVERT(CHAR(19),DATEADD(n,10,GETDATE()),121)"
				elseif db_RESERVETIME_3 = "2" then
					SQL = SQL & "		,	RESERVEDATE = CONVERT(CHAR(19),DATEADD(n,30,GETDATE()),121)"
				elseif db_RESERVETIME_3 = "3" then
					SQL = SQL & "		,	RESERVEDATE = CONVERT(CHAR(19),DATEADD(n,60,GETDATE()),121)"
				elseif db_RESERVETIME_3 = "4" then
					SQL = SQL & "		,	RESERVEDATE = CONVERT(CHAR(19),DATEADD(n,120,GETDATE()),121)"
				elseif db_RESERVETIME_3 <> "" then
					SQL = SQL & "		,	RESERVEDATE = '" & db_RESERVEDATE_3 & " " & db_RESERVETIME_3 & ":00:00'"
				else
					SQL = SQL & "		,	RESERVEDATE = '" & db_RESERVEDATE_3 & " " & db_RESERVEHOUR_3 & ":"&  db_RESERVEMIN_3 &":00'"
				end if
			else
				SQL = SQL & "		,	RESERVEDATE = ''"
			end if
			SQL = SQL & "		,	Remark = '" & db_Remark_3 & "'"
			SQL = SQL & "		,	Remark1 = '" & db_Remark1_3 & "'"
			SQL = SQL & "		,	monitorpoint = '" & db_TOT_3 & "'"
			SQL = SQL & "	where factPeoplenum = '" & db_factPeoplenum_3 & "' and factnum = '" & db_RECEIPTFACTNUM & "'"

			db.execute(SQL)
		end if

		'## 1번
		SQL = "	delete from armyinformix.dbo.monitor where factnum = '" & db_RECEIPTFACTNUM & "' and factPeoplenum = '" & db_factPeoplenum_3 & "'"
		db.execute(SQL)
		for i = 1 to icnt
			point9 =  Request("QUESTION_3"&i)
			if point9 = "9" then
				ipoint9 = "1"
			else
				ipoint9 = "0"
			end if
			if point9 = "8" then
				ipoint8 = "1"
			else
				ipoint8 = "0"
			end if
			if point9 = "7" then
				ipoint7 = "1"
			else
				ipoint7 = "0"
			end if
			pointplus =  Request("QUESTIONP_3"&i)
			if pointplus = "" then
				pointplus = "0"
			end if
			totpoint =  Request("POINT_3"&i)
			if totpoint = "" then
				totpoint = "0"
			end if

			SQL = "insert into armyinformix.dbo.monitor ( factnum, factPeoplenum, seqno, point9, point8, point7, pointplus, totpoint, monitordate, monitoruser)"
			SQL = SQL & " values ( '" & db_RECEIPTFACTNUM & "', '" & db_factPeoplenum_3 & "', " & i & ", " & ipoint9& ", " & ipoint8& ", " & ipoint7
			SQL = SQL & " , " & pointplus & ","&totpoint&",getdate(),'" &INCODE&"')"
			db.execute(SQL)
			'Response.write SQL
		next
	else
		'## 1번

		SQL = "delete from armyinformix.dbo.factpeople"
		SQL = SQL & "	where factPeoplenum = '" & db_factPeoplenum_3 & "' and factnum = '" & db_RECEIPTFACTNUM & "'"
		db.execute(SQL)

		SQL = "	delete from armyinformix.dbo.monitor where factnum = '" & db_RECEIPTFACTNUM & "' and factPeoplenum = '" & db_factPeoplenum_3 & "'"
		db.execute(SQL)
	end if

	'## 4번-------------------------------------------------------------------------------------------------------------------
	if db_FRM4 = "ON" then	'수정또는 Insert
		
		if db_factPeoplenum_4 = "" then

			SQL = "select convert(varchar,convert(int,isnull(rtrim(max(convert(int,substring(factPeoplenum,2,10)))),'0'))+1) from armyinformix.dbo.factpeople where left(factPeoplenum,1) = 'M'"
	
			set RS = db.execute(SQL)
			if isnull(rs(0)) then
				db_factPeoplenum_4 = "M1"
			else
				db_factPeoplenum_4 = "M"&rs(0)
			end if
			'신규로 임의 입력건
			SQL = "INSERT INTO armyinformix.dbo.factpeople ( factPeoplenum, factnum"
			SQL = SQL & "		, SECTION2, NAME, level"
			SQL = SQL & "		,homephone, mobilephone, etcphone"
			SQL = SQL & "		,MONITORDATE, MONITORRESULT"
			SQL = SQL & "		,RESERVEDATE, Remark, Remark1, monitorpoint )"
			SQL = SQL & " VALUES ('" & db_factPeoplenum_4 & "' , '" & db_RECEIPTFACTNUM & "'"
			SQL = SQL & "		,'" & db_SECTION2_4 & "','" & db_name_4 & "','" & db_level_4 & "'"
			SQL = SQL & "		,'" & db_homephone_4 & "','" & db_mobilephone_4 & "','" & db_etcphone_4 & "'"
			if db_MONITORDATE_4 = "" then
				SQL = SQL & "		,	getdate()"
			else
				SQL = SQL & "		,	'" & db_MONITORDATE_4 & "'"
			end if
			SQL = SQL & "		,'" & db_MONITORRESULT_4 & "'"

			if db_MONITORRESULT_4 = "4" then	'예약이라면
				if db_RESERVETIME_4 = "1" then
					SQL = SQL & "		,	CONVERT(CHAR(19),DATEADD(n,10,GETDATE()),121)"
				elseif db_RESERVETIME_4 = "2" then
					SQL = SQL & "		,	CONVERT(CHAR(19),DATEADD(n,30,GETDATE()),121)"
				elseif db_RESERVETIME_4 = "3" then
					SQL = SQL & "		,	CONVERT(CHAR(19),DATEADD(n,60,GETDATE()),121)"
				elseif db_RESERVETIME_4 = "4" then
					SQL = SQL & "		,	CONVERT(CHAR(19),DATEADD(n,120,GETDATE()),121)"
				elseif db_RESERVETIME_4 <> "" then
					SQL = SQL & "		,	'" & db_RESERVEDATE_4 & " " & db_RESERVETIME_4 & ":00:00'"
				else
					SQL = SQL & "		,	'" & db_RESERVEDATE_4 & " " & db_RESERVEHOUR_4 & ":"&  db_RESERVEMIN_4 &":00'"
				end if
			else
				SQL = SQL & "		,	''"
			end if
			SQL = SQL & "		,	'" & db_Remark_4 & "'"
			SQL = SQL & "		,	'" & db_Remark1_4 & "'"
			SQL = SQL & "		,	'" & db_TOT_4 & "')"

			db.execute(SQL)
	
		else

			SQL = " update armyinformix.dbo.factpeople set SECTION2 = '" & db_SECTION2_4 & "'"
			SQL = SQL & "		,	name = '" & db_name_4 & "',	level = '" & db_level_4 & "', homephone = '" & db_homephone_4 & "'"
			SQL = SQL & "		,	mobilephone = '" & db_mobilephone_4 & "', etcphone = '" & db_etcphone_4 & "'"

			if db_MONITORDATE_4 = "" then
				SQL = SQL & "		,	MONITORDATE = getdate()"
			else
				SQL = SQL & "		,	MONITORDATE = convert(datetime,'"&db_MONITORDATE_4&"')"
			end if
			SQL = SQL & "		,	MONITORRESULT = '" & db_MONITORRESULT_4 & "'"

			if db_MONITORRESULT_4 = "4" then	'예약이라면
				if db_RESERVETIME_4 = "1" then
					SQL = SQL & "		,	RESERVEDATE = CONVERT(CHAR(19),DATEADD(n,10,GETDATE()),121)"
				elseif db_RESERVETIME_4 = "2" then
					SQL = SQL & "		,	RESERVEDATE = CONVERT(CHAR(19),DATEADD(n,30,GETDATE()),121)"
				elseif db_RESERVETIME_4 = "3" then
					SQL = SQL & "		,	RESERVEDATE = CONVERT(CHAR(19),DATEADD(n,60,GETDATE()),121)"
				elseif db_RESERVETIME_4 = "4" then
					SQL = SQL & "		,	RESERVEDATE = CONVERT(CHAR(19),DATEADD(n,120,GETDATE()),121)"
				elseif db_RESERVETIME_4 <> "" then
					SQL = SQL & "		,	RESERVEDATE = '" & db_RESERVEDATE_4 & " " & db_RESERVETIME_4 & ":00:00'"
				else
					SQL = SQL & "		,	RESERVEDATE = '" & db_RESERVEDATE_4 & " " & db_RESERVEHOUR_4 & ":"&  db_RESERVEMIN_4 &":00'"
				end if
			else
				SQL = SQL & "		,	RESERVEDATE = ''"
			end if
			SQL = SQL & "		,	Remark = '" & db_Remark_4 & "'"
			SQL = SQL & "		,	Remark1 = '" & db_Remark1_4 & "'"
			SQL = SQL & "		,	monitorpoint = '" & db_TOT_4 & "'"
			SQL = SQL & "	where factPeoplenum = '" & db_factPeoplenum_4 & "' and factnum = '" & db_RECEIPTFACTNUM & "'"

	'response.write SQL

			db.execute(SQL)
		end if

		'## 1번
		SQL = "	delete from armyinformix.dbo.monitor where factnum = '" & db_RECEIPTFACTNUM & "' and factPeoplenum = '" & db_factPeoplenum_4 & "'"
		db.execute(SQL)
		for i = 1 to icnt
			point9 =  Request("QUESTION_4"&i)
			if point9 = "9" then
				ipoint9 = "1"
			else
				ipoint9 = "0"
			end if
			if point9 = "8" then
				ipoint8 = "1"
			else
				ipoint8 = "0"
			end if
			if point9 = "7" then
				ipoint7 = "1"
			else
				ipoint7 = "0"
			end if
			pointplus =  Request("QUESTIONP_4"&i)
			if pointplus = "" then
				pointplus = "0"
			end if
			totpoint =  Request("POINT_4"&i)
			if totpoint = "" then
				totpoint = "0"
			end if

			SQL = "insert into armyinformix.dbo.monitor ( factnum, factPeoplenum, seqno, point9, point8, point7, pointplus, totpoint, monitordate, monitoruser)"
			SQL = SQL & " values ( '" & db_RECEIPTFACTNUM & "', '" & db_factPeoplenum_4 & "', " & i & ", " & ipoint9& ", " & ipoint8& ", " & ipoint7
			SQL = SQL & " , " & pointplus & ","&totpoint&",getdate(),'" &INCODE&"')"
			db.execute(SQL)
			'Response.write SQL
		next
	else
		'## 1번
		SQL = "delete from armyinformix.dbo.factpeople"
		SQL = SQL & "	where factPeoplenum = '" & db_factPeoplenum_4 & "' and factnum = '" & db_RECEIPTFACTNUM & "'"
		db.execute(SQL)

		SQL = "	delete from armyinformix.dbo.monitor where factnum = '" & db_RECEIPTFACTNUM & "' and factPeoplenum = '" & db_factPeoplenum_4 & "'"
		db.execute(SQL)
	end if



	'## 5번-------------------------------------------------------------------------------------------------------------------
	if db_FRM5 = "ON" then	'수정또는 Insert
		
		if db_factPeoplenum_5 = "" then

			SQL = "select convert(varchar,convert(int,isnull(rtrim(max(convert(int,substring(factPeoplenum,2,10)))),'0'))+1) from armyinformix.dbo.factpeople where left(factPeoplenum,1) = 'M'"
	
			set RS = db.execute(SQL)
			if isnull(rs(0)) then
				db_factPeoplenum_5 = "M1"
			else
				db_factPeoplenum_5 = "M"&rs(0)
			end if
			'신규로 임의 입력건
			SQL = "INSERT INTO armyinformix.dbo.factpeople ( factPeoplenum, factnum"
			SQL = SQL & "		, SECTION2, NAME, level"
			SQL = SQL & "		,homephone, mobilephone, etcphone"
			SQL = SQL & "		,MONITORDATE, MONITORRESULT"
			SQL = SQL & "		,RESERVEDATE, Remark,Remark1, monitorpoint )"
			SQL = SQL & " VALUES ('" & db_factPeoplenum_5 & "' , '" & db_RECEIPTFACTNUM & "'"
			SQL = SQL & "		,'" & db_SECTION2_5 & "','" & db_name_5 & "','" & db_level_5 & "'"
			SQL = SQL & "		,'" & db_homephone_5 & "','" & db_mobilephone_5 & "','" & db_etcphone_5 & "'"
			if db_MONITORDATE_5 = "" then
				SQL = SQL & "		,	getdate()"
			else
				SQL = SQL & "		,	'" & db_MONITORDATE_5 & "'"
			end if
			SQL = SQL & "		,'" & db_MONITORRESULT_5 & "'"

			if db_MONITORRESULT_5 = "4" then	'예약이라면
				if db_RESERVETIME_5 = "1" then
					SQL = SQL & "		,	CONVERT(CHAR(19),DATEADD(n,10,GETDATE()),121)"
				elseif db_RESERVETIME_5 = "2" then
					SQL = SQL & "		,	CONVERT(CHAR(19),DATEADD(n,30,GETDATE()),121)"
				elseif db_RESERVETIME_5 = "3" then
					SQL = SQL & "		,	CONVERT(CHAR(19),DATEADD(n,60,GETDATE()),121)"
				elseif db_RESERVETIME_5 = "4" then
					SQL = SQL & "		,	CONVERT(CHAR(19),DATEADD(n,120,GETDATE()),121)"
				elseif db_RESERVETIME_5 <> "" then
					SQL = SQL & "		,	'" & db_RESERVEDATE_5 & " " & db_RESERVETIME_5 & ":00:00'"
				else
					SQL = SQL & "		,	'" & db_RESERVEDATE_5 & " " & db_RESERVEHOUR_5 & ":"&  db_RESERVEMIN_5 &":00'"
				end if
			else
				SQL = SQL & "		,	''"
			end if
			SQL = SQL & "		,	'" & db_Remark_5 & "'"
			SQL = SQL & "		,	'" & db_Remark1_5 & "'"
			SQL = SQL & "		,	'" & db_TOT_5 & "')"

			db.execute(SQL)
	
		else

			SQL = " update armyinformix.dbo.factpeople set SECTION2 = '" & db_SECTION2_5 & "'"
			SQL = SQL & "		,	name = '" & db_name_5 & "',	level = '" & db_level_5 & "', homephone = '" & db_homephone_5 & "'"
			SQL = SQL & "		,	mobilephone = '" & db_mobilephone_5 & "', etcphone = '" & db_etcphone_5 & "'"

			if db_MONITORDATE_5 = "" then
				SQL = SQL & "		,	MONITORDATE = getdate()"
			else
				SQL = SQL & "		,	MONITORDATE = convert(datetime,'"&db_MONITORDATE_5&"')"
			end if
			SQL = SQL & "		,	MONITORRESULT = '" & db_MONITORRESULT_5 & "'"

			if db_MONITORRESULT_5 = "4" then	'예약이라면
				if db_RESERVETIME_5 = "1" then
					SQL = SQL & "		,	RESERVEDATE = CONVERT(CHAR(19),DATEADD(n,10,GETDATE()),121)"
				elseif db_RESERVETIME_5 = "2" then
					SQL = SQL & "		,	RESERVEDATE = CONVERT(CHAR(19),DATEADD(n,30,GETDATE()),121)"
				elseif db_RESERVETIME_5 = "3" then
					SQL = SQL & "		,	RESERVEDATE = CONVERT(CHAR(19),DATEADD(n,60,GETDATE()),121)"
				elseif db_RESERVETIME_5 = "4" then
					SQL = SQL & "		,	RESERVEDATE = CONVERT(CHAR(19),DATEADD(n,120,GETDATE()),121)"
				elseif db_RESERVETIME_5 <> "" then
					SQL = SQL & "		,	RESERVEDATE = '" & db_RESERVEDATE_5 & " " & db_RESERVETIME_5 & ":00:00'"
				else
					SQL = SQL & "		,	RESERVEDATE = '" & db_RESERVEDATE_5 & " " & db_RESERVEHOUR_5 & ":"&  db_RESERVEMIN_5 &":00'"
				end if
			else
				SQL = SQL & "		,	RESERVEDATE = ''"
			end if
			SQL = SQL & "		,	Remark = '" & db_Remark_5 & "'"
			SQL = SQL & "		,	Remark1 = '" & db_Remark1_5 & "'"
			SQL = SQL & "		,	monitorpoint = '" & db_TOT_5 & "'"
			SQL = SQL & "	where factPeoplenum = '" & db_factPeoplenum_5 & "' and factnum = '" & db_RECEIPTFACTNUM & "'"

			response.write SQL

			db.execute(SQL)
		end if

		'## 1번
		SQL = "	delete from armyinformix.dbo.monitor where factnum = '" & db_RECEIPTFACTNUM & "' and factPeoplenum = '" & db_factPeoplenum_5 & "'"
		db.execute(SQL)
		for i = 1 to icnt
			point9 =  Request("QUESTION_5"&i)
			if point9 = "9" then
				ipoint9 = "1"
			else
				ipoint9 = "0"
			end if
			if point9 = "8" then
				ipoint8 = "1"
			else
				ipoint8 = "0"
			end if
			if point9 = "7" then
				ipoint7 = "1"
			else
				ipoint7 = "0"
			end if
			pointplus =  Request("QUESTIONP_5"&i)
			if pointplus = "" then
				pointplus = "0"
			end if
			totpoint =  Request("POINT_5"&i)
			if totpoint = "" then
				totpoint = "0"
			end if

			SQL = "insert into armyinformix.dbo.monitor ( factnum, factPeoplenum, seqno, point9, point8, point7, pointplus, totpoint, monitordate, monitoruser)"
			SQL = SQL & " values ( '" & db_RECEIPTFACTNUM & "', '" & db_factPeoplenum_5 & "', " & i & ", " & ipoint9& ", " & ipoint8& ", " & ipoint7
			SQL = SQL & " , " & pointplus & ","&totpoint&",getdate(),'" &INCODE&"')"
			db.execute(SQL)
			'Response.write SQL
		next
	else
		'## 1번
		SQL = "delete from armyinformix.dbo.factpeople"
		SQL = SQL & "	where factPeoplenum = '" & db_factPeoplenum_5 & "' and factnum = '" & db_RECEIPTFACTNUM & "'"
		db.execute(SQL)

		SQL = "	delete from armyinformix.dbo.monitor where factnum = '" & db_RECEIPTFACTNUM & "' and factPeoplenum = '" & db_factPeoplenum_5 & "'"
		db.execute(SQL)
	end if


	'## 6번-------------------------------------------------------------------------------------------------------------------
	if db_FRM6 = "ON" then	'수정또는 Insert
		
		if db_factPeoplenum_6 = "" then

			SQL = "select convert(varchar,convert(int,isnull(rtrim(max(convert(int,substring(factPeoplenum,2,10)))),'0'))+1) from armyinformix.dbo.factpeople where left(factPeoplenum,1) = 'M'"
	
			set RS = db.execute(SQL)
			if isnull(rs(0)) then
				db_factPeoplenum_6 = "M1"
			else
				db_factPeoplenum_6 = "M"&rs(0)
			end if
			'신규로 임의 입력건
			SQL = "INSERT INTO armyinformix.dbo.factpeople ( factPeoplenum, factnum"
			SQL = SQL & "		, SECTION2, NAME, level"
			SQL = SQL & "		,homephone, mobilephone, etcphone"
			SQL = SQL & "		,MONITORDATE, MONITORRESULT"
			SQL = SQL & "		,RESERVEDATE, Remark,Remark1, monitorpoint )"
			SQL = SQL & " VALUES ('" & db_factPeoplenum_6 & "' , '" & db_RECEIPTFACTNUM & "'"
			SQL = SQL & "		,'" & db_SECTION2_6 & "','" & db_name_6 & "','" & db_level_6 & "'"
			SQL = SQL & "		,'" & db_homephone_6 & "','" & db_mobilephone_6 & "','" & db_etcphone_6 & "'"
			if db_MONITORDATE_6 = "" then
				SQL = SQL & "		,	getdate()"
			else
				SQL = SQL & "		,	'" & db_MONITORDATE_6 & "'"
			end if
			SQL = SQL & "		,'" & db_MONITORRESULT_6 & "'"

			if db_MONITORRESULT_6 = "4" then	'예약이라면
				if db_RESERVETIME_6 = "1" then
					SQL = SQL & "		,	CONVERT(CHAR(19),DATEADD(n,10,GETDATE()),121)"
				elseif db_RESERVETIME_6 = "2" then
					SQL = SQL & "		,	CONVERT(CHAR(19),DATEADD(n,30,GETDATE()),121)"
				elseif db_RESERVETIME_6 = "3" then
					SQL = SQL & "		,	CONVERT(CHAR(19),DATEADD(n,60,GETDATE()),121)"
				elseif db_RESERVETIME_6 = "4" then
					SQL = SQL & "		,	CONVERT(CHAR(19),DATEADD(n,120,GETDATE()),121)"
				elseif db_RESERVETIME_6 <> "" then
					SQL = SQL & "		,	'" & db_RESERVEDATE_6 & " " & db_RESERVETIME_6 & ":00:00'"
				else
					SQL = SQL & "		,	'" & db_RESERVEDATE_6 & " " & db_RESERVEHOUR_6 & ":"&  db_RESERVEMIN_6 &":00'"
				end if
			else
				SQL = SQL & "		,	''"
			end if
			SQL = SQL & "		,	'" & db_Remark_6 & "'"
			SQL = SQL & "		,	'" & db_Remark1_6 & "'"
			SQL = SQL & "		,	'" & db_TOT_6 & "')"

			db.execute(SQL)
	
		else

			SQL = " update armyinformix.dbo.factpeople set SECTION2 = '" & db_SECTION2_6 & "'"
			SQL = SQL & "		,	name = '" & db_name_6 & "',	level = '" & db_level_6 & "', homephone = '" & db_homephone_6 & "'"
			SQL = SQL & "		,	mobilephone = '" & db_mobilephone_6 & "', etcphone = '" & db_etcphone_6 & "'"

			if db_MONITORDATE_6 = "" then
				SQL = SQL & "		,	MONITORDATE = getdate()"
			else
				SQL = SQL & "		,	MONITORDATE = convert(datetime,'"&db_MONITORDATE_6&"')"
			end if
			SQL = SQL & "		,	MONITORRESULT = '" & db_MONITORRESULT_6 & "'"

			if db_MONITORRESULT_6 = "4" then	'예약이라면
				if db_RESERVETIME_6 = "1" then
					SQL = SQL & "		,	RESERVEDATE = CONVERT(CHAR(19),DATEADD(n,10,GETDATE()),121)"
				elseif db_RESERVETIME_6 = "2" then
					SQL = SQL & "		,	RESERVEDATE = CONVERT(CHAR(19),DATEADD(n,30,GETDATE()),121)"
				elseif db_RESERVETIME_6 = "3" then
					SQL = SQL & "		,	RESERVEDATE = CONVERT(CHAR(19),DATEADD(n,60,GETDATE()),121)"
				elseif db_RESERVETIME_6 = "4" then
					SQL = SQL & "		,	RESERVEDATE = CONVERT(CHAR(19),DATEADD(n,120,GETDATE()),121)"
				elseif db_RESERVETIME_6 <> "" then
					SQL = SQL & "		,	RESERVEDATE = '" & db_RESERVEDATE_6 & " " & db_RESERVETIME_6 & ":00:00'"
				else
					SQL = SQL & "		,	RESERVEDATE = '" & db_RESERVEDATE_6 & " " & db_RESERVEHOUR_6 & ":"&  db_RESERVEMIN_6 &":00'"
				end if
			else
				SQL = SQL & "		,	RESERVEDATE = ''"
			end if
			SQL = SQL & "		,	Remark = '" & db_Remark_6 & "'"
			SQL = SQL & "		,	Remark1 = '" & db_Remark1_6 & "'"
			SQL = SQL & "		,	monitorpoint = '" & db_TOT_6 & "'"
			SQL = SQL & "	where factPeoplenum = '" & db_factPeoplenum_6 & "' and factnum = '" & db_RECEIPTFACTNUM & "'"

			db.execute(SQL)
		end if

		'## 1번
		SQL = "	delete from armyinformix.dbo.monitor where factnum = '" & db_RECEIPTFACTNUM & "' and factPeoplenum = '" & db_factPeoplenum_6 & "'"
		db.execute(SQL)
		for i = 1 to icnt
			point9 =  Request("QUESTION_6"&i)
			if point9 = "9" then
				ipoint9 = "1"
			else
				ipoint9 = "0"
			end if
			if point9 = "8" then
				ipoint8 = "1"
			else
				ipoint8 = "0"
			end if
			if point9 = "7" then
				ipoint7 = "1"
			else
				ipoint7 = "0"
			end if
			pointplus =  Request("QUESTIONP_6"&i)
			if pointplus = "" then
				pointplus = "0"
			end if
			totpoint =  Request("POINT_6"&i)
			if totpoint = "" then
				totpoint = "0"
			end if

			SQL = "insert into armyinformix.dbo.monitor ( factnum, factPeoplenum, seqno, point9, point8, point7, pointplus, totpoint, monitordate, monitoruser)"
			SQL = SQL & " values ( '" & db_RECEIPTFACTNUM & "', '" & db_factPeoplenum_6 & "', " & i & ", " & ipoint9& ", " & ipoint8& ", " & ipoint7
			SQL = SQL & " , " & pointplus & ","&totpoint&",getdate(),'" &INCODE&"')"
			db.execute(SQL)
			'Response.write SQL
		next
	else
		'## 6번
		SQL = "delete from armyinformix.dbo.factpeople"
		SQL = SQL & "	where factPeoplenum = '" & db_factPeoplenum_6 & "' and factnum = '" & db_RECEIPTFACTNUM & "'"
		db.execute(SQL)

		SQL = "	delete from armyinformix.dbo.monitor where factnum = '" & db_RECEIPTFACTNUM & "' and factPeoplenum = '" & db_factPeoplenum_6 & "'"
		db.execute(SQL)
	end if


	'## 7번-------------------------------------------------------------------------------------------------------------------
	if db_FRM7 = "ON" then	'수정또는 Insert
		
		if db_factPeoplenum_7 = "" then

			SQL = "select convert(varchar,convert(int,isnull(rtrim(max(convert(int,substring(factPeoplenum,2,10)))),'0'))+1) from armyinformix.dbo.factpeople where left(factPeoplenum,1) = 'M'"
	
			set RS = db.execute(SQL)
			if isnull(rs(0)) then
				db_factPeoplenum_7 = "M1"
			else
				db_factPeoplenum_7 = "M"&rs(0)
			end if
			'신규로 임의 입력건
			SQL = "INSERT INTO armyinformix.dbo.factpeople ( factPeoplenum, factnum"
			SQL = SQL & "		, SECTION2, NAME, level"
			SQL = SQL & "		,homephone, mobilephone, etcphone"
			SQL = SQL & "		,MONITORDATE, MONITORRESULT"
			SQL = SQL & "		,RESERVEDATE, Remark,Remark1, monitorpoint )"
			SQL = SQL & " VALUES ('" & db_factPeoplenum_7 & "' , '" & db_RECEIPTFACTNUM & "'"
			SQL = SQL & "		,'" & db_SECTION2_7 & "','" & db_name_7 & "','" & db_level_7 & "'"
			SQL = SQL & "		,'" & db_homephone_7 & "','" & db_mobilephone_7 & "','" & db_etcphone_7 & "'"
			if db_MONITORDATE_7 = "" then
				SQL = SQL & "		,	getdate()"
			else
				SQL = SQL & "		,	'" & db_MONITORDATE_7 & "'"
			end if
			SQL = SQL & "		,'" & db_MONITORRESULT_7 & "'"

			if db_MONITORRESULT_7 = "4" then	'예약이라면
				if db_RESERVETIME_7 = "1" then
					SQL = SQL & "		,	CONVERT(CHAR(19),DATEADD(n,10,GETDATE()),121)"
				elseif db_RESERVETIME_7 = "2" then
					SQL = SQL & "		,	CONVERT(CHAR(19),DATEADD(n,30,GETDATE()),121)"
				elseif db_RESERVETIME_7 = "3" then
					SQL = SQL & "		,	CONVERT(CHAR(19),DATEADD(n,60,GETDATE()),121)"
				elseif db_RESERVETIME_7 = "4" then
					SQL = SQL & "		,	CONVERT(CHAR(19),DATEADD(n,120,GETDATE()),121)"
				elseif db_RESERVETIME_7 <> "" then
					SQL = SQL & "		,	'" & db_RESERVEDATE_7 & " " & db_RESERVETIME_7 & ":00:00'"
				else
					SQL = SQL & "		,	'" & db_RESERVEDATE_7 & " " & db_RESERVEHOUR_7 & ":"&  db_RESERVEMIN_7 &":00'"
				end if
			else
				SQL = SQL & "		,	''"
			end if
			SQL = SQL & "		,	'" & db_Remark_7 & "'"
			SQL = SQL & "		,	'" & db_Remark1_7 & "'"
			SQL = SQL & "		,	'" & db_TOT_7 & "')"

			db.execute(SQL)
	
		else

			SQL = " update armyinformix.dbo.factpeople set SECTION2 = '" & db_SECTION2_7 & "'"
			SQL = SQL & "		,	name = '" & db_name_7 & "',	level = '" & db_level_7 & "', homephone = '" & db_homephone_7 & "'"
			SQL = SQL & "		,	mobilephone = '" & db_mobilephone_7 & "', etcphone = '" & db_etcphone_7 & "'"

			if db_MONITORDATE_7 = "" then
				SQL = SQL & "		,	MONITORDATE = getdate()"
			else
				SQL = SQL & "		,	MONITORDATE = convert(datetime,'"&db_MONITORDATE_7&"')"
			end if
			SQL = SQL & "		,	MONITORRESULT = '" & db_MONITORRESULT_7 & "'"

			if db_MONITORRESULT_7 = "4" then	'예약이라면
				if db_RESERVETIME_7 = "1" then
					SQL = SQL & "		,	RESERVEDATE = CONVERT(CHAR(19),DATEADD(n,10,GETDATE()),121)"
				elseif db_RESERVETIME_7 = "2" then
					SQL = SQL & "		,	RESERVEDATE = CONVERT(CHAR(19),DATEADD(n,30,GETDATE()),121)"
				elseif db_RESERVETIME_7 = "3" then
					SQL = SQL & "		,	RESERVEDATE = CONVERT(CHAR(19),DATEADD(n,60,GETDATE()),121)"
				elseif db_RESERVETIME_7 = "4" then
					SQL = SQL & "		,	RESERVEDATE = CONVERT(CHAR(19),DATEADD(n,120,GETDATE()),121)"
				elseif db_RESERVETIME_7 <> "" then
					SQL = SQL & "		,	RESERVEDATE = '" & db_RESERVEDATE_7 & " " & db_RESERVETIME_7 & ":00:00'"
				else
					SQL = SQL & "		,	RESERVEDATE = '" & db_RESERVEDATE_7 & " " & db_RESERVEHOUR_7 & ":"&  db_RESERVEMIN_7 &":00'"
				end if
			else
				SQL = SQL & "		,	RESERVEDATE = ''"
			end if
			SQL = SQL & "		,	Remark = '" & db_Remark_7 & "'"
			SQL = SQL & "		,	Remark1 = '" & db_Remark1_7 & "'"
			SQL = SQL & "		,	monitorpoint = '" & db_TOT_7 & "'"
			SQL = SQL & "	where factPeoplenum = '" & db_factPeoplenum_7 & "' and factnum = '" & db_RECEIPTFACTNUM & "'"

			db.execute(SQL)
		end if

		'## 1번
		SQL = "	delete from armyinformix.dbo.monitor where factnum = '" & db_RECEIPTFACTNUM & "' and factPeoplenum = '" & db_factPeoplenum_7 & "'"
		db.execute(SQL)
		for i = 1 to icnt
			point9 =  Request("QUESTION_7"&i)
			if point9 = "9" then
				ipoint9 = "1"
			else
				ipoint9 = "0"
			end if
			if point9 = "8" then
				ipoint8 = "1"
			else
				ipoint8 = "0"
			end if
			if point9 = "7" then
				ipoint7 = "1"
			else
				ipoint7 = "0"
			end if
			pointplus =  Request("QUESTIONP_7"&i)
			if pointplus = "" then
				pointplus = "0"
			end if
			totpoint =  Request("POINT_7"&i)
			if totpoint = "" then
				totpoint = "0"
			end if

			SQL = "insert into armyinformix.dbo.monitor ( factnum, factPeoplenum, seqno, point9, point8, point7, pointplus, totpoint, monitordate, monitoruser)"
			SQL = SQL & " values ( '" & db_RECEIPTFACTNUM & "', '" & db_factPeoplenum_7 & "', " & i & ", " & ipoint9& ", " & ipoint8& ", " & ipoint7
			SQL = SQL & " , " & pointplus & ","&totpoint&",getdate(),'" &INCODE&"')"
			db.execute(SQL)
			'Response.write SQL
		next
	else
		'## 7번
		SQL = "delete from armyinformix.dbo.factpeople"
		SQL = SQL & "	where factPeoplenum = '" & db_factPeoplenum_7 & "' and factnum = '" & db_RECEIPTFACTNUM & "'"
		db.execute(SQL)

		SQL = "	delete from armyinformix.dbo.monitor where factnum = '" & db_RECEIPTFACTNUM & "' and factPeoplenum = '" & db_factPeoplenum_7 & "'"
		db.execute(SQL)
	end if


	'## 8번-------------------------------------------------------------------------------------------------------------------
	if db_FRM8 = "ON" then	'수정또는 Insert
		
		if db_factPeoplenum_8 = "" then

			SQL = "select convert(varchar,convert(int,isnull(rtrim(max(convert(int,substring(factPeoplenum,2,10)))),'0'))+1) from armyinformix.dbo.factpeople where left(factPeoplenum,1) = 'M'"
	
			set RS = db.execute(SQL)
			if isnull(rs(0)) then
				db_factPeoplenum_8 = "M1"
			else
				db_factPeoplenum_8 = "M"&rs(0)
			end if
			'신규로 임의 입력건
			SQL = "INSERT INTO armyinformix.dbo.factpeople ( factPeoplenum, factnum"
			SQL = SQL & "		, SECTION2, NAME, level"
			SQL = SQL & "		,homephone, mobilephone, etcphone"
			SQL = SQL & "		,MONITORDATE, MONITORRESULT"
			SQL = SQL & "		,RESERVEDATE, Remark,Remark1, monitorpoint )"
			SQL = SQL & " VALUES ('" & db_factPeoplenum_8 & "' , '" & db_RECEIPTFACTNUM & "'"
			SQL = SQL & "		,'" & db_SECTION2_8 & "','" & db_name_8 & "','" & db_level_8 & "'"
			SQL = SQL & "		,'" & db_homephone_8 & "','" & db_mobilephone_8 & "','" & db_etcphone_8 & "'"
			if db_MONITORDATE_8 = "" then
				SQL = SQL & "		,	getdate()"
			else
				SQL = SQL & "		,	'" & db_MONITORDATE_8 & "'"
			end if
			SQL = SQL & "		,'" & db_MONITORRESULT_8 & "'"

			if db_MONITORRESULT_8 = "4" then	'예약이라면
				if db_RESERVETIME_8 = "1" then
					SQL = SQL & "		,	CONVERT(CHAR(19),DATEADD(n,10,GETDATE()),121)"
				elseif db_RESERVETIME_8 = "2" then
					SQL = SQL & "		,	CONVERT(CHAR(19),DATEADD(n,30,GETDATE()),121)"
				elseif db_RESERVETIME_8 = "3" then
					SQL = SQL & "		,	CONVERT(CHAR(19),DATEADD(n,60,GETDATE()),121)"
				elseif db_RESERVETIME_8 = "4" then
					SQL = SQL & "		,	CONVERT(CHAR(19),DATEADD(n,120,GETDATE()),121)"
				elseif db_RESERVETIME_8 <> "" then
					SQL = SQL & "		,	'" & db_RESERVEDATE_8 & " " & db_RESERVETIME_8 & ":00:00'"
				else
					SQL = SQL & "		,	'" & db_RESERVEDATE_8 & " " & db_RESERVEHOUR_8 & ":"&  db_RESERVEMIN_8 &":00'"
				end if
			else
				SQL = SQL & "		,	''"
			end if
			SQL = SQL & "		,	'" & db_Remark_8 & "'"
			SQL = SQL & "		,	'" & db_Remark1_8 & "'"
			SQL = SQL & "		,	'" & db_TOT_8 & "')"

			db.execute(SQL)
	
		else

			SQL = " update armyinformix.dbo.factpeople set SECTION2 = '" & db_SECTION2_8 & "'"
			SQL = SQL & "		,	name = '" & db_name_8 & "',	level = '" & db_level_8 & "', homephone = '" & db_homephone_8 & "'"
			SQL = SQL & "		,	mobilephone = '" & db_mobilephone_8 & "', etcphone = '" & db_etcphone_8 & "'"

			if db_MONITORDATE_8 = "" then
				SQL = SQL & "		,	MONITORDATE = getdate()"
			else
				SQL = SQL & "		,	MONITORDATE = convert(datetime,'"&db_MONITORDATE_8&"')"
			end if
			SQL = SQL & "		,	MONITORRESULT = '" & db_MONITORRESULT_8 & "'"

			if db_MONITORRESULT_8 = "4" then	'예약이라면
				if db_RESERVETIME_8 = "1" then
					SQL = SQL & "		,	RESERVEDATE = CONVERT(CHAR(19),DATEADD(n,10,GETDATE()),121)"
				elseif db_RESERVETIME_8 = "2" then
					SQL = SQL & "		,	RESERVEDATE = CONVERT(CHAR(19),DATEADD(n,30,GETDATE()),121)"
				elseif db_RESERVETIME_8 = "3" then
					SQL = SQL & "		,	RESERVEDATE = CONVERT(CHAR(19),DATEADD(n,60,GETDATE()),121)"
				elseif db_RESERVETIME_8 = "4" then
					SQL = SQL & "		,	RESERVEDATE = CONVERT(CHAR(19),DATEADD(n,120,GETDATE()),121)"
				elseif db_RESERVETIME_8 <> "" then
					SQL = SQL & "		,	RESERVEDATE = '" & db_RESERVEDATE_8 & " " & db_RESERVETIME_8 & ":00:00'"
				else
					SQL = SQL & "		,	RESERVEDATE = '" & db_RESERVEDATE_8 & " " & db_RESERVEHOUR_8 & ":"&  db_RESERVEMIN_8 &":00'"
				end if
			else
				SQL = SQL & "		,	RESERVEDATE = ''"
			end if
			SQL = SQL & "		,	Remark = '" & db_Remark_8 & "'"
			SQL = SQL & "		,	Remark1 = '" & db_Remark1_8 & "'"
			SQL = SQL & "		,	monitorpoint = '" & db_TOT_8 & "'"
			SQL = SQL & "	where factPeoplenum = '" & db_factPeoplenum_8 & "' and factnum = '" & db_RECEIPTFACTNUM & "'"

			db.execute(SQL)
		end if

		'## 1번
		SQL = "	delete from armyinformix.dbo.monitor where factnum = '" & db_RECEIPTFACTNUM & "' and factPeoplenum = '" & db_factPeoplenum_8 & "'"
		db.execute(SQL)
		for i = 1 to icnt
			point9 =  Request("QUESTION_8"&i)
			if point9 = "9" then
				ipoint9 = "1"
			else
				ipoint9 = "0"
			end if
			if point9 = "8" then
				ipoint8 = "1"
			else
				ipoint8 = "0"
			end if
			if point9 = "7" then
				ipoint7 = "1"
			else
				ipoint7 = "0"
			end if
			pointplus =  Request("QUESTIONP_8"&i)
			if pointplus = "" then
				pointplus = "0"
			end if
			totpoint =  Request("POINT_8"&i)
			if totpoint = "" then
				totpoint = "0"
			end if

			SQL = "insert into armyinformix.dbo.monitor ( factnum, factPeoplenum, seqno, point9, point8, point7, pointplus, totpoint, monitordate, monitoruser)"
			SQL = SQL & " values ( '" & db_RECEIPTFACTNUM & "', '" & db_factPeoplenum_8 & "', " & i & ", " & ipoint9& ", " & ipoint8& ", " & ipoint7
			SQL = SQL & " , " & pointplus & ","&totpoint&",getdate(),'" &INCODE&"')"
			db.execute(SQL)
			'Response.write SQL
		next
	else
		'## 8번
		SQL = "delete from armyinformix.dbo.factpeople"
		SQL = SQL & "	where factPeoplenum = '" & db_factPeoplenum_8 & "' and factnum = '" & db_RECEIPTFACTNUM & "'"
		db.execute(SQL)

		SQL = "	delete from armyinformix.dbo.monitor where factnum = '" & db_RECEIPTFACTNUM & "' and factPeoplenum = '" & db_factPeoplenum_8 & "'"
		db.execute(SQL)
	end if



	'## 9번-------------------------------------------------------------------------------------------------------------------
	if db_FRM9 = "ON" then	'수정또는 Insert
		
		if db_factPeoplenum_9 = "" then

			SQL = "select convert(varchar,convert(int,isnull(rtrim(max(convert(int,substring(factPeoplenum,2,10)))),'0'))+1) from armyinformix.dbo.factpeople where left(factPeoplenum,1) = 'M'"
	
			set RS = db.execute(SQL)
			if isnull(rs(0)) then
				db_factPeoplenum_9 = "M1"
			else
				db_factPeoplenum_9 = "M"&rs(0)
			end if
			'신규로 임의 입력건
			SQL = "INSERT INTO armyinformix.dbo.factpeople ( factPeoplenum, factnum"
			SQL = SQL & "		, SECTION2, NAME, level"
			SQL = SQL & "		,homephone, mobilephone, etcphone"
			SQL = SQL & "		,MONITORDATE, MONITORRESULT"
			SQL = SQL & "		,RESERVEDATE, Remark,Remark1, monitorpoint )"
			SQL = SQL & " VALUES ('" & db_factPeoplenum_9 & "' , '" & db_RECEIPTFACTNUM & "'"
			SQL = SQL & "		,'" & db_SECTION2_9 & "','" & db_name_9 & "','" & db_level_9 & "'"
			SQL = SQL & "		,'" & db_homephone_9 & "','" & db_mobilephone_9 & "','" & db_etcphone_9 & "'"
			if db_MONITORDATE_9 = "" then
				SQL = SQL & "		,	getdate()"
			else
				SQL = SQL & "		,	'" & db_MONITORDATE_9 & "'"
			end if
			SQL = SQL & "		,'" & db_MONITORRESULT_9 & "'"

			if db_MONITORRESULT_9 = "4" then	'예약이라면
				if db_RESERVETIME_9 = "1" then
					SQL = SQL & "		,	CONVERT(CHAR(19),DATEADD(n,10,GETDATE()),121)"
				elseif db_RESERVETIME_9 = "2" then
					SQL = SQL & "		,	CONVERT(CHAR(19),DATEADD(n,30,GETDATE()),121)"
				elseif db_RESERVETIME_9 = "3" then
					SQL = SQL & "		,	CONVERT(CHAR(19),DATEADD(n,60,GETDATE()),121)"
				elseif db_RESERVETIME_9 = "4" then
					SQL = SQL & "		,	CONVERT(CHAR(19),DATEADD(n,120,GETDATE()),121)"
				elseif db_RESERVETIME_9 <> "" then
					SQL = SQL & "		,	'" & db_RESERVEDATE_9 & " " & db_RESERVETIME_9 & ":00:00'"
				else
					SQL = SQL & "		,	'" & db_RESERVEDATE_9 & " " & db_RESERVEHOUR_9 & ":"&  db_RESERVEMIN_9 &":00'"
				end if
			else
				SQL = SQL & "		,	''"
			end if
			SQL = SQL & "		,	'" & db_Remark_9 & "'"
			SQL = SQL & "		,	'" & db_Remark1_9 & "'"
			SQL = SQL & "		,	'" & db_TOT_9 & "')"

			db.execute(SQL)
	
		else

			SQL = " update armyinformix.dbo.factpeople set SECTION2 = '" & db_SECTION2_9 & "'"
			SQL = SQL & "		,	name = '" & db_name_9 & "',	level = '" & db_level_9 & "', homephone = '" & db_homephone_9 & "'"
			SQL = SQL & "		,	mobilephone = '" & db_mobilephone_9 & "', etcphone = '" & db_etcphone_9 & "'"

			if db_MONITORDATE_9 = "" then
				SQL = SQL & "		,	MONITORDATE = getdate()"
			else
				SQL = SQL & "		,	MONITORDATE = convert(datetime,'"&db_MONITORDATE_9&"')"
			end if
			SQL = SQL & "		,	MONITORRESULT = '" & db_MONITORRESULT_9 & "'"

			if db_MONITORRESULT_9 = "4" then	'예약이라면
				if db_RESERVETIME_9 = "1" then
					SQL = SQL & "		,	RESERVEDATE = CONVERT(CHAR(19),DATEADD(n,10,GETDATE()),121)"
				elseif db_RESERVETIME_9 = "2" then
					SQL = SQL & "		,	RESERVEDATE = CONVERT(CHAR(19),DATEADD(n,30,GETDATE()),121)"
				elseif db_RESERVETIME_9 = "3" then
					SQL = SQL & "		,	RESERVEDATE = CONVERT(CHAR(19),DATEADD(n,60,GETDATE()),121)"
				elseif db_RESERVETIME_9 = "4" then
					SQL = SQL & "		,	RESERVEDATE = CONVERT(CHAR(19),DATEADD(n,120,GETDATE()),121)"
				elseif db_RESERVETIME_9 <> "" then
					SQL = SQL & "		,	RESERVEDATE = '" & db_RESERVEDATE_9 & " " & db_RESERVETIME_9 & ":00:00'"
				else
					SQL = SQL & "		,	RESERVEDATE = '" & db_RESERVEDATE_9 & " " & db_RESERVEHOUR_9 & ":"&  db_RESERVEMIN_9 &":00'"
				end if
			else
				SQL = SQL & "		,	RESERVEDATE = ''"
			end if
			SQL = SQL & "		,	Remark = '" & db_Remark_9 & "'"
			SQL = SQL & "		,	Remark1 = '" & db_Remark1_9 & "'"
			SQL = SQL & "		,	monitorpoint = '" & db_TOT_9 & "'"
			SQL = SQL & "	where factPeoplenum = '" & db_factPeoplenum_9 & "' and factnum = '" & db_RECEIPTFACTNUM & "'"

			db.execute(SQL)
		end if

		'## 1번
		SQL = "	delete from armyinformix.dbo.monitor where factnum = '" & db_RECEIPTFACTNUM & "' and factPeoplenum = '" & db_factPeoplenum_9 & "'"
		db.execute(SQL)
		for i = 1 to icnt
			point9 =  Request("QUESTION_9"&i)
			if point9 = "9" then
				ipoint9 = "1"
			else
				ipoint9 = "0"
			end if
			if point9 = "8" then
				ipoint8 = "1"
			else
				ipoint8 = "0"
			end if
			if point9 = "7" then
				ipoint7 = "1"
			else
				ipoint7 = "0"
			end if
			pointplus =  Request("QUESTIONP_9"&i)
			if pointplus = "" then
				pointplus = "0"
			end if
			totpoint =  Request("POINT_9"&i)
			if totpoint = "" then
				totpoint = "0"
			end if

			SQL = "insert into armyinformix.dbo.monitor ( factnum, factPeoplenum, seqno, point9, point8, point7, pointplus, totpoint, monitordate, monitoruser)"
			SQL = SQL & " values ( '" & db_RECEIPTFACTNUM & "', '" & db_factPeoplenum_9 & "', " & i & ", " & ipoint9& ", " & ipoint8& ", " & ipoint7
			SQL = SQL & " , " & pointplus & ","&totpoint&",getdate(),'" &INCODE&"')"
			db.execute(SQL)
			'Response.write SQL
		next
	else
		'## 9번
		SQL = "delete from armyinformix.dbo.factpeople"
		SQL = SQL & "	where factPeoplenum = '" & db_factPeoplenum_9 & "' and factnum = '" & db_RECEIPTFACTNUM & "'"
		db.execute(SQL)

		SQL = "	delete from armyinformix.dbo.monitor where factnum = '" & db_RECEIPTFACTNUM & "' and factPeoplenum = '" & db_factPeoplenum_9 & "'"
		db.execute(SQL)
	end if


	'## 10번-------------------------------------------------------------------------------------------------------------------
	if db_FRM10 = "ON" then	'수정또는 Insert
		
		if db_factPeoplenum_10 = "" then

			SQL = "select convert(varchar,convert(int,isnull(rtrim(max(convert(int,substring(factPeoplenum,2,10)))),'0'))+1) from armyinformix.dbo.factpeople where left(factPeoplenum,1) = 'M'"
	
			set RS = db.execute(SQL)
			if isnull(rs(0)) then
				db_factPeoplenum_10 = "M1"
			else
				db_factPeoplenum_10 = "M"&rs(0)
			end if
			'신규로 임의 입력건
			SQL = "INSERT INTO armyinformix.dbo.factpeople ( factPeoplenum, factnum"
			SQL = SQL & "		, SECTION2, NAME, level"
			SQL = SQL & "		,homephone, mobilephone, etcphone"
			SQL = SQL & "		,MONITORDATE, MONITORRESULT"
			SQL = SQL & "		,RESERVEDATE, Remark,Remark1, monitorpoint )"
			SQL = SQL & " VALUES ('" & db_factPeoplenum_10 & "' , '" & db_RECEIPTFACTNUM & "'"
			SQL = SQL & "		,'" & db_SECTION2_10 & "','" & db_name_10 & "','" & db_level_10 & "'"
			SQL = SQL & "		,'" & db_homephone_10 & "','" & db_mobilephone_10 & "','" & db_etcphone_10 & "'"
			if db_MONITORDATE_10 = "" then
				SQL = SQL & "		,	getdate()"
			else
				SQL = SQL & "		,	'" & db_MONITORDATE_10 & "'"
			end if
			SQL = SQL & "		,'" & db_MONITORRESULT_10 & "'"

			if db_MONITORRESULT_10 = "4" then	'예약이라면
				if db_RESERVETIME_10 = "1" then
					SQL = SQL & "		,	CONVERT(CHAR(19),DATEADD(n,10,GETDATE()),121)"
				elseif db_RESERVETIME_10 = "2" then
					SQL = SQL & "		,	CONVERT(CHAR(19),DATEADD(n,30,GETDATE()),121)"
				elseif db_RESERVETIME_10 = "3" then
					SQL = SQL & "		,	CONVERT(CHAR(19),DATEADD(n,60,GETDATE()),121)"
				elseif db_RESERVETIME_10 = "4" then
					SQL = SQL & "		,	CONVERT(CHAR(19),DATEADD(n,120,GETDATE()),121)"
				elseif db_RESERVETIME_10 <> "" then
					SQL = SQL & "		,	'" & db_RESERVEDATE_10 & " " & db_RESERVETIME_10 & ":00:00'"
				else
					SQL = SQL & "		,	'" & db_RESERVEDATE_10 & " " & db_RESERVEHOUR_10 & ":"&  db_RESERVEMIN_10 &":00'"
				end if
			else
				SQL = SQL & "		,	''"
			end if
			SQL = SQL & "		,	'" & db_Remark_10 & "'"
			SQL = SQL & "		,	'" & db_Remark1_10 & "'"
			SQL = SQL & "		,	'" & db_TOT_10 & "')"

			db.execute(SQL)
	
		else

			SQL = " update armyinformix.dbo.factpeople set SECTION2 = '" & db_SECTION2_10 & "'"
			SQL = SQL & "		,	name = '" & db_name_10 & "',	level = '" & db_level_10 & "', homephone = '" & db_homephone_10 & "'"
			SQL = SQL & "		,	mobilephone = '" & db_mobilephone_10 & "', etcphone = '" & db_etcphone_10 & "'"

			if db_MONITORDATE_10 = "" then
				SQL = SQL & "		,	MONITORDATE = getdate()"
			else
				SQL = SQL & "		,	MONITORDATE = convert(datetime,'"&db_MONITORDATE_10&"')"
			end if
			SQL = SQL & "		,	MONITORRESULT = '" & db_MONITORRESULT_10 & "'"

			if db_MONITORRESULT_10 = "4" then	'예약이라면
				if db_RESERVETIME_10 = "1" then
					SQL = SQL & "		,	RESERVEDATE = CONVERT(CHAR(19),DATEADD(n,10,GETDATE()),121)"
				elseif db_RESERVETIME_10 = "2" then
					SQL = SQL & "		,	RESERVEDATE = CONVERT(CHAR(19),DATEADD(n,30,GETDATE()),121)"
				elseif db_RESERVETIME_10 = "3" then
					SQL = SQL & "		,	RESERVEDATE = CONVERT(CHAR(19),DATEADD(n,60,GETDATE()),121)"
				elseif db_RESERVETIME_10 = "4" then
					SQL = SQL & "		,	RESERVEDATE = CONVERT(CHAR(19),DATEADD(n,120,GETDATE()),121)"
				elseif db_RESERVETIME_10 <> "" then
					SQL = SQL & "		,	RESERVEDATE = '" & db_RESERVEDATE_10 & " " & db_RESERVETIME_10 & ":00:00'"
				else
					SQL = SQL & "		,	RESERVEDATE = '" & db_RESERVEDATE_10 & " " & db_RESERVEHOUR_10 & ":"&  db_RESERVEMIN_10 &":00'"
				end if
			else
				SQL = SQL & "		,	RESERVEDATE = ''"
			end if
			SQL = SQL & "		,	Remark = '" & db_Remark_10 & "'"
			SQL = SQL & "		,	Remark1 = '" & db_Remark1_10 & "'"
			SQL = SQL & "		,	monitorpoint = '" & db_TOT_10 & "'"
			SQL = SQL & "	where factPeoplenum = '" & db_factPeoplenum_10 & "' and factnum = '" & db_RECEIPTFACTNUM & "'"

			db.execute(SQL)
		end if

		'## 1번
		SQL = "	delete from armyinformix.dbo.monitor where factnum = '" & db_RECEIPTFACTNUM & "' and factPeoplenum = '" & db_factPeoplenum_10 & "'"
		db.execute(SQL)
		for i = 1 to icnt
			point9 =  Request("QUESTION_10"&i)
			if point9 = "9" then
				ipoint9 = "1"
			else
				ipoint9 = "0"
			end if
			if point9 = "8" then
				ipoint8 = "1"
			else
				ipoint8 = "0"
			end if
			if point9 = "7" then
				ipoint7 = "1"
			else
				ipoint7 = "0"
			end if
			pointplus =  Request("QUESTIONP_10"&i)
			if pointplus = "" then
				pointplus = "0"
			end if
			totpoint =  Request("POINT_10"&i)
			if totpoint = "" then
				totpoint = "0"
			end if

			SQL = "insert into armyinformix.dbo.monitor ( factnum, factPeoplenum, seqno, point9, point8, point7, pointplus, totpoint, monitordate, monitoruser)"
			SQL = SQL & " values ( '" & db_RECEIPTFACTNUM & "', '" & db_factPeoplenum_10 & "', " & i & ", " & ipoint9& ", " & ipoint8& ", " & ipoint7
			SQL = SQL & " , " & pointplus & ","&totpoint&",getdate(),'" &INCODE&"')"
			db.execute(SQL)
			'Response.write SQL
		next
	else
		'## 10번
		SQL = "delete from armyinformix.dbo.factpeople"
		SQL = SQL & "	where factPeoplenum = '" & db_factPeoplenum_10 & "' and factnum = '" & db_RECEIPTFACTNUM & "'"
		db.execute(SQL)

		SQL = "	delete from armyinformix.dbo.monitor where factnum = '" & db_RECEIPTFACTNUM & "' and factPeoplenum = '" & db_factPeoplenum_10 & "'"
		db.execute(SQL)
	end if


	'## 11번-------------------------------------------------------------------------------------------------------------------
	if db_FRM11 = "ON" then	'수정또는 Insert
		
		if db_factPeoplenum_11 = "" then

			SQL = "select convert(varchar,convert(int,isnull(rtrim(max(convert(int,substring(factPeoplenum,2,10)))),'0'))+1) from armyinformix.dbo.factpeople where left(factPeoplenum,1) = 'M'"
	
			set RS = db.execute(SQL)
			if isnull(rs(0)) then
				db_factPeoplenum_11 = "M1"
			else
				db_factPeoplenum_11 = "M"&rs(0)
			end if
			'신규로 임의 입력건
			SQL = "INSERT INTO armyinformix.dbo.factpeople ( factPeoplenum, factnum"
			SQL = SQL & "		, SECTION2, NAME, level"
			SQL = SQL & "		,homephone, mobilephone, etcphone"
			SQL = SQL & "		,MONITORDATE, MONITORRESULT"
			SQL = SQL & "		,RESERVEDATE, Remark,Remark1, monitorpoint )"
			SQL = SQL & " VALUES ('" & db_factPeoplenum_11 & "' , '" & db_RECEIPTFACTNUM & "'"
			SQL = SQL & "		,'" & db_SECTION2_11 & "','" & db_name_11 & "','" & db_level_11 & "'"
			SQL = SQL & "		,'" & db_homephone_11 & "','" & db_mobilephone_11 & "','" & db_etcphone_11 & "'"
			if db_MONITORDATE_11 = "" then
				SQL = SQL & "		,	getdate()"
			else
				SQL = SQL & "		,	'" & db_MONITORDATE_11 & "'"
			end if
			SQL = SQL & "		,'" & db_MONITORRESULT_11 & "'"

			if db_MONITORRESULT_11 = "4" then	'예약이라면
				if db_RESERVETIME_11 = "1" then
					SQL = SQL & "		,	CONVERT(CHAR(19),DATEADD(n,10,GETDATE()),121)"
				elseif db_RESERVETIME_11 = "2" then
					SQL = SQL & "		,	CONVERT(CHAR(19),DATEADD(n,30,GETDATE()),121)"
				elseif db_RESERVETIME_11 = "3" then
					SQL = SQL & "		,	CONVERT(CHAR(19),DATEADD(n,60,GETDATE()),121)"
				elseif db_RESERVETIME_11 = "4" then
					SQL = SQL & "		,	CONVERT(CHAR(19),DATEADD(n,120,GETDATE()),121)"
				elseif db_RESERVETIME_11 <> "" then
					SQL = SQL & "		,	'" & db_RESERVEDATE_11 & " " & db_RESERVETIME_11 & ":00:00'"
				else
					SQL = SQL & "		,	'" & db_RESERVEDATE_11 & " " & db_RESERVEHOUR_11 & ":"&  db_RESERVEMIN_11 &":00'"
				end if
			else
				SQL = SQL & "		,	''"
			end if
			SQL = SQL & "		,	'" & db_Remark_11 & "'"
			SQL = SQL & "		,	'" & db_Remark1_11 & "'"
			SQL = SQL & "		,	'" & db_TOT_11 & "')"

			db.execute(SQL)
	
		else

			SQL = " update armyinformix.dbo.factpeople set SECTION2 = '" & db_SECTION2_11 & "'"
			SQL = SQL & "		,	name = '" & db_name_11 & "',	level = '" & db_level_11 & "', homephone = '" & db_homephone_11 & "'"
			SQL = SQL & "		,	mobilephone = '" & db_mobilephone_11 & "', etcphone = '" & db_etcphone_11 & "'"

			if db_MONITORDATE_11 = "" then
				SQL = SQL & "		,	MONITORDATE = getdate()"
			else
				SQL = SQL & "		,	MONITORDATE = convert(datetime,'"&db_MONITORDATE_11&"')"
			end if
			SQL = SQL & "		,	MONITORRESULT = '" & db_MONITORRESULT_11 & "'"

			if db_MONITORRESULT_11 = "4" then	'예약이라면
				if db_RESERVETIME_11 = "1" then
					SQL = SQL & "		,	RESERVEDATE = CONVERT(CHAR(19),DATEADD(n,10,GETDATE()),121)"
				elseif db_RESERVETIME_11 = "2" then
					SQL = SQL & "		,	RESERVEDATE = CONVERT(CHAR(19),DATEADD(n,30,GETDATE()),121)"
				elseif db_RESERVETIME_11 = "3" then
					SQL = SQL & "		,	RESERVEDATE = CONVERT(CHAR(19),DATEADD(n,60,GETDATE()),121)"
				elseif db_RESERVETIME_11 = "4" then
					SQL = SQL & "		,	RESERVEDATE = CONVERT(CHAR(19),DATEADD(n,120,GETDATE()),121)"
				elseif db_RESERVETIME_11 <> "" then
					SQL = SQL & "		,	RESERVEDATE = '" & db_RESERVEDATE_11 & " " & db_RESERVETIME_11 & ":00:00'"
				else
					SQL = SQL & "		,	RESERVEDATE = '" & db_RESERVEDATE_11 & " " & db_RESERVEHOUR_11 & ":"&  db_RESERVEMIN_11 &":00'"
				end if
			else
				SQL = SQL & "		,	RESERVEDATE = ''"
			end if
			SQL = SQL & "		,	Remark = '" & db_Remark_11 & "'"
			SQL = SQL & "		,	Remark1 = '" & db_Remark1_11 & "'"
			SQL = SQL & "		,	monitorpoint = '" & db_TOT_11 & "'"
			SQL = SQL & "	where factPeoplenum = '" & db_factPeoplenum_11 & "' and factnum = '" & db_RECEIPTFACTNUM & "'"

			db.execute(SQL)
		end if

		'## 1번
		SQL = "	delete from armyinformix.dbo.monitor where factnum = '" & db_RECEIPTFACTNUM & "' and factPeoplenum = '" & db_factPeoplenum_11 & "'"
		db.execute(SQL)
		for i = 1 to icnt
			point9 =  Request("QUESTION_11"&i)
			if point9 = "9" then
				ipoint9 = "1"
			else
				ipoint9 = "0"
			end if
			if point9 = "8" then
				ipoint8 = "1"
			else
				ipoint8 = "0"
			end if
			if point9 = "7" then
				ipoint7 = "1"
			else
				ipoint7 = "0"
			end if
			pointplus =  Request("QUESTIONP_11"&i)
			if pointplus = "" then
				pointplus = "0"
			end if
			totpoint =  Request("POINT_11"&i)
			if totpoint = "" then
				totpoint = "0"
			end if

			SQL = "insert into armyinformix.dbo.monitor ( factnum, factPeoplenum, seqno, point9, point8, point7, pointplus, totpoint, monitordate, monitoruser)"
			SQL = SQL & " values ( '" & db_RECEIPTFACTNUM & "', '" & db_factPeoplenum_11 & "', " & i & ", " & ipoint9& ", " & ipoint8& ", " & ipoint7
			SQL = SQL & " , " & pointplus & ","&totpoint&",getdate(),'" &INCODE&"')"
			db.execute(SQL)
			'Response.write SQL
		next
	else
		'## 11번
		SQL = "delete from armyinformix.dbo.factpeople"
		SQL = SQL & "	where factPeoplenum = '" & db_factPeoplenum_11 & "' and factnum = '" & db_RECEIPTFACTNUM & "'"
		db.execute(SQL)

		SQL = "	delete from armyinformix.dbo.monitor where factnum = '" & db_RECEIPTFACTNUM & "' and factPeoplenum = '" & db_factPeoplenum_11 & "'"
		db.execute(SQL)
	end if


	'## 12번-------------------------------------------------------------------------------------------------------------------
	if db_FRM12 = "ON" then	'수정또는 Insert
		
		if db_factPeoplenum_12 = "" then

			SQL = "select convert(varchar,convert(int,isnull(rtrim(max(convert(int,substring(factPeoplenum,2,10)))),'0'))+1) from armyinformix.dbo.factpeople where left(factPeoplenum,1) = 'M'"
	
			set RS = db.execute(SQL)
			if isnull(rs(0)) then
				db_factPeoplenum_12 = "M1"
			else
				db_factPeoplenum_12 = "M"&rs(0)
			end if
			'신규로 임의 입력건
			SQL = "INSERT INTO armyinformix.dbo.factpeople ( factPeoplenum, factnum"
			SQL = SQL & "		, SECTION2, NAME, level"
			SQL = SQL & "		,homephone, mobilephone, etcphone"
			SQL = SQL & "		,MONITORDATE, MONITORRESULT"
			SQL = SQL & "		,RESERVEDATE, Remark,Remark1, monitorpoint )"
			SQL = SQL & " VALUES ('" & db_factPeoplenum_12 & "' , '" & db_RECEIPTFACTNUM & "'"
			SQL = SQL & "		,'" & db_SECTION2_12 & "','" & db_name_12 & "','" & db_level_12 & "'"
			SQL = SQL & "		,'" & db_homephone_12 & "','" & db_mobilephone_12 & "','" & db_etcphone_12 & "'"
			if db_MONITORDATE_12 = "" then
				SQL = SQL & "		,	getdate()"
			else
				SQL = SQL & "		,	'" & db_MONITORDATE_12 & "'"
			end if
			SQL = SQL & "		,'" & db_MONITORRESULT_12 & "'"

			if db_MONITORRESULT_12 = "4" then	'예약이라면
				if db_RESERVETIME_12 = "1" then
					SQL = SQL & "		,	CONVERT(CHAR(19),DATEADD(n,10,GETDATE()),121)"
				elseif db_RESERVETIME_12 = "2" then
					SQL = SQL & "		,	CONVERT(CHAR(19),DATEADD(n,30,GETDATE()),121)"
				elseif db_RESERVETIME_12 = "3" then
					SQL = SQL & "		,	CONVERT(CHAR(19),DATEADD(n,60,GETDATE()),121)"
				elseif db_RESERVETIME_12 = "4" then
					SQL = SQL & "		,	CONVERT(CHAR(19),DATEADD(n,120,GETDATE()),121)"
				elseif db_RESERVETIME_12 <> "" then
					SQL = SQL & "		,	'" & db_RESERVEDATE_12 & " " & db_RESERVETIME_12 & ":00:00'"
				else
					SQL = SQL & "		,	'" & db_RESERVEDATE_12 & " " & db_RESERVEHOUR_12 & ":"&  db_RESERVEMIN_12 &":00'"
				end if
			else
				SQL = SQL & "		,	''"
			end if
			SQL = SQL & "		,	'" & db_Remark_12 & "'"
			SQL = SQL & "		,	'" & db_Remark1_12 & "'"
			SQL = SQL & "		,	'" & db_TOT_12 & "')"

			db.execute(SQL)
	
		else

			SQL = " update armyinformix.dbo.factpeople set SECTION2 = '" & db_SECTION2_12 & "'"
			SQL = SQL & "		,	name = '" & db_name_12 & "',	level = '" & db_level_12 & "', homephone = '" & db_homephone_12 & "'"
			SQL = SQL & "		,	mobilephone = '" & db_mobilephone_12 & "', etcphone = '" & db_etcphone_12 & "'"

			if db_MONITORDATE_12 = "" then
				SQL = SQL & "		,	MONITORDATE = getdate()"
			else
				SQL = SQL & "		,	MONITORDATE = convert(datetime,'"&db_MONITORDATE_12&"')"
			end if
			SQL = SQL & "		,	MONITORRESULT = '" & db_MONITORRESULT_12 & "'"

			if db_MONITORRESULT_12 = "4" then	'예약이라면
				if db_RESERVETIME_12 = "1" then
					SQL = SQL & "		,	RESERVEDATE = CONVERT(CHAR(19),DATEADD(n,10,GETDATE()),121)"
				elseif db_RESERVETIME_12 = "2" then
					SQL = SQL & "		,	RESERVEDATE = CONVERT(CHAR(19),DATEADD(n,30,GETDATE()),121)"
				elseif db_RESERVETIME_12 = "3" then
					SQL = SQL & "		,	RESERVEDATE = CONVERT(CHAR(19),DATEADD(n,60,GETDATE()),121)"
				elseif db_RESERVETIME_12 = "4" then
					SQL = SQL & "		,	RESERVEDATE = CONVERT(CHAR(19),DATEADD(n,120,GETDATE()),121)"
				elseif db_RESERVETIME_12 <> "" then
					SQL = SQL & "		,	RESERVEDATE = '" & db_RESERVEDATE_12 & " " & db_RESERVETIME_12 & ":00:00'"
				else
					SQL = SQL & "		,	RESERVEDATE = '" & db_RESERVEDATE_12 & " " & db_RESERVEHOUR_12 & ":"&  db_RESERVEMIN_12 &":00'"
				end if
			else
				SQL = SQL & "		,	RESERVEDATE = ''"
			end if
			SQL = SQL & "		,	Remark = '" & db_Remark_12 & "'"
			SQL = SQL & "		,	Remark1 = '" & db_Remark1_12 & "'"
			SQL = SQL & "		,	monitorpoint = '" & db_TOT_12 & "'"
			SQL = SQL & "	where factPeoplenum = '" & db_factPeoplenum_12 & "' and factnum = '" & db_RECEIPTFACTNUM & "'"

			db.execute(SQL)
		end if

		'## 1번
		SQL = "	delete from armyinformix.dbo.monitor where factnum = '" & db_RECEIPTFACTNUM & "' and factPeoplenum = '" & db_factPeoplenum_12 & "'"
		db.execute(SQL)
		for i = 1 to icnt
			point9 =  Request("QUESTION_12"&i)
			if point9 = "9" then
				ipoint9 = "1"
			else
				ipoint9 = "0"
			end if
			if point9 = "8" then
				ipoint8 = "1"
			else
				ipoint8 = "0"
			end if
			if point9 = "7" then
				ipoint7 = "1"
			else
				ipoint7 = "0"
			end if
			pointplus =  Request("QUESTIONP_12"&i)
			if pointplus = "" then
				pointplus = "0"
			end if
			totpoint =  Request("POINT_12"&i)
			if totpoint = "" then
				totpoint = "0"
			end if

			SQL = "insert into armyinformix.dbo.monitor ( factnum, factPeoplenum, seqno, point9, point8, point7, pointplus, totpoint, monitordate, monitoruser)"
			SQL = SQL & " values ( '" & db_RECEIPTFACTNUM & "', '" & db_factPeoplenum_12 & "', " & i & ", " & ipoint9& ", " & ipoint8& ", " & ipoint7
			SQL = SQL & " , " & pointplus & ","&totpoint&",getdate(),'" &INCODE&"')"
			db.execute(SQL)
			'Response.write SQL
		next
	else
		'## 12번
		SQL = "delete from armyinformix.dbo.factpeople"
		SQL = SQL & "	where factPeoplenum = '" & db_factPeoplenum_12 & "' and factnum = '" & db_RECEIPTFACTNUM & "'"
		db.execute(SQL)

		SQL = "	delete from armyinformix.dbo.monitor where factnum = '" & db_RECEIPTFACTNUM & "' and factPeoplenum = '" & db_factPeoplenum_12 & "'"
		db.execute(SQL)
	end if


	'## 13번-------------------------------------------------------------------------------------------------------------------
	if db_FRM13 = "ON" then	'수정또는 Insert
		
		if db_factPeoplenum_13 = "" then

			SQL = "select convert(varchar,convert(int,isnull(rtrim(max(convert(int,substring(factPeoplenum,2,10)))),'0'))+1) from armyinformix.dbo.factpeople where left(factPeoplenum,1) = 'M'"
	
			set RS = db.execute(SQL)
			if isnull(rs(0)) then
				db_factPeoplenum_13 = "M1"
			else
				db_factPeoplenum_13 = "M"&rs(0)
			end if
			'신규로 임의 입력건
			SQL = "INSERT INTO armyinformix.dbo.factpeople ( factPeoplenum, factnum"
			SQL = SQL & "		, SECTION2, NAME, level"
			SQL = SQL & "		,homephone, mobilephone, etcphone"
			SQL = SQL & "		,MONITORDATE, MONITORRESULT"
			SQL = SQL & "		,RESERVEDATE, Remark,Remark1, monitorpoint )"
			SQL = SQL & " VALUES ('" & db_factPeoplenum_13 & "' , '" & db_RECEIPTFACTNUM & "'"
			SQL = SQL & "		,'" & db_SECTION2_13 & "','" & db_name_13 & "','" & db_level_13 & "'"
			SQL = SQL & "		,'" & db_homephone_13 & "','" & db_mobilephone_13 & "','" & db_etcphone_13 & "'"
			if db_MONITORDATE_13 = "" then
				SQL = SQL & "		,	getdate()"
			else
				SQL = SQL & "		,	'" & db_MONITORDATE_13 & "'"
			end if
			SQL = SQL & "		,'" & db_MONITORRESULT_13 & "'"

			if db_MONITORRESULT_13 = "4" then	'예약이라면
				if db_RESERVETIME_13 = "1" then
					SQL = SQL & "		,	CONVERT(CHAR(19),DATEADD(n,10,GETDATE()),121)"
				elseif db_RESERVETIME_13 = "2" then
					SQL = SQL & "		,	CONVERT(CHAR(19),DATEADD(n,30,GETDATE()),121)"
				elseif db_RESERVETIME_13 = "3" then
					SQL = SQL & "		,	CONVERT(CHAR(19),DATEADD(n,60,GETDATE()),121)"
				elseif db_RESERVETIME_13 = "4" then
					SQL = SQL & "		,	CONVERT(CHAR(19),DATEADD(n,120,GETDATE()),121)"
				elseif db_RESERVETIME_13 <> "" then
					SQL = SQL & "		,	'" & db_RESERVEDATE_13 & " " & db_RESERVETIME_13 & ":00:00'"
				else
					SQL = SQL & "		,	'" & db_RESERVEDATE_13 & " " & db_RESERVEHOUR_13 & ":"&  db_RESERVEMIN_13 &":00'"
				end if
			else
				SQL = SQL & "		,	''"
			end if
			SQL = SQL & "		,	'" & db_Remark_13 & "'"
			SQL = SQL & "		,	'" & db_Remark1_13 & "'"
			SQL = SQL & "		,	'" & db_TOT_13 & "')"

			db.execute(SQL)
	
		else

			SQL = " update armyinformix.dbo.factpeople set SECTION2 = '" & db_SECTION2_13 & "'"
			SQL = SQL & "		,	name = '" & db_name_13 & "',	level = '" & db_level_13 & "', homephone = '" & db_homephone_13 & "'"
			SQL = SQL & "		,	mobilephone = '" & db_mobilephone_13 & "', etcphone = '" & db_etcphone_13 & "'"

			if db_MONITORDATE_13 = "" then
				SQL = SQL & "		,	MONITORDATE = getdate()"
			else
				SQL = SQL & "		,	MONITORDATE = convert(datetime,'"&db_MONITORDATE_13&"')"
			end if
			SQL = SQL & "		,	MONITORRESULT = '" & db_MONITORRESULT_13 & "'"

			if db_MONITORRESULT_13 = "4" then	'예약이라면
				if db_RESERVETIME_13 = "1" then
					SQL = SQL & "		,	RESERVEDATE = CONVERT(CHAR(19),DATEADD(n,10,GETDATE()),121)"
				elseif db_RESERVETIME_13 = "2" then
					SQL = SQL & "		,	RESERVEDATE = CONVERT(CHAR(19),DATEADD(n,30,GETDATE()),121)"
				elseif db_RESERVETIME_13 = "3" then
					SQL = SQL & "		,	RESERVEDATE = CONVERT(CHAR(19),DATEADD(n,60,GETDATE()),121)"
				elseif db_RESERVETIME_13 = "4" then
					SQL = SQL & "		,	RESERVEDATE = CONVERT(CHAR(19),DATEADD(n,120,GETDATE()),121)"
				elseif db_RESERVETIME_13 <> "" then
					SQL = SQL & "		,	RESERVEDATE = '" & db_RESERVEDATE_13 & " " & db_RESERVETIME_13 & ":00:00'"
				else
					SQL = SQL & "		,	RESERVEDATE = '" & db_RESERVEDATE_13 & " " & db_RESERVEHOUR_13 & ":"&  db_RESERVEMIN_13 &":00'"
				end if
			else
				SQL = SQL & "		,	RESERVEDATE = ''"
			end if
			SQL = SQL & "		,	Remark = '" & db_Remark_13 & "'"
			SQL = SQL & "		,	Remark1 = '" & db_Remark1_13 & "'"
			SQL = SQL & "		,	monitorpoint = '" & db_TOT_13 & "'"
			SQL = SQL & "	where factPeoplenum = '" & db_factPeoplenum_13 & "' and factnum = '" & db_RECEIPTFACTNUM & "'"

			db.execute(SQL)
		end if

		'## 1번
		SQL = "	delete from armyinformix.dbo.monitor where factnum = '" & db_RECEIPTFACTNUM & "' and factPeoplenum = '" & db_factPeoplenum_13 & "'"
		db.execute(SQL)
		for i = 1 to icnt
			point9 =  Request("QUESTION_13"&i)
			if point9 = "9" then
				ipoint9 = "1"
			else
				ipoint9 = "0"
			end if
			if point9 = "8" then
				ipoint8 = "1"
			else
				ipoint8 = "0"
			end if
			if point9 = "7" then
				ipoint7 = "1"
			else
				ipoint7 = "0"
			end if
			pointplus =  Request("QUESTIONP_13"&i)
			if pointplus = "" then
				pointplus = "0"
			end if
			totpoint =  Request("POINT_13"&i)
			if totpoint = "" then
				totpoint = "0"
			end if

			SQL = "insert into armyinformix.dbo.monitor ( factnum, factPeoplenum, seqno, point9, point8, point7, pointplus, totpoint, monitordate, monitoruser)"
			SQL = SQL & " values ( '" & db_RECEIPTFACTNUM & "', '" & db_factPeoplenum_13 & "', " & i & ", " & ipoint9& ", " & ipoint8& ", " & ipoint7
			SQL = SQL & " , " & pointplus & ","&totpoint&",getdate(),'" &INCODE&"')"
			db.execute(SQL)
			'Response.write SQL
		next
	else
		'## 13번
		SQL = "delete from armyinformix.dbo.factpeople"
		SQL = SQL & "	where factPeoplenum = '" & db_factPeoplenum_13 & "' and factnum = '" & db_RECEIPTFACTNUM & "'"
		db.execute(SQL)

		SQL = "	delete from armyinformix.dbo.monitor where factnum = '" & db_RECEIPTFACTNUM & "' and factPeoplenum = '" & db_factPeoplenum_13 & "'"
		db.execute(SQL)
	end if


	'## 14번-------------------------------------------------------------------------------------------------------------------
	if db_FRM14 = "ON" then	'수정또는 Insert
		
		if db_factPeoplenum_14 = "" then

			SQL = "select convert(varchar,convert(int,isnull(rtrim(max(convert(int,substring(factPeoplenum,2,10)))),'0'))+1) from armyinformix.dbo.factpeople where left(factPeoplenum,1) = 'M'"
	
			set RS = db.execute(SQL)
			if isnull(rs(0)) then
				db_factPeoplenum_14 = "M1"
			else
				db_factPeoplenum_14 = "M"&rs(0)
			end if
			'신규로 임의 입력건
			SQL = "INSERT INTO armyinformix.dbo.factpeople ( factPeoplenum, factnum"
			SQL = SQL & "		, SECTION2, NAME, level"
			SQL = SQL & "		,homephone, mobilephone, etcphone"
			SQL = SQL & "		,MONITORDATE, MONITORRESULT"
			SQL = SQL & "		,RESERVEDATE, Remark,Remark1, monitorpoint )"
			SQL = SQL & " VALUES ('" & db_factPeoplenum_14 & "' , '" & db_RECEIPTFACTNUM & "'"
			SQL = SQL & "		,'" & db_SECTION2_14 & "','" & db_name_14 & "','" & db_level_14 & "'"
			SQL = SQL & "		,'" & db_homephone_14 & "','" & db_mobilephone_14 & "','" & db_etcphone_14 & "'"
			if db_MONITORDATE_14 = "" then
				SQL = SQL & "		,	getdate()"
			else
				SQL = SQL & "		,	'" & db_MONITORDATE_14 & "'"
			end if
			SQL = SQL & "		,'" & db_MONITORRESULT_14 & "'"

			if db_MONITORRESULT_14 = "4" then	'예약이라면
				if db_RESERVETIME_14 = "1" then
					SQL = SQL & "		,	CONVERT(CHAR(19),DATEADD(n,10,GETDATE()),121)"
				elseif db_RESERVETIME_14 = "2" then
					SQL = SQL & "		,	CONVERT(CHAR(19),DATEADD(n,30,GETDATE()),121)"
				elseif db_RESERVETIME_14 = "3" then
					SQL = SQL & "		,	CONVERT(CHAR(19),DATEADD(n,60,GETDATE()),121)"
				elseif db_RESERVETIME_14 = "4" then
					SQL = SQL & "		,	CONVERT(CHAR(19),DATEADD(n,120,GETDATE()),121)"
				elseif db_RESERVETIME_14 <> "" then
					SQL = SQL & "		,	'" & db_RESERVEDATE_14 & " " & db_RESERVETIME_14 & ":00:00'"
				else
					SQL = SQL & "		,	'" & db_RESERVEDATE_14 & " " & db_RESERVEHOUR_14 & ":"&  db_RESERVEMIN_14 &":00'"
				end if
			else
				SQL = SQL & "		,	''"
			end if
			SQL = SQL & "		,	'" & db_Remark_14 & "'"
			SQL = SQL & "		,	'" & db_Remark1_14 & "'"
			SQL = SQL & "		,	'" & db_TOT_14 & "')"

			db.execute(SQL)
	
		else

			SQL = " update armyinformix.dbo.factpeople set SECTION2 = '" & db_SECTION2_14 & "'"
			SQL = SQL & "		,	name = '" & db_name_14 & "',	level = '" & db_level_14 & "', homephone = '" & db_homephone_14 & "'"
			SQL = SQL & "		,	mobilephone = '" & db_mobilephone_14 & "', etcphone = '" & db_etcphone_14 & "'"

			if db_MONITORDATE_14 = "" then
				SQL = SQL & "		,	MONITORDATE = getdate()"
			else
				SQL = SQL & "		,	MONITORDATE = convert(datetime,'"&db_MONITORDATE_14&"')"
			end if
			SQL = SQL & "		,	MONITORRESULT = '" & db_MONITORRESULT_14 & "'"

			if db_MONITORRESULT_14 = "4" then	'예약이라면
				if db_RESERVETIME_14 = "1" then
					SQL = SQL & "		,	RESERVEDATE = CONVERT(CHAR(19),DATEADD(n,10,GETDATE()),121)"
				elseif db_RESERVETIME_14 = "2" then
					SQL = SQL & "		,	RESERVEDATE = CONVERT(CHAR(19),DATEADD(n,30,GETDATE()),121)"
				elseif db_RESERVETIME_14 = "3" then
					SQL = SQL & "		,	RESERVEDATE = CONVERT(CHAR(19),DATEADD(n,60,GETDATE()),121)"
				elseif db_RESERVETIME_14 = "4" then
					SQL = SQL & "		,	RESERVEDATE = CONVERT(CHAR(19),DATEADD(n,120,GETDATE()),121)"
				elseif db_RESERVETIME_14 <> "" then
					SQL = SQL & "		,	RESERVEDATE = '" & db_RESERVEDATE_14 & " " & db_RESERVETIME_14 & ":00:00'"
				else
					SQL = SQL & "		,	RESERVEDATE = '" & db_RESERVEDATE_14 & " " & db_RESERVEHOUR_14 & ":"&  db_RESERVEMIN_14 &":00'"
				end if
			else
				SQL = SQL & "		,	RESERVEDATE = ''"
			end if
			SQL = SQL & "		,	Remark = '" & db_Remark_14 & "'"
			SQL = SQL & "		,	Remark1 = '" & db_Remark1_14 & "'"
			SQL = SQL & "		,	monitorpoint = '" & db_TOT_14 & "'"
			SQL = SQL & "	where factPeoplenum = '" & db_factPeoplenum_14 & "' and factnum = '" & db_RECEIPTFACTNUM & "'"

			db.execute(SQL)
		end if

		'## 1번
		SQL = "	delete from armyinformix.dbo.monitor where factnum = '" & db_RECEIPTFACTNUM & "' and factPeoplenum = '" & db_factPeoplenum_14 & "'"
		db.execute(SQL)
		for i = 1 to icnt
			point9 =  Request("QUESTION_14"&i)
			if point9 = "9" then
				ipoint9 = "1"
			else
				ipoint9 = "0"
			end if
			if point9 = "8" then
				ipoint8 = "1"
			else
				ipoint8 = "0"
			end if
			if point9 = "7" then
				ipoint7 = "1"
			else
				ipoint7 = "0"
			end if
			pointplus =  Request("QUESTIONP_14"&i)
			if pointplus = "" then
				pointplus = "0"
			end if
			totpoint =  Request("POINT_14"&i)
			if totpoint = "" then
				totpoint = "0"
			end if

			SQL = "insert into armyinformix.dbo.monitor ( factnum, factPeoplenum, seqno, point9, point8, point7, pointplus, totpoint, monitordate, monitoruser)"
			SQL = SQL & " values ( '" & db_RECEIPTFACTNUM & "', '" & db_factPeoplenum_14 & "', " & i & ", " & ipoint9& ", " & ipoint8& ", " & ipoint7
			SQL = SQL & " , " & pointplus & ","&totpoint&",getdate(),'" &INCODE&"')"
			db.execute(SQL)
			'Response.write SQL
		next
	else
		'## 14번
		SQL = "delete from armyinformix.dbo.factpeople"
		SQL = SQL & "	where factPeoplenum = '" & db_factPeoplenum_14 & "' and factnum = '" & db_RECEIPTFACTNUM & "'"
		db.execute(SQL)

		SQL = "	delete from armyinformix.dbo.monitor where factnum = '" & db_RECEIPTFACTNUM & "' and factPeoplenum = '" & db_factPeoplenum_14 & "'"
		db.execute(SQL)
	end if


	'## 15번-------------------------------------------------------------------------------------------------------------------
	if db_FRM15 = "ON" then	'수정또는 Insert
		
		if db_factPeoplenum_15 = "" then

			SQL = "select convert(varchar,convert(int,isnull(rtrim(max(convert(int,substring(factPeoplenum,2,10)))),'0'))+1) from armyinformix.dbo.factpeople where left(factPeoplenum,1) = 'M'"
	
			set RS = db.execute(SQL)
			if isnull(rs(0)) then
				db_factPeoplenum_15 = "M1"
			else
				db_factPeoplenum_15 = "M"&rs(0)
			end if
			'신규로 임의 입력건
			SQL = "INSERT INTO armyinformix.dbo.factpeople ( factPeoplenum, factnum"
			SQL = SQL & "		, SECTION2, NAME, level"
			SQL = SQL & "		,homephone, mobilephone, etcphone"
			SQL = SQL & "		,MONITORDATE, MONITORRESULT"
			SQL = SQL & "		,RESERVEDATE, Remark,Remark1, monitorpoint )"
			SQL = SQL & " VALUES ('" & db_factPeoplenum_15 & "' , '" & db_RECEIPTFACTNUM & "'"
			SQL = SQL & "		,'" & db_SECTION2_15 & "','" & db_name_15 & "','" & db_level_15 & "'"
			SQL = SQL & "		,'" & db_homephone_15 & "','" & db_mobilephone_15 & "','" & db_etcphone_15 & "'"
			if db_MONITORDATE_15 = "" then
				SQL = SQL & "		,	getdate()"
			else
				SQL = SQL & "		,	'" & db_MONITORDATE_15 & "'"
			end if
			SQL = SQL & "		,'" & db_MONITORRESULT_15 & "'"

			if db_MONITORRESULT_15 = "4" then	'예약이라면
				if db_RESERVETIME_15 = "1" then
					SQL = SQL & "		,	CONVERT(CHAR(19),DATEADD(n,10,GETDATE()),121)"
				elseif db_RESERVETIME_15 = "2" then
					SQL = SQL & "		,	CONVERT(CHAR(19),DATEADD(n,30,GETDATE()),121)"
				elseif db_RESERVETIME_15 = "3" then
					SQL = SQL & "		,	CONVERT(CHAR(19),DATEADD(n,60,GETDATE()),121)"
				elseif db_RESERVETIME_15 = "4" then
					SQL = SQL & "		,	CONVERT(CHAR(19),DATEADD(n,120,GETDATE()),121)"
				elseif db_RESERVETIME_15 <> "" then
					SQL = SQL & "		,	'" & db_RESERVEDATE_15 & " " & db_RESERVETIME_15 & ":00:00'"
				else
					SQL = SQL & "		,	'" & db_RESERVEDATE_15 & " " & db_RESERVEHOUR_15 & ":"&  db_RESERVEMIN_15 &":00'"
				end if
			else
				SQL = SQL & "		,	''"
			end if
			SQL = SQL & "		,	'" & db_Remark_15 & "'"
			SQL = SQL & "		,	'" & db_Remark1_15 & "'"
			SQL = SQL & "		,	'" & db_TOT_15 & "')"

			db.execute(SQL)
	
		else

			SQL = " update armyinformix.dbo.factpeople set SECTION2 = '" & db_SECTION2_15 & "'"
			SQL = SQL & "		,	name = '" & db_name_15 & "',	level = '" & db_level_15 & "', homephone = '" & db_homephone_15 & "'"
			SQL = SQL & "		,	mobilephone = '" & db_mobilephone_15 & "', etcphone = '" & db_etcphone_15 & "'"

			if db_MONITORDATE_15 = "" then
				SQL = SQL & "		,	MONITORDATE = getdate()"
			else
				SQL = SQL & "		,	MONITORDATE = convert(datetime,'"&db_MONITORDATE_15&"')"
			end if
			SQL = SQL & "		,	MONITORRESULT = '" & db_MONITORRESULT_15 & "'"

			if db_MONITORRESULT_15 = "4" then	'예약이라면
				if db_RESERVETIME_15 = "1" then
					SQL = SQL & "		,	RESERVEDATE = CONVERT(CHAR(19),DATEADD(n,10,GETDATE()),121)"
				elseif db_RESERVETIME_15 = "2" then
					SQL = SQL & "		,	RESERVEDATE = CONVERT(CHAR(19),DATEADD(n,30,GETDATE()),121)"
				elseif db_RESERVETIME_15 = "3" then
					SQL = SQL & "		,	RESERVEDATE = CONVERT(CHAR(19),DATEADD(n,60,GETDATE()),121)"
				elseif db_RESERVETIME_15 = "4" then
					SQL = SQL & "		,	RESERVEDATE = CONVERT(CHAR(19),DATEADD(n,120,GETDATE()),121)"
				elseif db_RESERVETIME_15 <> "" then
					SQL = SQL & "		,	RESERVEDATE = '" & db_RESERVEDATE_15 & " " & db_RESERVETIME_15 & ":00:00'"
				else
					SQL = SQL & "		,	RESERVEDATE = '" & db_RESERVEDATE_15 & " " & db_RESERVEHOUR_15 & ":"&  db_RESERVEMIN_15 &":00'"
				end if
			else
				SQL = SQL & "		,	RESERVEDATE = ''"
			end if
			SQL = SQL & "		,	Remark = '" & db_Remark_15 & "'"
			SQL = SQL & "		,	Remark1 = '" & db_Remark1_15 & "'"
			SQL = SQL & "		,	monitorpoint = '" & db_TOT_15 & "'"
			SQL = SQL & "	where factPeoplenum = '" & db_factPeoplenum_15 & "' and factnum = '" & db_RECEIPTFACTNUM & "'"

			db.execute(SQL)
		end if

		'## 1번
		SQL = "	delete from armyinformix.dbo.monitor where factnum = '" & db_RECEIPTFACTNUM & "' and factPeoplenum = '" & db_factPeoplenum_15 & "'"
		db.execute(SQL)
		for i = 1 to icnt
			point9 =  Request("QUESTION_15"&i)
			if point9 = "9" then
				ipoint9 = "1"
			else
				ipoint9 = "0"
			end if
			if point9 = "8" then
				ipoint8 = "1"
			else
				ipoint8 = "0"
			end if
			if point9 = "7" then
				ipoint7 = "1"
			else
				ipoint7 = "0"
			end if
			pointplus =  Request("QUESTIONP_15"&i)
			if pointplus = "" then
				pointplus = "0"
			end if
			totpoint =  Request("POINT_15"&i)
			if totpoint = "" then
				totpoint = "0"
			end if

			SQL = "insert into armyinformix.dbo.monitor ( factnum, factPeoplenum, seqno, point9, point8, point7, pointplus, totpoint, monitordate, monitoruser)"
			SQL = SQL & " values ( '" & db_RECEIPTFACTNUM & "', '" & db_factPeoplenum_15 & "', " & i & ", " & ipoint9& ", " & ipoint8& ", " & ipoint7
			SQL = SQL & " , " & pointplus & ","&totpoint&",getdate(),'" &INCODE&"')"
			db.execute(SQL)
			'Response.write SQL
		next
	else
		'## 15번
		SQL = "delete from armyinformix.dbo.factpeople"
		SQL = SQL & "	where factPeoplenum = '" & db_factPeoplenum_15 & "' and factnum = '" & db_RECEIPTFACTNUM & "'"
		db.execute(SQL)

		SQL = "	delete from armyinformix.dbo.monitor where factnum = '" & db_RECEIPTFACTNUM & "' and factPeoplenum = '" & db_factPeoplenum_15 & "'"
		db.execute(SQL)
	end if


	if db_IDX1 <> "" then

		sql = "update armyinformix.dbo.contactlist set remark = '" & db_Remark1_1 & "', recordyn = '"& db_RECYN1 & "' where idx = " & db_IDX1
		db.execute(SQL)

	end if

	if db_IDX2 <> "" then

		sql = "update armyinformix.dbo.contactlist set remark = '" & db_Remark1_2 & "', recordyn = '"& db_RECYN1 & "'  where idx = " & db_IDX2
		db.execute(SQL)

	end if

	if db_IDX3 <> "" then

		sql = "update armyinformix.dbo.contactlist set remark = '" & db_Remark1_3 & "', recordyn = '"& db_RECYN1 & "'  where idx = " & db_IDX3
		db.execute(SQL)

	end if

	if db_IDX4 <> "" then

		sql = "update armyinformix.dbo.contactlist set remark = '" & db_Remark1_4 & "', recordyn = '"& db_RECYN1 & "'  where idx = " & db_IDX4
		db.execute(SQL)

	end if

	if db_IDX5 <> "" then

		sql = "update armyinformix.dbo.contactlist set remark = '" & db_Remark1_5 & "', recordyn = '"& db_RECYN1 & "'  where idx = " & db_IDX5
		db.execute(SQL)

	end if

	if db_IDX6 <> "" then

		sql = "update armyinformix.dbo.contactlist set remark = '" & db_Remark1_6 & "', recordyn = '"& db_RECYN1 & "'  where idx = " & db_IDX6
		db.execute(SQL)

	end if

	if db_IDX7 <> "" then

		sql = "update armyinformix.dbo.contactlist set remark = '" & db_Remark1_7 & "', recordyn = '"& db_RECYN1 & "'  where idx = " & db_IDX7
		db.execute(SQL)

	end if

	if db_IDX8 <> "" then

		sql = "update armyinformix.dbo.contactlist set remark = '" & db_Remark1_8 & "', recordyn = '"& db_RECYN1 & "'  where idx = " & db_IDX8
		db.execute(SQL)

	end if

	if db_IDX9 <> "" then

		sql = "update armyinformix.dbo.contactlist set remark = '" & db_Remark1_9 & "', recordyn = '"& db_RECYN1 & "'  where idx = " & db_IDX9
		db.execute(SQL)

	end if

	if db_IDX10 <> "" then

		sql = "update armyinformix.dbo.contactlist set remark = '" & db_Remark1_10 & "', recordyn = '"& db_RECYN1 & "'  where idx = " & db_IDX10
		db.execute(SQL)

	end if

	if db_IDX11 <> "" then

		sql = "update armyinformix.dbo.contactlist set remark = '" & db_Remark1_11 & "', recordyn = '"& db_RECYN1 & "'  where idx = " & db_IDX11
		db.execute(SQL)

	end if

	if db_IDX12 <> "" then

		sql = "update armyinformix.dbo.contactlist set remark = '" & db_Remark1_12 & "', recordyn = '"& db_RECYN1 & "'  where idx = " & db_IDX12
		db.execute(SQL)

	end if

	if db_IDX13 <> "" then

		sql = "update armyinformix.dbo.contactlist set remark = '" & db_Remark1_13 & "', recordyn = '"& db_RECYN13 & "'  where idx = " & db_IDX13
		db.execute(SQL)

	end if

	if db_IDX14 <> "" then

		sql = "update armyinformix.dbo.contactlist set remark = '" & db_Remark1_14 & "', recordyn = '"& db_RECYN14 & "'  where idx = " & db_IDX14
		db.execute(SQL)

	end if

	if db_IDX15 <> "" then

		sql = "update armyinformix.dbo.contactlist set remark = '" & db_Remark1_15 & "', recordyn = '"& db_RECYN15 & "'  where idx = " & db_IDX15
		db.execute(SQL)

	end if
'----------------------------------------------------------------------------------------
'#### 마스타테이블에 update시키기 - 한건이라도 예약건이 있으면 예약
'								    모두
'----------------------------------------------------------------------------------------
'한건이라도 예약이 있으면. 예약이 되며,
' 한건이라도 진행하지 않은 건이 있다면. 진행중이 된다.

	SQL = "select processgb from armyinformix.dbo.receiptfact where receiptfactnum = '" & db_RECEIPTFACTNUM & "'"
	set Rs = db.execute(SQL)

	sOldProcessgb = rs(0)

	SQL = "	SELECT  COUNT(*) FROM  armyinformix.dbo.factpeople"
	SQL = SQL & "	where factnum = '" & db_RECEIPTFACTNUM & "' and MONITORRESULT = '4'"

	set rs = db.execute(SQL)
	if rs(0) > 0 then
		'예약건이다.
		SQL = " update armyinformix.dbo.receiptfact set processdate = getdate(), processgb = '2', Date1 = '" & trim(db_Date2) &"',	Date2 = '" & trim(db_Date3) &"', receiptkind = '" & db_receiptkind & "'"
		SQL = SQL & " where receiptfactnum = '" & db_RECEIPTFACTNUM &"'"

		db.execute(SQL)

			pageUrl = "/menu01/submenu0101/monitoring_input.asp?receiptfactnum="&db_RECEIPTFACTNUM&"&sGubun="&db_SECTION2_1&"&sGubunName=피의자"
			Call MsgGoUrl( "정상적으로 등록되었습니다.",pageUrl)

	else

		SQL = "	SELECT  COUNT(*) FROM  armyinformix.dbo.factpeople"
		SQL = SQL & "	where factnum = '" & db_RECEIPTFACTNUM & "' and ( MONITORRESULT not in ('1','2','3','9') or MONITORRESULT is null ) "
		set rs = db.execute(SQL)
		if rs(0)>0 then

			SQL = " update armyinformix.dbo.receiptfact set processdate = getdate(), processgb = '1', Date1 = '" & trim(db_Date2) &"',	Date2 = '" & trim(db_Date3) &"', receiptkind = '" & db_receiptkind & "'"
			SQL = SQL & " where receiptfactnum = '" & db_RECEIPTFACTNUM &"'"
		
			db.execute(SQL)

			pageUrl = "/menu01/submenu0101/monitoring_input.asp?receiptfactnum="&db_RECEIPTFACTNUM&"&sGubun="&db_SECTION2_1&"&sGubunName="&db_GetCodeName("B01",db_SECTION2_1)
			Call MsgGoUrl( "정상적으로 등록되었습니다.",pageUrl)

		else
			'완료된 상태이다.
			'완료된 상태일때 평균점수를 다시 낸다.
			SQL = "	SELECT  COUNT(*) cnt, sum(isnull(monitorpoint,0)) sum FROM  armyinformix.dbo.factpeople"
			SQL = SQL & "	where factnum = '" & db_RECEIPTFACTNUM & "' and MONITORRESULT = '9'" '설문완료했을 때의 평균점수를 낸다.
			
			set rs = db.execute(SQL)

			if rs(0) > 0 then
								
				SQL = " update armyinformix.dbo.receiptfact set processdate = (select max(monitordate) from armyinformix.dbo.factpeople where factnum = '" & db_RECEIPTFACTNUM &"'), processgb = '9', Date1 = '" & db_Date2 &"',	Date2 = '" & db_Date3 &"', receiptkind = '" & db_receiptkind & "'"
				SQL = SQL & ",	monitorpoint = round(" & rs(1) / rs(0)&",2)"
				SQL = SQL & " where receiptfactnum = '" & db_RECEIPTFACTNUM &"'"
				db.execute(SQL)

%>
				<script>
					parent.location.reload();
				</script>

<%

			else

				SQL = " update armyinformix.dbo.receiptfact set processdate = getdate(), processgb = '9', Date1 = '" & db_Date2 &"',	Date2 = '" & db_Date3 &"', receiptkind = '" & db_receiptkind & "'"
				SQL = SQL & ",	monitorpoint = '0.00'"
				SQL = SQL & " where receiptfactnum = '" & db_RECEIPTFACTNUM &"'"
				db.execute(SQL)

%>
				<script>
					parent.location.reload();
				</script>

<%
			end if	

		end if

''response.write		SQL

	end if


			' 완료건일 때는 화면 전체를 refresh
			' 예약 또는 진행건은 해당 frame만 refresh한다.




%>