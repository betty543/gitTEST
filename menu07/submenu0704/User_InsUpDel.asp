<!-- #include virtual="/include/CacheNo.asp" -->
<!-- #include virtual="/include/common.asp" -->

<%

guboon = Request("guboon")
sUSERID = Request("sUSERID")
sUSERNAME = Request("sUSERNAME")
sPASSWORD = Request("sPASSWORD")
sSECGROUP = Request("sSECGROUP")
sGRADE = Request("sGRADE")	
sCTIYN = Request("sCTIYN")	
sCTIID = Request("sCTIID")
sCTIPASSWORD = ConvertString(Request("sCTIPASSWORD"))
sEXTNO = ConvertString(Request("sEXTNO"))
sUSEYN = Request("sUSEYN")	
sIPDATE = Request("sIPDATE")	
sOUTDATE = trim(Request("sOUTDATE"))
Incode = SESSION("SS_LoginID")
sGunNumber= Request("sGunNumber")
sLevel= Request("sLevel")
sSosok= Request("sSosok")

If sOUTDATE <> "" Then
	sUSEYN = "N"
End if

If guboon = "INS"Then 

	Set Rs = server.createObject("ADODB.Recordset")
	SQL = "select userid from TB_USERINFO where userid = '" & sUSERID & "'"

	Rs.open SQL,db
	
	If Rs.Eof Or Rs.bof Then
		On Error Resume next
		db.begintrans


		SQL =				" insert into TB_USERINFO (USERID,USERNAME,PASSWORD,SECGROUP,GRADE,USEYN,IPDATE,OUTDATE,CTIYN,CTIID,CTIPASSWORD,EXTNO,INCODE,SOSOK,LEVEL)"
		SQL = SQL &	"  values('" & sUSERID & "', '" & sUSERNAME & "', '" & sPASSWORD & "', '" & sSECGROUP & "', '" & sGRADE & "', '" & sUSEYN 
		SQL = SQL &	"', '" & sIPDATE & "', '" & sOUTDATE & "', '"& sCTIYN & "', '" & sCTIID & "', '"& sCTIPASSWORD & "', '" & sEXTNO & "', '" & Incode & "', '" & sSOSOK & "', '" & sLEVEL & "')"

		'Response.Write SQL
		db.execute(SQL)

		If db.Errors.count = 0 Then 
			db.CommitTrans
			response.write "<script>parent.ListFrame.location.href='User_List.asp';</script>"
			Call MsgGoUrl("정상적으로 저장되었습니다.", "User_Detail.asp?guboon=INS")
		Else 
			db.RollBackTrans
			Call UrlBack("저장중 에러가 발생했습니다.\n\n다시 시도해 주세요")
		End If 

	Else

		Call UrlBack("이미 존재하는 아이디 입니다..\n\n다른 아이디를 입력해 주세요")
	
	end If
	
	Rs.Close
	set Rs = NOTHING	



ElseIf guboon = "UP" Then

		On Error Resume next
		db.begintrans

		SQL =				" update TB_USERINFO set"
		SQL = SQL &	"	USERNAME = '" & sUSERNAME & "', "
		SQL = SQL &	"	PASSWORD = '" & sPASSWORD & "', "
		SQL = SQL &	"	SECGROUP = '" & sSECGROUP & "', "
		SQL = SQL &	"	SOSOK = '" & sSOSOK & "', "
		SQL = SQL &	"	[LEVEL] = '" & sLEVEL & "', "
		SQL = SQL &	"	GUNNUMBER = '" & sGUNNUMBER & "', "
		SQL = SQL &	"	GRADE = '" & sGRADE & "', "
		SQL = SQL &	"	USEYN = '" & sUSEYN & "', "
		SQL = SQL &	"	IPDATE = '" & sIPDATE & "', "
		SQL = SQL &	"	OUTDATE = '" & sOUTDATE & "', "
		SQL = SQL &	"	CTIYN = '" & sCTIYN & "', "
		SQL = SQL &	"	CTIID = '" & sCTIID & "', "
		SQL = SQL &	"	CTIPASSWORD = '" & sCTIPASSWORD & "', "

		SQL = SQL &	"	EXTNO = '" & sEXTNO & "', "
		SQL = SQL &	"	MOCODE = '" & Incode & "', "
		SQL = SQL &	"	MODATE = getdate()"
		SQL = SQL &	" where USERID = '"& sUSERID & "'"
'
		db.execute(SQL)


		If db.Errors.count = 0 Then 
			db.CommitTrans
			response.write "<script>parent.ListFrame.location.reload();</script>"
			Call MsgGoUrl("정상적으로 저장되었습니다.", "User_Detail.asp?guboon=INS")
		Else 
			db.RollBackTrans
			Call UrlBack("저장중 에러가 발생했습니다.\n\n다시 시도해 주세요")
		End If 


ElseIf guboon = "DEL" Then 

	On Error Resume next
	db.begintrans

	SQL ="delete from TB_USERINFO where userid = '" & sUSERID & "'"
	db.execute(SQL)

	If db.Errors.count = 0 Then 
		db.CommitTrans
		response.write "<script>parent.ListFrame.location.href='User_List.asp';</script>"
		Call MsgGoUrl("정상적으로 삭제되었습니다.", "/Setup/User/User_Detail.asp?guboon=INS")
	Else 
		db.RollBackTrans
		Call UrlBack("삭제중 에러가 발생했습니다.\n\n다시 시도해 주세요")
	End If 

End If 

%>