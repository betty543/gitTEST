<!-- #include virtual="/include/CacheNo.asp" -->
<!-- #include virtual="/include/common.asp" -->

<%

guboon = Request("guboon")						'저장/수정/삭제 FLAG
'월요일
jobGb = Request("jobGb")	
DNIS = Request("DNIS")	

if jobGb = "C" then


	SQL = "UPDATE TB_TransferNo SET	OnPhone = 'N'"
	SQL = SQL & " Where DNIS = " & DNIS
	
	db.execute(SQL)


	Response.Write "<SCRIPT LANGUAGE=JavaScript>alert('정상적으로 수정되었습니다');" &_
						"document.location.href = '/Menu07/submenu0713/Transfer.asp';" &_
						"</SCRIPT>"

else

		TransferNo1 = Request("TransferNo1")
		TransferNo2 = Request("TransferNo2")
		TransferNo3 = Request("TransferNo3")
		TransferNo4 = Request("TransferNo4")

		UserId1 = Request("UserId1")
		UserId2 = Request("UserId2")
		UserId3 = Request("UserId3")
		UserId4 = Request("UserId4")


			'#################################################################################################################'
			'착신 전환 순서 관리하기'
			'#################################################################################################################'

			'DNIS, TransferNo, OnPhone,  UserId,	UpdateDate,	IN_DNIS
			SQL = "select * from TB_TransferNo 	where DNIS = 1"

			Set Rs = server.createObject("ADODB.Recordset")
			Rs.open SQL,db
			On Error Resume next
			db.begintrans

			If Rs.Eof Or Rs.bof Then
			
			Else

		'		SQL ="	UPDATE TB_PERSON_ETC	SET	 TB_PERSON_ETC (GIJUNDATE, INCODE, JOBGB,  WORKHOUR,	WORKDESC,	CHANGEHOUR,	PROCESSHOUR) VALUES "
				SQL = "UPDATE TB_TransferNo SET	TransferNo = '" &  trim(TransferNo1) & "', UserId = '" & UserId1 & "', updateDate = getdate()"
				SQL = SQL & " Where DNIS = 1"
				
				db.execute(SQL)

				SQL = "UPDATE TB_TransferNo SET	TransferNo = '" &  trim(TransferNo2) & "', UserId = '" & UserId2 & "', updateDate = getdate()"
				SQL = SQL & " Where DNIS = 2"
				
				db.execute(SQL)

				SQL = "UPDATE TB_TransferNo SET	TransferNo = '" &  trim(TransferNo3) & "', UserId = '" & UserId3 & "', updateDate = getdate()"
				SQL = SQL & " Where DNIS = 3"
				
				db.execute(SQL)

				SQL = "UPDATE TB_TransferNo SET	TransferNo = '" &  trim(TransferNo4) & "', UserId = '" & UserId4 & "', updateDate = getdate()"
				SQL = SQL & " Where DNIS = 4"
				
				db.execute(SQL)


				if db.Errors.count = 0 then
					db.CommitTrans

					Response.Write "<SCRIPT LANGUAGE=JavaScript>alert('정상적으로 수정되었습니다');" &_
										"document.location.href = '/Menu07/submenu0713/Transfer.asp';" &_
										"</SCRIPT>"
				else
					db.RollBackTrans
					'Response.Write SQL

					Call UrlBack("저장중 에러가 발생했습니다.\n\n다시 시도해 주세요")
				end if		
			
			End If

	end if

%>