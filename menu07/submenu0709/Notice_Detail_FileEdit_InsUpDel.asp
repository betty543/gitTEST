<!-- #include virtual="/Include/Top_PopUp.asp" -->
<%
	'####### multipart/form-data #######################################################################
	Dim UploadForm
	Set UploadForm = Server.CreateObject("DEXT.FileUpload")
		UploadForm.DefaultPath = Server.MapPath("\Upload\Board\Notice") 	'업로드 경로

	'####### 폼값 받기 #################################################################################
	isType = UploadForm("isType")
	SEQ = UploadForm("SEQ")
	FILENAME_OLD = UploadForm("FILENAME_OLD")

	'####### 디버깅 코드 ###############################################################################
	'Response.Write("SEQ=" &SEQ& "<br>")
	'Response.Write("FILENAME_OLD=" &FILENAME_OLD& "<br>")
	'Response.End()
%>
<%
	On Error Resume next
	db.begintrans
	
	SELECT CASE UCASE(isType)
		'####### UPDATE ################################################################################
		CASE "UP"
			'======= 파일 업로드 =======================================================================
			Dim FilenameArr()		'업로드할 파일명을 위한 배열선언
			UploadCount = UploadForm("aFilename").Count		'업로드할 파일수
		
			ReDim FilenameArr(UploadCount-1) 	'배열 재선언
				FOR i = 1 to UploadCount
					'----> 멀티 화일업로드 모듈
					Dim attachfile,upfile,upfile1,filesize
					IF UploadForm("aFilename")(i) <> "" THEN
						if UploadForm("aFilename")(i).FileLen = 0 then
								Response.Write("<script>")
								Response.Write("	alert('올바른 파일이 아닙니다.');")
								Response.Write("	history.back();")
								Response.Write("</script>")
								Response.End
						end if
						attachfile = UploadForm("aFilename")(i).FilePath '전체 파일
						FileUpName = mid(attachfile,instrrev(attachfile,"\")+1) 'aaa.zip
						upfile = mid(FileUpName,1,instr(FileUpName,".")-1) 'aaa
						upfile1 = mid(FileUpName,instr(FileUpName,".")+1) 'zip
		
						if UploadForm("aFilename")(i).MimeType = "text/html" or upfile1 = Ucase("asp") or upfile1 = Ucase("dll") then
							Response.Write "<script>"
							Response.Write "	alert('압축해서 올려주시기 바랍니다.');"
							Response.Write "	history.back();"
							Response.Write "</script>"
							Response.End
						end if
		
						UploadForm.MaxFileLen = 5120000
						if UploadForm("aFilename")(i).FileLen > UploadForm.MaxFileLen then
							Response.Write("<script>")
							Response.Write("	alert('파일 용량은 5MG를 넘을 수 없습니다.');")
							Response.Write("	history.back();")
							Response.Write("</script>")
							Response.End
						end if
		
						Dim objFS,fexist,strfilenameAdd,count
						Set objFS = Server.CreateObject("Scripting.FileSysTemObject")
							'파일이 존재한다고 가정..
							fexist = true
							'저장한 파일의 완전한 이름을 만든다.
							strfilenameAdd = UploadForm.DefaultPath & "\" & FileUpName
							'만약에 파일이 존재할 경우 파일뒤에 붙일 번호를 만든다.
							count = 0
							while fexist = true
								if (objFS.FileExists(strfilenameAdd)) then
									count = count +1
									strfilenameAdd = UploadForm.DefaultPath & "\" & upfile & count & "." & upfile1
									FileUpName = upfile & count & "." & upfile1
								else
									fexist = false
								end if
							wend
						Set objFS = Nothing
		
						UploadForm("aFilename")(i).SaveAs strfilenameAdd
						FilenameArr(i-1) = FileUpName	'업로드할 화일명을 배열에 넣는다.
		
						'----> 업로드 에러메세지
						If Err then
						    Response.Write(Err.number & "<br>" & Err.source & "<br>" & Err.description)
						    Set UploadForm = Nothing
						    UploadForm.DeleteAllSavedFiles
						    Response.End()
						End if
					END IF
				NEXT
				
			SQL = "UPDATE TB_BOARD_NOTICE SET FILENAME1='" &FilenameArr(0)& "' WHERE IDX='" &SEQ& "'"
%>
<%
		'####### DELETE ################################################################################
		CASE "DEL"
			SQL = "UPDATE TB_BOARD_NOTICE SET FILENAME1='' WHERE IDX='" &SEQ& "'"
%>

<%
	END SELECT
%>

<%
	'####### DB 처리 ###################################################################################
	db.execute(SQL)
	'Response.Write(SQL&"<br>")
	'Response.Write("db.Errors.count=" &db.Errors.count)
	'Response.End()
	

	
	IF db.Errors.count <> 0 THEN
		'LogWrite "[SQL] "&SQL, "Notice_Detail_FileEdit_InsUpDel.asp", ""
		db.RollBackTrans
		Call UrlBack("처리중 에러가 발생했습니다.\n\n다시 시도해 주세요.")
	ELSE
		db.CommitTrans
		
		'====> 기존파일 삭제
		set Fso = Server.CreateObject("Scripting.FileSystemObject")
			tempFILENAME1 = UploadForm.DefaultPath & "\" & FILENAME_OLD
			IF Fso.FileExists(tempFILENAME1) THEN Fso.DeleteFile(tempFILENAME1) END IF
		set Fso = nothing
		
		Response.Write("<script>parent.location.reload();</script>")
		
	END IF




%>
<!-- #include virtual="/Include/Bottom_PopUp.asp" -->