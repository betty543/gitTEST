<!-- #include virtual="/Include/Top_PopUp.asp" -->
<%
	'####### multipart/form-data #######################################################################
	Dim UploadForm
	Set UploadForm = Server.CreateObject("DEXT.FileUpload")
		UploadForm.DefaultPath = Server.MapPath("\Upload\Board\Notice") 	'���ε� ���

	'####### ���� �ޱ� #################################################################################
	isType = UploadForm("isType")
	SEQ = UploadForm("SEQ")
	FILENAME_OLD = UploadForm("FILENAME_OLD")

	'####### ����� �ڵ� ###############################################################################
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
			'======= ���� ���ε� =======================================================================
			Dim FilenameArr()		'���ε��� ���ϸ��� ���� �迭����
			UploadCount = UploadForm("aFilename").Count		'���ε��� ���ϼ�
		
			ReDim FilenameArr(UploadCount-1) 	'�迭 �缱��
				FOR i = 1 to UploadCount
					'----> ��Ƽ ȭ�Ͼ��ε� ���
					Dim attachfile,upfile,upfile1,filesize
					IF UploadForm("aFilename")(i) <> "" THEN
						if UploadForm("aFilename")(i).FileLen = 0 then
								Response.Write("<script>")
								Response.Write("	alert('�ùٸ� ������ �ƴմϴ�.');")
								Response.Write("	history.back();")
								Response.Write("</script>")
								Response.End
						end if
						attachfile = UploadForm("aFilename")(i).FilePath '��ü ����
						FileUpName = mid(attachfile,instrrev(attachfile,"\")+1) 'aaa.zip
						upfile = mid(FileUpName,1,instr(FileUpName,".")-1) 'aaa
						upfile1 = mid(FileUpName,instr(FileUpName,".")+1) 'zip
		
						if UploadForm("aFilename")(i).MimeType = "text/html" or upfile1 = Ucase("asp") or upfile1 = Ucase("dll") then
							Response.Write "<script>"
							Response.Write "	alert('�����ؼ� �÷��ֽñ� �ٶ��ϴ�.');"
							Response.Write "	history.back();"
							Response.Write "</script>"
							Response.End
						end if
		
						UploadForm.MaxFileLen = 5120000
						if UploadForm("aFilename")(i).FileLen > UploadForm.MaxFileLen then
							Response.Write("<script>")
							Response.Write("	alert('���� �뷮�� 5MG�� ���� �� �����ϴ�.');")
							Response.Write("	history.back();")
							Response.Write("</script>")
							Response.End
						end if
		
						Dim objFS,fexist,strfilenameAdd,count
						Set objFS = Server.CreateObject("Scripting.FileSysTemObject")
							'������ �����Ѵٰ� ����..
							fexist = true
							'������ ������ ������ �̸��� �����.
							strfilenameAdd = UploadForm.DefaultPath & "\" & FileUpName
							'���࿡ ������ ������ ��� ���ϵڿ� ���� ��ȣ�� �����.
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
						FilenameArr(i-1) = FileUpName	'���ε��� ȭ�ϸ��� �迭�� �ִ´�.
		
						'----> ���ε� �����޼���
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
	'####### DB ó�� ###################################################################################
	db.execute(SQL)
	'Response.Write(SQL&"<br>")
	'Response.Write("db.Errors.count=" &db.Errors.count)
	'Response.End()
	

	
	IF db.Errors.count <> 0 THEN
		'LogWrite "[SQL] "&SQL, "Notice_Detail_FileEdit_InsUpDel.asp", ""
		db.RollBackTrans
		Call UrlBack("ó���� ������ �߻��߽��ϴ�.\n\n�ٽ� �õ��� �ּ���.")
	ELSE
		db.CommitTrans
		
		'====> �������� ����
		set Fso = Server.CreateObject("Scripting.FileSystemObject")
			tempFILENAME1 = UploadForm.DefaultPath & "\" & FILENAME_OLD
			IF Fso.FileExists(tempFILENAME1) THEN Fso.DeleteFile(tempFILENAME1) END IF
		set Fso = nothing
		
		Response.Write("<script>parent.location.reload();</script>")
		
	END IF




%>
<!-- #include virtual="/Include/Bottom_PopUp.asp" -->