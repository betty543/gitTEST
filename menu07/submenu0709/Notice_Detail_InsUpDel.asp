<!-- #include virtual="/Include/Top2.asp" -->
<%
	'####### multipart/form-data #######################################################################
	Dim UploadForm
	Set UploadForm = Server.CreateObject("DEXT.FileUpload")
		UploadForm.DefaultPath = Server.MapPath("\Board\Notice") 	'���ε� ���

	'####### ���� �ޱ� #################################################################################
	isType = UploadForm("isType")
	curPage = UploadForm("curPage")
	SEQ = UploadForm("SEQ")
	ACLASS = UploadForm("ACLASS")
	TITLE = ConvertString(UploadForm("TITLE"))
	CONTENTS = ConvertString(UploadForm("CONTENTS"))
	FILENAME1 = UploadForm("FILENAME1")
	FRONTYN = UploadForm("FRONTYN")
	IF FRONTYN="" THEN FRONTYN="N" END IF
		
	'####### ����� �ڵ� ###############################################################################
	'Response.Write("isType=" &isType& "<br>")
	'Response.Write("SEQ=" &SEQ& "<br>")
	'Response.Write("curPage=" &curPage& "<br>")
	'Response.End()
%>

<%
	On Error Resume next
	db.begintrans
	
	SELECT CASE UCASE(isType)
		'####### INSERT ################################################################################
		CASE "INS"
			pageURL ="Notice.asp"

			'======= ���� ���ε� =======================================================================
			Dim FilenameArr()		'���ε��� ���ϸ��� ���� �迭����
			UploadCount = UploadForm("aFilename").Count		'���ε��� ���ϼ�

			ReDim FilenameArr(UploadCount-1) 	'�迭 �缱��
				for i = 1 to UploadCount
					'----> ��Ƽ ȭ�Ͼ��ε� ���
					Dim attachfile,upfile,upfile1,filesize
					if UploadForm("aFilename")(i) <> "" then
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
					end if
				next

			'======= ����� �ڵ� ======================================================================
			'----> �迭�� �ִ� ȭ�ϸ�
			'for j = 0 to UploadCount-1
			'	Response.Write("FileName(" &j& " ) = " &FilenameArr(j)& "<br>")
			'next
			'Response.End()
			SQL = "INSERT INTO TB_BOARD_NOTICE (	ACLASS"
			SQL = SQL& ", TITLE"
			SQL = SQL& ", CONTENTS"
			SQL = SQL& ", FILENAME1"
			SQL = SQL& ", INCODE,	INDATE"
			SQL = SQL& ", FRONTYN"
			SQL = SQL& ")"
			SQL = SQL& " values ("
			SQL = SQL& "'" &ACLASS& "'"
			SQL = SQL& ",'" &TITLE& "'"
			SQL = SQL& ",'" &Left(CONTENTS,3000)& "'"
			SQL = SQL& ",'" &FilenameArr(0)& "'"
			SQL = SQL& ",'" &SESSION("SS_LoginID")& "',GETDATE()"
			SQL = SQL& ",'" &FRONTYN& "'"
			SQL = SQL& ")"
%>
<%
		'####### UPDATE ################################################################################
		CASE "UP"
			pageURL = "Notice_Detail.asp?isType=VIEW&SEQ=" &SEQ& "&curPage=" & curPage
			
			SQL = "UPDATE TB_BOARD_NOTICE SET"
			SQL = SQL& " ACLASS='" &ACLASS& "'"
			SQL = SQL& ", TITLE='" &TITLE& "'"
			SQL = SQL& ", CONTENTS='" &Left(CONTENTS,3000)& "'"
			SQL = SQL& ", FRONTYN='" &FRONTYN& "'"
			SQL = SQL& ", MOCODE='" &SESSION("SS_LoginID")& "'"
			SQL = SQL& ", MODATE=GETDATE()"
			SQL = SQL& " WHERE IDX='" &SEQ& "'"
%>
<%
		'####### DELETE ################################################################################
		CASE "DEL"
			pageURL = "Notice.asp"
			
			SQL = "DELETE FROM	TB_BOARD_NOTICE WHERE IDX='" &SEQ& "'"

	END SELECT
%>

<%
	'####### DB ó�� ###################################################################################
	db.execute(SQL)
	'Response.Write(SQL&"<br>")
	'Response.Write("db.Errors.count=" &db.Errors.count)
	'Response.End()

	IF db.Errors.count <> 0 THEN
		LogWrite "[SQL] "&SQL, "Notice_Detail_InsUpDel.asp", ""
		db.RollBackTrans
		Call UrlBack("[#1]ó���� ������ �߻��߽��ϴ�.\n\n�ٽ� �õ��� �ּ���.")
	Else
	
		If UCASE(isType) = "INS" Or UCASE(isType) = "UP" Then
		
			If UCASE(isType) = "INS" Then
				SQL = "SELECT	MAX(IDX) HSEQ	FROM	TB_BOARD_NOTICE"
				Set RS = db.execute(SQL)
				If IsNull(RS("HSEQ")) = False Then
					SEQ = RS("HSEQ")
				End IF
			End if
			ErrorsFlag = db_TextINS("TB_BOARD_NOTICE_DETAIL","HIDX",SEQ,CONTENTS)
		Else
			SQL = "DELETE FROM	TB_BOARD_NOTICE_DETAIL WHERE HIDX='" &SEQ& "'"
			db.execute(SQL)
			IF db.Errors.count <> 0 Then
				ErrorsFlag = "N"
			Else
				ErrorsFlag  = "Y"
			End if
		End If
		IF ErrorsFlag = "Y" THEN
			db.CommitTrans			
			'====> ���� ����
			set Fso = Server.CreateObject("Scripting.FileSystemObject")
				tempFILENAME1 = UploadForm.DefaultPath & "\" & FILENAME1
				IF Fso.FileExists(tempFILENAME1) THEN Fso.DeleteFile(tempFILENAME1) END IF
			set Fso = nothing 
				
			Call MsgGoUrl( "���������� ó�� �Ǿ����ϴ�.",pageURL)
		Else
			LogWrite "[SQL] "&SQL, "Notice_Detail_InsUpDel.asp", ""
			db.RollBackTrans
			Call UrlBack("[#2]ó���� ������ �߻��߽��ϴ�.\n\n�ٽ� �õ��� �ּ���.")		
		End if
	END IF
%>
<!-- #include virtual="/Include/Bottom.asp" -->>