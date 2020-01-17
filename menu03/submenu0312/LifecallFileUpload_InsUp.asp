<!-- #include virtual="/Include/Top_PopUp.asp" -->
<%
	'####### multipart/form-data #######################################################################
	Dim UploadForm
	Set UploadForm = Server.CreateObject("DEXT.FileUpload")
		UploadForm.DefaultPath = Server.MapPath("\Upload\Lifecall") 	'업로드 경로

	'####### 폼값 받기 #################################################################################
	isType = UploadForm("isType")
%>

<%
	SELECT CASE UCASE(isType)
		'####### INSERT ################################################################################
		CASE "INS"
			'======= 폼값 받기 =========================================================================
			frmTYPE = UploadForm("frmTYPE")

			'======= 디버깅 코드 =======================================================================
			'Response.Write("frmTYPE=" &frmTYPE& "<br>")

			'======= 이미지 업로드 =====================================================================
			Dim FilenameArr()		'업로드할 파일명을 위한 배열선언
			UploadCount = UploadForm("aFilename").Count		'업로드할 파일수

			ReDim FilenameArr(UploadCount-1) 	'배열 재선언
				for i = 1 to UploadCount
					'----> 멀티 화일업로드 모듈
					Dim attachfile,upfile,upfile1,filesize
					if UploadForm("aFilename")(i) <> "" then
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
					end if
				next

			'======= 디버깅 코드 ======================================================================
			'----> 배열에 있는 화일명
			'for j = 0 to UploadCount-1
			'	Response.Write("FileName(" &j& " ) = " &FilenameArr(j)& "<br>")
			'next
			'Response.End()
		%>

		<%'======= 화일업로드 완료 =====================================================================%>
		<script language="javascript">
		<!--
			parent.document.all.FILENAME1.value = "<%=FilenameArr(0)%>";
			parent.document.getElementById('txtFILENAME1').innerHTML = "<%=FilenameArr(0)%>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='<%=FilenameArr(0)%> 삭제' style='cursor:hand;' align='absmiddle' onClick=\"FileDel('<%=frmTYPE%>','<%=FilenameArr(0)%>')\">&nbsp;";
			parent.HddnPOPLayer();
		//-->
		</script>


<%
		'####### UPDATE ################################################################################
		CASE "UP"
%>

<%
		'####### DELETE ################################################################################
		CASE "DEL"

	END SELECT
%>

<!-- #include virtual="/Include/Bottom_PopUp.asp" -->