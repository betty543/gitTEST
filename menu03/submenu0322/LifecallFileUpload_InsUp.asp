<!-- #include virtual="/Include/Top_PopUp.asp" -->
<%
	'####### multipart/form-data #######################################################################
	Dim UploadForm
	Set UploadForm = Server.CreateObject("DEXT.FileUpload")
		UploadForm.DefaultPath = Server.MapPath("\Upload\Lifecall") 	'���ε� ���

	'####### ���� �ޱ� #################################################################################
	isType = UploadForm("isType")
%>

<%
	SELECT CASE UCASE(isType)
		'####### INSERT ################################################################################
		CASE "INS"
			'======= ���� �ޱ� =========================================================================
			frmTYPE = UploadForm("frmTYPE")

			'======= ����� �ڵ� =======================================================================
			'Response.Write("frmTYPE=" &frmTYPE& "<br>")

			'======= �̹��� ���ε� =====================================================================
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
		%>

		<%'======= ȭ�Ͼ��ε� �Ϸ� =====================================================================%>
		<script language="javascript">
		<!--
			parent.document.all.FILENAME1.value = "<%=FilenameArr(0)%>";
			parent.document.getElementById('txtFILENAME1').innerHTML = "<%=FilenameArr(0)%>&nbsp;<img src='/Images/Comm/IconDel2.gif' title='<%=FilenameArr(0)%> ����' style='cursor:hand;' align='absmiddle' onClick=\"FileDel('<%=frmTYPE%>','<%=FilenameArr(0)%>')\">&nbsp;";
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