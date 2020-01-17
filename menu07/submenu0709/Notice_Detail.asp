<!-- #include virtual="/Include/Top.asp" -->
<%
	'####### 파라미터 ##################################################################################
	isType = Trim(Request("isType"))
	SEQ = Trim(Request("SEQ"))
	curPage = Trim(Request("curPage"))
	SMODE = TRIM(request("SMODE"))
	SWORD = TRIM(request("SWORD"))
	ACLASS = TRIM(request("ACLASS"))
	
	'####### 디버깅 코드 ###############################################################################
	'Response.Write("isType=" &isType& "<br>")
	'Response.Write("curPage=" &curPage& "<br>")
	'Response.Write("TemplateName=" &TemplateName& "<br>")
	'Response.Write("UseYN=" &UseYN& "<br>")

	pageWHERE= "curPage=" &curPage& "&ACLASS=" &ACLASS& "&SMODE=" &SMODE& "&SWORD=" &SWORD
%>

<script>
<!--
	function fn_inup() {
		if(!ComboChk(inUpFrm.ACLASS,"분류")) return false;
		if(!FieldChk(inUpFrm.TITLE,"제목")) return false;
		if(!FieldChk(inUpFrm.CONTENTS,"내용")) return false;
	}

	function fn_list() {
		location.href="Notice.asp?<%=pageWHERE%>"
	}

	function fn_edit() {
		location.href="Notice_Detail.asp?isType=UP&SEQ=<%=SEQ%>&<%=pageWHERE%>"
	}

	function fn_del() {
		var answer = confirm("정말 삭제하시겠습니까?");
		if(answer == true){
			inUpFrm.submit();
		}
	}

	function FileEdit(fn,f){
		POPLayerURL = "Notice_Detail_FileEdit.asp?FILENAME=" +fn+ "&SEQ=" +f;
		ShowPOPLayer(POPLayerURL,'500','160');
	}

	function FileDel(fn,f){
		var answer = confirm("「" +fn+ " 」을 정말 삭제하시겠습니까?");
		if(answer == true){
			FileDelFrm.submit();
			//document.location.href = "Notice_Detail_InsUpDel.asp?isType=DEL&SEQ=" +f;
		}
	}
//-->
</script>

<%
	SELECT CASE UCASE(isType)
		'####### INSERT ################################################################################
		CASE "INS"
%>
	<script>

	</script>

	<form name="inUpFrm" method="post" action="Notice_Detail_InsUpDel.asp" onsubmit="return fn_inup(this);" encType="multipart/form-data" style="margin:0">
	<input type="hidden" name="isType" value="<%=isType%>">
	<table width="940" border="0" cellspacing="1" cellpadding="1" bgcolor="#CCCCCC" align="center">
	    <tr>
	        <td width="100" nowrap bgcolor="#EEF6FF" class="TDCont" align='center'>분류</td>
	        <td bgcolor="#FFFFFF">
	        	<select name="ACLASS" size="1" class="ComboFFFCE7">
					<option value="">분류 선택</option>
					<%=db_getTBCodeSelect("Z04", "", "N")%>
				</select>
	        </td>
	    </tr>
	    <tr>
	        <td width="100" nowrap bgcolor="#EEF6FF" class="TDCont" align='center'>제목</td>
	        <td bgcolor="#FFFFFF"><input type="text" name="TITLE" value="" maxlength="50" size="50" onfocus="setFocusColor(this);" onblur="setOutColor(this);"> <input type="checkbox" name="FRONTYN" value="Y" class="none" checked> 체크하시면, 상단에 보여집니다.</td>
	    </tr>
	    <tr>
	        <td bgcolor="#EEF6FF" class="TDCont" align='center'>내용</td>
	        <td bgcolor="#FFFFFF"><textarea name="CONTENTS" style="width:100%; height:300" wrap="soft" class="TextareaInput"><%=db_Contents%></textarea> </td>
		</tr>
		<tr>
	        <td bgcolor="#EEF6FF" class="TDCont" align='center'>첨부파일</td>
	        <td bgcolor="#FFFFFF">
	        	<input type="file" size="30" name="aFilename" onfocus="setFocusColor(this);" onblur="setOutColor(this);">
	        </td>
		</tr>
	</table>
	<table width="940" border="0" cellspacing="0" cellpadding="0" align="center">
		<tr>
			<td height="30" align="right">
				<input type="image" src="/Images/Btn/BtnSubmit.gif" style="cursor:hand;" align="absmiddle" class="none">
				<img src="/Images/Btn/BtnReset.gif" style="cursor:hand;" align="absmiddle" onClick="inUpFrm.reset();">
				<img src="/Images/Btn/BtnList.gif" style="cursor:hand;" align="absmiddle" onClick="javascript:fn_list();">
			</td>
		</tr>
	</form>
	</table>
<%
		'####### SELECT ################################################################################
		CASE "VIEW"
%>
	<%
		IF NOT(Request.Cookies("CK_Board_Notice_SEQ")=SEQ) THEN	'---> 조회수 증가
			db.execute("UPDATE TB_BOARD_NOTICE SET READCNT=READCNT+1 WHERE IDX='" &SEQ& "'")
		END IF
		Response.Cookies("CK_Board_Notice_SEQ") = SEQ

		SqlContact = "SELECT ACLASS, TITLE, FRONTYN, CONTENTS, FILENAME1, READCNT, INDATE, INCODE"
		SqlContact = SqlContact& " FROM TB_BOARD_NOTICE"
		SqlContact = SqlContact& " WHERE IDX='" &SEQ& "'"
		set RsContact = db.execute(SqlContact)

		IF NOT(RsContact.Eof Or RsContact.bof) THEN
			db_ACLASS = RsContact("ACLASS")
			db_FRONTYN = RsContact("FRONTYN")
			db_TITLE = RsContact("TITLE")
			'db_CONTENTS = OriginalContent(RsContact("CONTENTS"))
			db_READCOUNT = RsContact("READCNT")
			db_INDATE = RsContact("INDATE")
			db_INCODE = db_getUserName(RsContact("INCODE"))
			db_FILENAME1 = RsContact("FILENAME1")	'화일확장자 구분
			IF len(db_FILENAME1)>0 THEN
				Filename_Temp = split(db_FILENAME1,".")
				FileType = FormatFile(Filename_Temp(1))
			END If
			db_CONTENTS = db_TextSELECT("TB_BOARD_NOTICE_DETAIL","HIDX",SEQ)
		END IF

		RsContact.Close
		set RsContact = NOTHING
	%>
	<form name="inUpFrm" method="post" action="Notice_Detail_InsUpDel.asp" encType="multipart/form-data" style="margin:0">
	<input type="hidden" name="isType" value="DEL">
	<input type="hidden" name="SEQ" value="<%=SEQ%>">
	<input type="hidden" name="FILENAME1" value="<%=db_FILENAME1%>">
	</form>
	<table width="940" border="0" cellspacing="1" cellpadding="1" bgcolor="#CCCCCC" align="center">
	    <tr>
	        <td bgcolor="#EEF6FF" class="TDCont" align='center'>제목</td>
	        <td bgcolor="#FFFFFF" colspan="3"><input type="checkbox" name="FRONTYN" <%IF db_FRONTYN="Y" THEN%>checked<%END IF%> class="none" title="상단에 우선적으로 보이도록 셋팅!" disabled> <font color="#FF0000">[<%=db_getCodeName("Z04",db_ACLASS)%>]</font> <%=db_TITLE%></td>
	    </tr>
	    <tr>
	        <td bgcolor="#EEF6FF" class="TDCont" align='center'>게시자</td>
	        <td bgcolor="#FFFFFF" colspan="3" align='center'><%=db_INCODE%></td>
	    </tr>
	    <tr>
	        <td width="100" bgcolor="#EEF6FF" class="TDCont" align='center'>등록일</td>
	        <td width="900" bgcolor="#FFFFFF" class="TDCont"><%=db_INDATE%></td>
	        <td width="100" bgcolor="#EEF6FF" class="TDCont" align='center'>조회수</td>
	        <td width="100" bgcolor="#FFFFFF" class="TDCont"><%=db_READCOUNT%></td>
		</tr>
	    <tr>
	        <td bgcolor="#EEF6FF" class="TDCont" align='center'>내용</td>
	        <td bgcolor="#FFFFFF" colspan="3" class="TDCont"><textarea name="CONTENTS" style="width:100%; height:300" wrap="soft" class="TextareaInput" readonly><%=db_CONTENTS%></textarea></td>
		</tr>
		<%IF len(db_FILENAME1)>0 THEN%>
		<tr>
			<td bgcolor="#EEF6FF" class="TDCont" align='center'>첨부파일</td>
			<td bgcolor="#FFFFFF" colspan="3" class="TDCont"><a href="/Upload/Board/Notice/Download.asp?filename=<%=db_FILENAME1%>"><img src="/Images/File/<%=FileType%>" align="absmiddle"> <%=db_FILENAME1%></a></td>
		</tr>
		<%END IF%>
	</table>
	<table width="940" border="0" cellspacing="0" cellpadding="0" align="center">
		<tr>
			<td height="30" align="right">
				<img src="/Images/Btn/BtnDel.gif" style="cursor:hand;" align="absmiddle" onClick="javascript:fn_del();">
				<img src="/Images/Btn/BtnEdit.gif" style="cursor:hand;" align="absmiddle" onClick="javascript:fn_edit();">
				<img src="/Images/Btn/BtnList.gif" style="cursor:hand;" align="absmiddle" onClick="javascript:fn_list();">
			</td>
		</tr>
	</table>
<%
		'####### UPDATE ################################################################################
		CASE "UP"
%>
	<%
		SqlContact = "SELECT ACLASS, FRONTYN, TITLE, CONTENTS, FILENAME1, READCNT, INDATE"
		SqlContact = SqlContact& " FROM TB_BOARD_NOTICE"
		SqlContact = SqlContact& " WHERE IDX='" &SEQ& "'"
		set RsContact = db.execute(SqlContact)

		IF NOT(RsContact.Eof Or RsContact.bof) THEN
			db_ACLASS = RsContact("ACLASS")
			db_FRONTYN = RsContact("FRONTYN")
			db_TITLE = RsContact("TITLE")
			'db_CONTENTS = RsContact("CONTENTS")
			db_READCOUNT = RsContact("READCNT")
			db_INDATE = RsContact("INDATE")
			db_FILENAME1 = RsContact("FILENAME1")	'화일확장자 구분
			IF len(db_FILENAME1)>0 THEN
				Filename_Temp = split(db_FILENAME1,".")
				FileType = FormatFile(Filename_Temp(1))
			END If
			db_CONTENTS = db_TextSELECT("TB_BOARD_NOTICE_DETAIL","HIDX",SEQ)
		END IF

		RsContact.Close
		set RsContact = NOTHING
	%>
	<form name="inUpFrm" method="post" action="Notice_Detail_InsUpDel.asp" onsubmit="return fn_inup(this);" encType="multipart/form-data" style="margin:0">
	<input type="hidden" name="isType" value="UP">
	<input type="hidden" name="curPage" value="<%=curPage%>">
	<input type="hidden" name="SEQ" value="<%=SEQ%>">

	<table width="940" border="0" cellspacing="1" cellpadding="1" bgcolor="#CCCCCC" align="center">
	    <tr>
	        <td width="100" nowrap bgcolor="#EEF6FF" class="TDCont" align='center'>분류</td>
	        <td bgcolor="#FFFFFF">
	        	<select name="ACLASS" size="1" class="ComboFFFCE7">
					<%=db_getTBCodeSelect("Z04", db_ACLASS, "N")%>
				</select>
	        </td>
	    </tr>
	    <tr>
	        <td width="100" nowrap bgcolor="#EEF6FF" class="TDCont" align='center'>제목</td>
	        <td bgcolor="#FFFFFF"><input type="text" name="TITLE" value="<%=db_TITLE%>" maxlength="50" size="50" onfocus="setFocusColor(this);" onblur="setOutColor(this);"> <input type="checkbox" name="FRONTYN" value="Y" class="none" <%IF db_FRONTYN="Y" THEN%>checked<%END IF%> onfocus="blur();"> 체크하시면, 상단에 보여집니다.</td>
	    </tr>
	    <tr>
	        <td bgcolor="#EEF6FF" class="TDCont" align='center'>내용</td>
	        <td bgcolor="#FFFFFF"><textarea name="CONTENTS" style="width:100%; height:300" wrap="soft" class="TextareaInput"><%=db_CONTENTS%></textarea> </td>
		</tr>
		<tr>
	        <td bgcolor="#EEF6FF" class="TDCont" align='center'>첨부파일</td>
	        <td bgcolor="#FFFFFF" class="TDCont">
	        	<%IF len(db_FILENAME1)>0 THEN%>
				<img src="/Images/File/<%=FileType%>" align="absmiddle"> <%=db_FILENAME1%> &nbsp;&nbsp;
				<img src="/Images/Btn/BtnFileEdit.gif" style="cursor:hand;" align="absmiddle" onClick="FileEdit('<%=db_FILENAME1%>','<%=SEQ%>');">  
				<img src="/Images/Btn/BtnFileDel.gif" style="cursor:hand;" align="absmiddle" onClick="FileDel('<%=db_FILENAME1%>','<%=SEQ%>');">
				<%ELSE%>
				<img src="/Images/Btn/BtnFileUpload.gif" style="cursor:hand;" align="absmiddle" onClick="FileEdit('','<%=SEQ%>');">
				<%END IF%>
	        </td>
		</tr>
	</table>
	<table width="940" border="0" cellspacing="0" cellpadding="0" align="center">
		<tr>
			<td height="30" align="right">
				<input type="image" src="/Images/Btn/BtnSubmit.gif" style="cursor:hand;" align="absmiddle" class="none">
				<img src="/Images/Btn/BtnReset.gif" style="cursor:hand;" align="absmiddle" onClick="inUpFrm.reset();">
				<img src="/Images/Btn/BtnDel.gif" style="cursor:hand;" align="absmiddle" onClick="fn_del();">
				<img src="/Images/Btn/BtnList.gif" style="cursor:hand;" align="absmiddle" onClick="fn_list();">
			</td>
		</tr>
	</form>
	</table>
<%
	END SELECT
%>


<%'======= 화일삭제 =======================================================================================%>
<DIV id="hiddenIframe" style="display:none;">
	<iframe SRC="about:blank" scrolling="auto" frameborder="0" border="0" width="940" height="50" name="hiddenIframe"></iframe>
	<form name="FileDelFrm" method="post" target="hiddenIframe" action="Notice_Detail_FileEdit_InsUpDel.asp" encType="multipart/form-data" style="margin:0">
		<input type="hidden" name="isType" value="DEL">
		<input type="hidden" name="SEQ" value="<%=SEQ%>">		
		<input type="hidden" name="FILENAME_OLD" value="<%=db_FILENAME1%>">
	</form>
</DIV>

<!-- #include virtual="/Include/PopLayer.asp" -->
<!-- #include virtual="/Include/Bottom.asp" -->