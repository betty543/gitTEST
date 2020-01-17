<!-- #include virtual="/Include/Top.asp" -->
<!-- #include virtual="/Include/PopLayer.asp" -->
<%

	SS_LoginID = SESSION("SS_LoginID")
	SS_Login_Secgroup = SESSION("SS_Login_Secgroup")
	SS_Login_Grade = SESSION("SS_Login_Grade")

	whereCD1 = Trim(request("whereCD1"))
	whereCD2 = Trim(request("whereCD2"))
	whereCD3 = Trim(request("whereCD3"))
	whereCD4 = Trim(request("whereCD4"))
	QueryYN = Trim(request("QueryYN"))

	i = 0
	if QueryYN = "Y" then


		SQL = "SELECT	a.*, g.groupname	FROM TB_SMSADDR a, TB_SMSGROUP g	"
		SQL = SQL & "	WHERE	groupgb in ('2') and a.group_idx = g.idx"
		if whereCD1 <> "" then
			SQL = SQL & "	AND	group_idx= " & whereCD1
		end if
		if whereCD2 <> "" then	'소속
			SQL = SQL & "	AND		sosok_name like '%" & whereCD2 &"%'"
		end if
		if whereCD3 <> "" then	'전화번호
			SQL = SQL & "	AND		( cellphone like '%" & whereCD3 &"%' or gunphone like '%" & whereCD3 &"%')"
		end if
		if whereCD4 <> "" then	'전화번호
			SQL = SQL & "	AND		name like '%" & whereCD4 &"%'"
		end if
		SQL = SQL & "	union all "
		SQL = SQL & "	SELECT	a.*, g.groupname	FROM TB_SMSADDR a, TB_SMSGROUP g	"
		SQL = SQL & "	WHERE	groupgb in ('1') and a.group_idx = g.idx"
		if whereCD1 <> "" then
			SQL = SQL & "	AND	group_idx= " & whereCD1
		end if
		if whereCD2 <> "" then	'소속
			SQL = SQL & "	AND		sosok_name like '%" & whereCD2 &"%'"
		end if
		if whereCD3 <> "" then	'전화번호
			SQL = SQL & "	AND		( cellphone like '%" & whereCD3 &"%' or gunphone like '%" & whereCD3 &"%')"
		end if
		if whereCD4 <> "" then	'전화번호
			SQL = SQL & "	AND		name like '%" & whereCD4 &"%'"
		end if

		if SS_Login_Secgroup = "A" then
			SQL = SQL & "	and	g.incode = '" & SS_LoginID & "'"
		elseif SS_Login_Secgroup = "B" then
			'우리팀껏만 보이도록 한다
			SQL = SQL & "	and	g.incode in ( select userid from tb_userinfo where Grade = '" & SS_Login_Grade & "')"
		end if

		SQL = SQL & " order by name"
		i = 0

		SET Rs1 = db.execute(SQL)

	end if

%>


<table border="0" width="940" cellpadding="0" cellspacing="3" bgcolor="#EFECE5" align="center">
	<tr bgcolor="#FFFFFF">
		<td>
		
			<form method="post" name="inUpFrm" style="margin:0">
			<input type="hidden" name="QueryYN" value="">
			<table width="100%" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">

			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=80 align='center'>그룹</td>
					<td bgcolor="#FFFFFF" colspan=4>
					<select name="whereCD1" size="1" class="ComboFFFCE7" >
						<option value="">그룹선택</option>
<%					
							SQL = "	select * from TB_SMSGROUP where groupgb in ('2') and useyn = 'Y'" '공용
							SQL = SQL & " union all	select * from TB_SMSGROUP where groupgb in ('1')  and useyn = 'Y'"
							if SS_Login_Secgroup = "A" then
								SQL = SQL & "	and	incode = '" & SS_LoginID & "'"
							elseif SS_Login_Secgroup = "B" then
								'우리팀껏만 보이도록 한다
								SQL = SQL & "	and	incode in ( select userid from tb_userinfo where Grade = '" & SS_Login_Grade & "')"
							end if
							SET Rs = DB.execute(SQL)
							do until Rs.eof
									CODE = Rs("idx")
									if Rs("groupgb")="1" then
										CODENAME = db_GetUserName(Rs("incode"))&"-"&Rs("groupname")
									else
										CODENAME = "공통-"&Rs("groupname")
									end if
								%>

									<%=printSelect("" &CODENAME& "","" &CODE& "","" &whereCD1& "")%>
								<%
								Rs.movenext
							loop
%>
						</select> 				
						</td>

						<td bgcolor="#EEF6FF" class="TDCont" width=80 align='center'>소속</td>
						<td bgcolor="#FFFFFF" width=100 colspan=2><input type="text" name="whereCD2" value="<%=whereCD2%>" maxlength="20" size="14" onfocus="setFocusColor(this);" onblur="setOutColor(this);" align="absmiddle">
						</td>

						<td bgcolor="#EEF6FF" class="TDCont" width=80 align='center'>전화번호</td>
						<td bgcolor="#FFFFFF"><input type="text" name="whereCD3" value="<%=whereCD3%>" maxlength="10" size="14" onfocus="setFocusColor(this);" onblur="setOutColor(this);" align="absmiddle"></td>


						<td bgcolor="#EEF6FF" class="TDCont" width=80 align='center'>성명</td>
						<td bgcolor="#FFFFFF"><input type="text" name="whereCD4" value="<%=whereCD3%>" maxlength="10" size="14" onfocus="setFocusColor(this);" onblur="setOutColor(this);" align="absmiddle"></td>


			        <td colspan='2' rowspan="2" bgcolor="#FFFFFF" align="center">
			        	<img src="/Images/Btn/BtnSearch.gif" style="cursor:hand;" onClick="fn_Search();">
			        </td>
				</tr>

			</table>
			</form>
		</td>
	</tr>
</table>
<table width="940" border="0" cellpadding="0" cellspacing="0" align="center">
	<tr valign="top">
		<td width="500" height="750">
			<!--<iframe src="User_List.html" name="ListFrame" width="100%" height="100%" frameborder=0 marginheight=0 marginwidth=0 scrolling="no"></iframe>-->

			<table width="100%" border="0" cellspacing="0" cellpadding="0">
			    <tr height="5">
					<td align="left" class="TDCont" colspan="8"></td>
				</tr>
        	</table>
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
			    <tr height="30">
					<td align="left" bgcolor="#FFFFFF" class="TDCont" colspan="8">&nbsp;<img src="/Images/dot_01.gif" >&nbsp;<b><font color="#ff00ff"></font> 사용자목록</b></td>
				</tr>
        	</table>
        	<table width="100%" border="0" cellspacing="1" cellpadding="1" bgcolor="#CCCCCC">
        		<tr height="20" bgcolor="#EEF6FF" align="center">
        			<td>NO</td>
        			<td>그룹</td>
        			<td>소속</td>
        			<td>계급</td>
        			<td>성명</td>
        			<td>휴대폰</td>
        			<td>전체선택<br><input type="checkbox" name="chkALL" class="None" onClick="fn_select();"></td>
        		</tr>
        		<tr><td colspan="11" height="1" bgcolor="#FFFFFF"></td></tr>
<%

	if QueryYN = "Y" then

		do until Rs1.eof	
		
			i = i + 1
			idx = Rs1("idx")
			groupname = Rs1("groupname")
			sosok_name = Rs1("sosok_name")
			sclass = Rs1("class")
			sname = Rs1("name")
			cellphone = Rs1("cellphone")
			gunphone = Rs1("gunphone")
	%>
					<tr onmouseover="this.style.background='#FFFCE7'" bgcolor="#FFFFFF" height="30">
						<td align="center" width="30"><%=i%></td>
						<td align="center" width="70"><%=groupname%></td>
						<td align="center" nowrap ><%=sosok_name%></td>
						<td align="center"  width="70"><%=sclass%></td>
						<td align="center"  width="70"><%=sname%></td>
						<td align="center"  width="70"><%=cellphone%></td>
						<td align="center"  width="50"><input type="checkbox" name="chk" value="<%=idx%>" class="None"></td>
					</tr>
	<%
			Rs1.movenext
		loop

	end if
%>
<!--
				<tr onmouseover="this.style.background='#FFFCE7'" bgcolor="#FFFFFF" >
					<td align="center"><img src="/Images/Btn/icon_02.gif" style="cursor:hand;" align="absmiddle"></td>
					<td align="center">2</td>
					<td align="center">A그룹</td>
					<td align="center">소속1-소속2-소속3</td>
					<td align="center"> </td>
					<td align="center"> </td>
					<td align="center">김아무개</td>
					<td align="center">010-234-1234</td>
					<td align="center"><img src="/Images/Btn/icon_02.gif" style="cursor:hand;" align="absmiddle"></td>
				</tr>
-->
<% if i > 0 then 
	i = 0
%>
				<tr bgcolor="#FFFFFF" height="30">
					<td align="right" colspan=11 ><img src="/Images/Btn/BtnPlus.gif" style="cursor:hand;" align="absmiddle" title="선택대상 SMS전송목록으로 추가" onclick="fn_rightmove();"></td>
				</tr>
<% end if%>
        	</table>     

		</td>
		<td width="10" height="750" align='center' valign='center'>
			<!--<iframe src="User_List.html" name="ListFrame" width="100%" height="100%" frameborder=0 marginheight=0 marginwidth=0 scrolling="no"></iframe>-->

	
			<!--<table width="80%" border="0" cellspacing="0" cellpadding="0">
        		<tr><td height="22" colspan="2" class="FBlk"></td></tr>
        	</table>
        	<table width="80%" border="0" cellspacing="1" cellpadding="1" bgcolor="#CCCCCC">
        		<tr height="20" bgcolor="#ffffff" align="center">
        			<td></td>        		
        		</tr>
        	</table> -->   
		</td>
		<td width="450" height="750">
			<!--<iframe src="User_List.html" name="ListFrame" width="100%" height="100%" frameborder=0 marginheight=0 marginwidth=0 scrolling="no"></iframe>-->
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
			    <tr height="5">
					<td align="left" class="TDCont" colspan="8"></td>
				</tr>
        	</table>
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
			    <tr height="30">
					<td align="left" bgcolor="#FFFFFF" class="TDCont" colspan="8">&nbsp;<img src="/Images/dot_01.gif" >&nbsp;<b><font color="#ff00ff"></font> SMS전송대상</b></td>
				</tr>
        	</table>
        	<table width="100%" border="0" cellspacing="1" cellpadding="1" bgcolor="#CCCCCC">
        		<tr bgcolor="#EEF6FF" align="center" height="30">
        			<td width="50">전체선택<br><input type="checkbox" name="chkALL1" class="None" onClick="fn_select1();"></td>
        			<td>NO</td>
        			<td>소속</td>
        			<td>계급</td>
        			<td>성명</td>
        			<td>휴대폰</td>
        		</tr>
        		<tr><td colspan="7" height="1" bgcolor="#FFFFFF"></td></tr>
<%
			SQL = "	select c.idx,sosok_name, class, name, c.cellphone,c.gunphone	"
			SQL = SQL & "	from	temp_conference c, TB_SMSADDR a"
			SQL = SQL & "	where	addr_idx = a.idx and userid = '" & SS_LoginID & "' and datagb = '2' order by c.idx"

			set RS2 = db.execute(SQL)
			do until RS2.eof
				i = i + 1
				idx = RS2("idx")
				sosok_name = RS2("sosok_name")
				sclass = RS2("class")
				sname = RS2("name")
				cellphone = RS2("cellphone")
%>
        		<tr bgcolor="#fffff" align="center" height="30">
        			<td width="50"><input type="checkbox" name="chk1" value="<%=idx%>" class="None"></td>
        			<td><%=i%></td>
        			<td><%=sosok_name%></td>
        			<td><%=sclass%></td>
        			<td><%=sname%></td>
        			<td><%=cellphone%></td>
        		</tr>
<%
				RS2.movenext

			loop


%>

				<tr bgcolor="#FFFFFF" height="30">
					<td align="left" colspan=7>
					     <table width="100%" border="0" cellspacing="1" cellpadding="1" bgcolor="#FFFFFF">
						 <tr><td bgcolor="#FFFFFF" align='left'>
								<img src="/Images/Btn/BtnMinus.gif" style="cursor:hand;" align="absmiddle" title="선택대상 전송대상에서 제외" onclick="fn_leftmove();"></td><td align="center" ><img src="/Images/Btn/BtnSendSMS.GIF" style="cursor:hand;" align="absmiddle" onclick="fn_start();"></td>
						</tr></table>
					</td>
				</tr>


        	</table>       	


		</td>

	</tr>
</table>

<iframe src="about:blank" name="DBFrame" width="0" height="0" frameborder=0 marginheight=0 marginwidth=0 scrolling="no"></iframe>

<script>

	function fn_select() {
		var i, chked=0;

		for(i=0;i<document.all.chk.length;i++) {
				document.all.chk[i].checked=document.all.chkALL.checked;
		}

		if ( document.all.chk[0] == null)
		{
			document.all.chk.checked=document.all.chkALL.checked;
		}

	}

	function fn_select1() {
		var i, chked=0;

		try {
			for(i=0;i<document.all.chk1.length;i++) {
					document.all.chk1[i].checked=document.all.chkALL1.checked;
			}

			if ( document.all.chk1[0] == null)
			{
				document.all.chk1.checked=document.all.chkALL1.checked;
			}
		}
		catch(e){}
	}
	function fn_rightmove() {
		var i, chked=0;
		var j = 0;
		for(i=0;i<document.all.chk.length;i++) {
				if ( document.all.chk[i].checked )
				{
					j = j + 1;
					DBFrame.location= "/menu05/submenu0502/sms_move.asp?FRM=right&idx="+document.all.chk[i].value;
				}
		}

		if ( document.all.chk[0] == null && document.all.chk.checked  )
		{
			j = j + 1;
			DBFrame.location= "/menu05/submenu0502/sms_move.asp?FRM=right&idx="+document.all.chk.value;
		}

		if ( j > 0 )
		{
			document.inUpFrm.submit();
		}

	}
	function fn_leftmove() {

		var i, chked=0;
		var j = 0;
		for(i=0;i<document.all.chk1.length;i++) {
				if ( document.all.chk1[i].checked )
				{
					j = j + 1;
					DBFrame.location= "/menu05/submenu0502/sms_move.asp?FRM=left&idx="+document.all.chk1[i].value;
				}
		}

		if ( document.all.chk1[0] == null && document.all.chk1.checked  )
		{
			j = j + 1;
			DBFrame.location= "/menu05/submenu0502/sms_move.asp?FRM=left&idx="+document.all.chk1.value;
		}

		if ( j > 0 )
		{
			document.inUpFrm.QueryYN.value = "Y";
			document.inUpFrm.submit();
		}

	}

	function fn_Search() {
		document.inUpFrm.QueryYN.value = "Y";
		document.inUpFrm.submit();
	}


	function fn_start() {

		ShowPOPLayer("sms.asp",'620','430');		
//				sms = window.open("sms.asp","sms","toolbar=no,status=yes,location=no,width=620,height=500,top=0,left=0,scrollbars=yes,resizable=no");
			//	sms.focus();
	}

</script>

<!-- #include virtual="/Include/Bottom.asp" -->