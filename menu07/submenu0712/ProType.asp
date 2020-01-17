<!-- #include virtual="/Include/Top.asp" -->
<%
	Aclass = Request("Aclass")				'1차분류
	Bclass = Request("Bclass")				'2차분류
	Cclass = Request("Cclass")				'3차분류
	Dclass = Request("Dclass")				'3차분류
	'Response.write Aclass

%>
<table width="940" border="0" cellspacing="0" cellpadding="0" align="center">
	<tr><td colspan="9" height="10"></td></tr>
    <tr align="center">
        <td width="300" class="FBlkB" valign="top">
        	<!-- 1차 분류 시작 -->
			<iframe src="ProType_1.asp?Aclass=<%=Aclass%>" frameborder=0 marginheight=0 marginwidth=0 width="100%" scrolling="no" name="Pro1fr"> </iframe>
			<!-- 1차 분류 끝 -->
		</td>
        <td width="5" rowspan="4"></td>
        <td width="300" class="FBlkB" valign="top">
        	<!-- 2차 분류 시작 -->
			<iframe src="ProType_2.asp?Aclass=<%=Aclass%>" frameborder=0 marginheight=0 marginwidth=0 width="100%" scrolling="no" name="Pro2fr"> </iframe>
			<!-- 2차 분류 끝 -->
		</td>
        <td width="5" rowspan="4"></td>
        <td width="300" class="FBlkB" valign="top">
        	<!-- 3차 분류 시작 -->
			<iframe src="ProType_3.asp?Aclass=<%=Aclass%>&Bclass=<%=Bclass%>" frameborder=0 marginheight=0 marginwidth=0 width="100%" scrolling="no" name="Pro3fr"> </iframe>
			<!-- 3차 분류 끝 -->
        </td>
        <td width="5" rowspan="4"></td>
        <td width="300" class="FBlkB" valign="top">
        	<!-- 4차 분류 시작 -->
			<iframe src="ProType_4.asp?Aclass=<%=Aclass%>&Bclass=<%=Bclass%>&Cclass=<%=Cclass%>" frameborder=0 marginheight=0 marginwidth=0 width="100%" scrolling="no" name="Pro4fr"> </iframe>
			<!-- 4차 분류 끝 -->
        </td>
    </tr> 
</table>

<!--제품분류코드입력 시작 -->
<span id="ProTypeDetail" style="display:none;">
<table width="940" border="0" cellspacing="0" cellpadding="0" align="center">
	<tr><td>&nbsp;</td></tr>
	<tr align="left">
		<td width="100%">
			<table border="0" cellspacing="0" width="100%" bordercolor="#CCCCCC" bordercolordark="white" bordercolorlight="#CCCCCC">
				<tr>
					<td width="100%" height="50" class="TDL35px">
						<iframe frameborder=0 marginheight=0 marginwidth=0 width="100%" scrolling="no" name="ProTypeDetailFrame"></iframe>
					</td>
				</tr>
			</table>
		</td>
		<td width="100%">&nbsp;</td>
	</tr>
</table>
</span>
<!--제품분류코드입력 끝 -->

<script>
	function TempleteADD(URL, layerOF){
		ProTypeDetailFrame.location.href=URL
		if (layerOF == "ON") {
			LayerON('ProTypeDetail');
		} else if (layerOF == "OFF") {
			LayerOFF('','ProTypeDetail');
		}
	}
</script>

<!-- #include virtual="/Include/Bottom.asp" -->