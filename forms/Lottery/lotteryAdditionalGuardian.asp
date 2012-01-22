<%
set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()
session.Contents("strInstructions") = "Fill out the form below to add an additional parent/guardian. " & _
									  "To save this information you must click the 'Save Guardian Info' " & _
									  "button at the bottom of the page. * denotes required information."
									  
%>
				<tr>
					<td class="svplain10">
						<table cellspacing="0" cellpadding="4" bordercolor="e6e6e6" border="1" ID="Table3">
							<tr>
								<td class="svplain10">
									<b>Instructions:</b><br>
									<% = session.Contents("strInstructions") %>
									<BR>
									<% = session.Contents("strError") %>
								</td>
							</tr>
						</table>
						<br>
					</td>
				</tr>
				<tr>
					<input type=hidden name="bolAdditionalGuardian" value="true">
					<td class="NavyHeader">
						&nbsp;<B>Parent/Guardian Information</B> &nbsp;&nbsp;&nbsp;&nbsp;
					</td>
				</tr>
				<tr>
					<td>
						<table ID="Table8">
							<tr>
								<Td colspan="6">
									<font class="svplain10"><b><i>Additional Parent/Guardian</i></b> </font>
								</Td>
							</tr>
							<tr>
								<td class="gray">
									&nbsp; Last Name*
								</td>
								<td class="gray">
									&nbsp;First Name*
								</td>
								<td class="gray">
									&nbsp;MI
								</td>
								<td class="gray">
									&nbsp;Email Address&nbsp;
								</td>
							</tr>
							<tr>
								<td>
									<input type="text" name="szLast_Name" value="<%= szLast_Name%>" maxlength="50" size="17"   ID="Text15">
								</td>
								<td>
									<input type="text" name="szFirst_Name" value="<%= szFirst_Name%>" maxlength="50" size="15"   ID="Text20">
								</td>
								<td>
									<input type="text" name="sMid_Initial" value="<%= sMid_Initial%>" maxlength="1" size="2"   ID="Text21">
								</td>
								<td class="svplain10">
									<input type=text name="szEmail" size=25 value="<%= szEmail%>" maxlength=128 ID="Text4">
								</td>
							</tr>
						</table>
						<table ID="Table9">
							<tr>
								<Td class="gray">
									&nbsp;Employer
								</Td>
								<td class="gray">
									&nbsp;Active Military&nbsp;
								</td>
								<td class="gray">
									&nbsp;Rank&nbsp;
								</td>
								<td class="gray">
									&nbsp;Pager
								</td>
							</tr>
							<tr>
								<td>
									<input type="text" name="szEmployer" value="<%= szEmployer%>" maxlength="128" size="30">
								</td>
								<td>
									<select name="bolActive_Military" ID="Select4">
										<option value="">- - - - - - - - - - -
											<%
											Response.Write oFunc.MakeList("TRUE,FALSE","Yes,No", oFunc.TFText(bolActive_Military2))
										%>
									</select>
								</td>
								<td>
									<input type="text" name="szRank" value="<%= szRank%>" maxlength="20" size="4"   ID="Text23">
								</td>
								<td>
									<input type="text" name="szPager" value="<%= szPager%>" maxlength="15" size="15"   ID="Text24">
								</td>
							</tr>
						</table>
						<table ID="Table10">
							<tr>
								<td class="gray">
									&nbsp;Home Phone&nbsp;
								</td>
								<td class="gray">
									&nbsp;Business Phone&nbsp;
								</td>
								<td class="gray">
									&nbsp;Ext.
								</td>
								<td class="gray">
									&nbsp;Cell Phone
								</td>
							</tr>
							<tr>
								<td align="center">
									<input type="text" name="szHome_Phone" value="<%= szHome_Phone%>" maxlength="15" size="15"   ID="Text3">
								</td>
								<td align="center">
									<input type="text" name="szBusiness_Phone" value="<%= szBusiness_Phone%>" maxlength="15" size="15"   ID="Text25">
								</td>
								<td>
									<input type="text" name="intPhone_Ext" value="<%= intPhone_Ext%>" maxlength="4" size="4"   ID="Text26">
								</td>
								<td>
									<input type="text" name="szCell_Phone" value="<%= szCell_Phone%>" maxlength="15" size="15"   ID="Text27">
								</td>
							</tr>
						</table>
						<table ID="Table11">
							<tr>
								<td class="gray">
									&nbsp;Address (if different)
								</td>
								<td class="gray">
									&nbsp;City
								</td>
								<td class="gray">
									&nbsp;State
								</td>
								<Td class="gray">
									&nbsp;Country
								</Td>
								<Td class="gray">
									&nbsp;Zip
								</Td>
							</tr>
							<tr>
								<td>
									<input type="text" name="szAddress" value="<%= szAddress%>" maxlength="256" size="30"   ID="Text28">
								</td>
								<td>
									<input type="text" name="szCity" value="<%= szCity%>" maxlength="50" size="10"   ID="Text29">
								</td>
								<td>
									<select name="szState" ID="Select5">
										<option value="">
											<%
								'Create State select list									
								sql = "select strValue, strText " & _
										"from Common_Lists " & _
										"where intList_ID=3 order by strValue "
								' Set Alaska as default state
								if szState = "" then szState = "AK"
								response.Write oFunc.MakeListSQL(sql,"strValue","strText",szState)
								
								if szCountry = "" then szCountry = "USA"
							%>
									</select>
								</td>
								<td>
									<input type="text" name="szCountry" value="<%= szCountry%>" maxlength="25" size="7"   ID="Text30">
								</td>
								<td>
									<input type="text" name="szZip_Code" value="<%= szZip_Code%>" maxlength="10" size="5"   ID="Text31">
								</td>
							</tr>
						</table>
					</td>
				</tr>
<%
session.Contents("strError") = ""
call oFunc.CloseCN()
set oFunc = nothing
%>			