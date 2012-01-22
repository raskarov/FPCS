<%@ Language=VBScript %>
<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Name:		addSponsorTeacher.asp
'Purpose:	Form for adding/inserting a sponsor teacher that will aid a 
'			parent in creating students ciriculum.			
'Date:		9-04-01
'Author:	Scott Bacon (ThreeShapes.com LLC)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

call vbfHeader("Parent Teacher Contract","")
%>
<form action="classInsert.asp" method=Post name=main>
<input type=hidden name=resourceCount value="<%=intCount%>">
<input type=hidden name=changed value="">
<input type=hidden name=bolValidated value="<% = request("bolValidated") %>">
<input type=hidden name=intInstructor_ID value="<% = intInstructor_ID %>">
<table width=100%>
	<tr>	
		<Td class=yellowHeader>
				&nbsp;<b>Parent Teacher Contract for <% = strClassTitle %></b> <i> 			
		</td>
	</tr>
	<tr>
		<td bgcolor=f7f7f7>
			<table>
				<tr>	
					<Td colspan=2>
						<font class=svplain11>
							<b><i>Parties Involved</I></B> 
						</font>
						<font class=svplain>
						</font>
					</td>
				</tr>
				<tr>
					<td class=gray>
							&nbsp;Select a Parent/Guardian
					</td>
					<td class=gray>
							&nbsp;Select a Teacher
					</td>													
				</tr>
				<tr>
					<td>
						<select name="intGuardian_id">
						<%
							dim sqlGuardian	
							sqlGuardian = "select g.intGuardian_id, g.szFirst_Name + ' ' + g.szLast_name as Name " & _
								  "from tblGuardian g, tascStudent_Guardian sg " & _
								  "where sg.intStudent_id = " & Request("intStudent_id") & _
								  " and g.intGuardian_id = sg.intGuardian_id order by Name "
								  
							Response.Write vbfMakeListSQL(sqlGuardian,"intGuardian_id","Name",intGuardian_id)
						%>
						</select>		
										
					</td>
					<td>
						<select name="intCourse_ID" onChange="jfChanged();">
							<option value="">
						<%	
							vbfPrint "end:"
							dim sqlTeachers
							sqlTeachers = "select intInstructor_id, szFirst_Name + ' ' + szLast_name as Name " & _
								         "from tblInstructor order by Name"
							Response.Write vbfMakeListSQL(sqlTeachers,"intInstructor_id","Name",intInstructor_id)
						%>
						</select>	
					</td>
					<!--
					<td>
						<input type=text name="szSubject" value="<% = szSubject%>" maxlength=64 size=20 onChange="jfChanged();">
					</td> -->
					<% if bolLocation = true then %>	
					<td>
						<input type=text name="szLocation" value="<% = szLocation%>" maxlength=50 size=17 onChange="jfChanged();">
					</td>	
					<%end if%>												
				</tr>
			</table>
			<table>
				<tr>	
					<Td colspan=2>
						<font class=svplain11>
							<b><i>Class Information</I></B> 
						</font>
						<font class=svplain>
						</font>
					</td>
				</tr>
				<tr>
					<td class=gray>
							&nbsp;Name of Class
					</td>
					<td class=gray>
							&nbsp;Select a Course Category
					</td>
					<!--<td class=gray>
						&nbsp;Subject
					</td>-->
					<% if bolLocation = true then%>	
					<td class=gray>
						&nbsp;Location
					</td>			
					<%end if%>														
				</tr>
				<tr>
					<td>
						<input type=text name="szClass_Name" value="<% = szClass_Name%>" maxlength=64 size=20 onChange="jfChanged();">
					</td>
					<td>
						<select name="intCourse_ID" onChange="jfChanged();">
							<option value="">
						<%
							dim sqlCourses
							sqlCourses = "select intCourse_ID, strCourse_Name " & _
								         "from trefCourse_Categories order by strCourse_Name"
							Response.Write vbfMakeListSQL(sqlCourses,"","",intCourse_ID)
						%>
						</select>	
					</td>
					<!--
					<td>
						<input type=text name="szSubject" value="<% = szSubject%>" maxlength=64 size=20 onChange="jfChanged();">
					</td> -->
					<% if bolLocation = true then %>	
					<td>
						<input type=text name="szLocation" value="<% = szLocation%>" maxlength=50 size=17 onChange="jfChanged();">
					</td>	
					<%end if%>												
				</tr>
			</table>
			<table>
				<tr>
					<td class=gray colspan=3>
						&nbsp;Registration Deadline
					</td>
					<% if bolStudentNum = true then %>
					<td class=gray>
						&nbsp;Min # Students
					</td>	
					<td class=gray>
						&nbsp;Max # Students
					</td>	
					<% end if %>
					<td class=gray>
						&nbsp;Grade Level
					</td>																	
				</tr>
				<tr>
					<td>
						<select name="month" onChange="jfChanged();">
							<% 
							dim sqlMonth
							sqlMonth = "select strValue,strText from common_lists where intList_ID = 1 order by intOrder"
							Response.Write vbfMakeListSQL(sqlMonth,"","",month)								
							%>
						</select>
					</td>		
					<td>
						<select name="day" onChange="jfChanged();">
							<% 
							dim sqlDay
							sqlDay = "select strValue,strText from common_lists where intList_ID = 2 order by intOrder"
							Response.Write vbfMakeListSQL(sqlDay,"","",day)								
							%>
						</select>
					</td>											
					<td>
						<select name="year" onChange="jfChanged();">
							<% = vbfMakeYearList(2,0,year) %>
						</select>
					</td>	
					<% if bolStudentNum = true then %>
					<td align=center>
						<input type=text name="intMin_Students" value="<% = intMin_Students%>" maxlength=3 size=4 onChange="jfChanged();">
					</td>	
					<td align=center>
						<input type=text name="intMax_Students" value="<% = intMax_Students%>" maxlength=3 size=4 onChange="jfChanged();">
					</td>	
					<% end if %>
					<td align=center>
						<select name="sGrade_Level" onChange="jfChanged();">
							<option value="">
							<% 
							dim strGradeList
							strGradeList = "K,1,2,3,4,5,6,7,8,9,10,11,12"
							Response.Write vbfMakeList(strGradeList,strGradeList,sGrade_Level)								
							%>
						</select>
					</td>										
				</tr>
			</table>
			<% if bolClassMeets = true then %>
			<table>				
				<tr>
					<td class=gray colspan=3>
							&nbsp;Class Start Date
					</td>
					<td class=gray colspan=3>
						&nbsp;Class End Date
					</td>	
					<td class=gray>
						&nbsp;Meets Every
					</td>																
				</tr>
				<tr>
					<td valign=top>
						<select name="monthStart" onChange="jfChanged();">
							<% 
							Response.Write vbfMakeListSQL(sqlMonth,"","",monthStart)								
							%>
						</select>
					</td>		
					<td valign=top>
						<select name="dayStart" onChange="jfChanged();">
							<% 
							Response.Write vbfMakeListSQL(sqlDay,"","",dayStart)								
							%>
						</select>
					</td>											
					<td valign=top>
						<select name="yearStart" onChange="jfChanged();">
							<% = vbfMakeYearList(2,0,yearStart) %>
						</select>		
					</td>				
					<td valign=top>
						<select name="monthEnd" onChange="jfChanged();">
							<% 
							Response.Write vbfMakeListSQL(sqlMonth,"","","")								
							%>
						</select>
					</td>		
					<td valign=top>
						<select name="dayEnd" onChange="jfChanged();">
							<% 
							Response.Write vbfMakeListSQL(sqlDay,"","","")								
							%>
						</select>
					</td>											
					<td valign=top>
						<select name="yearEnd" onChange="jfChanged();">
							<% = vbfMakeYearList(2,0,yearEnd) %>
						</select>		
					</td>				
					<td>
						<select name="szDays_Meet_On" onChange="jfChanged();" multiple size=2>
							<% 
							dim sqlDays
							sqlDays = "select strValue,strText from common_lists where intList_ID = 4 order by intOrder"
							Response.Write vbfMakeListSQL(sqlDays,"","",szDays_Meet_On)								
							%>
						</select>
					</td>
				</tr>
			</table>		
			<% end if %>	
			<table>				
				<tr>
					<td class=gray colspan=4>
							&nbsp;Class Start Time
					</td>
					<td class=gray>
							&nbsp;
					</td>
					<td class=gray colspan=4>
						&nbsp;Class End Time
					</td>		
					<td class=gray colspan=4>
						&nbsp;Schedule Comments
					</td>													
				</tr>
				<tr>
					<td valign=top>
						<select name="hourStart" onChange="jfChanged();">
							<% 
							dim strHour
							strHour = "1,2,3,4,5,6,7,8,9,10,11,12"
							Response.Write vbfMakeList(strHour,strHour,hourStart)								
							%>
						</select>
					</td>	
					<td valign=top>
						:
					</td>	
					<td valign=top>
						<select name="minuteStart" onChange="jfChanged();">
							<% 
							dim strMinute
							dim str0
							strMinute = "00,01"
							for i = 2 to 60
								if i < 10 then str0 = "0"
								strMinute = strMinute & "," &  str0 & i
								str0 = ""
							next 
							Response.Write vbfMakeList(strMinute,strMinute,minuteStart)								
							%>
						</select>
					</td>											
					<td valign=top>
						<select name="amPmStart" onChange="jfChanged();">
							<% 
							dim strAmPm
							strAmPm = "AM,PM"
							Response.Write vbfMakeList(strAmPm,strAmPm,amPmStart)								
							%>
						</select>		
					</td>			
					<td>
							&nbsp;
					</td>	
					<td valign=top>
						<select name="hourEnd" onChange="jfChanged();">
							<% 
							Response.Write vbfMakeList(strHour,strHour,hourEnd)								
							%>
						</select>
					</td>		
					<td valign=top>
						:
					</td>	
					<td valign=top>
						<select name="minuteEnd" onChange="jfChanged();">
							<% 
							Response.Write vbfMakeList(strMinute,strMinute,minuteEnd)								
							%>
						</select>
					</td>											
					<td valign=top>
						<select name="amPmEnd" onChange="jfChanged();">
							<% 
							Response.Write vbfMakeList(strAmPm,strAmPm,amPmEnd)								
							%>
						</select>		
					</td>	
					<td align=center>
						<textarea cols=20 rows=2 name="szSchedule_Comments" wrap=virtual><% = szSchedule_Comments%></textarea>						
					</td>	
				</tr>
			</table>			
			<table>
				<tr>	
					<Td colspan=2>
						<font class=svplain11>
							<b><i>Class Costs</I></B> 
						</font>
						<font class=svplain11>
							 (Resources Required)
						</font>
					</td>
				</tr>
				<tr>
					<td class=gray>
							&nbsp;<i>Itemized Costs for an <u>Individual</u> Student Only.</i>. 
							<%  if request("viewing") = "" and request("plain") = "" then %> 
							Add a needed item.
							<% if bolStudentNum = "" then strParam = "jfAddMaterials();" %>
							<input type=button name="addItem" value="add" id=btSmallGray onClick="jfAddResource('<%=strParam%>');">
							<% end if %>
					</td>										
				</tr>
				<tr>
					<td id="resources">
						<% = strMaterials %>
					</td>
				</tr>				
			</table>
			<table>
				<% if bolStudentNum = true then %>
				<tr>	
					<Td colspan=2>
						<font class=svplain11>
							<b><i>Class Costs</I></B> 
						</font>
						<font class=svplain11>
							 (Teachers Time)
						</font>
					</td>
				</tr>
				<tr>
					<td class=gray>
						<input type=text name="decHours_Student" value="<% = decHours_Student %>" size=4 maxlength=3 onChange="jfChanged();">
					</td>		
					<td class=gray>
						&nbsp;Number of teacher hours with student.
					</td>							
				</tr>		
				<tr>
					<td class=gray>
						<input type=text name="decHours_Planning" value="<% = decHours_Planning %>" size=4 maxlength=3 onChange="jfChanged()">
					</td>		
					<td class=gray>
						&nbsp;Number of hours for teacher planning.
					</td>							
				</tr>	
				<tr>
					<td class=gray align=center>
						&nbsp;=&nbsp;
					</td>
					<td class=gray>
						<input type=button value="calculate totals" onClick="jfAddHRS();" id=btSmallGray>
					</td>							
				</tr>	
				<tr>
					<td class=gray>
						<input type=text name="totalHours"  size=4 maxlength=4 disabled>
					</td>		
					<td class=gray>
						&nbsp;<B>Total teacher hours.</b>
					</td>							
				</tr>	
				<tr>
					<td class=gray>
						<input type=text name="intMin_Charged" value="" size=4 maxlength=10 disabled>
					</td>		
					<td class=gray>
						&nbsp;Minimum number of hours to be charged to each student.
					</td>							
				</tr>	
				<tr>
					<td class=gray>
						<input type=text name="intMax_Charged" value="" size=4 maxlength=10 disabled>
					</td>		
					<td class=gray>
						&nbsp;Maximum number of hours to be charged to each student.
					</td>							
				</tr>	
				<tr>
					<td class=gray>
						<input type=text name="curRate" value="$<% = curPay_Rate%>" size=4 maxlength=10 disabled>
					</td>		
					<td class=gray>
						&nbsp;Teachers hourly rate.
					</td>							
				</tr>
				<tr>
					<td class=gray>
						<input type=text name="intMinTeacherCost" value="" size=4 maxlength=10 disabled>
					</td>		
					<td class=gray>
						&nbsp;Minimum total teacher cost per student.
					</td>							
				</tr>	
				<tr>
					<td class=gray>
						<input type=text name="intMaxTeacherCost" value="" size=4 maxlength=10 disabled>
					</td>		
					<td class=gray>
						&nbsp;Maximum total teacher cost per student.
					</td>							
				</tr>
				<% end if %>
				<tr>
					<td class=gray>
						<input type=text name="intMiscCost" value="" size=4 maxlength=10 disabled>
					</td>		
					<td class=gray>
						&nbsp;Total miscellaneous costs per student.
					</td>							
				</tr>		
				<% if bolStudentNum = true then %>
				<tr>
					<td class=gray>
						<input type=text name="intMinTotalCost" value="" size=4 maxlength=10 disabled>
					</td>		
					<td class=gray>
						&nbsp;<B>Minimum total deduction per student account.</b>
					</td>							
				</tr>	
				<tr>
					<td class=gray>
						<input type=text name="intMaxTotalCost" value="" size=4 maxlength=10 disabled>
					</td>		
					<td class=gray>
						&nbsp;<B>Maximum total deduction per student account.</b>
					</td>							
				</tr>		
				<% end if %>	
			</table>
		</td>
	</tr>
</table>
<% if request("viewing") = "" and request("plain") = "" then %>
<input type=submit value="ADD CLASS >" id="btSmallGray">
<% end if %>
</form>
<script language=javascript>	
	function jfAddMaterials(){
		var intRcount = parseInt(document.main.resourceCount.value);
		var curMiscTotal = 0;
		var curItemTotal;
		for(i=0;i<intRcount;i++){
			curItemTotal = document.all.item("itemTotal" + i).value;
			curItemTotal = curItemTotal.replace("$","");
			curMiscTotal += parseFloat(curItemTotal);
		}
		document.main.intMiscCost.value = "$" + curMiscTotal;
	}
	
	function jfAddHRS(){
		var intStudentHours = document.main.decHours_Student.value;
		var intHRS_Planning = document.main.decHours_Planning.value;
		var intMinStudent = document.main.intMin_Students.value;
		var intMaxStudent = document.main.intMax_Students.value;
		var intTotalHours = parseInt(intStudentHours) + parseInt(intHRS_Planning);
		var intRate = document.main.curRate.value;
		var intRcount = parseInt(document.main.resourceCount.value);
		var curMiscTotal = 0;
		var curItemTotal;
		if (intStudentHours == "" || intHRS_Planning == "" || intMinStudent == ""
				|| intMaxStudent == "") {
			var strMessage;
			strMessage = "To Calculate totals you must provide a value for \n";
			strMessage += "'Min # of Students'\n'Max # of Students'\n";
			strMessage += "'Number of teacher hours with student'\n";
			strMessage += "'Number of hours for teacher planning'.";
			alert(strMessage);
			return;
		}
		
		intRate = intRate.replace("$","");
		document.main.totalHours.value = intTotalHours;
		document.main.intMax_Charged.value = intTotalHours / parseInt(intMinStudent);
		document.main.intMin_Charged.value = intTotalHours / parseInt(intMaxStudent);
		document.main.intMaxTeacherCost.value  =  "$" + (parseFloat(document.main.intMax_Charged.value) * parseFloat(intRate));
		document.main.intMinTeacherCost.value  =  "$" + (parseFloat(document.main.intMin_Charged.value) * parseFloat(intRate));	
		jfAddMaterials();
		
		var max1 = document.main.intMaxTeacherCost.value;
		var min1 = document.main.intMinTeacherCost.value
		var materials = document.main.intMiscCost.value;
		
		max1 = max1.replace("$","");
		min1 = min1.replace("$","");
		materials = materials.replace("$","");
		document.main.intMinTotalCost.value  =  "$" + (parseFloat(min1) + parseFloat(materials));
		document.main.intMaxTotalCost.value  =  "$" + (parseFloat(max1) + parseFloat(materials));	
	}
	<% if intClass_id <> ""  then 	
			for i = 0 to intCount -1 %>
				jfCalcTotal('<% = i %>');
		<%next%>
		jfAddHRS();
	<% end if %>
</script>
</BODY>
</HTML>
<% 
 call vbfCloseCN
%>