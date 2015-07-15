<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Index.aspx.vb" Inherits="Index" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>BVI Directory</title>
    <link href="css/BVI.css" rel="stylesheet" type="text/css" />
</head>
<body>
    <form id="form1" runat="server">
    <div>
    
    </div>
			<input type="hidden" name="hdAction"/> <input type="hidden" name="hdSortField"/> <input type="hidden" name="hdEnrollerID"/>
			<input id="hdFilterShowHideToggle" type="hidden" value="0" name="hdFilterShowHideToggle"/>
			<table class="PrimaryTbl" style="POSITION: absolute; TOP: 14px; LEFT: 14px" cellspacing="0"
				cellpadding="0" width="650" border="0">
				<tr style="DISPLAY: none">
					<td width="100"></td>
					<td width="300"></td>
				</tr>
				<tr>
					<td class="PrimaryTblTitle" colspan="2">BVI Directory</td>
				</tr>
				<tr style="vertical-align:top">
					<td class="HdPhoneL">HBG/REM: (800) 810-2200</td>
					<td class="HdPhoneR">LA: (800) 499-9190</td>
				</tr>								
				<tr>
					<td class="CellSeparator" colspan="2"></td>
				</tr>
				<tr>
					<td style="height:20px; vertical-align:top" colspan="2"><asp:Literal ID="litDBHost" runat="server" EnableViewState="False"></asp:Literal></td>				
				</tr>
				

				<tr>
					<td colspan="2"><asp:literal id="litDG" runat="server" enableviewstate="False"></asp:literal></td>
				</tr>
			</table>
			<asp:literal id="litMsg" runat="server" enableviewstate="False"></asp:literal><asp:literal id="litFilterHiddens" runat="server" enableviewstate="False"></asp:literal>
			<script type="text/javascript"> 
			
			function fnOnLoad()  {
				<asp:literal id="litJumpToAnchor" runat="server" enableviewstate="False"></asp:literal>
			}			

			function Sort(vField) {
				form1.hdAction.value = "Sort"
				form1.hdSortField.value = vField
				form1.submit()
			}
	
			function ToggleShowFilter()  {
				form1.hdFilterShowHideToggle.value = 1
				form1.hdAction.value = "ApplyFilter"
				form1.submit()
			}				
	
			function ApplyFilter()
			{		
				form1.hdAction.value = "ApplyFilter"
				form1.submit()				
			}						 						
						

			function ExistingRecord(vEnrollerID)
			{
				form1.hdAction.value = "ExistingRecord"
				form1.hdEnrollerID.value = vEnrollerID
				form1.submit()
			}
			function Permissions(vEnrollerID)
			{
				form1.hdAction.value = "Permissions"
				form1.hdEnrollerID.value = vEnrollerID
				form1.submit()
			}	
			
			function CallWorklist(vEnrollerID)
			{
				form1.hdAction.value = "CallWorklist"
				form1.hdEnrollerID.value = vEnrollerID
				form1.submit()
			}						
			
			function DailyHoursWorklist(vEnrollerID)
			{
				form1.hdAction.value = "DailyHoursWorklist"
				form1.hdEnrollerID.value = vEnrollerID
				form1.submit()
			}			
			
			function NewRecord() {
				form1.hdAction.value = "NewRecord"
				form1.submit()
			}				

			function SubmitOnEnterKey(e) {
				var keypressevent = e ? e : window.event
				if (keypressevent.keyCode == 13) {	
					form1.hdAction.value = "ApplyFilter"						 	
					form1.submit()
				}			
			}		
			
			function Delete(vEnrollerID, vUserName)  {
				vUserName = vUserName.replace("~", "'");
				var OKToDelete = confirm("Are you sure you wish to delete " + vUserName + "?");
				if (OKToDelete == true) {
					form1.hdAction.value = "Delete"
					form1.hdEnrollerID.value = vEnrollerID
					form1.submit()		
				}		
			}				
																												
			</script>    
    </form>
</body>
</html>
