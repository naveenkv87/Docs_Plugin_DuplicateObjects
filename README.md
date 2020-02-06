# TexasInstruments.Plugins.DuplicatePublications
Plugin developed to Duplicate a publication along with its objects.
1.	Adding button to ribbon bar
Paste below code in file “FolderButtionbar.xml” which will be in location “C:\InfoShare\Web\Author\ASP\XSL”

<BUTTON CHECKACCESS="Y">
    <!-- Duplicate Publications -->
    <CARDTYPE>VDOCTYPEPUBLICATION</CARDTYPE>
    <INPUT type="button" NAME="Duplicate Publication" onClick="DuplicatePublication()" VALUE="Duplicate Publication" class="button" ICON="UIFramework/export.32x32.png" SHOWTEXT="Y"/>
    <SCRIPT language="JAVASCRIPT">
      <![CDATA[
        function DuplicatePublication()
        {
		  var iPos;
          var FolderName;
          var FolderId;
          var FolderPath;
          var FolderUsergroup;
          var FolderType;
          var SelectedCards = parent.DataFrame.aDragLogicalArray;
          var SelectedCard;

          SelectedCard ='';
		  
		  FolderId = parent.DataFrame.document.FolderContentData.FOLDER.value;
          FolderName = parent.DataFrame.document.FolderContentData.CurrentFolderName.value;
          FolderPath=parent.DataFrame.document.FolderContentData.PathNames.value;
          IPos=FolderPath.lastIndexOf("|");
          if (IPos>0) FolderPath=FolderPath.substring(IPos+1);
		  
		  FolderUsergroup = parent.DataFrame.document.FolderContentData.FolderUsrGrp.value;
          FolderType =  'VDOCTYPENONE'; 
		  
		  for( i = 0, j=0, lArray=parent.DataFrame.aDragLogicalArray; i < parent.DataFrame.aDragLogicalArray.length ; i++)
          {
            if (lArray[i] != undefined)
            {
              IPos = lArray[i].lastIndexOf("|");
              if (IPos > 0) SelectedCard= lArray[i].substring(IPos+1)
            }
          }

          var WindowStr = "dependent=yes,toolbar=no,width=450,height=250,resizable=yes,scrollbars=auto,menubar=no,status=yes";
		  
		  if (SelectedCard.length==0)
          {
            alert("Please select an object");
          }
		  else
          {            
			FolderAssistForObjects('Duplicate Publications', FolderId, FolderName, FolderUsergroup, FolderType, FolderPath, 'PerformDuplicatePublication()', '');
          }
        }
    ]]>
    </SCRIPT>
  </BUTTON>
 
2.	JavaScript function to call new ASP file

Add below code in file “FolderButton.js” which will be in location “C:\InfoShare\Web\Author\ASP\FolderAssist”

function PerformDuplicatePublication()
{
// This function is executed when FolderAssist OK button is pressed.
//   - used for "moving objects to another location

  	var NewFolderId = parent.TreeFrame.GetSelected();
	var NewFolderName = parent.TreeFrame.GetCurrentName();
	var OldFolderId = parent.TreeFrmFolderId;
	var OldFolderName = parent.TreeFrmFolderName;
  
// Make String of array of selected cards. String is not part of QueryString, but will be submitted in "TreeFrm"-form.
  
   var lNoOfDocsSelected =window.top.opener.parent.DataFrame.aDragLogicalArray.length;
	var SelectedCards = window.top.opener.parent.DataFrame.aDragLogicalArray;

	var CardIds;
	var IPos;
	var Total = 0;
	CardIds='';
	if (lNoOfDocsSelected>0)
	{
		for( i = 0, j=0, lArray=SelectedCards; i < lArray.length ; i++)
		{
			if (lArray[i] != undefined)
			{
				IPos = lArray[i].lastIndexOf("|");
				if (IPos > -1)
				{
					if (Total++==0)
						CardIds= lArray[i].substring(IPos+1)
					else
						CardIds= CardIds.toString() + ', ' + lArray[i].substring(IPos+1);
				}
			}
		}
		parent.TreeFrmCardIds = CardIds.toString();
		document.SubmitTreeFrm.CardIds.value = CardIds.toString();
   }
   
   var url =  "../DuplicatePublication.asp?OldFolderId=" + encodeURIComponent(OldFolderId) + "&NewFolderId=" + encodeURIComponent(NewFolderId) + "&DestinationFolderName=" + encodeURIComponent(NewFolderName) + "&SourceFolderName=" + encodeURIComponent(OldFolderName) + "&Total=" + encodeURIComponent(Total);

   document.SubmitTreeFrm.action = url;
   document.SubmitTreeFrm.submit();
}
 

3.	Create new ASP file to trigger the event: 
    A new ASP file “DuplicatePublication.asp” is created to call the duplicate background task. Save it in location "C:\InfoShare\Web\Author\ASP"
    
4.	Update “XML Background Task Setting” from CM web client with below lines of code:

    <!-- Clone Publication ============================================================ -->
    <handler eventType="CLONEPUBLICATION">
      <scheduler executeSynchronously="false" />
      <authorization type="authenticationContext" />
      <execution timeout="01:00:00" recoveryGracePeriod="00:10:00" isolationLevel="None" useSingleThreadApartment="false" />
      <activator>
        <net name="ClonePublication">
          <parameters>
            <parameter name="Action">Clone Publication</parameter>
          </parameters>
        </net>
      </activator>
      <errorHandler maximumRetries="0" />
    </handler>
  
  Add te reference in both default and console roles.
 <add ref="CLONEPUBLICATION" /> 

5. Compile the code and put the compiled dll in location "C:\InfoShare\App\Plugins".

6. Restart BackgroundTask service and Component services in CM server.
