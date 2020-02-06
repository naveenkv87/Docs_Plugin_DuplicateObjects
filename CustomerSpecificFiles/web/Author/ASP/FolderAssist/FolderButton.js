function FolderAssist(AssistTitle, FolderId, FolderName, UserGroup, ParentFolderId, FolderType, ActionOnOK, RelativePath)
{
	// Calls the folderAssist
	// - RequiredFolderType = N: folders of different type than the specified foldertype will be active.
	// - ActionOnOk is execute when OK-button is pressed.

  //alert('FolderAssist id: ' + FolderId + ' name: ' + FolderName + ' type: ' + FolderType);

	var URL = RelativePath + "FolderAssist/FolderAssistFrame.asp?NewTree=Y&AssistTitle=" + encodeURIComponent(AssistTitle) + "&FolderUserGroup=" + encodeURIComponent(UserGroup) + "&FolderId=" + encodeURIComponent(FolderId) + "&FolderName=" + encodeURIComponent(FolderName) + "&ParentFolderId=" + encodeURIComponent(ParentFolderId) + "&FolderType=" + encodeURIComponent(FolderType) + "&FolderTypeIsRequired=N&ActionOnOk=" + encodeURIComponent(ActionOnOK);
	var WindowStr = "dependent=yes,toolbar=no,width=450,height=350,resizable=yes,scrollbars=no,menubar=no,status=yes,screenX=50,screenY=20";
	FolderAssistWindow = window.open(URL, "FolderName", WindowStr);
	FolderAssistWindow.opener = self;
}

function FolderAssistForObjects(AssistTitle, FolderId, FolderName, UserGroup, FolderType, FolderPath, ActionOnOK, RelativePath)
{
	// Calls the folderAssist
	// - RequiredFolderType = Y: only folder of the specified foldertype will be active.
	// - ActionOnOk is execute when OK-button is pressed.

  //alert('FolderAssistForObjects id: ' + FolderId + ' name: ' + FolderName + ' type: ' + FolderType);
  var URL = RelativePath + "FolderAssist/FolderAssistFrame.asp?NewTree=Y&AssistTitle=" + encodeURIComponent(AssistTitle) + "&FolderUserGroup=" + encodeURIComponent(UserGroup) + "&FolderId=" + encodeURIComponent(FolderId) + "&FolderName=" + encodeURIComponent(FolderName) + "&FolderType=" + encodeURIComponent(FolderType) + "&FolderTypeIsRequired=Y&FolderPath=" + encodeURIComponent(FolderPath) + "&ActionOnOk=" + encodeURIComponent(ActionOnOK);
  
	var WindowStr = "dependent=yes,toolbar=no,width=450,height=350,resizable=yes,scrollbars=no,menubar=no,status=yes,screenX=50,screenY=20";
	FolderAssistWindow = window.open(URL, "FolderName", WindowStr);
	FolderAssistWindow.opener = self;
}

function PerformCreateShortcut()
{
	// This function is executed when FolderAssist OK button is pressed.
	// - used for "creating shortcuts of objects.
	
	//alert('perform create shortcut');
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
   
   var url =  "../PerformCreateShortcut.asp?ShortcutFolderId=" + encodeURIComponent(NewFolderId) + "&DestinationFolderName=" + encodeURIComponent(NewFolderName) + "&SourceFolderName=" + encodeURIComponent(OldFolderName) + "&Total=" + encodeURIComponent(Total);

   document.SubmitTreeFrm.action = url;
   document.SubmitTreeFrm.submit();
}

function PerformMoveObjects()
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
 
   if (parent.TreeFrmFolderUserGroupSelected != parent.TreeFrmFolderUserGroup)
	{
		// First confirm the move between 2 different usergroups!!!
      var url =  "../MoveFromToRepositoryConfirm.asp?NewUsergroup=" + encodeURIComponent(parent.TreeFrmFolderUserGroupSelected) + "&OldUsergroup=" + encodeURIComponent(parent.TreeFrmFolderUserGroup)  + "&OldFolderId=" + encodeURIComponent(OldFolderId) + "&NewFolderId=" + encodeURIComponent(NewFolderId) + "&DestinationFolderName=" + encodeURIComponent(NewFolderName) + "&SourceFolderName=" + encodeURIComponent(OldFolderName) + "&Total=" + encodeURIComponent(Total);
	}
	else
	{
      var url =  "../PerformMove.asp?OldFolderId=" + encodeURIComponent(OldFolderId) + "&NewFolderId=" + encodeURIComponent(NewFolderId) + "&DestinationFolderName=" + encodeURIComponent(NewFolderName) + "&SourceFolderName=" + encodeURIComponent(OldFolderName) + "&Total=" + encodeURIComponent(Total);
	}

   document.SubmitTreeFrm.action = url;
   document.SubmitTreeFrm.submit();
}

function PerformMoveFolder()
{
// This function is executed when FolderAssist OK button is pressed.
//   - used for "moving a folder to another location

	var NewParentFolderId = parent.TreeFrame.GetSelected();
	var OldParentFolderId = parent.TreeFrmParentFolderId;
	var FolderId = parent.TreeFrmFolderId;

	if( NewParentFolderId.length > 0 )
	{	

		if( NewParentFolderId == OldParentFolderId )
		{
			alert("Source and Destination Folder are the same.");
		}
      else
      {
		   if( NewParentFolderId == FolderId )
		   {
			   alert("Folder cannot be moved to one of its children.");
		   }
		   else
		   {
            var url =  "../PerformMoveFolder.asp?OldParentFolderId=" + encodeURIComponent(OldParentFolderId) + "&NewParentFolderId=" + encodeURIComponent(NewParentFolderId) + "&FolderId=" + encodeURIComponent(FolderId);

            document.SubmitTreeFrm.action = url;
            document.SubmitTreeFrm.submit();
		   }	
      }
	}
}

function ShowFolderAssist(FolderId, FolderName, UserGroup, FolderType, RelativePath)
{
   var AssistTitle;
   var ObjectType;
   
   if (FolderType =='')
   {
      AssistTitle = 'Select a folder';
      FolderAssist(AssistTitle, FolderId, FolderName, UserGroup, '', '', 'PerformCopyFolderInfoToDestination()', RelativePath)
   }
   else
   {
      switch( FolderType) 
		      {
            case 'VDOCTYPEMASTER':
		         ObjectType ='maps';
	            break;
   		   case 'VDOCTYPEMAP':
		         ObjectType ='topics';
	            break;
      	   case 'VDOCTYPELIB':
		         ObjectType ='library topics';
	            break;
   		   case 'VDOCTYPEILLUSTRATION':
		         ObjectType ='images';
	            break;
            case 'VDOCTYPETEMPLATE':
		         ObjectType ='other (Word, PDF, ...)';
	            break;
		      case 'VDOCTYPEPUBLICATION':
		         ObjectType ='publications';
	            break;
            }
                  		
			AssistTitle = 'Select a folder for ' + ObjectType;
      FolderAssistForObjects(AssistTitle, FolderId, FolderName, UserGroup, FolderType, '', 'PerformCopyFolderInfoToDestination()', RelativePath)
   }
}

function PerformCopyFolderInfoToDestination()
{
// This function is executed when FolderAssist OK button is pressed.
//  - the selected folderid and foldername is copied into a fixed field: ishfolderref and IshFolderLabel
//  - used in import/Batchimport

		var FieldId
		var NewParentFolderId = parent.TreeFrame.GetSelected() ;
		var NewParentFolderName = parent.TreeFrame.GetCurrentName() ;
		var NewFolderPath = new String();
		NewFolderPath= unescape(parent.TreeFrmFolderPathSelected);
		
		var OldParentFolderId = parent.TreeFrmParentFolderId;
		var FolderId = parent.TreeFrmFolderId;
		
		FieldId = window.top.opener.document.inputForm.ishfolderref;
    FieldId.value=NewParentFolderId;
		FieldId=window.top.opener.document.inputForm.ishfolderlabel;
    FieldId.value=unescape( NewParentFolderName);
    FieldId.title=  unescape( NewFolderPath.replace(/\|/g, " / "));
    parent.window.close();
}

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
