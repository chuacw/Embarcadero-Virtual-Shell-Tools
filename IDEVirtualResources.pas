unit IDEVirtualResources;

// Version 1.3.0
//
// The contents of this file are subject to the Mozilla Public License
// Version 1.1 (the "License"); you may not use this file except in compliance
// with the License. You may obtain a copy of the License at http://www.mozilla.org/MPL/
//
// Alternatively, you may redistribute this library, use and/or modify it under the terms of the
// GNU Lesser General Public License as published by the Free Software Foundation;
// either version 2.1 of the License, or (at your option) any later version.
// You may obtain a copy of the LGPL at http://www.gnu.org/copyleft/.
//
// Software distributed under the License is distributed on an "AS IS" basis,
// WITHOUT WARRANTY OF ANY KIND, either express or implied. See the License for the
// specific language governing rights and limitations under the License.
//
// The initial developer of this code is Jim Kueneman <jimdk@mindspring.com>
//
//----------------------------------------------------------------------------

interface

uses
  Messages;

{$include Compilers.inc}
{$include VSToolsAddIns.inc}

{$IFNDEF T2H}
const
  WM_SHELLNOTIFY = WM_APP + 51;                    // Change in the Shell occurred
  WM_VTSETICONINDEX = WM_SHELLNOTIFY + 1;          // Threaded Icons
  WM_INVALIDFILENAME = WM_SHELLNOTIFY + 2;         // VET
  WM_SHELLNOTIFYTHREADQUIT = WM_SHELLNOTIFY + 3;   // ShellNotifier
  WM_CHANGENOTIFY_NT = WM_SHELLNOTIFY + 4;         // ShellNotifier
  WM_CHANGENOTIFY = WM_SHELLNOTIFY + 5;            // ShellNotifier
  WM_SHELLNOTIFYTHREADEVENT = WM_SHELLNOTIFY + 6;  // ShellNotifier
  WM_SHELLNOTIFYRELEASE = WM_SHELLNOTIFY + 7;      // ShellNotifier
  WM_REMOVEBUTTON = WM_APP + 8;                    // VirtualShellToolbar
  WM_CHANGENOTIFY_CUSTOM = WM_APP + 9;
  WM_UPDATESCROLLBAR = WM_APP + 10;                // Signal TDropDownWnd to update scrollbar

resourcestring
   S_WARNING = 'Warning';
   S_OPEN = 'Open';

   sName = 'Name';
   sSize = 'Size';
   sType = 'Type';
   sModified = 'Modified';
   sAttributes = 'Attributes';
   sCreated = 'Created';
   sAccessed = 'Accessed';
   sPath = 'Path';
   sDOSName = 'DOS Name';
   sCustom = 'Custom';

const
  // Names shown in column headers if toShellColumnDetails is not used. In that
  // case the shell handles the names based on local.
   VET_COLUMN_NAMES: array[0..9] of string = (
    sName,
    sSize,
    sType,
    sModified,
    sAttributes,
    sCreated,
    sAccessed,
    sPath,
    sDOSName,
    sCustom
  );

resourcestring
    // Error given when the root of VET is set to a path that does not exist.  The
  // end user should never see this message.
  STR_ERR_INVALID_CUSTOMPATH = 'Invalid pathname for Custom Root Path';

  // Menu item text shown when the column popup menu has more than a specified
  // number of items or the column has told VET it should not be shown in the
  // menu but only show it in the dialog box.
  STR_COLUMNMENU_MORE = 'More ...';

    // --------------------------------------------------------------------------
  // TExplorerComboBox messages
  // --------------------------------------------------------------------------

  S_PATH_ERROR = 'Invalid Path';
  S_COMBOEDIT_DEFAULT_ERROR = 'Windows can not find "%s". Please check the spelling and try again.';
  // --------------------------------------------------------------------------

    // --------------------------------------------------------------------------
  // VirtualShellNewMenu strings
  // --------------------------------------------------------------------------
  // The string used as a prefix to the new file being created based on the file type
  // of the extension.  For example a Notepad TXT file will create a new file named
  // New Text Document.txt, where the "New " string is this constant.
  S_NEW = 'New ';
  // These are the strings that will be appended to New when the additional menu
  // item "New Folder" and "New ShortCut" are requested in the menu.
  S_FOLDER = 'Folder';
  S_SHORTCUT = 'Shortcut';
  // This is the string that is shown in the Messagebox if the new file will
  // overwrite an existing file.
  S_OVERWRITE_EXISTING_FILE = 'File exists.  Overwrite existing file?';


const
  // SHOULD NOT HAVE TO EDIT THESE STRINGS
  // applet launcher file, should never have to change this
  S_RUNDLL32 = '\rundll32.exe';
  // The dll launched by RunDll32.exe for the Briefcase with the necessary mods
  // Do not Modify!
  // The %1 seems to be Boolean for create on desktop on not (not being true)
  S_BRIEFCASE_HACK_STRING = 'syncui.dll,Briefcase_Create 1!d! ';
  // Test string to match up with Command string to see if the menu item will
  // create a new Briefcase or Link. Do not Modify!
  S_BRIEFCASE_IDENTIFIER = ',Briefcase_Create';
  S_CREATELINK_IDENTIFIER = ',NewLinkHere';
  S_NULLFILE = 'NullFile';
  S_FILENAME = 'FileName';
  S_COMMAND = 'Command';
  S_DATA = 'Data';
  S_SHELLNEW_PATH = '\ShellNew';
  // --------------------------------------------------------------------------


resourcestring
    sAssociationChanged = 'Association Changed';
    sItemCreate = 'Item Create';
    sItemDelete = 'Item Delete';
    sDriveAdd = 'Drive Add';
    sDriveAddGUI = 'Drive Add GUI';
    sDriveRemoved = 'Drive Removed';
    sFreeSpace = 'Free Space';
    sMediaInserted = 'Media Inserted';
    sMediaRemoved = 'Media Removed';
    sMakeDirectory = 'Make Directory';
    sNetworkShare = 'Network Share';
    sNetworkUnshare = 'Network Unshare';
    sFolderRename = 'Folder Rename';
    sItemRename = 'Item Rename';
    sRemoveDirectory = 'Remove Directory';
    sServerDisconnect = 'Server Disconnect';
    sUpdateDirectory = 'Update Directory';
    sUpdateImage = 'Update Image';
    sUpdateItem = 'Update Item';

const
  // Literal translations of TShellNotifyEvent type.  Useful when using the
  // OnShellNotify event to print out what event occurred.  VirtualShellUtilities.pas
  // has a helper function ShellNotifyEventToStr that uses these.
  VET_NOTIFY_EVENTS: array[0..19] of WideString = (
    sAssociationChanged,
    sAttributes,
    sItemCreate,
    sItemDelete,
    sDriveAdd,
    sDriveAddGUI,
    sDriveRemoved,
    sFreeSpace,
    sMediaInserted,
    sMediaRemoved,
    sMakeDirectory,
    sNetworkShare,
    sNetworkUnshare,
    sFolderRename,
    sItemRename,
    sRemoveDirectory,
    sServerDisconnect,
    sUpdateDirectory,
    sUpdateImage,
    sUpdateItem
  );

resourcestring
  // General Error message
  S_ERROR = 'Error';

  // Strings that format the Details view in KB. Only applies if toShellColumnDetails
  // is not used.
  STR_FILE_SIZE_IN_KB = '%.0n KB';
  STR_ZERO_KB = '0 KB';
  STR_ONE_KB = '1 KB';

  // Strings that are used to show the attributes of a file in Details view.  Only
  // applies if toShellColumnDetails is not used.
  STR_ARCHIVE = 'A';
  STR_HIDDEN = 'H';
  STR_READONLY = 'R';
  STR_SYSTEM = 'S';
  STR_COMPRESS = 'C';

  // What is displayed in the FileType column if VET could not get the information
  // the normal way and it had determined that the item is a system folder.
  STR_SYSTEMFOLDER = 'System Folder';
  STR_FILE = ' File'; // NT is lax in the FileType column if the file is not registered
                                  // it returns nothing.  This is tacked on the end of the file extension
                                  // for example 'PAS Files', BAK Files, ZIP Files and so on.

  // The verb sent to the context menu notification events if the selected context
  // menu item is a non standard verb.
  STR_UNKNOWNCOMMAN = 'Unknown Command';

  // Shown when an operation in the TNamspace is being done on item that are not the
  // direct children of the TNamespace.  This is only a debugging aid, if the tree
  // is set up right this should never occur, i.e. restricted multiselect to one level
  STR_ERR_BAD_PIDL_RELATIONSHIP = 'PIDLs to operate on are not siblings of the Namespace doing the operation.';


  // The name given a new folder when CreateNewFolder is called.
  STR_NEWFOLDER = 'New Folder';

const
  STR_DRIVELETTER_A = 'a';

resourcestring
  // --------------------------------------------------------------------------
  // VirtualShellLink strings
  // --------------------------------------------------------------------------
  // Message shown if an attempt to create a new link is made with no target defined
  S_NOTARGETDEFINED = 'No target application defined';
  // --------------------------------------------------------------------------

  // --------------------------------------------------------------------------
  // Common strings
  // --------------------------------------------------------------------------

  S_PRINT = 'print';
  S_PROPERTIES = 'properties';

  // --------------------------------------------------------------------------

  // ASSERT Strings
  S_KERNELNOTIFERREGISTERED = 'A KernelChangeNotifier is still registered with a Control';
  S_SHELLNOTIFERREGISTERED = 'A ShellChangeThread is still registered with a Control';
  S_SHELLNOTIFERDISPATCHTHREAD = 'ChangeDispatchThread is still registered';
  S_KERNELSPECIALFOLDERWATCH = 'A control must be registered with ' +
     'the VirtualChangeNotifier using RegisterShellChangeNotify before this ' +
     'method may be used.';

{$ENDIF T2H}
implementation

end.
