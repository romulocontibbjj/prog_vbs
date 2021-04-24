VBPJ Getting Started, March 2001
Stan Schultes - stan@vbexpert.com

Code Copyright © 2001 by Stan Schultes, All Rights Reserved.

This code demonstrates how email is sent from a VB application using the VB MAPI controls (frmMAPI.frm), the CDO library (CEmailCDO.cls), and using Outlook objects (CEmailOL). 

The MAPI controls come with VB, in all versions since VB3. You need to download and install the CDO library if you don't already have it. If you need the CDO library, see Microsoft Knowledge Base article Q171440. If you don't have Outlook on your machine, you won't be able to use the CEmailOL class - just remove it from the project.

When you run the EmailTest code for the first time, a default Email List is created (DistList). You need to enter at least one Email address and click Save List Info to store the entry. You can enter more than one address by putting a semicolon between entries (ex: ReportList; SupportStaff).

Email settings for EmailTest are stored under Registry key \HKEY_CURRENT_USER\Software\VB and VBA Program Settings\EmailTest\DistList. This structure is designed so that you can have many programs (in place of EmailTest), each having many email lists (in place of DistList). 

See comments in the frmMain.frm and CEmailReg.cls modules for details on implementation. 

If you receive error 32002 using the MAPI controls, this probably indicates that you've forgotten to supply a SendTo address, or there's a problem with the address.

You may experience intermittent errors using CDO - error MAPI_E_LOGON_FAILED (error number -2147221231 or hex 80040111). This appears to be due to a bug in CDO, and the message gets sent anyway. See MS Knowledge Base article Q181739 for information.

MAPI Error codes can be found in MS Knowledge Base articles Q119647 and Q238119.

The sample requires VB 6. It is highly recommended that you install the latest service pack (SP4 or later). Except for the VB6-specific Split and InstrRev functions, the code should work on VB5 with suitable replacements.

