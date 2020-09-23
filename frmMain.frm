VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Associate Files"
   ClientHeight    =   4350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4350
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Restore the .Demo extension"
      Height          =   555
      Left            =   1110
      TabIndex        =   3
      Top             =   2820
      Width           =   2445
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Backup the .Demo extension"
      Height          =   555
      Left            =   1110
      TabIndex        =   2
      Top             =   2220
      Width           =   2445
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Remove The Associate of the .Demo file"
      Height          =   555
      Left            =   1110
      TabIndex        =   1
      Top             =   1620
      Width           =   2445
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Associate .Demo file extension with this application"
      Height          =   555
      Left            =   1110
      TabIndex        =   0
      Top             =   1020
      Width           =   2445
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Associate Files, by Max Raskin (maxim13@internet-zahav.net)
'Fill free to change and modify, but please leave some credits :-)
'April 08 2000

Private Sub Command1_Click()
    'Syntax - Associate(AppTitle, FileExtension, FileType, IconFileName, Parameters)
    '
    'AppTitle - The key that will be created in classes root directory (e.g. "MyApp"
    'FileExtension - The extension of the file to associate (e.g. ".EXT")
    'FileType - The file type of the application, that will be used as a short description (e.g. "Extension Text")
    'IconFileName - Specifies the path of the icon to be used in the application (e.g. "C:\MyApp\Icon.icon"), can be also icon in a libary (e.g. "C:\MyApp\MyDll.dll,2")
    'Parameters - Any parameters that might be used for the applicationf ile (e.g. "/parameter" or "/param1 /param2" etc..)
    Associate "Associate", ".Demo", "Demo File", "Shell32.dll,25"
End Sub

Private Sub Command2_Click()
    'Syntax - RemoveAssociate(AppTitle, FileExtension) - remove association create by associate function
    '
    'AppTitle - The key that will be created in classes root directory (e.g. "MyApp"
    'FileExtension - The extension of the file to associate (e.g. ".EXT")
    RemoveAssociate "Associate", ".Demo"
End Sub

Private Sub Command3_Click()
    'Syntax - BackupAssoc(FileName, AppTitle, FileExtension, FileType) - remove association create by associate function
    
    'FileName - File that the backup should be saved in
    'AppTitle - The key that will be created in classes root directory (e.g. "MyApp"
    'FileExtension - The extension of the file to associate (e.g. ".EXT")
    'FileType - Type of file, e.g. "Demo File"
    BackupAssoc "C:\demoregfile.reg", "Associate", ".demo", "Demo File"
End Sub

Private Sub Command4_Click()
    'Syntax RestoreAssoc(FileName) - mereges a reg file to the registry using regedit.exe (with the /s parameter which means silent merege)
    
    'FileName - specifies the file that should be mereged
    RestoreAssoc "C:\demoregfile.reg"
End Sub

'Loading file from command line example...
Private Sub Form_Load()
    If Command <> "" Then 'file found
        MsgBox "The File " & Command & " was loaded!", vbInformation
    End If
End Sub
