Attribute VB_Name = "modSHFileOperation"
Option Explicit

         Private Const FO_COPY = &H2&   'Copies the files specified
                                        'in the pFrom member to the
                                        'location specified in the
                                        'pTo member.

         Private Const FO_DELETE = &H3& 'Deletes the files specified
                                        'in pFrom (pTo is ignored.)

         Private Const FO_MOVE = &H1&   'Moves the files specified
                                        'in pFrom to the location
                                        'specified in pTo.

         Private Const FO_RENAME = &H4& 'Renames the files
                                        'specified in pFrom.

         Private Const FOF_ALLOWUNDO = &H40&   'Preserve Undo information.

         Private Const FOF_CONFIRMMOUSE = &H2& 'Not currently implemented.

         Private Const FOF_CREATEPROGRESSDLG = &H0& 'handle to the parent
                                                    'window for the
                                                    'progress dialog box.

         Private Const FOF_FILESONLY = &H80&        'Perform the operation
                                                    'on files only if a
                                                    'wildcard file name
                                                    '(*.*) is specified.

         Private Const FOF_MULTIDESTFILES = &H1&    'The pTo member
                                                    'specifies multiple
                                                    'destination files (one
                                                    'for each source file)
                                                    'rather than one
                                                    'directory where all
                                                    'source files are
                                                    'to be deposited.

         Private Const FOF_NOCONFIRMATION = &H10&   'Respond with Yes to
                                                    'All for any dialog box
                                                    'that is displayed.

         Private Const FOF_NOCONFIRMMKDIR = &H200&  'Does not confirm the
                                                    'creation of a new
                                                    'directory if the
                                                    'operation requires one
                                                    'to be created.

         Private Const FOF_RENAMEONCOLLISION = &H8& 'Give the file being
                                                    'operated on a new name
                                                    'in a move, copy, or
                                                    'rename operation if a
                                                    'file with the target
                                                    'name already exists.

         Private Const FOF_SILENT = &H4&            'Does not display a
                                                    'progress dialog box.

         Private Const FOF_SIMPLEPROGRESS = &H100&  'Displays a progress
                                                    'dialog box but does
                                                    'not show the
                                                    'file names.

         Private Const FOF_WANTMAPPINGHANDLE = &H20&
                                   'If FOF_RENAMEONCOLLISION is specified,
                                   'the hNameMappings member will be filled
                                   'in if any files were renamed.

         ' The SHFILOPSTRUCT is not double-word aligned. If no steps are
         ' taken, the last 3 variables will not be passed correctly. This
         ' has no impact unless the progress title needs to be changed.

         Private Type SHFILEOPSTRUCT
            hwnd As Long
            wFunc As Long
            pFrom As String
            pTo As String
            fFlags As Integer
            fAnyOperationsAborted As Long
            hNameMappings As Long
            lpszProgressTitle As String
         End Type

         Private Declare Function SHFileOperation Lib "Shell32.dll" _
               Alias "SHFileOperationA" _
               (lpFileOp As Any) As Long

         Public Sub MyMoveFileOrFolder(lHWND As Long, sFrom As Variant, sTo As String)
            Dim result As Long
            Dim fileop As SHFILEOPSTRUCT

            With fileop
               .hwnd = lHWND

               .wFunc = FO_MOVE
               .fFlags = FOF_ALLOWUNDO Or FOF_NOCONFIRMATION Or FOF_SILENT
               ' The files to copy separated by Nulls and terminated by two
               ' nulls
                .pFrom = sFrom & vbNullChar & vbNullChar

               .pTo = sTo & vbNullChar & vbNullChar

            End With

            result = SHFileOperation(fileop)

         End Sub


         Public Sub MyCopyFileOrFolder(lHWND As Long, sFrom As Variant, sTo As String)
            Dim result As Long
            Dim fileop As SHFILEOPSTRUCT

            With fileop
               .hwnd = lHWND

               .wFunc = FO_COPY
               .fFlags = FOF_ALLOWUNDO Or FOF_NOCONFIRMATION Or FOF_SILENT
               ' The files to copy separated by Nulls and terminated by two
               ' nulls
                .pFrom = sFrom & vbNullChar & vbNullChar

               .pTo = sTo & vbNullChar & vbNullChar

            End With

            result = SHFileOperation(fileop)

         End Sub

         Public Sub MyDeleteFileOrFolder(lHWND As Long, sFileOrFolderPath As Variant)
            Dim result As Long
            Dim fileop As SHFILEOPSTRUCT

            With fileop
               .hwnd = lHWND

               .wFunc = FO_DELETE
               .fFlags = FOF_ALLOWUNDO Or FOF_NOCONFIRMATION Or FOF_SILENT
               ' The files to copy separated by Nulls and terminated by two
               ' nulls
                .pFrom = sFileOrFolderPath & vbNullChar & vbNullChar

'               .pTo = sTo & vbNullChar & vbNullChar

            End With

            result = SHFileOperation(fileop)

         End Sub

