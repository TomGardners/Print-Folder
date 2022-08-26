<#
auth: Tomdav
asks for a zip file, tells you whats inside it, extracts it, prints everything in it, deletes extracted folder
this will print just PDF's in the folder
#>


#Functions
#extracts and prints from a zip file
function extractprint {
    expand-archive "$zselectedpath" -DestinationPath "$zselectedpath\..\TEMP $zselectedfile" -force  
    dir -path "$zselectedpath\..\TEMP $zselectedfile" -Recurse | Where-Object { $_.Extension -like '*.pdf' } | ForEach-Object {Start-Process -FilePath $_.FullName -Verb Print}
}

#prompts user for a file with a nice gui thing
function askforfile {
Add-Type -AssemblyName System.Windows.Forms
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.Title = "Please Select Zip File"
    $OpenFileDialog.InitialDirectory = $InitialDirectory
    $OpenFileDialog.filter = “Zip File (*.zip) |*.zip”
    If ($OpenFileDialog.ShowDialog() -eq "Cancel") 
    {
    [System.Windows.Forms.MessageBox]::Show("closed. Please select a ZIP file", "Error", 0, 
    [System.Windows.Forms.MessageBoxIcon]::Exclamation)
    }   # stores the zip file as a variable
        $Global:zSelectedFile = $OpenFileDialog.safeFileNames
        # shows the contents of the file
        $Global:zcompdocs = $OpenFileDialog.FileNames
        # stores the path way in variable
        $Global:zselectedpath = $OpenFileDialog.FileName
}

#start of script#
askforfile

#reads the files inside the zip folder
Add-Type -assembly "system.io.compression.filesystem"
[io.compression.zipfile]::OpenRead("$zcompdocs").Entries.Name

#prompt for user to print files listed above
$Ans = Read-Host -prompt "would you like to print the above files? (y / n)"
    If ($Ans -eq "y") 
    {
        #tests if they already have a folder from a previous export with the same name it cleans it out so it won't print anything that's not needed
        If (test-path "$zselectedpath\..\TEMP $zselectedfile") 
           {
           Remove-Item –path "$zselectedpath\..\printed $zselectedfile" –recurse -Force
           extractprint
           }
        Else 
           {
           extractprint
           }
    }
    else
    {
    read-host "press enter key to exit" 
    break
    }
