<#
Auth: tomdav
this script looks at everything in a folder and prints it
changes:
- need to give the option to limit to specific file types
- need to give option to aim it towards a specific printer
#>

Function Get-Folder($initialDirectory="")
<#this looks for a folder directory with a nice gui familiar thing for people
this asks the user for the folder and saves it as a get-folder variable, then allows it to be refrenced outside the function
#>
{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")|Out-Null

    $foldername = New-Object System.Windows.Forms.FolderBrowserDialog
    $foldername.Description = "Select a folder"
    $foldername.rootfolder = "MyComputer"
    $foldername.SelectedPath = $initialDirectory

    if($foldername.ShowDialog() -eq "OK")
    {
        $folder += $foldername.SelectedPath
    }
    return $folder
}

# shows user the name of everything that's in the folder incase they clicked the wrong folder, prompts for a confirmation to print to default printer
$a = Get-Folder
dir $a -Name
$Ans = Read-Host -prompt "do you want to print these files? (y / N)"
    If ($Ans -eq "y") 
        {write-host "printing to your default printer"
    dir "$a" | ForEach-Object {Start-Process -FilePath $_.FullName -Verb Print}
        }
        else
        {read-host "press any key to exit" 
        break}
