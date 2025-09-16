import-module ActiveDirectory

#create a form using layout from design file
Add-Type -AssemblyName System.Windows.Forms
. (Join-Path $PSScriptRoot 'main.designer.ps1')

#callable function to get list of groups to populate groups_ListBox from ADUC
function getgroups {
    $Groups = Get-ADGroup -Searchbase "OU=365 Groups,OU=Groups,OU=PLP,OU=IFAS,OU=Departments,OU=UF,DC=ad,DC=ufl,DC=edu" -Filter * | Select-Object Name
    return $Groups
}

#callable function to get list of group members of selected group from ADUC
function getgroupmembers {
    param ($Group)
    $GroupMembers = Get-ADGroupMember -Identity $Group
    return $GroupMembers
}

#callable function to remove selected user from selected group
function removefromgroup {
    param ($Group, $Member)
    #$creds = Get-Credential
    try {
        Remove-ADGroupMember -Identity $Group -Members $Member -Confirm:$false <#-Credential $creds#> -ErrorAction stop
        [System.Windows.Forms.MessageBox]::Show("It Worked. $Member is no longer a member of $Group","Success","OK","Warning")
         }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Something went wrong.","Error","OK","Hand")
    }
}

#callable function to add selected user to selected group
function addtogroup {
    param ($Group, $Member)
    $tempmember = Get-ADUser -Filter 'mail -like $Member'
    #$creds = Get-Credential
    try {
        Add-ADGroupMember -Identity $Group -Members $tempmember -Confirm:$false <#-Credential $creds#> -ErrorAction stop
        #Write-Host("It worked")
        [System.Windows.Forms.MessageBox]::Show("It Worked. $Member is now a member of $Group","Success","OK","Warning")
         }
    catch {
         [System.Windows.Forms.MessageBox]::Show("Something went wrong.","Error","OK","Hand")
    }
}

#call function to get the main list of groups and add them to groups_Listbox
#this should be last bit of code before showing the form
$Groups = getgroups
foreach ($Group in $Groups) {
    $Groups_ListBox.Items.Add($Group.Name)
}

#show the form
$Form1.ShowDialog()