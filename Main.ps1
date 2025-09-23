import-module ActiveDirectory

#create a form using layout from design file
Add-Type -AssemblyName System.Windows.Forms
. (Join-Path $PSScriptRoot 'main.designer.ps1')
. (Join-Path $PSScriptRoot 'sorting.ps1')

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
        $removal_messagebox = [System.Windows.Forms.MessageBox]::Show("It Worked. $Member is no longer a member of $Group","Success","OK","Warning")
         }
    catch {
        $removal_messagebox = [System.Windows.Forms.MessageBox]::Show("Something went wrong.","Error","OK","Hand")
    }
    if ($removal_messagebox -eq "OK") {refreshlistview $Group} 
}

#callable function to add selected user to selected group
function addtogroup {
    param ($Group, $Member)
    $tempmember = Get-ADUser -Filter 'mail -like $Member'
    #$creds = Get-Credential
    try {
        Add-ADGroupMember -Identity $Group -Members $tempmember -Confirm:$false <#-Credential $creds#> -ErrorAction stop
        $added_messagebox = [System.Windows.Forms.MessageBox]::Show("It Worked. $Member is now a member of $Group","Success","OK","Warning")
         }
    catch {
         $added_messagebox = [System.Windows.Forms.MessageBox]::Show("Something went wrong.","Error","OK","Hand")
    }
    if ($added_messagebox -eq "OK") {refreshlistview $Group} 
}

#callable function to initialize or refresh the Members_Listview with members of selected group
function refreshlistview {
    param ($Group)
    #call function to get members of selected group
    $Members = getgroupmembers -Group $Group
    #clear members list before recreating it
    $Members_Listview.Items.Clear()
    #iterate and populate members of selected group into Members_Listview
    foreach ($Member in $Members) {
        $item = New-Object System.Windows.Forms.ListViewItem($Member.Name)
                #only add users, not groups
                if ( $member.objectclass -eq "user" ) {
                    $real_user = Get-ADUser -Identity $Member -Properties SurName, GivenName, UserPrincipalName
                    $item.Subitems.Add($real_user.SurName)
                    $item.Subitems.Add($real_user.GivenName)
                    $item.Subitems.Add($real_user.UserPrincipalName)
                    $Members_Listview.Items.Add($item) | Out-Null
                }
                #show groups with just name so lists arent empty
                if ( $member.objectclass -eq "group" ) {
                    $Members_Listview.Items.Add($item) | Out-Null
                }
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