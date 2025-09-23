$Form1 = New-Object -TypeName System.Windows.Forms.Form
[System.Windows.Forms.ListBox]$Groups_Listbox = $null
[System.Windows.Forms.ListView]$Members_Listview = $null
[System.Windows.Forms.ColumnHeader]$Gatorlink = $null
[System.Windows.Forms.ColumnHeader]$Surname = $null
[System.Windows.Forms.ColumnHeader]$GivenName = $null
[System.Windows.Forms.ColumnHeader]$Email = $null
[System.Windows.Forms.Label]$email_Label = $null
[System.Windows.Forms.TextBox]$email_TextBox = $null
[System.Windows.Forms.Button]$search_Button = $null
[System.Windows.Forms.TextBox]$results_Textbox = $null
[System.Windows.Forms.Button]$add_Button = $null
[System.Windows.Forms.Button]$remove_Button = $null
[System.Windows.Forms.Label]$group_Label = $null
function InitializeComponent {
    $Groups_Listbox = (New-Object -TypeName System.Windows.Forms.ListBox)
    $Members_Listview = (New-Object -TypeName System.Windows.Forms.ListView)
    $Gatorlink = (New-Object -TypeName System.Windows.Forms.ColumnHeader)
    $Surname = (New-Object -TypeName System.Windows.Forms.ColumnHeader)
    $GivenName = (New-Object -TypeName System.Windows.Forms.ColumnHeader)
    $Email = (New-Object -TypeName System.Windows.Forms.ColumnHeader)
    $email_Label = (New-Object -TypeName System.Windows.Forms.Label)
    $email_TextBox = (New-Object -TypeName System.Windows.Forms.TextBox)
    $search_Button = (New-Object -TypeName System.Windows.Forms.Button)
    $results_Textbox = (New-Object -TypeName System.Windows.Forms.TextBox)
    $add_Button = (New-Object -TypeName System.Windows.Forms.Button)
    $remove_Button = (New-Object -TypeName System.Windows.Forms.Button)
    $group_Label = (New-Object -TypeName System.Windows.Forms.Label)
    $Form1.SuspendLayout()
    #
    #Groups_Listbox
    #
    $Groups_Listbox.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Tahoma', [System.Single]12))
    $Groups_Listbox.FormattingEnabled = $true
    $Groups_Listbox.ItemHeight = [System.Int32]19
    $Groups_Listbox.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]12, [System.Int32]52))
    $Groups_Listbox.Name = [System.String]'Groups_Listbox'
    $Groups_Listbox.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]303, [System.Int32]669))
    $Groups_Listbox.TabIndex = [System.Int32]0
    $Groups_Listbox.add_SelectedIndexChanged($Groups_Listbox_Click)
    #
    #Members_Listview
    #
    $Members_Listview.Columns.AddRange([System.Windows.Forms.ColumnHeader[]]@($Gatorlink, $Surname, $GivenName, $Email))
    $Members_Listview.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Tahoma', [System.Single]12))
    $Members_Listview.FullRowSelect = $true
    $Members_Listview.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]321, [System.Int32]52))
    $Members_Listview.MultiSelect = $false
    $Members_Listview.Name = [System.String]'Members_Listview'
    $Members_Listview.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]1006, [System.Int32]667))
    $Members_Listview.TabIndex = [System.Int32]1
    $Members_Listview.UseCompatibleStateImageBehavior = $false
    $Members_Listview.GridLines = $true
    $Members_Listview.View = [System.Windows.Forms.View]::Details
    $Members_Listview.add_SelectedIndexChanged($Members_Listview_Click)

    #attach sorting functionality to members_listview
    $sorter = [ListViewColumnSorter]::new()
    $Members_Listview.ListViewItemSorter = $sorter
    $null = $Members_Listview.Add_ColumnClick({
            param($sender, $e)

            # If we click the same column we flip the order
            if ($sorter.SortColumn -eq $e.Column) {
                if ($sorter.Order -eq [System.Windows.Forms.SortOrder]::Ascending) {
                    $sorter.Order = [System.Windows.Forms.SortOrder]::Descending
                }
                else {
                    $sorter.Order = [System.Windows.Forms.SortOrder]::Ascending
                }
            }
            else {
                # New column – start with ascending
                $sorter.SortColumn = $e.Column
                $sorter.Order = [System.Windows.Forms.SortOrder]::Ascending
            }

            # Tell the ListView to re‑sort based on the new settings
            $Members_Listview.Sort()
        })
    #
    #Gatorlink
    #
    $Gatorlink.Text = [System.String]'Gatorlink'
    $Gatorlink.Width = [System.Int32]276
    #
    #Surname
    #
    $Surname.Text = [System.String]'Surname'
    $Surname.Width = [System.Int32]193
    #
    #GivenName
    #
    $GivenName.Text = [System.String]'Given Name'
    $GivenName.Width = [System.Int32]195
    #
    #Email
    #
    $Email.Text = [System.String]'Email'
    $Email.Width = [System.Int32]321
    #
    #email_Label
    #
    $email_Label.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Tahoma', [System.Single]12))
    $email_Label.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]1333, [System.Int32]53))
    $email_Label.Name = [System.String]'email_Label'
    $email_Label.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]180, [System.Int32]23))
    $email_Label.TabIndex = [System.Int32]2
    $email_Label.Text = [System.String]'Enter an email address:'
    #
    #email_TextBox
    #
    $email_TextBox.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Tahoma', [System.Single]12))
    $email_TextBox.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]1333, [System.Int32]79))
    $email_TextBox.Name = [System.String]'email_TextBox'
    $email_TextBox.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]400, [System.Int32]27))
    $email_TextBox.TabIndex = [System.Int32]3
    #
    #search_Button
    #
    $search_Button.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Tahoma', [System.Single]12))
    $search_Button.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]1333, [System.Int32]106))
    $search_Button.Name = [System.String]'search_Button'
    $search_Button.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]150, [System.Int32]30))
    $search_Button.TabIndex = [System.Int32]4
    $search_Button.Text = [System.String]'Search'
    $search_Button.UseVisualStyleBackColor = $true
    $search_Button.add_Click($searchButton_Click)
    #
    #results_Textbox
    #
    $results_Textbox.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Tahoma', [System.Single]12))
    $results_Textbox.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]1333, [System.Int32]142))
    $results_Textbox.Multiline = $true
    $results_Textbox.Name = [System.String]'results_Textbox'
    $results_Textbox.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]400, [System.Int32]192))
    $results_Textbox.TabIndex = [System.Int32]5
    #
    #add_Button
    #
    $add_Button.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Tahoma', [System.Single]12))
    $add_Button.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]1333, [System.Int32]340))
    $add_Button.Name = [System.String]'add_Button'
    $add_Button.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]150, [System.Int32]30))
    $add_Button.TabIndex = [System.Int32]6
    $add_Button.Text = [System.String]'Add Member'
    $add_Button.UseVisualStyleBackColor = $true
    $add_Button.add_Click($add_Button_Click)
    #
    #remove_Button
    #
    $remove_Button.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Tahoma', [System.Single]12))
    $remove_Button.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]1333, [System.Int32]376))
    $remove_Button.Name = [System.String]'remove_Button'
    $remove_Button.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]150, [System.Int32]30))
    $remove_Button.TabIndex = [System.Int32]7
    $remove_Button.Text = [System.String]'Remove Member'
    $remove_Button.UseVisualStyleBackColor = $true
    $remove_Button.add_Click($remove_Button_Click)
    #
    #group_Label
    #
    $group_Label.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Tahoma', [System.Single]12))
    $group_Label.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]321, [System.Int32]19))
    $group_Label.Name = [System.String]'group_Label'
    $group_Label.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]455, [System.Int32]30))
    $group_Label.TabIndex = [System.Int32]8
    $group_Label.Text = [System.String]'Selected Group'
    #
    #Form1
    #
    $Form1.ClientSize = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]1800, [System.Int32]800))
    $Form1.Controls.Add($group_Label)
    $Form1.Controls.Add($remove_Button)
    $Form1.Controls.Add($add_Button)
    $Form1.Controls.Add($results_Textbox)
    $Form1.Controls.Add($search_Button)
    $Form1.Controls.Add($email_TextBox)
    $Form1.Controls.Add($email_Label)
    $Form1.Controls.Add($Members_Listview)
    $Form1.Controls.Add($Groups_Listbox)
    $Form1.Text = [System.String]'Email List Editor'
    $Form1.ResumeLayout($false)
    $Form1.PerformLayout()
    Add-Member -InputObject $Form1 -Name Groups_Listbox -Value $Groups_Listbox -MemberType NoteProperty
    Add-Member -InputObject $Form1 -Name Members_Listview -Value $Members_Listview -MemberType NoteProperty
    Add-Member -InputObject $Form1 -Name Gatorlink -Value $Gatorlink -MemberType NoteProperty
    Add-Member -InputObject $Form1 -Name Surname -Value $Surname -MemberType NoteProperty
    Add-Member -InputObject $Form1 -Name GivenName -Value $GivenName -MemberType NoteProperty
    Add-Member -InputObject $Form1 -Name Email -Value $Email -MemberType NoteProperty
    Add-Member -InputObject $Form1 -Name email_Label -Value $email_Label -MemberType NoteProperty
    Add-Member -InputObject $Form1 -Name email_TextBox -Value $email_TextBox -MemberType NoteProperty
    Add-Member -InputObject $Form1 -Name search_Button -Value $search_Button -MemberType NoteProperty
    Add-Member -InputObject $Form1 -Name results_Textbox -Value $results_Textbox -MemberType NoteProperty
    Add-Member -InputObject $Form1 -Name add_Button -Value $add_Button -MemberType NoteProperty
    Add-Member -InputObject $Form1 -Name remove_Button -Value $remove_Button -MemberType NoteProperty
    Add-Member -InputObject $Form1 -Name group_Label -Value $group_Label -MemberType NoteProperty
}

#####                      ######
#####functions for buttons ######
#####                      ######

#function to run when add button is pressed
$add_Button_Click = {
    $x = $Groups_ListBox.SelectedItem
    $y = $email_TextBox.Text
    $answer = [System.Windows.Forms.MessageBox]::Show( "Are you sure you want to add the user ' $y ' to $x ?", "Confirmation", "YesNo", "Warning" )
    if ($answer -eq 'yes') {
        addtogroup $x $y
    }
}

#function to run when item is selected from group_Listbox
$Groups_Listbox_Click = {
    #update the group_Label above Members_Listview to reflect selected group
    $group_Label.Text = $Groups_Listbox.SelectedItem
    #clear members list before recreating it
    $Members_Listview.Items.Clear()
    #call function to populate members of selected group into Members_Listview
    refreshlistview $Groups_Listbox.SelectedItem
}

#function to run when members_listview is clicked. it will populate the email_Textbox with the selected member and run a search on it
$Members_Listview_Click = {
    $selected_member = $Members_Listview.SelectedItems.text
    $email_Textbox.Text = $selected_member
    &$searchButton_Click
}

#function to run when remove_Button is clicked
$remove_Button_Click = {
    $x = $Groups_ListBox.SelectedItem
    $y = $Members_Listview.SelectedItems.text
    #verification box before removing user
    $answer = [System.Windows.Forms.MessageBox]::Show( "Are you sure you want to remove the user ' $y ' from $x ?", " Removal Confirmation", "YesNo", "Warning" )
    if ($answer -eq 'yes') {
        #call the removefromgroup function
        removefromgroup $x $y
    }
}

#function to run when search button is pressed
$searchButton_Click = {
    $email = $email_Textbox.Text
    $results_Textbox.Text = ""
    # Validate email input
    if ([string]::IsNullOrWhiteSpace($email)) {
        $results_Textbox.Text = "Please enter a valid email address."
        return
    }

    # Append domain if missing
    if (-not $email.Contains("@")) {
        $email = "$email@ufl.edu"
        $email_Textbox.Text = $email
    }

    # Search Active Directory for user with the provided email
    try {
        $user = Get-ADUser -Filter "Mail -eq '$email'" -Properties DisplayName, SamAccountName, Mail, Title, EmployeeID

        if ($user) {
            $results_Textbox.Text = "Display Name: $($user.DisplayName)`r`n" +
            "Username: $($user.SamAccountName)`r`n" +
            "Email: $($user.Mail)`r`n" +
            "Title: $($user.Title)`r`n" +
            "EmployeeID: $($user.EmployeeID)"
        }
        else {
            $results_Textbox.Text = "No user found with that email address."
        }
    }
    catch {
        $results_Textbox.Text = "Error searching Active Directory: $_"
    }
}

#should be last line of this file
. InitializeComponent
