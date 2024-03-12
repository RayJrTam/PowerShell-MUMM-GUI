<#
An upgrade to the CLI + CSV MUMM script.

If you're reading this, you're not supposed to!!!
Instead, right-click this file, then click "Run with PowerShell"!!!

Author: Raymond Tamse Jr
Date: March 12 2024
#>

Connect-ExchangeOnline
Add-Type -assembly System.Windows.Forms

# --- Functions ---
function Get-NewPosition {
    param (
        [System.Drawing.Point]$BasePosition,
        [int]$XOffset,
        [int]$YOffset
    )

    $NewX = $BasePosition.X + $XOffset
    $NewY = $BasePosition.Y + $YOffset
    $NewPosition = New-Object System.Drawing.Point($NewX, $NewY)
    return $NewPosition
}

# --- Style Variables ---
$ParagraphSize = 12
$ParagraphFamily = "Poppins"
$Paragraph = New-Object System.Drawing.Font($ParagraphFamily, $ParagraphSize)
$StartPosition = New-Object System.Drawing.Point(10, 10)

# --- Windows ---
$MainForm = New-Object System.Windows.Forms.Form
$MainForm.Text = 'MUMM Script'
$MainForm.Width = 900
$MainForm.Height = 768
#$MainForm.AutoSize = $true

# --- Mailbox Form ---
$MailboxLabelLocation = $StartPosition
$MailboxLabel = New-Object System.Windows.Forms.Label
$MailboxLabel.Text = "Enter the mailbox name or email address:"
$MailboxLabel.Location = $MailboxLabelLocation
$MailboxLabel.AutoSize = $true
$MailboxLabel.Font = $Paragraph
$MainForm.Controls.Add($MailboxLabel)

<#
To-do: You forgot to implement multiple mailboxes, it's not just for multiple users!!!!
Enable multiline, or implement something where it lets users enter multiple
names or email addresses in one-line, separated by ';'
#>
$MailboxTextbox = New-Object System.Windows.Forms.TextBox
$MailboxTextbox.Location = Get-NewPosition -BasePosition $StartPosition -XOffset 0 -YOffset 30
$MailboxTextbox.Width = $MainForm.ClientSize.Width - 30
$MailboxTextbox.Font = $Paragraph
$MainForm.Controls.Add($MailboxTextbox)

$MailboxSelectBtn = New-Object System.Windows.Forms.RadioButton
$MailboxSelectBtn.Location = Get-NewPosition $StartPosition 0 60
$MailboxSelectBtn.Width = ($MainForm.ClientSize.Width / 2) - 30
$MailboxSelectBtn.Font = $Paragraph
$MailboxSelectBtn.Text = "Mailbox"
$CalendarSelectBtn = New-Object System.Windows.Forms.RadioButton
$CalendarSelectBtn.Location = Get-NewPosition $StartPosition ($MainForm.ClientSize.Width / 2) 60
$CalendarSelectBtn.Width = ($MainForm.ClientSize.Width / 2) - 30
$CalendarSelectBtn.Font = $Paragraph
$CalendarSelectBtn.Text = "Calendar"
$MainForm.Controls.Add($MailboxSelectBtn)
$MainForm.Controls.Add($CalendarSelectBtn)

$MailboxSearchBtn = New-Object System.Windows.Forms.Button
$MailboxSearchBtn.Location = Get-NewPosition $StartPosition 0 90
$MailboxSearchBtn.Width = $MainForm.ClientSize.Width - 30
$MailboxSearchBtn.Font = $Paragraph
$MailboxSearchBtn.Text = "SEARCH"
$MainForm.Controls.Add($MailboxSearchBtn)

$AddUserList = New-Object System.Windows.Forms.ListView
$AddUserList.Location = Get-NewPosition $StartPosition 0 150
$AddUserList.Width = ($MainForm.ClientSize.Width / 2) - 30
$AddUserList.Height = 300
$AddUserList.Columns.Add("Users to add", ($MainForm.ClientSize.Width / 2) - 35)
$AddUserList.MultiSelect = $true
$AddUserList.View = "Details"
$AddUserList.Font = $Paragraph
$MainForm.Controls.Add($AddUserList)

$AddUserListBtn = New-Object System.Windows.Forms.Button
$AddUserListBtn.Location = Get-NewPosition $StartPosition 0 500
$AddUserListBtn.Text = "Add"
$AddUserListBtn.Font = $Paragraph
$AddUserListBtn.Width = ($MainForm.ClientSize.Width / 2) - 30
$MainForm.Controls.Add($AddUserListBtn)

$RemoveUserListBtn = New-Object System.Windows.Forms.Button
$RemoveUserListBtn.Location = Get-NewPosition $StartPosition 0 520
$RemoveUserListBtn.Text = "Remove"
$RemoveUserListBtn.Font = $Paragraph
$RemoveUserListBtn.Width = ($MainForm.ClientSize.Width / 2) - 30
$MainForm.Controls.Add($RemoveUserListBtn)

$MailboxUserList = New-Object System.Windows.Forms.ListView
$MailboxUserList.Location = Get-NewPosition $StartPosition ($MainForm.ClientSize.Width / 2) 150
$MailboxUserList.Width = ($MainForm.ClientSize.Width / 2) - 30
$MailboxUserList.Height = 300
$MailboxUserList.Columns.Add("Current users", ($MainForm.ClientSize.Width / 2) - 35)
$MailboxUserList.MultiSelect = $true
$MailboxUserList.View = "Details"
$MailboxUserList.Font = $Paragraph
$MailboxUserList.AutoSize = $true
$MainForm.Controls.Add($MailboxUserList)



# --- Event Handlers ---
# To-do: Modularize this....
$MailboxSearchBtn.Add_Click({
    $MailboxUserList.Items.Clear()
    $MailboxName = $MailboxTextbox.Text
    if ($MailboxSelectBtn.Checked) {
        $CurrentList = (Get-MailboxPermission $MailboxName | Sort-Object User).User
    } elseif ($CalendarSelectBtn.Checked) {
        $CalendarName = $MailboxName + ":\Calendar"
        $CurrentList = (Get-MailboxFolderPermission $CalendarName | Sort-Object {$_.User.DisplayName}).User
    } else {
        Write-Host "Please select Mailbox or Calendar to confirm which operation to perform!"
    }
    Write-Host "Searching $MailboxName"
    foreach ($Username in $CurrentList) {
        if ($Username -notlike "NT AUTHORITY*" -and $Username.DisplayName -ne "Default" -and $Username.DisplayName -ne "Anonymous") {
            $ListViewItem = New-Object System.Windows.Forms.ListViewItem
            $ListViewItem.Text = $Username
            $MailboxUserList.Items.Add($ListViewItem)
        }
    }
})
cls
$MainForm.ShowDialog()
