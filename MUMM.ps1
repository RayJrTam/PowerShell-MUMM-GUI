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
$ParagraphFamily = "Arial"
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
$MailboxLabelLeading = 40
$MailboxLabel = New-Object System.Windows.Forms.Label
$MailboxLabel.Text = "Enter the mailbox name or email address:"
$MailboxLabel.Location = $MailboxLabelLocation
$MailboxLabel.AutoSize = $true
$MailboxLabel.Font = $Paragraph
$MainForm.Controls.Add($MailboxLabel)

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

$MailboxUserList = New-Object System.Windows.Forms.ListView
$MailboxUserList.Location = Get-NewPosition $StartPosition 0 150
$MailboxUserList.Width = ($MainForm.ClientSize.Width / 2) - 30
$MailboxUserList.Height = 400
$MailboxUserList.Columns.Add("Users", ($MainForm.ClientSize.Width / 2) - 40)
$MailboxUserList.MultiSelect = $true
$MailboxUserList.View = "Details"
$MailboxUserList.Font = $Paragraph
$MailboxUserList.AutoSize = $true
$MainForm.Controls.Add($MailboxUserList)




# --- Event Handlers ---
$MailboxSearchBtn.Add_Click({
    $MailboxUserList.Items.Clear()
    $MailboxName = $MailboxTextbox.Text
    Write-Host "Searching $MailboxName"
    $CurrentList = (Get-MailboxPermission $MailboxName | Sort-Object User).User
    foreach ($Username in $CurrentList) {
        if ($Username -notlike "NT AUTHORITY*") {
            $ListViewItem = New-Object System.Windows.Forms.ListViewItem
            $ListViewItem.Text = $Username
            $MailboxUserList.Items.Add($ListViewItem)
        }
    }
    
    Write-Host $UserList
})

$MainForm.ShowDialog()
cls