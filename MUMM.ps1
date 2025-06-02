<#
An upgrade to the CLI + CSV MUMM script.

If you're reading this, you're not supposed to!!!
Instead, right-click this file, then click "Run with PowerShell"!!!

Author: Raymond Tamse Jr
Date: March 12 2024
#>

try {
    Get-Mailbox TestMailboxAzure -ErrorAction Stop
} catch {
    Connect-ExchangeOnline
}
Add-Type -Assembly System.Windows.Forms


# --- Hide PowerShell console window ---
Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("kernel32.dll")]
public static extern IntPtr GetConsoleWindow();

[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
'

$ConsoleHandle = [Console.Window]::GetConsoleWindow()

if ($ConsoleHandle -ne [IntPtr]::Zero) {
    [Console.Window]::ShowWindow($ConsoleHandle, 0)
}
# --- end ---

# --- Style Function ---
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

# --- MAIN WINDOW ---
$MainForm = New-Object System.Windows.Forms.Form
$MainForm.Text = 'MUMM Script'
$MainForm.Width = 1000
$MainForm.Height = 600
$MainForm.MaximumSize = 100

# --- Style Variables ---
$ParagraphSize = 12
$ParagraphFamily = "Poppins"
$Paragraph = New-Object System.Drawing.Font($ParagraphFamily, $ParagraphSize)
$StartPosition = New-Object System.Drawing.Point(10, 10)
$Width100 = $MainForm.ClientSize.Width - 20
$Width50 = ($MainForm.ClientSize.Width/2) - 10

# --- Mailbox Form ---
$RadioGroupBox1 = New-Object System.Windows.Forms.GroupBox
$RadioGroupBox1.Location = Get-NewPosition $StartPosition 0 0
$RadioGroupBox1.Font = $Paragraph
$RadioGroupBox1.Width = $Width50
$RadioGroupBox1.Text = "Is this for a mailbox or calendar?"
$RadioGroupBox1.Height = 60
$MainForm.Controls.Add($RadioGroupBox1)

$MailboxSelectBtn = New-Object System.Windows.Forms.RadioButton
$MailboxSelectBtn.Location = Get-NewPosition $StartPosition 0 20
$MailboxSelectBtn.Font = $Paragraph
$MailboxSelectBtn.Text = "Mailbox"
$CalendarSelectBtn = New-Object System.Windows.Forms.RadioButton
$CalendarSelectBtn.Location = Get-NewPosition $StartPosition ($RadioGroupBox1.Width / 2) 20
$CalendarSelectBtn.Font = $Paragraph
$CalendarSelectBtn.Text = "Calendar"
$RadioGroupBox1.Controls.Add($MailboxSelectBtn)
$RadioGroupBox1.Controls.Add($CalendarSelectBtn)

$RadioGroupBox2 = New-Object System.Windows.Forms.GroupBox
$RadioGroupBox2.Location = Get-NewPosition $StartPosition $Width50 0
$RadioGroupBox2.Font = $Paragraph
$RadioGroupBox2.Width = $Width50
$RadioGroupBox2.Text = "Add or remove?"
$RadioGroupBox2.Height = 60
$MainForm.Controls.Add($RadioGroupBox2)

$AddRadioBtn = New-Object System.Windows.Forms.RadioButton
$AddRadioBtn.Location = Get-NewPosition $StartPosition 0 20
$AddRadioBtn.Font = $Paragraph
$AddRadioBtn.Text = "Add"
$RemoveRadioBtn = New-Object System.Windows.Forms.RadioButton
$RemoveRadioBtn.Location = Get-NewPosition $StartPosition ($RadioGroupBox1.Width / 2) 20
$RemoveRadioBtn.Font = $Paragraph
$RemoveRadioBtn.Text = "Remove"
$RadioGroupBox2.Controls.Add($AddRadioBtn)
$RadioGroupBox2.Controls.Add($RemoveRadioBtn)

$AutoMapCheckBox = New-Object System.Windows.Forms.CheckBox
$AutoMapCheckBox.Location = Get-NewPosition $StartPosition 0 60
$AutoMapCheckBox.Font = $Paragraph
$AutoMapCheckBox.AutoSize = $true
$AutoMapCheckBox.Text = "Enable auto-mapping? (mailbox only)"
$MainForm.Controls.Add($AutoMapCheckBox)

$Spreadsheet = New-Object System.Windows.Forms.DataGridView
$Spreadsheet.Location = Get-NewPosition $StartPosition 0 100
$Spreadsheet.Width = $Width100
$Spreadsheet.Height = 200
$Spreadsheet.Font = $Paragraph
$Spreadsheet.ColumnCount = 2
$Spreadsheet.ColumnHeadersVisible = $true
$Spreadsheet.Columns[0].Name = "Username"
$Spreadsheet.Columns[1].Name = "Mailbox"
$Spreadsheet.AutoSizeColumnsMode = "Fill"
$Spreadsheet.AllowDrop = $true;
$Spreadsheet.MultiSelect = $true
$MainForm.Controls.Add($Spreadsheet)

$StartBtn = New-Object System.Windows.Forms.Button
$StartBtn.Location = Get-NewPosition $StartPosition 0 300
$StartBtn.Width = $Width50
$StartBtn.Height = 30
$StartBtn.Font = $Paragraph
$StartBtn.Text = "START"
$MainForm.Controls.Add($StartBtn)

$ClearBtn = New-Object System.Windows.Forms.Button
$ClearBtn.Location = Get-NewPosition $StartPosition $Width50 300
$ClearBtn.Width = $Width50
$ClearBtn.Height = 30
$ClearBtn.Font = $Paragraph
$ClearBtn.Text = "CLEAR"
$MainForm.Controls.Add($ClearBtn)

$OutputBox = New-Object System.Windows.Forms.Textbox
$OutputBox.Location = Get-NewPosition $StartPosition 0 340
$OutputBox.Width = $Width100
$OutputBox.Height = 100
$OutputBox.Font = $Paragraph
$OutputBox.ReadOnly = $true
$OutputBox.Multiline = $true
$OutputBox.BackColor = "#FFFFFF"
$MainForm.Controls.Add($OutputBox)

# To-do: add a pop-up that shows permissions for each mailbox/calendar actioned

# --- Event Handlers ---
$StartBtn.Add_Click({
    if ($MailboxSelectBtn.Checked) {
        $Type = "Mailbox"
    }
    if ($CalendarSelectBtn.Checked) {
        $Type = "Calendar"
    }

    if ($AddRadioBtn.Checked) {
        $Operation = "Adding"
    }
    if ($RemoveRadioBtn.Checked) {
        $Operation = "Removing"
    }
    
    # Removing a row mid-loop will stop the for-loop below.
    # Need to save the rows to be deleted here.
    $RowsToRemove = @()
    
    if ($Type -eq "Mailbox") {
        foreach ($Row in $Spreadsheet.Rows) {
            $Username = $Row.Cells['Username'].Value
            $Mailbox = $Row.Cells['Mailbox'].Value
            if ($Username -ne $null -and $Mailbox -ne $null) {
                $OutputBox.AppendText("$Operation $Username to/from $Mailbox...")
                Set-MailboxPermission $Operation $Username $Mailbox
                $OutputBox.AppendText(" done!")
                $OutputBox.AppendText([System.Environment]::NewLine)
                $RowsToRemove += $Row
            }
        }
    }
    if ($Type -eq "Calendar") {
        foreach ($Row in $Spreadsheet.Rows) {
            $Username = $Row.Cells['Username'].Value
            $Mailbox = $Row.Cells['Mailbox'].Value
            $PermissionLevel = $Row.Cells['Permission Level'].Value
            if ($Username -ne $null -and $Mailbox -ne $null -and $PermissionLevel -ne $null) {
                $OutputBox.AppendText("$Operation $Username to/from $Mailbox...")
                $CalendarName = $Mailbox + ":\Calendar"
                Set-CalendarPermission $Operation $Username $CalendarName $PermissionLevel
                $OutputBox.AppendText(" done!")
                $OutputBox.AppendText([System.Environment]::NewLine)
                $RowsToRemove += $Row
            }
        }
    }

    foreach ($Row in $RowsToRemove) {
        $Spreadsheet.Rows.Remove($Row)
    }
})

$ClearBtn.Add_Click({
    Write-Host $Spreadsheet | Select-Object -Property *
    $Spreadsheet.Rows.Clear()
    $Spreadsheet.Refresh()
})

$MailboxSelectBtn.add_CheckedChanged({
    $Spreadsheet.Refresh()
    if ($MailboxSelectBtn.Checked) {
        $Spreadsheet.ColumnCount = 2
    }
})

$CalendarSelectBtn.add_CheckedChanged({
    $Spreadsheet.Refresh()
    if ($CalendarSelectBtn.Checked) {
        $Spreadsheet.ColumnCount = 3
        $Spreadsheet.Columns[2].Name = "Permission Level"
    }
})

$Spreadsheet.add_KeyDown({
    param($sender, $e)

    $test = $e | Select-Object -Property *
    #$test = $Spreadsheet | Select-Object -Property *

    $V = 86
    $Backspace = 8

    # Backspace
    if ($e.KeyValue -eq $Backspace -and $Spreadsheet.SelectedCells.Count -gt 1) {
        foreach ($Cell in $Spreadsheet.SelectedCells) {
            $Cell.Value = ""
        }
    }

    # Ctrl + V
    if ($e.KeyValue -eq $V -and $e.Modifiers -eq "Control") {
        $CopyPaste = Get-Clipboard
        $CurrentCell = $Spreadsheet.CurrentCell
        $Row = $CurrentCell.RowIndex
        $Col = $CurrentCell.ColumnIndex
        
        # The behavior for the asterisk-row (or new row) is weird.
        # It is not an 'actual' row.

        if ($CopyPaste -is [String]) {
            if ($Row -ge $Spreadsheet.Rows.Count-1) {
                $Spreadsheet.Rows.Add()
            }
            $Spreadsheet.Rows[$Row].Cells[$Col].Value = $CopyPaste
            $Row += 1
        } else {
            for (($i = 0); $i -lt $CopyPaste.Length; $i++) {
                if ($Row -ge $Spreadsheet.Rows.Count-1) {
                    $Spreadsheet.Rows.Add()
                }
                $Spreadsheet.Rows[$Row].Cells[$Col].Value = $CopyPaste[$i]
                $Row += 1
            }
        }
    }
})

$Spreadsheet.Add_CellClick({
    #Write-Host "Cell clicked!"
    #Write-Host $Spreadsheet.CurrentCell | Select-Object -Property *
})

# --- Functions ---
function Set-MailboxPermission {
    param (
        $Operation,
        $Username,
        $Mailbox
    )
    $Result = "done"
    $AutoMapping = $False
    if ($AutoMapCheckBox.Checked) {
        $AutoMapping = $True
    }

    switch ($Operation) {
        "Adding" {
            Add-MailboxPermission $Mailbox -User $Username -AccessRights FullAccess -InheritanceType All -AutoMapping $AutoMapping
            Set-Mailbox $Mailbox -GrantSendOnBehalfTo @{Add="$Username"}
        }
        "Removing" {
            Remove-MailboxPermission -Identity $Mailbox -User $Username -AccessRights fullaccess -Confirm:$False -ErrorAction Stop

            # Removing possible auto-map
            Add-MailboxPermission $Mailbox -User $Username -AccessRights fullaccess -InheritanceType All -AutoMapping $False -ErrorAction Stop | Out-Null
            Remove-MailboxPermission -Identity $Mailbox -User $Username -AccessRights fullaccess -Confirm:$False -ErrorAction Stop

            Set-Mailbox $Mailbox -GrantSendOnBehalfTo @{remove="$Username"}
        }
    }
    return $Result
}

function Set-CalendarPermission {
    param (
        $Operation,
        $Username,
        $Calendar,
        $PermissionLevel
    )

    switch ($Operation) {
        "Adding" {
            Add-MailboxFolderPermission $Calendar -User $Username -AccessRights $PermissionLevel -ErrorAction SilentlyContinue
        }
        "Removing" {
            Remove-MailboxFolderPermission $Calendar -User $Username -Confirm:$False -ErrorAction Stop
        }
    }
}

cls
$MainForm.ShowDialog()