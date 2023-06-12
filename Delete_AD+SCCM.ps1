Add-Type -AssemblyName System.Windows.Forms

# Define function to check if computer exists in Active Directory
function CheckInAD($computerName) {
    $adComputer = Get-ADComputer -Identity $computerName -ErrorAction SilentlyContinue -Properties Description
    if ($adComputer) {
        return $adComputer.Description
    }
    else {
        return "Not Found"
    }
}

# Define function to check build success flag of a remote computer
function CheckBuildSuccess($computerName) {
    $pingResult = Test-Connection -ComputerName $computerName -Count 1 -ErrorAction SilentlyContinue
    if (!$pingResult) {
        return "Not Reachable"
    }

    $filePath = "\\$computerName\C$\CliffordChance\Build Successful.flg"
    if (!(Test-Path $filePath)) {
        return "File Not Found: $($pingResult.IPV4Address.IPAddressToString)"
    }

    $nslookupResult = Resolve-DnsName -Name $computerName
    if ($pingResult.IPV4Address.IPAddressToString -ne $nslookupResult.IPAddress) {
        return "nslookup mismatch"
    }

    return "File Found: $($pingResult.IPV4Address.IPAddressToString)"
}

# Define function to search for computer in SCCM
function SearchInSCCM($computerName) {
    $sccmServer = "LON-SCCM12-SMS"
    $sccmSiteCode = "C01"
    $sccmNamespace = "root\sms\site_$sccmSiteCode"
    $sccmQuery = "SELECT * FROM SMS_R_System WHERE Name = '$computerName'"
    $sccmResults = Get-WmiObject -Query $sccmQuery -Namespace $sccmNamespace -ComputerName $sccmServer -ErrorAction SilentlyContinue
    if ($sccmResults) {
        return "Found"
    }
    else {
        return "Not Found"
    }
}

# Define form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Machine Check - AD SCCM v1.5"
$form.Width = 850
$form.Height = 460
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
$form.StartPosition = "CenterScreen"

# Define label
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10, 20)
$label.Size = New-Object System.Drawing.Size(810, 20)
$label.Text = "Please enter machine numbers separated by commas:"
$form.Controls.Add($label)

# Define textbox
$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(10, 50)
$textBox.Size = New-Object System.Drawing.Size(810, 20)
$form.Controls.Add($textBox)

# Define label
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10, 80)
$label.Size = New-Object System.Drawing.Size(810, 20)
$label.Text = "SELECT the row(s) to remove machine(s) from AD and SCCM"
$form.Controls.Add($label)

# Define label
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10, 100)
$label.Size = New-Object System.Drawing.Size(810, 20)
$label.Text = "Note: Build Success Flag result only shows when machine is connected and PINGABLE via Trusted Network with Internal IP"
$form.Controls.Add($label)

# Define table
$table = New-Object System.Windows.Forms.DataGridView
$table.Location = New-Object System.Drawing.Point(10, 120)
$table.Size = New-Object System.Drawing.Size(810, 100)
$table.ColumnCount = 5
$table.Columns[0].Name = "Computer Name"
$table.Columns[1].Name = "Description"
$table.Columns[2].Name = "Active Directory"
$table.Columns[3].Name = "SCCM"
$table.Columns[4].Name = "Build Success Flag"
$table.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
$table.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells
$table.AutoSizeRowsMode = [System.Windows.Forms.DataGridViewAutoSizeRowsMode]::AllCells
$table.RowCount = 10
$table.Visible = $true
$form.Controls.Add($table)

# Adjust table size to show 10 rows
$numberOfRows = 10
$rowHeight = $table.Rows[0].Height
$headerHeight = $table.ColumnHeadersHeight
$tableHeight = ($numberOfRows * $rowHeight) + $headerHeight
$table.Size = New-Object System.Drawing.Size(810, $tableHeight)

# Define the function to delete computer from Active Directory
function DeleteFromAD($computerName) {
    $computerName = "$computerName$"
    $adComputer = Get-ADComputer -Identity $computerName -ErrorAction SilentlyContinue
    if ($adComputer) {
        Remove-ADObject -Identity $adComputer.DistinguishedName -Recursive -Confirm:$false
        return "Deleted"
    }
    else {
        return "Not Found"
    }
}

# Define the function to delete the computer from SCCM
function DeleteFromSCCM($computerName) {
    $sccmServer = "LON-SCCM12-SMS"
    $sccmSiteCode = "C01"
    $sccmNamespace = "root\sms\site_$sccmSiteCode"
    $sccmQuery = "SELECT * FROM SMS_R_System WHERE Name = '$computerName'"
    $sccmResults = Get-WmiObject -Query $sccmQuery -Namespace $sccmNamespace -ComputerName $sccmServer -ErrorAction SilentlyContinue
    if ($sccmResults) {
        try {
            $sccmResults | Remove-WmiObject
            return "Deleted"
        }
        catch {
            return "Not Found"
        }
    }
    else {
        return "Not Found"
    }
}

# Define check button
$checkButton = New-Object System.Windows.Forms.Button
$checkButton.Location = New-Object System.Drawing.Point(275, 370)
$checkButton.Size = New-Object System.Drawing.Size(80, 40)
$checkButton.Text = "Check"
$checkButton.Add_Click(
    {
        # Create the loading window popup
        $loadingForm = New-Object System.Windows.Forms.Form
        $loadingForm.Text = "Checking machines..."
        $loadingForm.ControlBox = $false
        $loadingForm.MaximizeBox = $false
        $loadingForm.MinimizeBox = $false
        $loadingForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
        $loadingForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
        $loadingLabel = New-Object System.Windows.Forms.Label
        $loadingLabel.Text = "Please wait while the search is in progress...

Do not move this window"
        $loadingLabel.AutoSize = $true
        $loadingLabel.Location = New-Object System.Drawing.Point(10, 10)
        $loadingForm.Controls.Add($loadingLabel)
        $loadingForm.Show()

        # Perform the search
        $table.Rows.Clear()
        $computerNumbers = $textBox.Text -split ","
        foreach ($computerNumber in $computerNumbers) {
            $adResult = CheckInAD($computerNumber.Trim())
            $sccmResult = SearchInSCCM($computerNumber.Trim())
            
#            $buildResult = CheckBuildSuccess($computerNumber.Trim())
            if ($adResult -eq "Not Found" -or $sccmResult -eq "Not Found") {
                $buildResult = "Not Applicable"
                $row = New-Object System.Windows.Forms.DataGridViewRow
                $cell1 = New-Object System.Windows.Forms.DataGridViewTextBoxCell
                $cell1.Value = $computerNumber.Trim()
                $cell2 = New-Object System.Windows.Forms.DataGridViewTextBoxCell
                $cell2.Value = $adResult
                $cell3 = New-Object System.Windows.Forms.DataGridViewTextBoxCell
                $cell3.Value = $adResult -ne "Not Found"
                $cell4 = New-Object System.Windows.Forms.DataGridViewTextBoxCell
                $cell4.Value = $sccmResult
                $cell5 = New-Object System.Windows.Forms.DataGridViewTextBoxCell
                $cell5.Value = $buildResult
                $row.Cells.Add($cell1)
                $row.Cells.Add($cell2)
                $row.Cells.Add($cell3)
                $row.Cells.Add($cell4)
                $row.Cells.Add($cell5)
                $table.Rows.Add($row)
            }
            else {
                $buildResult = CheckBuildSuccess($computerNumber.Trim())
                $row = New-Object System.Windows.Forms.DataGridViewRow
                $cell1 = New-Object System.Windows.Forms.DataGridViewTextBoxCell
                $cell1.Value = $computerNumber.Trim()
                $cell2 = New-Object System.Windows.Forms.DataGridViewTextBoxCell
                $cell2.Value = $adResult
                $cell3 = New-Object System.Windows.Forms.DataGridViewTextBoxCell
                $cell3.Value = $adResult -ne "Not Found"
                $cell4 = New-Object System.Windows.Forms.DataGridViewTextBoxCell
                $cell4.Value = $sccmResult
                $cell5 = New-Object System.Windows.Forms.DataGridViewTextBoxCell
                $cell5.Value = $buildResult
                $row.Cells.Add($cell1)
                $row.Cells.Add($cell2)
                $row.Cells.Add($cell3)
                $row.Cells.Add($cell4)
                $row.Cells.Add($cell5)
                $table.Rows.Add($row)
            }

            # Append the confirmed action to the log file
            $user = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
            $date = Get-Date
            $logFile = "C:\TEMP\623620\Delete_AD+SCCM_LOG.txt"
            Add-Content $logFile -Value "[Check Action] - $user - $date - $computerNumber - AD: $adResult - SCCM: $sccmResult - Build Success Flag: $buildResult" -Force
        }

        # Close the loading window popup
        $loadingForm.Close()
    }
)
$form.Controls.Add($checkButton)

# Define clear button
$clearButton = New-Object System.Windows.Forms.Button
$clearButton.Location = New-Object System.Drawing.Point(365, 370)
$clearButton.Size = New-Object System.Drawing.Size(80, 40)
$clearButton.Text = "Clear"
$clearButton.Add_Click({
        $textBox.Text = ""
        $table.Rows.Clear()
    })
$form.Controls.Add($clearButton)

# Define delete button
$deleteButton = New-Object System.Windows.Forms.Button
$deleteButton.Location = New-Object System.Drawing.Point(455, 370)
$deleteButton.Size = New-Object System.Drawing.Size(100, 40)
$deleteButton.Text = "Delete from AD and SCCM"
$deleteButton.Add_Click({
        $selectedRowIndices = $table.SelectedRows | ForEach-Object { $_.Index }
        if ($selectedRowIndices.Count -gt 0) {
            $result = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to delete the selected machine(s)?", "Confirmation", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
            if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
                foreach ($selectedRowIndex in $selectedRowIndices) {
                    $computerName = $table.Rows[$selectedRowIndex].Cells[0].Value

                    # Create the loading window popup
                    $loadingForm = New-Object System.Windows.Forms.Form
                    $loadingForm.Text = "Deleting selected machine(s) from AD and SCCM..."
                    $loadingForm.ControlBox = $false
                    $loadingForm.MaximizeBox = $false
                    $loadingForm.MinimizeBox = $false
                    $loadingForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
                    $loadingForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
                    $loadingLabel = New-Object System.Windows.Forms.Label
                    $loadingLabel.Text = "Please wait while machines are being deleted...

Do not move this window"
                    $loadingLabel.AutoSize = $true
                    $loadingLabel.Location = New-Object System.Drawing.Point(10, 10)
                    $loadingForm.Controls.Add($loadingLabel)
                    $loadingForm.Show()
            
                    $adDeleteResult = DeleteFromAD($computerName)
                    $sccmDeleteResult = DeleteFromSCCM($computerName)

                    if ($adDeleteResult -eq "Deleted" -and $sccmDeleteResult -eq "Deleted") {
                        $loadingForm.Close()
                        [System.Windows.Forms.MessageBox]::Show("Computer deleted. $computerName AD: $adDeleteResult SCCM: $sccmDeleteResult", "Delete Confirmed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)  
                        
                        # Append the confirmed action to the log file
                        $user = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
                        $date = Get-Date
                        $logFile = "C:\TEMP\623620\Delete_AD+SCCM_LOG.txt"
                        Add-Content $logFile -Value "[Delete Confirmed] - $user - $date - $computerName - AD: $adDeleteResult - SCCM: $sccmDeleteResult" -Force

                        $table.Rows.Remove($table.Rows[$selectedRowIndex])
                    } 
                    else {
                        $loadingForm.Close()
                        $computerName = $table.Rows[$selectedRowIndex].Cells[0].Value
                        [System.Windows.Forms.MessageBox]::Show("Error deleting computer. $computerName AD: $adDeleteResult SCCM: $sccmDeleteResult", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                        
                        # Append the confirmed action to the log file
                        $user = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
                        $date = Get-Date
                        $logFile = "C:\TEMP\623620\Delete_AD+SCCM_LOG.txt"
                        Add-Content $logFile -Value "[Delete Error] - $user - $date - $computerName - AD: $adDeleteResult - SCCM: $sccmDeleteResult" -Force
                    }
                }
            }
        }
    })
$form.Controls.Add($deleteButton)

# Show form
$table.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
$table.AutoSizeRowsMode = [System.Windows.Forms.DataGridViewAutoSizeRowsMode]::AllCells
$form.ShowDialog() | Out-Null