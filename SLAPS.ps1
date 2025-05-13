Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.DirectoryServices.AccountManagement
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# ===== COLOR SCHEME =====
$primaryColor   = [System.Drawing.Color]::FromArgb(0, 94, 184)  # Primary blue
$secondaryColor = [System.Drawing.Color]::FromArgb(255, 199, 44) # Accent color
$lightGray      = [System.Drawing.Color]::FromArgb(240, 240, 240)
$darkGray       = [System.Drawing.Color]::FromArgb(64, 64, 64)
$errorColor     = [System.Drawing.Color]::FromArgb(220, 53, 69)

# ===== FONT SETTINGS =====
$mainFont       = New-Object System.Drawing.Font("Segoe UI", 10)
$boldFont       = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$titleFont      = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
$monoFont       = New-Object System.Drawing.Font("Consolas", 12)

# ===== MAIN FORM =====
$form = New-Object System.Windows.Forms.Form
$form.Text = "LAPS Password Viewer"
$form.Size = New-Object System.Drawing.Size(500, 350)
$form.StartPosition = "CenterScreen"
$form.BackColor = $lightGray
$form.FormBorderStyle = "FixedSingle"
$form.MaximizeBox = $false

# Header panel
$headerPanel = New-Object System.Windows.Forms.Panel
$headerPanel.Dock = "Top"
$headerPanel.Height = 40
$headerPanel.BackColor = $primaryColor
$form.Controls.Add($headerPanel)

# Header label
$headerLabel = New-Object System.Windows.Forms.Label
$headerLabel.Text = "LAPS PASSWORD VIEWER"
$headerLabel.Font = $titleFont
$headerLabel.ForeColor = [System.Drawing.Color]::White
$headerLabel.Dock = "Left"
$headerLabel.Padding = New-Object System.Windows.Forms.Padding(10, 0, 0, 0)
$headerLabel.AutoSize = $true
$headerPanel.Controls.Add($headerLabel)

# Main content panel
$mainPanel = New-Object System.Windows.Forms.Panel
$mainPanel.Dock = "Fill"
$mainPanel.BackColor = $lightGray
$mainPanel.Padding = New-Object System.Windows.Forms.Padding(10)
$form.Controls.Add($mainPanel)

# Input group
$inputGroup = New-Object System.Windows.Forms.GroupBox
$inputGroup.Text = "Computer Information"
$inputGroup.Font = $boldFont
$inputGroup.ForeColor = $darkGray
$inputGroup.Location = New-Object System.Drawing.Point(10, 10)
$inputGroup.Size = New-Object System.Drawing.Size(460, 90)
$inputGroup.BackColor = [System.Drawing.Color]::White
$mainPanel.Controls.Add($inputGroup)

# Computer name label
$computerLabel = New-Object System.Windows.Forms.Label
$computerLabel.Text = "Computer Name:"
$computerLabel.Location = New-Object System.Drawing.Point(20, 25)
$computerLabel.Size = New-Object System.Drawing.Size(120, 20)
$computerLabel.Font = $mainFont
$inputGroup.Controls.Add($computerLabel)

# Computer name textbox
$computerTextBox = New-Object System.Windows.Forms.TextBox
$computerTextBox.Location = New-Object System.Drawing.Point(150, 22)
$computerTextBox.Size = New-Object System.Drawing.Size(200, 26)
$computerTextBox.Font = $mainFont
$inputGroup.Controls.Add($computerTextBox)

# Domain label
$domainLabel = New-Object System.Windows.Forms.Label
$domainLabel.Text = ".yourdomain.com" # Change this to match your domain
$domainLabel.Location = New-Object System.Drawing.Point(355, 25)
$domainLabel.Size = New-Object System.Drawing.Size(90, 20)
$domainLabel.Font = $mainFont
$inputGroup.Controls.Add($domainLabel)

# Get Password button
$getPasswordButton = New-Object System.Windows.Forms.Button
$getPasswordButton.Text = "Get Password"
$getPasswordButton.Location = New-Object System.Drawing.Point(150, 55)
$getPasswordButton.Size = New-Object System.Drawing.Size(120, 28)
$getPasswordButton.Font = $boldFont
$getPasswordButton.BackColor = $primaryColor
$getPasswordButton.ForeColor = [System.Drawing.Color]::White
$getPasswordButton.FlatStyle = "Flat"
$getPasswordButton.FlatAppearance.BorderSize = 0
$inputGroup.Controls.Add($getPasswordButton)

# Results group
$resultsGroup = New-Object System.Windows.Forms.GroupBox
$resultsGroup.Text = "Password"
$resultsGroup.Font = $boldFont
$resultsGroup.ForeColor = $darkGray
$resultsGroup.Location = New-Object System.Drawing.Point(10, 110)
$resultsGroup.Size = New-Object System.Drawing.Size(460, 160)
$resultsGroup.BackColor = [System.Drawing.Color]::White
$mainPanel.Controls.Add($resultsGroup)

# Password display
$passwordTextBox = New-Object System.Windows.Forms.TextBox
$passwordTextBox.Location = New-Object System.Drawing.Point(20, 25)
$passwordTextBox.Size = New-Object System.Drawing.Size(420, 60)
$passwordTextBox.Multiline = $true
$passwordTextBox.ReadOnly = $true
$passwordTextBox.Font = $monoFont
$passwordTextBox.BackColor = $lightGray
$passwordTextBox.BorderStyle = "FixedSingle"
$resultsGroup.Controls.Add($passwordTextBox)

# Copy button
$copyButton = New-Object System.Windows.Forms.Button
$copyButton.Text = "Copy to Clipboard"
$copyButton.Location = New-Object System.Drawing.Point(20, 95)
$copyButton.Size = New-Object System.Drawing.Size(150, 28)
$copyButton.Font = $boldFont
$copyButton.BackColor = $secondaryColor
$copyButton.ForeColor = $darkGray
$copyButton.FlatStyle = "Flat"
$copyButton.FlatAppearance.BorderSize = 0
$resultsGroup.Controls.Add($copyButton)

# Status label
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Location = New-Object System.Drawing.Point(180, 100)
$statusLabel.Size = New-Object System.Drawing.Size(260, 20)
$statusLabel.Font = $mainFont
$resultsGroup.Controls.Add($statusLabel)

# ===== FUNCTIONALITY =====
$getPasswordButton.Add_Click({
    $computerName = $computerTextBox.Text.Trim()

    if ([string]::IsNullOrWhiteSpace($computerName)) {
        $statusLabel.Text = "Please enter a computer name"
        $statusLabel.ForeColor = $errorColor
        return
    }

    try {
        $statusLabel.Text = "Retrieving password..."
        $statusLabel.ForeColor = $darkGray
        $resultsGroup.Refresh()

        # Change 'YOURDOMAIN' to your actual domain name
        $context = New-Object System.DirectoryServices.AccountManagement.PrincipalContext('Domain', 'YOURDOMAIN')
        $computer = [System.DirectoryServices.AccountManagement.ComputerPrincipal]::FindByIdentity($context, $computerName)

        if ($computer -eq $null) {
            $statusLabel.Text = "Computer not found or no LAPS data"
            $statusLabel.ForeColor = $errorColor
            $passwordTextBox.Text = ""
        }
        else {
            $searcher = New-Object System.DirectoryServices.DirectorySearcher
            $searcher.Filter = "(&(objectClass=computer)(sAMAccountName=$computerName`$))"
            $searcher.PropertiesToLoad.Add("ms-Mcs-AdmPwd") | Out-Null
            $result = $searcher.FindOne()

            if ($result -and $result.Properties["ms-Mcs-AdmPwd"]) {
                $password = $result.Properties["ms-Mcs-AdmPwd"][0]
                $passwordTextBox.Text = $password
                $statusLabel.Text = "Password retrieved for $computerName"
                $statusLabel.ForeColor = $primaryColor
            }
            else {
                $passwordTextBox.Text = ""
                $statusLabel.Text = "No LAPS password found for $computerName"
                $statusLabel.ForeColor = $errorColor
            }
        }
    }
    catch {
        $statusLabel.Text = "Error: $_"
        $statusLabel.ForeColor = $errorColor
    }
})

$copyButton.Add_Click({
    $password = $passwordTextBox.Text
    if (-not [string]::IsNullOrWhiteSpace($password)) {
        [System.Windows.Forms.Clipboard]::SetText($password)
        $statusLabel.Text = "Password copied to clipboard!"
        $statusLabel.ForeColor = $primaryColor
        
        # Visual feedback
        $copyButton.BackColor = [System.Drawing.Color]::FromArgb(200, 230, 200)
        $copyButton.Text = "Copied!"
        $resultsGroup.Refresh()
        Start-Sleep -Milliseconds 1000
        $copyButton.BackColor = $secondaryColor
        $copyButton.Text = "Copy to Clipboard"
    } else {
        $statusLabel.Text = "No password to copy"
        $statusLabel.ForeColor = $errorColor
    }
})

# ===== SHOW FORM =====
$form.Add_Shown({$form.Activate()})
[void]$form.ShowDialog()