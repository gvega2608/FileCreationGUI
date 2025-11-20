Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ===========================
# =     CREATE EE FILE      =
# ===========================

$BaseDir = "C:\Users\Gabriela.Vega\OneDrive - WFS\Desktop\BUILDINGS"

$BadgingDir = "C:\Users\Gabriela.Vega\OneDrive - WFS\Desktop\Badging"
$RenewalFiles = @("CBP CL.pdf","CBP Form 3078.pdf","CBP Q37.pdf","CBP COC.pdf", "MDAD.pdf")
$NewIDFiles = @("MDAD.pdf")
$TransferFiles = @("MDAD.pdf","CBP COC.pdf")
$SecondIDFiles = @("CBP CL.pdf","CBP Form 3078.pdf","CBP Q37.pdf","CBP COC.pdf")

# === CREATE FORM === #

$form = New-Object System.Windows.Forms.Form
$form.Text = "Create Employee Folder"
$form.Size = New-Object System.Drawing.Size(500,350)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::FromArgb(240,240,240)
$form.Font = New-Object System.Drawing.Font("Segoe UI",10)

# === TITLE LABEL === #

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = "Create Employee Folder"
$lblTitle.Font = New-Object System.Drawing.Font("Segoe UI",14,[System.Drawing.FontStyle]::Bold)
$lblTitle.AutoSize = $true
$lblTitle.Location = New-Object System.Drawing.Point(120,10)
$form.Controls.Add($lblTitle)

# === FORM LABELS === #

$lblFirst = New-Object System.Windows.Forms.Label
$lblFirst.Text = "First Name:"
$lblFirst.Location = New-Object System.Drawing.Point(20,60)
$lblFirst.AutoSize = $true
$form.Controls.Add($lblFirst)

$txtFirst = New-Object System.Windows.Forms.TextBox
$txtFirst.Location = New-Object System.Drawing.Point(150,60)
$txtFirst.Width = 250
$form.Controls.Add($txtFirst)

$lblLast = New-Object System.Windows.Forms.Label
$lblLast.Text = "Last Name:"
$lblLast.Location = New-Object System.Drawing.Point(20,100)
$lblLast.AutoSize = $true
$form.Controls.Add($lblLast)

$txtLast = New-Object System.Windows.Forms.TextBox
$txtLast.Location = New-Object System.Drawing.Point(150,100)
$txtLast.Width = 250
$form.Controls.Add($txtLast)

$lblBuilding = New-Object System.Windows.Forms.Label
$lblBuilding.Text = "Building Number:"
$lblBuilding.Location = New-Object System.Drawing.Point(20,140)
$lblBuilding.AutoSize = $true
$form.Controls.Add($lblBuilding)

$txtBuilding = New-Object System.Windows.Forms.TextBox
$txtBuilding.Location = New-Object System.Drawing.Point(150,140)
$txtBuilding.Width = 250
$form.Controls.Add($txtBuilding)

$lblType = New-Object System.Windows.Forms.Label
$lblType.Text = "Type:"
$lblType.Location = New-Object System.Drawing.Point(20, 180)
$lblType.AutoSize = $true
$form.Controls.Add($lblType)

$comboType = New-Object System.Windows.Forms.ComboBox
$comboType.Location = New-Object System.Drawing.Point(150,180)
$comboType.Width = 250
$comboType.DropDownStyle = "DropDownList"
$comboType.Items.AddRange(@("R - Renewal","N - New Hire"))
$form.Controls.Add($comboType)

$lblSubType = New-Object System.Windows.Forms.Label
$lblSubType.Text = "Sub-Type:"
$lblSubType.Location = New-Object System.Drawing.Point(20,220)
$lblSubType.AutoSize = $true
$lblSubType.Visible = $false
$form.Controls.Add($lblSubType)

$comboSubType = New-Object System.Windows.Forms.ComboBox
$comboSubType.Location = New-Object System.Drawing.Point(150,220)
$comboSubType.Width = 250
$comboSubType.DropDownStyle = "DropDownList"
$comboSubType.Items.AddRange(@("Transfer","New ID","Second ID/Active FP"))
$comboSubType.Visible = $false
$form.Controls.Add($comboSubType)

# === DYNAMIC RESIZE === #

$comboType.Add_SelectedIndexChanged({
	if ($comboType.SelectedItem -like "N*") {
		$lblSubType.Visible = $true
		$comboSubType.Visible = $true
		$btnCreate.Location = New-Object System.Drawing.Point(150,300)
		$form.Size = New-Object System.Drawing.Size(500,400)
	} else {
		$lblSubType.Visible = $false
		$comboSubType.Visible = $false
		$comboSubType.SelectedIndex = -1
		$btnCreate.Location = New-Object System.Drawing.Point(150,260)
		$form.Size = New-Object System.Drawing.Size(500,350)
	}
})

# === BUTTON === #

$btnCreate = New-Object System.Windows.Forms.Button
$btnCreate.Text = "Create File"
$btnCreate.Location = New-Object System.Drawing.Point(150,260)
$btnCreate.Width = 150
$btnCreate.BackColor = [System.Drawing.Color]::SteelBlue
$btnCreate.ForeColor = [System.Drawing.Color]::White
$btnCreate.Font = New-Object System.Drawing.Font("Segoe UI",10,[System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnCreate)

# === BUTTON ACTION ===#

$btnCreate.Add_Click({
    $FirstName = $txtFirst.Text.Trim()
    $FirstName = $FirstName.Substring(0,1).ToUpper() + $FirstName.Substring(1).ToLower()
    $LastName = $txtLast.Text.Trim()
    $LastName = $LastName.ToUpper()
    $BuildingNumber = $txtBuilding.Text.Trim()

    if ($FirstName -and $LastName -and $BuildingNumber) {
        $BuildingPath = Join-Path $BaseDir $BuildingNumber

        if (-Not (Test-Path $BuildingPath)) {
            [System.Windows.Forms.MessageBox]::Show("Building folder '$BuildingNumber' does not exist.","Error","OK","Error")
        } else {
            $EmployeeFolderName = "$LastName, $FirstName"
            $EmployeeFolderPath = Join-Path $BuildingPath $EmployeeFolderName

            if (-Not (Test-Path $EmployeeFolderPath)) {
                New-Item -ItemType Directory -Path $EmployeeFolderPath | Out-Null
                [System.Windows.Forms.MessageBox]::Show("Folder created: $EmployeeFolderPath","Success","OK","Information")

                ##OPEN FILES##
                $selectedType = $comboType.SelectedItem
                if ($selectedType -like "R*") {
                    $filesToCopy = $RenewalFiles
                } elseif ($selectedType -like "N*") {
                    switch ($comboSubType.SelectedItem) {
						"Transfer" {$filesToCopy = $TransferFiles}
						"New ID" {$filesToCopy = $NewIDFiles}
						"Second ID/Active FP" {$filesToCopy = $SecondIDFiles}
						default {$filesToCopy = @()}
					}
                } else {
                    $filesToCopy = @()
                }

                $specialCasesSecondWord = @("CBP CL","CBP Q37","CBP COC")
                $specialCaseFirstWord = "CBP Form 3078"

                foreach ($file in $filesToCopy) {
                    $sourcePath = Join-Path $BadgingDir $file
                    if (Test-Path $sourcePath) {
                        $baseName = $file.Split('.')[0]
                        $extension = $file.Split('.')[1]

                        if ($baseName -eq $specialCaseFirstWord) {
                            $tag = $baseName.Split(' ')[0]
                        } elseif ($specialCasesSecondWord -contains $baseName) {
                            $tag = $baseName.Split(' ')[-1]
                        } else {
                            $tag = $baseName
                        }

                        $newFileName = "$LastName, $FirstName [$tag].$extension"
                        $destPath = Join-Path $EmployeeFolderPath $newFileName

                        Copy-Item $sourcePath $destPath
                        Start-Process $destPath
                    }
                }

                ##CLEAR FORM##
                $txtFirst.Text = ""
                $txtLast.Text = ""
                $txtBuilding.Text = ""
                $comboType.SelectedIndex = -1

            } else {
                [System.Windows.Forms.MessageBox]::Show("Folder already exists: $EmployeeFolderPath","Warning","OK","Warning")
            } 
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("Please fill in all fields.","Warning","OK","Warning")
    }
})

$form.AcceptButton = $btnCreate

# === SHOW FORM === #
$form.ShowDialog()
