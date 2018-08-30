Function Add-Properties {
param ([Parameter(ValueFromPipeline=$true)]$ishObject)
  
  foreach($ishField in $ishObject.IshField)
  { 
    if ($ishField.ValueType -ne "Element" -and $ishField.Value -ne ""){
    $ishObject = $ishObject | Add-Member -MemberType NoteProperty -Name $ishField.Name -Value $ishField.Value -PassThru -Force }
    elseif ($ishField.ValueType -eq "Element"){
        $ishObject = $ishObject | Add-Member -MemberType NoteProperty -Name "$($ishField.Name)ELEMENT" -Value $ishField.Value -PassThru -Force }
    }
  Write-Output $ishObject
}

Function Get-Selected-Outputs {
    param($outputs)
    $selectedObjects = $outputs | Select-Object -Property FISHOUTPUTFORMATREF,FISHPUBLNGCOMBINATION,FGARMINPARTNUMBER,FISHPUBSTATUS,MODIFIED-ON,FTITLE,VERSION,IshRef,FISHOUTPUTFORMATREFELEMENT | Out-GridView -OutputMode Multiple -Title "Select one or more outputs"
    return $selectedObjects
}

Function Publish-Outputs {
    param($outputs, [boolean]$delay)
    $outputCount = 0
    $total = $outputs.count
    Write-Host "****************************`nPublishing $total output(s).`n****************************`n"
    $outputs | ForEach-Object {
        If($delay -and ++$outputCount -gt 3){
            Write-Host "Waiting 1 minute to start next output..."
            Start-Sleep -Seconds 60}

        Write-Host "$(get-date -Format 'M/d/yyyy h:mm:ss tt'): Started publishing $($_.FISHOUTPUTFORMATREF) $($_.FISHPUBLNGCOMBINATION) for $($_.FTITLE) v.$($_.VERSION)"
        Publish-IshPublicationOutput -IshSession $ishSession -LogicalId $_.IshRef -Version $_.VERSION -LanguageCombination $_.FISHPUBLNGCOMBINATION -OutputFormat $_.FISHOUTPUTFORMATREFELEMENT        
    }
}

#Install ISHRemote module if needed.
If (-not (Get-Module -ListAvailable -Name "ISHRemote")) {
    Install-Module ISHRemote -Repository PSGallery -Scope CurrentUser -Force
}

Function Get-ISHLogin {
    param($messageText)
    $credential = Get-Credential -Message $messageText
    if ($credential) {$testSession = Test-IshSession -WsBaseUrl https://garmindevaws01.sdlproducts.com/ISHWS/ -PSCredential $credential
    }else{Exit}
    if ($testSession) {return $credential
    }else{Get-ISHLogin -messageText "Username or password was incorrect. Re-enter your username and password for the Garmin DEV CMS Server"}
}

#Create session to connect to CMS server
$IshCredential=Get-ISHLogin -messageText "Enter your username and password for the Garmin DEV CMS Server"
Write-Host "Connecting to Garmin DEV CMS Server...`n"
$ishSession = New-IshSession -WsBaseUrl https://garmindevaws01.sdlproducts.com/ISHWS/ -PSCredential $IshCredential
Write-Host "Successfully connected.`n"

$metadata = Set-IshRequestedMetadataField -IshSession $ishSession -Name 'FISHPUBLICATIONTYPE' -Level "Logical" |
            Set-IshRequestedMetadataField -IshSession $ishSession -Name 'FTITLE' -Level "Logical" |
            Set-IshRequestedMetadataField -IshSession $ishSession -Name 'VERSION' -Level "Version" |
            Set-IshRequestedMetadataField -IshSession $ishSession -Name 'MODIFIED-ON' -Level "Version" |
            Set-IshRequestedMetadataField -IshSession $ishSession -Name 'FGARMINPARTNUMBER' -Level "Lng" |
            Set-IshRequestedMetadataField -IshSession $ishSession -Name 'FISHPUBSTATUS' -Level "Lng" 
 
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
$Form1 = New-Object System.Windows.Forms.Form
$Form1.Text = "Publish Outputs"
$Form1.Size = New-Object System.Drawing.Size(400,250)
$Form1.StartPosition = "CenterScreen"
$Form1.Topmost = $True 

# Icon
$Form1.Icon = [Drawing.Icon]::ExtractAssociatedIcon((Get-Command powershell).Path)
 
$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Point(125,170)
$OKButton.Size = New-Object System.Drawing.Size(75,23)
$OKButton.Text = "OK"
$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$Form1.AcceptButton = $OKButton
$Form1.Controls.Add($OKButton)
 
$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Point(200,170)
$CancelButton.Size = New-Object System.Drawing.Size(75,23)
$CancelButton.Text = "Cancel"
$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$Form1.CancelButton = $CancelButton
$Form1.Controls.Add($CancelButton)
 
$LabelGUID = New-Object System.Windows.Forms.Label
$LabelGUID.Location = New-Object System.Drawing.Point(10,20)
$LabelGUID.Size = New-Object System.Drawing.Size(260,20)
$LabelGUID.Text = "Publication GUID"
$Form1.Controls.Add($LabelGUID)
 
$textboxGUID = New-Object System.Windows.Forms.TextBox
$textboxGUID.Location = New-Object System.Drawing.Point(10,40)
$textboxGUID.Size = New-Object System.Drawing.Size(260,20)
$Form1.Controls.Add($textboxGUID)

$LabelVersion = New-Object System.Windows.Forms.Label
$LabelVersion.Location = New-Object System.Drawing.Point(280,20)
$LabelVersion.Size = New-Object System.Drawing.Size(50,20)
$LabelVersion.Text = "Version"
$Form1.Controls.Add($LabelVersion)
 
$textBoxVersion = New-Object System.Windows.Forms.TextBox
$textBoxVersion.Location = New-Object System.Drawing.Point(280,40)
$textBoxVersion.Size = New-Object System.Drawing.Size(50,20)
$Form1.Controls.Add($textBoxVersion)

$groupboxSelectOutputs = New-Object System.Windows.Forms.GroupBox
$groupboxSelectOutputs.Location = New-Object System.Drawing.Point(10,70)
$groupboxSelectOutputs.Size = New-Object System.Drawing.Size(210, 70)
$groupboxSelectOutputs.Text = "Select Outputs"

$radioButtonPublishAll = New-Object System.Windows.Forms.RadioButton
$radioButtonPublishAll.Location = New-Object System.Drawing.Point (10,20)
$radioButtonPublishAll.Size = New-Object System.Drawing.Size(190,20)
$radioButtonPublishAll.Checked = $false
$radioButtonPublishAll.Text = "Publish All Web Outputs"

$radioButtonPublishSelect = New-Object System.Windows.Forms.RadioButton
$radioButtonPublishSelect.Location = New-Object System.Drawing.Point (10,40)
$radioButtonPublishSelect.Size = New-Object System.Drawing.Size(190,20)
$radioButtonPublishSelect.Checked = $true
$radioButtonPublishSelect.Text = "Select Outputs to Publish"

$groupboxSelectOutputs.Controls.Add($radioButtonPublishAll)
$groupboxSelectOutputs.Controls.Add($radioButtonPublishSelect)
$Form1.Controls.Add($groupboxSelectOutputs)

$LabelDelay = New-Object System.Windows.Forms.Label
$LabelDelay.Location = New-Object System.Drawing.Point(230,100)
$LabelDelay.Size = New-Object System.Drawing.Size(160,40)
$LabelDelay.Text = "Adds a delay between outputs to reduce publishing queue backlog"
$Form1.Controls.Add($LabelDelay)

$checkboxDelay = New-Object System.Windows.Forms.CheckBox
$checkboxDelay.Location = New-Object System.Drawing.Point (230,80)
$checkboxDelay.Size = New-Object System.Drawing.Size (70,20)
$checkboxDelay.Checked = $true
$checkboxDelay.Text = "Delay"
$Form1.Controls.Add($checkboxDelay)

$LoadDialog = {
$Form1.Add_Shown({$textBoxGUID.Select()})
$result = $Form1.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK)
{
    $pubGUID = $textBoxGUID.Text
    $pubVersion = $textBoxVersion.Text

    if ($radioButtonPublishSelect.Checked){
        $publishSelect = $true}
    else {$publishSelect = $false}

    if ($checkboxDelay.Checked){
        $publishDelay = $true}
    else {$publishDelay = $false}


    if ($pubGUID -eq "" -Or $pubVersion -eq "") {
        [System.Windows.Forms.MessageBox]::Show("Please Enter All Required Fields", "Error")
        & $LoadDialog
         }
    else {
        $metadataFilterRetrieve = Set-IshMetadataFilterField -IshSession $ishSession -Name 'VERSION' -Level 'Version' -ValueType "Value" -FilterOperator "Equal" -Value $pubVersion
        if (-not $publishSelect){$metadataFilterRetrieve = $metadataFilterRetrieve | Set-IshMetadataFilterField -IshSession $ishSession -Name 'FISHOUTPUTFORMATREF' -Level 'Lng' -ValueType "Element" -FilterOperator "In" -Value 'VOUTPUTFORMATGARMINL4AAWEBHELP'}
        $pubObjects = Get-IshPublicationOutput -IshSession $ishSession -LogicalId @($pubGUID) -RequestedMetadata $metadata -MetadataFilter $metadataFilterRetrieve | ForEach-Object { Add-Properties $_ }
        if ($publishSelect) {
            $selectedOutputs = Get-Selected-Outputs $pubObjects
        }
        else {$selectedOutputs = $pubObjects}
        }
    
    $startTime = $(get-date)
    $startTimeString = $startTime.ToString('dd/MM/yyyy HH:mm:ss')
    $publishResult = Publish-Outputs -outputs $selectedOutputs -delay $publishDelay
    $eventMetadata = Set-IshRequestedMetadataField -IshSession $ishSession -Name "EVENTID" -Level "Progress" |
                     Set-IshRequestedMetadataField -IshSession $ishSession -Name "DESCRIPTION" -Level "Progress" |
                     Set-IshRequestedMetadataField -IshSession $ishSession -Name "STATUS" -Level "Progress"

    $publishBusy = Get-IshEvent -IshSession $ishSession -EventTypes 'EXPORTFORPUBLICATION' -ProgressStatusFilter Busy -UserFilter Current -ModifiedSince $startTime -RequestedMetadata $eventMetadata 
    While ($null -ne $publishBusy) {
        Start-Sleep -Seconds 30
        $publishBusy = Get-IshEvent -IshSession $ishSession -EventTypes 'EXPORTFORPUBLICATION' -ProgressStatusFilter Busy -UserFilter Current -ModifiedSince $startTime -RequestedMetadata $eventMetadata 
    }
    $global:publishFailed = Get-IshEvent -IshSession $ishSession -EventTypes 'EXPORTFORPUBLICATION' -ProgressStatusFilter Failed -UserFilter Current -ModifiedSince $startTime -RequestedMetadata $eventMetadata   

    $LoadFailedDialog = {
        
        $totalFailed = $global:publishFailed.count

        $global:selectedFailedOutputs = @()
        $metadata = Set-IshRequestedMetadataField -IshSession $ishSession -Name 'FISHPUBLICATIONTYPE' -Level "Logical" |
            Set-IshRequestedMetadataField -IshSession $ishSession -Name 'FTITLE' -Level "Logical" |
            Set-IshRequestedMetadataField -IshSession $ishSession -Name 'VERSION' -Level "Version" |
            Set-IshRequestedMetadataField -IshSession $ishSession -Name 'MODIFIED-ON' -Level "Version" |
            Set-IshRequestedMetadataField -IshSession $ishSession -Name 'FGARMINPARTNUMBER' -Level "Lng" |
            Set-IshRequestedMetadataField -IshSession $ishSession -Name 'FISHPUBSTATUS' -Level "Lng" 

        $FormFailed = New-Object System.Windows.Forms.Form
        $FormFailed.Text = "$totalFailed Outputs FAILED"
        $FormFailed.Size = New-Object System.Drawing.Size(980,500)
        $FormFailed.StartPosition = "CenterScreen"

        # Icon
        $FormFailed.Icon = [Drawing.Icon]::ExtractAssociatedIcon((Get-Command powershell).Path)
 
        $RepublishButton = New-Object System.Windows.Forms.Button
        $RepublishButton.Location = New-Object System.Drawing.Point(880,420)
        $RepublishButton.Size = New-Object System.Drawing.Size(75,23)
        $RepublishButton.Text = "Republish"
        $RepublishButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $FormFailed.AcceptButton = $RepublishButton
        $FormFailed.Controls.Add($RepublishButton)
 
        $SkipButton = New-Object System.Windows.Forms.Button
        $SkipButton.Location = New-Object System.Drawing.Point(800,420)
        $SkipButton.Size = New-Object System.Drawing.Size(75,23)
        $SkipButton.Text = "Skip"
        $SkipButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $FormFailed.CancelButton = $SkipButton
        $FormFailed.Controls.Add($SkipButton)
 
        $LabelFailedDesc = New-Object System.Windows.Forms.Label
        $LabelFailedDesc.Location = New-Object System.Drawing.Point(10,10)
        $LabelFailedDesc.Size = New-Object System.Drawing.Size(940,60)
        $LabelFailedDesc.Text = "$totalFailed outputs completed with a status of Failed.`n`nTo republish outputs, select the failed outputs from the list, and click Republish.`nTo skip republishing outputs, click Skip."
        $FormFailed.Controls.Add($LabelFailedDesc)
        
        $FailedGridView = New-Object System.Windows.Forms.DataGridView
        $FailedGridView.Location = New-Object System.Drawing.Point(10,80)
        $FailedGridView.Size=New-Object System.Drawing.Size(940,330)
        $FailedGridView.SelectionMode = 'FullRowSelect'
        $FailedGridView.MultiSelect = $true
       

        $failedFilterRetrieve = Set-IshMetadataFilterField -IshSession $ishSession -Name 'VERSION' -Level 'Version' -ValueType "Value" -FilterOperator "Equal" -Value $pubVersion |
                                Set-IshMetadataFilterField -IshSession $ishSession -Name 'FISHPUBSTATUS' -Level 'Lng' -ValueType "Value" -FilterOperator "Equal" -Value 'Failed' |
                                Set-IshMetadataFilterField -IshSession $ishSession -Name 'MODIFIED-ON' -Level 'Lng' -ValueType "Value" -FilterOperator "greaterthanorequal" -Value $startTimeString
                                
        $failedObjects = Get-IshPublicationOutput -IshSession $ishSession -LogicalId @($pubGUID) -RequestedMetadata $metadata -MetadataFilter $failedFilterRetrieve | ForEach-Object { Add-Properties $_ }
        $failedSource = $failedObjects | Select-Object -Property FISHOUTPUTFORMATREF,FISHPUBLNGCOMBINATION,FGARMINPARTNUMBER,FISHPUBSTATUS,MODIFIED-ON,FTITLE,VERSION,IshRef,FISHOUTPUTFORMATREFELEMENT
        $failedarraylist = New-Object System.Collections.ArrayList
        $failedarraylist.AddRange(($failedSource))
        
        $FailedGridView.DataSource = $failedarraylist

        $FormFailed.Controls.Add($FailedGridView)
        $FormFailed.Topmost = $True 
        $failedResult = $FormFailed.ShowDialog()

        if ($failedResult -eq [System.Windows.Forms.DialogResult]::OK) {
            $FailedGridView.SelectedRows | ForEach-Object {
                $global:selectedFailedOutputs += [pscustomobject]@{
                    FISHOUTPUTFORMATREF = $FailedGridView.Rows[$_.Index].Cells[0].Value
                    FISHPUBLNGCOMBINATION = $FailedGridView.Rows[$_.Index].Cells[1].Value
                    FGARMINPARTNUMBER = $FailedGridView.Rows[$_.Index].Cells[2].Value
                    FISHPUBSTATUS = $FailedGridView.Rows[$_.Index].Cells[3].Value
                    FTITLE = $FailedGridView.Rows[$_.Index].Cells[5].Value
                    VERSION = $FailedGridView.Rows[$_.Index].Cells[6].Value
                    IshRef = $FailedGridView.Rows[$_.Index].Cells[7].Value
                    FISHOUTPUTFORMATREFELEMENT = $FailedGridView.Rows[$_.Index].Cells[8].Value                
                }
            }
        }

        $failedstartTime = $(get-date)
        $publishResult = Publish-Outputs -outputs $selectedFailedOutputs -delay $publishDelay
        $publishBusy = Get-IshEvent -IshSession $ishSession -EventTypes 'EXPORTFORPUBLICATION' -ProgressStatusFilter Busy -UserFilter Current -ModifiedSince $failedstartTime -RequestedMetadata $eventMetadata 
        While ($null -ne $publishBusy) {
            Start-Sleep -Seconds 30
            $publishBusy = Get-IshEvent -IshSession $ishSession -EventTypes 'EXPORTFORPUBLICATION' -ProgressStatusFilter Busy -UserFilter Current -ModifiedSince $failedstartTime -RequestedMetadata $eventMetadata 
        }
        $global:publishFailed = Get-IshEvent -IshSession $ishSession -EventTypes 'EXPORTFORPUBLICATION' -ProgressStatusFilter Failed -UserFilter Current -ModifiedSince $failedstartTime -RequestedMetadata $eventMetadata   
    }
    If ($null -ne $global:publishFailed){& $LoadFailedDialog}

    $publishStatus = Get-IshEvent -IshSession $ishSession -EventTypes 'EXPORTFORPUBLICATION' -ProgressStatusFilter All -UserFilter Current -ModifiedSince $startTime -RequestedMetadata $eventMetadata
    $failedOutputs = $global:publishFailed.count
    $successfulOutputs = $publishStatus.count - $failedOutputs
    $startTimeString = $startTime.ToString('dd/MM/yyyy HH:mm:ss')
    $metadataFilterRetrieve = Set-IshMetadataFilterField -IshSession $ishSession -Name 'VERSION' -Level 'Version' -ValueType "Value" -FilterOperator "Equal" -Value $pubVersion |
                              Set-IshMetadataFilterField -IshSession $ishSession -Name 'MODIFIED-ON' -Level 'Lng' -ValueType "Value" -FilterOperator "greaterthanorequal" -Value $startTimeString

    $pubStatusFinal = Get-IshPublicationOutput -IshSession $ishSession -LogicalId @($pubGUID) -RequestedMetadata $metadata -MetadataFilter $metadataFilterRetrieve | ForEach-Object { Add-Properties $_ }
    $pubStatusFinal | Select-Object -Property FISHOUTPUTFORMATREF,FISHPUBLNGCOMBINATION,FGARMINPARTNUMBER,FISHPUBSTATUS,MODIFIED-ON,FTITLE,VERSION | Out-GridView -Wait -Title "Completed Outputs: $failedOutputs FAILED, $successfulOutputs SUCCESSFUL"
    
    Exit
}elseif ($result -eq [System.Windows.Forms.DialogResult]::Cancel){
    Exit
}

}
& $LoadDialog

Exit