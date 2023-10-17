    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true)]
        $CI_ID
    )
 
function Show-CMReportForm_psf {
 
    #----------------------------------------------
    #region Import the Assemblies
    #----------------------------------------------
    [void][reflection.assembly]::Load('System.Drawing, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
    [void][reflection.assembly]::Load('System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
    #endregion Import Assemblies
 
    #----------------------------------------------
    #region Generated Form Objects
    #----------------------------------------------
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $formTex = New-Object 'System.Windows.Forms.Form'
    $labelHostname = New-Object 'System.Windows.Forms.Label'
    $richtextbox1 = New-Object 'System.Windows.Forms.RichTextBox'
    $buttonFind = New-Object 'System.Windows.Forms.Button'
    $textboxFind = New-Object 'System.Windows.Forms.TextBox'
    $buttonCopy = New-Object 'System.Windows.Forms.Button'
    $buttonExit = New-Object 'System.Windows.Forms.Button'
    $buttonLoad = New-Object 'System.Windows.Forms.Button'
    $InitialFormWindowState = New-Object 'System.Windows.Forms.FormWindowState'
    #endregion Generated Form Objects
 
    #----------------------------------------------
    # User Generated Script
    #----------------------------------------------
 
# -------------------------------------------------------------------------------------------------------------
    # Manually Declare variables
 
    # Temp file that stores report
    $CSVOutputFile = 'c:\scripts\CMUpdateReport.csv'
    # Add your report server ex: http://SCCMServer/ReportServer
    $ReportServerURL = "http://SCCMServer/ReportServer"
    # Add your report path. Ex: /ConfigMgr_ABC/Software Updates - A Compliance/Compliance 8 - Computers in a specific compliance state for an update (secondary)
    $ReportPath = "/ConfigMgr_ABC/Software Updates - A Compliance/Compliance 8 - Computers in a specific compliance state for an update (secondary)"
    # Identity collection you want the report ran against. SMS00001 is the 'All Systems' collection
    $CMCollectionID = "SMS00001"
# -------------------------------------------------------------------------------------------------------------
     
    import-module ($Env:SMS_ADMIN_UI_PATH.Substring(0, $Env:SMS_ADMIN_UI_PATH.Length - 5) + '\ConfigurationManager.psd1')
    $Drive = Get-PSDrive -PSProvider CMSite
    CD "$($Drive):"
     
    $UpdateSource = Get-cmsoftwareupdate -ID $CI_ID -fast
    $ReqdUpdate = ($UpdateSource).CI_UniqueID
    [string]$CMUpdate = ($UpdateSource).LocalizedDisplayName
 
    # Remove old temp file if it exists
    if (test-path $CSVOutputFile) { Remove-Item –path $CSVOutputFile }  
     
    #region FindFunction
    function FindText
    {   
        if($textboxFind.Text.Length -eq 0)
        {
            return
        }
         
        $index = $richtextbox1.Find($textboxFind.Text,$richtextbox1.SelectionStart+ $richtextbox1.SelectedText.Length,[System.Windows.Forms.RichTextBoxFinds]::None)
        if($index -ge 0)
        {   
            $richtextbox1.Select($index,$textboxFind.Text.Length)
            $richtextbox1.ScrollToCaret()
            #$richtextbox1.Focus()
        }
        else
        {
            $index = $richtextbox1.Find($textboxFind.Text,0,$richtextbox1.SelectionStart,[System.Windows.Forms.RichTextBoxFinds]::None)
            #
            if($index -ge 0)
            {   
                $richtextbox1.Select($index,$textboxFind.Text.Length)
                $richtextbox1.ScrollToCaret()
                #$richtextbox1.Focus()
            }
            else
            {
                $richtextbox1.SelectionStart = 0    
            }
        }
         
    }
    #endregion
     
    $formTex_Load={
        #TODO: Initialize Form Controls here
         
    }
     
    $buttonExit_Click={
        Remove-Item –path $CSVOutputFile
        $formTex.Close()
    }
     
    $buttonLoad_Click={
        Update-Text
    }
     
    $buttonCopy_Click={
        $richtextbox1.SelectAll() #Select all the text
        $richtextbox1.Copy()    #Copy selected text to clipboard
        $richtextbox1.Select(0,0); #Unselect all the text
    }
     
    $textboxFind_TextChanged={
        $buttonFind.Enabled = $textboxFind.Text.Length -gt 0
    }
     
    $buttonFind_Click={
        FindText
    }
     
    #################################################
    # Customize LoadText Function
    #################################################
     
    function Update-Text
    {
     
        ## Load report viewer assemblies
        Add-Type -AssemblyName "Microsoft.ReportViewer.WinForms, Version=12.0.0.0, Culture=neutral, PublicKeyToken=89845dcd8080cc91"
         
        ## Create a ReportViewer object
         
        $rv = New-Object Microsoft.Reporting.WinForms.ReportViewer
         
        $rv.ServerReport.ReportServerUrl = $ReportServerURL
        $rv.ServerReport.ReportPath = $ReportPath
        $rv.ProcessingMode = "Remote"
         
        $inputParams = @{
            "CollID"   = $CMCollectionID;
            "UpdateID" = $ReqdUpdate;
            "Status"   = "Update is required"
        }
         
        #create an array based on how many incoming parameters 
        $params = New-Object 'Microsoft.Reporting.WinForms.ReportParameter[]' $inputParams.Count
         
        $i = 0
        foreach ($p in $inputParams.GetEnumerator())
        {
            $params[$i] = New-Object Microsoft.Reporting.WinForms.ReportParameter($p.Name, $p.Value, $false)
            $i++
        }
         
        $rv.ServerReport.SetParameters($params)
        # These variables are used for remdering PDF's. I left them in anyways.
        $mimeType = $null
        $encoding = $null
        $extension = $null
        $streamids = $null
        $warnings = $null
         
        $fileName = '$CSVOutputFile'
        $fileStream = New-Object System.IO.FileStream($fileName, [System.IO.FileMode]::OpenOrCreate)
        $fileStream.Write($bytes, 0, $bytes.Length)
        $fileStream.Close()
         
         
        # render the SSRS report in CSV 
        $bytes = $null
        $bytes = $rv.ServerReport.Render("CSV",
            $null,
            [ref]$mimeType,
            [ref]$encoding,
            [ref]$extension,
            [ref]$streamids,
            [ref]$warnings)
         
     
        # save the report to a file
        $fileStream = New-Object System.IO.FileStream($CSVOutputFile, [System.IO.FileMode]::OpenOrCreate)
        $fileStream.Write($bytes, 0, $bytes.Length)
        $fileStream.Close()
         
        # Re-import file and remove first nine lines
         
        get-content -LiteralPath $CSVOutputFile |
        select -Skip 9 |
        set-content "$CSVOutputFile-temp"
        move "$CSVOutputFile-temp" $CSVOutputFile -Force
         
        ###################################################
         
     
        $Finalvalues = Import-CSV -LiteralPath $CSVOutputFile | select -ExpandProperty Details_Table0_ComputerName0
        $richtextbox1.Text = $Finalvalues | Format-Table | out-string
             
    }
     
     
     
    # --End User Generated Script--
    #----------------------------------------------
    #region Generated Events
    #----------------------------------------------
     
    $Form_StateCorrection_Load=
    {
        #Correct the initial state of the form to prevent the .Net maximized form issue
        $formTex.WindowState = $InitialFormWindowState
    }
     
    $Form_Cleanup_FormClosed=
    {
        #Remove all event handlers from the controls
        try
        {
            $buttonFind.remove_Click($buttonFind_Click)
            $textboxFind.remove_TextChanged($textboxFind_TextChanged)
            $buttonCopy.remove_Click($buttonCopy_Click)
            $buttonExit.remove_Click($buttonExit_Click)
            $buttonLoad.remove_Click($buttonLoad_Click)
            $formTex.remove_Load($formTex_Load)
            $formTex.remove_Load($Form_StateCorrection_Load)
            $formTex.remove_FormClosed($Form_Cleanup_FormClosed)
        }
        catch { Out-Null <# Prevent PSScriptAnalyzer warning #> }
    }
    #endregion Generated Events
 
    #----------------------------------------------
    #region Generated Form Code
    #----------------------------------------------
    $formTex.SuspendLayout()
    #
    # formTex
    #
 
 
 
 
    $formTex.Controls.Add($labelHostname)
    $formTex.Controls.Add($richtextbox1)
    $formTex.Controls.Add($buttonFind)
    $formTex.Controls.Add($textboxFind)
    $formTex.Controls.Add($buttonCopy)
    $formTex.Controls.Add($buttonExit)
    $formTex.Controls.Add($buttonLoad)
    $formTex.AcceptButton = $buttonFind
    $formTex.AutoScaleDimensions = '6, 13'
    $formTex.AutoScaleMode = 'Font'
    $formTex.ClientSize = '584, 362'
    $formTex.Name = 'formTex'
    $formTex.StartPosition = 'CenterScreen'
    $formTex.Text = "Devices Requiring $CMUpdate"
    $formTex.AutoSize = $True
    $formTex.add_Load($formTex_Load)
    #
    # labelHostname
    #
    $labelHostname.AutoSize = $True
    $labelHostname.Location = '12, 16'
    $labelHostname.Name = 'labelHostname'
    $labelHostname.Size = '56, 17'
    $labelHostname.TabIndex = 7
    $labelHostname.Text = 'Hostname'
    $labelHostname.UseCompatibleTextRendering = $True
    #
    # richtextbox1
    #
    $richtextbox1.Anchor = 'Top, Bottom, Left, Right'
    $richtextbox1.BackColor = 'Window'
    $richtextbox1.Font = 'Courier New, 8.25pt'
    $richtextbox1.HideSelection = $False
    $richtextbox1.Location = '12, 36'
    $richtextbox1.Name = 'richtextbox1'
    $richtextbox1.ReadOnly = $True
    $richtextbox1.RightToLeft = 'No'
    $richtextbox1.Size = '559, 281'
    $richtextbox1.TabIndex = 6
    $richtextbox1.Text = ''
    $richtextbox1.WordWrap = $False
    #
    # buttonFind
    #
    $buttonFind.Anchor = 'Top, Right'
    $buttonFind.Enabled = $False
    $buttonFind.Location = '536, 8'
    $buttonFind.Name = 'buttonFind'
    $buttonFind.Size = '36, 23'
    $buttonFind.TabIndex = 5
    $buttonFind.Text = '&amp;Find'
    $buttonFind.UseCompatibleTextRendering = $True
    $buttonFind.UseVisualStyleBackColor = $True
    $buttonFind.add_Click($buttonFind_Click)
    #
    # textboxFind
    #
    $textboxFind.Anchor = 'Top, Right'
    $textboxFind.Location = '339, 10'
    $textboxFind.Name = 'textboxFind'
    $textboxFind.Size = '191, 20'
    $textboxFind.TabIndex = 4
    $textboxFind.add_TextChanged($textboxFind_TextChanged)
    #
    # buttonCopy
    #
    $buttonCopy.Anchor = 'Bottom'
    $buttonCopy.Location = '255, 327'
    $buttonCopy.Name = 'buttonCopy'
    $buttonCopy.Size = '75, 23'
    $buttonCopy.TabIndex = 3
    $buttonCopy.Text = '&amp;Copy'
    $buttonCopy.UseCompatibleTextRendering = $True
    $buttonCopy.UseVisualStyleBackColor = $True
    $buttonCopy.add_Click($buttonCopy_Click)
    #
    # buttonExit
    #
    $buttonExit.Anchor = 'Bottom, Right'
    $buttonExit.Location = '501, 327'
    $buttonExit.Name = 'buttonExit'
    $buttonExit.Size = '75, 23'
    $buttonExit.TabIndex = 2
    $buttonExit.Text = 'E&amp;xit'
    $buttonExit.UseCompatibleTextRendering = $True
    $buttonExit.UseVisualStyleBackColor = $True
    $buttonExit.add_Click($buttonExit_Click)
    #
    # buttonLoad
    #
    $buttonLoad.Anchor = 'Bottom, Left'
    $buttonLoad.Location = '12, 327'
    $buttonLoad.Name = 'buttonLoad'
    $buttonLoad.Size = '75, 23'
    $buttonLoad.TabIndex = 1
    $buttonLoad.Text = '&amp;Load'
    $buttonLoad.UseCompatibleTextRendering = $True
    $buttonLoad.UseVisualStyleBackColor = $True
    $buttonLoad.add_Click($buttonLoad_Click)
    $formTex.ResumeLayout()
    #endregion Generated Form Code
 
    #----------------------------------------------
 
    #Save the initial state of the form
    $InitialFormWindowState = $formTex.WindowState
    #Init the OnLoad event to correct the initial state of the form
    $formTex.add_Load($Form_StateCorrection_Load)
    #Clean up the control events
    $formTex.add_FormClosed($Form_Cleanup_FormClosed)
    #Show the Form
    return $formTex.ShowDialog()
 
} #End Function
 
#Call the form
Show-CMReportForm_psf | Out-Null
