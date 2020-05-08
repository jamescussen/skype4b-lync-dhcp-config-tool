########################################################################
# Name: Skype4B / Lync DHCP Config Tool
# Version: v1.0.1 (30/11/2015)
# Original Release Date: 29/11/2015
# Created By: James Cussen
# Web Site: http://www.myskypelab.com
# Notes: For more information on the requirements for setting up and using this tool please visit http://www.myskypelab.com
#
# Copyright: Copyright (c) 2015, James Cussen (www.myskypelab.com) All rights reserved.
# Licence: 	Redistribution and use of script, source and binary forms, with or without modification, are permitted provided that the following conditions are met:
#				1) Redistributions of script code must retain the above copyright notice, this list of conditions and the following disclaimer.
#				2) Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.
#				3) Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution.
#				4) This license does not include any resale or commercial use of this software.
#				5) Any portion of this software may not be reproduced, duplicated, copied, sold, resold, or otherwise exploited for any commercial purpose without express written consent of James Cussen.
#			THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; LOSS OF GOODWILL OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
#
#
# Pre-requisistes:
#		- You must run this tool on a Windows DHCP Server to edit Windows server settings. However, you can run the Export Cisco configuration function on any Windows PC with Powershell v2+.
#
# Known Issues: 
#		- None.
#
# Release Notes:
# 1.00 Initial Release.
#	- The script will display, upload and remove Options 120 and Options 1,2,3,4,5 of MS-UC-Client Vendor Class.
#	- Generate Cisco IOS DHCP options commands for configuring Option 120, and Vendor Class Options on a switch or router.
#
# 1.01 Bug Fix
#	- The netsh command sometimes doesn't report server scopes information properly (no header for the standard options section). Implemented work around for this inconsistent output. (Thanks Greig for reporting!)
#
########################################################################


$theVersion = $PSVersionTable.PSVersion
$MajorVersion = $theVersion.Major

$CiscoExportAvailable = $true

Write-Host ""
Write-Host "--------------------------------------------------------------"
Write-Host "Powershell Version Check..." -foreground "yellow"
if($MajorVersion -eq  "1")
{
	Write-Host "This machine only has Version 1 Powershell installed.  This version of Powershell is not supported." -foreground "red"
}
elseif($MajorVersion -eq  "2")
{
	Write-Host "This machine has version 2 Powershell installed. CHECK PASSED!" -foreground "green"
	Write-Host "INFO: Cisco Export as file is not available on Powershell V2. Output will be displayed in the Powershell window." -foreground "yellow"
	$CiscoExportAvailable = $false
}
elseif($MajorVersion -eq  "3")
{
	Write-Host "This machine has version 3 Powershell installed. CHECK PASSED!" -foreground "green"
}
elseif($MajorVersion -eq  "4")
{
	Write-Host "This machine has version 4 Powershell installed. CHECK PASSED!" -foreground "green"
}
else
{
	Write-Host "This machine has version $MajorVersion Powershell installed. Unknown level of support for this version." -foreground "yellow"
}
Write-Host "--------------------------------------------------------------"
Write-Host ""


if(Get-Command "netsh" -errorAction SilentlyContinue)
{
	Write-Host "PRE REQ CHECK - netsh is available." -foreground "green"
}
else
{
	Write-Host "ERROR: The netsh command is required to run this script. Please ensure that netsh is available in the root PATH." -foreground "red"
	exit
}


# Set up the form  ============================================================

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 

$objForm = New-Object System.Windows.Forms.Form 
$objForm.Text = "Skype4B / Lync DHCP Config Tool 1.01"
$objForm.Size = New-Object System.Drawing.Size(450,325) 
$objForm.MinimumSize = New-Object System.Drawing.Size(430,325)
$objForm.MaximumSize = New-Object System.Drawing.Size(5000,340)
$objForm.StartPosition = "CenterScreen"
[byte[]]$WindowIcon = @(66, 77, 56, 3, 0, 0, 0, 0, 0, 0, 54, 0, 0, 0, 40, 0, 0, 0, 16, 0, 0, 0, 16, 0, 0, 0, 1, 0, 24, 0, 0, 0, 0, 0, 2, 3, 0, 0, 18, 11, 0, 0, 18, 11, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114,0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0,198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 205, 132, 32, 234, 202, 160,255, 255, 255, 244, 229, 208, 205, 132, 32, 202, 123, 16, 248, 238, 224, 198, 114, 0, 205, 132, 32, 234, 202, 160, 255,255, 255, 255, 255, 255, 244, 229, 208, 219, 167, 96, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 248, 238, 224, 198, 114, 0, 198, 114, 0, 223, 176, 112, 255, 255, 255, 219, 167, 96, 198, 114, 0, 198, 114, 0, 255, 255, 255, 255, 255, 255, 227, 185, 128, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 241, 220, 192, 198, 114, 0, 198,114, 0, 248, 238, 224, 255, 255, 255, 244, 229, 208, 198, 114, 0, 198, 114, 0, 255, 255, 255, 255, 255, 255, 227, 185, 128, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 241, 220, 192, 198, 114, 0, 216, 158, 80, 255, 255, 255, 255, 255, 255, 252, 247, 240, 209, 141, 48, 198, 114, 0, 255, 255, 255, 255, 255, 255, 227, 185, 128, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 241, 220, 192, 198, 114, 0, 241, 220, 192, 255, 255, 255, 252, 247, 240, 212, 149, 64, 234, 202, 160, 198, 114, 0, 255, 255, 255, 255, 255, 255, 227, 185, 128, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 241, 220, 192, 205, 132, 32, 255, 255, 255, 255, 255, 255, 227, 185, 128, 198, 114, 0, 248, 238, 224, 202, 123, 16, 255, 255, 255, 255, 255, 255, 227, 185, 128, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 241, 220, 192, 234, 202, 160, 255, 255, 255, 255, 255, 255, 205, 132, 32, 198, 114, 0, 223, 176, 112, 223, 176, 112, 255, 255, 255, 255, 255, 255, 227, 185, 128, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 244, 229, 208, 252, 247, 240, 255, 255, 255, 237, 211, 176, 198, 114, 0, 198, 114, 0, 202, 123, 16, 248, 238, 224, 255, 255, 255, 255, 255, 255, 227, 185, 128, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 212, 149, 64, 255, 255, 255, 255, 255, 255, 255, 255, 255, 212, 149, 64, 198, 114, 0, 198, 114, 0, 198, 114, 0, 234, 202, 160, 255, 255,255, 255, 255, 255, 241, 220, 192, 205, 132, 32, 198, 114, 0, 198, 114, 0, 205, 132, 32, 227, 185, 128, 227, 185, 128, 227, 185, 128, 227, 185, 128, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 205, 132, 32, 227, 185, 128, 227, 185,128, 227, 185, 128, 219, 167, 96, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 0, 0)
$ico = New-Object IO.MemoryStream($WindowIcon, 0, $WindowIcon.Length)
$objForm.Icon = [System.Drawing.Icon]::FromHandle((new-object System.Drawing.Bitmap -argument $ico).GetHIcon())
$objForm.KeyPreview = $True
$objForm.TabStop = $false

$MyLinkLabel = New-Object System.Windows.Forms.LinkLabel
$MyLinkLabel.Location = New-Object System.Drawing.Size(312,2)
$MyLinkLabel.Size = New-Object System.Drawing.Size(120,15)
$MyLinkLabel.DisabledLinkColor = [System.Drawing.Color]::Red
$MyLinkLabel.VisitedLinkColor = [System.Drawing.Color]::Blue
$MyLinkLabel.LinkBehavior = [System.Windows.Forms.LinkBehavior]::HoverUnderline
$MyLinkLabel.LinkColor = [System.Drawing.Color]::Navy
$MyLinkLabel.TabStop = $False
$MyLinkLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$MyLinkLabel.Text = "www.myskypelab.com"
$MyLinkLabel.add_click(
{
	 [system.Diagnostics.Process]::start("http://www.myskypelab.com")
})
$objForm.Controls.Add($MyLinkLabel)


$ScopeLabel = New-Object System.Windows.Forms.Label
$ScopeLabel.Location = New-Object System.Drawing.Size(1,27) 
$ScopeLabel.Size = New-Object System.Drawing.Size(95,15) 
$ScopeLabel.Text = "Scope:"
$ScopeLabel.TextAlign = [System.Drawing.ContentAlignment]::TopRight
$ScopeLabel.TabStop = $False
$objForm.Controls.Add($ScopeLabel)

# Add Pool Dropdown box ============================================================
$ScopeDropDownBox = New-Object System.Windows.Forms.ComboBox 
$ScopeDropDownBox.Location = New-Object System.Drawing.Size(100,25) 
$ScopeDropDownBox.Size = New-Object System.Drawing.Size(300,20) 
$ScopeDropDownBox.DropDownHeight = 200 
$ScopeDropDownBox.tabIndex = 1
$ScopeDropDownBox.DropDownStyle = "DropDownList"
$ScopeDropDownBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Left
$objForm.Controls.Add($ScopeDropDownBox) 

[void] $ScopeDropDownBox.Items.Add("Server Scope")

[string]$FullOutput = "NOT SET"

Write-Host "---------------------------------------------------------------------------------------------"
Write-Host "COMMAND: netsh dhcp server show optionvalue" -foreground "green"
$FullOutput = netsh dhcp server show scope | Out-String

Write-Host "SCOPES:"
Write-Host "$FullOutput"

$ErrorResponse = "The following command was not found"
if($FullOutput -eq $null -or $FullOutput -eq "")
{
	Write-Host "ERROR: No response from netsh command. Check that netsh commands for DHCP server are available" -foreground "red"
}
elseif($FullOutput -imatch $ErrorResponse)
{
	Write-Host "ERROR: netsh command could not access DHCP commands. Please ensure that you are running the script on a DHCP server." -foreground "red"
}
else
{
	$ScopeSplit = $FullOutput.Split("`n")
	foreach ($line in $ScopeSplit) 
	{
		#Write-Host "LINE: $line"
		if ($line -match ".*(([01]?\d?\d|2[0-4]\d|25[0-5])\.){3}([01]?\d?\d|2[0-4]\d|25[0-5]).*")
		{
			$lineSplit = $line.Split(" ")
			
			$loopNo = 0
			foreach($aline in $lineSplit)
			{
				#Write-Host "$loopNo $aline"
				$loopNo++
			}

			$ScopeIP = $lineSplit[1].Trim()
			[void] $ScopeDropDownBox.Items.Add($ScopeIP)
			#Write-Host "FOUND: $ScopeIP"
		}
	}
}
Write-Host "---------------------------------------------------------------------------------------------"
	
	
$ScopeDropDownBox.add_SelectedValueChanged(
{
	$StatusLabel.Text = ""
	$PoolTextBox.text = ""
	$PoolTextBox.ForeColor = "Black"
	$WebServerTextBox.text = ""
	$WebServerTextBox.ForeColor = "Black"
	$VendorClassTextBox.text = "" 
	$VendorClassTextBox.ForeColor = "Black"
	$ProtocolTextBox.text = ""
	$ProtocolTextBox.ForeColor = "Black"
	$PortTextBox.text = ""
	$PortTextBox.ForeColor = "Black"
	$ServiceURLTextBox.text = "" 
	$ServiceURLTextBox.ForeColor = "Black"
	
	GetDHCPOptions
	
})


$PoolLabel = New-Object System.Windows.Forms.Label
$PoolLabel.Location = New-Object System.Drawing.Size(1,63) 
$PoolLabel.Size = New-Object System.Drawing.Size(95,15) 
$PoolLabel.Text = "SIP Pool FQDN:"
$PoolLabel.TextAlign = [System.Drawing.ContentAlignment]::TopRight
$PoolLabel.TabStop = $False
$objForm.Controls.Add($PoolLabel)

#Pool Text box ============================================================
$PoolTextBox = New-Object System.Windows.Forms.TextBox
$PoolTextBox.location = new-object system.drawing.size(100,60)
$PoolTextBox.size = new-object system.drawing.size(280,23)
#$URITextBox.Enabled = $False
$PoolTextBox.tabIndex = 2
$PoolTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Left
$PoolTextBox.text = ""   
$objform.controls.add($PoolTextBox)
$PoolTextBox.Add_TextChanged(
{
	$PoolTextBox.ForeColor = "Black"
	$StatusLabel.Text = ""
}
)

$PoolCheckBox = New-Object System.Windows.Forms.Checkbox 
$PoolCheckBox.Location = New-Object System.Drawing.Size(385,60) 
$PoolCheckBox.Size = New-Object System.Drawing.Size(20,20)
$PoolCheckBox.tabIndex = 12
$PoolCheckBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$PoolCheckBox.Add_Click(
{
	$StatusLabel.Text = ""
	if($PoolCheckBox.Checked)
	{
		$PoolTextBox.Enabled = $true
	}
	else
	{
		$PoolTextBox.Enabled = $false
	}
}
)
$objForm.Controls.Add($PoolCheckBox) 
$PoolCheckBox.Checked = $true

$WebServerLabel = New-Object System.Windows.Forms.Label
$WebServerLabel.Location = New-Object System.Drawing.Size(1,88) 
$WebServerLabel.Size = New-Object System.Drawing.Size(95,15) 
$WebServerLabel.Text = "Web FQDN:"
$WebServerLabel.TextAlign = [System.Drawing.ContentAlignment]::TopRight
$WebServerLabel.TabStop = $False
$objForm.Controls.Add($WebServerLabel)

#$WebServer Text box ============================================================
$WebServerTextBox = New-Object System.Windows.Forms.TextBox
$WebServerTextBox.location = new-object system.drawing.size(100,85)
$WebServerTextBox.size = new-object system.drawing.size(280,23)
#$URITextBox.Enabled = $False
$WebServerTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Left
$WebServerTextBox.tabIndex = 3
$WebServerTextBox.text = ""   
$objform.controls.add($WebServerTextBox)
$WebServerTextBox.Add_TextChanged(
{
	$WebServerTextBox.ForeColor = "Black"
	$StatusLabel.Text = ""
}
)

$WebServerCheckBox = New-Object System.Windows.Forms.Checkbox 
$WebServerCheckBox.Location = New-Object System.Drawing.Size(385,85) 
$WebServerCheckBox.Size = New-Object System.Drawing.Size(20,20)
$WebServerCheckBox.tabIndex = 13
$WebServerCheckBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$WebServerCheckBox.Add_Click(
{
	$StatusLabel.Text = ""
	if($WebServerCheckBox.Checked)
	{
		$WebServerTextBox.Enabled = $true
	}
	else
	{
		$WebServerTextBox.Enabled = $false
	}
}
)
$objForm.Controls.Add($WebServerCheckBox) 
$WebServerCheckBox.Checked = $true


$ProtocolLabel = New-Object System.Windows.Forms.Label
$ProtocolLabel.Location = New-Object System.Drawing.Size(1,113) 
$ProtocolLabel.Size = New-Object System.Drawing.Size(95,15) 
$ProtocolLabel.Text = "Protocol:"
$ProtocolLabel.TextAlign = [System.Drawing.ContentAlignment]::TopRight
$ProtocolLabel.TabStop = $False
$objForm.Controls.Add($ProtocolLabel)


#Identity Text box ============================================================
$ProtocolTextBox = New-Object System.Windows.Forms.TextBox
$ProtocolTextBox.location = new-object system.drawing.size(100,110)
$ProtocolTextBox.size = new-object system.drawing.size(280,23)
#$IdentityTextBox.Enabled = $False
$ProtocolTextBox.tabIndex = 4
$ProtocolTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Left
$ProtocolTextBox.text = "https"  
$objform.controls.add($ProtocolTextBox)
$ProtocolTextBox.Add_TextChanged(
{
	$ProtocolTextBox.ForeColor = "Black"
	$StatusLabel.Text = ""
}
)

$ProtocolCheckBox = New-Object System.Windows.Forms.Checkbox 
$ProtocolCheckBox.Location = New-Object System.Drawing.Size(385,110) 
$ProtocolCheckBox.Size = New-Object System.Drawing.Size(20,20)
$ProtocolCheckBox.tabIndex = 14
$ProtocolCheckBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$ProtocolCheckBox.Add_Click(
{
	$StatusLabel.Text = ""
	if($ProtocolCheckBox.Checked)
	{
		$ProtocolTextBox.Enabled = $true
	}
	else
	{
		$ProtocolTextBox.Enabled = $false
	}
}
)
$objForm.Controls.Add($ProtocolCheckBox) 
$ProtocolCheckBox.Checked = $true

$PortLabel = New-Object System.Windows.Forms.Label
$PortLabel.Location = New-Object System.Drawing.Size(1,138) 
$PortLabel.Size = New-Object System.Drawing.Size(95,15) 
$PortLabel.Text = "Web Port:"
$PortLabel.TextAlign = [System.Drawing.ContentAlignment]::TopRight
$PortLabel.TabStop = $False
$objForm.Controls.Add($PortLabel)


#$Port Text box ============================================================
$PortTextBox = New-Object System.Windows.Forms.TextBox
$PortTextBox.location = new-object system.drawing.size(100,135)
$PortTextBox.size = new-object system.drawing.size(280,23)
#$IdentityTextBox.Enabled = $False
$PortTextBox.tabIndex = 5
$PortTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Left
$PortTextBox.text = "443"   
$objform.controls.add($PortTextBox)
$PortTextBox.Add_TextChanged(
{
	$PortTextBox.ForeColor = "Black"
	$StatusLabel.Text = ""
}
)

$PortCheckBox = New-Object System.Windows.Forms.Checkbox 
$PortCheckBox.Location = New-Object System.Drawing.Size(385,135) 
$PortCheckBox.Size = New-Object System.Drawing.Size(20,20)
$PortCheckBox.tabIndex = 15
$PortCheckBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$PortCheckBox.Add_Click(
{
	$StatusLabel.Text = ""
	if($PortCheckBox.Checked)
	{
		$PortTextBox.Enabled = $true
	}
	else
	{
		$PortTextBox.Enabled = $false
	}
}
)
$objForm.Controls.Add($PortCheckBox) 
$PortCheckBox.Checked = $true

$ServiceURLLabel = New-Object System.Windows.Forms.Label
$ServiceURLLabel.Location = New-Object System.Drawing.Size(1,163) 
$ServiceURLLabel.Size = New-Object System.Drawing.Size(95,15) 
$ServiceURLLabel.Text = "Cert Prov URL:"
$ServiceURLLabel.TextAlign = [System.Drawing.ContentAlignment]::TopRight
$ServiceURLLabel.TabStop = $False
$objForm.Controls.Add($ServiceURLLabel)


#$ServiceURL Text box ============================================================
$ServiceURLTextBox = New-Object System.Windows.Forms.TextBox
$ServiceURLTextBox.location = new-object system.drawing.size(100,160)
$ServiceURLTextBox.size = new-object system.drawing.size(280,23)
#$IdentityTextBox.Enabled = $False
$ServiceURLTextBox.tabIndex = 6
$ServiceURLTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Left
$ServiceURLTextBox.text = "/CertProv/CertProvisioningService.svc"   
$objform.controls.add($ServiceURLTextBox)
$ServiceURLTextBox.Add_TextChanged(
{
	$ServiceURLTextBox.ForeColor = "Black"
	$StatusLabel.Text = ""
}
)

$ServiceURLCheckBox = New-Object System.Windows.Forms.Checkbox 
$ServiceURLCheckBox.Location = New-Object System.Drawing.Size(385,160) 
$ServiceURLCheckBox.Size = New-Object System.Drawing.Size(20,20)
$ServiceURLCheckBox.tabIndex = 16
$ServiceURLCheckBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$ServiceURLCheckBox.Add_Click(
{
	$StatusLabel.Text = ""
	if($ServiceURLCheckBox.Checked)
	{
		$ServiceURLTextBox.Enabled = $true
	}
	else
	{
		$ServiceURLTextBox.Enabled = $false
	}
}
)
$objForm.Controls.Add($ServiceURLCheckBox) 
$ServiceURLCheckBox.Checked = $true

$VendorClassLabel = New-Object System.Windows.Forms.Label
$VendorClassLabel.Location = New-Object System.Drawing.Size(1,188) 
$VendorClassLabel.Size = New-Object System.Drawing.Size(95,15) 
$VendorClassLabel.Text = "Vendor Class:"
$VendorClassLabel.TextAlign = [System.Drawing.ContentAlignment]::TopRight
$VendorClassLabel.TabStop = $False
$objForm.Controls.Add($VendorClassLabel)

#$VendorClass Text box ============================================================
$VendorClassTextBox = New-Object System.Windows.Forms.TextBox
$VendorClassTextBox.location = new-object system.drawing.size(100,185)
$VendorClassTextBox.size = new-object system.drawing.size(280,23)
#$URITextBox.Enabled = $False
$VendorClassTextBox.tabIndex = 7
$VendorClassTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Left
$VendorClassTextBox.text = "MS-UC-Client"   
$objform.controls.add($VendorClassTextBox)
$VendorClassTextBox.Add_TextChanged(
{
	$VendorClassTextBox.ForeColor = "Black"
	$StatusLabel.Text = ""
}
)

$VendorClassCheckBox = New-Object System.Windows.Forms.Checkbox 
$VendorClassCheckBox.Location = New-Object System.Drawing.Size(385,185) 
$VendorClassCheckBox.Size = New-Object System.Drawing.Size(20,20)
$VendorClassCheckBox.tabIndex = 17
$VendorClassCheckBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$VendorClassCheckBox.Add_Click(
{
	$StatusLabel.Text = ""
	if($VendorClassCheckBox.Checked)
	{
		$VendorClassTextBox.Enabled = $true
	}
	else
	{
		$VendorClassTextBox.Enabled = $false
	}
}
)
$objForm.Controls.Add($VendorClassCheckBox) 
$VendorClassCheckBox.Checked = $true


$CiscoButton = New-Object System.Windows.Forms.Button
$CiscoButton.Location = New-Object System.Drawing.Size(100,210)
$CiscoButton.Size = New-Object System.Drawing.Size(80,18)
$CiscoButton.Text = "Export Cisco"
$CiscoButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$CiscoButton.tabIndex = 8
$CiscoButton.Add_Click(
{
	$StatusLabel.Text = ""
	Export-Cisco
}
)
$objForm.Controls.Add($CiscoButton)




$DefaultsButton = New-Object System.Windows.Forms.Button
$DefaultsButton.Location = New-Object System.Drawing.Size(300,210)
$DefaultsButton.Size = New-Object System.Drawing.Size(80,18)
$DefaultsButton.Text = "Defaults"
$DefaultsButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$DefaultsButton.tabIndex = 9
$DefaultsButton.Add_Click(
{
	$StatusLabel.Text = ""
	#Set Defaults
	$PoolTextBox.text = "<Enter your Server Pool FQDN>"
	$PoolTextBox.ForeColor = "Black"
	$WebServerTextBox.text = "<Enter your Web Services FQDN>"
	$WebServerTextBox.ForeColor = "Black"
	$VendorClassTextBox.text = "MS-UC-Client" 
	$VendorClassTextBox.ForeColor = "Black"
	$ProtocolTextBox.text = "https"
	$ProtocolTextBox.ForeColor = "Black"
	$PortTextBox.text = "443"
	$PortTextBox.ForeColor = "Black"
	$ServiceURLTextBox.text = "/CertProv/CertProvisioningService.svc" 
	$ServiceURLTextBox.ForeColor = "Black"
}
)
$objForm.Controls.Add($DefaultsButton)


$SetDHCPButton = New-Object System.Windows.Forms.Button
$SetDHCPButton.Location = New-Object System.Drawing.Size(90,245)
$SetDHCPButton.Size = New-Object System.Drawing.Size(120,23)
$SetDHCPButton.Text = "Upload Settings"
#$SetDHCPButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$SetDHCPButton.tabIndex = 10
$SetDHCPButton.Add_Click(
{
	$StatusLabel.Text = ""
	SetDHCP
	
	$PoolTextBox.text = ""
	$PoolTextBox.ForeColor = "Black"
	$WebServerTextBox.text = ""
	$WebServerTextBox.ForeColor = "Black"
	$VendorClassTextBox.text = "" 
	$VendorClassTextBox.ForeColor = "Black"
	$ProtocolTextBox.text = ""
	$ProtocolTextBox.ForeColor = "Black"
	$PortTextBox.text = ""
	$PortTextBox.ForeColor = "Black"
	$ServiceURLTextBox.text = "" 
	$ServiceURLTextBox.ForeColor = "Black"
	
	GetDHCPOptions
}
)
$objForm.Controls.Add($SetDHCPButton)


$RemoveDHCPButton = New-Object System.Windows.Forms.Button
$RemoveDHCPButton.Location = New-Object System.Drawing.Size(230,245)
$RemoveDHCPButton.Size = New-Object System.Drawing.Size(120,23)
$RemoveDHCPButton.Text = "Remove Settings"
#$RemoveDHCPButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$RemoveDHCPButton.tabIndex = 11
$RemoveDHCPButton.Add_Click(
{
	$StatusLabel.Text = ""
	RemoveDHCP
	
	$PoolTextBox.text = ""
	$PoolTextBox.ForeColor = "Black"
	$WebServerTextBox.text = ""
	$WebServerTextBox.ForeColor = "Black"
	$VendorClassTextBox.text = "" 
	$VendorClassTextBox.ForeColor = "Black"
	$ProtocolTextBox.text = ""
	$ProtocolTextBox.ForeColor = "Black"
	$PortTextBox.text = ""
	$PortTextBox.ForeColor = "Black"
	$ServiceURLTextBox.text = "" 
	$ServiceURLTextBox.ForeColor = "Black"
	
	GetDHCPOptions
}
)
$objForm.Controls.Add($RemoveDHCPButton)


# Add the Status Label ============================================================
$StatusLabel = New-Object System.Windows.Forms.Label
$StatusLabel.Location = New-Object System.Drawing.Size(7,273) 
$StatusLabel.Size = New-Object System.Drawing.Size(300,15) 
$StatusLabel.Text = ""
$StatusLabel.forecolor = "red"
$StatusLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left
$StatusLabel.TabStop = $false
$objForm.Controls.Add($StatusLabel)



function GetDHCPOptions
{
	$Scope = $ScopeDropDownBox.Text
	Write-Host "---------------------------------------------------------------------------------------------"
	[string]$FullOutput = "NOT SET"
	$ServerScopeCheck = $false
	$WaitForVendor1 = $true
	$WaitForVendor2 = $true
	$WaitForVendor3 = $true
	$WaitForVendor4 = $true
	$WaitForVendor5 = $true
	$Usable120VendorClass = $true
	
	
	if($ScopeDropDownBox.Text -eq "Server Scope")
	{
		#[string]$cmd = "netsh dhcp server show optionvalue"
		Write-Host "COMMAND: netsh dhcp server show optionvalue" -foreground "green"
		$FullOutput = netsh dhcp server show optionvalue | Out-String
		$ServerScopeCheck = $true
	}
	else
	{
		Write-Host "COMMAND: netsh dhcp server scope $Scope show optionvalue" -foreground "green"
		#[string]$cmd = "netsh dhcp server scope $Scope show optionvalue"
		$FullOutput = netsh dhcp server scope $Scope show optionvalue | Out-String
		$ServerScopeCheck = $false
	}
	##
	
	$ErrorResponse = "The following command was not found"
	if($FullOutput -eq "" -or $FullOutput -eq $null)
	{
		Write-Host "ERROR: There was no response from the netsh command." -foreground "red"
	}
	elseif($FullOutput -imatch $ErrorResponse)
	{
		Write-Host "ERROR: netsh command could not access DHCP commands. Please ensure that you are running the script on a DHCP server." -foreground "red"
	}
	else
	{
	
		Write-Host "SCOPE OPTIONS:"
		Write-Host "$FullOutput"
		Write-Host ""
		
		#FIND THE OPTIONS
		$ScopeSplit = $FullOutput.Split("`n")
		$CurrentOptionValue = ""
		$CurrentElementType = ""
		$MSClientSection = $false
		foreach ($line in $ScopeSplit) 
		{
			#Write-Host "LINE: $line"
			if ($line -match "OptionId")
			{
				$lineSplit = $line.Split(":")
				
				$loopNo = 0
				foreach($aline in $lineSplit)
				{
					#Write-Host "$loopNo $aline"
					$loopNo++
				}

				$OptionID = $lineSplit[1].Trim()
				#Write-Host "FOUND OPTION: $OptionID" -foreground "green"
				$CurrentOptionValue = $OptionID
			}
			elseif($line -match "For vendor class \[MSUCClient\]")
			{
				Write-Host "Found MSUCClient Vendor Class" -foreground "yellow"
				$MSClientSection = $true
				$Usable120VendorClass = $true #NEED THIS HERE BECAUSE OF WEIRD SERVER SCOPE FORMATTING IN SOME CASES
			}
			elseif($line -match "DHCP Standard Options")
			{
				Write-Host "Options is DHCP Standard Options" -foreground "yellow"
				$MSClientSection = $false
				$Usable120VendorClass = $true #OPTION 120 SHOULD BE IN THE STANDARD VENDOR OPTIONS SECTION
			}
			elseif($line -match "For vendor class")
			{
				Write-Host "Options in non usable Vendor Class" -foreground "yellow"
				$MSClientSection = $false
				$Usable120VendorClass = $false
			}
			elseif($line -match "Option Element Type")
			{
				$lineSplit = $line.Split("=")
				
				$CurrentElementType = $lineSplit[1].Trim()
				#Write-Host "Current Element Type: $CurrentElementType"
			}
			elseif($line -match "Option Element Value")
			{
				$lineSplit = $line.Split("=")
				
				$loopNo = 0
				foreach($aline in $lineSplit)
				{
					#Write-Host "$loopNo $aline"
					$loopNo++
				}

				$OptionValue = $lineSplit[1].Trim()
				Write-Host "FOUND OPTION: $CurrentOptionValue $OptionValue" -foreground "green"
				
				
				if($CurrentOptionValue -eq "1" -and $MSClientSection -and $WaitForVendor1)
				{
					if($CurrentElementType -eq "BINARY")
					{
						$resultASCII = ConvertToASCII -InputString $OptionValue
						$VendorClassTextBox.Text = $resultASCII
						$VendorClassTextBox.ForeColor = "green"
						#TO FIXUP BUG IN SERVER SCOPE OUTPUT OF NETSH
						$WaitForVendor1 = $false
						Write-Host "OPTION 1 ASCII - $resultASCII" -foreground "green"
						Write-Host
					}
					else
					{
						Write-Host "INFO: Element Type $CurrentElementType is not supported. Element type must be binary. If you upload a new value for this option using this tool it will replace the existing unsupported setting." -foreground "red"
						$VendorClassTextBox.Text = "<UNSUPPORTED TYPE>"
						$VendorClassTextBox.ForeColor = "red"
					}
				}
				elseif($CurrentOptionValue -eq "2" -and $MSClientSection -and $WaitForVendor2)
				{
					if($CurrentElementType -eq "BINARY")
					{
						$resultASCII = ConvertToASCII -InputString $OptionValue
						$ProtocolTextBox.Text = $resultASCII
						$ProtocolTextBox.ForeColor = "green"
						#TO FIXUP BUG IN SERVER SCOPE OUTPUT OF NETSH
						$WaitForVendor2 = $false
						Write-Host "OPTION 2 ASCII - $resultASCII" -foreground "green"
						Write-Host
					}
					else
					{
						Write-Host "INFO: Element Type $CurrentElementType is not supported. Element type must be binary. If you upload a new value for this option using this tool it will replace the existing unsupported setting." -foreground "red"
						$ProtocolTextBox.Text = "<UNSUPPORTED TYPE>"
						$ProtocolTextBox.ForeColor = "red"
					}
				}
				elseif($CurrentOptionValue -eq "3" -and $MSClientSection -and $WaitForVendor3)
				{
					if($CurrentElementType -eq "BINARY")
					{
						$resultASCII = ConvertToASCII -InputString $OptionValue
						#Write-Host "FOUND 3" -foreground "green"
						$WebServerTextBox.text = $resultASCII
						$WebServerTextBox.ForeColor = "green"
						#TO FIXUP BUG IN SERVER SCOPE OUTPUT OF NETSH
						$WaitForVendor3 = $false
						Write-Host "OPTION 3 ASCII - $resultASCII" -foreground "green"
						Write-Host
					}
					else
					{
						Write-Host "INFO: Element Type $CurrentElementType is not supported. Element type must be binary. If you upload a new value for this option using this tool it will replace the existing unsupported setting." -foreground "red"
						$WebServerTextBox.Text = "<UNSUPPORTED TYPE>"
						$WebServerTextBox.ForeColor = "red"
					}
				}
				elseif($CurrentOptionValue -eq "4" -and $MSClientSection -and $WaitForVendor4)
				{
					if($CurrentElementType -eq "BINARY")
					{
						$resultASCII = ConvertToASCII -InputString $OptionValue
						#Write-Host "FOUND 4" -foreground "green"
						$PortTextBox.Text = $resultASCII
						$PortTextBox.ForeColor = "green"
						#TO FIXUP BUG IN SERVER SCOPE OUTPUT OF NETSH
						$WaitForVendor4 = $false
						Write-Host "OPTION 4 ASCII - $resultASCII" -foreground "green"
						Write-Host
					}
					else
					{
						Write-Host "INFO: Element Type $CurrentElementType is not supported. Element type must be binary. If you upload a new value for this option using this tool it will replace the existing unsupported setting." -foreground "red"
						$PortTextBox.Text = "<UNSUPPORTED TYPE>"
						$PortTextBox.ForeColor = "red"
					}
				}
				elseif($CurrentOptionValue -eq "5" -and $MSClientSection -and $WaitForVendor5)
				{
					if($CurrentElementType -eq "BINARY")
					{
						$resultASCII = ConvertToASCII -InputString $OptionValue
						#Write-Host "FOUND 5" -foreground "green"
						$ServiceURLTextBox.Text = $resultASCII
						$ServiceURLTextBox.ForeColor = "green"
						#TO FIXUP BUG IN SERVER SCOPE OUTPUT OF NETSH
						$WaitForVendor5 = $false
						Write-Host "OPTION 5 ASCII - $resultASCII" -foreground "green"
						Write-Host
					}
					else
					{
						Write-Host "INFO: Element Type $CurrentElementType is not supported. Element type must be binary. If you upload a new value for this option using this tool it will replace the existing unsupported setting." -foreground "red"
						$ServiceURLTextBox.Text = "<UNSUPPORTED TYPE>"
						$ServiceURLTextBox.ForeColor = "red"
					}
				}
				elseif($CurrentOptionValue -eq "120" -and $Usable120VendorClass)  #NEED TO HANDLE SPECIAL ENCODING FOR OPTION 120
				{
					if($CurrentElementType -eq "BINARY")
					{
						$CroppedOptionsString = $OptionValue.Substring(2,($OptionValue.length - 2))
						$EncodingValue = $OptionValue.Substring(0,2)
						#Write-Host "Encoding Value: $EncodingValue"
						
						if($EncodingValue -eq "00")
						{	
							Write-Host "INFO: Option 120 - FQDN Encoding is used." -foreground "yellow"
							
							#COUNT THE NUMBER OF FQDNs
							$NoOfFQDNs = 0
							for($i=0; $i -le ($CroppedOptionsString.length - 1); $i=$i+2)
							{
								$ByteToCheck = $CroppedOptionsString.Substring($i,2)
								if($ByteToCheck -eq "00")
								{
									#Write-Host "FOUND A 00!"
									$NoOfFQDNs++
								}
							}
							
							if($NoOfFQDNs -gt 1)
							{
								Write-Host "INFO: Found more than 1 FQDN." -foreground "yellow"
								Write-Host "ERROR: More than 1 FQDN found. Having more than 1 FQDN is not supported by this tool. If you upload a new value for Option 120 using this tool it will replace the existing unsupported setting." -foreground "red"
								$PoolTextBox.Text = "<UNSUPPORTED ENCODING>"
								$PoolTextBox.ForeColor = "red"
							}
							else
							{
								$LoopCount = 0
								$FinalString = ""
								$SizeOfOption120 = $CroppedOptionsString.length
								
								#LENGTH CHECKS OPTION 120
								$LoopNo = 0
								Write-Host "INFO: Option 120 - Checking length values in Option 120" -foreground "yellow"
								$CurrentByte = 0
								$ByteSizeOfOption120 = $SizeOfOption120 * 2
								$lengthCheck = $false
								for($i=0; $i -le ($ByteSizeOfOption120 - 1); $i++)
								{
									$LoopNo++
									$LengthOfSection = $CroppedOptionsString.Substring($CurrentByte,2)
									if($LengthOfSection -eq "00")
									{
										Write-Host "INFO: Option 120 - Length check passed!" -foreground "green"
										$lengthCheck = $true
										break
									}
									[int]$intLength = [convert]::ToInt32($LengthOfSection, 16)
									$intLength = $intLength * 2
									if($intLength -gt ($ByteSizeOfOption120 - $CurrentByte))
									{
										Write-Host "ERROR: Option 120 - Impossible length value found." -foreground "red"
										$FinalString = "<UNSUPPORTED ENCODING>"
										$lengthCheck = $false
										break
									}
									else
									{
										#MOVE TO THE NEXT LENGTH BYTE
										$CurrentByte = $CurrentByte + $intLength + 2
									}
								}
								
								#DECODE OPTION 120
								if($lengthCheck)
								{
									Write-Host "INFO: Decoding Option 120" -foreground "yellow"
									while($CroppedOptionsString.length -gt 1) 
									{
										$LengthOfSection = $CroppedOptionsString.Substring(0,2)
										[int]$intLength = [convert]::ToInt32($LengthOfSection, 16)
										if($intLength -gt $SizeOfOption120)
										{
											Write-Host "ERROR: Option 120 - Impossible length value found." -foreground "red"
											$FinalString = "<UNSUPPORTED ENCODING>"
											break
										}
										
										$intLength = $intLength * 2   # NEED TO MULTIPLY BY 2 TO CONVERT FROM BYTES TO CHARS
										
										$theSection = $CroppedOptionsString.Substring(2,$intLength)
										$CroppedOptionsString = $CroppedOptionsString.Substring((2 + $intLength),($CroppedOptionsString.length - (2 + $intLength)))
										
										if($CroppedOptionsString -match "^00")
										{
											$FinalString += "${theSection}"
											break
										}
										else
										{
											$FinalString += "${theSection}2E"
										}
										$LoopCount++
										if($LoopCount -gt 100){Write-Host "ERROR: Option 120 - Too many sections. Break Loop." -foreground "red";break}
									}
									
									if($FinalString -eq "<UNSUPPORTED ENCODING>")
									{
										$PoolTextBox.Text = ConvertToASCII -InputString $FinalString
										$PoolTextBox.ForeColor = "red"
									}
									else
									{
										$resultASCII = ConvertToASCII -InputString $FinalString
										$PoolTextBox.Text = $resultASCII
										$PoolTextBox.ForeColor = "green"
										Write-Host "OPTION 120 ASCII - $resultASCII" -foreground "green"
									}
								}
								else
								{
									$PoolTextBox.Text = "<UNSUPPORTED ENCODING>"
									$PoolTextBox.ForeColor = "red"
									Write-Host "ERROR: Option 120 - Encoding length check failed. If you upload a new value for Option 120 using this tool it will replace the existing unsupported setting." -foreground "red"
								}
								
								#The Internet standards (Requests for Comments) for protocols mandate that component hostname labels may contain 
								#only the ASCII letters 'a' through 'z' (in a case-insensitive manner), the digits '0' through '9', and the hyphen ('-'). 
								#The original specification of hostnames in RFC 952, mandated that labels could not start with a digit or with a hyphen, 
								#and must not end with a hyphen. However, a subsequent specification (RFC 1123) permitted hostname labels to start with 
								#digits. No other symbols, punctuation characters, or white space are permitted.
								
								#CHECK CHARACTERS IN FQDN
								if($FinalString -ne "<UNSUPPORTED ENCODING>")
								{
									$FQDNtoTest = ConvertToASCII -InputString $FinalString
									if($FQDNtoTest -imatch "(?=^.{4,253}$)(^((?!-)[a-zA-Z0-9-]{0,62}[a-zA-Z0-9]\.)+[a-zA-Z]{2,63}$)")
									{
										Write-Host "INFO: Option 120 - FQDN format check passed!" -foreground "green"
										Write-Host
									}
									else
									{
										Write-Host "INFO: FQDN format check failed. If you upload a new value for Option 120 using this tool it will replace the existing unsupported setting." -foreground "red"
										$PoolTextBox.Text = "<UNSUPPORTED ENCODING>"
										$PoolTextBox.ForeColor = "red"
										Write-Host
									}
								}
								else
								{
									Write-Host "INFO: Ignore FQDN Check." -foreground "yellow"
									Write-Host
								}
							}
						}
						elseif($EncodingValue -eq "01")
						{
							Write-Host "INFO: INFO: Option 120 - IP Address Encoding is used." -foreground "yellow"
							Write-Host "ERROR: IP Address Encoding. Unsupported Option 120 encoding. If you upload a new value for Option 120 using this tool it will replace the existing unsupported setting." -foreground "red"
							$PoolTextBox.Text = "<UNSUPPORTED ENCODING>"
							$PoolTextBox.ForeColor = "red"
						}
						else
						{
							Write-Host "ERROR: Unsupported Option 120 encoding. If you upload a new value for Option 120 using this tool it will replace the existing unsupported setting." -foreground "red"
							$PoolTextBox.Text = "<UNSUPPORTED ENCODING>"
							$PoolTextBox.ForeColor = "red"
						}
					}
					else
					{
						Write-Host "INFO: Element Type $CurrentElementType is not supported. Element type must be binary. If you upload a new value for this option using this tool it will replace the existing unsupported setting." -foreground "red"
						$ServiceURLTextBox.Text = "<UNSUPPORTED TYPE>"
						$ServiceURLTextBox.ForeColor = "red"
					}
				}
				$CurrentElementType = ""
			}
		}
	}
	Write-Host "---------------------------------------------------------------------------------------------"
	Write-Host 
}

<#
Options for Scope 192.168.0.0:

        DHCP Standard Options :
        General Option Values:
        OptionId : 160
        Option Value:
                Number of Option Elements = 1
                Option Element Type = STRING
                Option Element Value = ftp://ConfigServer.mylynclab.com/
        OptionId : 2
        Option Value:
                Number of Option Elements = 1
                Option Element Type = DWORD
                Option Element Value = 34200

#>


function SetDHCP
{
	$Scope = $ScopeDropDownBox.Text
	$VendorClassNameString = "MSUCClient"
	
	#CHECK EXISTING OPTION CONFIGURATION
	$OutputVariable = netsh dhcp server show optiondef 2>&1 | Out-String
	Write-Host "COMMAND: netsh dhcp server show optiondef" -foreground "green"
	Write-Host "RESULT: $OutputVariable"
	
	$OutputVariableSplit = $OutputVariable.Split("`n")
	$MSClientSection = "FALSE"
	foreach ($line in $OutputVariableSplit) 
	{
		if($line -match "Options\[vendor= MSUCClient\]")
		{
			Write-Host "INFO: Found MSUCClient Vendor Class" -foreground "yellow"
			$MSClientSection = "TRUE"
		}
		elseif($line -match "Options \[Non Vendor specific\]")
		{
			Write-Host "INFO: Found standard options section" -foreground "yellow"
			$MSClientSection = "FALSE"
		}
		elseif($line -match "Options\[vendor=")
		{
			Write-Host "INFO: In other Vendor Class" -foreground "yellow"
			$MSClientSection = "OTHER"
		}
		elseif ($line -match " 120 " -and $MSClientSection -ne "OTHER")
		{
			if($line -match "BINARY")
			{
				Write-Host "INFO: Option 120 is BINARY. OK" -foreground "Yellow"
			}
			else
			{
				Write-Host "INFO: Option 120 is NOT BINARY. FAIL" -foreground "red"
				
				if($PoolCheckBox.Checked)
				{
					$a = new-object -comobject wscript.shell 
					$intAnswer = $a.popup("Existing Option 120 (SIP Server) definition is not of binary type. Do you want to delete the existing definition and replace it?`r`n`r`nNote: This will affect all scopes that use this option.",0,"Change Definition Type",4) 
					if ($intAnswer -eq 6) 
					{
						$OutputVariable = netsh dhcp server delete optiondef 120 2>&1 | Out-String
						Write-Host "COMMAND: netsh dhcp server delete optiondef 120" -foreground "green"
						Write-Host "RESULT: $OutputVariable"
					}
					else
					{
						Write-Host "INFO: Cancel upload. Definition retained." -foreground "Yellow"
					}
				}
			}
		}
		elseif($line -match " 1 " -and $MSClientSection -eq "TRUE")
		{
			if($line -match "BINARY")
			{
				Write-Host "INFO: Option 1 (UCIdentifier) is BINARY. OK" -foreground "Yellow"
			}
			else
			{
				Write-Host "INFO: Option 1 (UCIdentifier) is NOT BINARY. FAIL" -foreground "red"
				if($VendorClassCheckBox.Checked)
				{
					$a = new-object -comobject wscript.shell 
					$intAnswer = $a.popup("Existing Option 1 (UCIdentifier) definition is not of binary type. Do you want to delete the existing definition and replace it?`r`n`r`nNote: This will affect all scopes that use this option.",0,"Change Definition Type",4) 
					if ($intAnswer -eq 6) 
					{
						$OutputVariable = netsh dhcp server delete optiondef 1 Vendor=MSUCClient 2>&1 | Out-String
						Write-Host "COMMAND: netsh dhcp server delete optiondef 1 Vendor=MSUCClient" -foreground "green"
						Write-Host "RESULT: $OutputVariable"
					}
					else
					{
						Write-Host "INFO: Cancel upload. Definition retained." -foreground "Yellow"
					}
				}
			}
		}
		elseif ($line -match " 2 " -and $MSClientSection -eq "TRUE")
		{
			if($line -match "BINARY")
			{
				Write-Host "INFO: Option 2 (URLScheme) is BINARY. OK" -foreground "Yellow"
			}
			else
			{
				Write-Host "INFO: Option 2 (URLScheme) is NOT BINARY. FAIL" -foreground "red"
				if($ProtocolCheckBox.Checked)
				{
					$a = new-object -comobject wscript.shell 
					$intAnswer = $a.popup("Existing Option 2 (URLScheme) definition is not of binary type. Do you want to delete the existing definition and replace it?`r`n`r`nNote: This will affect all scopes that use this option.",0,"Change Definition Type",4) 
					if ($intAnswer -eq 6) 
					{
						$OutputVariable = netsh dhcp server delete optiondef 2 Vendor=MSUCClient 2>&1 | Out-String
						Write-Host "COMMAND: netsh dhcp server delete optiondef 2 Vendor=MSUCClient" -foreground "green"
						Write-Host "RESULT: $OutputVariable"
					}
					else
					{
						Write-Host "INFO: Cancel upload. Definition retained." -foreground "Yellow"
					}
				}
			}
		}
		elseif ($line -match " 3 " -and $MSClientSection -eq "TRUE")
		{
			if($line -match "BINARY")
			{
				Write-Host "WARNING: Option 3 (WebServerFQDN) is BINARY. OK" -foreground "Yellow"
			}
			else
			{
				Write-Host "INFO: Option 3 (WebServerFQDN) is NOT BINARY. FAIL" -foreground "red"
				if($WebServerCheckBox.Checked)
				{
					$a = new-object -comobject wscript.shell 
					$intAnswer = $a.popup("Existing Option 3 (WebServerFQDN) definition is not of binary type. Do you want to delete the existing definition and replace it?`r`n`r`nNote: This will affect all scopes that use this option.",0,"Change Definition Type",4) 
					if ($intAnswer -eq 6) 
					{
						$OutputVariable = netsh dhcp server delete optiondef 3 Vendor=MSUCClient 2>&1 | Out-String
						Write-Host "COMMAND: netsh dhcp server delete optiondef 3 Vendor=MSUCClient" -foreground "green"
						Write-Host "RESULT: $OutputVariable"
					}
					else
					{
						Write-Host "INFO: Cancel upload. Definition retained." -foreground "Yellow"
					}
				}
			}
		}
		elseif ($line -match " 4 " -and $MSClientSection -eq "TRUE")
		{
			if($line -match "BINARY")
			{
				Write-Host "INFO: Option 4 (WebServerPort) is BINARY. OK" -foreground "Yellow"
			}
			else
			{
				Write-Host "INFO: Option 4 (WebServerPort) is NOT BINARY. FAIL" -foreground "red"
				if($PortCheckBox.Checked)
				{
					$a = new-object -comobject wscript.shell 
					$intAnswer = $a.popup("Existing Option 4 (WebServerPort) definition is not of binary type. Do you want to delete the existing definition and replace it?`r`n`r`nNote: This will affect all scopes that use this option.",0,"Change Definition Type",4) 
					if ($intAnswer -eq 6) 
					{
						$OutputVariable = netsh dhcp server delete optiondef 4 Vendor=MSUCClient 2>&1 | Out-String
						Write-Host "COMMAND: netsh dhcp server delete optiondef 4 Vendor=MSUCClient" -foreground "green"
						Write-Host "RESULT: $OutputVariable"
					}
					else
					{
						Write-Host "INFO: Cancel upload. Definition retained." -foreground "Yellow"
					}
				}
			}
		}
		elseif ($line -match " 5 " -and $MSClientSection -eq "TRUE")
		{
			if($line -match "BINARY")
			{
				Write-Host "INFO: Option 5 (CertProvRelPath) is BINARY. OK" -foreground "Yellow"
			}
			else
			{
				Write-Host "INFO: Option 5 (CertProvRelPath) is NOT BINARY. FAIL" -foreground "red"
				if($ServiceURLCheckBox.Checked)
				{
					$a = new-object -comobject wscript.shell 
					$intAnswer = $a.popup("Existing Option 5 (CertProvRelPath) definition is not of binary type. Do you want to delete the existing definition and replace it?`r`n`r`nNote: This will affect all scopes that use this option.",0,"Change Definition Type",4) 
					if ($intAnswer -eq 6) 
					{
						$OutputVariable = netsh dhcp server delete optiondef 5 Vendor=MSUCClient 2>&1 | Out-String
						Write-Host "COMMAND: netsh dhcp server delete optiondef 5 Vendor=MSUCClient" -foreground "green"
						Write-Host "RESULT: $OutputVariable"
					}
					else
					{
						Write-Host "INFO: Cancel upload. Definition retained." -foreground "Yellow"
					}
				}
			}
		}
	}
	
	
	if($PoolCheckBox.Checked)
	{
		if($PoolTextBox.Text -ne "") #OPTION 120
		{
			if($PoolTextBox.Text -imatch "(?=^.{4,253}$)(^((?!-)[a-zA-Z0-9-]{0,62}[a-zA-Z0-9]\.)+[a-zA-Z]{2,63}$)")
			{
				$PoolString = $PoolTextBox.Text
				
				#CREATE OPTION 120 FQDN ENCODING
				$PoolStringSplit = $PoolString.Split(".")
				$HexPool = "00"  #ENCODING IS FQDN
				
				foreach($Section in $PoolStringSplit)
				{
					$NoOfChars = $Section.length
					$SectionHex = ConvertToHex -InputString $Section
					$NumberOfCharsHex = "{0:X2}" -f $NoOfChars
					$HexPool += "${NumberOfCharsHex}$SectionHex".ToUpper()
				}
				
				Write-Host "---------------------------------------------------------------------------------------------"
				$HexPool = "${HexPool}00"
				Write-Host "SIP POOL FQDN: $HexPool"
				
				if($ScopeDropDownBox.Text -eq "Server Scope")
				{
					$OutputVariable = netsh dhcp server add optiondef 120 UCSipServer Binary 0 comment="Sip Server Fqdn" 2>&1 | Out-String
					Write-Host "COMMAND: netsh dhcp server add optiondef 120 UCSipServer Binary 0 comment=`"Sip Server Fqdn`"" -foreground "green"
					Write-Host "RESULT: $OutputVariable"
					
					$OutputVariable = netsh dhcp server set optionvalue 120 Binary $HexPool 2>&1 | Out-String
					Write-Host "COMMAND: netsh dhcp server set optionvalue 120 Binary $HexPool" -foreground "green"
					Write-Host "RESULT: $OutputVariable"
				}
				else
				{
					$OutputVariable = netsh dhcp server add optiondef 120 UCSipServer Binary 0 comment="Sip Server Fqdn" 2>&1 | Out-String
					Write-Host "COMMAND: netsh dhcp server add optiondef 120 UCSipServer Binary 0 comment=`"Sip Server Fqdn`"" -foreground "green"
					Write-Host "RESULT: $OutputVariable"
					
					$OutputVariable = netsh dhcp server scope $Scope set optionvalue 120 Binary $HexPool 2>&1 | Out-String
					Write-Host "COMMAND: netsh dhcp server scope $Scope set optionvalue 120 Binary $HexPool" -foreground "green"
					Write-Host "RESULT: $OutputVariable"
				}
				Write-Host "---------------------------------------------------------------------------------------------"
				Write-Host
			}
			else
			{
				Write-Host "ERROR: FQDN format check failed on the SIP Pool FQDN. Supported characters include ASCII letters a-z (case insensitive), the digits 0-9, and the hyphen (`"-`")." -foreground "red"
				$StatusLabel.Text = "ERROR: FQDN format check failed on the SIP Pool FQDN."
			}
		}
		else
		{
			Write-Host "ERROR: No Pool FQDN has been supplied." -foreground "red"
			$StatusLabel.Text = "ERROR: No Pool FQDN has been supplied."
		}
	}
	else
	{
		Write-Host "INFO: Pool FQDN not selected." -foreground "yellow"
	}
	
	
	
	if($VendorClassCheckBox.Checked)
	{
		if($VendorClassTextBox.Text -ne "") #VENDOR CLASS NAME
		{
			Write-Host "---------------------------------------------------------------------------------------------"
			if($VendorClassTextBox.Text -ne "MS-UC-Client")
			{
				Write-Host "WARNING: You are using a non-standard setting for the DHCP Vendor Class name. The standard setting for Skype4B / Lync is `"MS-UC-Client`"" -foreground "red"
			}

			#Write-Host
			$VendorClassString = $VendorClassTextBox.Text
			$HexVendorClass = ConvertToHex -InputString $VendorClassString
			$HexVendorClass = $HexVendorClass.ToUpper()
			Write-Host "VENDORCLASS: $VendorClassString"
			
			if($ScopeDropDownBox.Text -eq "Server Scope")
			{
				#CREATE VENDOR CLASS
				$OutputVariable = netsh dhcp server add class $VendorClassNameString "UC Vendor Class Id" "$VendorClassString" 1 2>&1 | Out-String
				Write-Host "COMMAND: netsh dhcp server add class $VendorClassNameString `"UC Vendor Class Id`" `"$VendorClassString`"" -foreground "green"
				Write-Host "RESULT: $OutputVariable"
				
				#SET THE VENDOR CLASS OPTION
				$OutputVariable = netsh dhcp server add optiondef 1 UCIdentifier Binary 0 Vendor=$VendorClassNameString comment="UC Identifier" 2>&1 | Out-String
				Write-Host "COMMAND: netsh dhcp server add optiondef 1 UCIdentifier Binary 0 Vendor=$VendorClassNameString comment=`"UC Identifier`"" -foreground "green"
				Write-Host "RESULT: $OutputVariable"
				$OutputVariable = netsh dhcp server set optionvalue 1 Binary vendor=$VendorClassNameString $HexVendorClass 2>&1 | Out-String
				Write-Host "COMMAND: netsh dhcp server set optionvalue 1 Binary vendor=$VendorClassNameString $HexVendorClass" -foreground "green"
				Write-Host "RESULT: $OutputVariable"
			}
			else
			{
				#CREATE VENDOR CLASS
				$OutputVariable = netsh dhcp server add class $VendorClassNameString "UC Vendor Class Id" "$VendorClassString" 1 2>&1 | Out-String
				Write-Host "COMMAND: netsh dhcp server add class $VendorClassNameString `"UC Vendor Class Id`" `"$VendorClassString`"" -foreground "green"
				Write-Host "RESULT: $OutputVariable"
				
				#SET THE VENDOR CLASS OPTION
				$OutputVariable = netsh dhcp server add optiondef 1 UCIdentifier Binary 0 Vendor=$VendorClassNameString comment="UC Identifier" 2>&1 | Out-String
				Write-Host "COMMAND: netsh dhcp server add optiondef 1 UCIdentifier Binary 0 Vendor=$VendorClassNameString comment=`"UC Identifier`"" -foreground "green"
				Write-Host "RESULT: $OutputVariable"
				$OutputVariable = netsh dhcp server scope $Scope set optionvalue 1 Binary vendor=$VendorClassNameString $HexVendorClass 2>&1 | Out-String
				Write-Host "COMMAND: netsh dhcp server scope $Scope set optionvalue 1 Binary vendor=$VendorClassNameString $HexVendorClass" -foreground "green"
				Write-Host "RESULT: $OutputVariable"
			}
			Write-Host "---------------------------------------------------------------------------------------------"
			Write-Host
		
		}
		else
		{
			Write-Host "ERROR: No vendor class has been supplied. Cannot add vendor specific options." -foreground "red"
		}
	}
	else
	{
		Write-Host "INFO: Vendor Class not seleted." -foreground "yellow"
	}
	
	if($WebServerCheckBox.Checked)
	{	
		if($WebServerTextBox.text -ne "")
		{
			if($WebServerTextBox.text -imatch "(?=^.{4,253}$)(^((?!-)[a-zA-Z0-9-]{0,62}[a-zA-Z0-9]\.)+[a-zA-Z]{2,63}$)")
			{
				Write-Host "---------------------------------------------------------------------------------------------"
				$WebServerString = $WebServerTextBox.text 
				$HexWebServer = ConvertToHex -InputString $WebServerString
				$HexWebServer = $HexWebServer.ToUpper()
				Write-Host "WEB SERVER: $HexWebServer"
					
				if($ScopeDropDownBox.Text -eq "Server Scope")
				{
					$OutputVariable = netsh dhcp server add optiondef 3 WebServerFqdn Binary 0 Vendor=$VendorClassNameString comment="Web Server Fqdn" 2>&1 | Out-String
					Write-Host "COMMAND: netsh dhcp server add optiondef 3 WebServerFqdn Binary 0 Vendor=$VendorClassNameString comment=`"Web Server Fqdn`"" -foreground "green"
					Write-Host "RESULT: $OutputVariable"
					$OutputVariable = netsh dhcp server set optionvalue 3 Binary vendor=$VendorClassNameString $HexWebServer 2>&1 | Out-String
					Write-Host "COMMAND: netsh dhcp server set optionvalue 3 Binary vendor=$VendorClassNameString $HexWebServer" -foreground "green"
					Write-Host "RESULT: $OutputVariable"
				}
				else
				{
					$OutputVariable = netsh dhcp server add optiondef 3 WebServerFqdn Binary 0 Vendor=$VendorClassNameString comment="Web Server Fqdn" 2>&1 | Out-String
					Write-Host "COMMAND: netsh dhcp server add optiondef 3 WebServerFqdn Binary 0 Vendor=$VendorClassNameString comment=`"Web Server Fqdn`"" -foreground "green"
					Write-Host "RESULT: $OutputVariable"
					$OutputVariable = netsh dhcp server scope $Scope set optionvalue 3 Binary vendor=$VendorClassNameString $HexWebServer 2>&1 | Out-String
					Write-Host "COMMAND: netsh dhcp server scope $Scope set optionvalue 3 Binary vendor=$VendorClassNameString $HexWebServer" -foreground "green"
					Write-Host "RESULT: $OutputVariable"
				}
				Write-Host "---------------------------------------------------------------------------------------------"
				Write-Host
			}
			else
			{
				Write-Host "ERROR: FQDN format check failed on the Web FQDN field. Supported characters include ASCII letters a-z (case insensitive), the digits 0-9, and the hyphen (`"-`")." -foreground "red"
				$StatusLabel.Text = "ERROR: FQDN format check failed on the Web FQDN field."
			}
		}
		else
		{
			Write-Host "ERROR: No Web Server FQDN has been supplied." -foreground "red"
		}
	}
	else
	{
		Write-Host "INFO: Web Server FQDN not seleted." -foreground "yellow"
	}
	
	if($ProtocolCheckBox.Checked)
	{
		if($ProtocolTextBox.Text -ne "")
		{
			Write-Host "---------------------------------------------------------------------------------------------"
			if($ProtocolTextBox.Text -ne "https")
			{
				Write-Host "WARNING: You are using a non-standard setting for the protocol type. The standard setting for Skype4B / Lync is `"https`"" -foreground "red"
			}

			$ProtocolString = $ProtocolTextBox.Text
			$HexProtocol = ConvertToHex -InputString $ProtocolString
			$HexProtocol = $HexProtocol.ToUpper()
			Write-Host "PROTOCOL: $HexProtocol"
			
			if($ScopeDropDownBox.Text -eq "Server Scope")
			{
				#SET THE PROTOCOL OPTION
				$OutputVariable = netsh dhcp server add optiondef 2 URLScheme Binary 0 Vendor=$VendorClassNameString comment="URL Scheme" 2>&1 | Out-String
				Write-Host "COMMAND: netsh dhcp server add optiondef 2 URLScheme Binary 0 Vendor=$VendorClassNameString comment=`"URL Scheme`"" -foreground "green"
				Write-Host "RESULT: $OutputVariable"
				$OutputVariable = netsh dhcp server set optionvalue 2 Binary vendor=$VendorClassNameString $HexProtocol 2>&1 | Out-String
				Write-Host "COMMAND: netsh dhcp server set optionvalue 2 Binary vendor=$VendorClassNameString $HexProtocol" -foreground "green"
				Write-Host "RESULT: $OutputVariable"
			}
			else
			{
				#SET THE PROTOCOL OPTION
				$OutputVariable = netsh dhcp server add optiondef 2 URLScheme Binary 0 Vendor=$VendorClassNameString comment="URL Scheme" 2>&1 | Out-String
				Write-Host "COMMAND: netsh dhcp server add optiondef 2 URLScheme Binary 0 Vendor=$VendorClassNameString comment=`"URL Scheme`"" -foreground "green"
				Write-Host "RESULT: $OutputVariable"
				$OutputVariable = netsh dhcp server scope $Scope set optionvalue 2 Binary vendor=$VendorClassNameString $HexProtocol 2>&1 | Out-String
				Write-Host "COMMAND: netsh dhcp server scope $Scope set optionvalue 2 Binary vendor=$VendorClassNameString $HexProtocol" -foreground "green"
				Write-Host "RESULT: $OutputVariable"
			}
			Write-Host "---------------------------------------------------------------------------------------------"
			Write-Host
		}
		else
		{
			Write-Host "ERROR: No protocol has been supplied." -foreground "red"
			#$StatusLabel.Text = "ERROR: No Web service URL has been supplied."
		}
	}
	else
	{
		Write-Host "INFO: Web Server FQDN not seleted." -foreground "yellow"
	}
	
	if($PortCheckBox.Checked)
	{
		if($PortTextBox.Text -ne "")
		{
			Write-Host "---------------------------------------------------------------------------------------------"
			if($PortTextBox.Text -ne "443")
			{
				Write-Host "WARNING: You are using a non-standard setting for the port number. The standard setting for Skype4B / Lync is `"443`"" -foreground "red"
			}
			
			$PortString = $PortTextBox.Text
			$HexPort = ConvertToHex -InputString $PortString
			$HexPort = $HexPort.ToUpper()
			Write-Host "PORT: $HexPort"
			
			if($ScopeDropDownBox.Text -eq "Server Scope")
			{
				#WEB SERVER PORT
				$OutputVariable = netsh dhcp server add optiondef 4 WebServerPort Binary 0 Vendor=$VendorClassNameString comment="Web Server Port" 2>&1 | Out-String
				Write-Host "COMMAND: netsh dhcp server add optiondef 4 WebServerPort Binary 0 Vendor=$VendorClassNameString comment=`"Web Server Port`"" -foreground "green"
				Write-Host "RESULT: $OutputVariable"
				$OutputVariable = netsh dhcp server set optionvalue 4 Binary vendor=$VendorClassNameString $HexPort 2>&1 | Out-String
				Write-Host "COMMAND: netsh dhcp server set optionvalue 4 Binary vendor=$VendorClassNameString $HexPort" -foreground "green"
				Write-Host "RESULT: $OutputVariable"
			}
			else
			{
				#WEB SERVER PORT
				$OutputVariable = netsh dhcp server add optiondef 4 WebServerPort Binary 0 Vendor=$VendorClassNameString comment="Web Server Port" 2>&1 | Out-String
				Write-Host "COMMAND: netsh dhcp server add optiondef 4 WebServerPort Binary 0 Vendor=$VendorClassNameString comment=`"Web Server Port`"" -foreground "green"
				Write-Host "RESULT: $OutputVariable"
				$OutputVariable = netsh dhcp server scope $Scope set optionvalue 4 Binary vendor=$VendorClassNameString $HexPort 2>&1 | Out-String
				Write-Host "COMMAND: netsh dhcp server scope $Scope set optionvalue 4 Binary vendor=$VendorClassNameString $HexPort" -foreground "green"
				Write-Host "RESULT: $OutputVariable"
			}
			Write-Host "---------------------------------------------------------------------------------------------"
			Write-Host
		}
		else
		{
			Write-Host "ERROR: No web service port has been supplied." -foreground "red"
			#$StatusLabel.Text = "ERROR: No Web service URL has been supplied."
		}
	}
	else
	{
		Write-Host "INFO: Web service port not seleted." -foreground "yellow"
	}
	
	if($ServiceURLCheckBox.Checked)
	{
		if($ServiceURLTextBox.Text -ne "")
		{	
			Write-Host "---------------------------------------------------------------------------------------------"
			if($ServiceURLTextBox.Text -ne "/CertProv/CertProvisioningService.svc")
			{
				Write-Host "WARNING: You are using a non-standard setting for the Cert Provisioning URL. The standard setting for Skype4B / Lync is `"/CertProv/CertProvisioningService.svc`"" -foreground "red"
			}
			$ServiceURLString = $ServiceURLTextBox.Text
			$HexServiceURL = ConvertToHex -InputString $ServiceURLString
			$HexServiceURL = $HexServiceURL.ToUpper()
			Write-Host "SERVICE URL: $HexServiceURL"
			
			if($ScopeDropDownBox.Text -eq "Server Scope")
			{			
				$OutputVariable = netsh dhcp server add optiondef 5 CertProvRelPath Binary 0 Vendor=$VendorClassNameString comment="Cert Prov Relative Path" 2>&1 | Out-String
				Write-Host "COMMAND: netsh dhcp server add optiondef 5 CertProvRelPath Binary 0 Vendor=$VendorClassNameString comment=`"Cert Prov Relative Path`"" -foreground "green"
				Write-Host "RESULT: $OutputVariable"
				$OutputVariable = netsh dhcp server set optionvalue 5 Binary vendor=$VendorClassNameString $HexServiceURL 2>&1 | Out-String
				Write-Host "COMMAND: netsh dhcp server set optionvalue 5 Binary vendor=$VendorClassNameString $HexServiceURL" -foreground "green"
				Write-Host "RESULT: $OutputVariable"
			}
			else
			{
				$OutputVariable = netsh dhcp server add optiondef 5 CertProvRelPath Binary 0 Vendor=$VendorClassNameString comment="Cert Prov Relative Path" 2>&1 | Out-String
				Write-Host "COMMAND: netsh dhcp server add optiondef 5 CertProvRelPath Binary 0 Vendor=$VendorClassNameString comment=`"Cert Prov Relative Path`"" -foreground "green"
				Write-Host "RESULT: $OutputVariable"
				$OutputVariable = netsh dhcp server scope $Scope set optionvalue 5 Binary vendor=$VendorClassNameString $HexServiceURL 2>&1 | Out-String
				Write-Host "COMMAND: netsh dhcp server scope $Scope set optionvalue 5 Binary vendor=$VendorClassNameString $HexServiceURL" -foreground "green"
				Write-Host "RESULT: $OutputVariable"
			}
			Write-Host "---------------------------------------------------------------------------------------------"
			Write-Host
		}
		else
		{
			Write-Host "ERROR: No Cert Prov URL has been supplied." -foreground "red"
			#$StatusLabel.Text = "ERROR: No Web service URL has been supplied."
		}
	}
	else
	{
		Write-Host "INFO: Cert Prov URL not seleted." -foreground "yellow"
	}
}


function ConvertToHex([string] $InputString)
{
	$enc = [system.Text.Encoding]::ASCII 
	$data1 = $enc.GetBytes($InputString)

	[string]$data1String = $data1
	$dataArray = $data1String.Split(" ")

	$finalString = ""
	foreach($byte in $dataArray)
	{
		$hexString = [Convert]::ToString($byte, 16)
		$finalString += $hexString
	}
	Return $finalString 
}


function ConvertToASCII([string] $InputString)
{
	$finalString = ""
	for($i=0; $i -le ($InputString.length - 1); $i=$i+2)
	{
		$charString = $InputString.Substring($i,2)
		$char = ([CHAR][BYTE]([CONVERT]::toint16($charString,16)))
		$finalString += $char
	}
	return $finalString
}

function RemoveDHCP
{
	$Scope = $ScopeDropDownBox.Text
	$VendorClassNameString = "MSUCClient"
	
	if($PoolCheckBox.Checked)
	{
		if($ScopeDropDownBox.Text -eq "Server Scope")
		{
			$OutputVariable = netsh dhcp server delete optionvalue 120 2>&1 | Out-String
			Write-Host "COMMAND: netsh dhcp server delete optionvalue 120" -foreground "green"
			Write-Host "RESULT: $OutputVariable"
		}
		else
		{
			$OutputVariable = netsh dhcp server scope $Scope delete optionvalue 120 2>&1 | Out-String
			Write-Host "COMMAND: netsh dhcp server scope $Scope delete optionvalue 120" -foreground "green"
			Write-Host "RESULT: $OutputVariable"
		}
	}
	if($WebServerCheckBox.Checked)
	{
		if($ScopeDropDownBox.Text -eq "Server Scope")
		{
			$OutputVariable = netsh dhcp server delete optionvalue 3 vendor=$VendorClassNameString 2>&1 | Out-String
			Write-Host "COMMAND: netsh dhcp server scope $Scope delete optionvalue 3 vendor=$VendorClassNameString" -foreground "green"
			Write-Host "RESULT: $OutputVariable"
		}
		else
		{
			$OutputVariable = netsh dhcp server scope $Scope delete optionvalue 3 vendor=$VendorClassNameString 2>&1 | Out-String
			Write-Host "COMMAND: netsh dhcp server scope $Scope delete optionvalue 3 vendor=$VendorClassNameString" -foreground "green"
			Write-Host "RESULT: $OutputVariable"
		}
	}
	if($VendorClassCheckBox.Checked)
	{
		if($ScopeDropDownBox.Text -eq "Server Scope")
		{
			$OutputVariable = netsh dhcp server delete optionvalue 1 Vendor=$VendorClassNameString 2>&1 | Out-String
			Write-Host "COMMAND: netsh dhcp server delete optionvalue 1 Vendor=$VendorClassNameString" -foreground "green"
			Write-Host "RESULT: $OutputVariable"
		}
		else
		{
			$OutputVariable = netsh dhcp server scope $Scope delete optionvalue 1 Vendor=$VendorClassNameString 2>&1 | Out-String
			Write-Host "COMMAND: netsh dhcp server scope $Scope delete optionvalue 1 Vendor=$VendorClassNameString" -foreground "green"
			Write-Host "RESULT: $OutputVariable"
		}
	}
	if($ProtocolCheckBox.Checked)
	{
		if($ScopeDropDownBox.Text -eq "Server Scope")
		{
			$OutputVariable = netsh dhcp server delete optionvalue 2 Vendor=$VendorClassNameString 2>&1 | Out-String
			Write-Host "COMMAND: netsh dhcp server delete optionvalue 2 Vendor=$VendorClassNameString" -foreground "green"
			Write-Host "RESULT: $OutputVariable"
		}
		else
		{
			$OutputVariable = netsh dhcp server scope $Scope delete optionvalue 2 Vendor=$VendorClassNameString 2>&1 | Out-String
			Write-Host "COMMAND: netsh dhcp server scope $Scope delete optionvalue 2 Vendor=$VendorClassNameString" -foreground "green"
			Write-Host "RESULT: $OutputVariable"
		}
	}
	if($PortCheckBox.Checked)
	{
		if($ScopeDropDownBox.Text -eq "Server Scope")
		{
			$OutputVariable = netsh dhcp server delete optionvalue 4 Vendor=$VendorClassNameString 2>&1 | Out-String
			Write-Host "COMMAND: netsh dhcp server delete optionvalue 4 Vendor=$VendorClassNameString" -foreground "green"
			Write-Host "RESULT: $OutputVariable"
		}
		else
		{
			$OutputVariable = netsh dhcp server scope $Scope delete optionvalue 4 Vendor=$VendorClassNameString 2>&1 | Out-String
			Write-Host "COMMAND: netsh dhcp server scope $Scope delete optionvalue 4 Vendor=$VendorClassNameString" -foreground "green"
			Write-Host "RESULT: $OutputVariable"
		}
	}
	if($ServiceURLCheckBox.Checked)
	{
		if($ScopeDropDownBox.Text -eq "Server Scope")
		{
			$OutputVariable = netsh dhcp server delete optionvalue 5 Vendor=$VendorClassNameString 2>&1 | Out-String
			Write-Host "COMMAND: netsh dhcp server delete optionvalue 5 Vendor=$VendorClassNameString" -foreground "green"
			Write-Host "RESULT: $OutputVariable"
		}
		else
		{
			$OutputVariable = netsh dhcp server scope $Scope delete optionvalue 5 Vendor=$VendorClassNameString 2>&1 | Out-String
			Write-Host "COMMAND: netsh dhcp server scope $Scope delete optionvalue 5 Vendor=$VendorClassNameString" -foreground "green"
			Write-Host "RESULT: $OutputVariable"
		}
	}
	
}

function Export-Cisco
{
	#EXAMPLE
	#option 120 hex 00066C796E637365056B6C6F7564036E657400
	#option 43 hex 010C4D532D55432D436C69656E740205687474707303106C796E6373652E6B6C6F75642E6E65740403
	$FQDNError = $false
	
	$PoolText = $PoolTextBox.text
	if($PoolText -imatch "(?=^.{4,253}$)(^((?!-)[a-zA-Z0-9-]{0,62}[a-zA-Z0-9]\.)+[a-zA-Z]{2,63}$)")
	{
		
		$PoolString = $PoolText
		
		#CREATE OPTION 120 FQDN ENCODING
		$PoolStringSplit = $PoolString.Split(".")
		$HexPool = "00"  #ENCODING IS FQDN
		
		foreach($Section in $PoolStringSplit)
		{
			$NoOfChars = $Section.length
			
			$SectionHex = ConvertToHex -InputString $Section
			
			$NumberOfCharsHex = "{0:X2}" -f $NoOfChars
			$HexPool += "${NumberOfCharsHex}$SectionHex".ToUpper()
		}
		$HexPool = "${HexPool}00"
		Write-Host "INFO: Cisco Export Option 120: $HexPool" -foreground "yellow"
	}
	else
	{
		Write-Host "ERROR: The SIP Pool is not an FQDN." -foreground "red"
		$StatusLabel.Text = "ERROR: The SIP Pool is not an FQDN."
		$FQDNError = $true
	}
	
	$WebServerText = $WebServerTextBox.text
	if($WebServerText -imatch "(?=^.{4,253}$)(^((?!-)[a-zA-Z0-9-]{0,62}[a-zA-Z0-9]\.)+[a-zA-Z]{2,63}$)")
	{
		$NoOfChars = $WebServerText.length
		$SectionHex = ConvertToHex -InputString $WebServerText
		$NumberOfCharsHex = "{0:X2}" -f $NoOfChars
		$HexWebServer = "03${NumberOfCharsHex}$SectionHex".ToUpper()
	}
	else
	{
		Write-Host "ERROR: The Web Server is not an FQDN." -foreground "red"
		$StatusLabel.Text = "ERROR: The SIP Pool is not an FQDN."
		$FQDNError = $true
	}
	
	#EXIT IF FQDN CHECK FAILS
	if(!$FQDNError)
	{
	
		$VendorClassText = $VendorClassTextBox.text
		$HexVendorClass = ""
		if($VendorClassText -ne "")
		{
			if($VendorClassText -ne "MS-UC-Client")
			{
				Write-Host "WARNING: You are using a non-standard setting for the DHCP Vendor Class name. The standard setting for Skype4B / Lync is `"MS-UC-Client`"" -foreground "red"
			}
			$NoOfChars = $VendorClassText.length
			$SectionHex = ConvertToHex -InputString $VendorClassText
			$NumberOfCharsHex = "{0:X2}" -f $NoOfChars
			$HexVendorClass = "01${NumberOfCharsHex}$SectionHex".ToUpper()
		}
		
		$ProtocolText = $ProtocolTextBox.text
		$HexProtocol = ""
		if($ProtocolText -ne "")
		{
			if($ProtocolText -ne "https")
			{
				Write-Host "WARNING: You are using a non-standard setting for the protocol type. The standard setting for Skype4B / Lync is `"https`"" -foreground "red"
			}
			$NoOfChars = $ProtocolText.length
			$SectionHex = ConvertToHex -InputString $ProtocolText
			$NumberOfCharsHex = "{0:X2}" -f $NoOfChars
			$HexProtocol = "02${NumberOfCharsHex}$SectionHex".ToUpper()
		}
		
		$PortText = $PortTextBox.text
		$HexPort = ""
		if($PortText -ne "")
		{
			if($PortText -ne "443")
			{
				Write-Host "WARNING: You are using a non-standard setting for the port number. The standard setting for Skype4B / Lync is `"443`"" -foreground "red"
			}
			$NoOfChars = $PortText.length
			$SectionHex = ConvertToHex -InputString $PortText
			$NumberOfCharsHex = "{0:X2}" -f $NoOfChars
			$HexPort = "04${NumberOfCharsHex}$SectionHex".ToUpper()
		}
		
		$ServiceURLText = $ServiceURLTextBox.text 
		$HexServiceURL = ""
		if($ServiceURLText -ne "")
		{
			if($ServiceURLText -ne "/CertProv/CertProvisioningService.svc")
			{
				Write-Host "WARNING: You are using a non-standard setting for the Cert Provisioning URL. The standard setting for Skype4B / Lync is `"/CertProv/CertProvisioningService.svc`"" -foreground "red"
			}
			$NoOfChars = $ServiceURLText.length
			$SectionHex = ConvertToHex -InputString $ServiceURLText
			$NumberOfCharsHex = "{0:X2}" -f $NoOfChars
			$HexServiceURL = "05${NumberOfCharsHex}$SectionHex".ToUpper()
		}
		
		$Option43Cisco = "${HexVendorClass}${HexProtocol}${HexWebServer}${HexPort}${HexServiceURL}"
		Write-Host "INFO: Cisco Export Option 43: $Option43Cisco" -foreground "yellow"
		
		Write-Host
		Write-Host "Cisco Settings:"
		Write-Host
		Write-Host "option 43 hex ${Option43Cisco}"
		Write-Host
		Write-Host "option 120 hex ${HexPool}"
		Write-Host

		
		if($CiscoExportAvailable)
		{
			$filename = ""
			Write-Host "INFO: Exporting..." -foreground "yellow"
			[string] $pathVar = "c:\"
			[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
			$objDialog = New-Object System.Windows.Forms.SaveFileDialog
			$objDialog.FileName = "CiscoDHCPConfig.txt"
			$objDialog.Title = "Export File Name"
			$objDialog.CheckFileExists = $false
			$Show = $objDialog.ShowDialog()
			if ($Show -eq "OK")
			{
				[string] $filename = $objDialog.FileName
			}
				
			if($filename -ne "")
			{
				$fileContent = "!OPTIONS FOR CISCO DHCP CONFIG`r`n"
				$fileContent += "option 43 hex ${Option43Cisco}`r`n"
				$fileContent += "option 120 hex ${HexPool}`r`n"
				
				$fileContent | out-file -Encoding UTF8 -FilePath $filename -Force
				Write-Host "INFO: Completed Cisco Export." -foreground "yellow"
			}
			else
			{
				Write-Host "INFO: Cancelled Cisco Export." -foreground "yellow"
			}
		}
		else
		{
			Write-Host
			Write-Host "INFO: Output to file is not available on Powershell version 2. Output will be displayed above." -foreground "yellow"

		}
	}
	
}

$numberOfItems = $ScopeDropDownBox.Items.count
if($numberOfItems -gt 0)
{
	$ScopeDropDownBox.SelectedIndex = 0
}


# Activate the form ============================================================
$objForm.Add_Shown({$objForm.Activate()})
[void] $objForm.ShowDialog()	



# SIG # Begin signature block
# MIIcXAYJKoZIhvcNAQcCoIIcTTCCHEkCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU7E0X0Hm3gSuGIHPzO60XbB8V
# WqaggheLMIIFFDCCA/ygAwIBAgIQC7/jb7qrV/+uuRoaboA8vjANBgkqhkiG9w0B
# AQsFADByMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
# VQQLExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFz
# c3VyZWQgSUQgQ29kZSBTaWduaW5nIENBMB4XDTE1MDYyMjAwMDAwMFoXDTE2MDgy
# NTEyMDAwMFowWzELMAkGA1UEBhMCQVUxDDAKBgNVBAgTA1ZJQzEQMA4GA1UEBxMH
# TWl0Y2hhbTEVMBMGA1UEChMMSmFtZXMgQ3Vzc2VuMRUwEwYDVQQDEwxKYW1lcyBD
# dXNzZW4wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCoITj78CkXvlTw
# OquYWSDCpm4CxAgJfi4CdvJXtwnK4q/BeURGUi8AepOluIF12pQRrTAqLyfy+hJf
# kk2lE3n0Z5qaAmK3w3PjXf7yKem8vVttC1QknMpfkvW0Lu/k6TxcNKimSlVk86bs
# W5qw1Ql2mClLjRRL+5Nz9qM8F4QMzz1P1dH6oDWhhDetk2NLMd5JbrMUMj9QEsu5
# gh5zGBn4fdEcW9ujZSxU6bxGTzZVNtcCWcr+9r/MpDdFl+ExwpHl2iIqVdvO8OBI
# TZE5xNCkbUn4enWhJi1elhI0TMZbIfy9X729aSILz5+0KgHQLTzU6oYDoTeezgtU
# TO4CzhEvAgMBAAGjggG7MIIBtzAfBgNVHSMEGDAWgBRaxLl7KgqjpepxA8Bg+S32
# ZXUOWDAdBgNVHQ4EFgQUSuaSOXtLjdMP2pIoh+MLe/KqZX8wDgYDVR0PAQH/BAQD
# AgeAMBMGA1UdJQQMMAoGCCsGAQUFBwMDMHcGA1UdHwRwMG4wNaAzoDGGL2h0dHA6
# Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9zaGEyLWFzc3VyZWQtY3MtZzEuY3JsMDWgM6Ax
# hi9odHRwOi8vY3JsNC5kaWdpY2VydC5jb20vc2hhMi1hc3N1cmVkLWNzLWcxLmNy
# bDBCBgNVHSAEOzA5MDcGCWCGSAGG/WwDATAqMCgGCCsGAQUFBwIBFhxodHRwczov
# L3d3dy5kaWdpY2VydC5jb20vQ1BTMIGEBggrBgEFBQcBAQR4MHYwJAYIKwYBBQUH
# MAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBOBggrBgEFBQcwAoZCaHR0cDov
# L2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0U0hBMkFzc3VyZWRJRENvZGVT
# aWduaW5nQ0EuY3J0MAwGA1UdEwEB/wQCMAAwDQYJKoZIhvcNAQELBQADggEBANCV
# INu/j0Scmsg7BMcSIaQP07XH3PP20Z+U30L5/kmd7c8arjPqk1mavKmdYksixh3V
# RCneKVwalXit4FXPQKKx+teTh6tgkr6HlXLxIPBVVYRi71CAY4NyhsmNHg2ky9X9
# hNVzs2sG5215okFs6RI1rCb+iM6fSBxbmHGldzocw+uH8xHoOF3S2eVlsEvDPsgA
# W91+dKdgajFjb97HWdpzaku022HnHyCnqa9rD70S7gFhgu9AQK4VvhcIqZZqI8Ie
# CFaLPxP/2b2RN+QEJLw2foWRkPWRoWi/D8Xjqaneb9u1t+eZ1gDN+Wgj5W6sx1VF
# 1KThdJ7OKvUuFVfTFGkwggUwMIIEGKADAgECAhAECRgbX9W7ZnVTQ7VvlVAIMA0G
# CSqGSIb3DQEBCwUAMGUxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJ
# bmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xJDAiBgNVBAMTG0RpZ2lDZXJ0
# IEFzc3VyZWQgSUQgUm9vdCBDQTAeFw0xMzEwMjIxMjAwMDBaFw0yODEwMjIxMjAw
# MDBaMHIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNV
# BAsTEHd3dy5kaWdpY2VydC5jb20xMTAvBgNVBAMTKERpZ2lDZXJ0IFNIQTIgQXNz
# dXJlZCBJRCBDb2RlIFNpZ25pbmcgQ0EwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAw
# ggEKAoIBAQD407Mcfw4Rr2d3B9MLMUkZz9D7RZmxOttE9X/lqJ3bMtdx6nadBS63
# j/qSQ8Cl+YnUNxnXtqrwnIal2CWsDnkoOn7p0WfTxvspJ8fTeyOU5JEjlpB3gvmh
# hCNmElQzUHSxKCa7JGnCwlLyFGeKiUXULaGj6YgsIJWuHEqHCN8M9eJNYBi+qsSy
# rnAxZjNxPqxwoqvOf+l8y5Kh5TsxHM/q8grkV7tKtel05iv+bMt+dDk2DZDv5LVO
# pKnqagqrhPOsZ061xPeM0SAlI+sIZD5SlsHyDxL0xY4PwaLoLFH3c7y9hbFig3NB
# ggfkOItqcyDQD2RzPJ6fpjOp/RnfJZPRAgMBAAGjggHNMIIByTASBgNVHRMBAf8E
# CDAGAQH/AgEAMA4GA1UdDwEB/wQEAwIBhjATBgNVHSUEDDAKBggrBgEFBQcDAzB5
# BggrBgEFBQcBAQRtMGswJAYIKwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0
# LmNvbTBDBggrBgEFBQcwAoY3aHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0Rp
# Z2lDZXJ0QXNzdXJlZElEUm9vdENBLmNydDCBgQYDVR0fBHoweDA6oDigNoY0aHR0
# cDovL2NybDQuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNy
# bDA6oDigNoY0aHR0cDovL2NybDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJl
# ZElEUm9vdENBLmNybDBPBgNVHSAESDBGMDgGCmCGSAGG/WwAAgQwKjAoBggrBgEF
# BQcCARYcaHR0cHM6Ly93d3cuZGlnaWNlcnQuY29tL0NQUzAKBghghkgBhv1sAzAd
# BgNVHQ4EFgQUWsS5eyoKo6XqcQPAYPkt9mV1DlgwHwYDVR0jBBgwFoAUReuir/SS
# y4IxLVGLp6chnfNtyA8wDQYJKoZIhvcNAQELBQADggEBAD7sDVoks/Mi0RXILHwl
# KXaoHV0cLToaxO8wYdd+C2D9wz0PxK+L/e8q3yBVN7Dh9tGSdQ9RtG6ljlriXiSB
# ThCk7j9xjmMOE0ut119EefM2FAaK95xGTlz/kLEbBw6RFfu6r7VRwo0kriTGxycq
# oSkoGjpxKAI8LpGjwCUR4pwUR6F6aGivm6dcIFzZcbEMj7uo+MUSaJ/PQMtARKUT
# 8OZkDCUIQjKyNookAv4vcn4c10lFluhZHen6dGRrsutmQ9qzsIzV6Q3d9gEgzpkx
# Yz0IGhizgZtPxpMQBvwHgfqL2vmCSfdibqFT+hKUGIUukpHqaGxEMrJmoecYpJpk
# Ue8wggZqMIIFUqADAgECAhADAZoCOv9YsWvW1ermF/BmMA0GCSqGSIb3DQEBBQUA
# MGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsT
# EHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lDZXJ0IEFzc3VyZWQgSUQg
# Q0EtMTAeFw0xNDEwMjIwMDAwMDBaFw0yNDEwMjIwMDAwMDBaMEcxCzAJBgNVBAYT
# AlVTMREwDwYDVQQKEwhEaWdpQ2VydDElMCMGA1UEAxMcRGlnaUNlcnQgVGltZXN0
# YW1wIFJlc3BvbmRlcjCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAKNk
# Xfx8s+CCNeDg9sYq5kl1O8xu4FOpnx9kWeZ8a39rjJ1V+JLjntVaY1sCSVDZg85v
# Zu7dy4XpX6X51Id0iEQ7Gcnl9ZGfxhQ5rCTqqEsskYnMXij0ZLZQt/USs3OWCmej
# vmGfrvP9Enh1DqZbFP1FI46GRFV9GIYFjFWHeUhG98oOjafeTl/iqLYtWQJhiGFy
# GGi5uHzu5uc0LzF3gTAfuzYBje8n4/ea8EwxZI3j6/oZh6h+z+yMDDZbesF6uHjH
# yQYuRhDIjegEYNu8c3T6Ttj+qkDxss5wRoPp2kChWTrZFQlXmVYwk/PJYczQCMxr
# 7GJCkawCwO+k8IkRj3cCAwEAAaOCAzUwggMxMA4GA1UdDwEB/wQEAwIHgDAMBgNV
# HRMBAf8EAjAAMBYGA1UdJQEB/wQMMAoGCCsGAQUFBwMIMIIBvwYDVR0gBIIBtjCC
# AbIwggGhBglghkgBhv1sBwEwggGSMCgGCCsGAQUFBwIBFhxodHRwczovL3d3dy5k
# aWdpY2VydC5jb20vQ1BTMIIBZAYIKwYBBQUHAgIwggFWHoIBUgBBAG4AeQAgAHUA
# cwBlACAAbwBmACAAdABoAGkAcwAgAEMAZQByAHQAaQBmAGkAYwBhAHQAZQAgAGMA
# bwBuAHMAdABpAHQAdQB0AGUAcwAgAGEAYwBjAGUAcAB0AGEAbgBjAGUAIABvAGYA
# IAB0AGgAZQAgAEQAaQBnAGkAQwBlAHIAdAAgAEMAUAAvAEMAUABTACAAYQBuAGQA
# IAB0AGgAZQAgAFIAZQBsAHkAaQBuAGcAIABQAGEAcgB0AHkAIABBAGcAcgBlAGUA
# bQBlAG4AdAAgAHcAaABpAGMAaAAgAGwAaQBtAGkAdAAgAGwAaQBhAGIAaQBsAGkA
# dAB5ACAAYQBuAGQAIABhAHIAZQAgAGkAbgBjAG8AcgBwAG8AcgBhAHQAZQBkACAA
# aABlAHIAZQBpAG4AIABiAHkAIAByAGUAZgBlAHIAZQBuAGMAZQAuMAsGCWCGSAGG
# /WwDFTAfBgNVHSMEGDAWgBQVABIrE5iymQftHt+ivlcNK2cCzTAdBgNVHQ4EFgQU
# YVpNJLZJMp1KKnkag0v0HonByn0wfQYDVR0fBHYwdDA4oDagNIYyaHR0cDovL2Ny
# bDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEQ0EtMS5jcmwwOKA2oDSG
# Mmh0dHA6Ly9jcmw0LmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRENBLTEu
# Y3JsMHcGCCsGAQUFBwEBBGswaTAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGln
# aWNlcnQuY29tMEEGCCsGAQUFBzAChjVodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5j
# b20vRGlnaUNlcnRBc3N1cmVkSURDQS0xLmNydDANBgkqhkiG9w0BAQUFAAOCAQEA
# nSV+GzNNsiaBXJuGziMgD4CH5Yj//7HUaiwx7ToXGXEXzakbvFoWOQCd42yE5FpA
# +94GAYw3+puxnSR+/iCkV61bt5qwYCbqaVchXTQvH3Gwg5QZBWs1kBCge5fH9j/n
# 4hFBpr1i2fAnPTgdKG86Ugnw7HBi02JLsOBzppLA044x2C/jbRcTBu7kA7YUq/OP
# Q6dxnSHdFMoVXZJB2vkPgdGZdA0mxA5/G7X1oPHGdwYoFenYk+VVFvC7Cqsc21xI
# J2bIo4sKHOWV2q7ELlmgYd3a822iYemKC23sEhi991VUQAOSK2vCUcIKSK+w1G7g
# 9BQKOhvjjz3Kr2qNe9zYRDCCBs0wggW1oAMCAQICEAb9+QOWA63qAArrPye7uhsw
# DQYJKoZIhvcNAQEFBQAwZTELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0
# IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEkMCIGA1UEAxMbRGlnaUNl
# cnQgQXNzdXJlZCBJRCBSb290IENBMB4XDTA2MTExMDAwMDAwMFoXDTIxMTExMDAw
# MDAwMFowYjELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcG
# A1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEhMB8GA1UEAxMYRGlnaUNlcnQgQXNzdXJl
# ZCBJRCBDQS0xMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEA6IItmfnK
# wkKVpYBzQHDSnlZUXKnE0kEGj8kz/E1FkVyBn+0snPgWWd+etSQVwpi5tHdJ3InE
# Ctqvy15r7a2wcTHrzzpADEZNk+yLejYIA6sMNP4YSYL+x8cxSIB8HqIPkg5QycaH
# 6zY/2DDD/6b3+6LNb3Mj/qxWBZDwMiEWicZwiPkFl32jx0PdAug7Pe2xQaPtP77b
# lUjE7h6z8rwMK5nQxl0SQoHhg26Ccz8mSxSQrllmCsSNvtLOBq6thG9IhJtPQLnx
# TPKvmPv2zkBdXPao8S+v7Iki8msYZbHBc63X8djPHgp0XEK4aH631XcKJ1Z8D2Kk
# PzIUYJX9BwSiCQIDAQABo4IDejCCA3YwDgYDVR0PAQH/BAQDAgGGMDsGA1UdJQQ0
# MDIGCCsGAQUFBwMBBggrBgEFBQcDAgYIKwYBBQUHAwMGCCsGAQUFBwMEBggrBgEF
# BQcDCDCCAdIGA1UdIASCAckwggHFMIIBtAYKYIZIAYb9bAABBDCCAaQwOgYIKwYB
# BQUHAgEWLmh0dHA6Ly93d3cuZGlnaWNlcnQuY29tL3NzbC1jcHMtcmVwb3NpdG9y
# eS5odG0wggFkBggrBgEFBQcCAjCCAVYeggFSAEEAbgB5ACAAdQBzAGUAIABvAGYA
# IAB0AGgAaQBzACAAQwBlAHIAdABpAGYAaQBjAGEAdABlACAAYwBvAG4AcwB0AGkA
# dAB1AHQAZQBzACAAYQBjAGMAZQBwAHQAYQBuAGMAZQAgAG8AZgAgAHQAaABlACAA
# RABpAGcAaQBDAGUAcgB0ACAAQwBQAC8AQwBQAFMAIABhAG4AZAAgAHQAaABlACAA
# UgBlAGwAeQBpAG4AZwAgAFAAYQByAHQAeQAgAEEAZwByAGUAZQBtAGUAbgB0ACAA
# dwBoAGkAYwBoACAAbABpAG0AaQB0ACAAbABpAGEAYgBpAGwAaQB0AHkAIABhAG4A
# ZAAgAGEAcgBlACAAaQBuAGMAbwByAHAAbwByAGEAdABlAGQAIABoAGUAcgBlAGkA
# bgAgAGIAeQAgAHIAZQBmAGUAcgBlAG4AYwBlAC4wCwYJYIZIAYb9bAMVMBIGA1Ud
# EwEB/wQIMAYBAf8CAQAweQYIKwYBBQUHAQEEbTBrMCQGCCsGAQUFBzABhhhodHRw
# Oi8vb2NzcC5kaWdpY2VydC5jb20wQwYIKwYBBQUHMAKGN2h0dHA6Ly9jYWNlcnRz
# LmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcnQwgYEGA1Ud
# HwR6MHgwOqA4oDaGNGh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFz
# c3VyZWRJRFJvb3RDQS5jcmwwOqA4oDaGNGh0dHA6Ly9jcmw0LmRpZ2ljZXJ0LmNv
# bS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcmwwHQYDVR0OBBYEFBUAEisTmLKZ
# B+0e36K+Vw0rZwLNMB8GA1UdIwQYMBaAFEXroq/0ksuCMS1Ri6enIZ3zbcgPMA0G
# CSqGSIb3DQEBBQUAA4IBAQBGUD7Jtygkpzgdtlspr1LPUukxR6tWXHvVDQtBs+/s
# dR90OPKyXGGinJXDUOSCuSPRujqGcq04eKx1XRcXNHJHhZRW0eu7NoR3zCSl8wQZ
# Vann4+erYs37iy2QwsDStZS9Xk+xBdIOPRqpFFumhjFiqKgz5Js5p8T1zh14dpQl
# c+Qqq8+cdkvtX8JLFuRLcEwAiR78xXm8TBJX/l/hHrwCXaj++wc4Tw3GXZG5D2dF
# zdaD7eeSDY2xaYxP+1ngIw/Sqq4AfO6cQg7PkdcntxbuD8O9fAqg7iwIVYUiuOsY
# Gk38KiGtSTGDR5V3cdyxG0tLHBCcdxTBnU8vWpUIKRAmMYIEOzCCBDcCAQEwgYYw
# cjELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQ
# d3d3LmRpZ2ljZXJ0LmNvbTExMC8GA1UEAxMoRGlnaUNlcnQgU0hBMiBBc3N1cmVk
# IElEIENvZGUgU2lnbmluZyBDQQIQC7/jb7qrV/+uuRoaboA8vjAJBgUrDgMCGgUA
# oHgwGAYKKwYBBAGCNwIBDDEKMAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYKKwYB
# BAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0B
# CQQxFgQUmWtBMNY5p8XXE61ZRLS7WPTh8lowDQYJKoZIhvcNAQEBBQAEggEAkzSK
# yA7Y5Bd1LQmDaZ9WbTNOlTjthZfzP1VB9hpsGYZQNI1B8OOC5qgGuhTnSlgDGEXz
# 9vIXoCz5x+x92AVm3hZgzh1xp/262O+R/0COYEyMSUyeA2C8VtRIV1okHfBKvI0E
# R8o0LTpkBTwWbV9ppcrqu7yYi2xR68lnYxYV6AkQS7m+cccECnHT6lw3rGDVcUZt
# olIjheFJUAFMuJXYsx3LSBCrA3EUe+lLHrf2b8JjBLH7H+3MKKfATxGLqXqaAYlZ
# qFDP9jOb+s0I3WSDskN0yv0dZUw+lbfXqOQLyL9cHeelIpf8wjVYFaXINACUPtxI
# o8lNCCGBmcU0LNMv6qGCAg8wggILBgkqhkiG9w0BCQYxggH8MIIB+AIBATB2MGIx
# CzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3
# dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lDZXJ0IEFzc3VyZWQgSUQgQ0Et
# MQIQAwGaAjr/WLFr1tXq5hfwZjAJBgUrDgMCGgUAoF0wGAYJKoZIhvcNAQkDMQsG
# CSqGSIb3DQEHATAcBgkqhkiG9w0BCQUxDxcNMTUxMTMwMDMxMjExWjAjBgkqhkiG
# 9w0BCQQxFgQUqqU45lK5VroWJPuGogOw4rl2qzQwDQYJKoZIhvcNAQEBBQAEggEA
# EWiVUDAPq20YHpCAp71A3sOWfWI/XemVZQZGMtK9C64hxsSMvG0eXLEGAPgOsPN/
# JjeFW7ZhEOr8VkI6l5AhOwKpLIbaufAECASb93stGjveZ/RPBL9lYDirRkKOSLQN
# zLpEIOO9pysnIZBC4tBQu0r0eWRp/51EH7fmcCwhoU8YJCFLzHh/ZNDKJ+SAVbAO
# PY7QH3ECBunsKQ/HE1nWb6+EqI5VQzHzMcdCPCxN04zq2lIGGGt1oiRokShrd5rS
# +ANjIPLVi2GqWWMBSxcR9lV3vxP69yR7vDI5CIhvvde9Mdhu8fdJ8UTnTg+aEL9v
# yZ2NnnQimFWR/m1AIFI3aQ==
# SIG # End signature block
