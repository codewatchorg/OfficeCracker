function Invoke-OfficeCracker {
<#
  .SYNOPSIS

    This module loops through passwords and attempts to decrypt an office file

    PowerSniper Function: Invoke-OfficeCracker
    Author: Josh Berry (@codewatchorg)
    License: BSD 3-Clause
    Required Dependencies: None
    Optional Dependencies: None

  .DESCRIPTION

    This module loops through passwords and attempts to decrypt an office file.

  .PARAMETER officetype

    File type to open (should be one of either excel, word, or powerpoint).

  .PARAMETER officefile

    File to crack the password.

  .PARAMETER passlist

    Password list to use in the cracking attempts.

  .PARAMETER verbose

    The verbose parameter causes OfficeCracker to print off each attempt.


  .EXAMPLE

    C:\PS> Invoke-OfficeCracker -officetype excel -officefile 'C:\Users\username\Documents\file.xlsx' -passlist 'C:\wordlist\rockyou.txt'

    Description
    -----------
    This command will attempt to crack the Excel files password using the Rockyou password list.

#>

  # Do not report exceptions and set variables
  param($officetype, $officefile, $passlist, [switch]$verbose);

  # Set encoding to UTF8
  $EncodingForm = [System.Text.Encoding]::UTF8;

  # Function to attempt to decrypt the file given the path to the file and a password
  function DecryptFile {
    param($OType, $OFile, $OPass);

    # Set default status to failure
    $AuthStatus = "Failure";

    # Find file type and process accordingly
    If ($OType -eq "excel") {
      # Setup Excel COM object for PowerShell, don't open Excel or display notifications
      $OfficeObj = New-Object -comobject Excel.Application
      $OfficeObj.Visible = $False
      $OfficeObj.DisplayAlerts = $False

      # Try to open the file with the given password
      Try {
        # Open the file with the password
        $Workbook = $OfficeObj.Workbooks.Open($OFile,0,$True,5,$OPass);
        $Workbook.Close();
        $AuthStatus = "Success";
      } Catch {
        $AuthStatus = "Failure";
      }

      # Close down the object
      $OfficeObj.Quit();
    } ElseIf ($OType -eq "word") {
      # Setup Word COM object for PowerShell, don't open Word or display notifications
      $OfficeObj = New-Object -comobject Word.Application;
      $OfficeObj.Visible = $False;

      # Try to open the file with the given password
      Try {
        # Open the file with the password
        $Worddoc = $OfficeObj.Documents.Open($OFile,$False,$True,$False,$OPass);
        $Worddoc.Close($False);
        $AuthStatus = "Success";
      } Catch {
        $AuthStatus = "Failure";
      }

      # Close down the object
      $OfficeObj.Quit();
    } ElseIf ($OType -eq "powerpoint") {
      # Setup Word COM object for PowerShell, don't open PowerPoint or display notifications
      $OfficeObj = New-Object -comobject PowerPoint.Application

      # Try to open the file with the given password
      Try {
        # Open the file with the password
        $Preso = $OfficeObj.Presentations.Open("$($OFile)::$($OPass)::",$True,$False,$False);
        $Preso.Close();
        $AuthStatus = "Success";
      } Catch {
        $AuthStatus = "Failure";
      }

      # Close down the object
      $OfficeObj.Quit();
    } Else {
      Write-Host "Unsupported file type.";
      break;
    }

    # Return the result
    return $AuthStatus;
  }

  # Loop through passwords, combine with each user
  Get-Content $passlist | ForEach-Object -Process {
    $pass = $_;

    # Set status of decrypt to default of failure
    $DecryptStatus = "Failure";

    # Attempt to decrypt the file
    $DecryptStatus = DecryptFile -OType $officetype -OFile $officefile -OPass $pass;

    # If verbose mode is set, print off each connection attempt
    If ($verbose -eq $true) {
      Write-Host "Attempting to decrypt file $officefile using $pass";
    }

    # If decrypt succeeded, write out to prompt
    If ($DecryptStatus -eq "Success") {
      Write-Host "Decryption of file $officefile succeeded: $pass";
      break;
    }
  }
}