# OfficeCracker
PowerShell script for attempting to crack the password of an Office document (Word, Excel, or PowerPoint) through the use of common password lists.

OfficeCracker supports password guessing against the following document types at this time:

<ul>
<li>Microsoft Word</li>
<li>Microsoft Excel</li>
<li>Microsoft PowerPoint</li>
</ul>

Requirements
============
Office installation

Usage
=====
<pre>
NAME    
  Invoke-OfficeCracker
  
SYNOPSIS    
  This module loops through passwords and attempts to decrypt an Excel office file
  
    PowerSniper Function: Invoke-OfficeCracker    
    Author: Josh Berry (@codewatchorg)    
    License: BSD 3-Clause    
    Required Dependencies: None    
    Optional Dependencies: None

SYNTAX    
  Invoke-OfficeCracker [[-officetype] &lt;Object&gt;] [[-officefile] &lt;Object&gt;] [[-passlist] &lt;Object&gt;] [-verbose] [&lt;CommonParameters&gt;]

DESCRIPTION    
  This module loops through passwords and attempts to decrypt an office file.
</pre>

Example OfficeCracker.ps1 usage:
<pre>
    # Outlook Anywhere Test
    Invoke-OfficeCracker -officetype excel -officefile 'C:\Users\username\Documents\file.xlsx' -passlist 'C:\wordlist\rockyou.txt'
</pre>