Import-Module ActiveDirectory

Add-Type -AssemblyName System.Speech

Add-Type -AssemblyName System.Speech
$speak = New-Object System.Speech.Synthesis.SpeechSynthesizer

$All_AD_Students_Alphabetized = Import-CSV -Path "C:\Users\rviglione\Desktop\CSV-Comparison\All-AD-Students-Alphabetized.csv"

$All_PS_Students_Alphabetized = Import-CSV -Path "C:\Users\rviglione\Desktop\CSV-Comparison\All-Student-Numbers.csv"

$All_PS_Students_Alphabetized_Get_Content = Get-Content -Path "C:\Users\rviglione\Desktop\CSV-Comparison\All-Student-Numbers.csv"

#$speak.Speak("There are $Students_With_Number students with employee numbers in Active Directory and 611 students total that need employee numbers")

foreach ($Line in $All_AD_Students_Alphabetized) 
{
	$Name = $Line.Name
	$Employee_Number = $Line.EmployeeNumber
    $SAM_Account_Name = $Line.SamAccountName

	Write-Host $Name
	Write-Host $Employee_Number
	$speak.Speak($Name)
	$speak.Speak($Employee_Number)

    if($Employee_Number -eq ""){

    $speak.Speak("$Name does not have an employee number, looking it up in other CSV")

    $Name_Find = $All_PS_Students_Alphabetized_Get_Content | Select-String $Name | Out-String

    $Position = $Name_Find.IndexOf(",")

    $Name_In_Other_CSV = $Name_Find.Substring(0, $Position)

    $Employee_Number_In_Other_CSV = $Name_Find.Substring($Position+1)

    Write-Host "$Name_In_Other_CSV is like $Name"

    if($Name_In_Other_CSV -match $Name){

    $speak.Speak("$Name is in All PowerSchool students CSV, adding now")

    Write-Host "$Name is in All PowerSchool students CSV, adding $Employee_Number_In_Other_CSV as employee number" -ForegroundColor "Green"

    Write-Host "$SAM_Account_Name"

    $speak.Speak("Adding $Employee_Number_In_Other_CSV to $Name") 

    Set-ADUser -Identity $SAM_Account_Name -EmployeeNumber $Employee_Number_In_Other_CSV

    $Employee_Number_Input = Get-ADUser -Filter "SamAccountName -like '$Sam_Account_Name'" -Properties EmployeeNumber | Select-Object -expand EmployeeNumber

    Write-Host "Employee Number changed to $Employee_Number_Input for $Name" -ForegroundColor Green
    
    }

    <#
    
    Set-ADUser -Identity $SAM_Account_Name -EmployeeNumber $Employee_Number

    $Employee_Number_Input = Get-ADUser -Filter "SamAccountName -like '$Sam_Account_Name'" -Properties EmployeeNumber | Select-Object -expand EmployeeNumber

    Write-Host "Employee Number changed to $Employee_Number_Input for $Sam_Account_Name" -ForegroundColor Green

    #>
    }

    }

    $All_PS_Students_Alphabetized.student_number

    if($All_PS_Students_Alphabetized -notcontains $Name){

    $speak.Speak("$Name is not in All PowerSchool students CSV")

    
    }

<#


#>
