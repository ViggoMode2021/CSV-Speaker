$CSV = Import-CSV -Path "C:\Users\ryans\Desktop\Sample.CSV"

add-type -assemblyname system.speech

Add-Type -AssemblyName System.speech
$speak = New-Object System.Speech.Synthesis.SpeechSynthesizer

foreach ($Line in $CSV) 
{
	$Time = $Line.Time
	$Student = $Line.Student

	Write-Host $Time
	Write-Host $Student
	$speak.Speak($Time)
	$speak.Speak($Student)

}
