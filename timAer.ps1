#TimAer
#Clocking system for employee time logging
#Author: Alexander Bazzi
#Released June 18 2014

$title = "Check in/out"
$message = "Are you checking in or out?"
$in = New-Object System.Management.Automation.Host.ChoiceDescription "&In", `
    "Checks the user in"
$out = New-Object System.Management.Automation.Host.ChoiceDescription "&Out", `
    "Checks the user out"
$startLunch = New-Object System.Management.Automation.Host.ChoiceDescription "&Start Lunch", `
    "Starts the lunch"
$endLunch = New-Object System.Management.Automation.Host.ChoiceDescription "&End Lunch", `
    "Ends the lunch"
$new = New-Object System.Management.Automation.Host.ChoiceDescription "&New", `
    "Checks the user in on a new week"
$choices = [System.Management.Automation.Host.ChoiceDescription[]]($in, $out, $startLunch, $endLunch, $new, 0)
$result = $host.ui.PromptForChoice($title, $message, $choices, 0)


$timeSheet = New-Object System.Collections.Arraylist #Creates new arraylist for constructing a timeSheet object

$fileName = Import-Clixml "C:\Scripts\TimAer\Variables\fileName.xml"

switch ($result) #Just a little block of text output in the cli screen
{
  0 {"You checked in!"}
  1 {"You checked out!"}
  2 {"You went to lunch!"}
  3 {"You came from lunch!"}
  4 {"You checked in on a new week!"}
}

########################################################################################################################

Function FormatTimeOut
{
  param ([String]$timeOut)

  $timeOutHours = $timeOut.Substring(0, 2)
  $timeOutMinutes = $timeOut.Substring(3)
  [String]$timeOutMinutes = $timeOutMinutes

  if([int]$timeOutMinutes.Substring(1, 1) -lt 5)
  {
    $timeOutMinutes = $timeOutMinutes.Substring(0, 1) + 5
  }

  else
  {
    [int]$firstDigit = $timeOutMinutes.Substring(0, 1)

    if($firstDigit -ne 5)
    {
      $firstDigit++
    }

    else
    {
      $firstDigit = 0
      $timeOutHours = [int]$timeOutHours
      $timeOutHours++
    }

    [String]$timeOutMinutes = "" + $firstDigit + "0"
  }

  $timeOut = "$timeOutHours" + ":" +"$timeOutMinutes"
  return ,$timeOut
}

Function FormatTimeIn
{
  param ([String]$timeIn)

  $timeInHours = $timeIn.Substring(0, 2)
  $timeInMinutes = $timeIn.Substring(3)
  [String]$timeInMinutes = $timeInMinutes

  [int]$timeInMinutesSecondDigit = $timeInMinutes.Substring(1, 1) #Isolates the second digit of the minute variable

  if($timeInMinutesSecondDigit -ge 5) #if the second digit is greater than or equal to 5
  {
    $timeInMinutes = $timeInMinutes.Substring(0, 1) + 5 #Set the second digit to 5
  }

  else #If the second digit is less than 5
  {
    [int]$firstDigit = $timeInMinutes.Substring(0, 1)
    [String]$timeInMinutes = "" + $firstDigit + "0" #Set the second digit to zero
  }

  $timeIn = "$timeInHours" + ":" +"$timeInMinutes"
  return ,$timeIn
}

#######################################################################################################################################
if($result -eq 0) #Checking in on a new day
{
  #Registers the time of arrival and saves the variable to the path specified for later use
  $dateIn = Get-Date #Gets today's date
  $timeIn = Get-Date -format HH:mm #Arrival
  $timeInFormatted = FormatTimeIn $timeIn
  $timeInFormatted | Export-CliXml "C:\Scripts\TimAer\Variables\timeIn.xml" #Saves the variable to the path specified   

  #Adding a row to the timeSheet table. This represents the opening of a work day               
  $timeSheet.Add("<tr style=`"font-family:arial; text-align:center; background-color:#FFFFFF; `">" +
                 "<td>" + $dateIn.toShortDateString() + "</td>" +
                 "<td>" + $dateIn.DayOfWeek + "</td>" +
                 "<td>" + $timeInFormatted + "</td>")
}

elseif($result -eq 1) #Checking out of the day
{
  $timeIn = Import-Clixml "C:\Scripts\TimAer\Variables\timeIn.xml" #Arrival
  $timeOut = Get-Date -Format HH:mm #Departure
  $timeOutFormatted = FormatTimeOut $timeOut #Formatted departure

  $totalTimeLunchMath = Import-Clixml "C:\Scripts\TimAer\Variables\totalTimeLunchMath.xml"
  [String]$totalTime = New-Timespan -Start $timeIn -End $timeOutFormatted #Total time in the work day
  $totalTime = $totalTime.Substring(0, 5)

  $minutesMath = [int]$totalTime.Substring(3, 2) / 60
  $totalTimeMath = [int]$totalTime.Substring(0, 2) + $minutesMath - $totalTimeLunchMath

  $payRate = 8
  $payToday = $payRate * $totalTimeMath
  $payToday = "{0:C2}" -f $payToday

  $totalTimeLunch = Import-Clixml "C:\Scripts\TimAer\Variables\totalTimeLunch.xml"

  #Adding final values to a timeSheet row. This represents the closing of a work day
  $timeSheet.Add("<td>" + $timeOutFormatted + "</td>" + 
                "<td>" + $totalTime + " - " + $totalTimeLunch + "</td>" +
                "<td>" + $totalTimeMath + "</td>" +
                "<td>" + "$" + $payRate + "/hour" + "</td>" +
                "<td>" + $payToday + "</td></tr>")
}

elseif($result -eq 2) #Lunch starts
{
  $timeInLunch = Get-Date -format HH:mm #Arrival
  $timeInLunchFormatted = FormatTimeOut $timeInLunch #looks contradictory, but FormatTimeOut is correct
  $timeInLunchFormatted | Export-CliXml "C:\Scripts\TimAer\Variables\timeInLunch.xml"

  $timeSheet.Add("<td>" + $timeInLunchFormatted + "</td>")
}

elseif($result -eq 3) #Lunch ends
{
  $timeInLunch = Import-Clixml "C:\Scripts\TimAer\Variables\timeInLunch.xml"
  $timeOutLunch = Get-Date -Format HH:mm #Departure
  $timeOutLunchFormatted = FormatTimeIn $timeOutLunch #sounds contradictory, but FormatTimeIn is correct

  $timeSheet.Add("<td>" + $timeOutLunchFormatted + "</td>")

  [String]$totalTimeLunch = New-Timespan -Start $timeInLunch -End $timeOutLunchFormatted #Total time in the work day
  $totalTimeLunch = $totalTimeLunch.Substring(0, 5)
  $totalTimeLunch | Export-Clixml "C:\Scripts\TimAer\Variables\totalTimeLunch.xml"

  $timeSheet.Add("<td>" + $totalTimeLunch + "</td>")

  $totalTimeLunchMath = [int]$totalTimeLunch.Substring(3, 2) / 60
  $totalTimeLunchMath | Export-Clixml "C:\Scripts\TimAer\Variables\totalTimeLunchMath.xml"
}

else
{
  #Registers the time of arrival and saves the variable to the path specified for later use
  $dateIn = Get-Date #Gets today's date
  $timeIn = Get-Date -format HH:mm #Arrival
  $timeInFormatted = FormatTimeIn $timeIn
  $timeInFormatted | Export-CliXml "C:\Scripts\TimAer\Variables\timeIn.xml" #Saves the variable to the path specified
  
  #Date format for the caption of the table for that week
  $dayOfWeek = Get-Date -Format ddddd
  $month = Get-Date -Format MMMMM
  $day = Get-Date -Format dd
  $year = Get-Date -Format yyyy
  $dateNew = "$dayOfWeek, $month $day $year"
  
  #Naming for the html file containing the time/time sheet
  $fileName = Get-Date -Format MMM_d_ddd_yyyy
  $pathOut = "C:\Scripts\TimAer\Output\$fileName" + ".html"
  $fileName | Export-Clixml "C:\Scripts\TimAer\Variables\fileName.xml" #Saves the variable containing the date the week started to the path specified
   
  #Creating a new timeSheet table
  $timeSheet.Add("<br/><table align=`"center`" border=`"1`" style=`"border:#3366FF; border-collapse:collapse; width:1100px`">" +
                 "<caption style=`"font-family:arial; color:#000000; background-color:#3366FF; `"><b>Time Sheet. Week starting: $dateNew</b></caption>" +
                 "<th style=`"font-family:arial; color:#ffffff; background-color:#383838;`">Date</th>" +
                 "<th style=`"font-family:arial; color:#ffffff; background-color:#383838;`">Day</th>" +
                 "<th style=`"font-family:arial; color:#ffffff; background-color:#383838;`">Time In</th>" +
                 "<th style=`"font-family:arial; color:#ffffff; background-color:#383838;`">Lunch Started</th>" +
                 "<th style=`"font-family:arial; color:#ffffff; background-color:#383838;`">Lunch Ended</th>" +
                 "<th style=`"font-family:arial; color:#ffffff; background-color:#383838;`">Lunch Time</th>" +
                 "<th style=`"font-family:arial; color:#ffffff; background-color:#383838;`">Time Out</th>" +
                 "<th style=`"font-family:arial; color:#ffffff; background-color:#383838;`">Total Time - Lunch</th>" +
                 "<th style=`"font-family:arial; color:#ffffff; background-color:#383838;`">Decimal Time</th>" +
                 "<th style=`"font-family:arial; color:#ffffff; background-color:#383838;`">Pay Rate</th>" +
                 "<th style=`"font-family:arial; color:#ffffff; background-color:#383838;`">Total Pay</th>")
                 
  #Adding a row to the timeSheet table. This represents the opening of a work day               
  $timeSheet.Add("<tr style=`"font-family:arial; text-align:center; background-color:#FFFFFF; `">" +
                 "<td>" + $dateIn.toShortDateString() + "</td>" +
                 "<td>" + $dateIn.DayOfWeek + "</td>" +
                 "<td>" + $timeInFormatted + "</td>")
}

$fileName = Import-Clixml "C:\Scripts\TimAer\Variables\fileName.xml"
$pathOut = "C:\Scripts\TimAer\Output\$fileName" + ".html"
Add-Content $pathOut $timeSheet #Adds the contents of the timeSheet array to the html file
