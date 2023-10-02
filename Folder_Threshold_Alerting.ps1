# Script created by Anthony Mignona, 2023

# [README] - Enter the appropriate parameters for each "REQUIRED" section according to your needs. 
# All "REQUIRED" sections will be at the top of the script

# [REQUIRED] Your SMTP Parameters Parameters 
$smtp_server = 'smtp.EXAMPLE.com'
$email_from = 'Admin_test admin@EXAMPLE.com'
$email_footer = "Enter Email Footer Here."

# [REQUIRED] Your business hours accordingly. 
$business_hours_start = Get-Date -Hour 9 -Minute 0 -Second 0
$business_hours_end = Get-Date -Hour 17 -Minute 0 -Second 0

# [REQUIRED] Your folder_monitor.csv full path & where your log file will reside
$monitors = Import-Csv -Path "folder_monitors.csv"
$log_file = "folder_threshold_alerting.log"

# [END_OF_REQUIRED] - No need need to modify anything past this part of the script, unless you'd like to add/modify to the logic

# Determine if it's Weekday and Weekend logic. 
$today = Get-Date
if ($today.DayOfWeek -in @('Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday')) {
    # Set time threshold for the week days 
    $threshold = "weekday_minute_threshold"     
} else {
    # Set time threshold for the weekends
    $threshold = "weekend_minute_threshold"
}

# Determine if it's business or after-hours. Enter your business hours accordingly. 
if ($today -lt $business_hours_start -or $today -gt $business_hours_end) {
    $threshold ="ah_threshold"
}

# Iterate through each monitor (row) within folder_monitors.csv 
foreach($monitor in $monitors){
    write-host "Monitor : $($monitor)" -ForegroundColor DarkGreen 
    
    $results = @()
    # Skip if monitor is turned off
    if($monitor.on_bit -eq 0){
        write-host "[SKIP] - on_bit set to 0." -ForegroundColor white -BackgroundColor DarkYellow
        continue
    }

    # Skip if weekday and weekeday bit off
    if ($today.DayOfWeek -in @('Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday') -and $monitor.weekday_bit -eq 0) {
        write-host "[SKIP] - weekday_bit set to 0." -ForegroundColor white -BackgroundColor DarkYellow
        continue
    }
    
    # Skip if it's the weekend and weekend bit off 
    if ($today.DayOfWeek -in @('Saturday', 'Sunday') -and $monitor.weekend_bit -eq 0) {
        write-host "[SKIP] - weekend_bit set to 0." -ForegroundColor white -BackgroundColor DarkYellow
        continue
    }

    # Skip if it's After-Hours and After-Hours bit off 
    if($threshold -eq "ah_threshold" -and $monitor.ah_bit -eq 0){
        write-host "[SKIP] - ah_bit set to 0." -ForegroundColor white -BackgroundColor DarkYellow
        continue
    }
    
    write-host $monitor.$threshold
    # Find any documents within within criteria and threshold
    $alert_minutes_threshold = [int]::Parse($monitor.$threshold)
    if($monitor.Extension -eq '.*'){
        $results = Get-ChildItem -Path $monitor.path -File | Where-Object { $_.LastWriteTime -lt (Get-Date).AddMinutes($alert_minutes_threshold * -1)} 
    }
    else{
        $results = Get-ChildItem -Path $monitor.path -File | Where-Object  { $_.Extension -eq $monitor.Extension } | Where-Object { $_.LastWriteTime -lt (Get-Date).AddMinutes($alert_minutes_threshold * -1)} 
    }
    
    # Omit certain folders in the Where-Object pipe. 
    $table = $results | Where-Object { $_.PSIsContainer -eq $false } 
    $table_rows = [int](($table | Measure-Object -Line).Lines) 

    write-host $table -BackgroundColor Red -ForegroundColor White

    # Logic for sending an alert based on the number of table rows. 
    if($table_rows -gt 0){ # If there are any results, send the email, else don't send the email. 
        write-host ("There were " + $table_rows + " results detected. Sending Alert!") -BackgroundColor DarkRed -ForegroundColor White
        Add-Content -path $log_file -Value "($(Get-Date -Format "yyyy-MM-dd HH:mm:ss")) - [ALERT!] -  $($table_rows) file(s) within alerting threshold detected for job: $($monitor.friendly_name). SEND EMAIL ALERT!"
        # Add custom columns to the table, then convert table to HTML object
        $table = $results | Select-Object Name, Directory, LastWriteTime, 
            @{Name='Minutes';Expression={("{0:N0}" -f ((Get-Date) - $_.LastWriteTime).TotalMinutes)}} | ConvertTo-Html -Property Name, Directory, LastWriteTime, Minutes -Fragment -PreContent "<style>table {border-collapse: collapse; font-family: Arial, sans-serif;} th {background-color: navy; color: white; border: 1px solid black; padding: 10px;} td {border: 1px solid black; padding: 5px;}</style>"
    
        # Create the email
        $email_to = $monitor.Email_to
        $email_subject = $monitor.Email_Subj
        $email_body = "<html><body><h2>$($monitor.Email_body)</h2>$($table) <br> $($email_footer)</body></html>"
        $email_body = New-Object System.Net.Mail.MailMessage($email_from, $email_to, $email_subject, $email_body)
        $email_body.isBodyHtml = $true
      
        # Add CC'd recipients (must be comma-separated values within the cc_to column)
        if($monitor.cc_to){
            foreach ($cc in $monitor.cc_to) {
                $email_body.CC.Add($cc)
            }
        }

        # Send the email
        $smtp = New-Object Net.Mail.SmtpClient($smtp_server)
        $smtp.Send($email_body)
    }
    else{
        write-host ("There were " + $table_rows + " results detected. Do nothing!") -BackgroundColor DarkRed -ForegroundColor White
        Add-Content -path $log_file -Value "($(Get-Date -Format "yyyy-MM-dd HH:mm:ss")) - [OK] - $($table_rows) files detected within the alerting threshold for job: $($monitor.friendly_name). Do nothing."
    }

}
# Maintain the log file  
$content = get-content -Path $log_file -tail 1500; $content | Set-Content -Path $log_file -Force