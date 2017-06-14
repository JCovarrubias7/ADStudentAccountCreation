# This line adds the .Net Framework WPF to Powershell that makes the GUI work
Add-Type -AssemblyName PresentationFramework

#Function being called from the button add_Click with its variables 
Function Create-User($userF,$userL,$gradYr,$pass){
$lowerName = "$userF$userL".ToLower()
$fullName = "$userF $userL"
$email = "$lowerName"+"@<DOMAIN>.us"
$script = "<BATFILE>.bat"
$pwd = ConvertTo-SecureString -String "$pass" -AsPlainText -force

#Checking to see if account exist
$SAMAccountname = $lowerName
If(@(Get-ADObject -Filter { SAMAccountname -eq $SAMAccountname }).Count -ge 1){
[System.Windows.Messagebox]::Show("The object/user $SAMAccountname already exsists. Exiting now...", "Account exists")
$userFirst.text=""
$userLast.text=""
$gradYear.text=""
$id.text=""
$result.Text += "An account with the name `"$SAMAccountname`" already exist. `n`n"
return;
}else{
#Running command
New-ADUser -Name "$fullName" -SamAccountName "$lowerName" `
 -GivenName "$userF" -Surname "$userL" -DisplayName "$fullName" -UserPrincipalName "$email" `
 -Path "ou=Class of $gradYr,ou=Students,dc=<DOMAIN>,dc=local" `
 -Enabled $true -AccountPassword $pwd `
 -ChangePasswordAtLogon $false -PasswordNeverExpires $true -CannotChangePassword $true `
 -Description "Student$gradYr" -EmailAddress "$email" -Title "Student" -Department "$gradYr" -Company "$pass" `
 -ScriptPath "$gradYr$Script" -HomeDrive "H:" -HomeDirectory "\\<SERVER>\Students\$gradYr\$lowerName" `
 -OtherAttributes @{mailNickname="$pass"}
}
$result.Text += "$fullname account has been created:`nU: $email`nP: $pass `n`n"
}


#xml form was created using Visual Studio and then edited here to my liking. This is the GUI
[xml]$Form = @"
    <Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
            Title="AD Student Account Creation" Height="355" Width="355">
        <Grid>
            <Label HorizontalContentAlignment="Center" Content="First Name:" HorizontalAlignment="Left" Height="23" Margin="10,10,0,0" VerticalAlignment="Top" Width="100"/>
            <Label HorizontalContentAlignment="Center" Content="Last Name:" HorizontalAlignment="Left" Height="23" Margin="10,40,0,0" VerticalAlignment="Top" Width="100"/>
            <Label HorizontalContentAlignment="Center" Content="Graduating Year:" HorizontalAlignment="Left" Height="30" Margin="10,70,0,0" VerticalAlignment="Top" Width="100"/>
            <Label HorizontalContentAlignment="Center" Content="ID Number:" HorizontalAlignment="Left" Height="23" Margin="10,100,0,0" VerticalAlignment="Top" Width="100"/>
            <TextBox Name="FirstName" HorizontalAlignment="Left" Height="23" Margin="115,10,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="217"/>
            <TextBox Name="LastName" HorizontalAlignment="Left" Height="23" Margin="115,40,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="217"/>
            <TextBox Name="GraduatingYear" HorizontalAlignment="Left" Height="23" Margin="115,70,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="217"/>
            <TextBox Name="IDNumber" HorizontalAlignment="Left" Height="23" Margin="115,100,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="217"/>
            <TextBox Name="Results" HorizontalAlignment="Left" Height="142" Margin="10,130,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="322"/>
            <Button Name="CreateAccount" Content="Create Account" HorizontalAlignment="Left" Height="25" Margin="10,284,0,0" VerticalAlignment="Top" Width="322"/>
        </Grid>
    </Window>
"@
#The NodeReader will grab the form, read it, and the Win will load the form from NR
$NR=(New-Object System.Xml.XmlNodeReader $Form)
$Win=[Windows.Markup.XamlReader]::Load( $NR )

#Variables that will take the input form the GUI so they can be called later
$userFirst = $Win.FindName("FirstName")
$userLast = $Win.FindName("LastName")
$gradYear = $Win.FindName("GraduatingYear")
$id = $Win.FindName("IDNumber")
$result = $Win.FindName("Results")
$create = $Win.FindName("CreateAccount")

#On click, we will take our variables from the GUI, make them into text and run them by the function Create-User
$create.Add_Click({
$userF = $userFirst.Text
$userL = $userLast.Text
$gradYr = $gradYear.Text
$pass = $id.Text
$gradYrValue = $gradYr -as [Double]
$gradok = $gradYrValue -ne $NULL
$passValue = $pass -as [Double]
$passok = $passValue -ne $NULL
If($userF -eq "" -or $userL -eq "" -or $gradYr -eq "" -or $pass -eq ""){
[System.Windows.Messagebox]::Show("Please make sure all values are entered correctly", "Missing Values")
}ElseIf($gradYr -notmatch "[2018,2019,2020,2021,2022,2023,2024,2025]"){
[System.Windows.Messagebox]::Show("Please enter Graduating year", "Wrong Graduating Year")
}ElseIf( -not $gradok){
[System.Windows.Messagebox]::Show("Please enter a numeric value", "Numeric Value")
}ElseIf($gradYr.length -notmatch 4){
[System.Windows.Messagebox]::Show("ID must be 4 digits long", "Graduation: Four Digit Value")
}ElseIf( -not $passok){
[System.Windows.Messagebox]::Show("Please enter a numeric value", "Numeric Value")
}ElseIf($pass.length -notmatch 5){
[System.Windows.Messagebox]::Show("ID must be 5 digits long", "ID: Five Digit Value")
}else{
Create-User $userF $userL $gradYr $pass
$userFirst.text=""
$userLast.text=""
$gradYear.text=""
$id.text=""
}
})
$Win.ShowDialog()