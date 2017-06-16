#This line adds the .Net Framework WPF to Powershell that makes the GUI work
Add-Type -AssemblyName PresentationFramework

#Clear field Function
Function Clear-Fields(){
		$userFirst.text=""
		$userLast.text=""
		$gradYear.text=""
		$id.text=""
}

#Function being called from the button add_Click with its variables 
Function Create-User($userF,$userL,$gradYr,$pass){
		$lowerName = "$userF$userL".ToLower()
		$SAMAccountname = $lowerName
		$fullName = "$userF $userL"
		$email = "$lowerName"+"@<DOMAIN>.us"
		$script = "<BATFILE>.bat"
		$OUName = "Class of $gradYr"
		$OUPath = "ou=$OUName,ou=Students,dc=<DOMAIN>,dc=local"
		$homePath = "\\<SERVER>\Students\$gradYr\$lowerName"
		$pwd = ConvertTo-SecureString -String "$pass" -AsPlainText -force

#Checking to see if account and OU exist
If(@(Get-ADObject -Filter { SAMAccountname -eq $SAMAccountname }).Count -ge 1){
		#[System.Windows.Messagebox]::Show("An account with the name `"$SAMAccountname`" already exist. Starting over...", "ERROR: Account exists")
		$result.Text += "---ERROR---`nAn account with the name `"$SAMAccountname`" already exist. `n`n"
		Clear-Fields
	}ElseIf(@(Get-ADOrganizationalUnit -Filter "Name -like '$OUName'").Count -eq 0){
		#[System.Windows.Messagebox]::Show("The OU doesn't exit. Make sure the OU exist or the graduating year is correct.", "ERROR: OU doesn't exists")
		$result.Text += "---ERROR---`nThe OU doesn't exit. Make sure the OU exist or the graduating year is correct for `"$SAMAccountname`". `n`n"
		$gradYear.text=""
	}Else{
#Running command
		New-ADUser -Name "$fullName" -SamAccountName "$lowerName" `
		-GivenName "$userF" -Surname "$userL" -DisplayName "$fullName" -UserPrincipalName "$email" `
		-Path "$OUPath" -Enabled $true -AccountPassword $pwd `
		-ChangePasswordAtLogon $false -PasswordNeverExpires $true -CannotChangePassword $true `
		-Description "Student$gradYr" -EmailAddress "$email" -Title "Student" -Department "$gradYr" -Company "$pass" `
		-ScriptPath "$gradYr$Script" -HomeDrive "H:" -HomeDirectory "$homePath" `
		-OtherAttributes @{mailNickname="$pass"}

		$result.Text += "An account for $fullname has been successfully created:`nU: $email`nP: $pass `n`n"
		Clear-Fields
	}
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
			[System.Windows.Messagebox]::Show("Please make sure all values are entered correctly", "ERROR: Missing Values")
		}ElseIf( -not $gradok){
			[System.Windows.Messagebox]::Show("Graduation year must be a numeric value (no letters)", "ERROR: Graduation Numeric Value")
		}ElseIf($gradYr.length -notmatch 4){
			[System.Windows.Messagebox]::Show("Graduation year must be 4 digits long", "ERROR: Graduation: Four Digit Value")
		}ElseIf( -not $passok){
			[System.Windows.Messagebox]::Show("ID Number must be a numeric value (no letters)", "ERROR: ID Numeric Value")
		}ElseIf($pass.length -notmatch 5){
			[System.Windows.Messagebox]::Show("ID number must be 5 digits long", "ERROR: ID: Five Digit Value")
		}Else{
			Create-User $userF $userL $gradYr $pass
		}
})
$Win.ShowDialog()