#############################################################################
# Author  : Jorge E. Covarrubias
# Website : 
# LinkedIn  : https://www.linkedin.com/in/jorge-e-covarrubias-973217141/
#
# Version   : 3.1
# Created   : 9/14/2017
# Modified  :
# 8/19/2021   - Add Tab menu to select New User or CSV.
#			  - Add controls for the CSV tab.
# 8/18/2021   - Removing placing the password in the Company field.
#			  - Remove Homepath information when creating user account.
# 11/29/2017  - Adding heading for version tracking.
#				Removing commented out code.
# 
#
# Purpose : This script will create a GUI to input a students information so it can be created in AD.
#			Requires. First Name, Last Name, Graduation Year (2018), and Students five digit ID number.
#			Make sure to change Variables in Create-User Function for your environment.
#
#############################################################################

#This is done to remove the console window that appears after running the script. The info was found at
# https://stackoverflow.com/questions/1802127/how-to-run-a-powershell-script-without-displaying-a-window
$t = '[DllImport("user32.dll")] public static extern bool ShowWindow(int handle, int state);'
add-type -name win -member $t -namespace native
[native.win]::ShowWindow(([System.Diagnostics.Process]::GetCurrentProcess() | Get-Process).MainWindowHandle, 0)

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
Function New-User($userF,$userL,$gradYr,$pass){
		$lowerName = "$userF$userL".ToLower()
		$SAMAccountname = $lowerName
		$fullName = "$userF $userL"
		$email = "$lowerName"+"@<DOMAIN>.us"
		$OUName = "Class of $gradYr"
		$OUPath = "ou=$OUName,ou=Students,dc=<DOMAIN>,dc=local"
		$password = ConvertTo-SecureString -String "$pass" -AsPlainText -force

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
		$ADUserArguments = @{ Name = "$fullName";
			SamAccountName = "$lowerName";
			GivenName = "$userF";
			Surname = "$userL";
			DisplayName = "$fullName";
			UserPrincipalName = "$email";
			Path = "$OUPath";
			Enabled = $True ;
			AccountPassword = $password;
			ChangePasswordAtLogon = $False ;
			PasswordNeverExpires = $True ;
			CannotChangePassword = $True ;
			Description = "Student$gradYr";
			EmailAddress = "$email";
			Title = "Student";
			Department = "$gradYr";
			OtherAttributes = @{mailNickname = "$pass"}
		} 
		New-ADUser @ADUserArguments
		
		$result.Text += "An account for $fullname has been successfully created:`nU: $email`nP: $pass `n`n"
		Clear-Fields
	}
}


#xml form was created using Visual Studio and then edited here to my liking. This is the GUI
[xml]$Form = @"
    <Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
            Title="AD Student Account Creation" Height="405" Width="366">
        <Grid>
			<Grid>
				<TabControl HorizontalAlignment="Left" Height="195" Margin="10,5,0,0" VerticalAlignment="Top" Width="330">
					<TabItem Header="New User">
						<Grid Background="#FFE5E5E5">
							<Grid.ColumnDefinitions>
								<ColumnDefinition Width="53*"/>
								<ColumnDefinition Width="109*"/>
							</Grid.ColumnDefinitions>
							<Label HorizontalContentAlignment="Center" Content="First Name:" HorizontalAlignment="Left" Height="25" Margin="5,10,0,0" VerticalAlignment="Top" Width="100"/>
							<Label HorizontalContentAlignment="Center" Content="Last Name:" HorizontalAlignment="Left" Height="25" Margin="5,40,0,0" VerticalAlignment="Top" Width="100"/>
							<Label HorizontalContentAlignment="Center" Content="Graduating Year:" HorizontalAlignment="Left" Height="30" Margin="5,70,0,0" VerticalAlignment="Top" Width="100"/>
							<Label HorizontalContentAlignment="Center" Content="ID Number:" HorizontalAlignment="Left" Height="25" Margin="5,100,0,0" VerticalAlignment="Top" Width="100"/>
							<TextBox Name="FirstName" HorizontalAlignment="Left" Height="25" Margin="4,10,0,0" TextWrapping="Wrap" VerticalContentAlignment="Center" VerticalAlignment="Top" Width="204" Grid.Column="1"/>
							<TextBox Name="LastName" HorizontalAlignment="Left" Height="25" Margin="4,40,0,0" TextWrapping="Wrap" VerticalContentAlignment="Center" VerticalAlignment="Top" Width="204" Grid.Column="1"/>
							<TextBox Name="GraduatingYear" HorizontalAlignment="Left" Height="25" Margin="4,70,0,0" TextWrapping="Wrap" VerticalContentAlignment="Center" VerticalAlignment="Top" Width="204" Grid.Column="1"/>
							<TextBox Name="IDNumber" HorizontalAlignment="Left" Height="25" Margin="4,100,0,0" TextWrapping="Wrap" VerticalContentAlignment="Center" VerticalAlignment="Top" Width="204" Grid.Column="1"/>
							<Button Name="CreateAccount" Content="Create New Student Account" HorizontalAlignment="Left" Height="25" Margin="10,135,0,0" VerticalAlignment="Top" Width="304" Grid.ColumnSpan="2"/>
						</Grid>
					</TabItem>
					<TabItem Header="CSV Upload">
						<Grid Background="#FFE5E5E5">
							<Button Name="CSVStartButton" Content="Create CSV Accounts" HorizontalAlignment="Left" Height="25" Margin="10,135,0,0" VerticalAlignment="Top" Width="304"/>
							<TextBlock HorizontalAlignment="Left" Height="56" Margin="10,10,0,0" TextAlignment="Center" TextWrapping="Wrap" Text="Select CSV File. &#x0a;Header Row should include: &#x0a;FirstName, LastName, GraduatingYear, IDNumber" VerticalAlignment="Top" Width="304"/>
							<Button Name="BrowseButton" Content="Browse" HorizontalAlignment="Left" Height="25" Margin="10,88,0,0" VerticalAlignment="Top" Width="86"/>
							<TextBox HorizontalAlignment="Left" Height="25" Margin="101,88,0,0" TextWrapping="Wrap" VerticalContentAlignment="Center" VerticalAlignment="Top" Width="213"/>
						</Grid>
					</TabItem>
				</TabControl>
				<TextBox Name="Results" HorizontalAlignment="Left" Height="154" Margin="10,205,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="330" VerticalScrollBarVisibility="Visible"/>
			</Grid>
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
			New-User $userF $userL $gradYr $pass
		}
})
$Win.ShowDialog()