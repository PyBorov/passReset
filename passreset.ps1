Add-Type -assembly System.Windows.Forms

Function SearchUser {
    Try {
        $users_grid.Rows.Clear()
        if ($find_str.Text.Length -ne 0){
        $find_qry_str = 'mail -like "*' + $find_str.Text + '*"' + " -OR " + 'SamAccountName -like "*' + $find_str.Text + '*" -OR ' +'Name -like "*' + $find_str.Text + '*"'
        Get-ADUser -Filter $find_qry_str -Properties Name,SamAccountName, EmailAddress,PasswordLastSet | foreach { 
	$users_grid.Rows.Add($_.Name,$_.SamAccountName,$_.EmailAddress,(($_.PasswordLastSet).AddMonths(2)).ToString('dd.MM.yy HH:mm'))
	}
        } else { $Output = $wshell.Popup("Password not entered")}
        } 
        Catch
        {

        }
        
	}


Function ResetPass {
    $user_res = $users_grid[1, $users_grid.CurrentCell.RowIndex].Value
    Try 
    {
        if ($chg_pass.Checked){
        Set-ADAccountPassword -Identity $user_res -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $pass_str.Text -Force -Verbose) | Set-ADuser -ChangePasswordAtLogon $True
        $user_res = $users_grid[0, $users_grid.CurrentCell.RowIndex].Value
        $Output = $wshell.Popup($user_res+"'s password has been changed `nUser must change password at next logon")
        } else {
        Set-ADAccountPassword -Identity $user_res -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $pass_str.Text -Force -Verbose) 
        $user_res = $users_grid[0, $users_grid.CurrentCell.RowIndex].Value
        $Output = $wshell.Popup($user_res+"'s password has been changed ")
        }
        $find_str.Text = ''
    } 
    Catch
    {

        $Output = $wshell.Popup("Password does not meet policy requirements")
    }
    
}

$main_form = New-Object System.Windows.Forms.Form
$main_form.Text ='Change password'
$main_form.MaximumSize = New-Object System.Drawing.Size(639,522)
$main_form.MinimumSize = New-Object System.Drawing.Size(639,522)
$main_form.StartPosition = "CenterScreen"
$main_form.MaximizeBox = $false
$main_form.KeyPreview = $true
$main_form.Add_KeyDown({
    if($_.KeyCode -eq "Enter") {
       SearchUser
        } 
    if($_.KeyCode -eq "Escape") {
       $main_form.Close()
        } 
    })

$label_login = New-Object System.Windows.Forms.Label
$label_login.Location  = New-Object System.Drawing.Point(0,8)
$label_login.Text = 'User:'
$label_login.Size = '40,12'
$main_form.Controls.Add($label_login)

$find_str = New-Object System.Windows.Forms.TextBox
$find_str.Location  = New-Object System.Drawing.Point(40,3)
$find_str.Text = ''
$find_str.Size = '250,30'
$main_form.Controls.Add($find_str)

$find_btn = New-Object System.Windows.Forms.Button
$find_btn.Text = 'Find'
$find_btn.Location = New-Object System.Drawing.Point(520,3)
$find_btn.Size = '100,25'
$find_btn.Add_Click({SearchUser})
$main_form.Controls.Add($find_btn)

$label_pass = New-Object System.Windows.Forms.Label
$label_pass.Location  = New-Object System.Drawing.Point(0,43)
$label_pass.AutoSize = $true
$label_pass.Text = 'Password:'
$main_form.Controls.Add($label_pass)

$pass_str = New-Object System.Windows.Forms.TextBox
$pass_str.Location  = New-Object System.Drawing.Point(60,40)
$pass_str.Text = ''
$pass_str.Size = '210,30'
$main_form.Controls.Add($pass_str)

$set_pass_btn = New-Object System.Windows.Forms.Button
$set_pass_btn.Text = 'Change'
$set_pass_btn.Size = '75, 20'
$set_pass_btn.Location = New-Object System.Drawing.Point(520,40)
$set_pass_btn.Size = '100,25'
$set_pass_btn.Add_Click({ResetPass})
$main_form.Controls.Add($set_pass_btn)

$chg_pass = New-Object System.Windows.Forms.CheckBox
$chg_pass.Text = 'Change at next logon'
$chg_pass.Size = '75, 20'
$chg_pass.Location = New-Object System.Drawing.Point(280,40)
$chg_pass.Size = '150,25'
$chg_pass.Checked = $true
$main_form.Controls.Add($chg_pass)


$users_grid = New-Object System.Windows.Forms.DataGridView
$users_grid.Location = New-Object System.Drawing.Point(0,70)
$users_grid.Size = '620,410'
$users_grid.ColumnCount = 4
$users_grid.ColumnHeadersVisible = $true
$users_grid.Columns[0].Name = "Name"
$users_grid.Columns[1].Name = "Login"
$users_grid.Columns[2].Name = "Mail"
$users_grid.Columns[3].Name = "Password expires"
$users_grid.ReadOnly = $true
$users_grid.MultiSelect = $false
$users_grid.SelectionMode = 1 
$main_form.Controls.Add($users_grid)
$wshell = New-Object -ComObject Wscript.Shell

$main_form.ShowDialog() | Out-Null
