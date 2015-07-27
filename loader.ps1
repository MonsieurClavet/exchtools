Add-Type -AssemblyName presentationframework

[xml]$xaml = @"

<Window 

    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"

    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" x:Name="MainUI"

    Title="XAML :: Exchange Admin Tools" Height="349" Width="812" Background="#FF9E9E9E" Cursor="Hand" WindowStyle="ToolWindow" ResizeMode="NoResize" WindowStartupLocation="CenterScreen">

    <Grid Margin="0,0,-6,-11">

        <Grid.ColumnDefinitions>

            <ColumnDefinition Width="0*"/>

            <ColumnDefinition/>

        </Grid.ColumnDefinitions>

        <Grid.RowDefinitions>

            <RowDefinition Height="23*"/>

            <RowDefinition Height="316*"/>

        </Grid.RowDefinitions>

        <StatusBar HorizontalAlignment="Center" Height="30" Margin="-2,278,4,0" VerticalAlignment="Top" Width="810" BorderThickness="0" Grid.Row="1" Grid.ColumnSpan="2">

            <Label x:Name="Lbl_LoggedUser" Content="User logged in: " Width="619" FontStyle="Italic" FontWeight="SemiBold" Height="23"/>

            <Separator Visibility="Hidden"/>

            <Hyperlink x:Name="Link_feedbacks" NavigateUri="mailto:maxence.caron@bell.ca" ToolTip="Send me an email ! ">

                <Run Text="Questions, Comments, Feedbacks?"/>

            </Hyperlink>

            <Separator Visibility="Hidden"/>

        </StatusBar>

        <Menu x:Name="Menubar" Height="28" VerticalAlignment="Top" Grid.RowSpan="2" Grid.ColumnSpan="2" BorderThickness="0,0,0,5">

            <MenuItem Header="_Mailbox" Height="20" Margin="0" Width="60" FontSize="12" FontWeight="Bold" HorizontalAlignment="Center" HorizontalContentAlignment="Center">

                <MenuItem Header="Mailbox _Details"/>

                <MenuItem x:Name="Menu_MailboxPerm" Header="Mailbox _Permissions"/>

                <MenuItem x:Name="Menu_Quota" Header="Mailbox _Quota"/>

            </MenuItem>



            <Separator Visibility="Hidden"/>

            <MenuItem Header="MHC Intranet" Height="20" Margin="0" FontSize="12" FontWeight="Bold" HorizontalAlignment="Center" HorizontalContentAlignment="Center">

                <MenuItem x:Name="Menu_MHCIntraFID" Header="Shared Mailbox"/>

                <MenuItem Header="Distribution List"/>

                <MenuItem Header="SMTP Relay"/>

            </MenuItem>

            <Separator Visibility="Hidden"/>

            <MenuItem Header="Unified Messaging" Height="20" Margin="0" FontSize="12" FontWeight="Bold" HorizontalAlignment="Center" HorizontalContentAlignment="Center"/>

            <Separator Visibility="Hidden"/>

            <MenuItem x:Name="Menu_EAS" Header="ActiveSync" Height="20" Margin="0" FontSize="12" FontWeight="Bold" HorizontalAlignment="Center" HorizontalContentAlignment="Center"/>

            <Separator Visibility="Hidden"/>

            <MenuItem Header="Miscellaneous" Height="20" Margin="0" FontSize="12" FontWeight="Bold" HorizontalAlignment="Center" HorizontalContentAlignment="Center">

                <MenuItem x:Name="Menu_OOOF" Header="Out of Office"/>

            </MenuItem>

            <Separator Visibility="Hidden"/>

            <MenuItem x:Name="Menu_Help" Header="?" Height="20" Margin="0" FontSize="12" FontWeight="Bold" HorizontalAlignment="Center" HorizontalContentAlignment="Center"/>

            <Separator Margin="0" Width="62.334" Visibility="Hidden"/>

            <Button x:Name="Btn_SearchUser" Content="Search" Width="59" Height="17" FontSize="10"/>

            <TextBox x:Name="Txt_Username" Width="174" ToolTip="Specify User to search" IsHitTestVisible="True" AllowDrop="False"/>

        </Menu>

        <Grid x:Name="Grid_MailQuota" Grid.ColumnSpan="2" HorizontalAlignment="Left" Height="259" Margin="10,14,0,0" Grid.Row="1" VerticalAlignment="Top" Width="792" Visibility="Hidden">

            <TextBox x:Name="Txt_DefaultQuota" HorizontalAlignment="Left" Height="23" Margin="33,43,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="161" IsReadOnly="True" Background="#FFB6B6B6"/>

            <TextBox x:Name="Txt_CurrentMailboxSize" HorizontalAlignment="Left" Height="23" Margin="33,71,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="161" IsReadOnly="True" Background="#FFB6B6B6"/>

            <TextBox x:Name="Txt_TotalDeletedItems" HorizontalAlignment="Left" Height="23" Margin="33,99,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="161" IsReadOnly="True" Background="#FFB6B6B6"/>

            <TextBox x:Name="Txt_IssueWarning" HorizontalAlignment="Left" Height="23" Margin="33,127,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="161" IsReadOnly="True" Background="#FFB6B6B6"/>

            <TextBox x:Name="Txt_ProhibitSend" HorizontalAlignment="Left" Height="23" Margin="33,155,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="161" IsReadOnly="True" Background="#FFB6B6B6"/>

            <Label Content="Default Quota" HorizontalAlignment="Left" Margin="194,43,0,0" VerticalAlignment="Top" Width="149" RenderTransformOrigin="0.302,0.565"/>

            <Label Content="Current Mailbox Size" HorizontalAlignment="Left" Margin="194,71,0,0" VerticalAlignment="Top" Width="127"/>

            <Label Content="Total Deleted Item Size" HorizontalAlignment="Left" Margin="194,99,0,0" VerticalAlignment="Top" RenderTransformOrigin="1.056,0.478" Width="149"/>

            <Label Content="Issue Warning Quota" HorizontalAlignment="Left" Margin="194,127,0,0" VerticalAlignment="Top" Width="169"/>

            <TextBox x:Name="Txt_ProhibitSendReceive" HorizontalAlignment="Left" Height="23" Margin="33,183,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="161" IsReadOnly="True" Background="#FFB6B6B6"/>

            <Label Content="Prohibit Send Quota" HorizontalAlignment="Left" Margin="194,155,0,0" VerticalAlignment="Top" Width="113"/>

            <Label Content="Prohibit Send/Receive Quota" HorizontalAlignment="Left" Margin="194,183,0,0" VerticalAlignment="Top" Width="149"/>

            <Button x:Name="Btn_RefreshQuota" Content="Refresh" HorizontalAlignment="Left" Margin="93,220,0,0" VerticalAlignment="Top" Width="101"/>

            <TextBox x:Name="Txt_NewIssueWarn" HorizontalAlignment="Left" Height="23" Margin="454,73,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="107" Background="#FFB6B6B6" IsReadOnly="True"/>

            <TextBox x:Name="Txt_CalcNewProhibit" HorizontalAlignment="Left" Height="23" Margin="454,101,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="107" Background="White"/>

            <TextBox x:Name="Txt_NewProhibSendRecv" HorizontalAlignment="Left" Height="23" Margin="454,129,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="107" Background="#FFB6B6B6" IsReadOnly="True"/>

            <Label Content="Issue Warning Quota" HorizontalAlignment="Left" Margin="561,73,0,0" VerticalAlignment="Top" Width="152"/>

            <Label Content="Prohibit Send Quota" HorizontalAlignment="Left" Margin="561,101,0,0" VerticalAlignment="Top" Width="152"/>

            <Label Content="Prohibit Send/Receive Quota" HorizontalAlignment="Left" Margin="561,129,0,0" VerticalAlignment="Top" Width="152"/>

            <Button x:Name="Btn_Calculate" Content="Calculate" HorizontalAlignment="Left" Margin="486,163.993,0,0" VerticalAlignment="Top" Width="75"/>

            <Button x:Name="Btn_ApplyQuota" Content="Apply Quota" HorizontalAlignment="Left" Margin="566,163.993,0,0" VerticalAlignment="Top" Width="75"/>



        </Grid>

        <Grid x:Name="Grid_OOOF" Grid.ColumnSpan="2" HorizontalAlignment="Left" Height="268" Margin="-2,4,0,0" VerticalAlignment="Top" Width="808" Grid.Row="1" Visibility="Hidden">

            <Button Content="Get Status" HorizontalAlignment="Left" Margin="301,221.986,0,0" VerticalAlignment="Top" Width="75"/>

            <Button x:Name="Btn_ApplyOOOF" Content="Apply" HorizontalAlignment="Left" Margin="422,221.986,0,0" VerticalAlignment="Top" Width="75"/>

            <StackPanel HorizontalAlignment="Left" Height="68" Margin="301,133.993,0,0" VerticalAlignment="Top" Width="196">

                <RadioButton Content="Bell Canada Out of Office" FontSize="12"/>

                <RadioButton Content="Bell Mobility Out of Office" FontSize="12"/>

                <RadioButton Content="Custom Out of Office" FontSize="12"/>

                <RadioButton Content="Disable" FontSize="12"/>

            </StackPanel>

            <StackPanel HorizontalAlignment="Left" Height="59" Margin="301,52.993,0,0" VerticalAlignment="Top" Width="57">

                <Label Content="User" HorizontalAlignment="Right"/>

                <Label Content="Status" HorizontalAlignment="Right"/>

            </StackPanel>

            <StackPanel HorizontalAlignment="Left" Height="59" Margin="369,52.993,0,0" VerticalAlignment="Top" Width="157">

                <TextBox Height="23" TextWrapping="Wrap"/>

                <TextBox Height="23" TextWrapping="Wrap" IsReadOnly="True"/>

            </StackPanel>

        </Grid>

        <Grid x:Name="Grid_MHCIntraFID" HorizontalAlignment="Left" Height="263" VerticalAlignment="Top" Width="808"  Visibility="Hidden" Grid.ColumnSpan="2" Margin="4,10,0,0" Grid.Row="1" IsEnabled="False">

            <StackPanel HorizontalAlignment="Left" Height="61" Margin="140,31,0,0" VerticalAlignment="Top" Width="117">

                <RadioButton Content="Bell Canada" FontSize="12"/>

                <RadioButton Content="Bell Mobility" FontSize="12"/>

                <RadioButton Content="Bell Express Vu" FontSize="12"/>

                <RadioButton Content="Bell Distribution" FontSize="12"/>

            </StackPanel>

            <StackPanel HorizontalAlignment="Left" Height="95" Margin="140,104,0,0" VerticalAlignment="Top" Width="206">

                <TextBox Height="23" TextWrapping="Wrap"/>

                <TextBox Height="23" TextWrapping="Wrap"/>

                <TextBox Height="23" TextWrapping="Wrap"/>

                <TextBox Height="23" TextWrapping="Wrap"/>

            </StackPanel>

            <StackPanel HorizontalAlignment="Left" Height="100" Margin="39,103,0,0" VerticalAlignment="Top" Width="100">

                <Label Content="Display Name" HorizontalAlignment="Right"/>

                <Label Content="Email Address" HorizontalAlignment="Right"/>

                <Label Content="Owner" HorizontalAlignment="Right"/>

                <Label Content="Backup Owner" HorizontalAlignment="Right"/>

            </StackPanel>

            <StackPanel HorizontalAlignment="Left" Height="100" Margin="454,31,0,0" VerticalAlignment="Top" Width="303">

                <RadioButton Content="Send On Behalf and Full Mailbox Permissions" FontSize="12"/>

                <RadioButton Content="Send As and Full Mailbox Permissions" FontSize="12"/>

                <RadioButton Content="Send As Only" FontSize="12"/>

                <RadioButton Content="Send On Behalf Only" FontSize="12"/>

                <RadioButton Content="Full Mailbox Permissions - No Sending Access" FontSize="12"/>

            </StackPanel>

            <Button Content="Open CSV" HorizontalAlignment="Left" Margin="454,136,0,0" VerticalAlignment="Top" Width="111"/>

            <Button Content="Grant Permissions" HorizontalAlignment="Left" Margin="607,136,0,0" VerticalAlignment="Top" Width="111"/>

            <Button Content="Create FID" HorizontalAlignment="Left" Margin="140,216,0,0" VerticalAlignment="Top" Width="111" RenderTransformOrigin="-1.216,0.05"/>

        </Grid>

        <Grid x:Name="Grid_EAS" Grid.ColumnSpan="2" HorizontalAlignment="Left" Height="274" Margin="4,4,0,0" Grid.Row="1" VerticalAlignment="Top" Width="802" Visibility="Hidden" >

            <Label Content="User(s): " HorizontalAlignment="Left" Margin="125,53,0,0" VerticalAlignment="Top"/>

            <TextBox HorizontalAlignment="Left" Height="23" Margin="195,53,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="120"/>

            <TextBox HorizontalAlignment="Left" Height="167" Margin="395.378,97,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="397" IsUndoEnabled="True" IsReadOnly="True" IsEnabled="False" RenderTransformOrigin="0.5,0.5" AcceptsReturn="True" VerticalScrollBarVisibility="Visible">

                <TextBox.RenderTransform>

                    <TransformGroup>

                        <ScaleTransform/>

                        <SkewTransform AngleX="-1.273"/>

                        <RotateTransform/>

                        <TranslateTransform X="1.478"/>

                    </TransformGroup>

                </TextBox.RenderTransform>

            </TextBox>

            <StackPanel HorizontalAlignment="Left" Height="69" Margin="125,119,0,0" VerticalAlignment="Top" Width="187">

                <RadioButton Content="Enable Single User" FontSize="12"/>

                <RadioButton Content="Enable Bulk Users" FontSize="12"/>

                <RadioButton Content="Disable User" FontSize="12"/>

                <RadioButton Content="Check Status" FontSize="12"/>

            </StackPanel>

            <Button Content="Get-DeviceID" HorizontalAlignment="Left" Margin="398,18.993,0,0" VerticalAlignment="Top" Width="128" Height="19"/>

        </Grid>





    </Grid>



</Window>

"@



#region Generate and Create Variables

$reader=(New-Object System.Xml.XmlNodeReader $xaml)

$Window=[Windows.Markup.XamlReader]::Load( $reader )

$xaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | ForEach {

    Set-Variable -Name ($_.Name) -Value $Window.FindName($_.Name)

}

#endregion



#region Generic Variables

$whoami = $env:username

$Lbl_LoggedUser.Content = "User logged in: " + $whoami

#endregion



#region Functions

Function ClearView{

	$Grid_MailQuota.Visibility = "Hidden"

	$Grid_MHCIntraFID.Visibility = "Hidden"

	$Grid_OOOF.Visibility = "Hidden"

	$Grid_EAS.Visibility = "Hidden"

}

#endregion





#region Menu Button

$Menu_MailboxPerm.Add_Click({

	ClearView

	$Window.Title = "XAML :: Exchange Admin Tools - Mailbox Permission"

})



$Menu_Quota.Add_Click({

	Clearview

	$Grid_MailQuota.Visibility = "Visible"

	$Window.Title = "XAML :: Exchange Admin Tools - Mailbox Quota"

})



$Menu_MHCIntraFID.Add_Click({

	ClearView

	$Grid_MHCIntraFID.Visibility = "Visible"

	$Window.Title = "XAML :: Exchange Admin Tools - MHC Intranet FID Creation"

})



$Menu_OOOF.Add_Click({

	ClearView

	$Grid_OOOF.Visibility = "Visible"

	$Window.Title = "XAML :: Exchange Admin Tools - Out Of Office"

})



$Menu_EAS.Add_Click({

	ClearView

	$Grid_EAS.Visibility = "Visible"

	$Window.Title = "XAML :: Exchange Admin Tools - Active Sync"

})

#endregion





$Window.Showdialog() | Out-Null
