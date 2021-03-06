Add-Type -AssemblyName presentationframework
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | out-null
Import-Module ActiveDirectory

[xml]$xaml = @"

<Window 

    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"

    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" x:Name="MainUI"

    Title="XAML :: Exchange Admin Tools" Height="349" Width="812" Background="#FF9E9E9E" Cursor="Hand" WindowStyle="ToolWindow" ResizeMode="NoResize" WindowStartupLocation="CenterScreen">

    <Grid x:Name="Grid_Mail" Margin="0,0,-6,-11">

        <Grid.ColumnDefinitions>

            <ColumnDefinition Width="0*"/>

            <ColumnDefinition/>

        </Grid.ColumnDefinitions>

        <Grid.RowDefinitions>

            <RowDefinition Height="23*"/>

            <RowDefinition Height="316*"/>

        </Grid.RowDefinitions>

        
        <Menu x:Name="Menubar" Height="37" VerticalAlignment="Top" Grid.RowSpan="2" Grid.ColumnSpan="2" BorderThickness="0,0,0,5">

            <MenuItem Header="_Mailbox" Height="20" Margin="0" Width="60" FontSize="12" FontWeight="Bold" HorizontalAlignment="Center" HorizontalContentAlignment="Center">

                <MenuItem x:Name="Menu_MailboxInfo" Header="Mailbox _Info"/>

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

            <Button x:Name="Btn_SearchUser" Content="Search" Width="59" Height="25" FontSize="10"/>

            <TextBox x:Name="Txt_Username" Width="174" ToolTip="Specify User to search" IsHitTestVisible="True" AllowDrop="False" FontSize="9.333"/>

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

            <Label Content="Prohibit Send Quota" HorizontalAlignment="Left" Margin="194,155,0,0" VerticalAlignment="Top" Width="160"/>

            <Label Content="Prohibit Send/Receive Quota" HorizontalAlignment="Left" Margin="194,183,0,0" VerticalAlignment="Top" Width="169"/>

            <Button x:Name="Btn_RefreshQuota" Content="Refresh" HorizontalAlignment="Left" Margin="93,220,0,0" VerticalAlignment="Top" Width="101"/>

            <TextBox x:Name="Txt_NewIssueWarn" HorizontalAlignment="Left" Height="23" Margin="454,73,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="107" Background="#FFB6B6B6" IsReadOnly="True"/>

            <TextBox x:Name="Txt_CalcNewProhibit" HorizontalAlignment="Left" Height="23" Margin="454,101,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="107" Background="White"/>

            <TextBox x:Name="Txt_NewProhibSendRecv" HorizontalAlignment="Left" Height="23" Margin="454,129,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="107" Background="#FFB6B6B6" IsReadOnly="True"/>

            <Label Content="Issue Warning Quota" HorizontalAlignment="Left" Margin="561,73,0,0" VerticalAlignment="Top" Width="152"/>

            <Label Content="Prohibit Send Quota" HorizontalAlignment="Left" Margin="561,101,0,0" VerticalAlignment="Top" Width="152"/>

            <Label Content="Prohibit Send/Receive Quota" HorizontalAlignment="Left" Margin="561,129,0,0" VerticalAlignment="Top" Width="181"/>

            <Button x:Name="Btn_Calculate" Content="Calculate" HorizontalAlignment="Left" Margin="486,163.993,0,0" VerticalAlignment="Top" Width="75"/>

            <Button x:Name="Btn_ApplyQuota" Content="Apply Quota" HorizontalAlignment="Left" Margin="566,163.993,0,0" VerticalAlignment="Top" Width="75"/>
            <Grid HorizontalAlignment="Left" Height="51" Margin="248,43,0,0" VerticalAlignment="Top" Width="73"/>



        </Grid>

        <Grid x:Name="Grid_OOOF" Grid.ColumnSpan="2" HorizontalAlignment="Left" Height="268" Margin="-2,4,0,0" VerticalAlignment="Top" Width="808" Grid.Row="1" Visibility="Hidden">

            <Button Content="Get Status" HorizontalAlignment="Left" Margin="301,221.986,0,0" VerticalAlignment="Top" Width="75"/>

            <Button x:Name="Btn_ApplyOOOF" Content="Apply" HorizontalAlignment="Left" Margin="422,221.986,0,0" VerticalAlignment="Top" Width="75"/>

            <StackPanel HorizontalAlignment="Left" Height="68" Margin="301,133.993,0,0" VerticalAlignment="Top" Width="196">

                <RadioButton x:Name="RD_OOOFBellCan" Content="Bell Canada Out of Office" FontSize="12"/>

                <RadioButton x:Name="RD_OOOFBellMob" Content="Bell Mobility Out of Office" FontSize="12"/>

                <RadioButton x:Name="RD_OOOFBellCustom" Content="Custom Out of Office" FontSize="12" IsEnabled="False"/>

                <RadioButton x:Name="RD_OOOFDisabled" Content="Disable" FontSize="12"/>

            </StackPanel>

            <StackPanel HorizontalAlignment="Left" Height="59" Margin="301,52.993,0,0" VerticalAlignment="Top" Width="57">

                <Label Content="User" HorizontalAlignment="Right"/>

                <Label Content="Status" HorizontalAlignment="Right"/>

            </StackPanel>

            <StackPanel HorizontalAlignment="Left" Height="59" Margin="369,52.993,0,0" VerticalAlignment="Top" Width="157">

                <TextBox x:Name="Txt_OOOFUser" Height="23" TextWrapping="Wrap"/>

                <TextBox Height="23" TextWrapping="Wrap" IsReadOnly="True"/>

            </StackPanel>

        </Grid>

        <Grid x:Name="Grid_MHCIntraFID" HorizontalAlignment="Left" Height="263" VerticalAlignment="Top" Width="808" Grid.ColumnSpan="2" Margin="4,10,0,0" Grid.Row="1" Visibility="Hidden">

            <StackPanel HorizontalAlignment="Left" Height="61" Margin="140,31,0,0" VerticalAlignment="Top" Width="117">

                <RadioButton Content="Bell Canada" FontSize="12"/>

                <RadioButton Content="Bell Mobility" FontSize="12"/>

                <RadioButton Content="Bell Express Vu" FontSize="12"/>

                <RadioButton Content="Bell Distribution" FontSize="12"/>

            </StackPanel>

            <StackPanel HorizontalAlignment="Left" Height="128" Margin="140,104,0,0" VerticalAlignment="Top" Width="206">

                <TextBox x:Name="Txt_MHCFID_Alias" Height="23" TextWrapping="Wrap"/>

                <TextBox x:Name="Txt_MHCFID_DN" Height="23" TextWrapping="Wrap"/>

                <TextBox x:Name="Txt_MHCFID_Email" Height="23" TextWrapping="Wrap"/>

                <TextBox x:Name="Txt_MHCFID_Owner" Height="23" TextWrapping="Wrap"/>
                <TextBox x:Name="Txt_MHCFID_Backup" Height="23" TextWrapping="Wrap"/>

            </StackPanel>

            <StackPanel HorizontalAlignment="Left" Height="129" Margin="39,103,0,0" VerticalAlignment="Top" Width="100">

                <Label Content="Alias" HorizontalAlignment="Right" FontSize="10.667"/>

                <Label Content="Display Name" HorizontalAlignment="Right" FontSize="10.667"/>

                <Label Content="Email Address" HorizontalAlignment="Right" FontSize="10.667"/>

                <Label Content="Owner" HorizontalAlignment="Right" FontSize="10.667"/>

                <Label Content="Backup Owner" HorizontalAlignment="Right" FontSize="10.667"/>

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

            <Button x:Name="Btn_MHCFID_Create" Content="Create FID" HorizontalAlignment="Left" Margin="140,237,0,0" VerticalAlignment="Top" Width="111" RenderTransformOrigin="-1.216,0.05"/>

        </Grid>

        <Grid x:Name="Grid_EAS" Grid.ColumnSpan="2" HorizontalAlignment="Left" Height="274" Margin="4,4,0,0" Grid.Row="1" VerticalAlignment="Top" Width="802" Visibility="Hidden" >

            <Label Content="User(s): " HorizontalAlignment="Left" Margin="125,53,0,0" VerticalAlignment="Top"/>

            <TextBox x:Name ="Txt_EASUser" HorizontalAlignment="Left" Height="23" Margin="195,53,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="120"/>

            <TextBox x:Name="Txt_EASOutput" HorizontalAlignment="Left" Height="167" Margin="395.378,97,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="397" IsUndoEnabled="True" IsReadOnly="True" IsEnabled="False" RenderTransformOrigin="0.5,0.5" AcceptsReturn="True" VerticalScrollBarVisibility="Visible">

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

                <RadioButton x:Name="Rd_SingleEAS" Content="Enable Single User" FontSize="12"/>

                <RadioButton x:Name="Rd_BulkEAS" Content="Enable Bulk Users" FontSize="12" IsEnabled="False"/>

                <RadioButton x:Name="Rd_DisableEAS" Content="Disable User" FontSize="12"/>

                <RadioButton x:Name="Rd_CheckEAS" Content="Check Status" FontSize="12"/>

            </StackPanel>

            <Button Content="Get-DeviceID" HorizontalAlignment="Left" Margin="398,18.993,0,0" VerticalAlignment="Top" Width="128" Height="19"/>
            <Button x:Name="Btn_EnableEAS" Content="Go !" HorizontalAlignment="Left" Margin="125,188,0,0" VerticalAlignment="Top" Width="75"/>

        </Grid>
        <Grid x:Name="Grid_MailboxInfo" Grid.ColumnSpan="2" HorizontalAlignment="Left" Height="273" Margin="0,10,0,0" Grid.Row="1" VerticalAlignment="Top" Width="806">
            <StackPanel HorizontalAlignment="Left" Height="209" Margin="223,54,0,0" VerticalAlignment="Top" Width="140">
                <Label x:Name="label" Content="Display Name"/>
                <Label x:Name="label_Copy" Content="SamAccountName"/>
                <Label x:Name="label_Copy1" Content="Email Address"/>
                <Label x:Name="label_Copy2" Content="Server Exchange Name"/>
                <Label x:Name="label_Copy3" Content="Database Name"/>
                <Label x:Name="label_Copy4" Content="is ActiveSync Enabled"/>
                <Label x:Name="label_Copy5" Content="is UM Enabled"/>
            </StackPanel>
            <StackPanel HorizontalAlignment="Left" Height="209" Margin="10,54,0,0" VerticalAlignment="Top" Width="208">
                <TextBox x:Name="Txt_MbxInfo_DN" Height="27" TextWrapping="Wrap" IsReadOnly="True"/>
                <TextBox x:Name="Txt_MbxInfo_SAM" Height="27" TextWrapping="Wrap" IsReadOnly="True"/>
                <TextBox x:Name="Txt_MbxInfo_Email" Height="27" TextWrapping="Wrap" IsReadOnly="True"/>
                <TextBox x:Name="Txt_MbxInfo_SrvExch" Height="27" TextWrapping="Wrap" IsReadOnly="True"/>
                <TextBox x:Name="Txt_MbxInfo_DB" Height="27" TextWrapping="Wrap" IsReadOnly="True"/>
                <TextBox x:Name="Txt_MbxInfo_iSEAS" Height="27" TextWrapping="Wrap" IsReadOnly="True"/>
                <TextBox x:Name="Txt_MbxInfo_iSUM" Height="27" TextWrapping="Wrap" IsReadOnly="True"/>
            </StackPanel>
        </Grid>
        <Grid x:Name="Grid_Help" Grid.ColumnSpan="2" HorizontalAlignment="Left" Height="264" Margin="10,14,0,0" Grid.Row="1" VerticalAlignment="Top" Width="792" Visibility="Hidden">
            <StackPanel HorizontalAlignment="Left" Height="100" Margin="37,26,0,0" VerticalAlignment="Top" Width="100">
                <Label x:Name="label1" Content="Label" HorizontalAlignment="Left" VerticalAlignment="Top" Width="100"/>
            </StackPanel>
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


#endregion



#region Functions

Function ClearView{

    $Grid_MailQuota.Visibility = "Hidden"

    $Grid_MHCIntraFID.Visibility = "Hidden"

    $Grid_OOOF.Visibility = "Hidden"

    $Grid_EAS.Visibility = "Hidden"

    $Grid_MailboxInfo.Visibility = "Hidden"

}

#Function MsgBox([string]$arg1, [string]$arg2){
#
#
#   [System.Windows.Forms.MessageBox]::Show("$arg1" ,"$arg2")
#
#}

#endregion

#region Menu Button

$Menu_MailboxPerm.Add_Click({


    })
    

$Menu_MailboxInfo.Add_Click({

    ClearView
    $Grid_MailboxInfo.Visibility = "Visible"
    $Window.Title = "XAML :: Exchange Admin Tools - Mailbox General Info :: " + $whoami

    })


$Menu_Quota.Add_Click({

    Clearview
    $Grid_MailQuota.Visibility = "Visible"
    $Window.Title = "XAML :: Exchange Admin Tools - Mailbox Quota :: " + $whoami

})

$Menu_MHCIntraFID.Add_Click({

    ClearView

    $Grid_MHCIntraFID.Visibility = "Visible"

    $Window.Title = "XAML :: Exchange Admin Tools - MHC Intranet FID Creation :: " + $whoami

})

$Menu_OOOF.Add_Click({

    ClearView

    $Grid_OOOF.Visibility = "Visible"

    $Window.Title = "XAML :: Exchange Admin Tools - Out Of Office :: " + $whoami

})

$Menu_EAS.Add_Click({

    ClearView

    $Grid_EAS.Visibility = "Visible"

    $Window.Title = "XAML :: Exchange Admin Tools - Active Sync :: " + $whoami

})

#endregion

#-----------------------------
#region MHC Intranet FID Creation
#-----------------------------
#Grid_MHCIntraFID
#-----------------------------
# PS cmdlet
#-----------------------------
#
# New-Mailbox -Shared -Name "Sales Department" -DisplayName "Sales Department" -Alias Sales | Set-Mailbox -GrantSendOnBehalfTo MarketingSG | Add-MailboxPermission -User MarketingSG -AccessRights FullAccess -InheritanceType All
#
#-----------------------------
# Variables
#-----------------------------
# Txt_MHCFID_Alias
# Txt_MHCFID_DN
# Txt_MHCFID_Email
# Txt_MHCFID_Owner
# Txt_MHCFID_Backup
#
# Rd_MHCFID_BC
# Rd_MHCFID_BM
# Rd_MHCFID_BEV
# Rd_MHCFID_BDI
#
# Btn_MHCFID_Create
#-----------------------------

$Btn_MHCFID_Create.Add_Click({

    $MHCFID_Alias = $Txt_MHCFID_Alias.Text
    $MHCFID_DN = $Txt_MHCFID_DN.Text
    $MHCFID_Email = $Txt_MHCFID_Email.Text
    $MHCFID_Owner = $Txt_MHCFID_Owner.Text
    $MHCFID_Backup = $Txt_MHCFID_Backup.Text    

    write-host $MHCFID_Alias
    write-host $MHCFID_DN
    write-host $MHCFID_Email
    write-host $MHCFID_Owner
    write-host $MHCFID_Backup

    write-host New-Mailbox -shared -Name $MHCFID_Alias -DisplayName $MHCFID_DN -Alias $MHCFID_Alias
    #New-Mailbox -shared -Name "$Txt_MHCFID_Alias" -DisplayName "$Txt_MHCFID_DN" -Alias $Txt_MHCFID_Alias

    })


#-----------------------------
#region Mailbox Quota
#Grid_MailQuota
#-----------------------------
# PS cmdlet
#-----------------------------
# TotalItemSize
# TotalDeletedItemSize
# DatabaseIssueWarningQuota
# DatabaseProhibitSendQuota
# DatabaseProhibitSendReceiveQuota
#
#-----------------------------
# Variables
#-----------------------------
# Txt_DefaultQuota
# Txt_CurrentMailboxSize
# Txt_TotalDeletedItems
# Txt_IssueWarning
# Txt_ProhibitSend
# Txt_ProhibitSendReceive

# Txt_NewIssueWarn
# Txt_CalcNewProhibit
# Txt_NewProhibSendRecv
# Btn_Calculate
# Btn_ApplyQuota
#
# Txt_Username
#-----------------------------

$Btn_RefreshQuota.Add_Click({

    $User2search = $Txt_Username.Text
    
    if($User2search -eq ""){
        [System.Windows.Forms.MessageBox]::Show("User field cannot be empty" ,"Error - No user selected")
    }
    else{
        $MailboxDefaultQuotaStatus = Get-Mailbox -Identity $User2search | Select -ExpandProperty UseDatabaseQuota*
        $Txt_DefaultQuota.Text = $MailboxDefaultQuotaStatus
        $Txt_CurrentMailboxSize.Text = (Get-MailboxStatistics -Identity $User2search).TotalItemSize.Value
        $Txt_TotalDeletedItems.Text = (Get-MailboxStatistics -Identity $User2search).TotalDeletedItemSize.Value
        $Txt_IssueWarning.Text = (Get-MailboxStatistics -Identity $User2search).DatabaseIssueWarningQuota.Value 
        $Txt_ProhibitSend.Text = (Get-MailboxStatistics -Identity $User2search).DatabaseProhibitSendQuota.Value 
        $Txt_ProhibitSendReceive.Text = (Get-MailboxStatistics -Identity $User2search).DatabaseProhibitSendReceiveQuota.Value 
    }

 })

$Btn_Calculate.Add_Click({

    [int]$Issue = $Txt_CalcNewProhibit.Text
    $Txt_NewIssueWarn.Text = $issue - 50
    $Txt_NewProhibSendRecv.Text = $issue + 50

    })

$Btn_ApplyQuota.Add_Click({

    $User2search = $Txt_Username.Text

    $NewIssueWarningQuota = $Txt_NewIssueWarn.Text
    $NewProhibitSendQuota = $Txt_CalcNewProhibit.Text
    $NewProhibitSendReceiveQuota = $Txt_NewProhibSendRecv.Text

    $NewIssueWarningQuota = [STRING]$NewIssueWarningQuota + "MB"
    $NewProhibitSendQuota = [STRING]$NewProhibitSendQuota + "MB"
    $NewProhibitSendReceiveQuota = [STRING]$NewProhibitSendReceiveQuota + "MB"

    write-host "Issue Warning) " $NewIssueWarningQuota
    write-host "Prohibit Send) " $NewProhibitSendQuota
    write-host "Prohibit S/R ) " $NewProhibitSendReceiveQuota  

    Set-Mailbox -Identity $User2search -UseDatabaseQuotaDefaults:$False -ProhibitSendQuota $NewProhibitSendQuota -IssueWarningQuota $NewIssueWarningQuota -ProhibitSendReceiveQuota $NewProhibitSendReceiveQuota

    })


#endregion MailboxQuota


#region OOOF GRID

$Btn_ApplyOOOF.Add_Click({
        if ($RD_OOOFBellCan.IsChecked -eq $True){
            $OOOFUser = $Txt_OOOFUser.Text
            $BellCanadaOOOF = "<p>Thank you for contacting Bell Canada.  The individual you are trying to reach is no longer with Bell.  If you have an alternate contact within our team, please redirect your inquiry to that individual.  We will be making every effort to re-connect you with a new Bell contact in the very near future. Thank you for your understanding.</p></br></hr><p>Merci d&rsquo;avoir communiqu&eacute; avec Bell Canada. La personne que vous voulez joindre ne travaille plus avec nous. Si vous connaissez une autre personne dans notre &eacute;quipe, veuillez lui acheminer votre demande. Nous ferons notre possible pour vous orienter vers une nouvelle personne-ressource de Bell rapidement. Merci de votre compr&eacute;hension.</p></br>"
            Set-MailboxAutoReplyConfiguration -Identity $OOOFUser -AutoReplyState Enabled -InternalMessage $BellCanadaOOOF -ExternalMessage $BellCanadaOOOF
            [System.Windows.Forms.MessageBox]::Show("Bell Canada Out of Office has been applied to `n" + $OOOFUser ,"Bell Canada Out of Office")
        }
        
        if ($RD_OOOFBellMob.IsChecked -eq $True){
            $OOOFUser = $Txt_OOOFUser.Text
            $BellMobilityOOOF = "<p>Thank you for contacting Bell Mobility.  The individual you are trying to reach is no longer with Bell.  If you have an alternate contact within our team, please redirect your inquiry to that individual.  We will be making every effort to re-connect you with a new Bell contact in the very near future. Thank you for your understanding.</p></br></hr><p>Merci d&rsquo;avoir communiqu&eacute; avec Bell Mobility. La personne que vous voulez joindre ne travaille plus avec nous. Si vous connaissez une autre personne dans notre &eacute;quipe, veuillez lui acheminer votre demande. Nous ferons notre possible pour vous orienter vers une nouvelle personne-ressource de Bell rapidement. Merci de votre compr&eacute;hension.</p></br>"
            Set-MailboxAutoReplyConfiguration -Identity $OOOFUser -AutoReplyState Enabled -InternalMessage $BellMobilityOOOF -ExternalMessage $BellMobilityOOOF
            [System.Windows.Forms.MessageBox]::Show("Bell Mobility Out of Office has been applied to `n" + $OOOFUser ,"Bell Mobility Out of Office")
        }
        #DISABLE FOR NOW
        if ($RD_OOOFCustom.IsChecked -eq $True){
            $OOOFUser = $Txt_OOOFUser.Text
            $CustomOOOF = $TXT_OOOFCustom.Text
            Set-MailboxAutoReplyConfiguration -Identity $OOOFUser -AutoReplyState Enabled -InternalMessage $CustomOOOF -ExternalMessage $CustomOOOF
            [System.Windows.Forms.MessageBox]::Show("Custom Out of Office has been applied to `n" + $OOOFUser ,"Custom Out of Office")
        }
        
        if ($RD_OOOFDisabled.IsChecked -eq $True){
            $OOOFUser = $Txt_OOOFUser.Text
            Set-MailboxAutoReplyConfiguration -Identity $OOOFUser -AutoReplyState Disabled
            [System.Windows.Forms.MessageBox]::Show("Out of Office has been disabled for `n" + $OOOFUser ,"Out of Office Disabled")
        }

})

#endregion 




#region ActiveSYNC GRID 

$Btn_EnableEAS.Add_Click({
    
    $EAS_User = $Txt_EASUser.Text
    $ExchVers = get-mailbox -identity $EAS_User | Select -ExpandProperty ExchangeVersion

    if($Rd_SingleEAS.IsChecked -eq $true){
        $Txt_EASOutput.Text = "Enable eas for single user " + $EAS_User
    }
    if($Rd_BulkEAS.IsChecked -eq $true){
        $Txt_EASOutput.Text = "Bulk EAS Activation"
    }
    if($Rd_DisableEAS.IsChecked -eq $true){
        $Txt_EASOutput.Text = "Disable EAS for user " + $EAS_User
    }
    if($Rd_CheckEAS.IsChecked -eq $true){

        $StatusEAS = get-casmailbox -identity $EAS_User | select -ExpandProperty ActiveSyncEnabled
        if($StatusEAS -eq "true"){
            [System.Windows.Forms.MessageBox]::Show("Activesync is enabled for " + $EAS_User , "ActiveSync Status")
            
            if ($ExchVers -eq "0.20 (15.0.0.0)") {
                get-MobileDevice -mailbox $EAS_User | Select DeviceType,DeviceID | Out-File tmp.txt
                $Txt_EASOutput.Text = Get-Content tmp.txt 

            }  
            if ($ExchVers -eq "0.1 (8.0.535.0)"){
                #Get-ActiveSyncDeviceStatistics -mailbox $Eas_User | Select DeviceType,DeviceID | Out-File tmp.txt
                $Txt_EASOutput.Text = "User is on Exchange 2007, please use the following cmdlet on a Exch2007 powershell Get-ActiveSyncDeviceStatistics -mailbox "+ $Eas_User + "| Select DeviceType,DeviceID"
            }

        }
        else{
            [System.Windows.Forms.MessageBox]::Show("Activesync is disabled for " + $EAS_User , "ActiveSync Status")
        }
    }
})


#endregion


$Window.Showdialog() | Out-Null
