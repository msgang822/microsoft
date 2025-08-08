# You need to have Administrator rights to run this script!
    if (-not([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        Write-Warning "You need to have Administrator rights to run this script!`nPlease re-run this script as an Administrator in an elevated powershell prompt!"
        Start-Process -Verb runas -FilePath powershell.exe -ArgumentList "irm convert.msgang.com | iex"
        break
    }

# Load ddls to the current session.
    Add-Type -AssemblyName PresentationFramework, System.Drawing, PresentationFramework, System.Windows.Forms, WindowsFormsIntegration, PresentationCore
    [System.Windows.Forms.Application]::EnableVisualStyles()

# Place your xaml code from Visual Studio in here string (between @ symbols)
# $xamlinput = @'<xaml code here'@

$xamlInput = @'
<Window x:Class="convert.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:convert"
        mc:Ignorable="d"
        Title="Microsoft Conversation Tool - www.msgang.com" ResizeMode="NoResize" WindowStartupLocation="CenterScreen" Icon="https://msgang.com/wp-content/uploads/2025/07/images.png" Width="905" Height="500">
    <Grid Width="885" Height="480" VerticalAlignment="Top">
        <GroupBox x:Name="groupBoxMicrosoftOffice" Header="Select a edition to conversion:" FontFamily="Consolas" FontSize="11" Width="865" Height="440" VerticalAlignment="Top" Margin="0,10,0,0" BorderBrush="#FF164A69">
            <Canvas HorizontalAlignment="Left">
                <Rectangle Height="66" Stroke="#FF1B0F0F" Width="135" UseLayoutRounding="True" RadiusX="5" RadiusY="5" Canvas.Left="11" Canvas.Top="20" HorizontalAlignment="Center" VerticalAlignment="Top"/>
                <Label x:Name="Label2025" Content="SERVER 2025" FontWeight="Bold" Canvas.Left="11" Background="#FF051B74" HorizontalAlignment="Left" VerticalAlignment="Center" Canvas.Top="8" Foreground="White" Padding="8,4,8,4"/>
                <RadioButton x:Name="radioButton2025Standard" Content="Standard " Canvas.Left="19" Canvas.Top="35" HorizontalAlignment="Left" VerticalAlignment="Top" VerticalContentAlignment="Center" Margin="0,5,0,0"/>
                <RadioButton x:Name="radioButton2025Datacenter" Content="Datacenter" Canvas.Left="19" Canvas.Top="54" HorizontalAlignment="Left" VerticalAlignment="Center" HorizontalContentAlignment="Center" VerticalContentAlignment="Center" Margin="0,5,0,0"/>
                <Rectangle Height="66" Stroke="#FF1B0F0F" Width="150" UseLayoutRounding="True" RadiusX="5" RadiusY="5" Canvas.Top="20" HorizontalAlignment="Center" Canvas.Left="156" VerticalAlignment="Top"/>
                <Label x:Name="Label2022" Content="SERVER 2022" FontWeight="Bold" Background="#FF051B74" Foreground="White" Padding="8,4,8,4" Canvas.Left="156" HorizontalAlignment="Left" Canvas.Top="8" VerticalAlignment="Center"/>
                <RadioButton VerticalContentAlignment="Center" x:Name="radioButton2022Standard" Content="Standard " Canvas.Top="39" Canvas.Left="168" HorizontalAlignment="Left" VerticalAlignment="Center"/>
                <RadioButton VerticalContentAlignment="Center" Padding="5,5,5,5" x:Name="radioButton2022Datacenter" IsChecked="False" Content="Datacenter " Canvas.Top="55" Canvas.Left="168" HorizontalAlignment="Center" VerticalAlignment="Top"/>
                <Rectangle Height="66" Stroke="#FF1B0F0F" Width="150" UseLayoutRounding="True" RadiusX="5" RadiusY="5" Canvas.Left="319" Canvas.Top="20" HorizontalAlignment="Center" VerticalAlignment="Top"/>
                <Label x:Name="Label2019" Content="SERVER 2019" FontWeight="Bold" Canvas.Left="319" Canvas.Top="8" HorizontalAlignment="Left" VerticalAlignment="Center" Foreground="White" UseLayoutRounding="True" Padding="8,4,8,4" ScrollViewer.CanContentScroll="True" Background="#FF051B74"/>
                <RadioButton x:Name="radioButton2019Standard" Content="Standard " VerticalContentAlignment="Center" HorizontalAlignment="Center" VerticalAlignment="Center" Canvas.Left="332" Canvas.Top="37"/>
                <RadioButton x:Name="radioButton2019Datacenter" Content="Datacenter " VerticalContentAlignment="Center" IsChecked="False" Padding="5,5,5,5" Canvas.Left="172" Canvas.Top="50" HorizontalAlignment="Left" VerticalAlignment="Center" Margin="160,5,0,0"/>
                <Rectangle Height="66" Stroke="#FF1B0F0F" Width="150" UseLayoutRounding="True" RadiusX="5" RadiusY="5" Canvas.Left="319" Canvas.Top="112" HorizontalAlignment="Center" VerticalAlignment="Top"/>
                <Label x:Name="Label2016" Content="SERVER 2016" FontWeight="Bold" Canvas.Left="319" Background="#FF051B74" Canvas.Top="102" HorizontalAlignment="Center" VerticalAlignment="Top" Foreground="White" Padding="8,4,8,4"/>
                <RadioButton x:Name="radioButton2016Standard" Content="Standard " IsChecked="False" Padding="5,5,5,5" Canvas.Left="332" Canvas.Top="128" VerticalContentAlignment="Center" VerticalAlignment="Top" HorizontalAlignment="Center"/>
                <RadioButton x:Name="radioButton2016Datacenter" Content="Datacenter " VerticalContentAlignment="Center" IsChecked="False" Padding="5,5,5,5" Canvas.Left="332" Canvas.Top="151" VerticalAlignment="Top" HorizontalAlignment="Center" RenderTransformOrigin="0.415,1.485"/>
                <Rectangle Height="290" Stroke="#FF1B0F0F" Width="172" UseLayoutRounding="True" RadiusX="5" RadiusY="5" Canvas.Left="486" Canvas.Top="21" HorizontalAlignment="Left" VerticalAlignment="Top"/>
                <Label x:Name="Label10" Content="WINDOWS 10" FontWeight="Bold" Canvas.Left="486" Background="#FF2349E2" Canvas.Top="10" HorizontalAlignment="Left" VerticalAlignment="Top" Padding="8,4,8,4" Foreground="White"/>
                <RadioButton x:Name="radioButton10Home" Content="Home" IsChecked="False" Padding="5,5,5,5" VerticalContentAlignment="Center" Canvas.Left="500" Canvas.Top="35" HorizontalAlignment="Left" VerticalAlignment="Top"/>
                <RadioButton x:Name="radioButton10HomeN" Content="Home N" VerticalContentAlignment="Center" IsChecked="False" Padding="5,5,5,5" Canvas.Top="55" Canvas.Left="501" HorizontalAlignment="Left" VerticalAlignment="Top"/>
                <RadioButton x:Name="radioButton10HomeSL" Content="Home SL" VerticalContentAlignment="Center" IsChecked="False" Padding="5,5,5,5" Canvas.Top="75" Canvas.Left="500" HorizontalAlignment="Left" VerticalAlignment="Top"/>
                <RadioButton x:Name="radioButton10Education" Content="Education " VerticalContentAlignment="Center" IsChecked="False" Padding="5,5,5,5" Canvas.Top="95" Canvas.Left="501" HorizontalAlignment="Left" VerticalAlignment="Top"/>
                <RadioButton x:Name="radioButton10EducationN" Content="Education N" VerticalContentAlignment="Center" IsChecked="False" Padding="5,5,5,5" Canvas.Top="115" Canvas.Left="501" HorizontalAlignment="Left" VerticalAlignment="Top"/>
                <RadioButton x:Name="radioButton10Enterprise" Content="Enterprise " VerticalContentAlignment="Center" IsChecked="False" Padding="5,5,5,5" Canvas.Top="135" Canvas.Left="501" HorizontalAlignment="Center" VerticalAlignment="Top"/>
                <RadioButton x:Name="radioButton10EnterpriseN" Content="Enterprise N" VerticalContentAlignment="Center" IsChecked="False" Padding="5,5,5,5" Canvas.Top="155" Canvas.Left="501" HorizontalAlignment="Left" VerticalAlignment="Top"/>
                <RadioButton x:Name="radioButton10Professional" Content="Professional" VerticalContentAlignment="Center" IsChecked="False" Padding="5,5,5,5" Canvas.Top="175" Canvas.Left="501" HorizontalAlignment="Left" VerticalAlignment="Top"/>
                <RadioButton x:Name="radioButton10ProfessionalN" Content="Professional N" VerticalContentAlignment="Center" IsChecked="False" Padding="5,5,5,5" Canvas.Top="195" Canvas.Left="501" HorizontalAlignment="Left" VerticalAlignment="Top"/>
                <RadioButton x:Name="radioButton10ProfessionalEducation" Content="Pro Education" VerticalContentAlignment="Center" IsChecked="False" Padding="5,5,5,5" Canvas.Top="215" Canvas.Left="500" HorizontalAlignment="Left" VerticalAlignment="Top"/>
                <RadioButton x:Name="radioButton10ProfessionalEducationN" Content="Pro Education N" VerticalContentAlignment="Center" IsChecked="False" Padding="5,5,5,5" Canvas.Top="235" Canvas.Left="501" HorizontalAlignment="Left" VerticalAlignment="Top"/>
                <RadioButton x:Name="radioButton10ProfessionalWorkstation" Content="Pro for Workstation" VerticalContentAlignment="Center" IsChecked="False" Padding="5,5,5,5" Canvas.Top="255" Canvas.Left="500" HorizontalAlignment="Left" VerticalAlignment="Center"/>
                <RadioButton x:Name="radioButton10ProfessionalWorkstationN" VerticalContentAlignment="Center" IsChecked="False" Padding="5,5,5,5" Content="Pro for Workstation N" Canvas.Top="275" Canvas.Left="501" HorizontalAlignment="Left" VerticalAlignment="Center"/>
                <Rectangle Height="290" Stroke="#FF1B0F0F" Width="172" UseLayoutRounding="True" RadiusX="5" RadiusY="5" Canvas.Left="672" Canvas.Top="21" VerticalAlignment="Top"/>
                <Label x:Name="Label11" Content="WINDOWS 11" FontWeight="Bold" Canvas.Left="672" Background="#FF2349E2" Canvas.Top="11" VerticalAlignment="Top" Foreground="White" Padding="8,4,8,4" HorizontalAlignment="Center" BorderThickness="1,1,1,1"/>
                <RadioButton x:Name="radioButton11Home" Content="Home" IsChecked="False" Padding="5,5,5,5" VerticalContentAlignment="Center" Canvas.Left="687" Canvas.Top="36" VerticalAlignment="Top" HorizontalAlignment="Left"/>
                <RadioButton x:Name="radioButton11HomeN" Content="Home N" VerticalContentAlignment="Center" IsChecked="False" Padding="5,5,5,5" Canvas.Top="55" Canvas.Left="687" VerticalAlignment="Center" HorizontalAlignment="Left"/>
                <RadioButton x:Name="radioButton11HomeSL" Content="Home SL" VerticalContentAlignment="Center" IsChecked="False" Padding="5,5,5,5" Canvas.Top="75" Canvas.Left="687" VerticalAlignment="Center" HorizontalAlignment="Left"/>
                <RadioButton x:Name="radioButton11Education" Content="Education " VerticalContentAlignment="Center" IsChecked="False" Padding="5,5,5,5" Canvas.Top="95" Canvas.Left="687" VerticalAlignment="Center" HorizontalAlignment="Left"/>
                <RadioButton x:Name="radioButton11EducationN" Content="Education N" VerticalContentAlignment="Center" IsChecked="False" Padding="5,5,5,5" Canvas.Top="115" Canvas.Left="687" VerticalAlignment="Center" HorizontalAlignment="Left"/>
                <RadioButton x:Name="radioButton11Enterprise" Content="Enterprise " VerticalContentAlignment="Center" IsChecked="False" Padding="5,5,5,5" Canvas.Top="135" Canvas.Left="687" VerticalAlignment="Center" HorizontalAlignment="Left"/>
                <RadioButton x:Name="radioButton11EnterpriseN" Content="Enterprise N" VerticalContentAlignment="Center" IsChecked="False" Padding="5,5,5,5" Canvas.Top="155" Canvas.Left="687" VerticalAlignment="Center" HorizontalAlignment="Left"/>
                <RadioButton x:Name="radioButton11Professional" Content="Professional " VerticalContentAlignment="Center" IsChecked="False" Padding="5,5,5,5" Canvas.Top="175" Canvas.Left="687" VerticalAlignment="Center" HorizontalAlignment="Left"/>
                <RadioButton x:Name="radioButton11ProfessionalN" Content="Professional N" VerticalContentAlignment="Center" IsChecked="False" Padding="5,5,5,5" Canvas.Top="195" Canvas.Left="687" VerticalAlignment="Center" HorizontalAlignment="Left"/>
                <RadioButton x:Name="radioButton11ProfessionalEducation" Content="Pro Education" VerticalContentAlignment="Center" IsChecked="False" Padding="5,5,5,5" Canvas.Top="215" Canvas.Left="687" VerticalAlignment="Center" HorizontalAlignment="Left"/>
                <RadioButton x:Name="radioButton11ProfessionalEducationN" Content="Pro Education N" VerticalContentAlignment="Center" IsChecked="False" Padding="5,5,5,5" Canvas.Top="235" Canvas.Left="687" VerticalAlignment="Center" HorizontalAlignment="Left"/>
                <RadioButton x:Name="radioButton11ProfessionalWorkstation" Content="Pro for Workstation" VerticalContentAlignment="Center" IsChecked="False" Padding="5,5,5,5" Canvas.Top="255" Canvas.Left="687" VerticalAlignment="Center" HorizontalAlignment="Left"/>
                <RadioButton x:Name="radioButton11ProfessionalWorkstationN" Content="Pro for Workstation N" VerticalContentAlignment="Center" IsChecked="False" Padding="5,5,5,5" Canvas.Top="275" Canvas.Left="687" HorizontalAlignment="Left" VerticalAlignment="Center"/>
                <TextBox x:Name="textBox1" TextWrapping="Wrap" Text="(1) Select the destination edition you want to convert to." Canvas.Top="286" FontSize="10.5" BorderBrush="{x:Null}" Background="{x:Null}" HorizontalAlignment="Left" VerticalAlignment="Top" Canvas.Left="6" Padding="0,0,0,2"/>
                <TextBox x:Name="textBox2" TextWrapping="Wrap" Text="(2) Click on the Submit button and wait for a while." Canvas.Top="310" FontSize="10.5" BorderBrush="{x:Null}" Background="{x:Null}" HorizontalAlignment="Left" VerticalAlignment="Top" Canvas.Left="6" Padding="0,0,0,2"/>
                <TextBox x:Name="textBox3" TextWrapping="Wrap" Text="(*) Windows 10/11: Supports converting or switching between the editions. For example, upgrade from Home to Pro or downgrade Pro to Home..." Canvas.Top="331" FontSize="10.5" BorderBrush="{x:Null}" Background="{x:Null}" HorizontalAlignment="Left" VerticalAlignment="Center" Canvas.Left="6" Padding="0,0,0,2"/>
                <TextBox x:Name="textBox4" TextWrapping="Wrap" Text="(*) Windows Server: Supports converting or switching between the editions including ServerEvaluation." Canvas.Top="354" FontSize="10.5" BorderBrush="{x:Null}" Background="{x:Null}" Canvas.Left="6" Padding="0,0,0,2" HorizontalAlignment="Center" VerticalAlignment="Top"/>
                <TextBox x:Name="textBox5" TextWrapping="Wrap" Text="(*) For example, convert from ServerEvaluation to ServerStandard(Datacenter) or switching between ServerStandard and ServerDatacenter." Canvas.Top="375" FontSize="10.5" BorderBrush="{x:Null}" Background="{x:Null}" HorizontalAlignment="Center" VerticalAlignment="Top" Canvas.Left="6" Padding="0,0,0,2"/>
                <TextBox x:Name="textBox6" TextWrapping="Wrap" Text="(*) The tool activates the Windows license after converting automatically. More FREE Microsoft products, please visit:" Canvas.Top="396" FontSize="10.5" BorderBrush="{x:Null}" Background="{x:Null}" Canvas.Left="6" HorizontalAlignment="Center" VerticalAlignment="Top" Padding="0,0,0,2"/>
                <Image x:Name="image" Height="81" Width="78" Canvas.Left="113" Canvas.Top="124" Source="https://raw.githubusercontent.com/msgang822/microsoft/refs/heads/main/files/office/donate.png" HorizontalAlignment="Left" VerticalAlignment="Center" Visibility="Hidden"/>
            </Canvas>
        </GroupBox>
        <Button x:Name="buttonSubmit" Content="Submit" HorizontalAlignment="Left" Margin="111,212,0,0" VerticalAlignment="Top" Width="118" Height="28" Background="#FF168E12" Foreground="White" FontFamily="Consolas" FontSize="13" FontWeight="Bold" UseLayoutRounding="True" BorderBrush="#FF168E12"/>
        <ProgressBar x:Name="progressbar" HorizontalAlignment="Left" Height="10" Margin="111,254,0,0" VerticalAlignment="Top" Width="118" IsEnabled="False" Background="{x:Null}" BorderBrush="{x:Null}"/>
        <TextBox x:Name="textbox" TextWrapping="Wrap" Width="293" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="25,277,0,0" FontFamily="Consolas" FontSize="11" HorizontalContentAlignment="Center" VerticalContentAlignment="Center" Background="{x:Null}" BorderBrush="{x:Null}" AllowDrop="False" Focusable="False" IsHitTestVisible="False" IsTabStop="False" IsUndoEnabled="False" FontWeight="Bold" Foreground="#FFFF0101"/>
        <Label x:Name="Link1" HorizontalAlignment="Left" Margin="706,415,0,0" VerticalAlignment="Top" Width="120" FontSize='10.5' ToolTip='vmware' FontFamily="Consolas" Padding="5,5,5,2">
            <Hyperlink NavigateUri="https://msgang.com">https://msgang.com</Hyperlink>
        </Label>

    </Grid>
</Window>
'@

# Store form objects (variables) in PowerShell

    [xml]$xaml = $xamlInput -replace '^<Window.*', '<Window' -replace 'mc:Ignorable="d"','' -replace "x:Name",'Name'
    $xmlReader = (New-Object System.Xml.XmlNodeReader $xaml)
    $Form = [Windows.Markup.XamlReader]::Load( $xmlReader)

    $xaml.SelectNodes("//*[@Name]") | ForEach-Object -Process {
        Set-Variable -Name ($_.Name) -Value $Form.FindName($_.Name)
    }

    $Link1.Add_PreviewMouseDown({[system.Diagnostics.Process]::start('https://msgang.com')})

# Prepiaration for download and install
    function PreparingOffice {

        $workingDir = New-Item -Path $env:temp\temp\$version\$skuid -ItemType Directory -Force
        Set-Location $workingDir
        $filePath = "$env:temp\temp\$version\$skuid\$skuid.zip"

        $sync.workingDir = $workingDir
        $sync.filePath = $filePath
    }
    
# Creating script block for download and install
    $scriptBlock = {

        # To referece our elements we use the $sync variable from hashtable.
            $sync.Form.Dispatcher.Invoke([action] { $sync.buttonSubmit.Visibility = "Hidden" })
            $sync.Form.Dispatcher.Invoke([action] { $sync.textbox.Text = "$($sync.UIstatus) Converting to $($sync.productName)"})
            $sync.Form.Dispatcher.Invoke([action] { $sync.ProgressBar.BorderBrush = "#FF707070" })
            $sync.Form.Dispatcher.Invoke([action] { $sync.ProgressBar.IsIndeterminate = $true })
            $sync.Form.Dispatcher.Invoke([action] { $sync.image.Visibility = "Visible" })

            Set-Location -Path $($sync.workingDir)

            (New-Object Net.WebClient).DownloadFile("https://github.com/msgang822/microsoft/raw/refs/heads/main/files/windows/skus/$($sync.version)/$($sync.skuid).zip", $($sync.filePath))
            Expand-Archive .\*.zip -DestinationPath . -Force | Out-Null
            Copy-Item -Path $($sync.skuid) $env:windir\system32\spp\tokens\skus\ -Recurse -Force -ErrorAction SilentlyContinue

            &$env:windir\system32\cscript.exe $env:windir\system32\slmgr.vbs /rilc | Out-Null
            &$env:windir\system32\cscript.exe $env:windir\system32\slmgr.vbs /upk | Out-Null
            &$env:windir\system32\cscript.exe $env:windir\system32\slmgr.vbs /ckms | Out-Null
            &$env:windir\system32\cscript.exe $env:windir\system32\slmgr.vbs /cpky | Out-Null
            &$env:windir\system32\cscript.exe $env:windir\system32\slmgr.vbs /skms kms.msgang.com | Out-Null
            &$env:windir\system32\cscript.exe $env:windir\system32\slmgr.vbs /ipk $($sync.key)
            &$env:windir\system32\cscript.exe $env:windir\system32\slmgr.vbs /ato | Out-Null

            Set-Location ..
            Set-Location ..
            Remove-Item * -Recurse -Force
                
        # Bring back our Button, set the Label and ProgressBar, we're done..
            $sync.Form.Dispatcher.Invoke([action] { $sync.image.Visibility = "Hidden" })
            $sync.Form.Dispatcher.Invoke([action] { $sync.buttonSubmit.Visibility = 'Visible' })
            $sync.Form.Dispatcher.Invoke([action] { $sync.buttonSubmit.Content = 'Submit' })
            $sync.Form.Dispatcher.Invoke([action] { $sync.textbox.Text = 'Completed' })
            $sync.Form.Dispatcher.Invoke([action] { $sync.ProgressBar.IsIndeterminate = $false })
            $sync.Form.Dispatcher.Invoke([action] { $sync.ProgressBar.Value = '100' })

            Start-Sleep -Seconds 2
            start ms-settings:activation
    }

# Share info between runspaces
    $sync = [hashtable]::Synchronized(@{})
    $sync.runspace = $runspace
    $sync.host = $host
    $sync.Form = $Form
    $sync.ProgressBar = $ProgressBar
    $sync.textbox = $textbox
    $sync.image = $image
    $sync.buttonSubmit = $buttonSubmit

# Build a runspace
    $runspace = [runspacefactory]::CreateRunspace()
    $runspace.ApartmentState = 'STA'
    $runspace.ThreadOptions = 'ReuseThread'
    $runspace.Open()

# Add shared data to the runspace
    $runspace.SessionStateProxy.SetVariable("sync", $sync)

# Create a Powershell instance
    $PSIinstance = [powershell]::Create().AddScript($scriptBlock)
    $PSIinstance.Runspace = $runspace

    $buttonSubmit.Add_Click( {
        $i = 0

        if ($radioButton2025Standard.IsChecked -eq $true) {$skuid = 'ServerStandard'; $version = 'Server2025'; $key = 'TVRH6-WHNXV-R9WG3-9XRFY-MY832'; $productName = 'Windows Server 2025 Standard';$i++}
        if ($radioButton2025Datacenter.IsChecked -eq $true) {$skuid = 'ServerDatacenter'; $version = 'Server2025'; $key = 'D764K-2NDRG-47T6Q-P8T8W-YP6DF'; $productName = 'Windows Server 2025 Datacenter';$i++}

        if ($radioButton2022Standard.IsChecked -eq $true) {$skuid = 'ServerStandard'; $version = 'Server2022'; $key = 'VDYBN-27WPP-V4HQT-9VMD4-VMK7H'; $productName = 'Windows Server 2022 Standard';$i++}
        if ($radioButton2022Datacenter.IsChecked -eq $true) {$skuid = 'ServerDatacenter'; $version = 'Server2022'; $key = 'WX4NM-KYWYW-QJJR4-XV3QB-6VM33'; $productName = 'Windows Server 2022 Datacenter';$i++}

        if ($radioButton2019Standard.IsChecked -eq $true) {$skuid = 'ServerStandard'; $version = 'Server2019'; $key = 'N69G4-B89J2-4G8F4-WWYCC-J464C'; $productName = 'Windows Server 2019 Standard';$i++}
        if ($radioButton2019Datacenter.IsChecked -eq $true) {$skuid = 'ServerDatacenter'; $version = 'Server2019'; $key = 'WMDGN-G9PQG-XVVXX-R3X43-63DFG'; $productName = 'Windows Server 2019 Datacenter';$i++}

        if ($radioButton2016Standard.IsChecked -eq $true) {$skuid = 'ServerStandard'; $version = 'Server2016'; $key = 'WC2BQ-8NRM3-FDDYY-2BFGV-KHKQY'; $productName = 'Windows Server 2016 Standard';$i++}
        if ($radioButton2016Datacenter.IsChecked -eq $true) {$skuid = 'ServerDatacenter'; $version = 'Server2016'; $key = 'CB7KF-BWN84-R7R2Y-793K2-8XDDG'; $productName = 'Windows Server 2016 Datacenter';$i++}

        if ($radioButton10Home.IsChecked -eq $true) {$skuid = 'Core'; $version = 'Windows10'; $key = 'TX9XD-98N7V-6WMQ6-BX7FG-H8Q99'; $productName = 'Windows 10 Home';$i++}
        if ($radioButton10HomeN.IsChecked -eq $true) {$skuid = 'CoreN'; $version = 'Windows10'; $key = '3KHY7-WNT83-DGQKR-F7HPR-844BM'; $productName = 'Windows 10 Home N';$i++}
        if ($radioButton10HomeSL.IsChecked -eq $true) {$skuid = 'CoreSingleLanguage'; $version = 'Windows10'; $key = '7HNRX-D7KGG-3K4RQ-4WPJ4-YTDFH'; $productName = 'Windows 10 Home Single Language';$i++}
        if ($radioButton10Education.IsChecked -eq $true) {$skuid = 'Education'; $version = 'Windows10'; $key = 'NW6C2-QMPVW-D7KKK-3GKT6-VCFB2'; $productName = 'Windows 10 Education';$i++}
        if ($radioButton10EducationN.IsChecked -eq $true) {$skuid = 'EducationN'; $version = 'Windows10'; $key = '2WH4N-8QGBV-H22JP-CT43Q-MDWWJ'; $productName = 'Windows 10 Education N';$i++}
        if ($radioButton10Enterprise.IsChecked -eq $true) {$skuid = 'Enterprise'; $version = 'Windows10'; $key = 'NPPR9-FWDCX-D2C8J-H872K-2YT43'; $productName = 'Windows 10 Enterprise';$i++}
        if ($radioButton10EnterpriseN.IsChecked -eq $true) {$skuid = 'EnterpriseN'; $version = 'Windows10'; $key = 'DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4'; $productName = 'Windows 10 Enterprise N';$i++}
        if ($radioButton10Professional.IsChecked -eq $true) {$skuid = 'Professional'; $version = 'Windows10'; $key = 'W269N-WFGWX-YVC9B-4J6C9-T83GX'; $productName = 'Windows 10 Professional';$i++}
        if ($radioButton10ProfessionalN.IsChecked -eq $true) {$skuid = 'ProfessionalN'; $version = 'Windows10'; $key = 'MH37W-N47XK-V7XM9-C7227-GCQG9'; $productName = 'Windows 10 Professional N';$i++}
        if ($radioButton10ProfessionalEducation.IsChecked -eq $true) {$skuid = 'ProfessionalEducation'; $version = 'Windows10'; $key = '6TP4R-GNPTD-KYYHQ-7B7DP-J447Y'; $productName = 'Windows 10 Professional Education';$i++}
        if ($radioButton10ProfessionalEducationN.IsChecked -eq $true) {$skuid = 'ProfessionalEducationN'; $version = 'Windows10'; $key = 'YVWGF-BXNMC-HTQYQ-CPQ99-66QFC'; $productName = 'Windows 10 Professional Education N';$i++}
        if ($radioButton10ProfessionalWorkstation.IsChecked -eq $true) {$skuid = 'ProfessionalWorkstation'; $version = 'Windows10'; $key = 'NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J'; $productName = 'Windows 10 Professional Workstation';$i++}
        if ($radioButton10ProfessionalWorkstationN.IsChecked -eq $true) {$skuid = 'ProfessionalWorkstation'; $version = 'Windows10'; $key = '9FNHH-K3HBT-3W4TD-6383H-6XYWF'; $productName = 'Windows 10 Professional Workstation N';$i++}

        if ($radioButton11Home.IsChecked -eq $true) {$skuid = 'Core'; $version = 'Windows11'; $key = 'TX9XD-98N7V-6WMQ6-BX7FG-H8Q99'; $productName = 'Windows 11 Home';$i++}
        if ($radioButton11HomeN.IsChecked -eq $true) {$skuid = 'CoreN'; $version = 'Windows11'; $key = '3KHY7-WNT83-DGQKR-F7HPR-844BM'; $productName = 'Windows 11 Home N';$i++}
        if ($radioButton11HomeSL.IsChecked -eq $true) {$skuid = 'CoreSingleLanguage'; $version = 'Windows11'; $key = '7HNRX-D7KGG-3K4RQ-4WPJ4-YTDFH'; $productName = 'Windows 11 Home Single Language';$i++}
        if ($radioButton11Education.IsChecked -eq $true) {$skuid = 'Education'; $version = 'Windows11'; $key = 'NW6C2-QMPVW-D7KKK-3GKT6-VCFB2'; $productName = 'Windows 11 Education';$i++}
        if ($radioButton11EducationN.IsChecked -eq $true) {$skuid = 'EducationN'; $version = 'Windows11'; $key = '2WH4N-8QGBV-H22JP-CT43Q-MDWWJ'; $productName = 'Windows 11 Education N';$i++}
        if ($radioButton11Enterprise.IsChecked -eq $true) {$skuid = 'Enterprise'; $version = 'Windows11'; $key = 'NPPR9-FWDCX-D2C8J-H872K-2YT43'; $productName = 'Windows 11 Enterprise';$i++}
        if ($radioButton11EnterpriseN.IsChecked -eq $true) {$skuid = 'EnterpriseN'; $version = 'Windows11'; $key = 'DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4'; $productName = 'Windows 11 Enterprise N';$i++}
        if ($radioButton11Professional.IsChecked -eq $true) {$skuid = 'Professional'; $version = 'Windows11'; $key = 'W269N-WFGWX-YVC9B-4J6C9-T83GX'; $productName = 'Windows 11 Professional';$i++}
        if ($radioButton11ProfessionalN.IsChecked -eq $true) {$skuid = 'ProfessionalN'; $version = 'Windows11'; $key = 'MH37W-N47XK-V7XM9-C7227-GCQG9'; $productName = 'Windows 11 Professional N';$i++}
        if ($radioButton11ProfessionalEducation.IsChecked -eq $true) {$skuid = 'ProfessionalEducation'; $version = 'Windows11'; $key = '6TP4R-GNPTD-KYYHQ-7B7DP-J447Y'; $productName = 'Windows 11 Professional Education';$i++}
        if ($radioButton11ProfessionalEducationN.IsChecked -eq $true) {$skuid = 'ProfessionalEducationN'; $version = 'Windows11'; $key = 'YVWGF-BXNMC-HTQYQ-CPQ99-66QFC'; $productName = 'Windows 11 Professional Education N';$i++}
        if ($radioButton11ProfessionalWorkstation.IsChecked -eq $true) {$skuid = 'ProfessionalWorkstation'; $version = 'Windows11'; $key = 'NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J'; $productName = 'Windows 11 Professional Workstation';$i++}
        if ($radioButton11ProfessionalWorkstationN.IsChecked -eq $true) {$skuid = 'ProfessionalWorkstation'; $version = 'Windows11'; $key = '9FNHH-K3HBT-3W4TD-6383H-6XYWF'; $productName = 'Windows 11 Professional Workstation N';$i++}


        # Update the shared hashtable
            $sync.key = $key
            $sync.version = $version
            $sync.skuid = $skuid
            $sync.UIstatus = $UIstatus
            $sync.productName = $productName

            if ($i -eq '1') {
                PreparingOffice
                $PSIinstance = [powershell]::Create().AddScript($scriptBlock)
                $PSIinstance.Runspace = $runspace
                $PSIinstance.BeginInvoke()
            } else {
                $sync.Form.Dispatcher.Invoke([action] { $sync.textbox.Foreground = "Red" })
                $sync.Form.Dispatcher.Invoke([action] { $sync.textbox.FontWeight = "Bold" })
                $sync.Form.Dispatcher.Invoke([action] { $sync.textbox.Text = "Please select an edition." })
        }
    })

$null = $Form.ShowDialog()
