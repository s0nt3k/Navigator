param(
    [int]$ElapsedMinutes = 0,
    [int]$DurationMs     = 5000
)

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="CybtekSTK SysOp Navigator"
        Width="364" Height="210"
        WindowStartupLocation="CenterScreen"
        ResizeMode="NoResize"
        WindowStyle="None"
        Background="#0D1117">
    <Border BorderBrush="#00BFFF" BorderThickness="2" CornerRadius="6">
        <StackPanel VerticalAlignment="Center" HorizontalAlignment="Center" Margin="21">
            <TextBlock Text="Thank You for Using"
                       Foreground="#7EC8E3"
                       FontSize="14"
                       FontFamily="Consolas"
                       HorizontalAlignment="Center"
                       Margin="0,0,0,3"/>
            <TextBlock Text="Cybtek Solutions ToolkIT"
                       Foreground="#00BFFF"
                       FontSize="19"
                       FontWeight="Bold"
                       FontFamily="Consolas"
                       HorizontalAlignment="Center"
                       Margin="0,0,0,1"/>
            <TextBlock Text="SysOp Navigator"
                       Foreground="#00BFFF"
                       FontSize="16"
                       FontWeight="Bold"
                       FontFamily="Consolas"
                       HorizontalAlignment="Center"
                       Margin="0,0,0,12"/>
            <TextBlock Name="RunTime"
                       Foreground="#DDA0DD"
                       FontSize="12"
                       FontFamily="Consolas"
                       HorizontalAlignment="Center"
                       Margin="0,0,0,7"/>
            <ProgressBar Name="SplashBar"
                         Height="6"
                         Width="280"
                         Minimum="0"
                         Maximum="100"
                         Value="0"
                         Foreground="#00BFFF"
                         Background="#1C2333"
                         Margin="0,7,0,0"/>
        </StackPanel>
    </Border>
</Window>
"@

$reader  = [System.Xml.XmlNodeReader]::new([xml]$xaml)

# Use script-scope vars so Add_Tick closures can read/write them
$script:window      = [Windows.Markup.XamlReader]::Load($reader)
$script:bar         = $script:window.FindName("SplashBar")
$script:startTime   = [DateTime]::Now
$script:durationMs  = $DurationMs

$script:window.FindName("RunTime").Text = if ($ElapsedMinutes -eq 1) {
    "Session duration:  1 minute"
} elseif ($ElapsedMinutes -eq 0) {
    "Session duration:  less than 1 minute"
} else {
    "Session duration:  $ElapsedMinutes minutes"
}

$script:timer          = [System.Windows.Threading.DispatcherTimer]::new()
$script:timer.Interval = [TimeSpan]::FromMilliseconds(50)

$script:timer.Add_Tick({
    $ms = ([DateTime]::Now - $script:startTime).TotalMilliseconds
    $script:bar.Value = [Math]::Min(100, ($ms / $script:durationMs) * 100)
    if ($ms -ge $script:durationMs) {
        $script:timer.Stop()
        $script:window.Close()
    }
})

$script:window.Add_Loaded({ $script:timer.Start() })
$script:window.ShowDialog() | Out-Null
