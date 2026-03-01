param(
    [int]$DurationMs = 4000
)

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

# Local cache path — downloaded once, reused on subsequent runs.
# Delete splash-image.png to force a fresh download.
$script:localImage = Join-Path $PSScriptRoot "splash-image.png"
$rawUrl            = "https://raw.githubusercontent.com/s0nt3k/Scripts/main/Assets/Images/splash-image.png"

if (-not (Test-Path $script:localImage)) {
    try {
        $wc = [System.Net.WebClient]::new()
        $wc.DownloadFile($rawUrl, $script:localImage)
    } catch {
        # No network or download failed — splash will show without the image
    }
}

$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="CybtekSTK SysOp Navigator"
        Width="444" Height="524"
        WindowStartupLocation="CenterScreen"
        ResizeMode="NoResize"
        WindowStyle="None"
        Background="#0D1117">
    <Border BorderBrush="#00BFFF" BorderThickness="2" CornerRadius="6">
        <StackPanel VerticalAlignment="Center" HorizontalAlignment="Center" Margin="21">
            <Image Name="SplashImage"
                   Width="380" Height="380"
                   Stretch="Uniform"
                   Margin="0,0,0,10"/>
            <TextBlock Text="Cybtek Solutions ToolkIT"
                       Foreground="#00BFFF"
                       FontSize="15"
                       FontWeight="Bold"
                       FontFamily="Consolas"
                       HorizontalAlignment="Center"
                       Margin="0,0,0,1"/>
            <TextBlock Text="SysOp Navigator"
                       Foreground="#00BFFF"
                       FontSize="13"
                       FontWeight="Bold"
                       FontFamily="Consolas"
                       HorizontalAlignment="Center"
                       Margin="0,0,0,14"/>
            <ProgressBar Name="SplashBar"
                         Height="6"
                         Width="380"
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

$reader            = [System.Xml.XmlNodeReader]::new([xml]$xaml)
$script:window     = [Windows.Markup.XamlReader]::Load($reader)
$script:bar        = $script:window.FindName("SplashBar")
$script:startTime  = [DateTime]::Now
$script:durationMs = $DurationMs

# Load the PNG into the Image control if the file is available
$imgCtrl = $script:window.FindName("SplashImage")
if (Test-Path $script:localImage) {
    try {
        $bmp = [System.Windows.Media.Imaging.BitmapImage]::new()
        $bmp.BeginInit()
        $bmp.UriSource    = [Uri]::new($script:localImage)
        $bmp.CacheOption  = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
        $bmp.EndInit()
        $imgCtrl.Source   = $bmp
    } catch {}
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
