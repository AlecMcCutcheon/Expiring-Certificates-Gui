<# This form was created by Alec McCutcheon #>

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = New-Object System.Drawing.Point(730,400)
$Form.text                       = "Search for Expiring certifications"
$Form.TopMost                    = $false

$TextBox1                        = New-Object system.Windows.Forms.TextBox
$TextBox1.multiline              = $false
$TextBox1.width                  = 100
$TextBox1.height                 = 20
$TextBox1.location               = New-Object System.Drawing.Point(200,11)
$TextBox1.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Button1                         = New-Object system.Windows.Forms.Button
$Button1.text                    = "Search"
$Button1.width                   = 102
$Button1.height                  = 25
$Button1.location                = New-Object System.Drawing.Point(14,46)
$Button1.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$TextBox2                        = New-Object system.Windows.Forms.TextBox
$TextBox2.multiline              = $true
$TextBox2.width                  = 700
$TextBox2.height                 = 304
$TextBox2.location               = New-Object System.Drawing.Point(14,76)
$TextBox2.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Label1                          = New-Object system.Windows.Forms.Label
$Label1.text                     = "Find certificates expiring within"
$Label1.AutoSize                 = $true
$Label1.width                    = 25
$Label1.height                   = 10
$Label1.location                 = New-Object System.Drawing.Point(14,16)
$Label1.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Label2                          = New-Object system.Windows.Forms.Label
$Label2.text                     = "days."
$Label2.AutoSize                 = $true
$Label2.width                    = 25
$Label2.height                   = 10
$Label2.location                 = New-Object System.Drawing.Point(306,16)
$Label2.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$CheckBox1                       = New-Object system.Windows.Forms.CheckBox
$CheckBox1.text                  = "Ignore temporary Certifications"
$CheckBox1.AutoSize              = $true
$CheckBox1.width                 = 201
$CheckBox1.height                = 28
$CheckBox1.Anchor                = 'top'
$CheckBox1.location              = New-Object System.Drawing.Point(350,17)
$CheckBox1.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Form.controls.AddRange(@($TextBox1,$Button1,$TextBox2,$Label1,$Label2,$CheckBox1))

$Button1.Add_MouseClick({ GetCertifications $this $_ })

$checkbox1.Checked = $True

function GetCertifications ($sender,$event) { 

    $DayCount = $TextBox1.text
    
    if ($checkbox1.Checked){
$certs = Get-ChildItem -Path cert:\LocalMachine\My -Recurse | where { $_.notafter -le (get-date).AddDays($DayCount) -AND $_.notafter -gt (get-date) -AND $_.issuer -NotLike "*CN=MS-Organization-P2P-Access*"} | select thumbprint, subjectname, issuer
} else{
$certs = Get-ChildItem -Path cert:\LocalMachine\My -Recurse | where { $_.notafter -le (get-date).AddDays($DayCount) -AND $_.notafter -gt (get-date)} | select thumbprint, subjectname, issuer
}

IF([string]::IsNullOrEmpty($certs)) {            
    $TextBox2.text = "No certificates expiring in within that time frame"            
} else {
    $TextBox2.text = ("Certificates Expiring Within $DayCount Days:" + [Environment]::NewLine + [Environment]::NewLine) 
    foreach($cert in $certs){
    foreach($prop in $cert.psobject.properties){
        $TextBox2.text += ("$($prop.name) : $($prop.Value)" + [Environment]::NewLine) 
      }
    $TextBox2.text += ([Environment]::NewLine)
    }
  }         
}
[void]$Form.ShowDialog()