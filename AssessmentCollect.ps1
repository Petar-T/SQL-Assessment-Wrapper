function checkPrereqs{
param ([Parameter(Mandatory = $False)]
        [object[]] $ModuleList=(('PSWriteWord','1.1.14'),('ImportExcel','7.8.4'),('SqlServer','22.3.0'))
        )
    foreach ($mdl in $ModuleList)
    {
        if (Get-InstalledModule -Name  $mdl[0] -MinimumVersion $mdl[1]  ) {
            Write-Host "Module [$($mdl[0])] exists in required version"    } 
        else {
            Write-Host "Module [$($mdl[0])] does not exist in minimal version [$($mdl[1])] , you need to install it in order to properly run this script" -ForegroundColor Red
            Write-Host "You should : " -ForegroundColor Yellow
            Write-Host "    Install-Module -Name $($mdl[0]) -MinumumVersion $($mdl[1]) " -ForegroundColor Yellow
            Write-Host "    manually download file from [https://www.powershellgallery.com/packages/$($mdl[0])/] " -ForegroundColor Yellow
            Exit}
    }
}

function parseObj{
param( 
 [Parameter(Mandatory = $true)]
        [String]
        $ParseLine,
 [Parameter(Mandatory = $true)]
        [String]
        $ParseType)
        

 if ($ParseType -eq 'Database'){
    $ret =    $ParseLine.Substring($ParseLine.IndexOf('Database[@Name=')+16, $ParseLine.Length - ($ParseLine.IndexOf('Database[@Name=')+18) ) }
else {
    
    $ret =    $ParseLine.Substring($ParseLine.IndexOf('@Name=')+7 , $ParseLine.Length - ($ParseLine.IndexOf('@Name=')+9) )    }
     
     
     return $ret    
}

function colorSeverity{
param ([Parameter(Mandatory = $True)]
        [String] $SeverityType)
$ret='Black'
switch -Regex ($SeverityType) {
    'Critical' {$ret='Red'}
    'High'     {$ret='Red'}
    'Medium'   {$ret='Blue'}
    'Low'      {$ret='Green'}
    'Information' {$ret='Yellow'}
    Default   {$ret='Black'}
}
return $ret
}

function  collect-Assessment{
param( 
 [Parameter(Mandatory = $true)]
        [String[]] $ServerList,
 [Parameter(Mandatory = $true)]
        [String] $OutputPath,
 [Parameter(Mandatory = $false)]
        [switch] $doNotCreateWordOutput
        )

   CheckPrereqs 
   $sortTargetType = "High", "Medium", "Low", "Information"
   $sortSeverity = "Server", "Database"

    foreach($srv in $ServerList)
    {
        try
        {
            
            Write-Host "[$srv] connecting.. " -ForegroundColor Green

            Get-SqlInstance -ServerInstance $srv | Invoke-SqlAssessment -FlattenOutput -ErrorAction Ignore -OutVariable ResultSet  | Out-Null
            Get-SqlDatabase -ServerInstance $srv | Where-Object -Property Status -eq ‘Normal’| Invoke-SqlAssessment -FlattenOutput -ErrorAction Ignore -OutVariable +ResultSet    | Out-Null

            $csvPath = "$OutputPath\assesment_$srv.csv"
            $docPath = "$OutputPath\assesment_findings_$srv.docx"
            #$ResultSet | Sort TargetType, Severity |  Export-Csv -Path $csvPath -NoTypeInformation
            $ResultSet | Sort-Object { $sortTargetType.IndexOf($_.TargetType), $sortSeverity.IndexOf($_.Severity) } |  Export-Csv -Path $csvPath -NoTypeInformation

            Write-Host "[$srv] wrote file $csvPath" -ForegroundColor Green
            if (-not ($doNotCreateWordOutput.IsPresent ))
            {
                $WordDocument = New-WordDocument $docPath
                $Footer = Add-WordFooter -WordDocument $WordDocument -DifferentFirstPage $true -DifferentOddAndEvenPages $false
                $Header = Add-WordHeader -WordDocument $WordDocument

                foreach ($row in $ResultSet)
                {
                    $Row1 = New-Object -TypeName PSObject
                        $Row1 | Add-Member -Name 'Name'  -MemberType Noteproperty -Value 'Check Name'
                        $Row1 | Add-Member -Name 'Value' -MemberType Noteproperty -Value $row.CheckName
                    $Row2 = New-Object System.Object
                        $Row2 | Add-Member -Name 'Name'  -MemberType Noteproperty -Value 'Severity'
                        $Row2 | Add-Member -Name 'Value' -MemberType Noteproperty -Value $row.Severity
                    $Row3 = New-Object System.Object
                        $Row3 | Add-Member -Name 'Name'  -MemberType Noteproperty -Value 'Check Level'
                        $Row3 | Add-Member -Name 'Value' -MemberType Noteproperty -Value $row.TargetType
                    $Row4 = New-Object System.Object
                        $aff = ParseObj -ParseLine $row.TargetPath -ParseType $row.TargetType
                        $Row4 | Add-Member -Name 'Name'  -MemberType Noteproperty -Value 'Affected Object'
                        $Row4 | Add-Member -Name 'Value' -MemberType Noteproperty -Value $aff
                    $Row5 = New-Object System.Object
                        $Row5 | Add-Member -Name 'Name'  -MemberType Noteproperty -Value 'Message'
                        $Row5 | Add-Member -Name 'Value' -MemberType Noteproperty -Value $row.Message
                    $Row6 = New-Object System.Object
                        $Row6 | Add-Member -Name 'Name'  -MemberType Noteproperty -Value 'Help Link'
                        $Row6 | Add-Member -Name 'Value' -MemberType Noteproperty -Value  $row.HelpLink
                        #Add-WordHyperLink -WordDocument ??

                    $TableRows = @($Row1, $Row2, $Row3, $Row4, $Row5, $Row6)
                    
                    $clr = ColorSeverity -SeverityType $row.Severity
                    Add-WordTable -WordDocument $WordDocument -DataTable $TableRows -Design MediumShading1Accent1   -Percentage $true -ColumnWidth 25,75   -Supress $True -Color Black, Black, $clr  , Black, black , black -Bold $True, $True , $false, $false, $false, $false
                    Add-WordParagraph -WordDocument $WordDocument -Supress $True
                }
                            
                Save-WordDocument $WordDocument -Language 'en-US' -Supress $True -OpenDocument
                Write-Host "[$srv] wrote file $docPath" -ForegroundColor Green
            }
        }
        catch [Exception]
        {
                Write-Warning "Problem using [$srv] , Error message: $_ " 
        }
        Write-Host ""
    } 
    
    Write-Host "assessment completed !"
}

#Example
#collect-Assessment  -ServerList 'ServerA'  -OutputPath 'c:\Temp' 
#collect-Assessment  -ServerList 'ServerA','ServerB' -OutputPath 'c:\Temp'  -doNotCreateWordOutput  #just produce Excel file 


