
#not used anymore, but good for reference.
Function Get-ExtensionAttribute {
    [CmdletBinding()]
    Param (
        [Parameter(ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true,
            Position = 0)]
        [string[]]
        $FullName
    )
    DynamicParam {
        $Attributes = New-Object System.Management.Automation.ParameterAttribute
        $Attributes.ParameterSetName = "__AllParameterSets"
        $Attributes.Mandatory = $false
        $AttributeCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
        $AttributeCollection.Add($Attributes)
        $Values = @($Com = (New-Object -ComObject Shell.Application).NameSpace('C:\'); 1..400 | ForEach-Object { $com.GetDetailsOf($com.Items, $_) } | Where-Object { $_ } | ForEach-Object { $_ -replace '\s' })
        $AttributeValues = New-Object System.Management.Automation.ValidateSetAttribute($Values)
        $AttributeCollection.Add($AttributeValues)
        $DynParam1 = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("ExtensionAttribute", [string[]], $AttributeCollection)
        $ParamDictionary = New-Object -Type System.Management.Automation.RuntimeDefinedParameterDictionary
        $ParamDictionary.Add("ExtensionAttribute", $DynParam1)
        $ParamDictionary
    }
 
    begin {
        $ShellObject = New-Object -ComObject Shell.Application
        $DefaultName = $ShellObject.NameSpace('C:\')
        $ExtList = 0..400 | ForEach-Object {
            ($DefaultName.GetDetailsOf($DefaultName.Items, $_)).ToUpper().Replace(' ', '')
        }
    }
 
    process {
        foreach ($Object in $FullName) {
            Write-host 'Get attribs of ' $Object
            # Check if there is a fullname attribute, in case pipeline from Get-ChildItem is used
            if ($Object.FullName) {
                $Object = $Object.FullName
            }
 
            # Check if the path is a single file or a folder
            if (-not (Test-Path -Path $Object -PathType Container)) {
                $CurrentNameSpace = $ShellObject.NameSpace($(Split-Path -Path $Object))
                $CurrentNameSpace.Items() | Where-Object {
                    $_.Path -eq $Object
                } | ForEach-Object {
                    $HashProperties = @{
                        FullName = $_.Path
                    }
                    foreach ($Attribute in $MyInvocation.BoundParameters.ExtensionAttribute) {
                        $HashProperties.$($Attribute) = $CurrentNameSpace.GetDetailsOf($_, $($ExtList.IndexOf($Attribute.ToUpper())))
                    }
                    New-Object -TypeName PSCustomObject -Property $HashProperties
                }
            }
            elseif (-not $input) {
                $CurrentNameSpace = $ShellObject.NameSpace($Object)
                $CurrentNameSpace.Items() | ForEach-Object {
                    $HashProperties = @{
                        FullName = $_.Path
                    }
                    foreach ($Attribute in $MyInvocation.BoundParameters.ExtensionAttribute) {
                        $HashProperties.$($Attribute) = $CurrentNameSpace.GetDetailsOf($_, $($ExtList.IndexOf($Attribute.ToUpper())))
                    }
                    New-Object -TypeName PSCustomObject -Property $HashProperties
                }
            }
        }
    }
 
    end {
        Remove-Variable -Force -Name DefaultName
        Remove-Variable -Force -Name CurrentNameSpace
        Remove-Variable -Force -Name ShellObject
    }
} 