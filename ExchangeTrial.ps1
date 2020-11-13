#region Functions
Function Transpose-Object
{ [CmdletBinding()]
  Param([OBJECT][Parameter(ValueFromPipeline = $TRUE)]$InputObject)

  BEGIN
  { # initialize variables just to be "clean"
    $Props = @()
    $PropNames = @()
    $InstanceNames = @()
  }

  PROCESS
  {
  	if ($Props.Length -eq 0)
  	{ # when first object in pipeline arrives retrieve its property names
			$PropNames = $InputObject.PSObject.Properties | Select-Object -ExpandProperty Name
			# and create a PSCustomobject in an array for each property
			$InputObject.PSObject.Properties | %{ $Props += New-Object -TypeName PSObject -Property @{Property = $_.Name} }
		}

 		if ($InputObject.Name)
 		{ # does object have a "Name" property?
 			$Property = $InputObject.Name
 		} else { # no, take object itself as property name
 			$Property = $InputObject | Out-String
		}

 		if ($InstanceNames -contains $Property)
 		{ # does multiple occurence of name exist?
  		$COUNTER = 0
 			do { # yes, append a number in brackets to name
 				$COUNTER++
 				$Property = "$($InputObject.Name) ({0})" -f $COUNTER
 			} while ($InstanceNames -contains $Property)
 		}
 		# add current name to name list for next name check
 		$InstanceNames += $Property

  	# retrieve property values and add them to the property's PSCustomobject
  	$COUNTER = 0
  	$PropNames | %{
  		if ($InputObject.($_))
  		{ # property exists for current object
  			$Props[$COUNTER] | Add-Member -Name $Property -Type NoteProperty -Value $InputObject.($_)
  		} else { # property does not exist for current object, add $NULL value
  			$Props[$COUNTER] | Add-Member -Name $Property -Type NoteProperty -Value $NULL
  		}
 			$COUNTER++
  	}
  }

  END
  {
  	# return collection of PSCustomobjects with property values
  	$Props
  }
}
Function ConvertTo-MDTest {

    [cmdletbinding()]
        [outputtype([string[]])]
        [alias('ctm')]
    
        Param(
            [Parameter(Position = 0, ValueFromPipeline)]
            [object]$Inputobject,
            [Parameter()]
            [string]$Title,
            [string[]]$PreContent,
            [string[]]$PostContent,
            [ValidateScript( {$_ -ge 10})]
            [int]$Width = 80,
            #display results as a markdown table
            [switch]$AsTable
        )
    
        Begin {
            Write-Verbose "[BEGIN  ] Starting $($myinvocation.MyCommand)"
            #initialize an array to hold incoming data
            $data = @()
    
            #initialize an empty here string for markdown text
            $Text = @"

"@
            If ($title) {
                Write-Verbose "[BEGIN  ] Adding Title: $Title"
                $Text += "# $Title`n`n"
            }
            If ($precontent) {
                Write-Verbose "[BEGIN  ] Adding Precontent"
                $Text += $precontent
                $text += "`n`n"
            }
    
        } #begin
        Process {
            #add incoming objects to data array
            Write-Verbose "[PROCESS] Adding processed object"
            $data += $Inputobject
    
        } #process
        End {
            #add the data to the text
            if ($data) {
                if ($AsTable) {
                    Write-Verbose "[END    ] Formatting as a table"
                    $names = $data[0].psobject.Properties.name
                    $head = "| $($names -join " | ") |"
                    $text += $head
                    $text += "`n"
    
                    $bars = "| $(($names -replace '.','-') -join " | ") |"
    
                    $text += $bars
                    $text += "`n"
    
                    foreach ($item in $data) {
                        $line = "| "
                        $values = @()
                        for ($i = 0; $i -lt $names.count; $i++) {
                            
                            #if an item value contains return and new line replace them with <br> Issue #97
                            if ($item.($names[$i]) -match "`n") {
                                Write-Verbose "[END    ] Replacing line returns for property $($names[$i])"
                                [string]$val = $($item.($names[$i])).replace("`r`n","<br>") -join ""
                                Write-Verbose $val
                            }
                            else {
                                [string]$val = $item.($names[$i])
                            }
                            
                            $values += $val
                        }
                        $line += $values -join " | "
                        $line += " |"
                        $text += $line
                        $text += "`r"
                    }
                }
                else {
                    #convert data to strings and trim each line
                    Write-Verbose "[END    ] Converting data to strings"
                    [string]$trimmed = (($data | Out-String -Width $width).split("`n")).ForEach({ "$($_.trim())`n" })
                    Write-Verbose "[END    ] Adding to markdown"
                    $clean = $($trimmed.trimend())
                    $text += @"
    ``````text
    $clean
"@
            } #else as text
        } #if $data
        If ($postcontent) {
            Write-Verbose "[END    ] Adding postcontent"
            $text += "`n"
            $text += $postcontent
        }
    
        #write the markdown to the pipeline
        $text.TrimEnd()
        Write-Verbose "[END    ] Ending $($myinvocation.MyCommand)"
    } #end
    }

#endregion Functions

#region Variables
$VerbosePreference = 'Continue'
$UserPrincipalName = 'martin@precisionservices.biz'
$outfile = ".\TransportRules.md"
#endregion Variables

# Start writing document
Set-Content -Encoding UTF8 -Path $outfile -Value "# Overview"
Connect-ExchangeOnline -UserPrincipalName $UserPrincipalName
Add-Content -Path $outfile -Encoding UTF8 -Value "## Exchange Mail Flow Rules"
Add-Content -Path $outfile -Encoding UTF8 -Value "These values are retrieved using powershell 'Get-TransportRule'"
Add-Content -Path $outfile -Encoding UTF8 -Value "These values can be viewed in the GUI by accessing the Exchange Admin Center - Mail Flow - Rules"

$results = Get-TransportRule | Select-Object * -ExcludeProperty RunspaceId, ObjectClass, PSComputerName, PSShowComputerName
Write-Host $results
$allTransportRules = @()
foreach ($result in $results){
    $thisTransportRule = @()
# 	# only show non-null values
    foreach ($noteProperty in $result | Get-Member -MemberType NoteProperty){
        $PSObject = @()
        if ($null -ne $noteProperty.Value){
            $PSObject = New-Object PSObject -Property @{
                Name = $noteProperty.Name
                Value = $noteProperty.Value
            }    
        }
        $thisTransportRule += $PSObject
        # $PSObject = $null
    }
    
}
	$thisTransportRule | Transpose-Object | ConvertTo-MDTest -AsTable | Add-Content -Encoding UTF8 $outfile