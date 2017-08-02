<#
.SYNOPSIS
API для OWS
.DESCRIPTION
Загрузка списка обращений по типу. Закрытие обращения по номеру
.PARAMETER SCType
Тип обращения.
.PARAMETER Type
Тип действия (выгрузка, закрытие)
.PARAMETER SC
Номер обращения для закрытия
.EXAMPLE
PS C:\> Invoke-OWSRequest -SCType open_newfc -Type View
Получение списка обращений по доступу к файловым каталогам
.NOTES
NAME        :  Invoke-OWSRequest
VERSION     :  1.0.0   
LAST UPDATED:  05.10.2017
AUTHOR      :  Denis Kukavitsa
.INPUTS
None
.OUTPUTS
Array of Object
#> 

function Invoke-OWSRequest
{
    param (
        [Parameter (Mandatory = $true)]
        [ValidateSet("open_newworker", "open_newfc", "open_newmail")]
            [string] $SCType,
        [Parameter (Mandatory = $true)]
        [ValidateSet("View", "CloseSC", "UpdateSC")]
            [string] $Type,
        [Parameter (Mandatory = $False)]
            [string] $SC,   
        [Parameter (Mandatory = $False)]
            [string] $toGroup,  
        [Parameter (Mandatory = $False)]
            [string] $Comment                  
    )
    
    Switch ($Type) {
        "View" {$Method = "GET"}
        "CloseSC" {$Method = "POST"}
        "UpdateSC" {$Method = "POST"}
    }

    $serverName = "omni-app-3"

    $Uri = "http://$serverName/ows/json/nova/$Type"

    Switch ($Type) {
        "View" {
            $result = Invoke-WebRequest -Uri $Uri -UseDefaultCredentials -Method $Method -Body @{type=$SCType}
        }
        "CloseSC" {
            $postParams = @{
                ID=$SC;
                Solution="Доступ изменён. Для применения изменений необходимо не раньше чем через 30 минут перезайти в компьютер.
При возникновении вопросов в рамках данного обращения можете написать на адрес '5440 группа системного администрирования'";
                Class="СА Операции с файловым каталогом (создание/доступ…)"
            }
            $result = Invoke-WebRequest -Uri $Uri -UseDefaultCredentials -Method $Method -Body $postParams
        }
        "UpdateSC" {
            $postParams = @{
                ID=$SC;
                Comment=$Comment;
                Group=$toGroup
            }
            $result = Invoke-WebRequest -Uri $Uri -UseDefaultCredentials -Method $Method -Body $postParams
        }
    }
    
    $result = $result.Content | ConvertFrom-Json

    if ($result.status -eq 0) {
        if ($result.data.count -gt 0) {
            $data = @()
            foreach ($SCTmp in $result.data) 
            {
                $i = 0
                $tmp = New-Object PSObject
                $result.scheme | 
                ForEach{ 
                    if ($SCTmp[$i] -eq "Пусто") {
                        $SCTmp[$i] = $null
                    }
                    $tmp | Add-Member -type NoteProperty -name $_ -Value $SCTmp[$i]
                    $i++
                }
                $tmp | Add-Member -type NoteProperty -name "SCType" -Value $SCType
                $tmp | Add-Member -type NoteProperty -name "Result" -Value $False
                $data += $tmp
                remove-variable tmp
            }
            return $data
        } elseif ($Type -eq "CloseSC") {
            return $true
        } elseif ($Type -eq "UpdateSC") {
            return $true            
        } else {
            return $false
        }
    }
    return $false
}#Invoke-OWSRequest
