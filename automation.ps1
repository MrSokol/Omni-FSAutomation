#Включение отладки
#$DebugPreference = "Continue"

. .\Invoke-OWSRequest.ps1
. .\Set-OWSFolderPermissions.ps1
. .\Send-OWSEmail.ps1
$OWSRequest = Invoke-OWSRequest -SCType open_newfc -Type View
if ($OWSRequest -eq $false) {
    return
}
foreach ($SC in $OWSRequest) {
    $SC.ResourceName = $SC.ResourceName.split(";")
    $SC.ResourceType = $SC.ResourceType.split(";")
    $SC.ResourceApproval = $SC.ResourceApproval.split(";")
    $SC.ResourcePath = $SC.ResourcePath.split(";")
    $SC | Add-Member -type NoteProperty -name "Resources" -Value @()
    Write-Debug "КОСТЫЛЬ: Ресурсов в обращении SC-$($SC.ID) - $($SC.ResourceName.Count)"
    for ($i = 0; $i -lt $SC.ResourceName.Count; $i++) {
        $Resource = New-Object PSObject
        $Resource | Add-Member -type NoteProperty -name "Name" -Value $SC.ResourceName[$i]
        $Resource | Add-Member -type NoteProperty -name "Path" -Value $SC.ResourcePath[$i]
        $Resource | Add-Member -type NoteProperty -name "Type" -Value $SC.ResourceType[$i]
        $Resource | Add-Member -type NoteProperty -name "Approval" -Value $SC.ResourceApproval[$i]
        $Resource | Add-Member -type NoteProperty -name "Result" -Value $null
        $SC.Resources += $Resource
        Write-Debug "КОСТЫЛЬ: Добавлен ресурс $($Resource.Path)"
        remove-variable Resource
    }
    $SC.PSObject.Properties.Remove("ResourceName")
    $SC.PSObject.Properties.Remove("ResourceType")
    $SC.PSObject.Properties.Remove("ResourceApproval")
    $SC.PSObject.Properties.Remove("ResourcePath")
    $SC | Add-Member -type NoteProperty -name "Employees" -Value @()
    if (($SC.Users -eq $null) -and ($SC.UserIDs -eq $null)) {
            $SC.Users = $SC.Name
            $SC.UserIDs = $SC.EmployeeID
    }
    $SC.Users = $SC.Users.split(";")
    $SC.UserIDs = $SC.UserIDs.split(";")
    Write-Debug "КОСТЫЛЬ: Пользователей в обращении SC-$($SC.ID) - $($SC.UserIDs.Count)"
    for ($i = 0; $i -lt $SC.UserIDs.Count; $i++) {
        $Employees = New-Object PSObject
        $Employees | Add-Member -type NoteProperty -name "ID" -Value $SC.UserIDs[$i]
        $Employees | Add-Member -type NoteProperty -name "Name" -Value $SC.Users[$i]
        #$Employees | Add-Member -type NoteProperty -name "Result" -Value $null
        $SC.Employees += $Employees
        Write-Debug "КОСТЫЛЬ: Добавлен пользователь $($Employees.Name)"
        remove-variable Employees
    }
    $SC.PSObject.Properties.Remove("Users")
    $SC.PSObject.Properties.Remove("UserIDs")
    #---------------------------------------------------------
    # Решение обращения

    $resultText = "Результат решения обращения SC-$($SC.ID)<br/>"
    $resultText += "Заявитель: $($SC.ownerName)<br/>"

    $status = ""
    $SC.Result = $true
    foreach ($Resource in $SC.Resources) {
        try {
            Set-OWSFolderPermissions -resourcePath $Resource.Path -resourceName $Resource.Name  -resourceType $Resource.Type -EmployeeIDs $SC.Employees.ID
            $Resource.Result = "Выполнено"
            $status = "Решено"
        }
        catch [System.DirectoryServices.ActiveDirectory.ActiveDirectoryObjectNotFoundException] {
            $Resource.Result = "$($_.Exception.Message)"
            $status = "Проблема при решении"
            $SC.Result = $false
        }
        catch [System.Exception] {
            $Resource.Result = "$($_.Exception.Message)"
            $status = "Ошибка при решении"
            $SC.Result = $false
        }
        catch {
            $Resource.Result = "$($_.Exception.Message)"
            $status = "Ошибка при решении"
            $SC.Result = $false
        }
    }
    if ($SC.Result -eq $true) {
        if ((Invoke-OWSRequest -SCType open_newfc -Type CloseSC -SC $SC.ID) -eq $true) {
            $status = "Закрыто"
        } else {
            $status = "Ошибка при закрытии"
        }
    }    
    Send-OWSEmail -Resources ($SC.Resources) -Employees ($SC.Employees) -Subject "$status SC-$($SC.ID)" -Pre $resultText
}
