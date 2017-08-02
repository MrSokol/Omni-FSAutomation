. .\Invoke-OWSRequest.ps1
$ConfirmPreference="none"
$session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://nsk-mbx.reg.vtb24.ru/powershell -AllowRedirection
Import-PSSession $session -AllowClobber

$global:PSEmailServer = "nsk-mail.reg.vtb24.ru"
$global:smtp=@{	
	Encoding = [System.Text.Encoding]::UTF8
	From  = "Иннокентий <kesha@nsk.vtb24.ru>"
}
$global:to = "sysadmin@nsk.vtb24.ru"
#$global:to = "sokol@nsk.vtb24.ru"

$OfficeXML="\\nsk-fs\Otdels\ОИТ\_SysAdmin\_scripts\data\Office.xml"
$Offices = [xml](Get-Content $OfficeXML)  

$getParams = @{type="open_newworker"}
$result = Invoke-WebRequest -Uri "http://omni-app-3/OWS/json/nova/View" -UseDefaultCredentials -Method Get -Body $getParams
$result = $result.Content | ConvertFrom-Json
Write-Host $result
if ($result.status -eq 0) {
    if ($result.data.count -gt 0) {
        $data = @()
        foreach ($employee in $result.data) 
        {
            $i = 0
            $tmp = New-Object PSObject
            $result.scheme | 
            ForEach{ 
                $tmp | Add-Member -type NoteProperty -name $_ -Value $employee[$i]
                $i++
            }
            $tmp | Add-Member -type NoteProperty -name "Result" -Value $False
            $tmp | Add-Member -type NoteProperty -name "SPP" -Value $False
            $data += $tmp
            remove-variable tmp
        }
    } else {
        Write-Host "Нет новых обращений!"
        Remove-PSSession $Session
        return
    }
}
foreach ($employee in $data) {
    $user=@()
    $user=Get-ADUser -Filter "(EmployeeID -eq $($employee.EmployeeID)) -and (Enabled -eq 'true')" -Properties mail
    #if ($user.count -eq 0) {
    #    $user=Get-ADUser -Filter "(EmployeeID -eq $($employee.EmployeeID)) -and (Enabled -eq 'true')" -Properties mail -Server msk.vtb24.ru
    #} 
    if ($user.count -eq 0) {
        $body = "Пользователь <$($employee.Name)> с табельным номером: <$($employee.EmployeeID)> не найден!"
        Send-MailMessage @smtp -To $to -Subject ("Проблема при решении обращения  SC-$($employee.ID)") -body $body -BodyAsHtml
        write-Warning $body
    }
    elseif ($user.count -gt 1) {
        $body = "ДАННЫЕ ПОЛЬЗОВАТЕЛИ <$($employee.Name)> ИМЕЮТ ОДИНАКОВЫЙ EmployeeID"
        Send-MailMessage @smtp -To $to -Subject ("Проблема при решении обращения  SC-$($employee.ID)") -body $body -BodyAsHtml
        write-Warning $body   
    } else {
        Write-Host "Правим пользователя <$($user.Name)> по обращению SC-$($employee.ID)"
        $user = $user | %{$_ | select *,@{n='OU';e={$($_.distinguishedname  -split ",")[2] -replace "OU=",''}},@{n='code';e={$($employee.code)}}}

        $reg = $Offices.data.Region | where{$_.OU -eq $user.OU}
        if ($reg.count -ne 0) {
            Write-Host "Добавляем стандартные группы для региона <$($reg.Name)>"
            $reg.Groups | where{$_} | foreach {add-adgroupmember $_ -Members $user.SamAccountName}

            $tp = $reg.TP | where{$_.COD -eq $user.code}
            if ($tp.count -ne 0) {
                Write-Host "Добавляем стандартные группы для ТП <$($tp.Name)>"
                $tp.Groups | where{$_} | foreach {add-adgroupmember $_ -Members $User.SamAccountName}

                $otdel=$tp.Otdel | where{$_.Name -eq $employee.Office}
                if ($otdel.count -ne 0) {
                    Write-Host "Добавляем стандартные группы для отдела <$($otdel.Name)>"
                    $otdel.Groups | where{$_} | foreach {add-adgroupmember $_ -Members $User.SamAccountName}
                } else {
                    Write-Warning "Не найден отдел <$($employee.Office)>"
                }
            } else {
                Write-Warning "Не найдена ТП <$($employee.Office)>"
            }
        } else {
            Write-Warning "Не найден Регион <$($employee.Office)>"
        }

        #Создаем папку на диске U
        $folder = "\\$($reg.OU)-fs\users\$($user.SamAccountName)"
        $f_group = "$"+$reg.OU+"-AdminsFS"
        $r_group = $reg.OU+'_U_R'
        if(!(Test-Path $folder)) { 
            Write-Host "Создаем папку $folder"
            md -Path $folder
        } else {
            Write-Host "Папка <$folder> уже существует"
        }
        $acl = Get-Acl $folder
        $AccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule ("$($user.SamAccountName)","DeleteSubdirectoriesAndFiles, Modify, Synchronize","ContainerInherit, ObjectInherit", "None","Allow")
        $AccessRule1 = New-Object System.Security.AccessControl.FileSystemAccessRule ("Администраторы","FullControl","ContainerInherit, ObjectInherit", "None","Allow")
        $AccessRule2 = New-Object System.Security.AccessControl.FileSystemAccessRule ($f_group,"FullControl","ContainerInherit, ObjectInherit", "None","Allow")
        $AccessRule3 = New-Object System.Security.AccessControl.FileSystemAccessRule ($r_group,"ReadAndExecute, Synchronize","ContainerInherit, ObjectInherit", "None","Allow")
        $acl.SetAccessRule($AccessRule)
        $acl.SetAccessRule($AccessRule1)
        $acl.SetAccessRule($AccessRule2)
        $acl.SetAccessRule($AccessRule3)
        $acl.SetAccessRuleProtection($true,$false)
        Write-Host "Даём права на папку $folder"
        $acl | Set-Acl $folder

        $employee.Result = $true
        $employee.SPP = $reg.SPP
        if ($employee.Result -eq $true) {
            Write-Host "Переназначаем обращение SC-$($employee.ID)"
            Write-Host "Группа назначения <$($employee.SPP)>"
            $postParams = @{
                        ID=$($employee.ID);
                        Group="$($employee.SPP)"
                        Comment="Учетная запись доработана после участия FIM. Логин и пароль автоматически высылается системой FIM на группу технической поддержки вашего региона, но не непосредственному руководителю нового сотрудника. Такие дела...";
                        }
            $result = Invoke-WebRequest -Uri "http://omni-app-4/ows/json/nova/UpdateSC" -UseDefaultCredentials -Method POST -Body $postParams;
            $result = $result.Content | ConvertFrom-Json
            if ($result.status -eq 0) {
                $body = "Пользователю <$($employee.Name)> предоставлены стандартные права и обращение отправлено на группу <$($employee.SPP)>"
                Send-MailMessage @smtp -To $to -Subject ("Переназначено обращение SC-$($employee.ID)") -body $body -BodyAsHtml
            } else {
                $body = "Просьба вручную переназначить обращение на группу <$($employee.SPP)>."
                Send-MailMessage @smtp -To $to -Subject ("Ошибка при переназначении обращения SC-$($employee.ID)") -body $body -BodyAsHtml
            }    
        } else {
                $body = "Просьба просьба вручную обработать обращение."
                Send-MailMessage @smtp -To $to -Subject ("Ошибка при решении обращения SC-$($employee.ID)") -body $body -BodyAsHtml
        }
        Write-Host "========================================================================"
    }
}
Remove-PSSession $Session