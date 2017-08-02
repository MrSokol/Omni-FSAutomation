<#
.SYNOPSIS
Усановка прав на файловые каталоги по обращениям из OWS
.DESCRIPTION
Усановка прав на файловые каталоги по обращениям из OWS
.PARAMETER resourceName
Имя ресурса
.PARAMETER resourceType
Режим доступа
.PARAMETER resourcePath
Путь к ресурсу
.PARAMETER EmployeeIDs
Список табельных номеров пользователей
.EXAMPLE
PS C:\> Set-OWSFolderPermissions -resourcePath $Resource.Path -resourceName $Resource.Name  -resourceType $Resource.Type -EmployeeIDs $SC.Employees.ID
Получение списка обращений по доступу к файловым каталогам
.NOTES
NAME        :  Set-OWSFolderPermissions
VERSION     :  1.0.0   
LAST UPDATED:  05.10.2017
AUTHOR      :  Denis Kukavitsa
.INPUTS
None
.OUTPUTS
Array of Object
#> 

function Set-OWSFolderPermissions
{
    param (
        [Parameter (Mandatory = $true)] #Имя ресурса, указанное пользователем. Используется если нет пути
            [string] $resourceName,
        [Parameter (Mandatory = $true)] #Режим доступа ресурса
            [string] $resourceType,
        [Parameter (Mandatory = $true)] #Путь к ресурсу
            [string] $resourcePath,
        [Parameter (Mandatory = $true)] #Список табельных номеров пользователей
            [array] $EmployeeIDs           
    )
    #Включение отладки
    #$DebugPreference = "Continue"

    #Объявление переменных
    [string]$SearchBase = ""
    [array]$permittedOU = @("barn","brat","irk","kuz","krsn","nsk","omsk","tmsk","ulan","chi")
    [string]$attributeFolderPathname = ""
    [string]$attributeMode = ""
    [int]$resourceMode = -1
    [bool]$flagRemove = $false


    #Проверяем чей это ресурс (наш или ГО)
    $checkServer = [regex]"(?i)\\\\(\w{3,4})-(fs|nas).*"

    #Если ресурса нет в списке пытаемся угадать что хотел пользователь
    if ($resourcePath -eq "! Файловый ресурс отсутствует в списке") {
        $ou = ""
        $resourcePath = ""
        $resourceName.split("|").trim() | %{
            if ($ou -notin $permittedOU) {
                $regex = $checkServer.Matches($_)
                if ($regex.Count -gt 0) {
                    $ou = $regex.Groups[1].Value  #Корневая OU
                    $server = $regex.Groups[2].Value  #тип сервера
                    $resourcePath = $_
                }
            }
        }
    } else {
        $regex = $checkServer.Matches($resourcePath)
        if ($regex.Count -gt 0) {
            if ($resourcePath.Substring($resourcePath.Length - 1) -eq ' ') {
                throw [System.IO.PathTooLongException] "Лишние знаки в конце ресурса <$resourceName>"
            }
            $ou = $regex.Groups[1].Value  #Корневая OU
            $server = $regex.Groups[2].Value  #тип сервера
        }
    }
    if (($resourcePath -eq "") -or ($ou -notin $permittedOU)) {
        throw [System.Exception] "Ошибка анализа ресурса $resourceName"
    }

    #Используем разные места хранения атрибутов для разных серверов
    Switch ($server) {
        "fs" {
            $SearchBase = "OU=_FS,OU=Groups,OU=$ou,DC=reg,DC=vtb24,DC=ru"
            $attributeFolderPathname = "folderPathname"
            $attributeMode = "Flags"
        }
        "nas" {
            $SearchBase = "OU=$ou,OU=NETAPP,OU=GROUPS,DC=reg,DC=vtb24,DC=ru"
            $attributeFolderPathname = "extensionAttribute14"
            $attributeMode = "extensionAttribute15"
        }
    }
    Switch ($resourceType) {
        "Полный доступ" {
            $resourceMode = 7
            $flagRemove = $false
        }
        "Чтение" {
            $resourceMode = 5
            $flagRemove = $false
        }
        "Отменить права доступа" {
            $resourceMode = 0
            $flagRemove = $true
        }
    }

    if (($SearchBase -eq "") -or ($resourceMode -eq -1)) {
        throw [System.ApplicationException] "Неизвестная ошибка. Проверьте параметры."
    }

    $users=@()   #Список пользователей из AD
    Write-Debug "Получаем список пользователей"
    foreach ($EmployeeID in $EmployeeIDs) {
        $users += Get-ADUser -Filter "EmployeeID -eq $EmployeeID" -Properties mail
        $users += Get-ADUser -Filter "EmployeeID -eq $EmployeeID" -Properties mail -Server "msk.vtb24.ru"
    }

    $groups = @() #Cписок групп из AD
    Write-Debug "Получаем список групп"
    if ($flagRemove -eq $false) {
        $LDAPFilter = "(&(objectCategory=group)(objectClass=group)($attributeFolderPathname=$($resourcePath.Replace('\','\5c').Replace(' ','\20')))($attributeMode=$resourceMode))" 
    } else {
        $LDAPFilter = "(&(objectCategory=group)(objectClass=group)($attributeFolderPathname=$($resourcePath.Replace('\','\5c').Replace(' ','\20'))))" 
    }
    $groups = @(Get-ADGroup -LDAPFilter $LDAPFilter -SearchBase $SearchBase)
    #Костыль для переноса служебной информации из Description (наследние ручного метода обработки)
    if ($groups.Count -eq 0) {
        $access = ""
        Switch ($resourceMode) {
            7 {
                $access = "F"
            }
            5 {
                $access = "R"
            }
        }
        $description = "$resourcePath | $access"
        $LDAPfilter = "(&(objectCategory=group)(objectClass=group)(description="+$description.Replace('\','\5c').Replace(' ','\20')+"))"
        $groups = @(Get-ADGroup -LDAPFilter "$LDAPfilter" -SearchBase $SearchBase -Properties $attributeFolderPathname,  $attributeMode)
        Write-Debug "КОСТЫЛЬ: $description"
        Write-Debug "КОСТЫЛЬ: $LDAPfilter"
        if ($groups.Count -ne 0) {
            foreach ($group in $groups) {
                Set-ADGroup $group.Name -Replace @{
                    "$attributeFolderPathname" = $resourcePath
                }
                Set-ADGroup $group.Name -Replace @{
                    "$attributeMode" = $resourceMode
                }
                Write-Debug "КОСТЫЛЬ: Сработал на группе $group"
            }
        }
    }
    #Конец костыля
    if (($groups.Count -ne 0) -and ($users.Count -ne 0)) {
        if ($flagRemove -eq $false) {
            Write-Debug "Добавляю в группы"
            $groups | add-ADGroupmember -Members $users
        } else {
            Write-Debug "Удаляю из групп"
            $groups | Remove-ADGroupMember -Members $users
        }
    } elseif ($groups.Count -eq 0) {
        throw [System.DirectoryServices.ActiveDirectory.ActiveDirectoryObjectNotFoundException] "Группа для каталога <$resourcePath> не найдена."
    }
}#Set-OWSFolderPermissions
