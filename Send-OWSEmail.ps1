<#
.SYNOPSIS
Send an email with an SC data
.DESCRIPTION
Send email
.PARAMETER Resources
PSOBJECT of Resources
.PARAMETER Employees
PSOBJECT of Employees
.PARAMETER Subject
The Subject of the email
.PARAMETER To
The To field is who receives the email
.EXAMPLE
PS C:\> Send-OWSEmail -Resources @Resources -Employees @Employees -Subject $Subject -Pre $Pre -To sokol@nsk.vtb24.ru
An example to send some SC information to email recipient
.NOTES
NAME        :  Send-OWSEmail
VERSION     :  1.0.0   
LAST UPDATED:  05.10.2017
AUTHOR      :  Denis Kukavitsa
.INPUTS
None
.OUTPUTS
None
#> 

function Send-OWSEmail {
#Requires -Version 2.0
[CmdletBinding()]
 Param 
   ([Parameter(Mandatory=$True,
               Position = 0,
               HelpMessage="Please enter the Resources")]
    $Resources,
    [Parameter(Mandatory=$True,
               Position = 1,
               HelpMessage="Please enter the Employees")]
    $Employees,
    [Parameter(Mandatory=$True,
               Position = 2,
               HelpMessage="Please enter the Subject")]
    [String]$Subject,
    [Parameter(Mandatory=$False,
               Position = 3)]
    [String]$Pre,
    [Parameter(Mandatory=$False,
               Position = 4,
               HelpMessage="Please enter the To address")]    
    [String[]]$To = "sysadmin@nsk.vtb24.ru",
    [String]$From = "»ннокентий <kesha@nsk.vtb24.ru>",    
    [String]$CSS,
    [String]$SmtpServer ="nsk-mbx.reg.vtb24.ru"
   )#End Param

$CSS = @"
<style type="text/css">
    table {
        font-family: Verdana;
        border-style: dashed;
        border-width: 1px;
        border-color: #FF6600;
        padding: 5px;
        background-color: #FFFFCC;
        table-layout: auto;
        text-align: center;
        font-size: 8pt;
        width: 100%;
    }

    table th {
        border-bottom-style: solid;
        border-bottom-width: 1px;
        font: bold
    }
    table td {
        border-top-style: solid;
        border-top-width: 1px;
    }
    .style1 {
        font-family: Courier New, Courier, monospace;
        font-weight:bold;
        font-size:small;
    }
</style>
"@

$Employees = "$($Employees | ConvertTo-Html -Fragment -Pre "Список пользователей:`n")"
$Resources = "$($Resources | ConvertTo-Html -Fragment -Pre "Список ресурсов:`n")"

$xmlOut = [xml] (ConvertTo-Html -Title $Subject -Head $CSS -Body "$Pre<br/>`n$Employees<br/>`n$Resources<br/>`n")
$null = $xmlOut.html.body.LastChild.ParentNode.RemoveChild($xmlOut.html.body.LastChild)

$Splat = @{
    To         =$To
    Body       ="$($xmlOut.OuterXml)"
    Subject    =$Subject
    SmtpServer =$SmtpServer
    From       =$From
    BodyAsHtml =$True
    }
    Send-MailMessage @Splat -Encoding utf8
    
}#Send-OWSEmail
