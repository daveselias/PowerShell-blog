#* FileName: My-ADTools_v1.psm1
#*=============================================
#* Script Name: Create-ADVisio
#* Created: 07/01/2014-published
#* Author: David S. Elias
#* Email: daveselias@gmail.com
#* Requirements: Read rights in Active Directory, Server Local Admin rights, SQL SysAdmin rights
#* Keywords: AD, SQL, share, groups, nested
#*=============================================

#*=============================================
#* REVISION HISTORY
#*=============================================
#* Date: 07/01/2014
#* Time: 12:30PM
#* Issue: Local Server Share rights missing
#* Solution: Added function Get-SharePermissions
#*           to the "-ServerCheck" switch
#*=============================================

#*=============================================
#* FUNCTION LISTINGS
#*=============================================
#* Function: Get-LocalGroups
#* Created: 06/25/2014
#* Author: David S. Elias
#* Arguments: $ComputerName (Server)
#*=============================================
#* Purpose: Gathers users and groups that are
#*          members of local server groups
#*
#*=============================================
#* Function: Get-SQLUsers
#* Created: 06/26/2014
#* Author: David S. Elias
#* Arguments: $ComputerName (Server) $GroupName (AD Group)
#*=============================================
#* Purpose: Interigates SQL Server security login list
#*          gets server security and all databases
#*          containing the AD Group in it's securities
#*=============================================
#* Function: Get-SharePermissions
#* Created: 07/01/2014
#* Author: David S. Elias
#* Arguments: $ComputerName (Server)
#*=============================================
#* Purpose: Retrieves Share securities (not NTFS)
#*          from all shares on server that have
#*          more than default accounts added
#*=============================================
#* Function: Create-ADVisio
#* Created: 07/01/2014
#* Author: David S. Elias
#* Arguments: $ObjectName, $OutputPath, $DisplayDuplicates,
#*            $DuplicatesFilePath, $GroupMembers, $ServerCheck
#*=============================================
#* Purpose: Using the above functions creates a CSV
#*          that shows Nested AD Group Memberships
#*          that can be imported into MS Visio 2010
#*          without any additional data manipulation
#*=============================================

Function Get-LocalGroups{
    param(

        [parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true)]  
        [string]$ComputerName
    )  
      
    Begin{
        
        # Test that the computer is powered on and available on the network

        Try{
            
            $Connected = (Test-Connection -ComputerName $ComputerName -Count 1 -Quiet -ErrorAction Stop)
        
        }Catch{
        
            Write-Error "Unable to connect to $ComputerName" -Category ConnectionError -TargetObject $ComputerName
            PAUSE
            #EXIT 1
        }
        
        # Test that the user has rights to make WMI calls to the Computer

        Try{
            
            $hostname = (Get-WmiObject -ComputerName $ComputerName -Class Win32_ComputerSystem).Name
        
        }Catch{
        
            Write-Error "Unable to make a WMI call to $ComputerName" -Category AuthenticationError -TargetObject $ComputerName
            PAUSE
            #EXIT 1
        }
    }  
      
    Process{

        $GroupList=$Allmembers=$CorpMembers=@()
        $LocalGroupNames = (Get-WmiObject -ComputerName $hostname -Query "SELECT * FROM Win32_Group WHERE Domain='$Hostname'").name
        
        ForEach($GroupName in $LocalGroupNames){

            $wmi = Get-WmiObject -ComputerName $hostname -Query "SELECT * FROM Win32_GroupUser WHERE GroupComponent=`"Win32_Group.Domain='$Hostname',Name='$GroupName'`""

            if($wmi -ne $null) {

                ForEach($item in $wmi) {  

                    $data = $item.PartComponent -split "\," 
                    $domain = ($data[0] -split "=")[1] 
                    $name = ($data[1] -split "=")[1] 
                    $User = ("$domain\$name").Replace("""","") 

                    $member=@()
                    $member=New-Object PSObject
                    $member | Add-Member -MemberType NoteProperty -Name "ServerName" -Value $Hostname
                    $member | Add-Member -MemberType NoteProperty -Name "User/Group" -Value $User
                    $member | Add-Member -MemberType NoteProperty -Name "LocalGroup" -Value $GroupName

                    $Allmembers += $member

                }
            }
        }
    }
    
    End{
    
        $CorpMembers += $Allmembers | Where-Object {$_.'User/Group' -like "CORP*"} | Sort-Object localgroup
        
        Return $CorpMembers

    }
}

Function Get-SQLUsers{
    [CmdletBinding()]
    [OutputType([Object])]
    Param(

        # Computer Name to be checked for SQL.
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [string]
        $ComputerName,

        # Group Name to be looked for in SQL Server.
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=1)]
        [string]
        $GroupName
    )

    Begin{

        # Check for SQL PSsnappin and Cmdlets
        $Registered = (Get-PSSnapin -Registered).Name

        IF($Registered -notcontains "SqlServerCmdletSnapin100"){

            Write-Error "SQL Server Snapin is not present" -Category NotInstalled
            PAUSE
            # EXIT 1

        }ELSEIF($Registered -contains "SqlServerCmdletSnapin100"){

            Add-PSSnapin -Name SqlServerCmdletSnapin100 -ErrorAction Stop
            $Registered = $true

        }

        Try{

            Get-Command -Name invoke-sqlcmd | Out-Null

        }Catch{

            Write-Error "Unable to find Command invoke-sqlcmd" -Category ObjectNotFound invoke-sqlcmd
            PAUSE
            #EXIT 1

        }

        # Test Server Connection

        Write-host "Testing Server Connectivity - $ComputerName" -ForegroundColor Green

        Try{
            
            $Alive = Test-Connection -ComputerName $ComputerName -Count 2 -Quiet -ErrorAction Stop

        }Catch{
            
            Write-Error "Unable to connect to Server $ComputerName" -Category ConnectionError -TargetObject $ComputerName
            PAUSE
            #EXIT 1

        }

        # Test SQL Presence

        IF($Alive -eq $true){

        Write-host "Testing SQL Service Presence on Server - $ComputerName" -ForegroundColor Green

            Try{
            
                $SQLService = Get-Service -Name MSSQLSERVER -ComputerName $ComputerName -ErrorAction Stop

            }Catch{

                Write-Error "MSSQLSERVER is not on Server $ComputerName" -Category InvalidResult
                PAUSE
                #EXIT 1
            }

        }ELSEIF($Alive -eq $false){
            
            Write-Error "Can Not Contact Server" -Category ConnectionError -TargetObject $ComputerName
            PAUSE
            Exit 1
        }

        Write-host "Checking that SQL Service is running - $ComputerName" -ForegroundColor Green

        # Test SQL Connection and Database name

        IF($SQLService.Status -like "Running"){

            Try{
                
                $DatabaseList = ( invoke-sqlcmd -query "SELECT name FROM master..sysdatabases ORDER BY Name" -serverinstance "$ComputerName" -Database master -QueryTimeout 100)

            }Catch{

                Write-Error "The account you are running does not have SQL access on Server $ComputerName" -Category AuthenticationError -TargetObject $ComputerName
                PAUSE
                #EXIT 1
            }
        }
    }
    
    Process{

        IF($Registered -eq $true){

            Try{
    
                $SQLUserList = (invoke-sqlcmd -query "SELECT name, sysadmin FROM master..syslogins WHERE NAME LIKE 'CORP\%' ORDER BY Name" -database "master" -serverinstance "$ComputerName" -QueryTimeout 100)
    
            }Catch{

                $SQLUserList = $null
                Write-Error "Unable to retrieve list of users from $ComputerName" -Category ConnectionError -TargetObject $ComputerName

            }

            IF($SQLUserList -ne $null){

                IF($SQLUserList.Name -contains "CORP\$GroupName"){

                    $FoundUser = $SQLUserList | Where-Object {$_.name -like "CORP\$GroupName"}
                    IF($FoundUser.sysadmin -eq 1){

                        $SysAdmin = "SysAdmin"

                    }ELSE{

                        $SysAdmin = "N/A"

                    }

                    Try{
            
                        $DBGroups=@()
                        $ob = New-Object -TypeName PSObject
                        $ob | Add-Member -MemberType NoteProperty -Name "Server" -Value $ComputerName
                        $ob | Add-Member -MemberType NoteProperty -Name "GroupName" -Value $FoundUser.name
                        $ob | Add-Member -MemberType NoteProperty -Name "Database" -Value $SysAdmin
                        $DBGroups += $ob

                        ForEach($DB in $DatabaseList){
            
                            $DB=$DB.name
                            $SQLGroup=$null
                            $SQLGroup = (invoke-sqlcmd -query "SELECT * FROM sys.database_principals WHERE name LIKE 'CORP\$GroupName'" -database "$DB" -serverinstance "$ComputerName" -QueryTimeout 100)

                            IF(($SQLGroup -as [bool]) -eq $true){
                                
                                $ob = New-Object -TypeName PSObject
                                $ob | Add-Member -MemberType NoteProperty -Name "Server" -Value $ComputerName
                                $ob | Add-Member -MemberType NoteProperty -Name "GroupName" -Value $SQLGroup.name
                                $ob | Add-Member -MemberType NoteProperty -Name "Database" -Value $DB

                                $DBGroups += $ob

                            }ELSEIF(($SQLGroup -as [bool]) -eq $false){
            
                                #Write-Output "No User Group named CORP\$GroupName found in $DB"
            
                            }
                        }

                    }catch{

                        Write-Output "Unable to connect to $ComputerName.$Database"
        
                    }
                }ELSEIF($SQLUserList.Name -notcontains "CORP\$GroupName"){

                    Write-Error "CORP\$GroupName was not found in SQL Server on $ComputerName" -Category ObjectNotFound $GroupName

                }

            }ELSEIF($SQLUserList -eq $null){
        
                Write-Error "Unable to retrieve list of users from $ComputerName" -Category ConnectionError -TargetObject $ComputerName

            }
        }
    }
    
    End{

        IF($DBGroups.count -gt 0){
            
            Return $DBGroups
        
        }ELSEIF($DBGroups.count -lt 1){

        Write-Error "Ain't nothing worked"

        }
    }
}

Function Get-SharePermissions{
    [CmdletBinding()]
    [OutputType([Object])]
    Param(
        
        # ComputerName to look for Shares on.
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [string]
        $ComputerName

    )

    Begin{
        
        $Shares=@()
        
        Try{
            
            $Alive = Test-Connection -ComputerName $ComputerName -Count 1 -Quiet -ErrorAction Stop
            
            IF($Alive -eq $true){
                
                Try{
                    # retrieve all share names from target computer

                    $Shares = Get-WmiObject -Class win32_Share -ComputerName $ComputerName | select -ExpandProperty Name

                }Catch{

                    Write-Error "WMI Access denied on $ComputerName." -Category AuthenticationError -TargetObject $ComputerName
                    PAUSE
                    # EXIT 1

                }
            }
        
        }Catch{

            Write-Error "Unable to connect to $ComputerName" -Category ConnectionError -TargetObject $ComputerName
            PAUSE
            # EXIT 1
        }
    }
    
    Process{

        $Master=@()
        ForEach ($shareinstance in $shares){

            $ACL=@()
            
            # Retrieve security data from each Share in the $Shares array

            $SecuritySettings = Get-WMIObject -ComputerName $ComputerName -Class win32_LogicalShareSecuritySetting -Filter "name='$Shareinstance'"
    
            Try{
    
                $SecurityDescriptor = $SecuritySettings.GetSecurityDescriptor().Descriptor
                
                # Decipher UserIDs from share security data
                ForEach($ace in $SecurityDescriptor.DACL){
            
                    $UserName = $ace.Trustee.Name
        
                    IF($ace.Trustee.Domain -ne $Null){
                    
                        $UserName = "$($ace.Trustee.Domain)\$UserName"
                    
                    }
        
                    IF($ace.Trustee.Name -eq $Null){
                    
                        $UserName = $ace.Trustee.SIDString
                    
                    }
                    
                    # Cheating Hash object
                    [Array]$ACL += New-Object Security.AccessControl.FileSystemAccessRule($UserName, $ace.AccessMask, $ace.AceType)
                
                }

            }Catch{
                
                # This is where a log entry should go as it does not necessarily mean an error.
            
            }

            ForEach($Record in $ACL){

                $obj=New-Object PSObject
                $obj | Add-Member -MemberType NoteProperty -Name "Server" -Value $Computername
                $obj | Add-Member -MemberType NoteProperty -Name "ShareName" -Value $shareinstance
                $obj | Add-Member -MemberType NoteProperty -Name "IdentityReference" -Value $Record.IdentityReference
                $obj | Add-Member -MemberType NoteProperty -Name "AccessControlType" -Value $Record.AccessControlType
                $obj | Add-Member -MemberType NoteProperty -Name "FileSystemRights" -Value $Record.FileSystemRights
                $obj | Add-Member -MemberType NoteProperty -Name "IsInherited" -Value $Record.IsInherited
                $Master += $obj

            }
        }
    }
    
    End{
        IF($Master.count -gt 0){
        
            Return $Master

        }ELSEIF(($Master.count -lt 1) -and ($Shares -lt 1)){

            # No Shares were found on this server This is nor an error

        }ELSEIF(($Master.count -lt 1) -and ($Shares -gt 0)){
            
            ForEach($Share in $Shares){
            
                Write-Host "Unable to retrieve Share data from $ComputerName for Share Name $Share" -ForegroundColor Yellow
            
            }
        }
    }
}

<#
.Synopsis
   Enter a Valid Active Directory Group name or UserID and receive a CSV output of all memberships with an option of listing the user members of each of those sub-groups.
.DESCRIPTION
   Create-ADVisio creates a CSV file formatted for importation into Visio 2010.
   By entering a valid Active Directory UserID (samaccountname) or Group Name an organizational chart can then be generated showing memberships to within 4 nesting levels.
   As an option you can also list any user objects that are a member of each of the subsequent sub groups.
.EXAMPLE
   Create-ADVisio -ObjectName TestUser -OutputPath 'C:\temp\Groups'

   This Example uses a UserID (samaccountname)

   GlobalOversight_INT is not a member of any other groups
   IWB_Operations_INT is not a member of any other groups

   Unique_ID Department E-mail Name                                                     Telephone Title Reports_To Master_Shape
   --------- ---------- ------ ----                                                     --------- ----- ---------- ------------
   ID0                         TestUser                                                                            0           
   ID1                         Domain Users                                                             ID0        0           
   ID2                         IWB_Operations_DEV                                                       ID0        0           
   ID3                         GlobalOversight_DEV                                                      ID0        0           
   ID4                         GlobalOversight_INT                                                      ID0        0           
   ID5                         IWB_Operations_INT                                                       ID0        0           
   ID6                         IWB_Operations_UAT                                                       ID0        0           

.EXAMPLE
   Create-ADVisio -ObjectName IWB_Operations_UAT -OutputPath 'C:\temp\Groups'
   
   This Example uses a Group Name

   GENERICAPP61_LOGS_RW is not a member of any other groups
   GENERICAPP62_LOGS_RW is not a member of any other groups
   GENERICAPP128_LOGS_RW is not a member of any other groups
   GENERICAPP130_LOGS_RW is not a member of any other groups
   GENERICSQL164_DatabaseName_RO is not a member of any other groups
   GENERICSQL164_DatabaseName_RWE is not a member of any other groups
   
   Unique_ID Department E-mail Name                           Telephone Title Reports_To Master_Shape
   --------- ---------- ------ ----                           --------- ----- ---------- ------------
   ID0                         IWB_Operations_UAT                                        0           
   ID1                         GENERICAPP61_LOGS_RW                           ID0        0           
   ID2                         GENERICAPP62_LOGS_RW                           ID0        0           
   ID3                         GENERICAPP128_LOGS_RW                          ID0        0           
   ID4                         GENERICAPP130_LOGS_RW                          ID0        0           
   ID5                         GENERICSQL164_DatabaseName_RO                  ID0        0           
   ID6                         GENERICSQL164_DatabaseName_RWE                 ID0        0           

.EXAMPLE
   Create-ADVisio -ObjectName    IWB_Operations_UAT -OutputPath 'C:\temp\Groups' -GroupMembers

   This Example showes the difference between the previous example and the addition of the -GroupMembers switch.

   GENERICAPP61_LOGS_RW is not a member of any other groups
   GENERICAPP62_LOGS_RW is not a member of any other groups
   GENERICAPP128_LOGS_RW is not a member of any other groups
   GENERICAPP130_LOGS_RW is not a member of any other groups
   GENERICSQL164_DatabaseName_RO is not a member of any other groups
   GENERICSQL164_DatabaseName_RWE is not a member of any other groups
   Checking for User Members of GENERICAPP61_LOGS_RW
   Checking for User Members of GENERICAPP62_LOGS_RW
   Checking for User Members of GENERICAPP128_LOGS_RW
   Checking for User Members of GENERICAPP130_LOGS_RW
   Checking for User Members of GENERICSQL164_DatabaseName_RO
   Checking for User Members of GENERICSQL164_DatabaseName_RWE

   Unique_ID Department                     E-mail Name                                         Telephone Title Reports_To Master_Shape
   --------- ----------                     ------ ----                                         --------- ----- ---------- ------------
   ID0                                             IWB_Operations_UAT                                                      0           
   ID1                                             GENERICAPP61_LOGS_RW                                         ID0        0           
   ID2                                             GENERICAPP62_LOGS_RW                                         ID0        0           
   ID3                                             GENERICAPP128_LOGS_RW                                        ID0        0           
   ID4                                             GENERICAPP130_LOGS_RW                                        ID0        0           
   ID5                                             GENERICSQL164_DatabaseName_RO                                ID0        0           
   ID6                                             GENERICSQL164_DatabaseName_RWE                               ID0        0           
   ID7       GENERICAPP61_LOGS_RW                  SVC-FinanceAppUAT2 | SVC-FinanceAppUAT2                      ID1        4           
   ID8       GENERICAPP62_LOGS_RW                  SVC-FinanceAppUAT2 | SVC-FinanceAppUAT2                      ID2        4           
   ID9       GENERICAPP128_LOGS_RW                 HMSATH | Samual Thomas                                       ID3        4           
   ID10      GENERICAPP128_LOGS_RW                 SVC-FinanceAppUAT4 | SVC-FinanceAppUAT4                      ID3        4           
   ID11      GENERICAPP130_LOGS_RW                 HMSATH | Samual Thomas                                       ID4        4           
   ID12      GENERICAPP130_LOGS_RW                 SVC-FinanceAppUAT4 | SVC-FinanceAppUAT4                      ID4        4           
   ID13      GENERICSQL164_DatabaseName_RO         Svc-FinanceAppUATETL | Svc-FinanceAppUATETL                  ID5        4           
   ID14      GENERICSQL164_DatabaseName_RO         SQLSERVSQL163 |  SQLSERVSQL163                               ID5        4           
   ID15      GENERICSQL164_DatabaseName_RO         SVCFinanceAppSSRSUAT5 | SVCFinanceAppSSRSUAT5                ID5        4           
   ID16      GENERICSQL164_DatabaseName_RWE        Svc-FinanceAppUATETL | Svc-FinanceAppUATETL                  ID6        4           

.EXAMPLE
   Create-ADVisio -ObjectName FinanceApp_support -OutputPath 'C:\temp\Groups' -DisplayDuplicates

   Adding the -DisplayDuplicates switch will look for and display any duplicated group memberships.
   Duplicated membership inheritances can cause extreme difficulty when attempting to troubleshoot security issues.

   GENERICSQL91_ALL_DBs_RO is not a member of any other groups
   GENERICSQL51_ALL_DBs_RO is not a member of any other groups
   GENERICSQL136_ALL_DBs_RO is not a member of any other groups
   GENERICSQL169_ALL_DBs_RO is not a member of any other groups
   GENERICSQL94_All_DBs_RO is not a member of any other groups
   GENERICSQL68_ALL_DBs_RO is not a member of any other groups
   GENERICSQL96_ALL_DBs_RO is not a member of any other groups
   GENERICSQL164_ALL_DBs_RO is not a member of any other groups
   GENERICSQL90_ALL_DBs_RO is not a member of any other groups
   GENERICSQL133_ALL_DBs_RO is not a member of any other groups

   BEGIN DUPLICATE GROUP MEMBERSHIPS

   Unique_ID Department E-mail Name                                        Telephone Title Reports_To Master_Shape
   --------- ---------- ------ ----                                        --------- ----- ---------- ------------
   ID6                         GENERICSQL82_LOCALADMIN                                     ID0        0           
   ID7                         GENERICSQL83_LOCALADMIN                                     ID0        0           
   ID8                         GENERICSQL84_LOCALADMIN                                     ID0        0           
   ID9                         GENERICAPP89_LOCALADMIN                                     ID0        0           
   ID35                        GENERICAPP90_LOCALADMIN                                     ID0        0           
   ID38                        GENERICSQL84_STAGING_File_AIA_RO                            ID1        1           
   ID41                        GENERICSQL84_STAGING_File_EIA_RO                            ID2        1           
   ID58                        GENERICAPP178_LOGS$_RO                                      ID13       1           
   ID59                        GENERICAPP179_LOGS$_RO                                      ID13       1           
   ID108                       GENERICSQL82_LOCALADMIN                                     ID34       1           
   ID109                       GENERICSQL83_LOCALADMIN                                     ID34       1           
   ID110                       GENERICSQL84_LOCALADMIN                                     ID34       1           
   ID111                       GENERICAPP89_LOCALADMIN                                     ID34       1           
   ID112                       GENERICAPP90_LOCALADMIN                                     ID34       1           
   ID137                       GENERICAPP178_LOGS$_RO                                      ID57       2           
   ID138                       GENERICAPP179_LOGS$_RO                                      ID57       2           


   END DUPLICATE GROUP MEMBERSHIPS

   Unique_ID Department E-mail Name                                        Telephone Title Reports_To Master_Shape
   --------- ---------- ------ ----                                        --------- ----- ---------- ------------
   ID0                         FinanceApp_SUPPORT                                                     0           
   ID1                         Superusers_DEV                                              ID0        0           
   ID2                         GlobalOversight_DEV                                         ID0        0           
   ID3                         GlobalOversight_INT                                         ID0        0           
   ID4                         Superusers_INT                                              ID0        0           
   ID5                         DATA-FMG-FinanceApp-EXTRACT-RO                              ID0        0           
   ID6                         GENERICSQL82_LOCALADMIN                                     ID0        0           
   ID7                         GENERICSQL83_LOCALADMIN                                     ID0        0           
   ID8                         GENERICSQL84_LOCALADMIN                                     ID0        0           

#>

Function Create-ADVisio{

    [CmdletBinding()]
    [OutputType([object])]

    Param(

        # A valid Active Directory User ID or Group Name (samaccountname)
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [string]
        $ObjectName,

        # UNC or local path to create the CSV output file in
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=1)]
        [string]
        $OutputPath,

        # Enables looking for and displaying duplicate group memberships. Does not function if -GroupMembers is used
        [switch]
        $DisplayDuplicates,

        # UNC or local path to create a CSV output file of Dulicate memberships
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true,
                   Position=3)]
        [string]
        $DuplicatesFilePath,

        # This checks all groups
        [switch]
        $GroupMembers,

        # Enables looking for Group Membership in Server local security groups and SQL Server securities
        [switch]
        $ServerCheck,

        # Servername prefix (Assumes standard naming conventions for servers)
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true)]
        $Preefix

    )

    Begin{
        
        $ID = $ObjectName
        $FileLocation = $OutputPath

        <################################################################
    
            Check that the ID entered exists in Active Directory
    
        ################################################################>

        Try{

            $ID = Get-ADUser -Identity $ID -ErrorAction Stop

        }Catch{

            Try{

                $ID = Get-ADGroup -Identity $ID -ErrorAction Stop

            }Catch{

                Write-Error "The ID entered can not be found in Active Directory" -Category ObjectNotFound -ErrorAction Stop
                PAUSE
                Break

            }
        }

        # Set ID back to proper name after validation
        $ID = ($ID).samaccountname
        
        #Set Number to be used 
        $Number = 0

        # Create Table
        $tabName = "VisioTest01"
        $table = New-Object system.Data.DataTable “$tabName”

        # All Columns to be created
        $Col1 = New-Object system.Data.DataColumn Unique_ID,([string])
        $Col2 = New-Object system.Data.DataColumn Department,([string])
        $Col3 = New-Object system.Data.DataColumn E-mail,([string])
        $Col4 = New-Object system.Data.DataColumn Name,([string])
        $Col5 = New-Object system.Data.DataColumn Telephone,([string])
        $Col6 = New-Object system.Data.DataColumn Title,([string])
        $Col7 = New-Object system.Data.DataColumn Reports_To,([string])
        $Col8 = New-Object system.Data.DataColumn Master_Shape,([string])

        # Add Columns to Table
        $table.columns.add($Col1)
        $table.columns.add($Col2)
        $table.columns.add($Col3)
        $table.columns.add($Col4)
        $table.columns.add($Col5)
        $table.columns.add($Col6)
        $table.columns.add($Col7)
        $table.columns.add($Col8)

        # Create Master Row Record
        $Row = $table.NewRow()

        # The hyphen in "E-Mail" can't be used in this context
        $email = 'E-Mail'

        # Shape is a field designator for Visio to destiguish between objects and also designates membership relationship in org chart
        $Shape = 0

        $Row.Unique_ID = "ID$Number"
        $Row.Department = ""
        $Row.$email = ""
        $Row.Name = "$ID"
        $Row.Telephone = ""
        $Row.Title = ""
        $Row.Reports_To = ""
        $Row.Master_Shape = "$Shape"

        # Add the row to the table
        $table.Rows.Add($Row)

    }
    
    Process{

        <################################################################
    
               Table has been created and first record populated
    
            Now Checking Active Directory and populate table based
            
                     on Principal Group Memberships found

        ################################################################>

        # Start gathering Group Names at Level 1
        $List1 = Get-ADPrincipalGroupMembership -Identity $ID
        $List2 = ($List1).name
        $SubGroups=@()

        ForEach($L in $List2){
    
            $Number = $Number + 1

            $Row = $table.NewRow()
            $Row.Unique_ID = "ID$Number"
            $Row.Department = ""
            $Row.$email = ""
            $Row.Name = "$L"
            $Row.Telephone = ""
            $Row.Title = ""
            $Row.Reports_To = "ID0"
            $Row.Master_Shape = "$Shape"

            $table.Rows.Add($Row)

        }

        # Start gathering Group Names at Level 2

        $SubGroups = $Table | Where-Object {$_.Unique_ID -notlike "ID0"}
        $List10=$Empty1=$SubGroups1=$SubGroups2=$SubGroups3=$SubGroups4=@()
        $Shape = $Shape + 1

        ForEach($SubGroup in $SubGroups){

            $SubSubGroups = (Get-ADPrincipalGroupMembership -Identity $SubGroup.name)
    
            IF($SubSubGroups.count -ge 1){

                ForEach($Sub in $SubSubGroups){      
            
                    $List10 += $Sub

                    $Number = $Number + 1

                    $Row = $table.NewRow()
    
                    $Row.Unique_ID = "ID$Number"
                    $Row.Department = ""
                    $Row.$email = ""
                    $Row.Name = "$($Sub.name)"
                    $Row.Telephone = ""
                    $Row.Title = ""
                    $Row.Reports_To = "$($SubGroup.Unique_ID)"
                    $Row.Master_Shape = "$Shape"

                    $table.Rows.Add($Row)

                }

            }ELSE{
        
                Write-host "$($SubGroup.name) is not a member of any other groups" -ForegroundColor Cyan
                $Empty1 += $SubGroup.name
    
            }
        }

        # Start gathering Group Names at Level 3

        $SubGroups2 = $Table | Where-Object {($_.name -in $list10.name) -and ($_.name -notin $Empty)}
        $List11=$Empty2=@()
        $Shape = $Shape + 1

        ForEach($SubGroup in $SubGroups2){

            $SubSubGroups = (Get-ADPrincipalGroupMembership -Identity $SubGroup.name)
    
            IF($SubSubGroups.count -ge 1){

                ForEach($Sub in $SubSubGroups){      
            
                    $List11 += $Sub

                    $Number = $Number + 1

                    $Row = $table.NewRow()
    
                    $Row.Unique_ID = "ID$Number"
                    $Row.Department = ""
                    $Row.$email = ""
                    $Row.Name = "$($Sub.name)"
                    $Row.Telephone = ""
                    $Row.Title = ""
                    $Row.Reports_To = "$($SubGroup.Unique_ID)"
                    $Row.Master_Shape = "$Shape"

                    $table.Rows.Add($Row)

                }

            }ELSE{
        
                Write-host "$($SubGroup.name) is not a member of any other groups" -ForegroundColor Cyan
                $Empty2 += $SubGroup.name

            }
        }

        # Start gathering Group Names at Level 4

        $SubGroups3 = $Table | Where-Object {($_.name -in $list11.name) -and ($_.name -notin $Empty)}
        $List12=$Empty3=@()
        $Shape = $Shape + 1

        ForEach($SubGroup in $SubGroups3){

            $SubSubGroups = (Get-ADPrincipalGroupMembership -Identity $SubGroup.name)
    
            IF($SubSubGroups.count -ge 1){

                ForEach($Sub in $SubSubGroups){      
            
                    $List12 += $Sub

                    $Number = $Number + 1

                    $Row = $table.NewRow()
    
                    $Row.Unique_ID = "ID$Number"
                    $Row.Department = ""
                    $Row.$email = ""
                    $Row.Name = "$($Sub.name)"
                    $Row.Telephone = ""
                    $Row.Title = ""
                    $Row.Reports_To = "$($SubGroup.Unique_ID)"
                    $Row.Master_Shape = "$Shape"

                    $table.Rows.Add($Row)

                }

            }ELSE{
        
                Write-host "$($SubGroup.name) is not a member of any other groups" -ForegroundColor Cyan
                $Empty3 += $SubGroup.name

            }
        }

        <##################################################

            Checking for Duplicate Group Memberships

        ##################################################>

        
        IF(($DisplayDuplicates -eq $true) -and ($GroupMembers -eq $false)){

            $Single = $table | select name -Unique
            
            $Duplicates = Compare-Object -ReferenceObject ($Single).name -DifferenceObject ($table).name
            
            $Duplicates = $Duplicates.inputobject
        
            
            IF($Duplicates.count -ge 1){
            
                $AllDups = $Table | Where-Object {$_.name -in $Duplicates}

                Write-Host "BEGIN DUPLICATE GROUP MEMBERSHIPS" -ForegroundColor Magenta
                
                $AllDups | FT -AutoSize
                
                Write-Host "END DUPLICATE GROUP MEMBERSHIPS" -ForegroundColor Magenta

                $tabCsv = $AllDups | export-csv $DuplicatesFilePath\$ID-Duplicates.csv -noType

            }
        }

        <##################################################

                Checking for Group User Members

        ##################################################>

        Try{

            $IU = Get-ADUser -Identity $ID -ErrorAction Stop

        }Catch{

            $IU = Get-ADGroup -Identity $ID -ErrorAction Stop

        }

        IF(($IU).ObjectClass -eq "User"){
            
            $IDisUser = $true
        
        }ELSEIF(($IU).ObjectClass -ne "User"){
        
            $IDisUser = $false
        
        }


        IF($IDisUser -eq $false){

            IF($GroupMembers -eq $true){
            
                $Groups = $table | Where-Object {($_.name -ne $ID) -and ($_.name -notlike "Domain Users") -and ($_.name -notlike "APP*") -and ($_.name -notlike "DATA-ISS-ONCALL-RO")}

                ForEach($Group in $Groups){
                
                    $Memebers=$Users=@()
                    $Memebers = (Get-ADGroupMember -Identity ($Group).name).samaccountname
                
                    Write-Host "Checking for User Members of $($Group.name)" -ForegroundColor Green

                    ForEach($User in $Memebers){

                        Try{
                        
                            $IsUser = $false
                        
                            $IsUser = (Get-ADUser -Identity $User -ErrorAction SilentlyContinue) -as [bool]
                        
                            IF($IsUser -eq $true){

                                $Users += $User
                        
                            }
                    
                        }Catch{
                    
                            $IsUser = $false
                    
                        }
                    }

                    ForEach($UserID in $Users){
                    
                        $NN = Get-ADUser -Identity $UserID

                        $NiceName = $NN.SamAccountName + ' ' + '|' + ' ' + $NN.GivenName + ' ' + $NN.Surname

                        $Number = $Number + 1

                        $Row = $table.NewRow()
    
                        $Row.Unique_ID = "ID$Number"
                        $Row.Department = "$($Group.Name)"
                        $Row.$email = ""
                        $Row.Name = "$NiceName"
                        $Row.Telephone = ""
                        $Row.Title = ""
                        $Row.Reports_To = "$($Group.Unique_ID)"
                        $Row.Master_Shape = "4"

                        $table.Rows.Add($Row)

                    }
                }
            }

        }ELSEIF($IDisUser -eq $true){



        }


        <######################################################

            Checking for local computer group memberships

            based on AD Resource Group Names

        #######################################################>

        IF($ServerCheck -eq $true){

            $TableGroups = $Table
            $GroupNames = $TableGroups | Where-Object {($_.name -like "$Preefix*") -and ($_.name -match "\d_")}
            $ServerList=$ServerGroups=@()

            ForEach($TG in $GroupNames){
    
                # Seperate the Servername from the group name
                $Name = $TG.name.tostring()
                $Server = $name.Split('_')[0]
                $ServerList += $Server
    
                $Lookup = $TG.Name.ToString()

                Write-Host "Checking $Server for the presence of $Lookup"

                # Look through members of all local server groups for the AD Group
                Try{
    
                    $FoundGroup = Get-LocalGroups -ComputerName $Server -ErrorAction SilentlyContinue | Where-Object {$_.'User/Group' -like "*$Name"}
    
                }Catch{
    
                    $FoundGroup = $null
    
                }
    
                IF($FoundGroup -ne $null){
                    #$ServerGroups += $FoundGroup

                    $Number = $Number + 1
                    $Row = $table.NewRow()
    
                    $Row.Unique_ID = "ID$Number"
                    $Row.Department = ""
                    $Row.$email = ""
                    $Row.Name = "$($FoundGroup.ServerName)"
                    $Row.Telephone = ""
                    $Row.Title = "$($FoundGroup.LocalGroup)"
                    $Row.Reports_To = "$($TG.Unique_ID)"
                    $Row.Master_Shape = "4"

                    $table.Rows.Add($Row)


                }ELSEIF($FoundGroup -eq $null){
    
                # AD Group was not found in local windows group so look in SQL Server
                <######################################################

                            Checking for SQL Server Logins

                           based on AD Resource Group Names

                #######################################################>

                    Try{
    
                        $SQLGroup = Get-SQLUsers -ComputerName $Server -GroupName $lookup -ErrorAction Stop
    
                    }Catch{
    
                        $SQLGroup = $null
    
                    }

                    IF($SQLGroup -ne $null){
                    #$ServerGroups += $SQLGroup

                        IF (($SQLGroup.count -gt 1)-and ($SQLGroup.database -notcontains "SysAdmin")){
                        
                            $SG = $SQLGroup | Where-Object {$_.Database -notlike "N/A"}

                            ForEach($Login in $SG){

                                $Number = $Number + 1
                                $Row = $table.NewRow()
                        
                                $Row.Unique_ID = "ID$Number"
                                $Row.Department = ""
                                $Row.$email = ""
                                $Row.Name = "$($Login.Server)"
                                $Row.Telephone = ""
                                $Row.Title = "$($Login.Database)"
                                $Row.Reports_To = "$($TG.Unique_ID)"
                                $Row.Master_Shape = "4"

                                $table.Rows.Add($Row)
                            }

                        }ELSEIF(($SQLGroup.count -eq 1) -and ($SQLGroup.database -contains "SysAdmin")){

                            ForEach($Login in $SQLGroup){

                                $Number = $Number + 1
                                $Row = $table.NewRow()
                        
                                $Row.Unique_ID = "ID$Number"
                                $Row.Department = ""
                                $Row.$email = ""
                                $Row.Name = "$($Login.Server)"
                                $Row.Telephone = ""
                                $Row.Title = "$($Login.Database)"
                                $Row.Reports_To = "$($TG.Unique_ID)"
                                $Row.Master_Shape = "4"

                                $table.Rows.Add($Row)
                            }

                        }ELSE{

                            ForEach($Login in $SQLGroup){

                                $Number = $Number + 1
                                $Row = $table.NewRow()
                        
                                $Row.Unique_ID = "ID$Number"
                                $Row.Department = ""
                                $Row.$email = ""
                                $Row.Name = "$($Login.Server)"
                                $Row.Telephone = ""
                                $Row.Title = "$($Login.Database)"
                                $Row.Reports_To = "$($TG.Unique_ID)"
                                $Row.Master_Shape = "4"

                                $table.Rows.Add($Row)
                            }
                        }
                    }
                }
            }

                <######################################################

                                Checking Server Shares

                           based on AD Resource Group Names

                #######################################################>
            ForEach($GN in $GroupNames){
    
                # Seperate the Servername from the group name
                $Name = $GN.name.tostring()
                $Server = $name.Split('_')[0]
                
                $SharesFound = Get-SharePermissions -ComputerName $Server | Where-Object {$_.IdentityReference -like "*\$($GN.name)"}
                
                Write-output "Checking Shares on $Server"

                ForEach($Share in $SharesFound){

                    $Number = $Number + 1
                    $Row = $table.NewRow()
                        
                    $Row.Unique_ID = "ID$Number"
                    $Row.Department = ""
                    $Row.$email = ""
                    $Row.Name = "\\$($Share.Server)\$($Share.ShareName)"
                    $Row.Telephone = ""
                    $Row.Title = "$($Share.AccessControlType) - $($Share.FileSystemRights)"
                    $Row.Reports_To = "$($GN.Unique_ID)"
                    $Row.Master_Shape = "4"

                    $table.Rows.Add($Row)
                }

            }


        }
    }

    End{

        # TEST BY DISPLAYING TABLE #
        $table | format-table -AutoSize -Wrap

        Try{
    
            $TP = Test-path -path $FileLocation\$ID.csv -ErrorAction stop
        
            IF($TP -eq $true){
        
                Try{
                
                    Remove-Item -Path $FileLocation\$ID.csv -ErrorAction stop
                
                }Catch{
                    
                    Write-Error "Unable to remove existing File $FileLocation\$ID.csv" -Category PermissionDenied -ErrorAction Stop
                
                }
        
            }

        }Catch{
    
            $TP = Test-Path -path $FileLocation -ErrorAction stop

            IF($TP -ne $true){
        
                Write-Error "Unable to connect to $FileLocation" -Category ConnectionError -ErrorAction Stop
    
            }
        }

        <####################################################################
    
            Create output File in proper CSV format for Visio consumption

        ####################################################################>

        $tabCsv = $table | export-csv $FileLocation\$ID.csv -noType

    }
}