FUNCTION Confirm-TargetSystem{
    <#
        .SYNOPSIS
        Attempt to verify that a system is the intended target system and,
        optionally, is the intended logged on user.

        .DESCRIPTION
        Attempt to verify that a system is the intended target system and,
        optionally, is the intended logged on user.  Will use Get-CIMInstance
        to reach out to target system and attempt to match the computer name
        or IP address and, if desired, the username as well.

        .EXAMPLE
        Confirm-TargetSystem -sysID PSCenter

        Expected results of a confirmed system:

        WARNING: You have not entered identifying information for a user on
        the target system.  Confirm-TargetSystem will only attempt to verify
        the system, not the user at this time.


        sysID             : PSCenter
        sysName           : PSCENTER
        sysIP             : 169.254.255.162
        usrID             : NA
        usrLoggedOn       : NA
        isIntendedSystem  : True
        isIntendedUser    : False
        usingTechnician   : PSCENTER\PSAdmin
        targetSystemNotes : NA

        .EXAMPLE
        Confirm-TargetSystem -sysID PSCenter -Silent

        Optional warnings suppressed.
        Expected results of a confirmed system:

        sysID             : PSCenter
        sysName           : PSCENTER
        sysIP             : 169.254.255.162
        usrID             : NA
        usrLoggedOn       : NA
        isIntendedSystem  : True
        isIntendedUser    : False
        usingTechnician   : PSCENTER\PSAdmin
        targetSystemNotes : NA

        .EXAMPLE
        Confirm-TargetSystem -sysID PSCenters

        Expected results of a failed confirmation where the system could not
        be reached.

        sysID             : PSCenters
        sysName           : NA
        sysIP             : NA
        usrID             : NA
        usrLoggedOn       : NA
        isIntendedSystem  : False
        isIntendedUser    : False
        usingTechnician   : PSCENTER\PSAdmin
        targetSystemNotes : Name resolution of PSCenters failed

        .EXAMPLE
        Confirm-TargetSystem -sysID PSCenter -usrID PSAdmin

        If all results confirmed, feedback would appear as:

        sysID             : PSCenter
        sysName           : PSCENTER
        sysIP             : 169.254.255.162
        usrID             : PSAdmin
        usrLoggedOn       : PSAdmin
        isIntendedSystem  : True
        isIntendedUser    : True
        usingTechnician   : PSCENTER\PSAdmin
        targetSystemNotes : NA

        .EXAMPLE
        Confirm-TargetSystem -sysID 192.168.1.21

        Expected results if you attempt to check an IP address instead of
        a computer name.

        sysID             : NA
        sysName           : NA
        sysIP             : 192.168.1.21
        usrID             : NA
        usrLoggedOn       : NA
        isIntendedSystem  : False
        isIntendedUser    : False
        usingTechnician   : PSCENTER\PSAdmin
        targetSystemNotes : This function is intended to be used to verify you are
                            not facing DNS issues and are correctly reaching an
                            intended system by name. IP address checks aren't necessary
                            as they would bypass any DNS issues.

                            Thank you. Confirm-TargetSystem will now stop.

        .EXAMPLE
        Confirm-TargetSystem -sysID PSCenter -tCred BettyBoop

        Expected results if your credentials are invalid on the remote system.

        WARNING: You have not entered identifying information for a user on the target
        system.  Confirm-TargetSystem will only attempt to verify the system, not the
        user at this time.


        sysID             : PSCenter
        sysName           : NA
        sysIP             : 169.254.255.162
        usrID             : NA
        usrLoggedOn       : NA
        isIntendedSystem  : False
        isIntendedUser    : False
        usingTechnician   : BettyBoop
        targetSystemNotes : You do not have access to this system with your current
                            credentials. Perhaps try other credentials using the
                            -tCred option if you haven't already?
                            Error:
                            Access is denied.

        .EXAMPLE
        IF(Confirm-TargetSystem -sysID PSCenter -usrID PSAdmin -Silent |
            SELECT -ExpandProperty isIntendedSystem){Write-Host "Hello"}

        To use in another script to confirm whether or not to proceed with
        an intended action/configuration, above will return true and act if
        confirmation of the system is clear.

        .PARAMETER sysID
        sysID is looking for a computer name. Name or FQDN can be accepted.

        .PARAMETER usrID
        usrID is looking to match the user's login ID.  When it reaches out
        to the remote system, if a logged on user is found, it will grab the
        user's name.  This will be username only, not domain.  For example,
        if I'm logged in as PSCenter\PSAdmin, attempting to match
        -usrID PSAdmin will come back successful.  Matching
        -usrID PSCenter\PSAdmin will not match.

        .PARAMETER tCred
        tCred is aiming to collect alternate creds if you need to connect with the
        remote system using another account.  If a technician is running PowerShell
        as MSmith but needs to reach out with the tech cred PSAdmin, they would
        include -tCred PSAdmin. When they do, they'll be prompted for their current
        password and then PSAdmin will be the account used to reach the remote
        system.

        Domain accounts are supported here as well.  If the domain is needed, it
        can be added in the format Domain\Username.
        E.g. -tCred PSCenter\PSAdmin

        .LINK
        http://blog.everydaypowershell.com/2019/03/confirm-targetsystem.html

        .NOTES
        Mark Smith
        MarkTSmith@everydayPowerShell.com
        blog.everydayPowerShell.com
        Twitter: @edPowerShell
    #>
    [cmdletbinding()]
    param (
        [String]$sysID,
        [String]$usrID,
        [Switch]$Silent,
        [PSCredential]$tCred
    )
    IF(!($silent)){
        Clear-Host
        Start-Sleep -Seconds 1
    }
    #region Functions
        FUNCTION Initialize-CIMSplat {
            param(
                [PSCredential]$tCred,
                [String]$sysID
            )
            $dcom = New-CimSessionOption -Protocol Dcom
            IF($tCred){
                $cimSesSplat = @{
                    Credential = $tCred
                    ComputerName = $sysID
                    SkipTestConnection = $true
                }
                $cimSesAlt = @{
                    Credential = $tCred
                    ComputerName = $sysID
                    SkipTestConnection = $true
                    SessionOption = $dcom
                }
                $cimInstSplat = @{
                    ClassName = "CIM_ComputerSystem"
                    CimSession = (New-CimSession @cimSesSplat)
                }
                $cimInstAlt = @{
                    ClassName = "CIM_ComputerSystem"
                    CimSession = (New-CimSession @cimSesAlt)
                }
                return $cimInstSplat,$cimInstAlt
            }
            ELSE{
                $cimSesSplat = @{
                    ComputerName = $sysID
                    SkipTestConnection = $true
                }
                $cimSesAlt = @{
                    ComputerName = $sysID
                    SkipTestConnection = $true
                    SessionOption = $dcom
                }
                $cimInstSplat = @{
                    ClassName = "CIM_ComputerSystem"
                    CimSession = (New-CimSession @cimSesSplat)
                }
                $cimInstAlt = @{
                    ClassName = "CIM_ComputerSystem"
                    CimSession = (New-CimSession @cimSesAlt)
                }
                return $cimInstSplat,$cimInstAlt
            }
        }
        FUNCTION Redo-TargetSystemResult {
            param(
                [Parameter(Position=0)]
                $sysID,
                [Parameter(Position=1)]
                $sysName,
                [Parameter(Position=2)]
                $sysIP,
                [Parameter(Position=3)]
                $usrID,
                [Parameter(Position=4)]
                $usrLoggedOn,
                [Parameter(Position=5)]
                $isIntendedSystem,
                [Parameter(Position=6)]
                $isIntendedUser,
                [Parameter(Position=7)]
                $usingTechnician,
                [Parameter(Position=8)]
                $targetSystemNotes,
                [Parameter(Position=9)]
                $result
            )
            IF($sysID){
                $result.sysID = $sysID
            }
            IF($sysName){
                $result.sysName = $sysName
            }
            IF($sysIP){
                $result.sysIP = $sysIP
            }
            IF($usrID){
                $result.usrID = $usrID
            }
            IF($usrLoggedOn){
                $result.usrLoggedOn = $usrLoggedOn
            }
            IF($isIntendedSystem -eq "False"){
                $result.isIntendedSystem = $false
            }
            ELSEIF($isIntendedSystem -eq "True"){
                $result.isIntendedSystem = $true
            }
            IF($isIntendedUser -eq "False"){
                $result.isIntendedUser = $false
            }
            ELSEIF($isIntendedUser -eq "True"){
                $result.isIntendedUser = $true
            }
            IF($usingTechnician){
                $result.usingTechnician = $usingTechnician
            }
            IF($targetSystemNotes){
                $result.targetSystemNotes = $targetSystemNotes
            }
            return $result
        }
    #endregion
    #region Param Check & Variable Prep
        #region Establish End Game Object
            $result = "" | Select-Object sysID, sysName, sysIP, usrID,
            usrLoggedOn, isIntendedSystem, isIntendedUser, usingTechnician,
            targetSystemNotes
        #endregion
        #region Splats
            $stopSplat = @{
                WarningAction = "Stop"
                ErrorAction = "Stop"
            }
        #endregion
        #region Check tCred
            IF(!($tCred)){
                $rSplat = @{
                    usingTechnician = "$($env:USERDOMAIN)\$($env:USERNAME)"
                    result = $result
                }
                $result = Redo-TargetSystemResult @rSplat
            }
            ELSE{
                $rSplat = @{
                    usingTechnician = ($tCred.UserName)
                    result = $result
                }
                $result = Redo-TargetSystemResult @rSplat
            }
        #endregion
        #region Check SysID
            IF(!($sysID)){
                #No SysID Provided
                $Msg = -JOIN(
                    "You have not entered identifying information for a ",
                    "system that needs to be confirmed.  Confirm-TargetSystem ",
                    "cannot run without this information. ",
                    "Please use `"Get-Help Confirm-TargetSystem -Examples`" ",
                    "for examples of how to use this function.`n`n",
                    "Thank you.  Confirm-TargetSystem will now stop."
                )
                $TSRSplat = @{
                    sysID = "NA"
                    sysName = "NA"
                    sysIP = "NA"
                    usrID = "NA"
                    usrLoggedOn = "NA"
                    isIntendedSystem = "False"
                    isIntendedUser = "False"
                    targetSystemNotes = $Msg
                    result = $result
                }
                $result = Redo-TargetSystemResult @TSRSplat
                return $result
            }
            ELSEIF($sysID -like "*.*.*.*"){
                #SysID Is An IP
                $Msg = -JOIN(
                    "This function is intended to be used to verify ",
                    "you are not facing DNS issues and are correctly ",
                    "reaching an intended system by name. IP address ",
                    "checks aren't necessary as they would bypass ",
                    "any DNS issues.`n`n",
                    "Thank you. Confirm-TargetSystem will now stop."
                )
                $TSRSplat = @{
                    sysID = "NA"
                    sysName = "NA"
                    sysIP = $sysID
                    usrID = "NA"
                    usrLoggedOn = "NA"
                    isIntendedSystem = "False"
                    isIntendedUser = "False"
                    targetSystemNotes = $Msg
                    result = $result
                }
                $result = Redo-TargetSystemResult @TSRSplat
                return $result
            }
            ELSE{
                TRY{
                    #System Online
                    $result = Redo-TargetSystemResult -sysID $sysID -result $result
                    $tcSplat = @{
                        ComputerName = $sysID
                        Count = 1
                        TimeToLive = 24
                    }
                    $thisTC = Test-Connection @tcSplat @stopSplat
                    $thisTC = ($thisTC |
                        Select-Object -ExpandProperty IPV4Address).IPAddressToString
                    $result = Redo-TargetSystemResult -sysIP $thisTC -result $result
                }
                CATCH{
                    #System Offline
                    $Msg = Test-NetConnection -ComputerName $sysID *>&1 |
                        Select-Object -First 1
                    $TSRSplat = @{
                        sysName = "NA"
                        sysIP = "NA"
                        usrID = "NA"
                        usrLoggedOn = "NA"
                        isIntendedSystem = "False"
                        isIntendedUser = "False"
                        targetSystemNotes = $Msg
                        result = $result
                    }
                    $result = Redo-TargetSystemResult @TSRSplat
                    return $result
                }
            }
        #endregion
        #region Check UsrID
            IF(!($usrID)){
                $Msg = -JOIN(
                    "You have not entered identifying information for a user ",
                    "on the target system.  Confirm-TargetSystem will only ",
                    "attempt to verify the system, not the user at this time."
                )
                $TSRSplat = @{
                    usrID = "NA"
                    usrLoggedOn = "NA"
                    isIntendedUser = "False"
                    result = $result
                }
                $result = Redo-TargetSystemResult @TSRSplat
                IF(!($silent)){
                    Write-Warning -Message $Msg
                }
            }
            ELSE{
                $result = Redo-TargetSystemResult -usrID $usrID -result $result
            }
        #endregion
    #endregion
    #region Verify System
        Write-Verbose -Message "Attempting To Verify System $sysID now."
        TRY{
            #region Splats
                IF($tCred){
                    $cimInstSplat = (Initialize-CIMSplat -tCred $tCred -sysID $sysID)[0]
                    $cimInstAlt = (Initialize-CIMSplat -tCred $tCred -sysID $sysID)[1]
                }
                ELSE{
                    $cimInstSplat = (Initialize-CIMSplat -sysID $sysID)[0]
                    $cimInstAlt = (Initialize-CIMSplat -sysID $sysID)[1]
                }
            #endregion
            #region Attempting System Verification
                TRY{
                    $sysName = Get-CimInstance @cimInstSplat @stopSplat |
                        Select-Object -ExpandProperty Name
                }
                CATCH{
                    TRY{
                        $sysName = Get-CimInstance @cimInstAlt @stopSplat |
                            Select-Object -ExpandProperty Name
                    }
                    CATCH{
                        throw $_.Exception.Message
                    }
                }
                $result = Redo-TargetSystemResult -sysName $sysName -result $result
                IF($sysID -eq ($result.sysName)){
                    $result = Redo-TargetSystemResult -isIntendedSystem True -result $result
                }
                ELSE{
                    $result = Redo-TargetSystemResult -isIntendedSystem False -result $result
                }
            #endregion
        }
        CATCH [Microsoft.Management.Infrastructure.CimException] {
            IF(($_.Exception.Message) -like "*WinRM cannot complete the operation*"){
                $Msg = -JOIN (
                    "The system is unavailable to communicate with the system ",
                    "$sysID.`nError:`n$($_.Exception.Message)"
                )
                $TSRSplat = @{
                    targetSystemNotes = $Msg
                    sysName = "NA"
                    isIntendedSystem = "False"
                    result = $result
                }
                $result = Redo-TargetSystemResult @TSRSplat
                return $result
            }
            ELSEIF(($_.Exception.Message) -like "*Access is denied*"){
                $Msg = -JOIN (
                    "You do not have access to this system with your current ",
                    "credentials. Perhaps try other credentials using the ",
                    "-tCred option if you haven't already?`nError:`n",
                    "$($_.Exception.Message)"
                )
                $TSRSplat = @{
                    targetSystemNotes = $Msg
                    sysName = "NA"
                    isIntendedSystem = "False"
                    result = $result
                }
                $result = Redo-TargetSystemResult @TSRSplat
                return $result
            }
            ELSE{
                $Msg = -JOIN (
                    "$($_.Exception.GetType().FullName)`n",
                    "$($_.Exception.Message)"
                )
                $TSRSplat = @{
                    targetSystemNotes = $Msg
                    sysName = "NA"
                    isIntendedSystem = "False"
                    result = $result
                }
                $result = Redo-TargetSystemResult @TSRSplat
                return $result
            }
        }
        CATCH{
            $Msg = -JOIN (
                "$($_.Exception.GetType().FullName)`n",
                "$($_.Exception.Message)"
            )
            $TSRSplat = @{
                targetSystemNotes = $Msg
                sysName = "NA"
                isIntendedSystem = "False"
                result = $result
            }
            $result = Redo-TargetSystemResult @TSRSplat
            return $result
        }
    #endregion
    #region Verify User
        IF($usrID){
            #region Attempting User Verification
                Write-Verbose -Message "Attempting To Verify User $usrID Is On $sysID now."
                TRY{
                    TRY{
                        $loggedOnUser = (Get-CimInstance @cimInstSplat @stopSplat |
                            Select-Object -ExpandProperty Username).Split("\") |
                            Select-Object -Last 1
                    }
                    CATCH{
                        TRY{
                            $loggedOnUser = (Get-CimInstance @cimInstAlt @stopSplat |
                                Select-Object -ExpandProperty Username).Split("\") |
                                Select-Object -Last 1
                        }
                        CATCH{
                            throw $_.Exception.Message
                        }
                    }
                    $loggedOnUser = (Get-CimInstance @cimInstSplat @stopSplat |
                        Select-Object -ExpandProperty Username).Split("\") |
                        Select-Object -Last 1
                    $result = Redo-TargetSystemResult -usrLoggedOn $loggedOnUser -result $result
                    IF($usrID -eq ($result.usrLoggedOn)){
                        $result = Redo-TargetSystemResult -isIntendedUser True -result $result
                    }
                    ELSE{
                        $result = Redo-TargetSystemResult -isIntendedUser False -result $result
                    }
                }
                CATCH{
                    $Msg = -JOIN (
                        "$($_.Exception.GetType().FullName)`n",
                        "$($_.Exception.Message)"
                    )
                    $TSRSplat = @{
                        targetSystemNotes = $Msg
                        sysName = "NA"
                        isIntendedSystem = "False"
                        result = $result
                    }
                    $result = Redo-TargetSystemResult @TSRSplat
                    return $result
                }
            #endregion
        }
    #endregion
    #region Closure
        Get-PSSession | Remove-PSSession *>$null
        $result = Redo-TargetSystemResult -targetSystemNotes NA -result $result
        return $result
    #endregion
}
Export-ModuleMember -Function Confirm-TargetSystem
