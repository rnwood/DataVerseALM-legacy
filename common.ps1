$ErrorActionPreference = "Stop"

$commonroot = (split-path $MyInvocation.MyCommand.Source)
write-verbose "Root: $commonroot"

. $commonroot\settings.ps1

function restoreModules() {


    [bool] $packageschanged = $true

    if (test-path $commonroot\.Modules.scripts\this.config) {
        Write-Verbose "Modules folder exists. Checking for modules.config changes"
        $packageschanged = (get-content -raw $commonroot\modules.scripts.config) -ne (get-content -raw $commonroot\.modules.scripts\this.config)
    }

    if ($packageschanged) {

        Write-Verbose "Modules folder update needed. Restoring modules"

        if (test-path $commonroot\.Modules.scripts) {
            get-childitem $commonroot\.Modules.scripts -Recurse | remove-item -force -Recurse
            remove-item -Recurse -force $commonroot\.Modules.scripts
        }

        new-item $commonroot\.Modules.scripts -type Directory | out-null

        get-content $commonroot\modules.scripts.config | foreach-object {
            $modulename, $version = $_.Split(" ")
            save-module -Name $modulename -RequiredVersion $version -path $commonroot\.Modules.scripts
        }

        copy-item $commonroot\modules.scripts.config $commonroot\.modules.scripts\this.config
        Write-Verbose "Modules folder update completed ok"
    } else {
        Write-Verbose "Modules folder is already up to date"
    }

    get-content $commonroot\modules.scripts.config | foreach-object {
        $modulename, $version = $_.Split(" ")
        
        resolve-path ("$commonroot\.modules.scripts\$($modulename)*") | import-module
    }

}

function restoreNugetPackages() {

    [bool] $packageschanged = $true

    if (test-path $commonroot\.Packages.scripts\this.config) {
        Write-Verbose "Packages folder exists. Checking for packages.config changes"
        $packageschanged = (get-content -raw $commonroot\packages.scripts.config) -ne (get-content -raw $commonroot\.Packages.scripts\this.config)
    }

    if ($packageschanged) {

        Write-Verbose "Packages folder update needed. Restoring packages"

        if (test-path $commonroot\.Packages.scripts) {
            get-childitem $commonroot\.Packages.scripts -Recurse | remove-item -force -Recurse
            remove-item -Recurse -force $commonroot\.Packages.scripts
        }

        new-item $commonroot\.Packages.scripts -type Directory | out-null

        invoke-webrequest "https://dist.nuget.org/win-x86-commandline/latest/nuget.exe" -outfile $commonroot\.packages.scripts\nuget.exe
        
        & $commonroot\.packages.scripts\nuget.exe restore -packagesdirectory $commonroot\.Packages.scripts $commonroot\packages.scripts.config
        if ($LASTEXITCODE -ne 0) {
            throw "Nuget package restore failed"
        }

        copy-item $commonroot\packages.scripts.config $commonroot\.Packages.scripts\this.config
        Write-Verbose "Packages folder update completed ok"
    } else {
        Write-Verbose "Packages folder is already up to date"
    }

    
}

function getConfigValue([Parameter(Mandatory)][string] $key) {

    $file = "$commonroot\.user.config\$key"

    if (-not (test-path $file)) {
        $result = read-host -prompt "Enter value for $key"
        $result | out-file -encoding UTF8 $file
    } else {
        $result = get-content -encoding UTF8 $file
    }

    write-output $result
}

function publishCrmCustomizations([Parameter(Mandatory)][Microsoft.Xrm.Tooling.Connector.CrmServiceClient] $connection) {

	write-verbose "Publishing customizations..."

	$request = new-object Microsoft.Crm.Sdk.Messages.PublishAllXmlRequest
	$connection.Execute($request) | out-null

}


function loadDependencies() {
    
    [Reflection.Assembly]::LoadWithPartialName("System.IO.Compression") | out-null
    [Reflection.Assembly]::LoadWithPartialName("System.IO.Compression.FileSystem") | out-null    
    
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12;
    
    import-module $commonroot\.Modules.scripts\CredentialManager\*\CredentialManager.psd1

    Add-Type -assemblyname presentationframework

    if (-not (test-path "$commonroot\.user.config")) {
        new-item -type Directory "$commonroot\.user.config"
    }
}

function getCredentials($url) {

    $result = get-storedcredential -Type Generic -target $url
    if (-not $result) {
        $result = Get-Credential -Message "Enter credentials for $url"

        if ([System.Windows.MessageBox]::Show("Store the credentials for $url in Windows Credential Manager?",'Store Credentials Securely','YesNo','Question') -eq "Yes") {
            new-storedcredential -type Generic -persist LocalMachine -target $url -credentials $result | out-null
        }
    }

    write-output $result

}


function getCrmConnection($url) {

    $cred = getCredentials $url
    $netcred = $cred.GetNetworkCredential()
    $result = (get-crmconnection -ConnectionString "AuthType=OAuth;Username=$($netcred.Username);Password=$($netcred.Password);Url=$url;AppId=2ad88395-b77d-4561-9441-d0e40824f9bc;RedirectUri=app://5d3e90d6-aa8e-48a8-8f2c-58b45cc67315;TokenCacheStorePath=$commonroot\.user.config\tokens;LoginPrompt=Auto"-MaxCrmConnectionTimeOutMinutes 10)

    return $result
}



function getCrmSolutionFileInfo([Parameter(Mandatory)][string] $solutionfile) {

    $zipfile = [System.IO.Compression.ZipFile]::OpenRead($solutionfile)

    try {

            $reader = new-object System.IO.StreamReader] ($zipfile.GetEntry("solution.xml").Open())
            try {
                [xml] $solutionxml = $reader.ReadToEnd()
                return $solutionxml.ImportExportXml.SolutionManifest
            } finally {
                $reader.Dispose()
            }

    } finally {
        $zipfile.Dispose()
    }

}


function runSolutionPackager([string[]] $arguments) {
    $solutionpackager = resolve-path $commonroot\.packages.scripts\Microsoft.CrmSdk.CoreTools.*\content\bin\coretools\SolutionPackager.exe
    write-verbose "Solution Packager Executable: $solutionpackager"


    Write-Verbose "Running $solutionpackager $arguments"

    $output 
    (& $solutionpackager $arguments 2>&1) | tee-object -Variable output
    if ($LASTEXITCODE -ne 0) {
        throw "Solution packager failed. Exit code: $LASTEXITCODE"
    }

    Write-Verbose "Solution packager executed OK"


}


function packCrmSolution([Parameter(Mandatory)][string] $folder, [Parameter(Mandatory)][string] $zipfile, [Switch] $managed) {
    $packagetype = "Unmanaged"
    if ($managed) {
        $packagetype = "Managed"
    }

    runSolutionPackager -arguments "/action:Pack", "/zipfile:$zipfile", "/folder:$folder", "/packagetype:$packagetype"
}

function unpackCrmSolution([Parameter(Mandatory)][string] $folder, [Parameter(Mandatory)][string] $zipfile) {

    if (test-path $folder) {
        remove-item -force -recurse $folder
    }

    runSolutionPackager -arguments "/action:Extract", "/zipfile:$zipfile", "/folder:$folder", "/packagetype:Unmanaged", "/clobber"

    normaliseCrmSolutionFolder -solutionfolder $folder
}

function exportCrmSolution([Parameter(Mandatory)][Microsoft.Xrm.Tooling.Connector.CrmServiceClient] $connection, [Parameter(Mandatory)][string] $uniquename, [Parameter(Mandatory)][string] $solutionfile, [Switch] $managed) {
    
    Write-Verbose "Starting Export of solution $uniquename to $solutionfile"

    [Microsoft.Crm.Sdk.Messages.ExportSolutionRequest] $exportrequest = new-object Microsoft.Crm.Sdk.Messages.ExportSolutionRequest
    $exportrequest.Managed = $managed
    $exportrequest.SolutionName = $uniquename
    $exportresponse = $connection.Execute($exportrequest)

    [IO.File]::WriteAllBytes($solutionfile, $exportresponse.ExportSolutionFile)

    write-verbose "Export completed ok"
}    

function importCrmSolution([Parameter(Mandatory)][Microsoft.Xrm.Tooling.Connector.CrmServiceClient] $connection, [Parameter(Mandatory)][string] $solutionfile, [switch] $managed, [switch] $skipmanagedpromote) {
    
    $solutionfileinfo = getCrmSolutionFileInfo -solutionfile $solutionfile
    $uniquename = $solutionfileinfo.UniqueName

    $installedsolution = (getCrmSolutionInfo -connection $connection -uniquename $uniquename)
    $installedholdingsolution = (getCrmSolutionInfo -connection $connection -uniquename "${uniquename}_Upgrade")

    if ($installedsolution -and $installedsolution["ismanaged"] -ne $managed) {
        throw "Attempting to deploy managed solution to env that has unmanaged solution. If this is intended for conversion, delete the unmanaged solution first."
    }

    if ($installedholdingsolution -and $installedholdingsolution["version"] -ne $solutionfileinfo.Version) {
        throw "Error - Holding solution ${uniquename}_Upgrade is already present with non-matching version no $($installedholdingsolution["version"]). Promote or delete (warning: potential data loss) it and try again."
    }

    $ismanagedupgrade = $managed -and $installedsolution

    if ($managed -and $installedsolution -and $installedsolution["version"] -eq $solutionfileinfo.Version) {
        write-verbose "Skipping import of solution $solutionfile as version $($solutionfileinfo.Version) is already installed"
    } elseif ($ismanagedupgrade -and $installedholdingsolution -and $installedholdingsolution["version"] -eq  $solutionfileinfo.Version) {
        write-verbose "Skipping import of solution $solutionfile as holding solution version $($solutionfileinfo.Version) is already installed"
    } else {

        Write-Verbose "Starting Import of solution $solutionfile - managed:$managed managedupgrade:$ismanagedupgrade"

        [Microsoft.Crm.Sdk.Messages.ImportSolutionRequest] $importrequest = new-object Microsoft.Crm.Sdk.Messages.ImportSolutionRequest
        $importrequest.CustomizationFile = [IO.File]::ReadAllBytes($solutionfile)
        $importrequest.ImportJobId = [Guid]::NewGuid()
        $importrequest.PublishWorkflows = $true
        $importrequest.OverwriteUnmanagedCustomizations = $false
        $importrequest.HoldingSolution  = $ismanagedupgrade
        $importrequest.ConvertToManaged = $managed

        [Microsoft.Xrm.Sdk.Messages.ExecuteAsyncRequest] $asyncrequest = new-object Microsoft.Xrm.Sdk.Messages.ExecuteAsyncRequest
        $asyncrequest.Request = $importrequest

        
        $asyncresponse = $connection.Execute($asyncrequest)
        $asyncjobid = $asyncresponse.AsyncJobId


        $completed=$false;
        $errordetail=$null

        while(-not $completed) {
            start-sleep -Seconds 1
            $asyncjob = $connection.Retrieve("asyncoperation", $asyncjobid, (new-object Microsoft.Xrm.Sdk.Query.ColumnSet $true))

            $asyncjobstatuscode = $asyncjob["statuscode"].Value
            Write-Verbose "Waiting for import job to complete. Job status code: $asyncjobstatuscode"
            
            switch ($asyncjobstatuscode) {
                30 {
                    $completed = $true;
                }

                31 {
                    $completed = $true
                    $errordetail = $asyncjob["message"]
                }

                32 {
                    throw "Solution import cancelled"
                }

                10 {
                    write-verbose "Waiting for execution"
                }

                0 {
                    write-verbose "Waiting for execution"
                }

                20 {
                    try {
                        $importjob = $connection.Retrieve("importjob", $importrequest.ImportJobId, (new-object Microsoft.Xrm.Sdk.Query.ColumnSet $true))
                        write-verbose "Progress: $($importjob["progress"])%"
                    } catch {
                        write-verbose "Progress: unknown (importjob not found)"
                    }
                }
            }
        }

        try {
        $importjob = $connection.Retrieve("importjob", $importrequest.ImportJobId, (new-object Microsoft.Xrm.Sdk.Query.ColumnSet $true))
            [xml] $results = $importjob["data"]
            $importerrordetails = $results.SelectSingleNode("//*[@result='failure']")
            if ($importerrordetails) {
                $errordetail = $importerrordetails.ParentNode.OuterXml
            }
        } catch {
            
        }

        if ($errordetail) {
            throw "Solution import error:`n" + $errordetail
        }
    }

    $installedholdingsolution = (getCrmSolutionInfo -connection $connection -uniquename "${uniquename}_Upgrade")

    if ($installedholdingsolution -and (-not $skipmanagedpromote)) {
        promoteCrmSolution -connection $connection -uniquename $uniquename
    }
    
}

function promoteCrmSolution([Parameter(Mandatory)][Microsoft.Xrm.Tooling.Connector.CrmServiceClient] $connection, [Parameter(Mandatory)][string] $uniquename, [Switch]$allowmissing) {

    $installedholdingsolution = (getCrmSolutionInfo -connection $connection -uniquename "${uniquename}_Upgrade")
    if ($allowmissing -and (-not $installedholdingsolution)) {
        write-verbose "Skipping promote of solution $uniquename - no pending upgrade"
        return;
    }

    write-verbose "Promoting solution"
    $promoterequest = new-object Microsoft.Crm.Sdk.Messages.DeleteAndPromoteRequest
    $promoterequest.UniqueName = $uniquename
    $connection.Execute($promoterequest)
}

function getCrmSolutionInfo([Parameter(Mandatory)][Microsoft.Xrm.Tooling.Connector.CrmServiceClient] $connection, [Parameter(Mandatory)][string] $uniquename) {

    $query = new-object Microsoft.Xrm.Sdk.Query.QueryExpression "solution"
    $query.Criteria.AddCondition("uniquename", [Microsoft.Xrm.Sdk.Query.ConditionOperator]::Equal, $uniquename)
    $query.ColumnSet = new-object Microsoft.Xrm.Sdk.Query.ColumnSet $true

    return $connection.RetrieveMultiple($query).Entities[0]

}

function importManagedCrmSolution([Parameter(Mandatory)][Microsoft.Xrm.Tooling.Connector.CrmServiceClient] $connection, [Parameter(Mandatory)][string] $solutionfile, [Switch] $force) {
    $solutionfileinfo = getCrmSolutionFileInfo -solutionfile $solutionfile
    $solutioninfo = getCrmSolutionInfo -connection $connection -uniquename $solutionfileinfo.UniqueName

    write-verbose "Solution file version: $($solutionfileinfo.Version)"
    write-verbose "Installed version: $(if ($solutioninfo) { $solutioninfo["version"] })"

    if ($force -or (-not $solutioninfo) -or $solutioninfo["version"] -ne $solutionfileinfo.Version) {
        importCrmSolution -connection $connection -solutionfile $solutionfile
    }
}

function normaliseCrmSolutionFolder([Parameter(Mandatory)][string] $solutionfolder) {
    
    write-verbose "Normalising CRM solution folder $solutionfolder"
    
    $xmlsettings = new-object System.Xml.XmlWriterSettings
    $xmlsettings.Indent=$true
    $xmlsettings.Encoding=[Text.Encoding]::UTF8

    Get-ChildItem $solutionfolder -recurse -Include *.xml | foreach-object {
        $xml = [xml] (get-content -Encoding UTF8 $_.FullName -raw)


        $commonrootComponentElement = $xml.SelectNodes("/ImportExportXml/SolutionManifest/RootComponents");

        if ($commonrootComponentElement.Count) {
            $newNodes = $commonrootComponentElement.ChildNodes | sort-object -property type,schemaName,parentId,id
            $xml.SelectNodes("/ImportExportXml/SolutionManifest/RootComponents/RootComponent")| foreach-object {
                $_.ParentNode.RemoveChild($_) | out-null
            }
            $newnodes | foreach-object {
                $commonrootComponentElement.AppendChild($_) | out-null
            }
        }

        $xml.SelectNodes("//AppModuleRoleMaps") | foreach-object {
            $rolemapselement = $_
            $newNodes = $rolemapselement.ChildNodes | sort-object -property id
            $newNodes | foreach-object {
                $_.ParentNode.RemoveChild($_) | out-null
            }
            $newnodes | foreach-object {
                $rolemapselement.AppendChild($_) | out-null
            }
        }
    
        $xml.SelectNodes("//MissingDependencies/MissingDependency")| foreach-object {
            $_.ParentNode.RemoveChild($_) | out-null
        }

        $xml.SelectNodes("//EntityMaps/EntityMap[count(*)=0]")| foreach-object {
            $_.ParentNode.RemoveChild($_) | out-null
        }
    
        $xml.SelectNodes("//EntityInfo//attribute/Format[not(text())]")| foreach-object {
            $_.ParentNode.RemoveChild($_) | out-null
        }
        
        $xml.SelectNodes("//systemform/form//row[count(*)=0]")| foreach-object {
            $_.InnerText = ""
        }


        $xml.SelectNodes("//*")| foreach-object {

            $attributes = $_.GetType().GetProperty("Attributes").GetValue($_)

            $sortedattributes = [object[]] ( $attributes | Sort-ObJect {$_.GetType().GetProperty("LocalName").GetValue($_) } )

            $attributes.RemoveAll()

            foreach($attribute in $sortedattributes) {
                $attributes.Append($attribute) | out-null;
            }
        }

        $writer = [System.Xml.XmlWriter]::Create($_.FullName, $xmlsettings)
        try {
            $xml.WriteContentTo($writer)
        } finally {
            $writer.Dispose()
        }
    }

}

function importCrmData([Parameter(Mandatory)][Microsoft.Xrm.Tooling.Connector.CrmServiceClient] $connection, [Parameter(Mandatory)][string] $datadir) {

    Write-Verbose "Importing CRM data from $datadir"

    $buquery=new-object Microsoft.Xrm.Sdk.Query.QueryExpression "businessunit"
    $buquery.Criteria.AddCondition("parentbusinessunitid", "Null")
    $rootburef = [Microsoft.Xrm.Sdk.EntityReference] $connection.RetrieveMultiple($buquery).Entities[0].ToEntityReference()

    $config = get-content -raw $datadir/config.json | ConvertFrom-Json
    
    foreach($entity in $config.entities) {
        write-host "Importing $($entity.name) data"
        $em = $connection.GetEntityMetadata($entity.Name, "All")

        get-childitem "$datadir/data/$($entity.name)" -filter *.json | foreach-object {

            write-host ""
            $recorddata = get-content -raw $_.fullname | convertfrom-json

            write-host "Importing entity $($entity.name) record id $($recorddata.Id)"

            try {

                if ($em.IsIntersect) {

                    $relationship = new-object Microsoft.Xrm.Sdk.Relationship $em.ManyToManyRelationships[0].SchemaName
                    $collection = new-object Microsoft.Xrm.Sdk.EntityReferenceCollection
                    $collection.AddRange((new-object Microsoft.Xrm.Sdk.EntityReference $em.ManyToManyRelationships[0].Entity2LogicalName, ([guid] $recorddata."$($em.ManyToManyRelationships[0].Entity2IntersectAttribute)")))

                    try {
                    $connection.Associate($em.ManyToManyRelationships[0].Entity1LogicalName, ([guid] $recorddata."$($em.ManyToManyRelationships[0].Entity1IntersectAttribute)"), $relationship, $collection)
                    } catch {
                        if (-not $_.ToString().Contains("Cannot insert duplicate key")) {
                            throw $_
                        }
                    }

                } else {

                    $record = new-object Microsoft.Xrm.Sdk.Entity $entity.name, ([Guid]$recorddata.Id)

                    $recorddata | get-member -type NoteProperty | where-object {$_.Name -ne "Id" -and $_.Name -ne "statuscode" -and $_.Name -ne "statecode"} | foreach-object {
                        $field = $_
                        $value = $recorddata.($field.Name)
                        $fm = $em.Attributes | where-object {$_.LogicalName -eq $field.Name}

                        if (-not $fm) {
                            write-error "Attribute not found in metadata $($field.Name)"
                        }

                        if ($value -eq $null) {
                            $value = $null
                        } elseif ($fm.LogicalName -eq "businessunitid") {
                            $value = $rootburef
                        } elseif ($fm.AttributeType -eq "Decimal") {
                            $value = [decimal] $value
                        } elseif ($fm.AttributeType -eq "Integer") {
                            $value = [int] $value
                        } elseif ($fm.AttributeType -eq "String") {
                            $value = [string] $value
                        } elseif ($fm.AttributeType -eq "UniqueIdentifier") {
                            $value = [Guid] $value
                        } elseif ($fm.AttributeType -eq "Boolean") {
                            $value = [bool] $value
                        } elseif ($fm.AttributeType -eq "Picklist") {
                            $intvalue = [int] $value
                            $value = [ Microsoft.Xrm.Sdk.OptionSetValue](new-object Microsoft.Xrm.Sdk.OptionSetValue $intvalue)
                        }elseif ($fm.AttributeType -eq "Lookup") {
                            $value = [Microsoft.Xrm.Sdk.EntityReference] (new-object Microsoft.Xrm.Sdk.EntityReference $value.LogicalName, ([Guid]$value.Id))
                        }else {
                            write-error "Unhandled attribute type $($fm.AttributeType) for attribute $($fm.LogicalName)"
                        }

                        $record[$fm.LogicalName] = $value

                        if ($value -ne $null) {
                            write-host "Field $($field.Name) type $($fm.AttributeType) Value: $($value) type: $($value.GetType())"
                        } else {
                            write-host "Field $($field.Name) type $($fm.AttributeType) Value <null>"
                        }
                    }

                    
                    $existingrecord = $null;
                    
                    try {
                        $existingrecord = $connection.Retrieve($record.LogicalName, $record.Id, (new-object Microsoft.Xrm.Sdk.Query.ColumnSet $true))
                    } catch {
                        if (-not $_.ToString().Contains("Does Not Exist")) {
                            throw $_
                        }
                    }
                    if (-not $existingrecord) {
                        $connection.Create($record)
                    } else {
                        $connection.Update($record)
                    }

                    if ($recorddata.statuscode) {
                        write-host "Setting status to State:$($recorddata.statecode) Status:$($recorddata.statuscode)"
                        $recordupdate = new-object Microsoft.Xrm.Sdk.Entity $record.LogicalName, $record.Id
                        #$recordupdate["statecode"] = [ Microsoft.Xrm.Sdk.OptionSetValue](new-object Microsoft.Xrm.Sdk.OptionSetValue ([int] $recorddata.statecode))
                        $recordupdate["statuscode"] = [ Microsoft.Xrm.Sdk.OptionSetValue](new-object Microsoft.Xrm.Sdk.OptionSetValue ([int]$recorddata.statuscode))
                        $connection.Update($recordupdate)
                    
                    }
                }

            } catch {
                write-error ("Error importing record $($recorddata.Id): " + $_.ToString())
            }
        }
    }
    
}

function exportCrmData([Parameter(Mandatory)][Microsoft.Xrm.Tooling.Connector.CrmServiceClient] $connection, [Parameter(Mandatory)][string] $datadir) {

    Write-Verbose "Exporting CRM data to $datadir"

    if (test-path "$datadir/data") {
        remove-item -recurse "$datadir/data"
    }
	
    $config = get-content -raw $datadir/config.json | ConvertFrom-Json
    
    foreach($entity in $config.entities) {
       
        write-host "Exporting $($entity.name) data"
        $em = $connection.GetEntityMetadata($entity.Name, "All")
        new-item -type directory "$datadir/data/$($entity.name)"



        if ($entity.fetchxmlbody) {
            $fetchconvertrequest = new-object Microsoft.Crm.Sdk.Messages.FetchXmlToQueryExpressionRequest
            $fetchconvertrequest.FetchXml = "<fetch><entity name='$($entity.name)' mapping='logical'>$($entity.fetchxmlbody)</entity></fetch>"
            $query = $connection.Execute($fetchconvertrequest).Query
        } else {
            $query = new-object Microsoft.Xrm.Sdk.Query.QueryExpression $entity.name
        }

        $query.ColumnSet = new-object Microsoft.Xrm.Sdk.Query.ColumnSet $true

        $connection.RetrieveMultiple($query).Entities | foreach-object {
            $values = @{Id=$_.Id}
            foreach($attribute in $_.Attributes) {
                if ($entity.excludefields -and $entity.excludefields.Contains($attribute.Key)) {
                    continue;
                }

                if ($attribute.Key -eq "businessunitid") {
                    $values[$attribute.Key] = "<Root BU>"
                } elseif ($attribute.Value -is [Microsoft.Xrm.Sdk.OptionSetValue]) {
                    $values[$attribute.Key] = $attribute.Value.Value
                } elseif ($attribute.Value -is [Microsoft.Xrm.Sdk.EntityReference]) {
                    $values[$attribute.Key] = new-object PSObject -property @{Id=$attribute.Value.Id; LogicalName=$attribute.Value.LogicalName}; 
                }else {
                    $values[$attribute.Key] = $attribute.Value
                }
            }

            $filename = $_.Id

            if ($em.IsIntersect) {
                $filename = $_[$em.ManyToManyRelationships[0].Entity1IntersectAttribute].ToString() + "-" + $_[$em.ManyToManyRelationships[0].Entity2IntersectAttribute].ToString()

                $values.Remove("Id")
                $values.Remove($em.PrimaryIdAttribute)
            }

            new-object PSObject -property $values | convertto-json | out-file -encoding UTF8 "$datadir/data/$($entity.name)/${filename}.json"
        }
    }
}

function queryRecordsByAttributes($connection, [string] $entityname, [Hashtable] $values) {
    $query = new-object Microsoft.Xrm.Sdk.Query.QueryByAttribute($entityname)
    $query.ColumnSet = new-object Microsoft.Xrm.Sdk.Query.ColumnSet $true
    
    $values.Keys | foreach-object {
        $query.AddAttributeValue($_, $values[$_]);
    }
    

    $connection.RetrieveMultiple($query).Entities | write-output
}

function queryRecordByAttributes($connection, [string] $entityname, [Hashtable] $values) {
    $results = @(queryRecordsByAttributes -connection $connection -entityname $entityname -values $values)
    if ($results.Count -gt 1) {
        throw "More than one existing record found"
    }

    return $results[0]
}



restoreModules
restoreNugetPackages
loadDependencies