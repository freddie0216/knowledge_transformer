[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client")
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime")

# to work around uploading file size limitation of Add-PnPFile
function Add-KMFile($file_path, $target_site_url,  $target_lib_name, $credential) {
    $ctx = New-Object Microsoft.SharePoint.Client.ClientContext($target_site_url)
    $ctx.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($credential.UserName, $credential.Password)

    try {
        $web = $ctx.Web
        $ctx.Load($web)
        $ctx.ExecuteQuery()

        $fs = New-Object System.IO.FileStream($file_path, [System.IO.FileMode]::Open)
        $des_url = ("{0}/{1}/{2}" -f $web.ServerRelativeUrl.TrimEnd("/"), $target_lib_name, [System.IO.Path]::GetFileName($file_path), $fs, $true)
        [Microsoft.SharePoint.Client.File]::SaveBinaryDirect($ctx, $des_url, $fs, $true)
    }
    catch {
        Write-Host -f Red "error happened in uploading file!" $_.Exception.Message
    }
    finally {
        $fs.Dispose()
        $ctx.Dispose()
    }
}

function Set-CheckedInKMFile($item_id, $is_major_checkin, $credential){
    $ctx = New-Object Microsoft.SharePoint.Client.ClientContext($site_url)
    $ctx.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($credential.UserName, $credential.Password)

    try {
        $web = $ctx.Web
        $ctx.Load($web)
        $item=$web.Lists.GetByTitle($lib_name).GetItemById($ItemID)
        $file=$item.File
        $ctx.Load($file)
        $ctx.ExecuteQuery()

        if ($null -eq $file.CheckOutType) {
            Write-Host "file is not checked out..." -f Yellow
        }
        else {
            if ($is_major_checkin) {
                $file.CheckIn("", [Microsoft.SharePoint.Client.CheckinType]::MajorCheckIn)
            }
            else {
                $file.CheckIn("", [Microsoft.SharePoint.Client.CheckinType]::MinorCheckIn)
            }
            
            $ctx.ExecuteQuery()
        }
    }
    catch {

    }
    finally {
        $ctx.Dispose()
    }

}

# the copied file is checked out, need to check-in after the execution
function Copy-KMFile($source_url, $source_site_url, $target_site_url, $target_lib_name, $credential) {
    $src_ctx = New-Object Microsoft.SharePoint.Client.ClientContext($source_site_url)
    $tar_ctx = New-Object Microsoft.SharePoint.Client.ClientContext($target_site_url)

    $tar_ctx.RequestTimeout = -1

    try {
        $src_ctx.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($credential.UserName, $credential.Password)
        $tar_ctx.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($credential.UserName, $credential.Password)

        $spo_file_copy = $src_ctx.Web.GetFileByServerRelativeUrl($source_url)
        $src_ctx.Load($spo_file_copy)
        $src_ctx.ExecuteQuery()

        $tar_web = $tar_ctx.Web
        $tar_ctx.Load($tar_web)
        $tar_ctx.ExecuteQuery()

        $des_url = ("{0}/{1}/{2}" -f $tar_web.ServerRelativeUrl.TrimEnd("/"), $target_lib_name, $spo_file_copy.Name)
        
        if($src_ctx.HasPendingrequest) {
            $src_ctx.ExecuteQuery()
        }
        $file_info = [Microsoft.SharePoint.Client.File]::OpenBinaryDirect($src_ctx, $spo_file_copy.ServerRelativeUrl)
        [Microsoft.SharePoint.Client.File]::SaveBinaryDirect($tar_ctx, $des_url, $file_info.Stream, $true)
    }
    catch {
        Write-Host -f Red "error happened in uploading file!" $_.Exception.Message
    }
    finally {
        if ($null -eq $src_ctx) {
            $src_ctx.Dispose()
        }
        if ($null -eq $tar_ctx){
            $tar_ctx.Dispose()
        }
    }
}

# pnp move cannot be used in this scenario as the source will disappear
function Move-KMFile($source_url, $target_lib_name, $file_name, $conn_source, $target_credential) {
    try {
        Write-Output "downloading file from $source_url..."
        Get-PnPFile -Url $source_url -Path C:\Users\freddie.dong\tmp -FileName $file_name -AsFile -Connection $conn_source
        Write-Output "downloaded!"

        $file_path = "C:\Users\freddie.dong\tmp\$file_name"

        Write-Output "uploading $file_name to km3..."
        # $new_file = Add-PnPFile -Path $file_path -Folder $target_lib_name -Connection $conn_target
        # add-pnpfile has file size limitation to be <= 250 mb, workaround it by using customized upload
        Add-KMFile -file_path $file_path -target_lib_name $target_lib_name -target_site_url "https://frogdesign.sharepoint.com/sites/knowledge" -credential $target_credential
        Write-Output "uploaded!"

        # delete the cached file
        Remove-Item -Path $file_path

        Write-Host -f Green "file moved successfully..."
    }
    catch {
        Write-Host -f Red "error happened in moving files!" $_.Exception.Message
    }
}

$user_credential = Get-Credential

$conn_k = Connect-PnPOnline -Url https://frogdesign.sharepoint.com/sites/Knowledge -Credentials $user_credential -ReturnConnection
$conn_km = Connect-PnPOnline -Url https://frogdesign.sharepoint.com/sites/KnowledgeManagement -Credentials $user_credential -ReturnConnection

# term store in km3
$term_studios = Get-PnPTerm -TermGroup "Knowledge" -TermSet "Studios" -Connection $conn_k
$term_capabilities = Get-PnPTerm -TermGroup "Knowledge" -TermSet "Capabilities" -Connection $conn_k
$term_clients = Get-PnPTerm -TermGroup "Knowledge" -TermSet "Clients" -Connection $conn_k
$term_confidentiality = Get-PnPTerm -TermGroup "Knowledge" -TermSet "Confidentiality" -Connection $conn_k
$term_industries = Get-PnPTerm -TermGroup "Knowledge" -TermSet "Industries" -Connection $conn_k
$term_keywords = Get-PnPTerm -TermGroup "Knowledge" -TermSet "Keywords" -Connection $conn_k
$term_years = Get-PnPTerm -TermGroup "Knowledge" -TermSet "Years" -Connection $conn_k

$original_case_study_library = Get-PnPList -Identity "Case Study" -Connection $conn_km
$converted_case_study_library = Get-PnPList -Identity "PDF Conversion" -Connection $conn_km

# keynote, powerpoint items
$original_case_study_docs = Get-PnPListItem -List $original_case_study_library -Connection $conn_km

# converted pdf items
$converted_case_study_docs = Get-PnPListItem -List $converted_case_study_library -Connection $conn_km

foreach($original_case_study_doc in $original_case_study_docs) {
    # populate data, 11658 is a sample assest where studio is multiple valued. after testing, loop logic will be added around.
    # $case_study_togo = $original_case_study_docs | Where-Object {$_.Id -eq 2247}
    $case_study_togo = $original_case_study_doc

    if ($case_study_togo.FileSystemObjectType -eq "Folder") {
        continue
    }

    $original_case_study_url = $case_study_togo.FieldValues.FileRef
    $original_case_study_file_name = $case_study_togo.FieldValues.FileLeafRef
    $original_case_study_client = $case_study_togo.FieldValues["Client_test"].Label
    $original_case_study_studio = $case_study_togo.FieldValues["r0bf"]
    $original_case_study_capability = $case_study_togo.FieldValues["wq4j"]
    $original_case_study_industry = $case_study_togo.FieldValues["Industry"]
    $original_case_study_industry_subtype = $case_study_togo.FieldValues["Industry_x0020_Subtype"]
    $original_case_study_keywords = $case_study_togo.FieldValues["TaxKeyword"].Label
    $original_case_study_year = $case_study_togo.FieldValues["h1k7"]
    $original_case_study_confidentiality = $case_study_togo.FieldValues["_x0066_hp8"].Label

    # get corresponding pdf 
    $file_name_without_extension = $original_case_study_file_name.Substring(0, $original_case_study_file_name.LastIndexOf('.'))
    Write-Output "searching converted pdf: $file_name_without_extension"
    $converted_case_study = Get-PnPListItem -List $converted_case_study_library -Query ("<View Scope='RecursiveAll'><Query><Where><Contains><FieldRef Name='FileLeafRef'/><Value Type='Computed'>{0}</Value></Contains></Where></Query></View>" -f $file_name_without_extension) -Connection $conn_km

    if ($null -eq $converted_case_study) {
        Write-Output "no converted pdf found, please convert the original file: $original_case_study_file_name, skipping..."
        continue
    }
    else {

        if($converted_case_study.Length -gt 1) {
            Write-Output "found multiple matched results, skipping..."
            foreach($t in $converted_case_study) {
                Write-Output $t.FieldValues.FileLeafRef
            }
            continue
        }

        Write-Output "moving $file_name_without_extension"
        $converted_case_study_url = $converted_case_study.FieldValues.FileRef
        $converted_case_study_file_name = $converted_case_study.FieldValues.FileLeafRef
        $converted_case_study_client_id = ($term_clients | where-object {$_.Name -eq $original_case_study_client}).Id.Guid
        
        # todo: industry does not pretty match with the new term store in km 3.0
        # $converted_case_study_industry_id = $term_industries | Where-Object {$_.Name -eq $original_case_study_industry}

        # capability is multiple valued
        $converted_case_study_capability_id_list = @()
        for ($i = 0;$i -lt $original_case_study_capability.Length; $i++) {
            $tmp_capability = $term_capabilities | Where-Object {$_.Name -eq $original_case_study_capability[$i]}
            $tmp_capability_id = $tmp_capability.Id.Guid
            $converted_case_study_capability_id_list += "$tmp_capability_id"
        }

        # studio is multiple valued
        $converted_case_study_studio_id_list = @()
        for ($i = 0;$i -lt $original_case_study_studio.Length; $i++) {
            $tmp_studio = $term_studios | Where-Object {$_.Name -eq $original_case_study_studio[$i]}
            $tmp_studio_id = $tmp_studio.Id.Guid
            $converted_case_study_studio_id_list += "$tmp_studio_id"
        }

        # todo: better to populate keyword metadata from km2.0 to km3.0 first, i will develop it in another script

        $converted_case_study_year_id = ($term_years | Where-Object {$_.Name -eq $original_case_study_year}).Id.Guid

        # todo: confidentiality does not pretty match with the new term store in km 3.0
        # $converted_case_study_confidentiality = $term_confidentiality | Where-Object {$_.Name -eq $origina}

        # populate case study item on km 3.0
        # copy-pnpfile has 200 mb file size limit
        Write-Output "populating $converted_case_study_file_name into case studies on km 3.0..."
        # Write-Output "executing: Move-PnPFile -ServerRelativeUrl $converted_case_study_url -TargetUrl /sites/Knowledge/Case Studies/$converted_case_study_file_name -OverwriteIfAlreadyExists -Force"

        # Move-PnPFile -ServerRelativeUrl $converted_case_study_url -TargetUrl "/sites/Knowledge/Case Studies/$converted_case_study_file_name" -OverwriteIfAlreadyExists -Force
        Copy-KMFile -source_url $converted_case_study_url -source_site_url "https://frogdesign.sharepoint.com/sites/KnowledgeManagement" -target_site_url "https://frogdesign.sharepoint.com/sites/Knowledge" -target_lib_name 'Case Studies' -credential $user_credential

        # Move-KMFile -source_url $converted_case_study_url -target_lib_name "Case Studies" -file_name $converted_case_study_file_name -conn_source $conn_km -target_credential $user_credential

        $sources_km3 = Get-PnPList -Identity "Sources" -Connection $conn_k
        $case_study_library_km3 = Get-PnPList -Identity "Case Studies" -Connection $conn_k
        $case_study_km3 = Get-PnPListItem -List $case_study_library_km3 -Query "<View Scope='RecursiveAll'><Query><Where><Contains><FieldRef Name='FileLeafRef'/><Value Type='Computed'>$converted_case_study_file_name</Value></Contains></Where></Query></View>" -Connection $conn_k

        Write-Output "updating fields..."
        $case_study_km3 = Set-PnPListItem -List $case_study_library_km3 -Identity $case_study_km3.Id -Connection $conn_k -Values @{
            "Source" = ("https://frogdesign.sharepoint.com/sites/Knowledge/Sources/{0}?web=1, {1}" -f $original_case_study_file_name, $original_case_study_file_name);
            "Client" = "$converted_case_study_client_id";
            # "Industries" = ;
            "Capabilities" = $converted_case_study_capability_id_list;
            # "Keywords_x0020__x0028_frog_x0020_Knowledge_x0029_" = ;
            "Studios" = $converted_case_study_studio_id_list;
            "Year" = "$converted_case_study_year_id"
            # "Confidentiality" = ;
            # "Contacts" = ;
            # "Video" = 
        }
        # Set-PnPFileCheckedIn -Url $case_study_km3.FieldValues.FileDirRef -CheckinType MinorCheckIn
        Set-CheckedInKMFile -item_id $case_study_km3.Id -is_major_checkin $false -credential $user_credential

        Write-Output "done!"

        # populate source item on km 3.0
        # copy file has 200 mb file size limit
        Write-Output "populating $original_case_study_file_name into sources on km 3.0..."
        # Move-PnPFile -ServerRelativeUrl $original_case_study_url -TargetUrl "/sites/Knowledge/Sources/$original_case_study_file_name" -OverwriteIfAlreadyExists -Force

        # $source_url, $source_site_url, $target_site_url, $target_lib_name, $credential
        Copy-KMFile -source_url $original_case_study_url -source_site_url "https://frogdesign.sharepoint.com/sites/KnowledgeManagement" -target_site_url "https://frogdesign.sharepoint.com/sites/Knowledge" -target_lib_name 'Sources' -credential $user_credential

        # Move-KMFile -source_url $original_case_study_url -target_lib_name "Sources" -file_name $original_case_study_file_name -conn_source $conn_km -target_credential $user_credential

        $sources_case_study_km3 = Get-PnPListItem -List $sources_km3 -Query ("<View Scope='RecursiveAll'><Query><Where><Contains><FieldRef Name='FileLeafRef'/><Value Type='Computed'>{0}</Value></Contains></Where></Query></View>" -f $original_case_study_file_name)

        Write-Output "updating fields..."
        $sources_case_study_km3 = Set-PnPListItem -List $sources_km3 -Identity $sources_case_study_km3.Id -Values @{
            "Reference" = ("https://frogdesign.sharepoint.com/sites/Knowledge/Case Studies/{0}, {1}" -f $converted_case_study_file_name, $converted_case_study_file_name)
        }
        # Set-PnPFileCheckedIn -Url $sources_case_study_km3.FieldValues.FileDirRef -CheckinType MinorCheckIn
        Set-CheckedInKMFile -item_id $sources_case_study_km3.Id -is_major_checkin $false -credential $user_credential
    }
}

Disconnect-PnPOnline -Connection $conn_k
Disconnect-PnPOnline -Connection $conn_km