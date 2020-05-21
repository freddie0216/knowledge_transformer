# to work around unworking Move-PnPFile
function Move-KMFile($source_url, $target_lib_name, $file_name, $credential){

    Write-Output "downloading file from $source_url..."
    Connect-PnPOnline -Url https://frogdesign.sharepoint.com/sites/KnowledgeManagement -Credentials $user_credential
    Get-PnPFile -Url $source_url -Path C:\Users\freddie.dong\tmp -FileName $file_name -AsFile
    Write-Output "downloaded!"
    Disconnect-PnPOnline

    $file_path = "C:\Users\freddie.dong\tmp\$file_name"

    Write-Output "uploading $file_name to km3..."
    Connect-PnPOnline -Url https://frogdesign.sharepoint.com/sites/Knowledge -Credentials $credential
    Add-PnPFile -Path $file_path -Folder $target_lib_name
    Write-Output "uploaded!"
    Disconnect-PnPOnline

    #delete the source file
    Remove-Item -Path $file_path
}

$user_credential = Get-Credential

Connect-PnPOnline -Url https://frogdesign.sharepoint.com/sites/Knowledge -Credentials $user_credential

# term store in km3
$term_studios = Get-PnPTerm -TermGroup "Knowledge" -TermSet "Studios"
$term_capabilities = Get-PnPTerm -TermGroup "Knowledge" -TermSet "Capabilities"
$term_clients = Get-PnPTerm -TermGroup "Knowledge" -TermSet "Clients"
$term_confidentiality = Get-PnPTerm -TermGroup "Knowledge" -TermSet "Confidentiality"
$term_industries = Get-PnPTerm -TermGroup "Knowledge" -TermSet "Industries"
$term_keywords = Get-PnPTerm -TermGroup "Knowledge" -TermSet "Keywords"
$term_years = Get-PnPTerm -TermGroup "Knowledge" -TermSet "Years"

Disconnect-PnPOnline

Connect-PnPOnline -Url https://frogdesign.sharepoint.com/sites/KnowledgeManagement -Credentials $user_credential

$original_case_study_library = Get-PnPList -Identity "Case Study"
$converted_case_study_library = Get-PnPList -Identity "PDF Conversion"

# keynote, powerpoint items
$original_case_study_docs = Get-PnPListItem -List $original_case_study_library

# converted pdf items
$converted_case_study_docs = Get-PnPListItem -List $converted_case_study_library

# populate data, 11658 is a sample assest where studio is multiple valued. after testing, loop logic will be added around.
$case_study_togo = $original_case_study_docs | Where-Object{$_.Id -eq 2247}

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
$coverted_case_study = Get-PnPListItem -List $converted_case_study_library  -Query "<View Scope='RecursiveAll'><Query><Where><Contains><FieldRef Name='FileLeafRef'/><Value Type='Computed'>$file_name_without_extension</Value></Contains></Where></Query></View>"

if(!$converted_case_study){
    if($coverted_case_study.Length -gt 1){
        Write-Output "found multiple matched results, skipping..."
        continue
    }

    Write-Output "moving $file_name_without_extension"
    $converted_case_study_url = $coverted_case_study.FieldValues.FileRef
    $converted_case_study_file_name = $coverted_case_study.FieldValues.FileLeafRef
    $converted_case_study_client_id = $term_clients | where-object{$_.Name -eq $original_case_study_client}
    
    # todo: industry does not pretty match with the new term store in km 3.0
    # $converted_case_study_industry_id = $term_industries | Where-Object{$_.Name -eq $original_case_study_industry}

    # capability is multiple valued
    $converted_case_study_capability_id_list = ""
    for($i = 0;$i -lt $original_case_study_capability.Length; $i++){
        $tmp_capability = $term_capabilities | Where-Object{$_.Name -eq $original_case_study_capability[$i]}
        $tmp_capability_id = $tmp_capability.Id
        $converted_case_study_capability_id_list = "$converted_case_study_capability_id_list, '$tmp_capability_id'"
    }
    $converted_case_study_capability_id_list = $converted_case_study_capability_id_list.Substring(1, $converted_case_study_capability_id_list.Length-1)

    # studio is multiple valued
    $converted_case_study_studio_id_list = ""
    for($i = 0;$i -lt $original_case_study_studio.Length; $i++){
        $tmp_studio = $term_studios | Where-Object{$_.Name -eq $original_case_study_studio[$i]}
        $tmp_studio_id = $tmp_studio.Id
        $converted_case_study_studio_id_list = "$converted_case_study_studio_id_list, '$tmp_studio_id'"
    }
    $converted_case_study_studio_id_list = $converted_case_study_studio_id_list.Substring(1, $converted_case_study_studio_id_list.Length-1)

    # todo: better to populate keyword metadata from km2.0 to km3.0 first, i will develop it in another script

    $converted_case_study_year = $term_years | Where-Object{$_.Name -eq $original_case_study_year}

    # todo: confidentiality does not pretty match with the new term store in km 3.0
    # $converted_case_study_confidentiality = $term_confidentiality | Where-Object{$_.Name -eq $origina}

    # populate case study item on km 3.0
    # copy file has 200 mb file size limit
    Write-Output "populating $original_case_study_file_name into case studies on km 3.0..."
    # Write-Output "executing: Move-PnPFile -ServerRelativeUrl $converted_case_study_url -TargetUrl /sites/Knowledge/Case Studies/$converted_case_study_file_name -OverwriteIfAlreadyExists -Force"

    # Move-PnPFile -ServerRelativeUrl $converted_case_study_url -TargetUrl "/sites/Knowledge/Case Studies/$converted_case_study_file_name" -OverwriteIfAlreadyExists -Force
    Move-KMFile -source_url $converted_case_study_url -target_lib_name 'Case Studies'
                -file_name $converted_case_study_file_name -credential $user_credential
    Connect-PnPOnline -Url https://frogdesign.sharepoint.com/sites/Knowledge -Credentials $user_credential

    $sources_km3 = Get-PnPList -Identity "Sources"
    $case_study_library_km3 = Get-PnPList -Identity "Case Studies"
    $case_study_km3 = Get-PnPListItem -List $case_study_library_km3 -Query "<View Scope='RecursiveAll'><Query><Where><Contains><FieldRef Name='FileLeafRef'/><Value Type='Computed'>$converted_case_study_file_name</Value></Contains></Where></Query></View>"

    Write-Output "updating fields..."
    Set-PnPListItem -List $case_study_library_km3 -Identity $case_study_km3.Id -Values @{
        "Source" = "https://frogdesign.sharepoint.com/sites/Knowledge/Sources/$original_case_study_file_name?web=1, $original_case_study_file_name";
        "Client" = $converted_case_study_client_id;
        # "Industries" = ;
        "Capabilities" = $converted_case_study_capability_id_list;
        # "Keywords_x0020__x0028_frog_x0020_Knowledge_x0029_":;
        "Studios" = $converted_case_study_studio_id_list;
        "Year" = $converted_case_study_year
        # "Confidentiality":;
        # "Contacts":;
        # "Video":
    }
    Write-Output "done!"

    # populate source item on km 3.0
    # copy file has 200 mb file size limit
    Write-Output "populating $converted_case_study_file_name into sources on km 3.0..."
    # Move-PnPFile -ServerRelativeUrl $original_case_study_url -TargetUrl "/sites/Knowledge/Sources/$original_case_study_file_name" -OverwriteIfAlreadyExists -Force
    Move-KMFile -source_url $original_case_study_url -target_lib_name 'Sources'
                -file_name $original_case_study_file_name -credential $user_credential
    $sources_case_study_km3 = Get-PnPListItem -List $sources_km3 -Query "<View Scope='RecursiveAll'><Query><Where><Contains><FieldRef Name='FileLeafRef'/><Value Type='Computed'>$original_case_study_file_name</Value></Contains></Where></Query></View>"

    Write-Output "updating fields..."
    Set-PnPListItem -List $sources_km3 -Identity $sources_case_study_km3.Id -Values @{
        "Reference" = "https://frogdesign.sharepoint.com/sites/Knowledge/Case Studies/$converted_case_study_file_name, $converted_case_study_file_name"
    }

    Disconnect-PnPOnline

}
else{
    Write-Output "no converted pdf found, please convert the original file: $original_case_study_file_name, skipping..."
    continue
}


