$user_credential = Get-Credential

$conn_km3 = Connect-PnPOnline -Url https://frogdesign.sharepoint.com/sites/Knowledge -Credentials $user_credential -ReturnConnection
$conn_km2 = Connect-PnPOnline -Url https://frogdesign.sharepoint.com/sites/KnowledgeManagement -Credentials $user_credential -ReturnConnection

# term store in km3
$term_studios = Get-PnPTerm -TermGroup "Knowledge" -TermSet "Studios" -Connection $conn_km3
$term_capabilities = Get-PnPTerm -TermGroup "Knowledge" -TermSet "Capabilities" -Connection $conn_km3
$term_clients = Get-PnPTerm -TermGroup "Knowledge" -TermSet "Clients" -Connection $conn_km3
$term_confidentiality = Get-PnPTerm -TermGroup "Knowledge" -TermSet "Confidentiality" -Connection $conn_km3
$term_industries = Get-PnPTerm -TermGroup "Knowledge" -TermSet "Industries" -Connection $conn_km3
$term_keywords = Get-PnPTerm -TermGroup "Knowledge" -TermSet "Keywords" -Connection $conn_km3
$term_years = Get-PnPTerm -TermGroup "Knowledge" -TermSet "Years" -Connection $conn_km3

$case_study_list_km2 = Get-PnPList -Identity "Case Study" -Connection $conn_km2
$case_study_pdf_list_km2 = Get-PnPList -Identity "PDF Conversion" -Connection $conn_km2
$case_study_list_km3 = Get-PnPList -Identity "Case Studies" -Connection $conn_km3
$sources_list_km3 = Get-PnPList -Identity "Sources" -Connection $conn_km3

$case_study_items_km2 = Get-PnPListItem -List $case_study_list_km2 -Connection $conn_km2
# $case_study_pdf_items_km2 = Get-PnPListItem -List $case_study_pdf_list_km2 -Connection $conn_km2
$case_study_items_km3 = Get-PnPListItem -List $case_study_list_km3 -Connection $conn_km3
# $sources_items_km3 = Get-PnPListItem -List $sources_list_km3 -Connection $conn_km3

foreach ($case_study_item_km3 in $case_study_items_km3) {
    $file_name = $case_study_item_km3.FieldValues.FileLeafRef
    $file_name_without_extension = $file_name.Substring(0, $file_name.LastIndexOf('.'))

    write-host ("handling $file_name_without_extension ...")
    write-host "searching original file from km2..."

    $corr_case_study_km2 = Get-PnPListItem -List $case_study_list_km2 -Query ("<View Scope='RecursiveAll'><Query><Where><Contains><FieldRef Name='FileLeafRef'/><Value Type='Computed'>{0}</Value></Contains></Where></Query></View>" -f $file_name_without_extension) -Connection $conn_km2

    if ($null -eq $corr_case_study_km2) {
        write-host "no corresponding case study found, skipping..."
        # Remove-PnPListItem -List $case_study_list_km3 -Identity $case_study_item_km3.Id -Connection $km3
    }
    else{
        write-host "found corresponding case study, searching original files in sources on km3..."
        $corr_sources_item_km3 = Get-PnPListItem -List $sources_list_km3 -Query ("<View Scope='RecursiveAll'><Query><Where><Contains><FieldRef Name='FileLeafRef'/><Value Type='Computed'>{0}</Value></Contains></Where></Query></View>" -f $file_name_without_extension) -Connection $conn_km3

        if ($null -eq $corr_sources_item_km3){
            write-host "no corresponding source found, skipping..."
        }
        else{
            if ($corr_sources_item_km3.Length -gt 1){

                # to be discussed with rob
                write-host "multiple results returned, captured the first one..."
                $corr_sources_item_km3 = $corr_sources_item_km3[0]
            }
    
            write-host "found corresponding source, updating pdf fields..."

            try {
                $case_study_km2_url = $corr_case_study_km2.FieldValues.FileRef
                $case_study_km2_file_name = $corr_case_study_km2.FieldValues.FileLeafRef
                $case_study_km2_client = $corr_case_study_km2.FieldValues["Client_test"].Label
                $case_study_km2_studio = $corr_case_study_km2.FieldValues["r0bf"]
                $case_study_km2_capability = $corr_case_study_km2.FieldValues["wq4j"]
                $case_study_km2_industry = $corr_case_study_km2.FieldValues["Industry"]
                $case_study_industry_km2_subtype = $corr_case_study_km2.FieldValues["Industry_x0020_Subtype"]
                $case_study_km2_keywords = $corr_case_study_km2.FieldValues["TaxKeyword"].Label
                $case_study_km2_year = $corr_case_study_km2.FieldValues["h1k7"]
                $case_study_km2_confidentiality = $corr_case_study_km2.FieldValues["_x0066_hp8"].Label

                $case_study_client_km3_id = ($term_clients | where-object {$_.Name -eq $case_study_km2_client}).Id.Guid
            
                # todo: industry does not pretty match with the new term store in km 3.0
                # $converted_case_study_industry_id = $term_industries | Where-Object {$_.Name -eq $original_case_study_industry}

                # capability is multiple valued
                $case_study_capability_km3_id_list = @()
                for ($i = 0;$i -lt $case_study_km2_capability.Length; $i++) {
                    $tmp_capability = $term_capabilities | Where-Object {$_.Name -eq $case_study_km2_capability[$i]}
                    $tmp_capability_id = $tmp_capability.Id.Guid
                    $case_study_capability_km3_id_list += "$tmp_capability_id"
                }

                # studio is multiple valued
                $case_study_studio_km3_id_list = @()
                for ($i = 0;$i -lt $case_study_km2_studio.Length; $i++) {
                    $tmp_studio = $term_studios | Where-Object {$_.Name -eq $case_study_km2_studio[$i]}
                    $tmp_studio_id = $tmp_studio.Id.Guid
                    $case_study_studio_km3_id_list += "$tmp_studio_id"
                }

                # todo: better to populate keyword metadata from km2.0 to km3.0 first, i will develop it in another script

                $case_study_year_km3_id = ($term_years | Where-Object {$_.Name -eq $case_study_km2_year}).Id.Guid

                $case_study_item_km3 = Set-PnPListItem -List $case_study_list_km3 -Identity $case_study_item_km3.Id -Connection $conn_km3 -Values @{
                    "Source" = ("https://frogdesign.sharepoint.com/sites/Knowledge/Sources/{0}?web=1, {1}" -f $case_study_km2_file_name, $case_study_km2_file_name);
                    "Client" = "$case_study_client_km3_id";
                    # "Industries" = ;
                    # "Capabilities" = $case_study_capability_km3_id_list;
                    # "Keywords_x0020__x0028_frog_x0020_Knowledge_x0029_" = ;
                    "Studios" = $case_study_studio_km3_id_list;
                    "Year" = "$case_study_year_km3_id"
                    # "Confidentiality" = ;
                    # "Contacts" = ;
                    # "Video" = 
                }

                write-host "done!" -f Green
                write-host "updating source fields..."
                
                write-host "checking the file out first..."
                Set-PnPFileCheckedOut -Url $corr_sources_item_km3.FieldValues.FileRef -Connection $conn_km3
                write-host "checked out..."

                $corr_sources_item_km3 = Set-PnPListItem -List $sources_list_km3 -Identity $corr_sources_item_km3.Id -Connection $conn_km3 -Values @{
                    "Reference" = ("https://frogdesign.sharepoint.com/sites/Knowledge/Case Studies/{0}, {1}" -f $file_name, $file_name)
                }

                write-host "checking in..."
                Set-PnPFileCheckedIn -Url $corr_sources_item_km3.FieldValues.FileRef -Connection $conn_km3
                write-host "done!" -f Green

                write-host "mark this item on km2 as migrated..."
                $corr_case_study_km2 = Set-PnPListItem -List $case_study_list_km2 -Identity $corr_case_study_km2.Id -Connection $conn_km2 -Values @{
                    "has_migrated" = 1
                }
                write-host "done!" -f Green
            }
            catch {
                Write-Host -f Red "error happened in uploading file!" $_.Exception.Message
            }
        }
    }
}

# generate conversion report
$to_convert = @{}
foreach ($case_study_item_km2 in $case_study_items_km2) {
    $file_name = $case_study_item_km2.FieldValues.FileLeafRef
    $file_name_without_extension = $file_name.Substring(0, $file_name.LastIndexOf('.'))

    $corr_pdfs_item_km2 = Get-PnPListItem -List $case_study_pdf_list_km2 -Query ("<View Scope='RecursiveAll'><Query><Where><Contains><FieldRef Name='FileLeafRef'/><Value Type='Computed'>{0}</Value></Contains></Where></Query></View>" -f $file_name_without_extension) -Connection $conn_km2

    if ($null -eq $corr_pdfs_item_km2){
        $to_convert.Add($file_name, 0)
    }
    else {
        if ($to_convert.Length -gt 1){
            $to_convert.Add($file_name, $to_convert.Length)
        }
        else {
            $to_convert.Add($file_name, 1)
        }
    }
}

$to_convert.GetEnumerator() | Sort-Object "Value" -Descending| Export-Csv ".\pdf_converted.csv"

Disconnect-PnPOnline -Connection $conn_km2
Disconnect-PnPOnline -Connection $conn_km3