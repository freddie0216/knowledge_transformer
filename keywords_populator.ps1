$user_credential = Get-Credential

# get keywords term store from km 3.0
Connect-PnPOnline -Url https://frogdesign.sharepoint.com/sites/Knowledge -Credentials $user_credential
$term_keywords = Get-PnPTerm -TermGroup "Knowledge" -TermSet "Keywords"
Disconnect-PnPOnline

# get case studies from km 2.0
Connect-PnPOnline -Url https://frogdesign.sharepoint.com/sites/KnowledgeManagement -Credentials $user_credential
$original_case_study_library = Get-PnPList -Identity "Case Study"
$original_case_study_docs = Get-PnPListItem -List $original_case_study_library

$keywords_to_add = @{}
foreach($original_case_study_doc in $original_case_study_docs){
    $keywords = $original_case_study_doc.FieldValues["TaxKeyword"].Label
    
    foreach($keyword in $keywords){
        $keyword_km3 = $term_keywords | where-object{$_.Name -eq $keyword}
        if(!$keyword_km3){
            $keyword = $keyword.ToLower()
            if($keywords_to_add.Contains($keyword)){
                $keywords_to_add[$keyword] = $keywords_to_add[$keyword] + 1
            }
            else{
                $keywords_to_add.Add($keyword, 1)
            }
        }
    }
}

$keywords_to_add.GetEnumerator() | Sort-Object "Value" -Descending| Export-Csv ".\keywords_km2_v2.csv"