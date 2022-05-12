$session = New-Object -ComObject Microsoft.Update.Session
$manager = New-Object -ComObject Microsoft.Update.ServiceManager
$service = $manager.AddScanPackageService("Offline Sync Service", "C:\github\kbupdate\wsusscn2.cab", 1)
$searcher = $session.CreateUpdateSearcher()

Write-Host "Searching for updates..."

$searcher.ServerSelection = 3 #' ssOthers
$searcher.ServiceID = $service.ServiceID
$SearchResult = $searcher.Search("IsInstalled=0")
$Updates = $SearchResult.Updates

If ($Updates.Count -eq 0) {
    Write-Host "There are no applicable updates."
    break
}

Write-Host "List of applicable items on the machine when using wssuscan.cab:"

0..($Updates.Count - 1) | ForEach-Object {
    $update = $Updates.Item($psitem)
    Write-Host $update.Title
    Next
}