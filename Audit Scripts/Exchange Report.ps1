#This section will look at all the mailboxes and will compile a list of the ones with items in their outboxes.
Get-Mailbox -ResultSize Unlimited | 
Get-MailboxFolderStatistics | 
Where-Object {$_.Name -eq "Outbox" -and $_.ItemsInFolder -gt '0' } | 
Select-Object Identity, FolderType, ItemsinFolder, FolderSize | Format-Table Identity, FolderType, ItemsinFolder, FolderSize

#This section shows all disconnected mailboxes as well as the date the corresponding user account was deleted.
get-mailboxserver | 
get-mailboxstatistics | 
where { $_.DisconnectDate } | 
fl DisplayName, DisconnectDate

#This section shows all users who are configured in Exchange to forward elsewhere.
Get-Mailbox -Resultsize Unlimited | 
Where-Object {$_.ForwardingAddress} | 
Select-Object Name, @{Expression={$_.ForwardingAddress};Label="Forwarded to"}, @{Expression={$_.DeliverToMailboxAndForward};Label="Mailbox & Forward"}

#This section shows mailbox database sizes from largest to smallest. 
get-mailboxdatabase -includepre | 
foreach-object{select-object -inputobject $_ -property *,@{name="MailboxDBSizeinGB";expression={[math]::Round(((get-item ("\\" + $_.servername + "\" + $_.edbfilepath.pathname.replace(":","$"))).length / 1GB),2)}}} | 
Sort-Object mailboxdbsizeinGB -Descending | 
format-table identity,mailboxdbsizeinGB -autosize

#This section shows a table of mailbox sizes
get-mailbox -resultsize unlimited -erroraction silentlycontinue | 
Get-MailboxStatistics | 
select-object *,@{name="TotalItemSizeinMB";expression={[math]::Round($_.totalitemsize.value.ToBytes() / 1MB,2)}},@{name="TotalDeletedItemSizeinMB";expression={[math]::Round($_.totaldeleteditemsize.value.ToBytes() / 1MB,2)}},@{name="CombinedTotalSizeinMB";expression={[math]::Round($_.totalitemsize.value.ToBytes() / 1MB,2) + [math]::Round($_.totaldeleteditemsize.value.ToBytes() / 1MB,2)}} | 
sort-object combinedtotalsizeinMB -Descending | 
format-table displayname,*MB

#This section shows displays stats about the percentage of mailboxes that consume the total space of the database.
1..10 | 
foreach-object{"The top $($_*10) percent of the mailboxes consume $([int]((($mailboxsizelist | select-object -first ($mailboxsizelist.count * ($_ / 10)) | measure-object CombinedTotalSizeinMB -sum).sum / ($mailboxsizelist | 
measure-object combinedtotalsizeinMB -sum).sum) * 100)) percent of the Size of all the mailboxes."}
