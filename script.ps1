# Désactive l'affichage de la commande en cours d'exécution pour un script plus propre
$echoPreference = $false

# Affiche un message
Write-Host "Mise à jour de Word..."
Start-Sleep -Seconds 1

# Change le répertoire courant
Set-Location -Path "C:\Program Files\Microsoft Office\Office16"
Start-Sleep -Seconds 1

# Exécute la commande et capture la sortie
$output = & cscript ospp.vbs /dstatus
Start-Sleep -Seconds 1

# Trouve la ligne contenant "SKU ID" et extrait l'ID
$skuId = ($output | Where-Object { $_ -match "SKU ID:\s+([-\w]+)" }) -replace "SKU ID:\s+",""
echo $skuId

if ($skuId -ne $null) {
    # Exécute la commande avec le SKU ID récupéré
    for ($i=0; $i -lt 3; $i++) {
        Start-Sleep -Seconds 1
        & .\OSPPREARM.EXE $skuId
    }
} else {
    Write-Host "SKU ID non trouvé."
}
