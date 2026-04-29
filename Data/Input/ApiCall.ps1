param(
    [string]$User,
    [string]$Password,
    [string]$ContractNumbers,
    [string]$ContractsFile,
    [string]$BaseUrl      = "https://iabtgs-dev2.fa.ocs.oraclecloud.com",
    [string]$OutputFolder = ".",
    [int]   $BatchSize    = 25,
    [string]$TemplateFile = "05-IDF_TEMPLATE_SUBMIT_CONTRACT_V1.0.xlsx",
    [int]   $TemplateFirstDataRow = 6
)

$ErrorActionPreference = "Stop"

# ------------------------------------------------------------------
# Si -ContractsFile est fourni, on lit la liste depuis le fichier
# (contourne la limite CMD 8191 caracteres)
# ------------------------------------------------------------------
if (-not [string]::IsNullOrEmpty($ContractsFile)) {
    if (-not (Test-Path $ContractsFile)) {
        throw "Fichier ContractsFile introuvable : $ContractsFile"
    }
    $fileContent = Get-Content -Path $ContractsFile -Raw -Encoding UTF8
    if ([string]::IsNullOrEmpty($ContractNumbers)) {
        $ContractNumbers = $fileContent
    } else {
        $ContractNumbers = $ContractNumbers + "," + $fileContent
    }
}

# ------------------------------------------------------------------
# Log fichier
# ------------------------------------------------------------------
if (-not (Test-Path $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
}
$logFile = Join-Path $OutputFolder "requeteAPI.log"
"`n===== RUN $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') =====" | Out-File -FilePath $logFile -Encoding UTF8 -Append

function Log {
    param([string]$msg)
    $line = "{0} {1}" -f (Get-Date -Format "HH:mm:ss"), $msg
    $line | Out-File -FilePath $logFile -Encoding UTF8 -Append
    Write-Host $line
}

# ------------------------------------------------------------------
# Parametres
# ------------------------------------------------------------------
if ([string]::IsNullOrEmpty($User))            { throw "Parametre User manquant" }
if ([string]::IsNullOrEmpty($Password))        { throw "Parametre Password manquant" }
if ([string]::IsNullOrEmpty($ContractNumbers)) { throw "Parametre ContractNumbers vide" }

$ContractList = $ContractNumbers -split '[,;\s]+' | Where-Object { $_ -ne "" } | Select-Object -Unique
if ($ContractList.Count -eq 0) { throw "Aucun numero de contrat apres parsing" }

Log "User: $User | Contracts attendus: $($ContractList.Count) | Batch: $BatchSize | Output: $OutputFolder"

# ------------------------------------------------------------------
# Auth
# ------------------------------------------------------------------
$pair  = [string]("${User}:${Password}")
$bytes = [System.Text.Encoding]::UTF8.GetBytes($pair)
$b     = [System.Convert]::ToBase64String($bytes)
$hdr = @{
    "Authorization"          = "Basic $b"
    "REST-Framework-Version" = "4"
}

# ------------------------------------------------------------------
# Batching : decouper la liste en paquets de $BatchSize
# ------------------------------------------------------------------
$batches = @()
for ($i = 0; $i -lt $ContractList.Count; $i += $BatchSize) {
    $end = [Math]::Min($i + $BatchSize - 1, $ContractList.Count - 1)
    $batches += ,($ContractList[$i..$end])
}
Log ("Decoupage en {0} batch(es) de {1} contrats max" -f $batches.Count, $BatchSize)

# ------------------------------------------------------------------
# Appel API par batch, avec pagination interne
# Pas besoin d'expand : on ne lit que les champs du header de contrat
# ------------------------------------------------------------------
$all = [Collections.ArrayList]@()
$swApi = [System.Diagnostics.Stopwatch]::StartNew()
$batchIdx = 0
foreach ($batch in $batches) {
    $batchIdx++
    $contractListStr = ($batch | ForEach-Object { "'$_'" }) -join ","
    # VersionType='C' = Current version (filtre serveur pour ne pas charger les historiques)
    $q = "ContractNumber IN ($contractListStr) AND VersionType='C'"
    $off = 0
    do {
        $body = @{
            q        = $q
            onlyData = "true"
            limit    = 500
            offset   = $off
        }
        try {
            $r = Invoke-WebRequest -UseBasicParsing `
                -Uri "$BaseUrl/fscmRestApi/resources/11.13.18.05/contracts" `
                -Headers $hdr -Body $body
        } catch {
            Log "ERREUR batch $batchIdx offset=$off : $($_.Exception.Message)"
            throw
        }
        $r.RawContentStream.Position = 0
        $reader = [IO.StreamReader]::new($r.RawContentStream, [Text.Encoding]::UTF8)
        $json   = $reader.ReadToEnd()
        $reader.Close()
        $page = $json | ConvertFrom-Json
        foreach ($it in $page.items) { [void]$all.Add($it) }
        Log ("  batch $batchIdx/$($batches.Count) offset=$off charge=$($page.items.Count) total=$($all.Count) hasMore=$($page.hasMore)")
        $off += 500
    } while ($page.hasMore)
}
$swApi.Stop()
Log ("API OK : $($all.Count) contrats (version courante) recuperes en $([int]$swApi.Elapsed.TotalSeconds)s")

# ------------------------------------------------------------------
# Dedup de securite : ceinture + bretelles au cas ou le filtre serveur
# ne marcherait pas sur un autre pod. Sur dev5 avec VersionType='C',
# cette etape est un no-op.
# ------------------------------------------------------------------
$latestByNumber = @{}
foreach ($c in $all) {
    $num = [string]$c.ContractNumber
    if ([string]::IsNullOrEmpty($num)) { continue }
    $ver = 0
    if ($null -ne $c.MajorVersion) {
        [int]::TryParse([string]$c.MajorVersion, [ref]$ver) | Out-Null
    }
    if (-not $latestByNumber.ContainsKey($num)) {
        $latestByNumber[$num] = @{ Version = $ver; Contract = $c }
    }
    elseif ($ver -gt $latestByNumber[$num].Version) {
        $latestByNumber[$num] = @{ Version = $ver; Contract = $c }
    }
}
$before = $all.Count
$all = [Collections.ArrayList]@()
foreach ($entry in $latestByNumber.Values) { [void]$all.Add($entry.Contract) }
if ($before -ne $all.Count) {
    Log ("Dedup de securite : $before items -> $($all.Count) contrats")
}

# Liste des contrats non trouves (pour diagnostic)
$foundNumbers = @($all | ForEach-Object { $_.ContractNumber })
$missing = $ContractList | Where-Object { $foundNumbers -notcontains $_ }
if ($missing.Count -gt 0) {
    Log ("ATTENTION : $($missing.Count) contrats non trouves : " + ($missing -join ", "))
}

# ------------------------------------------------------------------
# Preparation 2D array (1 ligne par contrat, sans header : le
# template a deja sa propre ligne d'en-tete en ligne 1).
# Ordre des colonnes : A..O = Columns, Status, Error Message, Action,
# ContractNumber, StsCode, MajorVersion, ContractTypeName,
# LegalEntityName, Cognomen, CurrencyCode, StartDate, EndDate,
# EstimatedAmount, WebServiceFlag.
# ------------------------------------------------------------------
$colCount  = 15
$totalRows = $all.Count
Log "Total contrats a ecrire : $totalRows"
if ($totalRows -eq 0) {
    Log "Aucun contrat a ecrire, on s'arrete sans toucher au template"
    return
}

$grid = New-Object 'object[,]' $totalRows, $colCount

$swFill = [System.Diagnostics.Stopwatch]::StartNew()
$rowIdx = 0
foreach ($c in $all) {
    # Action : valeur fixe pour declencher la soumission a l'approbation
    $action = "SUBMIT FOR APPROVAL"

    # Normalisation des dates : on ne garde que YYYY-MM-DD si jamais l'API renvoie un timestamp
    $startDate = [string]$c.StartDate
    if ($startDate -and $startDate.Length -ge 10) { $startDate = $startDate.Substring(0,10) }
    $endDate = [string]$c.EndDate
    if ($endDate -and $endDate.Length -ge 10) { $endDate = $endDate.Substring(0,10) }

    $grid[$rowIdx, 0]  = ""                          # Columns        (vide)
    $grid[$rowIdx, 1]  = ""                          # Status         (vide)
    $grid[$rowIdx, 2]  = ""                          # Error Message  (vide)
    $grid[$rowIdx, 3]  = $action
    $grid[$rowIdx, 4]  = [string]$c.ContractNumber
    $grid[$rowIdx, 5]  = [string]$c.StsCode
    $grid[$rowIdx, 6]  = [string]$c.MajorVersion
    $grid[$rowIdx, 7]  = [string]$c.ContractTypeName
    $grid[$rowIdx, 8]  = [string]$c.LegalEntityName
    $grid[$rowIdx, 9]  = [string]$c.Cognomen
    $grid[$rowIdx, 10] = [string]$c.CurrencyCode
    $grid[$rowIdx, 11] = $startDate
    $grid[$rowIdx, 12] = $endDate
    $grid[$rowIdx, 13] = [string]$c.EstimatedAmount
    $grid[$rowIdx, 14] = [string]$c.WebServiceFlag
    $rowIdx++
}
$swFill.Stop()
Log ("Buffer memoire rempli : $rowIdx lignes en $([int]$swFill.Elapsed.TotalMilliseconds) ms")

# ------------------------------------------------------------------
# Excel : on ouvre une copie du template et on y append les lignes.
# Le template d'origine n'est jamais modifie.
# ------------------------------------------------------------------
# Resolution du chemin du template :
#  - chemin absolu fourni ?  on l'utilise tel quel
#  - sinon on cherche d'abord a cote du .ps1 ($PSScriptRoot)
#  - puis dans le repertoire courant
$templatePath = $null
if ([System.IO.Path]::IsPathRooted($TemplateFile) -and (Test-Path $TemplateFile)) {
    $templatePath = (Resolve-Path $TemplateFile).Path
} else {
    $candidates = @()
    if ($PSScriptRoot) { $candidates += (Join-Path $PSScriptRoot $TemplateFile) }
    $candidates += (Join-Path (Get-Location).Path $TemplateFile)
    foreach ($cand in $candidates) {
        if (Test-Path $cand) { $templatePath = (Resolve-Path $cand).Path; break }
    }
}
if (-not $templatePath) {
    throw "Template introuvable : $TemplateFile (cherche dans `$PSScriptRoot et le repertoire courant)"
}
Log "Template source : $templatePath"

$stamp     = Get-Date -Format "yyyyMMdd_HHmmss"
$baseName  = [System.IO.Path]::GetFileNameWithoutExtension($templatePath)
$ext       = [System.IO.Path]::GetExtension($templatePath)
$fileName  = "{0}_{1}{2}" -f $baseName, $stamp, $ext
$tempDir   = Join-Path $env:TEMP "OracleContractsExport"
if (-not (Test-Path $tempDir)) { New-Item -ItemType Directory -Path $tempDir -Force | Out-Null }
$tempPath  = Join-Path $tempDir $fileName
$finalPath = Join-Path $OutputFolder $fileName

# On copie le template vers le temp avant ouverture pour ne JAMAIS toucher l'original
Copy-Item -Path $templatePath -Destination $tempPath -Force
Log "Copie de travail : $tempPath"

Log "Ouverture Excel..."
$x = New-Object -ComObject Excel.Application
$x.Visible        = $false
$x.DisplayAlerts  = $false
$x.ScreenUpdating = $false
$x.EnableEvents   = $false

try {
    $wb = $x.Workbooks.Open($tempPath)
    # xlCalculationManual
    try { $x.Calculation = -4135 } catch { Log "Note: Calculation non supporte, ignore" }

    $ws = $wb.Worksheets.Item(1)

    # Ligne de depart pour les donnees :
    #  - on part de $TemplateFirstDataRow (defaut 6, apres les 5 lignes de descripteurs)
    #  - on scanne vers le bas dans la colonne D (Action) au cas ou des donnees
    #    seraient deja presentes (si le template a ete pre-rempli ou si on
    #    relance le script sur un fichier deja partiellement rempli)
    $startRow = $TemplateFirstDataRow
    while ($null -ne $ws.Cells.Item($startRow, 4).Value2 -and
           [string]$ws.Cells.Item($startRow, 4).Value2 -ne "") {
        $startRow++
    }
    Log "Premiere ligne libre detectee : $startRow"

    $endRow = $startRow + $totalRows - 1
    # Forcer le format texte sur la plage de donnees pour eviter qu'Excel
    # interprete les dates / numeros (ContractNumber peut commencer par 0, etc.)
    $rangeAll = $ws.Range($ws.Cells.Item($startRow,1), $ws.Cells.Item($endRow, $colCount))
    $rangeAll.NumberFormat = "@"

    $swWrite = [System.Diagnostics.Stopwatch]::StartNew()
    $rangeAll.Value2 = $grid
    $swWrite.Stop()
    Log ("Excel : ecriture bloc ($totalRows lignes, ligne $startRow a $endRow) en $([int]$swWrite.Elapsed.TotalMilliseconds) ms")

    # xlCalculationAutomatic
    try { $x.Calculation = -4105 } catch {}

    $wb.Save()
    $wb.Close($false)
    Log "Excel ferme"
}
finally {
    $x.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($x) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Move-Item -Path $tempPath -Destination $finalPath -Force
Log "Fichier final : $finalPath"
Write-Output $finalPath