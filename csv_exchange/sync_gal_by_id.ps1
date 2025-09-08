<#
.SYNOPSIS
  Sync Microsoft Lists CSV -> Exchange Online GAL (Mail Contacts)
  Clé d'unicité: CustomAttribute3 = "<NomListe>:<ID>"

.PARAMETERS
  -CsvPath "C:\Imports\Contacts Scolaires.csv"
  -Apply             # Appliquer (sinon dry-run)
  -EnableRemoval     # Masquer les objets absents de ce CSV (ListName courant)
  -HardDelete        # Supprimer au lieu de masquer

.NOTES
  Prérequis: Install-Module ExchangeOnlineManagement
#>

param(
  [Parameter(Mandatory=$true)] [string]$CsvPath,
  [switch]$Apply,
  [switch]$EnableRemoval,
  [switch]$HardDelete,
  [switch]$Offline,
  [string]$SmtpDomain
)

function Normalize-Email($s){ if ($null -eq $s) { return "" } ($s -replace '\s','').ToLower() }

# Corrige les erreurs simples de format SMTP (ex: '..' -> '.') et rogne les points en bord
function AutoFix-Email {
  param([string]$Address)
  $addr = Normalize-Email $Address
  if (-not $addr) { return "" }
  if ($addr -notmatch '@') { return $addr }
  $parts = $addr.Split('@')
  if ($parts.Count -ne 2) { return $addr }
  $local = $parts[0]
  $domain = $parts[1]
  # Compresse les points multiples
  $local = ($local -replace '\.\.+','.')
  $domain = ($domain -replace '\.\.+','.')
  # Supprime points en début/fin
  $local = $local.Trim('.')
  $domain = $domain.Trim('.')
  return "$local@$domain"
}

# Validation simple des adresses SMTP (format local@domaine, sans '..')
function Test-SmtpAddress {
  param([string]$Address)
  if ([string]::IsNullOrWhiteSpace($Address)) { return $false }
  # Doit ressembler à local@domaine.tld et ne pas contenir de '..'
  if ($Address -match '^([a-z0-9](?:[a-z0-9._%+\-]*[a-z0-9])?)@([a-z0-9](?:[a-z0-9\-]*[a-z0-9])?(?:\.[a-z0-9](?:[a-z0-9\-]*[a-z0-9])?)+)$' -and $Address -notmatch '\.\.') { return $true }
  return $false
}

# --- Déduire NomListe depuis le nom du fichier ---
$baseName = [System.IO.Path]::GetFileNameWithoutExtension($CsvPath)
# Extrait le mot après "Contacts " s'il existe, sinon prend tout le basename
$ListName = ($baseName -replace '^\s*Contacts\s+','').Trim()
if ([string]::IsNullOrWhiteSpace($ListName)) { $ListName = $baseName }

# --- Gestion du groupe dynamique pour la liste courante ---
function New-AliasFromListName {
  param(
    [Parameter(Mandatory=$true)][string]$ListName
  )
  $a = $ListName.ToLower()
  $a = ($a -replace "[^a-z0-9]+","-").Trim('-')
  return "ddg-contacts-$a"
}

function Ensure-DynamicDistributionGroupForList {
  param(
    [Parameter(Mandatory=$true)][string]$ListName
  )
  $name   = "Contacts $ListName"
  $alias  = New-AliasFromListName -ListName $ListName
  $filter = "CustomAttribute2 -eq 'List:$ListName'"
  $gCa1   = 'Source:Lists'
  $gCa2   = "List:$ListName"

  if ($Offline) {
    Write-Host "[DDG][DRY] $name"
    Write-Host "           filter: $filter"
    if ($SmtpDomain) { Write-Host "           smtp: $alias@$SmtpDomain" }
    return
  }

  try {
    $ddg = Get-DynamicDistributionGroup -Identity $name -ErrorAction Stop
    if ($ddg.RecipientFilter -ne $filter) {
      Write-Host "[DDG][SET] $name"
      if ($Apply) { Set-DynamicDistributionGroup -Identity $name -RecipientFilter $filter -ErrorAction Stop }
    } else {
      Write-Host "[DDG][OK] $name"
    }
    # Toujours forcer la visibilité en GAL si besoin
    if ($ddg.HiddenFromAddressListsEnabled) {
      Write-Host "[DDG][SHOW] $name"
      if ($Apply) { Set-DynamicDistributionGroup -Identity $name -HiddenFromAddressListsEnabled $false -ErrorAction Stop }
    }

    # Aligne attributs et adresse SMTP
    if ($Apply) {
      try { Set-DynamicDistributionGroup -Identity $name -CustomAttribute1 $gCa1 -CustomAttribute2 $gCa2 -ErrorAction SilentlyContinue -WarningAction SilentlyContinue } catch { }
      try { Set-DynamicDistributionGroup -Identity $name -DisplayName $name -ErrorAction SilentlyContinue -WarningAction SilentlyContinue } catch { }
      if ($SmtpDomain) {
        $smtp = "$alias@$SmtpDomain"
        try { Set-DynamicDistributionGroup -Identity $name -PrimarySmtpAddress $smtp -ErrorAction SilentlyContinue -WarningAction SilentlyContinue } catch { }
      }
    }
  } catch {
    Write-Host "[DDG][NEW] $name"
    if ($Apply) {
      $params = @{ Name = $name; Alias = $alias; RecipientFilter = $filter; ErrorAction = 'Stop' }
      if ($SmtpDomain) { $params['PrimarySmtpAddress'] = "$alias@$SmtpDomain" }
      $null = New-DynamicDistributionGroup @params
      # S'assure qu'il n'est pas caché
      Set-DynamicDistributionGroup -Identity $name -HiddenFromAddressListsEnabled $false -ErrorAction SilentlyContinue
      # Aligne attributs et affichage
      try { Set-DynamicDistributionGroup -Identity $name -CustomAttribute1 $gCa1 -CustomAttribute2 $gCa2 -DisplayName $name -ErrorAction SilentlyContinue -WarningAction SilentlyContinue } catch { }
    }
  }
}

# --- Connexion Exchange Online ---
if (-not $Offline) {
  Import-Module ExchangeOnlineManagement -ErrorAction Stop
  if (-not (Get-ConnectionInformation)) {
    Connect-ExchangeOnline -ShowBanner:$false
  }
} else {
  Write-Host "[OFFLINE] Skip Exchange connection/session"
}

# Crée/Met à jour le groupe dynamique pour cette liste
Ensure-DynamicDistributionGroupForList -ListName $ListName

# --- Charger CSV ---
if (-not (Test-Path $CsvPath)) { Write-Error "CSV introuvable: $CsvPath"; exit 65 }
$rows = Import-Csv -Path $CsvPath | ForEach-Object {
  $obj = $_ | Select-Object *; foreach ($p in $obj.PSObject.Properties) { if ($null -eq $p.Value){ $p.Value = "" } }; $obj
}
if (-not $rows) { Write-Warning "CSV vide."; exit 0 }

# --- Colonnes attendues (mappage standard) ---
$col = @{
  ID          = 'ID'
  FirstName   = 'Prénom'
  LastName    = 'Nom'
  Title       = 'Fonction'
  Company     = 'Organisation'
  Department  = 'OrgaType'
  City        = 'Commune'
  State       = 'Département'
  Street      = 'Adresse_1'
  PostalCode  = 'Code_postal'
  Phone       = 'Tel_Fixe'
  Mobile      = 'Tel_Mobile'
  MailPro     = 'Mail_Pro'
  MailPerso   = 'Mail_Perso'
  Notes       = 'Notes.'         # si présent
}

# --- Champs connus pour séparer les "extras" dynamiques ---
$known = @($col.Values + @('Niveau','Classes','Projets','Créé par')) | Select-Object -Unique

# --- Index CSV par clé "<ListName>:<ID>" + fallback email ---
function Build-Key($listName, $id){ if ([string]::IsNullOrWhiteSpace($id)) { return "" } return "$listName`:$id" }

$csvByKey   = @{}
$csvByEmail = @{}
foreach ($r in $rows) {
  $id  = $r.$($col.ID)
  $key = Build-Key $ListName $id

  $email = $r.$($col.MailPro); if ([string]::IsNullOrWhiteSpace($email)) { $email = $r.$($col.MailPerso) }
  $email = Normalize-Email $email
  if (-not (Test-SmtpAddress $email)) {
    $maybe = AutoFix-Email $email
    if (Test-SmtpAddress $maybe) { $email = $maybe } else { $email = "" }
  }
  if (-not (Test-SmtpAddress $email)) { $email = "" }

  if ($key)        { $csvByKey[$key] = $r }
  if ($email -ne "") { $csvByEmail[$email] = $r }
}

# --- Récupérer tous les Mail Contacts marqués Source:Lists pour diff/MAJ ---
if ($Offline) {
  $existing = @()
} else {
  $existing = Get-MailContact -ResultSize Unlimited -Filter "CustomAttribute1 -eq 'Source:Lists'" |
    Select-Object Guid,Identity,ExternalEmailAddress,DisplayName,CustomAttribute1,CustomAttribute2,CustomAttribute3,CustomAttribute4,HiddenFromAddressListsEnabled
}

# Index existants par CustomAttribute3 et par email
$exByKey   = @{}
$exByEmail = @{}
foreach ($ex in $existing) {
  if ($ex.CustomAttribute3) { $exByKey[$ex.CustomAttribute3] = $ex }
  $exByEmail[(Normalize-Email ($ex.ExternalEmailAddress -ireplace '^smtp:',''))] = $ex
}

# --- Helpers ---
function Build-DisplayName($r, $col) {
  $fn = $r.$($col.FirstName); $ln = $r.$($col.LastName)
  if ($fn -and $ln) { return "$fn $ln" }
  if ($ln) { return $ln }
  if ($fn) { return $fn }
  $company = $r.$($col.Company)
  if ($company -and $company.Trim() -ne "") { return $company }
  $mp = $r.$($col.MailPro); $ms = $r.$($col.MailPerso)
  return @($mp,$ms) | Where-Object { $_ -and $_.Trim() -ne "" } | Select-Object -First 1
}
function Build-ExtrasJson($r, $known){
  $map = @{}
  foreach ($p in $r.PSObject.Properties) {
    $k = $p.Name; $v = "$($p.Value)".Trim()
    if ($known -notcontains $k -and $v -ne "") { $map[$k] = $v }
  }
  if ($map.Keys.Count -eq 0) { return $null }
  return ($map | ConvertTo-Json -Compress)
}
function Build-Notes($r, $col, $known){
  $notesIn = ""
  if ($r.PSObject.Properties.Name -contains $col.Notes) { $notesIn = "$($r.$($col.Notes))".Trim() }
  $out = @(); if ($notesIn) { $out += $notesIn }

  $extras = @()
  foreach ($p in $r.PSObject.Properties) {
    $k = $p.Name; $v = "$($p.Value)".Trim()
    if ($known -notcontains $k -and $v -ne "") { $extras += "- $k : $v" }
  }
  if ($extras.Count -gt 0) {
    if ($out.Count -gt 0) { $out += "" }
    $out += "Champs supplémentaires :"
    $out += $extras
  }
  return ($out -join "`n")
}

$created = 0; $updated = 0; $hidden = 0; $removed = 0; $skippedConflict = 0

# --- UPSERT principal (par clé <ListName>:<ID> avec fallback email) ---
foreach ($k in $csvByKey.Keys) {
  $r = $csvByKey[$k]

  $first = $r.$($col.FirstName)
  $last  = $r.$($col.LastName)
  $name  = Build-DisplayName $r $col
  $company = $r.$($col.Company)
  $dept    = $r.$($col.Department)
  $title   = $r.$($col.Title)
  $city    = $r.$($col.City)
  $state   = $r.$($col.State)
  $street  = $r.$($col.Street)
  $postal  = $r.$($col.PostalCode)
  $phone   = $r.$($col.Phone)
  $mobile  = $r.$($col.Mobile)

  $mail = $r.$($col.MailPro); if ([string]::IsNullOrWhiteSpace($mail)) { $mail = $r.$($col.MailPerso) }
  $email0 = Normalize-Email $mail
  if (-not (Test-SmtpAddress $email0)) {
    $maybe = AutoFix-Email $email0
    if (Test-SmtpAddress $maybe) { Write-Host "[FIX] Email corrigé: <$email0> -> <$maybe> ($ca3)"; $email = $maybe } else { $email = "" }
  } else { $email = $email0 }

  $ca1 = 'Source:Lists'
  $ca2 = "List:$ListName"
  $ca3 = $k                                   # clé de synchro "<ListName>:<ID>"
  $ca4 = Build-ExtrasJson $r $known
  $notes = Build-Notes $r $col $known

  # Cherche existant par clé, sinon par email
  $exists = $exByKey[$k]
  if (-not $exists -and $email -ne "") { $exists = $exByEmail[$email] }

  if ($exists) {
    Write-Host "[UPD] $name <$email> ($ca3)"
    if ($Apply) {
      # Remet visible si masqué
      if ($exists.HiddenFromAddressListsEnabled) {
        Set-MailContact -Identity $exists.Guid -HiddenFromAddressListsEnabled $false -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
      }

      Set-MailContact -Identity $exists.Guid `
        -DisplayName $name `
        -ErrorAction Stop -WarningAction SilentlyContinue

      # S'assure que les attributs de synchro et le filtre DDG sont posés
      Set-MailContact -Identity $exists.Guid `
        -CustomAttribute1 $ca1 -CustomAttribute2 $ca2 -CustomAttribute3 $ca3 `
        -ErrorAction Stop -WarningAction SilentlyContinue

      # Met à jour l'adresse externe si elle est valable et a changé
      if ($email -ne "" -and (Test-SmtpAddress $email)) {
        $exEmailNorm = Normalize-Email ($exists.ExternalEmailAddress -ireplace '^smtp:','')
        if ($exEmailNorm -ne $email) {
          try { Set-MailContact -Identity $exists.Guid -ExternalEmailAddress $email -ErrorAction Stop -WarningAction SilentlyContinue } catch { }
        }
      }

      if ($ca4) { Set-MailContact -Identity $exists.Guid -CustomAttribute4 $ca4 -ErrorAction SilentlyContinue -WarningAction SilentlyContinue }
      else      { Set-MailContact -Identity $exists.Guid -CustomAttribute4 $null -ErrorAction SilentlyContinue -WarningAction SilentlyContinue }

      try {
        Set-Contact -Identity $exists.Guid `
          -FirstName $first -LastName $last `
          -Title $title -Company $company -Department $dept `
          -City $city -StateOrProvince $state `
          -StreetAddress $street -PostalCode $postal `
          -Phone $phone -MobilePhone $mobile `
          -Notes $notes `
          -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
      } catch { }
    }
    $updated++
  } else {
    Write-Host "[NEW] $name <$email> ($ca3)"
    if ($Apply) {
      if ([string]::IsNullOrWhiteSpace($email)) {
        # Informe si email présent mais invalide, sinon email manquant
        $rawPro = Normalize-Email $r.$($col.MailPro)
        $rawPer = Normalize-Email $r.$($col.MailPerso)
        if ($rawPro -and -not (Test-SmtpAddress $rawPro)) {
          Write-Warning "[SKIP] Email pro invalide: $rawPro (ID=$k, Nom=$name)"
        } elseif ($rawPer -and -not (Test-SmtpAddress $rawPer)) {
          Write-Warning "[SKIP] Email perso invalide: $rawPer (ID=$k, Nom=$name)"
        } else {
          Write-Warning "[SKIP] Impossible de créer sans email externe (ID=$k, Nom=$name)"
        }
      } else {
        # Conflit global: l'adresse est-elle déjà utilisée par un autre destinataire ?
        $conflict = @()
        if (-not $Offline) {
          try {
            $proxy = "smtp:$email"
            $filter = ("EmailAddresses -eq '{0}'" -f $proxy.Replace("'","''"))
            $conflict = Get-Recipient -ResultSize 5 -Filter $filter -ErrorAction Stop
          } catch { $conflict = @() }
        }
        if ($conflict -and $conflict.Count -gt 0) {
          $mcToReuse = $null
          if ($conflict.Count -eq 1 -and $conflict[0].RecipientTypeDetails -eq 'MailContact') {
            $mcToReuse = $conflict[0]
          }
          if ($mcToReuse) {
            Write-Host "[REUSE] $name <$email> ($ca3) -> $($mcToReuse.Identity)"
            if ($Apply) {
              try { Set-MailContact -Identity $mcToReuse.Guid -HiddenFromAddressListsEnabled $false -ErrorAction SilentlyContinue -WarningAction SilentlyContinue } catch { }
              Set-MailContact -Identity $mcToReuse.Guid `
                -CustomAttribute1 $ca1 -CustomAttribute2 $ca2 -CustomAttribute3 $ca3 `
                -ErrorAction Stop -WarningAction SilentlyContinue
              if ($ca4) { Set-MailContact -Identity $mcToReuse.Guid -CustomAttribute4 $ca4 -ErrorAction SilentlyContinue -WarningAction SilentlyContinue }
              else      { Set-MailContact -Identity $mcToReuse.Guid -CustomAttribute4 $null -ErrorAction SilentlyContinue -WarningAction SilentlyContinue }
              try {
                Set-Contact -Identity $mcToReuse.Guid `
                  -FirstName $first -LastName $last `
                  -Title $title -Company $company -Department $dept `
                  -City $city -StateOrProvince $state `
                  -StreetAddress $street -PostalCode $postal `
                  -Phone $phone -MobilePhone $mobile `
                  -Notes $notes `
                  -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
              } catch { }
            }
            $updated++
            continue
          } else {
            $types = ($conflict | Select-Object -ExpandProperty RecipientTypeDetails | Select-Object -Unique) -join ','
            Write-Warning ("[SKIP] Adresse déjà utilisée par {0} : {1} (ID={2}, Nom={3})" -f $types, $email, $k, $name)
            $skippedConflict++
            continue
          }
        }

        # Garantir un -Name unique et valide (<= 64 chars). Conserve DisplayName complet.
        $idSuffix = ($ca3 -split ':')[-1]
        $idSuffix = ($idSuffix -replace '[^A-Za-z0-9-]','')
        if ($idSuffix.Length -gt 10) { $idSuffix = $idSuffix.Substring($idSuffix.Length - 10) }
        $suffix = "-$idSuffix"
        $maxBase = 64 - $suffix.Length
        if ($maxBase -lt 1) { $maxBase = 63 }
        $base = $name
        if ($null -ne $base -and $base.Length -gt $maxBase) { $base = $base.Substring(0, $maxBase) }
        $safeName = "$base$suffix"
        $mc = New-MailContact -Name $safeName `
        -DisplayName $name `
        -FirstName $first -LastName $last `
        -ExternalEmailAddress $email `
        -ErrorAction Stop

        if ($mc) {
          Set-MailContact -Identity $mc.Identity `
            -CustomAttribute1 $ca1 -CustomAttribute2 $ca2 -CustomAttribute3 $ca3 `
            -ErrorAction Stop -WarningAction SilentlyContinue
          if ($ca4) { Set-MailContact -Identity $mc.Identity -CustomAttribute4 $ca4 -ErrorAction SilentlyContinue -WarningAction SilentlyContinue }
          try {
            Set-Contact -Identity $mc.Identity `
              -FirstName $first -LastName $last `
              -Title $title -Company $company -Department $dept `
              -City $city -StateOrProvince $state `
              -StreetAddress $street -PostalCode $postal `
              -Phone $phone -MobilePhone $mobile `
              -Notes $notes `
              -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
          } catch { }
        }
      }
    }
    $created++
  }
}

# --- Nettoyage optionnel : absents de la liste courante (même ListName) ---
if ($EnableRemoval) {
  # Tous les contacts Source:Lists pour ce ListName
  $existingForList = $existing | Where-Object { $_.CustomAttribute2 -eq "List:$ListName" }
  foreach ($ex in $existingForList) {
    $key = $ex.CustomAttribute3
    if (-not $csvByKey.ContainsKey($key)) {
      if ($HardDelete) {
        Write-Host "[DEL] $($ex.DisplayName) <$($ex.ExternalEmailAddress)> ($key)"
        if ($Apply) { Remove-MailContact -Identity $ex.Guid -Confirm:$false -ErrorAction Stop }
        $removed++
      } else {
        Write-Host "[HIDE] $($ex.DisplayName) <$($ex.ExternalEmailAddress)> ($key)"
        if ($Apply) { Set-MailContact -Identity $ex.Guid -HiddenFromAddressListsEnabled $true -ErrorAction Stop -WarningAction SilentlyContinue }
        $hidden++
      }
    }
  }
}

Write-Host "----"
Write-Host ("Résumé: {0} créés, {1} mis à jour, {2} masqués, {3} supprimés, {4} conflits (Apply={5})" -f $created,$updated,$hidden,$removed,$skippedConflict,$Apply)

# --- Post-check: aperçu des membres DDG ---
if (-not $Offline) {
  try {
    $gname = "Contacts $ListName"
    $g = Get-DynamicDistributionGroup -Identity $gname -ErrorAction Stop
    $recips = Get-Recipient -RecipientPreviewFilter $g.RecipientFilter -ErrorAction Stop
    $count = ($recips | Measure-Object).Count
    Write-Host ("[DDG][PREVIEW] {0} destinataires pour '{1}'" -f $count,$gname)
    $sample = $recips | Select-Object -First 10 -Property Name,RecipientTypeDetails,CustomAttribute2
    if ($sample) { $sample | Format-Table -AutoSize | Out-String | Write-Host }
  } catch { }
}
