

<# 
========================================================================
 Script : Exemple de script PowerShell pour créer un utilisateur AD
 Auteur : (Votre Nom)
 Date   : (Date de dernière modification)
 Objectif : 
    1) Vérifier qu'un administrateur s'authentifie.
    2) Créer un utilisateur Active Directory avec :
       - OU dynamique (sélection par Out-GridView ou sélection manuelle)
       - Définition des coordonnées (téléphone, adresse, etc.)
       - Ajout dans des groupes selon le type de PC
       - Héritage optionnel des droits d'un autre utilisateur
       - Paramétrage des adresses mail (principal + alias)
       - Assignation de licences Office 365 via des groupes
========================================================================
#>

# --- FIX 1 : Import du module AD en premier ---
Import-Module ActiveDirectory

##########################################################################
### 1) Contrôle d’accès administrateur
##########################################################################
function Test-AdminAccess {
    param (
        [Parameter(Mandatory=$true)]
        [System.Management.Automation.PSCredential]$Credential
    )

    # OU ou chemin dont doivent dépendre les comptes admin
    $adminOU = "OU=Admins,DC=MYDOMAIN,DC=COM"  # <-- Ajustez selon votre environnement

    try {
        # Tenter d'authentifier l'utilisateur avec les credentials fournis
        $user = Get-ADUser -Identity $Credential.UserName -Credential $Credential -Properties DistinguishedName

        # Vérifier si l'utilisateur est dans l'OU Admin (ou un chemin particulier)
        if ($user.DistinguishedName -like "*$adminOU") {
            return $true
        }
    }
    catch {
        Write-Host "Erreur d'authentification ou utilisateur non trouvé."
    }

    return $false
}

# --- Boucle principale pour s'authentifier ---
$maxAttempts = 3
$attemptCount = 0

do {
    $attemptCount++

    if ($attemptCount -gt $maxAttempts) {
        Write-Host "Nombre maximal de tentatives atteint. Le script va se terminer."
        exit
    }

    $credential = Get-Credential -Message "Entrez vos identifiants d'administrateur (Tentative $attemptCount sur $maxAttempts)"
    if ($null -eq $credential) {
        Write-Host "Opération annulée par l'utilisateur."
        exit
    }

    $isAuthorized = Test-AdminAccess -Credential $credential
    if (-not $isAuthorized) {
        Write-Host "Accès refusé. Utilisateur non autorisé ou identifiants incorrects."
    }

} while (-not $isAuthorized)

Write-Host "Accès autorisé. Bienvenue dans le script, $($credential.UserName)!"

##########################################################################
### 2) (Optionnel) Fonction d’affichage de menu
##########################################################################
function Show-Menu {
    Clear-Host
    Write-Host "===============================" -ForegroundColor Cyan
    Write-Host "   Menu Principal"
    Write-Host "===============================" -ForegroundColor Cyan
    Write-Host "1. Créer un utilisateur" -ForegroundColor Yellow
    Write-Host "2. Modifier/Supprimer un utilisateur" -ForegroundColor Yellow
    Write-Host "3. Gérer les alias (UPN)" -ForegroundColor Yellow
    Write-Host "4. Gérer les packs Office 365" -ForegroundColor Yellow
    Write-Host "5. Référentiel du script (lignes de code)" -ForegroundColor Yellow
    Write-Host "6. Suivi des modifications" -ForegroundColor Yellow
    Write-Host "7. Quitter" -ForegroundColor Yellow
    Write-Host "===============================" -ForegroundColor Cyan
}

##########################################################################
### 3) Fonction principale : Create-User
##########################################################################
function Create-User {

    # -----------------------------------------------------------------
    # MAPPING ENTRE COMPAGNIES ET DOMAINES EMAIL
    # -----------------------------------------------------------------
    $companyDomainsMap = @{
        "Compagnie" = @("Entite.fr")  # exemple
    }

    # OU de base par compagnie
    $companyBaseOUMap = @{
        # --- FIX 2 : Corrections de l'orthographe / DNs ---
        "Compagnie" = "OU=Compagnie,OU=Chemin,DC=MYDOMAIN,DC=COM"
    }

    # Départements par entreprise
    $companyDepartments = @{
        "Compagnie" = @(
            "Accueil", "Achats", "Administratif"
        )
    }

    # -----------------------------------------------------------------
    # Informations des sites (adresses, code postal, etc.)
    # -----------------------------------------------------------------
    $sitesInfo = @{
        "Compagnie" = @{
            Office = "Compagnie"
            Sites  = @(
                @{ Street = "Rue d'ici"; PostalCode = "60152"; City = "Hyrule" }
            )
        }
    }

    # -----------------------------------------------------------------
    # Liste des domaines possibles pour des alias supplémentaires
    # -----------------------------------------------------------------
    $aliasDomainsList = @(
        "Entite.fr"
    )

    # -----------------------------------------------------------------
    # Fonctions d'aide pour manipuler les groupes AD
    # -----------------------------------------------------------------
    function Get-GroupDN {
        param ([string]$GroupDN)
        $group = Get-ADGroup -Filter { DistinguishedName -eq $GroupDN }
        if ($group) {
            return $group.DistinguishedName
        }
        else {
            Write-Error "Groupe non trouvé : $GroupDN"
            return $null
        }
    }

    function Add-UserToGroup {
        param (
            [string]$UserDN,
            [string]$GroupDN
        )
        try {
            Add-ADGroupMember -Identity $GroupDN -Members $UserDN -ErrorAction Stop
            Write-Host "Utilisateur ajouté au groupe : $GroupDN" -ForegroundColor Green
        }
        catch {
            Write-Error "Erreur lors de l'ajout au groupe $GroupDN : $_"
        }
    }

    # -----------------------------------------------------------------
    # Fonctions d'aide supplémentaires (sélections, etc.)
    # -----------------------------------------------------------------
    function Select-Company {
        param ([hashtable]$companyBaseOUMap)
        Write-Host "`nSélectionnez une compagnie :"
        $companies = $companyBaseOUMap.Keys | Sort-Object
        $selectedCompany = $companies | Out-GridView -Title "Sélectionnez une compagnie" -OutputMode Single
        if (-not $selectedCompany) {
            Write-Host "Aucune compagnie sélectionnée. Abandon." -ForegroundColor Red
            exit
        }
        return $selectedCompany
    }

    function Select-SubOU {
        param ([string]$BaseOU)
        Write-Host "`nSélectionnez une sous-OU :"
        $subOUs = Get-ADOrganizationalUnit -Filter * -SearchBase $BaseOU -SearchScope Subtree |
            Select-Object @{Name='OUName';Expression={ $_.Name }},
                          @{Name='DistinguishedName';Expression={ $_.DistinguishedName }}

        if ($subOUs.Count -eq 0) {
            Write-Host "Aucune sous-OU trouvée sous '$BaseOU'. Abandon." -ForegroundColor Red
            return $null
        }

        $selectedSubOU = $subOUs | Out-GridView -Title "Sélectionnez une sous-OU" -OutputMode Single
        if (-not $selectedSubOU) {
            Write-Host "Aucune sous-OU sélectionnée. Abandon." -ForegroundColor Red
            return $null
        }
        Write-Host "Vous avez sélectionné l'OU : $($selectedSubOU.OUName)" -ForegroundColor Green
        return $selectedSubOU.DistinguishedName
    }

    function Select-Manager {
        param ([array]$allManagers)
        Write-Host "`nSélectionnez le manager :"
        $selectedManager = $allManagers | Out-GridView -Title "Sélectionnez le manager" -OutputMode Single
        return $selectedManager
    }

    function Select-PCType {
        param ([array]$pcTypes)
        Write-Host "`nSélectionnez le type de PC :"
        $selectedPCType = $pcTypes | Out-GridView -Title "Sélectionnez le type de PC" -OutputMode Single
        if (-not $selectedPCType) {
            Write-Host "Aucun type de PC sélectionné. Abandon." -ForegroundColor Red
            exit
        }
        return $selectedPCType
    }

    function Select-Domain {
        param ([array]$companyDomains)
        Write-Host "`nSélectionnez un domaine pour l'email principal :"
        $selectedDomain = $companyDomains | Out-GridView -Title "Sélectionnez un domaine pour l'email principal" -OutputMode Single
        if (-not $selectedDomain) {
            Write-Host "Aucun domaine principal sélectionné. Abandon." -ForegroundColor Red
            exit
        }
        return $selectedDomain
    }

    # -----------------------------------------------------------------
    # Fonctions pour vérifier/créer des OUs
    # -----------------------------------------------------------------
    function Ensure-OUExists {
        param ([string]$FullOUDN)
        try {
            $existingOU = Get-ADOrganizationalUnit -Filter { DistinguishedName -eq $FullOUDN } -ErrorAction SilentlyContinue
            if ($existingOU) {
                Write-Host "L'OU '$FullOUDN' existe déjà." -ForegroundColor Green
                return $existingOU.DistinguishedName
            }

            # Si l'OU n'existe pas, on peut la créer récursivement si nécessaire
            # (Ici, simplifié, à adapter selon vos besoins réels)
            New-ADOrganizationalUnit -Name ($FullOUDN -replace '^OU=','' -replace ',.*$','') -Path ($FullOUDN -replace '^OU=[^,]+,') -ProtectedFromAccidentalDeletion $false
            Write-Host "OU créée : $FullOUDN" -ForegroundColor Green
            return $FullOUDN
        }
        catch {
            Write-Error "Erreur lors de la vérification/création de l'OU '$FullOUDN' : $_"
            return $null
        }
    }

    # -----------------------------------------------------------------
    # 1) Demande du NOM et du PRÉNOM
    # -----------------------------------------------------------------
    $LastName = Read-Host -Prompt "Entrer le nom"
    $FirstName = Read-Host -Prompt "Entrer le prénom"

    # -----------------------------------------------------------------
    # 2) Téléphones (optionnel)
    # -----------------------------------------------------------------
    $yesVariants = @("o", "ou", "oui", "ouai", "ouais")

    # Téléphone mobile ?
    $useMobilePhone = Read-Host "Voulez-vous ajouter un numéro de téléphone mobile ? (Oui/Non)"
    if ($yesVariants -contains $useMobilePhone.ToLower()) {
        do {
            $Phone = Read-Host -Prompt "Entrer le numéro de téléphone mobile (format : 06 12 34 56 78)"
            if ($Phone -notmatch "^\d{2} \d{2} \d{2} \d{2} \d{2}$") {
                Write-Host "Format invalide. Le format doit être : 06 12 34 56 78" -ForegroundColor Red
                $Phone = $null
            }
        } while (-not $Phone)
    }
    else {
        $Phone = $null
    }

    # Téléphone fixe ?
    $useFixedPhone = Read-Host "Voulez-vous ajouter un numéro de téléphone fixe ? (Oui/Non)"
    if ($yesVariants -contains $useFixedPhone.ToLower()) {
        do {
            $HomePhone = Read-Host -Prompt "Entrer le numéro de téléphone fixe (format : 01 23 45 67 89)"
            if ($HomePhone -notmatch "^\d{2} \d{2} \d{2} \d{2} \d{2}$") {
                Write-Host "Format invalide. Le format doit être : 01 23 45 67 89" -ForegroundColor Red
                $HomePhone = $null
            }
        } while (-not $HomePhone)
    }
    else {
        $HomePhone = $null
    }

    # -----------------------------------------------------------------
    # 3) Sélection de la Compagnie
    # -----------------------------------------------------------------
    if (Get-Command Out-GridView -ErrorAction SilentlyContinue) {
        $company = $companyDomainsMap.Keys | Out-GridView -Title "Sélectionnez une compagnie" -OutputMode Single
    }
    else {
        Write-Host "Out-GridView n'est pas disponible. Sélection alternative..." -ForegroundColor Yellow
        $company = Select-Company -companyBaseOUMap $companyBaseOUMap
    }
    if (-not $company) {
        Write-Host "Aucune compagnie sélectionnée. Abandon." -ForegroundColor Red
        return
    }
    Write-Host "Compagnie sélectionnée : $company" -ForegroundColor Green

    # -----------------------------------------------------------------
    # 4) Sélection du Site (adresse, city, etc.)
    # -----------------------------------------------------------------
    $companySiteInfo = $sitesInfo[$company]
    if (-not $companySiteInfo -or -not $companySiteInfo.ContainsKey("Sites")) {
        Write-Host "Aucun site défini pour la compagnie '$company'. Abandon." -ForegroundColor Red
        return
    }
    $availableSites = $companySiteInfo.Sites
    $globalOffice  = if ($companySiteInfo.ContainsKey("Office")) { $companySiteInfo.Office } else { $null }

    # Formatage pour Out-GridView
    $sitesFormatted = $availableSites | ForEach-Object {
        [PSCustomObject]@{
            "Adresse"     = $_.Street
            "Code Postal" = $_.PostalCode
            "Ville"       = $_.City
            "Office"      = $_.Office
        }
    }

    # Si plusieurs sites, on laisse choisir
    if ($sitesFormatted.Count -gt 1) {
        $selectedSite = $sitesFormatted | Out-GridView -Title "Sélectionnez un site pour '$company'" -OutputMode Single
    }
    else {
        $selectedSite = $sitesFormatted[0]
    }

    if (-not $selectedSite) {
        Write-Host "Aucun site sélectionné. Abandon." -ForegroundColor Red
        return
    }
    Write-Host "Site sélectionné : $($selectedSite.Ville)" -ForegroundColor Green

    # Récup info du site
    $Street     = $selectedSite.Adresse
    $PostalCode = $selectedSite.'Code Postal'
    $City       = $selectedSite.Ville

    # Nom d'office
    if ($selectedSite.Office -and $selectedSite.Office -ne "") {
        $officeName = $selectedSite.Office
    }
    else {
        $officeName = $globalOffice
    }
    Write-Host "Nom d'office sélectionné : $officeName" -ForegroundColor Green

    # Vérification qu'on a bien les domaines pour cette compagnie
    if (-not $companyDomainsMap.ContainsKey($company)) {
        Write-Host "La compagnie '$company' n'est pas définie dans le mapping des domaines. Abandon." -ForegroundColor Red
        return
    }

    # -----------------------------------------------------------------
    # 5) Gestion de l'OU de base et sous-OU
    # -----------------------------------------------------------------
    $baseOU = $companyBaseOUMap[$company]
    $baseOU = Ensure-OUExists -FullOUDN $baseOU
    if (-not $baseOU) {
        Write-Host "Impossible de créer ou de retrouver l'OU de base. Abandon." -ForegroundColor Red
        return
    }

    $selectedOUDN = Select-SubOU -BaseOU $baseOU
    if (-not $selectedOUDN) {
        Write-Host "Aucune sous-OU sélectionnée. Abandon." -ForegroundColor Red
        return
    }
    Write-Host "OU sélectionnée : $selectedOUDN" -ForegroundColor Green

    # --- Si besoin de corrections spécifiques du DN, adaptez ici ---
    # (exemple de correction de doublons)
    # $selectedOUDN = $selectedOUDN -replace '^OU=OU=', 'OU='

    # -----------------------------------------------------------------
    # 6) Sélection du Département
    # -----------------------------------------------------------------
    if ($companyDepartments.ContainsKey($company) -and $companyDepartments[$company].Count -gt 0) {
        $availableDepartments = $companyDepartments[$company]
        if (Get-Command Out-GridView -ErrorAction SilentlyContinue) {
            $selectedDepartment = $availableDepartments | Out-GridView -Title "Sélectionnez un département pour $company" -OutputMode Single
        }
        else {
            Write-Host "Out-GridView n'est pas disponible. Veuillez saisir le département parmi la liste suivante :" -ForegroundColor Yellow
            $i = 1
            foreach ($dept in $availableDepartments) {
                Write-Host "$i) $dept"
                $i++
            }
            $index = Read-Host "Entrez le numéro du département souhaité"
            if ($index -match '^\d+$' -and $index -ge 1 -and $index -le $availableDepartments.Count) {
                $selectedDepartment = $availableDepartments[$index - 1]
            }
            else {
                Write-Host "Sélection invalide. Abandon." -ForegroundColor Red
                return
            }
        }
        Write-Host "Département sélectionné : $selectedDepartment" -ForegroundColor Green
    }
    else {
        $selectedDepartment = Read-Host "Aucun département prédéfini pour $company. Entrez le département manuellement"
    }
    $department = $selectedDepartment
    Write-Host "Département final assigné : $department" -ForegroundColor Green

    # -----------------------------------------------------------------
    # 7) Demande du Service/Emploi (pour Title, Description) --- FIX 3
    #    On demande AVANT la création de l'utilisateur
    # -----------------------------------------------------------------
    $Service = Read-Host -Prompt "Entrer l'emploi (intitulé de poste / service)"

    # -----------------------------------------------------------------
    # 8) Demande du mot de passe
    # -----------------------------------------------------------------
    $Password = Read-Host -Prompt "Entrer le mot de passe de l'utilisateur" -AsSecureString
    if (-not $Password) {
        $Password = ConvertTo-SecureString "MotDePasseDefaut123!" -AsPlainText -Force
    }

    # -----------------------------------------------------------------
    # 9) Génération du SamAccountName unique
    # -----------------------------------------------------------------
    $FirstNameClean     = ($FirstName -replace '[-\s]', '').ToLower()
    $LastNameFormatted  = ($LastName -replace '[-\s]', '').ToLower()

    $minPrefixLength = 2
    $maxPrefixLength = $FirstNameClean.Length
    $prefixLength    = $minPrefixLength
    $UserNameFound   = $false

    while (-not $UserNameFound -and $prefixLength -le $maxPrefixLength) {
        $prefix        = $FirstNameClean.Substring(0, $prefixLength)
        $tempUserName  = $prefix + $LastNameFormatted
        $existingUser  = Get-ADUser -Filter "SamAccountName -eq '$tempUserName'" -ErrorAction SilentlyContinue

        if ($existingUser) {
            $prefixLength++
        }
        else {
            $UserName      = $tempUserName
            $UserNameFound = $true
        }
    }

    if (-not $UserNameFound) {
        # Ajout d'un numéro incrémental si on est arrivé au bout
        $number = 1
        do {
            $tempUserName = $FirstNameClean + $LastNameFormatted + $number
            $existingUser = Get-ADUser -Filter "SamAccountName -eq '$tempUserName'" -ErrorAction SilentlyContinue
            if ($existingUser) {
                $number++
            }
        } while ($existingUser)

        $UserName = $tempUserName
    }

    # Définition du UPN et DisplayName
    $UPN         = "$UserName@MYDOMAIN.COM"  # <-- Adaptez à votre vrai domaine interne
    $DisplayName = "$LastName $FirstName"

    Write-Host "`n== Récapitulatif Avant Création :" -ForegroundColor Cyan
    Write-Host " Nom             : $LastName"
    Write-Host " Prénom          : $FirstName"
    Write-Host " SamAccountName  : $UserName"
    Write-Host " UPN             : $UPN"
    Write-Host " DisplayName     : $DisplayName"
    Write-Host " Service (Title) : $Service"

    # -----------------------------------------------------------------
    # 10) Création de l'utilisateur dans AD
    # -----------------------------------------------------------------
    if ([string]::IsNullOrEmpty($selectedOUDN)) {
        Write-Host "OU non valide. Abandon." -ForegroundColor Red
        return
    }
    try {
        New-ADUser `
            -Name              $DisplayName `
            -GivenName         $FirstName `
            -Surname           $LastName `
            -SamAccountName    $UserName `
            -UserPrincipalName $UPN `
            -Description       $Service `
            -DisplayName       $DisplayName `
            -AccountPassword   $Password `
            -Enabled           $true `
            -ChangePasswordAtLogon $true `
            -Path              $selectedOUDN `
            -Title             $Service

        Write-Host "Utilisateur '$UserName' créé avec succès dans l'OU : $selectedOUDN" -ForegroundColor Green

        $newUser = Get-ADUser -Filter { SamAccountName -eq $UserName } -Properties DistinguishedName
        Write-Host "Utilisateur créé à : $($newUser.DistinguishedName)" -ForegroundColor Cyan

        # Exemple : Ajout à un groupe par défaut
        $disabledGroupDN = "CN=Avant De Cliquer,OU=GROUPE,DC=MYDOMAIN,DC=COM"
        try {
            Add-ADGroupMember -Identity $disabledGroupDN -Members $newUser.DistinguishedName -ErrorAction Stop
            Write-Host "Utilisateur ajouté au groupe 'Avant de cliquer'." -ForegroundColor Green
        }
        catch {
            Write-Host "Erreur lors de l'ajout au groupe 'Avant de cliquer' : $_" -ForegroundColor Red
        }
    }
    catch {
        Write-Host "Erreur lors de la création de l'utilisateur : $_" -ForegroundColor Red
        return
    }

    # -----------------------------------------------------------------
    # 11) Sélection du Manager
    # -----------------------------------------------------------------
    $allManagers = Get-ADUser -Filter * -Properties Name, SamAccountName | Select-Object Name, SamAccountName
    if (Get-Command Out-GridView -ErrorAction SilentlyContinue) {
        $managerSelection = $allManagers | Out-GridView -Title "Sélectionnez le manager" -OutputMode Single
    }
    else {
        Write-Host "Out-GridView n'est pas disponible. Méthode alternative..." -ForegroundColor Yellow
        $managerSelection = Select-Manager -allManagers $allManagers
    }

    if ($managerSelection) {
        $ManagerUserName = $managerSelection.SamAccountName
        $Manager         = Get-ADUser -Identity $ManagerUserName
        $ManagerDistinguishedName = $Manager.DistinguishedName
    }
    else {
        Write-Host "Aucun manager sélectionné. Il ne sera pas assigné." -ForegroundColor Yellow
        $ManagerDistinguishedName = $null
    }

    # -----------------------------------------------------------------
    # 12) Mise à jour des propriétés de l'utilisateur
    # -----------------------------------------------------------------
    try {
        Set-ADUser -Identity $newUser `
            -HomePhone   $HomePhone `
            -MobilePhone $Phone `
            -OfficePhone $HomePhone `
            -Title       $Service `
            -Description $Service `
            -Office      $officeName `
            -Company     $company `
            -StreetAddress $Street `
            -City        $City `
            -PostalCode  $PostalCode `
            -Department  $department `
            -Manager     $ManagerDistinguishedName

        Write-Host "`nInformations de l'utilisateur mises à jour avec succès !" -ForegroundColor Green
    }
    catch {
        Write-Host "Erreur lors de la mise à jour des informations : $_" -ForegroundColor Red
    }

    # -----------------------------------------------------------------
    # 13) Sélection du Domaine Principal (email)
    # -----------------------------------------------------------------
    $companyDomains = $companyDomainsMap[$company]
    if (Get-Command Out-GridView -ErrorAction SilentlyContinue) {
        $selectedEmailDomain = $companyDomains | Out-GridView -Title "Sélectionnez un domaine pour l'email principal" -OutputMode Single
    }
    else {
        Write-Host "Out-GridView n'est pas disponible. Sélection alternative..." -ForegroundColor Yellow
        $selectedEmailDomain = Select-Domain -companyDomains $companyDomains
    }

    if (-not $selectedEmailDomain) {
        Write-Host "Aucun domaine principal sélectionné. Abandon." -ForegroundColor Red
        return
    }

    $Email = "$UserName@$selectedEmailDomain"
    Write-Host "Email principal : $Email" -ForegroundColor Green

    # -----------------------------------------------------------------
    # 14) Gestion des alias
    # -----------------------------------------------------------------
    $aliasPrompt = @("Oui", "Non")
    $aliasChoice = $aliasPrompt | Out-GridView -Title "Voulez-vous ajouter des alias ?" -OutputMode Single

    $selectedAliases = @()
    if ($aliasChoice -eq "Oui") {
        $aliasOptions = @()
        foreach ($domain in $aliasDomainsList) {
            $aliasOptions += "$UserName@$domain"
        }

        if (Get-Command Out-GridView -ErrorAction SilentlyContinue) {
            $selectedAliases = $aliasOptions | Out-GridView -Title "Sélectionnez les alias" -OutputMode Multiple
        }
        else {
            Write-Host "Out-GridView n'est pas disponible. Sélection manuelle..." -ForegroundColor Yellow
            for ($i = 0; $i -lt $aliasOptions.Count; $i++) {
                Write-Host "$($i+1)) $($aliasOptions[$i])"
            }
            $selectedIndices = Read-Host "Entrez les numéros des alias à ajouter (séparés par virgules)"
            $selectedIndices = $selectedIndices -split ',' | ForEach-Object { $_.Trim() }
            foreach ($index in $selectedIndices) {
                if ($index -match '^\d+$') {
                    $idx = [int]$index - 1
                    if ($idx -ge 0 -and $idx -lt $aliasOptions.Count) {
                        $selectedAliases += $aliasOptions[$idx]
                    }
                }
            }
        }

        if ($selectedAliases.Count -eq 0) {
            Write-Host "Aucun alias sélectionné." -ForegroundColor Yellow
        }
        else {
            Write-Host "Alias sélectionnés : $($selectedAliases -join ', ')" -ForegroundColor Green
        }
    }
    else {
        Write-Host "Aucun alias ne sera ajouté." -ForegroundColor Yellow
    }

    # -----------------------------------------------------------------
    # 15) Assignation des groupes Office 365
    # -----------------------------------------------------------------
    # --- Exemple de groupes : à adapter à votre orga ---
    $defaultGroupDN = "CN=Office365-Basic,OU=Groupes,DC=MYDOMAIN,DC=COM" 
    $mainGroups = @{
        "1" = "CN=Office365-E3,OU=Groupes,DC=MYDOMAIN,DC=COM"
        # Ajoutez d'autres groupes si besoin
    }
    # Groupe optionnel (par ex. Visio Plan 2)
    $optionalGroupDN = "CN=VisioPlan2,OU=Groupes,DC=MYDOMAIN,DC=COM"

    try {
        Add-ADGroupMember -Identity $defaultGroupDN -Members $newUser.DistinguishedName
        Write-Host "Utilisateur ajouté au groupe par défaut : $defaultGroupDN" -ForegroundColor Green
    }
    catch {
        Write-Host "Erreur lors de l'ajout au groupe par défaut : $_" -ForegroundColor Red
    }

    Write-Host "`nSélectionnez une licence Office 365 pour l'utilisateur :" -ForegroundColor Cyan

    $groupOptions = $mainGroups.Keys | ForEach-Object {
        [PSCustomObject]@{
            Numéro  = $_
            Licence = ($mainGroups[$_] -split ",")[0].Substring(3)  # Récup le 'CN='
        }
    }

    if (Get-Command Out-GridView -ErrorAction SilentlyContinue) {
        $selectedGroup = $groupOptions | Out-GridView -Title "Sélectionnez une licence Office 365" -OutputMode Single
    }
    else {
        $i = 1
        foreach ($option in $groupOptions) {
            Write-Host "$i) $($option.Licence)" -ForegroundColor Yellow
            $i++
        }
        $selectedIndex = Read-Host "Entrez le numéro correspondant à la licence souhaitée"
        if ($selectedIndex -match '^\d+$') {
            $idx = [int]$selectedIndex
            if ($idx -ge 1 -and $idx -le $groupOptions.Count) {
                $selectedGroup = $groupOptions[$idx - 1]
            }
        }
    }

    if ($selectedGroup) {
        $selectedGroupDN = $mainGroups[$selectedGroup.Numéro]
        try {
            Add-ADGroupMember -Identity $selectedGroupDN -Members $newUser.DistinguishedName
            Write-Host "Utilisateur ajouté au groupe : $selectedGroupDN" -ForegroundColor Green
        }
        catch {
            Write-Error "Erreur lors de l'ajout au groupe '$selectedGroupDN' : $_"
        }
    }
    else {
        Write-Host "Aucun groupe de licence sélectionné." -ForegroundColor Yellow
    }

    $visioChoice = @("Oui", "Non") | Out-GridView -Title "Voulez-vous ajouter 'Visio Plan2' ?" -OutputMode Single
    if ($visioChoice -eq "Oui") {
        try {
            Add-ADGroupMember -Identity $optionalGroupDN -Members $newUser.DistinguishedName
            Write-Host "Utilisateur ajouté au groupe optionnel : $optionalGroupDN" -ForegroundColor Green
        }
        catch {
            Write-Host "Erreur lors de l'ajout au groupe optionnel : $_" -ForegroundColor Red
        }
    }
    else {
        Write-Host "Aucun groupe optionnel ajouté." -ForegroundColor Yellow
    }

    # -----------------------------------------------------------------
    # 16) Sélection du Type de PC
    # -----------------------------------------------------------------
    $pcTypes = @("PC Fixe", "PC Portable", "Client léger")
    if (Get-Command Out-GridView -ErrorAction SilentlyContinue) {
        $selectedPCType = $pcTypes | Out-GridView -Title "Sélectionnez le type de PC" -OutputMode Single
    }
    else {
        Write-Host "Out-GridView n'est pas disponible. Méthode alternative..." -ForegroundColor Yellow
        $selectedPCType = Select-PCType -pcTypes $pcTypes
    }

    if (-not $selectedPCType) {
        Write-Host "Aucun type de PC sélectionné. Abandon." -ForegroundColor Red
        return
    }
    Write-Host "Type de PC sélectionné : $selectedPCType" -ForegroundColor Green

    # Groupes liés au PC
    $pcGroupsToAdd = @()

    if ($selectedPCType -eq "PC Fixe") {
        $mobileOptions = @("Avec téléphone mobile", "Sans téléphone mobile")
        $mobileChoice  = $mobileOptions | Out-GridView -Title "L'utilisateur dispose-t-il d'un téléphone mobile ?" -OutputMode Single
        if (-not $mobileChoice) {
            Write-Host "Aucune réponse pour le téléphone mobile. Abandon." -ForegroundColor Red
            return
        }

        if ($mobileChoice -eq "Avec téléphone mobile") {
            # Adaptez le DN
            $group = "CN=Utilisateurs_PC_Fixe_Mobile,OU=Chemin,DC=MYDOMAIN,DC=COM"
        }
        else {
            $group = "CN=Utilisateurs_PC_Fixe_SansMobile,OU=Chemin,DC=MYDOMAIN,DC=COM"
        }
        Write-Host "Ajout de l'utilisateur au groupe : $group" -ForegroundColor Cyan
        $pcGroupsToAdd += $group
    }
    elseif ($selectedPCType -eq "PC Portable") {
        $group = "CN=Utilisateurs_PC_Portable,OU=Chemin,DC=MYDOMAIN,DC=COM"
        Write-Host "Ajout de l'utilisateur au groupe : $group" -ForegroundColor Cyan
        $pcGroupsToAdd += $group

        # Exemple : ajout d'un groupe VPN
        $baseVPNGroup = "CN=Utilisateurs_VPNDeBase,OU=Chemin,DC=MYDOMAIN,DC=COM"
        Write-Host "Ajout de l'utilisateur au groupe : $baseVPNGroup" -ForegroundColor Cyan
        $pcGroupsToAdd += $baseVPNGroup

        $vpnOptions = @{
            "Applicatif" = "CN=SSLVPN_Applicatif,OU=Chemin,DC=MYDOMAIN,DC=COM"
            "Intranet"   = "CN=SSLVPN_Intranet,OU=Chemin,DC=MYDOMAIN,DC=COM"
        }

        $selectedVPN = $vpnOptions.Keys | Out-GridView -Title "Sélectionnez le type de VPN" -OutputMode Single
        if ($selectedVPN) {
            $vpnGroup = $vpnOptions[$selectedVPN]
            Write-Host "VPN sélectionné : $selectedVPN" -ForegroundColor Green
            Write-Host "Ajout de l'utilisateur au groupe : $vpnGroup" -ForegroundColor Cyan
            $pcGroupsToAdd += $vpnGroup
        }
        else {
            Write-Host "Aucun VPN sélectionné. Abandon." -ForegroundColor Red
            return
        }
    }
    elseif ($selectedPCType -eq "Client léger") {
        $group = "CN=Utilisateurs_ClientLeger,OU=Chemin,DC=MYDOMAIN,DC=COM"
        Write-Host "Ajout de l'utilisateur au groupe : $group" -ForegroundColor Cyan
        $pcGroupsToAdd += $group
    }

    # -----------------------------------------------------------------
    # 17) Option d’héritage des droits d’un autre utilisateur
    # -----------------------------------------------------------------
    $inheritRightsOptions = @("Oui", "Non")
    $inheritChoice = $inheritRightsOptions | Out-GridView -Title "Hériter des droits d'un autre utilisateur ?" -OutputMode Single

    $inheritRights = ($inheritChoice -eq "Oui")

    $inheritGroupsToAdd = @()

    if ($inheritRights) {
        $allUsers = Get-ADUser -Filter * -Properties Name, SamAccountName, MemberOf |
                    Select-Object Name, SamAccountName, MemberOf

        $sourceUser = $allUsers | Out-GridView -Title "Sélectionnez l'utilisateur source" -OutputMode Single
        if ($sourceUser) {
            $sourceUserDetails = Get-ADUser -Identity $sourceUser.SamAccountName -Properties MemberOf
            if ($sourceUserDetails.MemberOf) {
                foreach ($groupDN in $sourceUserDetails.MemberOf) {
                    if ($groupDN -match '^CN=([^,]+)') {
                        $groupName = $matches[1]
                        # On peut filtrer certains groupes pour ne pas les copier
                        if ($groupName -notmatch '(?i)(mfa|office|vpn|pc)') {
                            $inheritGroupsToAdd += $groupDN
                        }
                    }
                    else {
                        $inheritGroupsToAdd += $groupDN
                    }
                }
            }
        }
        else {
            Write-Host "Aucun utilisateur source sélectionné. Héritage ignoré." -ForegroundColor Yellow
        }
    }

    # -----------------------------------------------------------------
    # 18) Ajout de l’utilisateur aux groupes PC / VPN
    # -----------------------------------------------------------------
    foreach ($grp in $pcGroupsToAdd) {
        try {
            Add-ADGroupMember -Identity $grp -Members $newUser.DistinguishedName
            Write-Host "Utilisateur ajouté au groupe : $grp" -ForegroundColor Green
        }
        catch {
            Write-Host "Erreur d'ajout dans $grp : $_" -ForegroundColor Red
        }
    }

    # -----------------------------------------------------------------
    # 19) Héritage des droits (si demandé)
    # -----------------------------------------------------------------
    if ($inheritGroupsToAdd.Count -gt 0) {
        foreach ($groupDN in $inheritGroupsToAdd) {
            try {
                Add-ADGroupMember -Identity $groupDN -Members $newUser.DistinguishedName
                Write-Host "Utilisateur ajouté au groupe hérité : $groupDN" -ForegroundColor Green
            }
            catch {
                Write-Host "Erreur d'ajout dans $groupDN : $_" -ForegroundColor Red
            }
        }

        # Exemple : si on voulait aussi copier certaines propriétés
        # $propertiesToCopy = @("Department","Title","Description")
        # foreach ($property in $propertiesToCopy) {
        #     $value = (Get-ADUser -Identity $sourceUser.SamAccountName -Properties $property).$property
        #     if ($value) {
        #         Set-ADUser -Identity $newUser -Replace @{ $property = $value }
        #         Write-Host "Propriété copiée : $property = $value" -ForegroundColor Cyan
        #     }
        # }
    }

    # -----------------------------------------------------------------
    # 20) Ajout des alias et définition de l'email principal (ProxyAddresses)
    # -----------------------------------------------------------------
    try {
        # Mettre l'adresse principale en SMTP: (majuscules => principal)
        Set-ADUser -Identity $newUser -EmailAddress $Email -Replace @{ProxyAddresses = @("SMTP:$Email")}

        # Ajout des alias en smtp:
        foreach ($alias in $selectedAliases) {
            Set-ADUser -Identity $newUser -Add @{ProxyAddresses = "smtp:$alias"}
        }

        Write-Host "Email principal et alias mis à jour avec succès." -ForegroundColor Green
    }
    catch {
        Write-Host "Erreur lors de l'ajout des alias email : $_" -ForegroundColor Red
    }

    # -----------------------------------------------------------------
    # 21) Fin du processus
    # -----------------------------------------------------------------
    Write-Host "`nCréation de l'utilisateur terminée avec succès !" -ForegroundColor Green
}
<#
============================================================================
 Script : Import d’utilisateurs AD à partir d’un CSV
 Auteur : (Votre Nom)
 Date   : (Date de dernière modification)
 Objectif :
   - Charger le module ActiveDirectory et éventuellement AzureAD
   - Importer un fichier CSV listant des utilisateurs
   - Créer et configurer chaque utilisateur dans AD :
       * OU dynamique
       * Manager
       * Téléphone / Email / Alias / ProxyAddresses
       * Groupes en fonction du type de PC / VPN
       * (Optionnel) Copie de droits d’un utilisateur source
       * (Optionnel) Assignation de licences Office 365 via AzureAD
============================================================================
#>

# --- FIX 1 : Préparation de l'environnement ---

# Forcer l’utilisation du protocole TLS 1.2 pour les connexions
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Génération d’un identifiant unique pour la session d’import
$importSessionId = Get-Date -Format 'yyyyMMdd_HHmmss'

# Définition du chemin du fichier log
$logFile = "C:\Logs\CreateUserLog_$importSessionId.txt"

# Fonction de logging
function Write-Log {
    param (
        [string]$Message,
        [string]$Color = "White"
    )

    # Affichage console
    Write-Host $Message -ForegroundColor $Color

    # Écriture dans le log
    Add-Content -Path $logFile -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - [$importSessionId] - $Message"
}

# --- FIX 2 : Chargement des modules ---

function Load-Modules {
    # Charger le module ActiveDirectory si pas déjà importé
    if (-not (Get-Module ActiveDirectory)) {
        try {
            Import-Module ActiveDirectory -ErrorAction Stop
            Write-Log "Module ActiveDirectory chargé avec succès." "Green"
        }
        catch {
            Write-Log "Impossible de charger le module ActiveDirectory (RSAT/ADDS Tools requis)." "Red"
            throw
        }
    }

    # Charger le module AzureAD si souhaité
    $AzureADLoaded = $false
    if (-not (Get-Module AzureAD)) {
        try {
            Import-Module AzureAD -ErrorAction Stop
            $AzureADLoaded = $true
            Write-Log "Module AzureAD chargé avec succès." "Green"
        }
        catch {
            Write-Log "Impossible de charger le module AzureAD. Les licences Office 365 ne seront pas assignées." "Yellow"
            $AzureADLoaded = $false
        }
    }
    else {
        $AzureADLoaded = $true
    }

    # Tentative de connexion au module AzureAD
    if ($AzureADLoaded) {
        try {
            Connect-AzureAD -ErrorAction Stop
            Write-Log "Connecté à AzureAD avec succès." "Green"
        }
        catch {
            Write-Log "Erreur de connexion à AzureAD : $($_.Exception.Message)" "Red"
            $AzureADLoaded = $false
        }
    }

    return $AzureADLoaded
}

# Charger les modules dès le démarrage
$AzureADLoaded = Load-Modules

# --- Variables et mappings globaux ---

# DN du groupe Office 365 (exemple)
$Office365GroupDN = "CN=Office 365,OU=Groups,DC=MYDOMAIN,DC=COM"

# Mappage des sites (exemple)
$sitesInfo = @{
    "Site-A" = @{ Street = "Rue du Général"; PostalCode = "75001"; City = "Paris";    Company = "CompagnieA" }
    "Site-B" = @{ Street = "Avenue XYZ";     PostalCode = "69001"; City = "Lyon";     Company = "CompagnieB" }
    # ...
}

# Mapping entre les Companies et leurs OUs (exemples)
$OUs = @{
    "CompagnieA" = @(
        "OU=Direction,OU=Paris,DC=MYDOMAIN,DC=COM",
        "OU=Services,OU=Paris,DC=MYDOMAIN,DC=COM"
    )
    "CompagnieB" = @(
        "OU=Direction,OU=Lyon,DC=MYDOMAIN,DC=COM",
        "OU=Services,OU=Lyon,DC=MYDOMAIN,DC=COM"
    )
}

# Mapping des groupes PC
$PcGroups = @{
    "PC Fixe"      = @("CN=Utilisateurs_PC_Fixe,OU=PC,DC=MYDOMAIN,DC=COM")
    "PC Portable"  = @("CN=Utilisateurs_PC_Portable,OU=PC,DC=MYDOMAIN,DC=COM")
    "Client léger" = @("CN=Utilisateurs_ClientLeger,OU=PC,DC=MYDOMAIN,DC=COM")
}

# Mapping VPN
$VpnGroups = @{
    "Intranet / Extranet" = "CN=SSLVPN_Applicatif,OU=VPN,DC=MYDOMAIN,DC=COM"
    # Ajoutez d'autres mappings au besoin
}

# Liste de départements
$departments = @("Accueil", "Achats", "Informatique", "Comptabilité")

# --- Fonctions utilitaires ---

function Select-Department {
    param (
        [array]$Departments
    )
    if (Get-Command Out-GridView -ErrorAction SilentlyContinue) {
        $selectedDepartment = $Departments | Out-GridView -Title "Sélectionnez un département" -OutputMode Single
    }
    else {
        Write-Log "Out-GridView n'est pas disponible. Sélection par saisie..." "Yellow"
        $i = 1
        foreach ($dept in $Departments) {
            Write-Host "$i) $dept"
            $i++
        }
        $choice = Read-Host "Entrez le numéro du département"
        if ($choice -match '^\d+$') {
            $index = [int]$choice - 1
            if ($index -ge 0 -and $index -lt $Departments.Count) {
                $selectedDepartment = $Departments[$index]
            }
        }
    }

    if (-not $selectedDepartment) {
        Write-Log "Aucun département sélectionné." "Red"
        return $null
    }
    return $selectedDepartment
}

function Select-OU {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Company
    )

    if ([string]::IsNullOrWhiteSpace($Company)) {
        Write-Log "Company non définie, utilisation de l'OU par défaut." "Yellow"
        return "OU=Stagiaires,DC=MYDOMAIN,DC=COM"
    }

    if (-not $OUs.ContainsKey($Company)) {
        Write-Log "Aucune OU définie pour la Company '$Company'. Utilisation de l'OU par défaut." "Yellow"
        return "OU=Stagiaires,DC=MYDOMAIN,DC=COM"
    }

    $availableOUs = $OUs[$Company]

    # Construction d'objets pour Out-GridView
    $ouObjects = $availableOUs | ForEach-Object {
        $OUName = ($_ -split ",")[0] -replace "^OU=", ""
        [PSCustomObject]@{
            DisplayName        = $OUName
            DistinguishedName  = $_
        }
    }

    if (Get-Command Out-GridView -ErrorAction SilentlyContinue) {
        $selectedDisplayName = $ouObjects.DisplayName | Out-GridView -Title "Sélectionnez une OU pour la Company '$Company'" -OutputMode Single
    }
    else {
        Write-Log "Out-GridView indisponible. Sélection manuelle." "Yellow"
        $i = 1
        foreach ($item in $ouObjects) {
            Write-Host "$i) $($item.DisplayName)"
            $i++
        }
        $choice = Read-Host "Entrez le numéro de l'OU"
        if ($choice -match '^\d+$') {
            $index = [int]$choice - 1
            if ($index -ge 0 -and $index -lt $ouObjects.Count) {
                $selectedDisplayName = $ouObjects[$index].DisplayName
            }
        }
    }

    if (-not $selectedDisplayName) {
        Write-Log "Aucune OU sélectionnée. Utilisation de l'OU par défaut." "Yellow"
        return "OU=Stagiaires,DC=MYDOMAIN,DC=COM"
    }

    $selectedOU = $ouObjects | Where-Object { $_.DisplayName -eq $selectedDisplayName }
    if ($selectedOU) {
        return $selectedOU.DistinguishedName
    }
    else {
        Write-Log "Erreur lors de la sélection de l'OU. Utilisation de l'OU par défaut." "Red"
        return "OU=Stagiaires,DC=MYDOMAIN,DC=COM"
    }
}

function Get-SamAccountName {
    param(
        [string]$FirstName,
        [string]$LastName
    )
    # Nettoyage basique
    $FirstName = $FirstName -replace "[^a-zA-Z0-9\-]", ""
    $LastName  = $LastName  -replace "[^a-zA-Z0-9\-]", ""

    if ($FirstName -match "-") {
        # Cas d’un prénom avec tiret
        $fnParts = $FirstName -split "-"
        $firstNamePart = ($fnParts | ForEach-Object { $_.Substring(0,1).ToLower() }) -join "-"
        $lnClean = ($LastName -replace "-", "").ToLower()
        if ($lnClean.Length -gt 1) {
            $lnClean = $lnClean.Substring(0, $lnClean.Length - 1)
        }
        return "$firstNamePart$lnClean"
    }
    else {
        # Prénom sans tiret
        $firstLetter = $FirstName.Substring(0,1).ToLower()
        $lnClean     = ($LastName -replace "-", "").ToLower()
        return "$firstLetter$lnClean"
    }
}

function Copy-UserRights {
    param (
        [Parameter(Mandatory = $true)]
        [string]$SourceUserSamAccountName,
        [Parameter(Mandatory = $true)]
        [string]$TargetUserSamAccountName
    )

    try {
        $sourceUser = Get-ADUser -Identity $SourceUserSamAccountName -Properties MemberOf
        if (-not $sourceUser) {
            Write-Log "Utilisateur source '$SourceUserSamAccountName' introuvable." "Red"
            return
        }

        $sourceGroups = $sourceUser.MemberOf
        if ($sourceGroups.Count -eq 0) {
            Write-Log "L’utilisateur source n’appartient à aucun groupe." "Yellow"
            return
        }

        foreach ($groupDN in $sourceGroups) {
            # Exclure certains groupes sensibles
            $excludedGroups = @(
                "CN=Domain Admins,CN=Users,DC=MYDOMAIN,DC=COM",
                "CN=Enterprise Admins,CN=Users,DC=MYDOMAIN,DC=COM",
                "CN=Schema Admins,CN=Users,DC=MYDOMAIN,DC=COM"
            )
            if ($excludedGroups -contains $groupDN) {
                Write-Log "Exclusion du groupe sensible : $groupDN" "Yellow"
                continue
            }

            try {
                Add-ADGroupMember -Identity $groupDN -Members $TargetUserSamAccountName -ErrorAction Stop
                Write-Log "Ajouté au groupe : $groupDN" "Green"
            }
            catch {
                Write-Log "Erreur lors de l'ajout au groupe '$groupDN' : $($_.Exception.Message)" "Red"
            }
        }

        Write-Log "Droits copiés de '$SourceUserSamAccountName' vers '$TargetUserSamAccountName' !" "Green"
    }
    catch {
        Write-Log "Erreur lors de la copie des droits : $($_.Exception.Message)" "Red"
    }
}

function Add-UserToGroup {
    param (
        [Parameter(Mandatory = $true)]
        [string]$UserSamAccountName,
        [Parameter(Mandatory = $true)]
        [string]$GroupDN
    )

    try {
        $user = Get-ADUser -Identity $UserSamAccountName
        if (-not $user) {
            Write-Log "Utilisateur '$UserSamAccountName' non trouvé." "Red"
            return
        }

        $isMember = Get-ADGroupMember -Identity $GroupDN -Recursive | Where-Object { $_.SamAccountName -eq $UserSamAccountName }
        if ($isMember) {
            Write-Log "Utilisateur '$UserSamAccountName' déjà membre du groupe '$GroupDN'." "Yellow"
        }
        else {
            Add-ADGroupMember -Identity $GroupDN -Members $UserSamAccountName -ErrorAction Stop
            Write-Log "Utilisateur '$UserSamAccountName' ajouté au groupe '$GroupDN'." "Green"
        }
    }
    catch {
        Write-Log "Erreur d'ajout de l'utilisateur '$UserSamAccountName' au groupe '$GroupDN' : $($_.Exception.Message)" "Red"
    }
}

function Select-VPNGroup {
    param (
        [Parameter(Mandatory = $true)]
        [string]$UserName
    )
    $vpnGroupObjects = $VpnGroups.GetEnumerator() | ForEach-Object {
        [PSCustomObject]@{
            DisplayName       = $_.Key
            DistinguishedName = $_.Value
        }
    }

    if (Get-Command Out-GridView -ErrorAction SilentlyContinue) {
        $selectedVPNDisplayNames = $vpnGroupObjects.DisplayName | Out-GridView -Title "Sélectionnez les groupes VPN pour '$UserName'" -OutputMode Multiple
    }
    else {
        Write-Log "Out-GridView indisponible. Sélection VPN manuelle." "Yellow"
        # Sélection alternative (liste indices)
        $i = 1
        foreach ($g in $vpnGroupObjects) {
            Write-Host "$i) $($g.DisplayName)"
            $i++
        }
        $choices = Read-Host "Entrez les numéros séparés par des virgules"
        $selectedVPNDisplayNames = @()
        foreach ($choice in ($choices -split ",")) {
            if ($choice -match '^\d+$') {
                $idx = [int]$choice - 1
                if ($idx -ge 0 -and $idx -lt $vpnGroupObjects.Count) {
                    $selectedVPNDisplayNames += $vpnGroupObjects[$idx].DisplayName
                }
            }
        }
    }

    if (-not $selectedVPNDisplayNames -or $selectedVPNDisplayNames.Count -eq 0) {
        Write-Log "Aucun groupe VPN sélectionné." "Red"
        return
    }

    foreach ($displayName in $selectedVPNDisplayNames) {
        $selectedVPNGroup = $vpnGroupObjects | Where-Object { $_.DisplayName -eq $displayName }
        if ($selectedVPNGroup) {
            try {
                Add-ADGroupMember -Identity $selectedVPNGroup.DistinguishedName -Members $UserName -ErrorAction Stop
                Write-Log "Ajouté au groupe VPN : $($selectedVPNGroup.DisplayName)" "Green"
            }
            catch {
                Write-Log "Erreur lors de l'ajout au groupe VPN '$($selectedVPNGroup.DisplayName)' : $($_.Exception.Message)" "Red"
            }
        }
    }
}

function Select-PCAndAssignGroups {
    param (
        [Parameter(Mandatory = $true)]
        [string]$UserName
    )
    Write-Log "Sélection du type de PC pour '$UserName'..." "Cyan"

    $pcTypes = @("PC Fixe", "PC Portable", "Client léger")
    if (Get-Command Out-GridView -ErrorAction SilentlyContinue) {
        $selectedPCType = $pcTypes | Out-GridView -Title "Sélectionnez le type de PC pour '$UserName'" -OutputMode Single
    }
    else {
        Write-Log "Out-GridView indisponible. Sélection manuelle du type de PC." "Yellow"
        $i = 1
        foreach ($t in $pcTypes) {
            Write-Host "$i) $t"
            $i++
        }
        $choice = Read-Host "Entrez le numéro du type de PC"
        if ($choice -match '^\d+$') {
            $idx = [int]$choice - 1
            if ($idx -ge 0 -and $idx -lt $pcTypes.Count) {
                $selectedPCType = $pcTypes[$idx]
            }
        }
    }

    if (-not $selectedPCType) {
        Write-Log "Aucun type de PC sélectionné. Abandon." "Red"
        return
    }

    Write-Log "Type de PC sélectionné : $selectedPCType" "Green"

    if ($PcGroups.ContainsKey($selectedPCType)) {
        foreach ($group in $PcGroups[$selectedPCType]) {
            try {
                Add-ADGroupMember -Identity $group -Members $UserName -ErrorAction Stop
                Write-Log "Ajouté au groupe : $group" "Green"
            }
            catch {
                Write-Log "Erreur lors de l'ajout au groupe '$group' : $($_.Exception.Message)" "Red"
            }
        }
    }
    else {
        Write-Log "Aucun groupe défini pour le PC '$selectedPCType'." "Yellow"
    }

    if ($selectedPCType -eq "PC Portable") {
        $vpnOptions = @("Oui", "Non")
        if (Get-Command Out-GridView -ErrorAction SilentlyContinue) {
            $vpnRequired = $vpnOptions | Out-GridView -Title "L'utilisateur '$UserName' a-t-il besoin d'un VPN ?" -OutputMode Single
        }
        else {
            Write-Log "Out-GridView indisponible. Saisie manuelle pour le VPN." "Yellow"
            $vpnRequired = Read-Host "Besoin d’un VPN ? (Oui/Non)"
        }

        if ($vpnRequired -eq "Oui") {
            Select-VPNGroup -UserName $UserName
        }
        else {
            Write-Log "Aucun VPN requis pour '$UserName'." "Yellow"
        }
    }
}

function Assign-Office365License {
    param (
        [Parameter(Mandatory = $true)]
        [string]$UserUPN,
        [Parameter(Mandatory = $true)]
        [array]$Office365Packages
    )

    if (-not $AzureADLoaded) {
        Write-Log "Module AzureAD non chargé, impossible d’assigner des licences." "Red"
        return
    }

    foreach ($package in $Office365Packages) {
        try {
            $license = Get-AzureADSubscribedSku | Where-Object { $_.SkuPartNumber -eq $package }
            if ($license) {
                $azureUser = Get-AzureADUser -ObjectId $UserUPN
                if ($azureUser) {
                    Set-AzureADUserLicense -ObjectId $azureUser.ObjectId -AssignedLicenses @{AddLicenses = $license.ObjectId; RemoveLicenses = @()}
                    Write-Log "Licence '$package' assignée à '$UserUPN'." "Green"
                }
                else {
                    Write-Log "Utilisateur AzureAD '$UserUPN' introuvable." "Red"
                }
            }
            else {
                Write-Log "Licence '$package' non trouvée dans AzureAD." "Red"
            }
        }
        catch {
            Write-Log "Erreur lors de l’assignation de la licence '$package' à '$UserUPN' : $($_.Exception.Message)" "Red"
        }
    }
}

# --- Fonction d'import CSV ---

function Import-UserData {
    param (
        [Parameter(Mandatory = $true)]
        [string]$csvFilePath
    )

    if (-not (Test-Path $csvFilePath)) {
        Write-Log "Le fichier CSV '$csvFilePath' n’existe pas !" "Red"
        throw "Fichier CSV introuvable."
    }

    try {
        $users = Import-Csv -Path $csvFilePath
        Write-Log "Importation des utilisateurs depuis le CSV réussie." "Green"
    }
    catch {
        Write-Log "Erreur lors de l’importation du CSV : $($_.Exception.Message)" "Red"
        throw "Échec de l’importation du CSV."
    }

    try {
        $allUsers = Get-ADUser -Filter * -Properties SamAccountName, GivenName, Surname, DistinguishedName
        Write-Log "Récupération de la liste des utilisateurs AD réussie." "Green"
    }
    catch {
        Write-Log "Erreur lors de la récupération des utilisateurs AD : $($_.Exception.Message)" "Red"
        throw "Échec de la récupération des utilisateurs AD."
    }

    return @{ Users = $users; AllUsers = $allUsers }
}

# --- Fonction principale : création d’utilisateurs depuis le CSV ---

function Create-UserFromCSV {
    param (
        [Parameter(Mandatory = $true)]
        [string]$csvFilePath
    )

    $executingUser = $env:USERNAME
    Write-Log "Début de l’importation des utilisateurs par '$executingUser'." "Yellow"

    $importData = Import-UserData -csvFilePath $csvFilePath
    $users   = $importData.Users
    $allUsers = $importData.AllUsers

    foreach ($user in $users) {

        Write-Log "`n=== Configuration en cours : $($user.FirstName) $($user.LastName) ===" "Cyan"

        # Champs obligatoires
        if (-not $user.FirstName -or -not $user.LastName) {
            Write-Log "[!] Le CSV ne contient pas 'FirstName' OU 'LastName'. Saut de cette entrée." "Red"
            continue
        }

        # Site
        $selectedSiteName = $user.Site
        if (-not $selectedSiteName -or -not $sitesInfo.ContainsKey($selectedSiteName)) {
            Write-Log "Site '$selectedSiteName' non trouvé dans le mapping. Abandon." "Red"
            continue
        }
        $siteInfo  = $sitesInfo[$selectedSiteName]
        $Street    = $siteInfo.Street
        $PostalCode= $siteInfo.PostalCode
        $City      = $siteInfo.City
        $Company   = $siteInfo.Company
        $Office    = $selectedSiteName  # ou tout autre logique

        Write-Log "Site : $selectedSiteName" "Green"

        # OU
        $OUPath = if ($user.OrganizationalUnit) { 
            $user.OrganizationalUnit 
        } else { 
            Select-OU -Company $Company 
        }

        if (-not (Get-ADOrganizationalUnit -Filter "DistinguishedName -eq '$OUPath'" -ErrorAction SilentlyContinue)) {
            Write-Log "L’OU '$OUPath' n’existe pas. Abandon pour $($user.FirstName) $($user.LastName)." "Red"
            continue
        }

        Write-Log "OU sélectionnée : $OUPath" "Green"

        # Manager
        Write-Log "Sélection du manager..." "Blue"
        $ManagerDistinguishedName = $null
        $ManagerList = $allUsers | Select-Object Name, SamAccountName

        if (Get-Command Out-GridView -ErrorAction SilentlyContinue) {
            $managerSelection = $ManagerList | Out-GridView -Title "Sélectionnez le manager pour $($user.FirstName) $($user.LastName)" -OutputMode Single
        }
        else {
            Write-Log "Out-GridView indisponible. Sélection manuelle du manager." "Yellow"
            $i = 1
            foreach ($mgr in $ManagerList) {
                Write-Host "$i) $($mgr.Name) ($($mgr.SamAccountName))"
                $i++
            }
            $mgrChoice = Read-Host "Entrez le numéro du manager"
            if ($mgrChoice -match '^\d+$') {
                $index = [int]$mgrChoice - 1
                if ($index -ge 0 -and $index -lt $ManagerList.Count) {
                    $managerSelection = $ManagerList[$index]
                }
            }
        }

        if ($managerSelection) {
            $ManagerUserName = $managerSelection.SamAccountName
            $Manager         = Get-ADUser -Identity $ManagerUserName
            if ($Manager) {
                $ManagerDistinguishedName = $Manager.DistinguishedName
                Write-Log "Manager sélectionné : $($Manager.Name)" "Green"
            }
            else {
                Write-Log "Manager '$ManagerUserName' introuvable. Abandon." "Red"
                continue
            }
        }
        else {
            Write-Log "Aucun manager sélectionné. Abandon." "Red"
            continue
        }

        # Génération du SamAccountName
        $UserName = Get-SamAccountName -FirstName $user.FirstName -LastName $user.LastName
        Write-Log "SamAccountName généré : $UserName" "Yellow"

        # E-mail principal
        $emailDomains = @("mondomaine.fr", "autre.fr")
        $selectedEmailDomain = $emailDomains | Out-GridView -Title "Sélectionnez un domaine pour $($user.FirstName) $($user.LastName)" -OutputMode Single

        if (-not $selectedEmailDomain) {
            Write-Log "Aucun domaine sélectionné. Abandon." "Red"
            continue
        }
        $Email = "$UserName@$selectedEmailDomain"
        Write-Log "Adresse e-mail : $Email" "Green"

        # Alias
        $aliasDomains = $emailDomains | Where-Object { $_ -ne $selectedEmailDomain }
        $aliasList = $aliasDomains | ForEach-Object { "$UserName@$_" }

        Write-Log "Alias disponibles : $($aliasList -join ', ')" "Cyan"
        $selectedAliases = $aliasList | Out-GridView -Title "Sélectionnez les alias (multi-sélection possible)" -OutputMode Multiple

        if (-not $selectedAliases) {
            Write-Log "Aucun alias sélectionné." "Yellow"
            $selectedAliases = @()
        }

        # UPN
        $UPN = "$UserName@MYDOMAIN.COM"  # Adapter au domaine interne
        if ($AzureADLoaded) {
            $count = 1
            while (Get-AzureADUser -Filter "UserPrincipalName eq '$UPN'" -ErrorAction SilentlyContinue) {
                $UPN = "$UserName$count@MYDOMAIN.COM"
                $count++
            }
        }
        Write-Log "UPN défini : $UPN" "Yellow"

        # Mot de passe
        [System.Security.SecureString]$securePassword = $null
        do {
            Write-Host "Veuillez saisir le mot de passe pour $($user.FirstName) $($user.LastName) :" -ForegroundColor Yellow
            $securePassword = Read-Host -AsSecureString
            if (-not $securePassword) {
                Write-Log "Le mot de passe ne peut pas être vide." "Red"
            }
        } while (-not $securePassword)

        # JobTitle / Department / Téléphones
        $JobTitle = if ($user.JobTitle) { $user.JobTitle } else { "Titre non spécifié" }
        $Department = if ($user.Department) { 
            $user.Department 
        } else { 
            Select-Department -Departments $departments
        }
        if (-not $Department) {
            Write-Log "Aucun département défini. Abandon." "Red"
            continue
        }
        $Phone          = $user.Phone
        $HomePhoneNumber= $user.HomePhoneNumber
        $Office365Package = if ($user.Office365Package) { $user.Office365Package } else { "Non spécifié" }

        # Création de l’utilisateur AD
        Write-Log "Création de l’utilisateur dans AD..." "Blue"
        try {
            $newUser = New-ADUser `
                -Name "$($user.FirstName) $($user.LastName)" `
                -Surname $user.LastName `
                -GivenName $user.FirstName `
                -SamAccountName $UserName `
                -UserPrincipalName $UPN `
                -OfficePhone $Phone `
                -Title $JobTitle `
                -Description "$JobTitle - $Office365Package" `
                -Office $Office `
                -Company $Company `
                -Department $Department `
                -Manager $ManagerDistinguishedName `
                -MobilePhone $Phone `
                -HomePhone $HomePhoneNumber `
                -EmailAddress $Email `
                -StreetAddress $Street `
                -City $City `
                -PostalCode $PostalCode `
                -AccountPassword $securePassword `
                -Path $OUPath `
                -Enabled $true `
                -ChangePasswordAtLogon $true `
                -DisplayName "$($user.FirstName) $($user.LastName)" `
                -PassThru

            Write-Log "Utilisateur $($user.FirstName) $($user.LastName) créé avec succès !" "Green"

            # Ajout ProxyAddresses (adresse principale + alias)
            $proxyList = @("SMTP:$Email")  # principal en majuscule
            foreach ($alias in $selectedAliases) {
                $proxyList += "smtp:$alias"
            }

            Set-ADUser -Identity $UserName -Replace @{ ProxyAddresses = $proxyList }
            Write-Log "ProxyAddresses ajoutés : $($proxyList -join ', ')" "Green"

            # Attribution des groupes PC
            Write-Log "Attribution des groupes en fonction du type de PC..." "Blue"
            Select-PCAndAssignGroups -UserName $UserName

            # Copie des droits d’un utilisateur existant (optionnel)
            $allUserList = $allUsers | Select-Object Name, SamAccountName
            $sourceUser = $allUserList | Out-GridView -Title "Sélectionnez l’utilisateur source pour copier les droits (ou annulez)" -OutputMode Single
            if ($sourceUser) {
                Copy-UserRights -SourceUserSamAccountName $sourceUser.SamAccountName -TargetUserSamAccountName $UserName
            }
            else {
                Write-Log "Aucun utilisateur source sélectionné. Pas de copie de droits." "Yellow"
            }

            # Ajout au groupe Office 365 par défaut
            Add-UserToGroup -UserSamAccountName $UserName -GroupDN $Office365GroupDN

            # Assignation des licences O365
            if ($AzureADLoaded) {
                Write-Log "Assignation des licences Office 365..." "Blue"
                $mainPackages = @("OFFICE_365_E3", "M365_BUSINESS_BASIC") # Exemple de SKU
                $selectedSku = $mainPackages | Out-GridView -Title "Sélectionnez la licence O365 principale pour $UserName" -OutputMode Single
                if ($selectedSku) {
                    $licensesToAssign = @($selectedSku)

                    # Visio plan 2 ?
                    $visioChoice = @("Oui","Non") | Out-GridView -Title "Ajouter Visio Plan 2 pour $UserName ?" -OutputMode Single
                    if ($visioChoice -eq "Oui") {
                        $licensesToAssign += "VISIO_PLAN2"
                    }

                    Assign-Office365License -UserUPN $UPN -Office365Packages $licensesToAssign
                }
                else {
                    Write-Log "Aucune licence O365 principale sélectionnée." "Yellow"
                }
            }
            else {
                Write-Log "AzureAD non connecté, pas de licence attribuée." "Yellow"
            }

        }
        catch {
            Write-Log "Erreur lors de la création ou configuration de l’utilisateur : $($_.Exception.Message)" "Red"
        }
    }

    Write-Log "Fin de l’importation des utilisateurs par '$executingUser'." "Yellow"
}

# --- FIN ---

<#
Pour lancer l’import depuis un CSV :

PS C:\> Create-UserFromCSV -csvFilePath "C:\Chemin\Vers\MonFichier.csv"

Exemple de structure CSV :

FirstName,LastName,Site,Phone,HomePhoneNumber,JobTitle,Office365Package,Department,OrganizationalUnit
John,Doe,Site-A,0102030405,,Technicien,OFFICE_365_E3,Informatique,"OU=Direction,OU=Paris,DC=MYDOMAIN,DC=COM"
Jane,Smith,Site-B,0607080910,0499999999,Développeuse,OFFICE_365_E3,Informatique,
#>
<#
===============================================================================
 Script : Import & Modification d’utilisateurs AD
 Objectif :
   - Fournir une fonction `Import-UserData` pour charger depuis un CSV
   - Fournir une fonction `Modify-User` pour modifier un compte AD via GUI
   - Exemple de `Write-Log` pour tracer les opérations (intégrable dans un fichier)
   - Corriger les points à risque, les DN, etc.
===============================================================================
#>

# --- Préambule ---

# Pour charger le module Active Directory :
Import-Module ActiveDirectory

# Si vous voulez forcer TLS 1.2 (pour d’autres connexions éventuelles)
#[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# --- Fonction de log minimale ---
function Write-Log {
    param(
        [string]$Message,
        [string]$Color = "White"
    )
    # Simplement en console (vous pouvez rediriger dans un fichier si besoin)
    Write-Host $Message -ForegroundColor $Color
}

# --------------------------------------------
# Fonction pour importer les données des utilisateurs depuis un fichier CSV
# --------------------------------------------
function Import-UserData {
    param (
        [Parameter(Mandatory = $true)]
        [string]$csvFilePath # Chemin vers le fichier CSV contenant les données des utilisateurs
    )

    # Vérifier si le fichier CSV existe
    if (-Not (Test-Path $csvFilePath)) {
        Write-Log "Le fichier CSV '$csvFilePath' n'existe pas!" "Red"
        throw "Fichier CSV introuvable."
    }

    # Importer les utilisateurs depuis le CSV
    try {
        $users = Import-Csv -Path $csvFilePath
        Write-Log "Importation des utilisateurs depuis le CSV réussie." "Green"
    }
    catch {
        Write-Log "Erreur lors de l'importation du CSV : $($_.Exception.Message)" "Red"
        throw "Échec de l'importation du CSV."
    }

    # Récupérer la liste de tous les utilisateurs AD pour la sélection des managers
    try {
        $allUsers = Get-ADUser -Filter * -Properties SamAccountName, GivenName, Surname, DistinguishedName
        Write-Log "Récupération de la liste des utilisateurs AD réussie." "Green"
    }
    catch {
        Write-Log "Erreur lors de la récupération des utilisateurs AD : $($_.Exception.Message)" "Red"
        throw "Échec de la récupération des utilisateurs AD."
    }

    # Retourner les utilisateurs importés et tous les utilisateurs AD
    return @{ Users = $users; AllUsers = $allUsers }
}

# --------------------------------------------
# Début du script pour la modification d'utilisateur (Modify-User)
# --------------------------------------------

function Modify-User {
    # -------------------------------
    # Fonctions auxiliaires internes
    # -------------------------------

    # Récupère l'ensemble des utilisateurs AD avec un ensemble de propriétés
    function Get-AllUsers {
        return Get-ADUser -Filter * -Properties DisplayName, SamAccountName, UserPrincipalName, EmailAddress, `
            GivenName, Surname, Title, Department, Company, Office, `
            StreetAddress, City, PostalCode, MobilePhone, OfficePhone, Manager, Enabled |
        Select-Object `
            @{ Name = "Nom complet";      Expression = { $_.DisplayName } },
            @{ Name = "Nom d'utilisateur"; Expression = { $_.SamAccountName } },
            @{ Name = "UPN";              Expression = { $_.UserPrincipalName } },
            @{ Name = "Email";            Expression = { $_.EmailAddress } },
            @{ Name = "Prénom";           Expression = { $_.GivenName } },
            @{ Name = "Nom";              Expression = { $_.Surname } },
            @{ Name = "Titre";            Expression = { $_.Title } },
            @{ Name = "Département";      Expression = { $_.Department } },
            @{ Name = "Entreprise";       Expression = { $_.Company } },
            @{ Name = "Bureau";           Expression = { $_.Office } },
            @{ Name = "Adresse";          Expression = { $_.StreetAddress } },
            @{ Name = "Ville";            Expression = { $_.City } },
            @{ Name = "Code postal";      Expression = { $_.PostalCode } },
            @{ Name = "Téléphone mobile"; Expression = { $_.MobilePhone } },
            @{ Name = "Téléphone fixe";   Expression = { $_.OfficePhone } },
            @{ Name = "Manager";          Expression = {
                if ($_.Manager) {
                    try {
                        (Get-ADUser -Identity $_.Manager -Properties DisplayName).DisplayName
                    }
                    catch {
                        "Non défini"
                    }
                }
                else { "Non défini" }
            }},
            @{ Name = "Compte actif"; Expression = { if ($_.Enabled) { "Oui" } else { "Non" } } }
    }

    # Valide le format d'un numéro de téléphone (format français : 01 23 45 67 89)
    function Validate-PhoneNumber {
        param ([string]$PhoneNumber)
        $pattern = "^(0[1-9])(\s[0-9]{2}){4}$"
        return ($PhoneNumber -match $pattern)
    }

    # Génère un nom d'utilisateur à partir du prénom et du nom
    function Generate-Username {
        param (
            [string]$FirstName,
            [string]$LastName
        )
        $formattedFirstName = ($FirstName -split "[-\s]+" | ForEach-Object {
            $_.Substring(0,1).ToLower()
        }) -join ""
        
        $formattedLastName  = ($LastName -replace "[-\s]+", "").ToLower()

        # --- FIX : Remplacer caractères accentués ---
        $formattedFirstName = $formattedFirstName `
            -replace '[éèêë]', 'e' -replace '[àâä]', 'a' -replace '[îï]', 'i' `
            -replace '[ôö]', 'o'   -replace '[ûüù]', 'u' -replace 'ç', 'c'

        $formattedLastName  = $formattedLastName `
            -replace '[éèêë]', 'e' -replace '[àâä]', 'a' -replace '[îï]', 'i' `
            -replace '[ôö]', 'o'   -replace '[ûüù]', 'u' -replace 'ç', 'c'

        return "$formattedFirstName$formattedLastName"
    }

    # -------------------------------
    # Tables de correspondance
    # -------------------------------

    # 1) Mapping des domaines email par compagnie (exemple)
    $companyDomainsMap = @{
        " "   = @("exemple.fr")
    }

    # 2) Mapping de la Base OU par compagnie (exemple)
    $companyBaseOUMap = @{
        " " = "OU=01 ,OU=,DC=,DC="
        ""  = "OU=07 ,OU=,DC=,DC="
        ""  = "OU=02 ,OU=,DC=,DC="
    }

    # 3) Mapping des sites par compagnie
    $sitesInfo = @{
        " " = @{
            Office = " Office"
            Sites  = @(
                @{ Street = "Rue Ferrer"; PostalCode = "59450"; City = "Sin le Noble" }
            )
        },
        "" = @{
            Sites = @(
                @{ Street = ""; PostalCode = ""; City = ""; Office = "" },
            )
        },
        "" = @{
            Office = " Office"
            Sites  = @(
                @{ Street = ""; PostalCode = ""; City = "" }
            )
        }
    }

    # 4) Départements par compagnie
    $companyDepartments = @{
        " " = @("Accueil", "Achats"),
         "" = @("Administratif", "Accueil", "Production")
    }

    # -------------------------------
    # Corps principal de Modify-User
    # -------------------------------

    do {
        # Récupère tous les utilisateurs de l'AD
        $allUsers = Get-AllUsers
        if (-not $allUsers -or $allUsers.Count -eq 0) {
            Write-Host "Aucun utilisateur trouvé dans l'Active Directory." -ForegroundColor Red
            return
        }

        # Sélection de l'utilisateur à modifier via une interface graphique
        $selectedUser = $allUsers | Out-GridView -Title "Sélectionnez l'utilisateur à modifier" -OutputMode Single
        if (-not $selectedUser) {
            Write-Host "Aucun utilisateur sélectionné. Fin de la fonction." -ForegroundColor Yellow
            break
        }

        # Récupère l'utilisateur complet depuis AD
        $user = Get-ADUser -Identity $selectedUser."Nom d'utilisateur" -Properties *
        if (-not $user) {
            Write-Host "Erreur : Impossible de récupérer l'utilisateur dans AD." -ForegroundColor Red
            return
        }

        Write-Host "`nUtilisateur sélectionné : $($user.DisplayName)" -ForegroundColor Cyan
        Write-Host "Prêt à effectuer des modifications..." -ForegroundColor Green

        # Boucle interne pour modifier plusieurs champs avant de terminer
        do {
            $fields = @(
                "Nom et Prénom (avec génération du nom d'utilisateur)",
                "Adresse e-mail",
                "Numéro de téléphone mobile",
                "Numéro de téléphone fixe",
                "Département",
                "Site et OU",
                "Titre de poste",
                "Type de PC",
                "Manager",
                "Mot de passe",
                "Désactiver l'utilisateur",
                "Attribuer les mêmes droits qu'un utilisateur existant",
                "Modifier le package Office d'un utilisateur",
                "Terminer"
            )

            $selectedField = $fields | Out-GridView -Title "Que souhaitez-vous modifier pour $($user.DisplayName) ?" -OutputMode Single
            if (-not $selectedField) {
                Write-Host "Aucune modification sélectionnée. Fin du sous-menu." -ForegroundColor Yellow
                break
            }

            switch ($selectedField) {

                # 1. Modification Nom/Prénom + SamAccountName
                "Nom et Prénom (avec génération du nom d'utilisateur)" {
                    do {
                        $newFirstName = Read-Host "Entrez le nouveau prénom"
                        if (-not $newFirstName) {
                            Write-Host "Aucun prénom saisi. Modification annulée." -ForegroundColor Yellow
                            break
                        }
                        $newLastName = Read-Host "Entrez le nouveau nom"
                        if (-not $newLastName) {
                            Write-Host "Aucun nom saisi. Modification annulée." -ForegroundColor Yellow
                            break
                        }

                        $newUsername = Generate-Username -FirstName $newFirstName -LastName $newLastName
                        Write-Host "Nouveau nom d'utilisateur généré : $newUsername" -ForegroundColor Green

                        try {
                            $oldDistinguishedName = $user.DistinguishedName

                            # --- FIX : Adapter UPN à votre vrai domaine interne ---
                            $newUPN = "$newUsername@mondomaine.local"

                            Set-ADUser -Identity $user.SamAccountName `
                                       -GivenName $newFirstName `
                                       -Surname   $newLastName `
                                       -SamAccountName $newUsername `
                                       -UserPrincipalName $newUPN `
                                       -DisplayName "$newLastName $newFirstName"

                            Rename-ADObject -Identity $oldDistinguishedName -NewName "$newLastName $newFirstName"

                            Write-Host "Nom, prénom, UPN, DisplayName et CN mis à jour." -ForegroundColor Green
                            $user = Get-ADUser -Identity $newUsername -Properties *
                        }
                        catch {
                            Write-Host "Erreur lors de la mise à jour du nom et prénom : $_" -ForegroundColor Red
                        }
                        break
                    } while ($true)
                }

                # 2. Modification de l'adresse e-mail
                "Adresse e-mail" {
                    do {
                        $username = $user.SamAccountName
                        if (-not $username) {
                            Write-Host "Erreur : Impossible de récupérer le nom d'utilisateur." -ForegroundColor Red
                            break
                        }
                        Write-Host "Le préfixe de l'adresse e-mail sera : $username" -ForegroundColor Cyan

                        # --- FIX : Adaptez la liste des domaines ---
                        $domains = @("exemple.fr", "autre.fr", "encore-un.fr")

                        $selectedDomain = $domains | Out-GridView -Title "Sélectionnez un domaine" -OutputMode Single
                        if (-not $selectedDomain) {
                            Write-Host "Aucun domaine sélectionné. Modification annulée." -ForegroundColor Yellow
                            break
                        }

                        $newEmailAddress = "$username@$selectedDomain"
                        Write-Host "Nouvelle adresse e-mail : $newEmailAddress" -ForegroundColor Green

                        try {
                            # Mettre à jour EmailAddress et UPN
                            Set-ADUser -Identity $user.SamAccountName -EmailAddress $newEmailAddress -UserPrincipalName $newEmailAddress
                            Write-Host "Adresse e-mail mise à jour." -ForegroundColor Green

                            # Recharger l'utilisateur
                            $user = Get-ADUser -Identity $user.SamAccountName -Properties *
                        }
                        catch {
                            Write-Host "Erreur lors de la mise à jour de l'adresse e-mail : $_" -ForegroundColor Red
                        }
                        break
                    } while ($true)
                }

                # 3. Modification du numéro de téléphone mobile
                "Numéro de téléphone mobile" {
                    do {
                        $newMobile = Read-Host "Entrez le nouveau numéro de mobile (format : 01 23 45 67 89)"
                        if (-not $newMobile) {
                            Write-Host "Aucune saisie. Modification annulée." -ForegroundColor Yellow
                            break
                        }
                        if (Validate-PhoneNumber -PhoneNumber $newMobile) {
                            try {
                                Set-ADUser -Identity $user.SamAccountName -Replace @{ mobile = $newMobile }
                                Write-Host "Téléphone mobile mis à jour." -ForegroundColor Green
                                $user = Get-ADUser -Identity $user.SamAccountName -Properties *
                                break
                            }
                            catch {
                                Write-Host "Erreur lors de la mise à jour du mobile : $_" -ForegroundColor Red
                                break
                            }
                        }
                        else {
                            Write-Host "Format invalide. Veuillez réessayer." -ForegroundColor Red
                        }
                    } while ($true)
                }

                # 4. Modification du numéro de téléphone fixe
                "Numéro de téléphone fixe" {
                    do {
                        $newPhone = Read-Host "Entrez le nouveau numéro de téléphone fixe (format : 01 23 45 67 89)"
                        if (-not $newPhone) {
                            Write-Host "Aucune saisie. Modification annulée." -ForegroundColor Yellow
                            break
                        }
                        if (Validate-PhoneNumber -PhoneNumber $newPhone) {
                            try {
                                Set-ADUser -Identity $user.SamAccountName -Replace @{
                                    telephoneNumber = $newPhone
                                    homePhone       = $newPhone
                                }
                                Write-Host "Téléphone fixe mis à jour." -ForegroundColor Green
                                $user = Get-ADUser -Identity $user.SamAccountName -Properties *
                                break
                            }
                            catch {
                                Write-Host "Erreur lors de la mise à jour du téléphone fixe : $_" -ForegroundColor Red
                                break
                            }
                        }
                        else {
                            Write-Host "Format invalide. Veuillez réessayer." -ForegroundColor Red
                        }
                    } while ($true)
                }

                # 5. Modification du Département
                "Département" {
                    $companyForDept = $user.Company
                    if (-not $companyForDept -or $companyForDept -eq "") {
                        Write-Host "La compagnie n'est pas définie. Impossible de sélectionner un département prédéfini." -ForegroundColor Red
                        break
                    }
                    if (-not $companyDepartments.ContainsKey($companyForDept)) {
                        Write-Host "Aucun département prédéfini pour la compagnie '$companyForDept'." -ForegroundColor Yellow
                        break
                    }
                    $departments = $companyDepartments[$companyForDept]
                    if (-not $departments -or $departments.Count -eq 0) {
                        Write-Host "Aucun département prédéfini pour la compagnie '$companyForDept'." -ForegroundColor Yellow
                        break
                    }
                    $selectedDepartment = $departments | Out-GridView -Title "Sélectionnez le département pour '$companyForDept'" -OutputMode Single
                    if ($selectedDepartment) {
                        try {
                            Set-ADUser -Identity $user.SamAccountName -Department $selectedDepartment
                            Write-Host "Département mis à jour : $selectedDepartment" -ForegroundColor Green
                            $user = Get-ADUser -Identity $user.SamAccountName -Properties *
                        }
                        catch {
                            Write-Host "Erreur lors de la mise à jour du département : $_" -ForegroundColor Red
                        }
                    }
                    else {
                        Write-Host "Aucun département sélectionné. Modification annulée." -ForegroundColor Yellow
                    }
                }

                # 6. Modification du Site et déplacement dans l'OU
                "Site et OU" {
                    # Sélection de la compagnie
                    $companyList = $sitesInfo.Keys | Sort-Object
                    $selectedCompany = $companyList | Out-GridView -Title "Sélectionnez la compagnie" -OutputMode Single
                    if (-not $selectedCompany) {
                        Write-Host "Aucune compagnie sélectionnée. Abandon." -ForegroundColor Yellow
                        break
                    }
                    Write-Host "Compagnie sélectionnée : $selectedCompany" -ForegroundColor Cyan

                    # Sélection du site
                    $availableSites = $sitesInfo[$selectedCompany].Sites
                    if (-not $availableSites) {
                        Write-Host "Aucun site défini pour la compagnie '$selectedCompany'." -ForegroundColor Red
                        break
                    }
                    $sitesFormatted = $availableSites | ForEach-Object {
                        [PSCustomObject]@{
                            "Adresse"     = $_.Street
                            "Code Postal" = $_.PostalCode
                            "Ville"       = $_.City
                        }
                    }
                    if ($sitesFormatted.Count -gt 1) {
                        $selectedSite = $sitesFormatted | Out-GridView -Title "Sélectionnez un site pour '$selectedCompany'" -OutputMode Single
                    }
                    else {
                        $selectedSite = $sitesFormatted[0]
                    }
                    if (-not $selectedSite) {
                        Write-Host "Aucun site sélectionné. Abandon." -ForegroundColor Yellow
                        break
                    }

                    $siteDetails = @{
                        Street     = $selectedSite.Adresse
                        PostalCode = $selectedSite.'Code Postal'
                        City       = $selectedSite.Ville
                    }

                    if (-not $companyBaseOUMap.ContainsKey($selectedCompany)) {
                        Write-Host "Aucune base OU définie pour '$selectedCompany'." -ForegroundColor Red
                        break
                    }
                    $baseOU = $companyBaseOUMap[$selectedCompany]
                    Write-Host "Base OU pour '$selectedCompany' : $baseOU" -ForegroundColor Cyan

                    try {
                        $childOUs = Get-ADOrganizationalUnit -SearchBase $baseOU -Filter * | Select-Object -ExpandProperty DistinguishedName
                    }
                    catch {
                        Write-Host "Erreur lors de la récupération des sous-OU : $_" -ForegroundColor Red
                        break
                    }
                    if (-not $childOUs -or $childOUs.Count -eq 0) {
                        Write-Host "Aucune sous-OU trouvée sous '$baseOU'." -ForegroundColor Yellow
                        break
                    }
                    $childOUsForDisplay = $childOUs | ForEach-Object {
                        [PSCustomObject]@{
                            "Nom Court"          = ($_ -split ',')[0] -replace '^OU=', ''
                            "DistinguishedName"  = $_
                        }
                    } | Sort-Object "Nom Court"

                    $selectedOUObject = $childOUsForDisplay | Out-GridView -Title "Sélectionnez l'OU (sous-OU)" -OutputMode Single
                    if ($selectedOUObject) {
                        $selectedOU = $selectedOUObject.DistinguishedName
                    }
                    else {
                        Write-Host "Aucune OU sélectionnée. Abandon." -ForegroundColor Yellow
                        break
                    }

                    try {
                        # Mise à jour des attributs liés au site (adresse, ville, cp, company)
                        Set-ADUser -Identity $user.SamAccountName `
                                   -Office $siteDetails.City `
                                   -StreetAddress $siteDetails.Street `
                                   -PostalCode $siteDetails.PostalCode `
                                   -City $siteDetails.City `
                                   -Company $selectedCompany

                        # Déplacement de l'utilisateur dans la nouvelle OU
                        Move-ADObject -Identity $user.DistinguishedName -TargetPath $selectedOU

                        Write-Host "`nSite, adresse et OU mis à jour pour l'utilisateur :" -ForegroundColor Green
                        Write-Host " - Utilisateur : $($user.SamAccountName)"
                        Write-Host " - Site : $($siteDetails.City)" -ForegroundColor Cyan
                        Write-Host " - Adresse : $($siteDetails.Street), $($siteDetails.PostalCode) $($siteDetails.City)" -ForegroundColor Cyan
                        Write-Host " - Compagnie : $selectedCompany" -ForegroundColor Cyan
                        Write-Host " - Nouvelle OU : $selectedOU" -ForegroundColor Cyan

                        $user = Get-ADUser -Identity $user.SamAccountName -Properties *
                    }
                    catch {
                        Write-Host "Erreur lors de la mise à jour ou du déplacement dans l'OU : $_" -ForegroundColor Red
                    }
                }

                # 7. Modification du Titre de poste
                "Titre de poste" {
                    do {
                        $newTitle = Read-Host "Entrez le nouveau titre de poste"
                        if (-not $newTitle) {
                            Write-Host "Aucun titre saisi. Abandon." -ForegroundColor Yellow
                            break
                        }
                        try {
                            Set-ADUser -Identity $user.SamAccountName -Title $newTitle
                            Write-Host "Titre de poste mis à jour : $newTitle" -ForegroundColor Green
                            $user = Get-ADUser -Identity $user.SamAccountName -Properties *
                        }
                        catch {
                            Write-Host "Erreur lors de la mise à jour du titre : $_" -ForegroundColor Red
                        }
                        break
                    } while ($true)
                }

                # 8. Modification du Type de PC et gestion des groupes associés
                "Type de PC" {
                    # --- FIX : Corriger orthographe dans la liste des groupes ---
                    $pcGroupList = @(
                        "CN=Utilisateurs_PC_Fixe,OU=PC,DC=MYDOMAIN,DC=COM",
                        "CN=Utilisateurs_PC_Portable,OU=PC,DC=MYDOMAIN,DC=COM",
                        "CN=Utilisateurs_PC_Mobile,OU=PC,DC=MYDOMAIN,DC=COM",
                        "CN=Utilisateurs_CL,OU=PC,DC=MYDOMAIN,DC=COM"
                    )

                    # Retirer l’utilisateur de tout groupe PC existant
                    $currentGroups = (Get-ADUser -Identity $user.SamAccountName -Properties MemberOf).MemberOf
                    foreach ($grp in $pcGroupList) {
                        if ($currentGroups -contains $grp) {
                            try {
                                Remove-ADGroupMember -Identity $grp -Members $user.SamAccountName -Confirm:$false -ErrorAction Stop
                                Write-Host "Retrait du groupe PC existant : $grp" -ForegroundColor Yellow
                            }
                            catch {
                                Write-Host "Erreur lors du retrait du groupe $grp : $_" -ForegroundColor Red
                            }
                        }
                    }

                    # Sélection du nouveau type de PC
                    $pcTypes = @("PC Fixe", "PC Portable", "Client léger")
                    if (Get-Command Out-GridView -ErrorAction SilentlyContinue) {
                        $selectedPCType = $pcTypes | Out-GridView -Title "Sélectionnez le type de PC" -OutputMode Single
                    }
                    else {
                        Write-Host "Out-GridView n'est pas disponible. Sélection manuelle." -ForegroundColor Yellow
                        $i = 1
                        foreach ($type in $pcTypes) {
                            Write-Host "$i) $type"
                            $i++
                        }
                        $choice = Read-Host "Entrez le numéro correspondant"
                        if ($choice -match '^\d+$' -and $choice -ge 1 -and $choice -le $pcTypes.Count) {
                            $selectedPCType = $pcTypes[$choice - 1]
                        }
                    }
                    if (-not $selectedPCType) {
                        Write-Host "Aucun type de PC sélectionné. Abandon." -ForegroundColor Red
                        break
                    }
                    Write-Host "Type de PC sélectionné : $selectedPCType" -ForegroundColor Green

                    # En fonction du type choisi, on ajoute les groupes
                    $pcGroupsToAdd = @()
                    if ($selectedPCType -eq "PC Fixe") {
                        # Exemple
                        $mobileOptions = @("Avec téléphone mobile", "Sans téléphone mobile")
                        $mobileChoice = $mobileOptions | Out-GridView -Title "Téléphone mobile ?" -OutputMode Single
                        if (-not $mobileChoice) {
                            Write-Host "Aucune réponse pour le mobile. Abandon." -ForegroundColor Red
                            break
                        }
                        if ($mobileChoice -eq "Avec téléphone mobile") {
                            $group = "CN=Utilisateurs_PC_Fixe_Mobile,OU=PC,DC=MYDOMAIN,DC=COM"
                        }
                        else {
                            $group = "CN=Utilisateurs_PC_Fixe_SansMobile,OU=PC,DC=MYDOMAIN,DC=COM"
                        }
                        $pcGroupsToAdd = @($group)
                    }
                    elseif ($selectedPCType -eq "PC Portable") {
                        $group = "CN=Utilisateurs_PC_Portable,OU=PC,DC=MYDOMAIN,DC=COM"
                        $pcGroupsToAdd = @($group)

                        # Exemple : un groupe VPN
                        $vpnGroup = "CN=Utilisateurs_VPN,OU=VPN,DC=MYDOMAIN,DC=COM"
                        $pcGroupsToAdd += $vpnGroup

                        # Vous pouvez ajouter une sélection de type VPN ici
                    }
                    elseif ($selectedPCType -eq "Client léger") {
                        $group = "CN=Utilisateurs_CL,OU=PC,DC=MYDOMAIN,DC=COM"
                        $pcGroupsToAdd = @($group)
                    }

                    # Ajouter l'utilisateur aux groupes
                    foreach ($grp in $pcGroupsToAdd) {
                        try {
                            Add-ADGroupMember -Identity $grp -Members $user.SamAccountName -ErrorAction Stop
                            Write-Host "Utilisateur ajouté au groupe : $grp" -ForegroundColor Green
                        }
                        catch {
                            Write-Host "Erreur lors de l'ajout de l'utilisateur au groupe $grp : $_" -ForegroundColor Red
                        }
                    }
                }

                # 9. Modification du Manager de l'utilisateur
                "Manager" {
                    do {
                        $allUsersMgr = Get-ADUser -Filter * -Properties DisplayName, SamAccountName |
                            Select-Object DisplayName, SamAccountName | Sort-Object DisplayName
                        if (-not $allUsersMgr -or $allUsersMgr.Count -eq 0) {
                            Write-Host "Aucun utilisateur trouvé pour le manager." -ForegroundColor Red
                            break
                        }
                        $selectedManager = $allUsersMgr | Out-GridView -Title "Sélectionnez un manager pour $($user.DisplayName)" -OutputMode Single
                        if (-not $selectedManager) {
                            Write-Host "Aucun manager sélectionné. Abandon." -ForegroundColor Yellow
                            break
                        }
                        try {
                            $mgrDN = (Get-ADUser -Identity $selectedManager.SamAccountName).DistinguishedName
                            Set-ADUser -Identity $user.SamAccountName -Manager $mgrDN
                            Write-Host "Manager mis à jour : $($selectedManager.DisplayName)" -ForegroundColor Green
                            $user = Get-ADUser -Identity $user.SamAccountName -Properties *
                        }
                        catch {
                            Write-Host "Erreur lors de la mise à jour du manager : $_" -ForegroundColor Red
                        }
                        break
                    } while ($true)
                }

                # 10. Modification du mot de passe
                "Mot de passe" {
                    do {
                        $newPassword = Read-Host -AsSecureString "Entrez le nouveau mot de passe pour $($user.DisplayName)"
                        if (-not $newPassword) {
                            Write-Host "Aucun mot de passe saisi. Abandon." -ForegroundColor Yellow
                            break
                        }
                        try {
                            Set-ADAccountPassword -Identity $user.SamAccountName -NewPassword $newPassword -Reset -ErrorAction Stop

                            # --- FIX : Probable erreur de frappe : remplacer Unk-ADAccount par Unlock-ADAccount
                            Unlock-ADAccount -Identity $user.SamAccountName -ErrorAction Stop

                            Write-Host "Mot de passe mis à jour pour $($user.DisplayName)." -ForegroundColor Green
                        }
                        catch {
                            Write-Host "Erreur lors de la mise à jour du mot de passe : $_" -ForegroundColor Red
                        }
                        break
                    } while ($true)
                }

                # 11. Désactivation de l'utilisateur
                "Désactiver l'utilisateur" {
                    do {
                        $confirmation = Read-Host "Êtes-vous sûr de vouloir désactiver $($user.DisplayName) ? (Oui/Non)"
                        if ($confirmation -notin @("Oui","O","o")) {
                            Write-Host "Action annulée." -ForegroundColor Yellow
                            break
                        }
                        # Désactivation du compte AD
                        try {
                            Disable-ADAccount -Identity $user.SamAccountName -ErrorAction Stop
                            Write-Host "$($user.DisplayName) désactivé." -ForegroundColor Green
                        }
                        catch {
                            Write-Host "Erreur lors de la désactivation : $_" -ForegroundColor Red
                            break
                        }

                        # Déplacement dans l'OU disabled
                        $disabledOU = "OU=A_verifier,OU=_Disabled,DC=,DC="
                        try {
                            Move-ADObject -Identity $user.DistinguishedName -TargetPath $disabledOU -ErrorAction Stop
                            Write-Host "Utilisateur déplacé vers l'OU : $disabledOU" -ForegroundColor Green
                        }
                        catch {
                            Write-Host "Erreur lors du déplacement vers l'OU '$disabledOU' : $_" -ForegroundColor Red
                        }

                        # Liste des groupes Office 365 dont on souhaite retirer l'utilisateur
                        $officeGroupsDN = @(
                            "CN=Office 365,OU=CHEMIN,DC=MYDOMAIN,DC=COM",
                            "CN=Office 365 - E3,OU=...,DC=,DC=",
                            # etc.
                        )

                        # Retrait de chacun des groupes O365
                        foreach ($groupDN in $officeGroupsDN) {
                            try {
                                Remove-ADGroupMember -Identity $groupDN -Members $user.SamAccountName -Confirm:$false -ErrorAction Stop
                                Write-Host "Retiré du groupe : $groupDN" -ForegroundColor Green
                            }
                            catch {
                                Write-Host "Erreur lors du retrait du groupe $groupDN : $_" -ForegroundColor Red
                            }
                        }
                        break
                    } while ($true)
                }

                # 12. Attribution des mêmes droits qu'un utilisateur existant
                "Attribuer les mêmes droits qu'un utilisateur existant" {
                    do {
                        $allUsersModel = Get-ADUser -Filter * -Properties DisplayName, SamAccountName |
                            Select-Object DisplayName, SamAccountName | Sort-Object DisplayName
                        if (-not $allUsersModel -or $allUsersModel.Count -eq 0) {
                            Write-Host "Aucun utilisateur trouvé." -ForegroundColor Red
                            break
                        }
                        $sourceUser = $allUsersModel | Out-GridView -Title "Sélectionnez l'utilisateur source" -OutputMode Single
                        if (-not $sourceUser) {
                            Write-Host "Aucun utilisateur modèle sélectionné. Abandon." -ForegroundColor Yellow
                            break
                        }
                        try {
                            $sourceGroups = (Get-ADUser -Identity $sourceUser.SamAccountName -Properties MemberOf).MemberOf
                        }
                        catch {
                            Write-Host "Erreur lors de la récupération des groupes de l'utilisateur modèle : $_" -ForegroundColor Red
                            break
                        }
                        if (-not $sourceGroups -or $sourceGroups.Count -eq 0) {
                            Write-Host "L'utilisateur $($sourceUser.DisplayName) n'appartient à aucun groupe." -ForegroundColor Yellow
                            break
                        }
                        foreach ($groupDN in $sourceGroups) {
                            $excludedGroups = @(
                                "CN=Domain Admins,CN=Users,DC=,DC=",
                                "CN=Enterprise Admins,CN=Users,DC=,DC=",
                                "CN=Schema Admins,CN=Users,DC=,DC="
                            )
                            if ($excludedGroups -contains $groupDN) {
                                Write-Host "Exclusion du groupe : $groupDN" -ForegroundColor Yellow
                                continue
                            }
                            try {
                                Add-ADGroupMember -Identity $groupDN -Members $user.SamAccountName -ErrorAction Stop
                                Write-Host "Ajouté au groupe : $groupDN" -ForegroundColor Green
                            }
                            catch {
                                Write-Host "Erreur pour le groupe $groupDN : $($_.Exception.Message)" -ForegroundColor Red
                            }
                        }
                        Write-Host "Droits copiés depuis $($sourceUser.DisplayName)." -ForegroundColor Green
                        break
                    } while ($true)
                }

                # 13. Modification du package Office d'un utilisateur
                "Modifier le package Office d'un utilisateur" {
                    do {
                        $officePackages = @{
                            "Office 365 - Business Basic"      = "CN=Office365_Basic,OU=..."
                            "Office 365 - Business Premium"    = "CN=Office365_Premium,OU=..."
                            "Office 365 - Business Standard"   = "CN=Office365_Standard,OU=..."
                            "Office 365 - E3"                  = "CN=Office365_E3,OU=..."
                            "Office 365 - Exchange Online P1"  = "CN=Office365_ExchangeP1,OU=..."
                            "Office 365 - Visio Plan2"         = "CN=Office365_VisioPlan2,OU=..."
                        }
                        $visioGroupDN = "CN=Office365_VisioPlan2,OU=..."
                        $standardOfficeGroupDNs = $officePackages.Values

                        try {
                            $currentOfficeGroups = (Get-ADUser -Identity $user.SamAccountName -Properties MemberOf).MemberOf |
                                Where-Object { $standardOfficeGroupDNs -contains $_ }
                        }
                        catch {
                            Write-Host "Erreur lors de la récupération des groupes Office : $_" -ForegroundColor Red
                            break
                        }

                        $standardOfficeLicenses = $officePackages.Keys
                        $selectedStandardPackage = $standardOfficeLicenses | Out-GridView -Title "Sélectionnez le package Office" -OutputMode Single
                        if (-not $selectedStandardPackage) {
                            Write-Host "Aucun package sélectionné. Abandon." -ForegroundColor Yellow
                            break
                        }

                        # Retirer l’utilisateur de tous les groupes Office qu’il a déjà
                        foreach ($groupDN in $currentOfficeGroups) {
                            try {
                                Remove-ADGroupMember -Identity $groupDN -Members $user.SamAccountName -Confirm:$false -ErrorAction Stop
                                Write-Host "Retiré du groupe : $groupDN" -ForegroundColor Green
                            }
                            catch {
                                Write-Host "Erreur lors du retrait du groupe $groupDN : $_" -ForegroundColor Red
                            }
                        }

                        # Ajouter le nouveau package
                        $newStandardGroupDN = $officePackages[$selectedStandardPackage]
                        try {
                            Add-ADGroupMember -Identity $newStandardGroupDN -Members $user.SamAccountName -ErrorAction Stop
                            Write-Host "Ajouté au groupe Office : $newStandardGroupDN" -ForegroundColor Green
                        }
                        catch {
                            Write-Host "Erreur lors de l'ajout au groupe $newStandardGroupDN : $_" -ForegroundColor Red
                        }

                        # Visio ?
                        $addVisio = Read-Host "Ajouter 'Office 365 - Visio Plan2' ? (Oui/Non)"
                        if ($addVisio -match "^(Oui|O|o)$") {
                            try {
                                Add-ADGroupMember -Identity $visioGroupDN -Members $user.SamAccountName -ErrorAction Stop
                                Write-Host "Ajouté au groupe Visio." -ForegroundColor Green
                            }
                            catch {
                                Write-Host "Erreur lors de l'ajout au groupe Visio : $_" -ForegroundColor Red
                            }
                        }

                        Write-Host "Modification des packages Office terminée." -ForegroundColor Green
                        break
                    } while ($true)
                }

                # 14. Terminer
                "Terminer" {
                    Write-Host "Modification terminée pour l'utilisateur $($user.DisplayName)." -ForegroundColor Cyan
                    break
                }

                default {
                    Write-Host "Option non gérée." -ForegroundColor Yellow
                }

            } # Fin du switch

        } while ($selectedField -ne "Terminer")

    } while ($true)

    Write-Host "Fin de Modify-User." -ForegroundColor Cyan
}

<#
=================================================================================
 Exemple d'utilisation :

 1) Pour importer des données depuis un CSV :
    $data = Import-UserData -csvFilePath "C:\MonFichier.csv"
    # Vous obtiendrez un objet contenant .Users et .AllUsers

 2) Pour lancer la modification d'un utilisateur existant :
    Modify-User

 Note : La fonction Modify-User s'appuie sur des sélections visuelles (Out-GridView).
        Assurez-vous d'exécuter en environnement compatible (Windows PowerShell, 
        ou PowerShell avec interface graphique).
=================================================================================
#>
<#
===============================================================================
 Script : Menu Principal & Fonctions (Alias, Groupes, Modification, etc.)
 Objectif :
   - Fournir un menu principal pour gérer les utilisateurs AD
   - Fonctions d'ajout / suppression / modification d'alias
   - Gestion de l’affectation de groupes Office 365
   - Edition de fonctions via Show-CodeMap
   - Suivi des modifications (Show-ModificationTracker)

 Note :
   - Adapter les chemins & DNs à votre infra
   - Tester sur un environnement de recette avant prod
===============================================================================
#>

# ---------------------
# 1. Fonctions utilitaires communes (Confirm-Exit, Show-UserList)
# ---------------------

function Confirm-Exit {
    param()
    $choice = Read-Host "Souhaitez-vous revenir au menu principal ? (O/N)"
    if ($choice -match '^(O|o|Oui|oui|y|Yes|yes|Y)$') {
        return $true
    }
    else {
        return $false
    }
}

function Show-UserList {
    # Vérifier et importer le module Active Directory
    if (-not (Get-Module ActiveDirectory)) {
        try {
            Import-Module ActiveDirectory -ErrorAction Stop
        }
        catch {
            Write-Host "Impossible de charger le module ActiveDirectory." -ForegroundColor Red
            Write-Host "Assurez-vous d’avoir les RSAT/ADDS Tools installés." -ForegroundColor Red
            return $null
        }
    }

    # Récupérer la liste de tous les utilisateurs, avec quelques propriétés utiles
    $allUsers = Get-ADUser -Filter * -Properties DisplayName, UserPrincipalName, SamAccountName |
        Select-Object `
            @{ Name = 'Name'; Expression = { $_.DisplayName } },
            SamAccountName,
            UserPrincipalName

    if (-not $allUsers) {
        Write-Host "Aucun utilisateur trouvé dans l'Active Directory." -ForegroundColor Red
        return $null
    }

    Write-Host "Veuillez sélectionner un utilisateur dans la liste ci-dessous..." -ForegroundColor Cyan

    # Sélection via Out-GridView
    $selectedUser = $allUsers | Out-GridView -Title "Sélectionnez un utilisateur" -OutputMode Single
    # Retourne l’objet sélectionné
    return $selectedUser
}

# ---------------------
# 2. Fonctions pour Gérer les Alias (Ajout, Modification, Suppression)
# ---------------------

function Add-AliasToUser {
    param (
        [string]$UserUPN
    )

    # Charger le module Active Directory
    if (-not (Get-Module ActiveDirectory)) {
        try {
            Import-Module ActiveDirectory -ErrorAction Stop
        }
        catch {
            Write-Host "Impossible de charger le module ActiveDirectory." -ForegroundColor Red
            Write-Host "Assurez-vous d’avoir les RSAT/ADDS Tools installés." -ForegroundColor Red
            return
        }
    }

    # --- NOTE : Adapter la liste des domaines pour vos alias ---
    $domains = @(
        "-.fr"
    )

    # Récupérer l'utilisateur visé et ses ProxyAddresses
    $User = Get-ADUser -Filter { UserPrincipalName -eq $UserUPN } -Properties SamAccountName, UserPrincipalName, ProxyAddresses
    if ($User) {
        $ProxyAddresses = $User.ProxyAddresses
        $AliasesAdded   = 0  # Compteur d’alias ajoutés

        # Récupérer automatiquement la partie avant le @ dans le UPN
        $usernameBeforeAt = ($User.UserPrincipalName -split "@")[0]

        while ($true) {
            Write-Host "`n===== Alias existants pour $($User.Name) =====" -ForegroundColor Cyan
            # Afficher chaque alias, en mettant en évidence l'alias principal
            $ProxyAddresses | ForEach-Object {
                if ($_ -eq $User.UserPrincipalName) {
                    Write-Host "$_ (Alias principal)" -ForegroundColor Green
                }
                else {
                    Write-Host $_ -ForegroundColor Yellow
                }
            }
            Write-Host "==============================================`n"

            Write-Host "Le nom d'utilisateur sera automatiquement utilisé : $usernameBeforeAt" -ForegroundColor Cyan

            # Sélection d'un domaine
            $selectedDomain = $domains | Out-GridView -Title "Sélectionnez un domaine pour l'alias" -OutputMode Single
            if (-not $selectedDomain) {
                Write-Host "Aucun domaine sélectionné. Modification annulée." -ForegroundColor Yellow
                break
            }

            # Construire l'alias
            $AliasWithPrefix = "smtp:$usernameBeforeAt@$selectedDomain"

            # Vérifier si l'alias existe déjà
            if ($ProxyAddresses -contains $AliasWithPrefix) {
                Write-Host "Alias déjà existant : $AliasWithPrefix" -ForegroundColor Yellow
            }
            else {
                # Ajout de l'alias
                $ProxyAddresses += $AliasWithPrefix
                Write-Host "Alias ajouté : $AliasWithPrefix" -ForegroundColor Green
                $AliasesAdded++
            }

            $continueAddition = Read-Host "Souhaitez-vous ajouter un autre alias ? (O/N)"
            if ($continueAddition -notmatch '^(O|o|Oui|oui|y|Yes|yes|Y)$') {
                break
            }
        }

        # Mettre à jour AD si au moins un alias a été ajouté
        if ($AliasesAdded -gt 0) {
            $ProxyAddresses = $ProxyAddresses | ForEach-Object { $_.ToString() }
            Set-ADUser -Identity $User.DistinguishedName -Replace @{ ProxyAddresses = $ProxyAddresses }
            Write-Host "`n$AliasesAdded alias ajouté(s) avec succès pour l'utilisateur $UserUPN." -ForegroundColor Green
        }
        else {
            Write-Host "`nAucun alias n'a été ajouté." -ForegroundColor Cyan
        }
    }
    else {
        Write-Host "Utilisateur non trouvé : $UserUPN" -ForegroundColor Red
    }

    # Proposer de recommencer ou revenir au menu
    $exitConfirmation = Read-Host "Appuyez sur Entrée pour quitter OU tapez 'restart' pour recommencer"
    if ($exitConfirmation -eq 'restart') {
        Add-AliasToUser -UserUPN $UserUPN
    }
    else {
        Write-Host "Retour au menu principal." -ForegroundColor Cyan
    }
}

function Modify-AliasForUser {
    param (
        [string]$UserUPN
    )

    # Charger le module Active Directory
    if (-not (Get-Module ActiveDirectory)) {
        try {
            Import-Module ActiveDirectory -ErrorAction Stop
        }
        catch {
            Write-Host "Impossible de charger le module ActiveDirectory." -ForegroundColor Red
            Write-Host "Assurez-vous d’avoir les RSAT/ADDS Tools installés." -ForegroundColor Red
            return
        }
    }

    # Récupérer l'utilisateur et ses ProxyAddresses
    $User = Get-ADUser -Filter { UserPrincipalName -eq $UserUPN } -Properties ProxyAddresses, DisplayName, SamAccountName
    if ($User) {
        $ProxyAddresses = $User.ProxyAddresses
        if (-not $ProxyAddresses -or $ProxyAddresses.Count -eq 0) {
            Write-Host "Cet utilisateur n'a aucun alias existant." -ForegroundColor Yellow
            return
        }
        while ($true) {
            Write-Host "`n===== Alias existants pour $($User.Name) =====" -ForegroundColor Cyan
            $ProxyAddresses | ForEach-Object {
                if ($_ -eq $User.UserPrincipalName) {
                    Write-Host "$_ (Alias principal)" -ForegroundColor Green
                }
                else {
                    Write-Host $_ -ForegroundColor Yellow
                }
            }
            Write-Host "==============================================`n"

            $oldAlias = Read-Host -Prompt "Entrez l'alias à modifier (sans 'smtp:') OU tapez 'quit' pour quitter"
            if ($oldAlias -eq 'quit') {
                Write-Host "Quitter l'option de modification d'alias." -ForegroundColor Cyan
                break
            }
            $formattedOldAlias = "smtp:$oldAlias"

            if ($ProxyAddresses -notcontains $formattedOldAlias) {
                Write-Host "Cet alias n'existe pas. Veuillez réessayer." -ForegroundColor Red
                continue
            }

            $newAlias         = Read-Host -Prompt "Entrez le nouvel alias (sans 'smtp:')"
            $formattedNewAlias= "smtp:$newAlias"

            # Vérifier si le nouvel alias existe déjà
            if ($ProxyAddresses -contains $formattedNewAlias) {
                Write-Host "Le nouvel alias existe déjà pour cet utilisateur." -ForegroundColor Yellow
                continue
            }

            # Remplacer l'ancien alias par le nouveau
            $ProxyAddresses = $ProxyAddresses | ForEach-Object {
                if ($_ -ieq $formattedOldAlias) {
                    $formattedNewAlias
                }
                else {
                    $_
                }
            }

            Write-Host "Alias modifié avec succès : $formattedOldAlias → $formattedNewAlias" -ForegroundColor Green

            $continueModifications = Read-Host "Souhaitez-vous modifier un autre alias ? (O/N)"
            if ($continueModifications -notmatch '^(O|o|Oui|oui|y|Yes|yes|Y)$') {
                break
            }
        }
        # Mettre à jour AD
        $ProxyAddresses = $ProxyAddresses | ForEach-Object { $_.ToString() }
        Set-ADUser -Identity $User.DistinguishedName -Replace @{ ProxyAddresses = $ProxyAddresses }
        Write-Host "`nAlias(s) modifié(s) avec succès pour l'utilisateur $UserUPN." -ForegroundColor Green
    }
    else {
        Write-Host "Utilisateur non trouvé : $UserUPN" -ForegroundColor Red
    }

    # Confirmation avant de quitter
    $exitConfirmation = Confirm-Exit
    if ($exitConfirmation) {
        Write-Host "Retour au menu principal." -ForegroundColor Cyan
        return
    }
    else {
        Modify-AliasForUser -UserUPN $UserUPN
    }
}

function Remove-AliasFromUser {
    param (
        [string]$UserUPN
    )

    # Charger le module Active Directory
    if (-not (Get-Module ActiveDirectory)) {
        try {
            Import-Module ActiveDirectory -ErrorAction Stop
        }
        catch {
            Write-Host "Impossible de charger le module ActiveDirectory." -ForegroundColor Red
            Write-Host "Assurez-vous d’avoir les RSAT/ADDS Tools installés." -ForegroundColor Red
            return
        }
    }

    # Récupérer l'utilisateur et ses ProxyAddresses
    $User = Get-ADUser -Filter { UserPrincipalName -eq $UserUPN } -Properties ProxyAddresses, DisplayName, SamAccountName
    if ($User) {
        $ProxyAddresses = $User.ProxyAddresses
        if (-not $ProxyAddresses -or $ProxyAddresses.Count -eq 0) {
            Write-Host "Cet utilisateur n'a aucun alias existant." -ForegroundColor Yellow
            return
        }
        while ($true) {
            Write-Host "`n===== Alias existants pour $($User.Name) =====" -ForegroundColor Cyan
            $ProxyAddresses | ForEach-Object {
                if ($_ -eq $User.UserPrincipalName) {
                    Write-Host "$_ (Alias principal)" -ForegroundColor Green
                }
                else {
                    Write-Host $_ -ForegroundColor Yellow
                }
            }
            Write-Host "==============================================`n"

            $aliasToRemove = Read-Host -Prompt "Entrez l'alias à supprimer (sans 'smtp:') OU tapez 'quit' pour quitter"
            if ($aliasToRemove -eq 'quit') {
                Write-Host "Quitter l'option de suppression d'alias." -ForegroundColor Cyan
                break
            }
            $formattedAlias = "smtp:$aliasToRemove"

            if ($ProxyAddresses -notcontains $formattedAlias) {
                Write-Host "Cet alias n'existe pas. Veuillez réessayer." -ForegroundColor Red
                continue
            }

            # Supprimer l'alias
            $ProxyAddresses = $ProxyAddresses | Where-Object { $_ -ne $formattedAlias }
            Write-Host "Alias supprimé : $formattedAlias" -ForegroundColor Green

            $continueRemoval = Read-Host "Souhaitez-vous supprimer un autre alias ? (O/N)"
            if ($continueRemoval -notmatch '^(O|o|Oui|oui|y|Yes|yes|Y)$') {
                break
            }
        }

        # Mettre à jour AD
        $ProxyAddresses = $ProxyAddresses | ForEach-Object { $_.ToString() }
        Set-ADUser -Identity $User.DistinguishedName -Replace @{ ProxyAddresses = $ProxyAddresses }
        Write-Host "`nAlias(s) supprimé(s) avec succès pour l'utilisateur $UserUPN." -ForegroundColor Green
    }
    else {
        Write-Host "Utilisateur non trouvé : $UserUPN" -ForegroundColor Red
    }

    $exitConfirmation = Confirm-Exit
    if ($exitConfirmation) {
        Write-Host "Retour au menu principal." -ForegroundColor Cyan
        return
    }
    else {
        Remove-AliasFromUser -UserUPN $UserUPN
    }
}

# ---------------------
# 3. Show-CodeMap (pour modifier le code d’une fonction dans un script)
# ---------------------

function Show-CodeMap {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ScriptPath
    )

    Clear-Host
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host " Analyse & Modification du script "
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host ""

    # Vérifier si le script existe
    if (-not (Test-Path $ScriptPath)) {
        Write-Host "Le fichier $ScriptPath n'existe pas." -ForegroundColor Red
        return
    }

    # Lire toutes les lignes du script
    $allLines = Get-Content -Path $ScriptPath

    # Liste pour stocker les fonctions détectées
    $functionList = New-Object System.Collections.Generic.List[PSObject]
    $currentFuncName   = $null
    $currentStartLine  = $null

    Write-Host "[1/5] Détection des fonctions dans le script..." -ForegroundColor Green
    for ($i = 0; $i -lt $allLines.Count; $i++) {
        $line = $allLines[$i]
        # Regex : "function NomDeFonction {"
        if ($line -match "^\s*function\s+([\w-]+)\s*\{?") {
            # Clôturer la fonction précédente si besoin
            if ($currentFuncName) {
                $endLine = $i
                $functionList.Add(
                    [PSCustomObject]@{
                        Name      = $currentFuncName
                        StartLine = $currentStartLine
                        EndLine   = $endLine
                        LineCount = $endLine - $currentStartLine
                        Body      = $allLines[$currentStartLine..($endLine - 1)] -join "`r`n"
                    }
                )
            }
            # Nouvelle fonction
            $currentFuncName  = $matches[1]
            $currentStartLine = $i + 1
        }
    }

    # Clôturer la dernière fonction
    if ($currentFuncName) {
        $functionList.Add(
            [PSCustomObject]@{
                Name      = $currentFuncName
                StartLine = $currentStartLine
                EndLine   = $allLines.Count
                LineCount = $allLines.Count - $currentStartLine
                Body      = $allLines[$currentStartLine..($allLines.Count - 1)] -join "`r`n"
            }
        )
    }

    if ($functionList.Count -eq 0) {
        Write-Host "Aucune fonction détectée dans $ScriptPath." -ForegroundColor Yellow
        return
    }

    Write-Host "[2/5] Liste des fonctions trouvées :" -ForegroundColor Green
    Write-Host ""
    $functionList | Sort-Object StartLine | Format-Table Name, StartLine, EndLine, LineCount -AutoSize
    Write-Host ""
    Write-Host "[3/5] Aperçu et sélection de la fonction à modifier" -ForegroundColor Green
    Write-Host "(Une fenêtre Out-GridView va s'ouvrir : double-cliquez OU sélectionnez + OK)" -ForegroundColor Yellow
    Write-Host ""

    # Construction d’une vue condensée pour Out-GridView
    $viewData = $functionList | ForEach-Object {
        [PSCustomObject]@{
            Name     = $_.Name
            Lines    = "$($_.StartLine) - $($_.EndLine)"
            LineCount= $_.LineCount
            Preview  = ($_.Body -split "`r`n")[0..([Math]::Min(2, ($_.Body -split "`r`n").Count - 1))] -join "`r`n"
            ObjectRef= $_
        }
    }
    $selected = $viewData | Out-GridView -Title "Sélectionnez la fonction à modifier" -OutputMode Single
    if (-not $selected) {
        Write-Host "Aucune fonction sélectionnée. Fin." -ForegroundColor Cyan
        return
    }

    $funcObj = $selected.ObjectRef
    Clear-Host
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host " Fonction sélectionnée : $($funcObj.Name)"
    Write-Host " Lignes : $($funcObj.StartLine) - $($funcObj.EndLine) (Nb : $($funcObj.LineCount))"
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host ""

    Write-Host "[4/5] Ouverture de la fonction dans Notepad..." -ForegroundColor Green
    Write-Host "Modifiez le code, puis fermez Notepad pour sauvegarder vos changements." -ForegroundColor Yellow
    Write-Host ""
    $tempPath = [System.IO.Path]::GetTempFileName().Replace(".tmp", ".ps1")
    # Écrire le corps de la fonction dans le fichier temporaire
    Set-Content -Path $tempPath -Value $funcObj.Body -Encoding UTF8

    # Ouvrir Notepad
    Start-Process notepad.exe -ArgumentList $tempPath
    Write-Host "En attente de la fermeture de Notepad..." -ForegroundColor DarkCyan

    # Boucle d’attente
    while (Get-Process notepad -ErrorAction SilentlyContinue | Where-Object { $_.Path -eq (Get-Command notepad.exe).Source }) {
        Start-Sleep -Seconds 1
    }

    # Lire le code mis à jour
    if (Test-Path $tempPath) {
        $updatedBody = Get-Content $tempPath -Raw
    }
    else {
        Write-Host "Fichier temporaire introuvable. Annulation de la modification." -ForegroundColor Red
        return
    }

    Write-Host "[5/5] Mise à jour du script principal..." -ForegroundColor Green
    $newLines = New-Object System.Collections.Generic.List[string]
    $idx      = 0

    foreach ($func in $functionList) {
        if ($func.Name -eq $funcObj.Name) {
            # Ajouter toutes les lignes avant la fonction
            for ($k = $idx; $k -lt $func.StartLine; $k++) {
                $newLines.Add($allLines[$k])
            }
            # Ajouter le nouveau code
            $newLines.AddRange($updatedBody -split "`r`n")
            # Sauter les anciennes lignes
            $idx = $func.EndLine
        }
        else {
            # Conserver le code des autres fonctions
            for ($k = $func.StartLine; $k -lt $func.EndLine; $k++) {
                $newLines.Add($allLines[$k])
            }
            $idx = $func.EndLine
        }
    }
    # Ajouter les lignes restantes après la dernière fonction
    for ($m = $idx; $m -lt $allLines.Count; $m++) {
        $newLines.Add($allLines[$m])
    }

    # Sauvegarde du script original
    $backupPath = "$ScriptPath.bak_$(Get-Date -Format "yyyyMMdd_HHmmss")"
    Copy-Item -Path $ScriptPath -Destination $backupPath -ErrorAction SilentlyContinue

    Set-Content -Path $ScriptPath -Value $newLines -Encoding UTF8
    Write-Host "`nMise à jour terminée !" -ForegroundColor Green
    Write-Host "Un backup a été créé : $backupPath" -ForegroundColor Yellow
    Write-Host "`n----------------------------------------------`n"
}

# ---------------------
# 4. Manage-UserGroups (pour affecter un user à des groupes Office 365, etc.)
# ---------------------

function Manage-UserGroups {
    # Importer ActiveDirectory si non déjà chargé
    if (-not (Get-Module ActiveDirectory)) {
        try {
            Import-Module ActiveDirectory -ErrorAction Stop
        }
        catch {
            Write-Host "Impossible de charger le module ActiveDirectory." -ForegroundColor Red
            Write-Host "Assurez-vous d’avoir les RSAT/ADDS Tools installés." -ForegroundColor Red
            return
        }
    }

    Write-Host "`nChargement de la liste des utilisateurs..." -ForegroundColor Cyan
    $allUsers = Get-ADUser -Filter * -Properties DisplayName, SamAccountName, MemberOf | Sort-Object DisplayName

    $userMenu = $allUsers | Select-Object @{Name='Nom';Expression={$_.DisplayName}}, SamAccountName
    $userSelection = $userMenu | Out-GridView -Title "Sélectionnez un utilisateur" -OutputMode Single
    if (-not $userSelection) {
        Write-Error "Aucun utilisateur sélectionné. Le script va s'arrêter."
        return
    }

    $user = Get-ADUser -Identity $userSelection.SamAccountName -Properties DistinguishedName, MemberOf

    # --- FIX : Adaptez les DNs ci-dessous à votre structure ---
    $defaultGroupDN = "CN=Office 365,OU=CHEMINDC=MYDOMAIN,DC=COM"
    $mainGroups = @{
        "1" = "CN=Office 365,OU=CHEMINDC=MYDOMAIN,DC=COM"
        "2" = "CN=Office 365 - Business Premium,OU=Office 365,OU=s,DC=,DC="
        "3" = "CN=Office 365 - Business Standard,OU=Office 365,OU=s,DC=,DC="
        "4" = "CN=Office 365 - E3,OU=Office 365,OU=s,DC=,DC="
        "5" = "CN=Office 365 - Exchange Online Plan 1,OU=Office 365,OU=s,DC=,DC="
    }
    $optionalGroupDN = "CN=Office 365,OU=CHEMINDC=MYDOMAIN,DC=COM"

    # Fonctions internes
    function Get-GroupDN {
        param ([string]$GroupDN)
        $group = Get-ADGroup -Filter { DistinguishedName -eq $GroupDN } -ErrorAction SilentlyContinue
        if ($group) {
            return $group.DistinguishedName
        }
        else {
            Write-Error "Groupe non trouvé : $GroupDN"
            return $null
        }
    }

    function Add-UserToGroup {
        param (
            [string]$UserDN,
            [string]$GroupDN
        )
        try {
            Add-ADGroupMember -Identity $GroupDN -Members $UserDN -ErrorAction Stop
            Write-Host "Utilisateur ajouté au groupe : $GroupDN" -ForegroundColor Green
        }
        catch {
            Write-Error "Erreur lors de l'ajout au groupe $GroupDN : $_"
        }
    }

    function Remove-UserFromGroup {
        param (
            [string]$UserDN,
            [string]$GroupDN
        )
        try {
            Remove-ADGroupMember -Identity $GroupDN -Members $UserDN -Confirm:$false -ErrorAction Stop
            Write-Host "Utilisateur retiré du groupe : $GroupDN" -ForegroundColor Yellow
        }
        catch {
            Write-Host "L'utilisateur n'était pas membre du groupe : $GroupDN" -ForegroundColor DarkYellow
        }
    }

    # Ajouter automatiquement l'utilisateur au groupe par défaut
    $validDefaultDN = Get-GroupDN -GroupDN $defaultGroupDN
    if ($validDefaultDN) {
        Add-UserToGroup -UserDN $user.DistinguishedName -GroupDN $validDefaultDN
    }
    else {
        Write-Error "Le groupe par défaut n'a pas été trouvé. L'opération s'arrête."
        return
    }

    # Préparer le menu pour choisir le groupe principal
    $groupOptions = foreach ($key in $mainGroups.Keys) {
        $dn = $mainGroups[$key]
        $groupName = ($dn -replace '^CN=', '') -split ',' | Select-Object -First 1
        [PSCustomObject]@{
            Numero = $key
            Nom    = $groupName
            DN     = $dn
        }
    }

    Write-Host "`nSélectionnez le groupe principal Office :" -ForegroundColor Cyan
    $selectedGroup = $groupOptions | Out-GridView -Title "Choisissez un groupe principal" -OutputMode Single
    if ($selectedGroup) {
        # Retirer l’utilisateur des autres groupes Office (sauf le groupe par défaut)
        $mainGroupDNs = $mainGroups.Values | Where-Object { $_ -ne $defaultGroupDN }
        foreach ($grpDN in $mainGroupDNs) {
            if ($user.MemberOf -contains $grpDN) {
                Remove-UserFromGroup -UserDN $user.DistinguishedName -GroupDN $grpDN
            }
        }

        # Ajouter le nouveau groupe principal
        $validMainDN = Get-GroupDN -GroupDN $selectedGroup.DN
        if ($validMainDN) {
            Add-UserToGroup -UserDN $user.DistinguishedName -GroupDN $validMainDN
        }
        else {
            Write-Error "Le groupe sélectionné est invalide."
        }
    }
    else {
        Write-Error "Aucun groupe principal n'a été sélectionné."
        return
    }

    # Gérer l'optionnel "Visio Plan2"
    $addOptional = Read-Host "`nVoulez-vous ajouter 'Office 365 - Visio Plan2' ? (O/N)"
    if ($addOptional -match '^(O|o|Oui|oui|y|Y)$') {
        $validOptionalDN = Get-GroupDN -GroupDN $optionalGroupDN
        if ($validOptionalDN) {
            Add-UserToGroup -UserDN $user.DistinguishedName -GroupDN $validOptionalDN
        }
        else {
            Write-Error "Le groupe optionnel est invalide."
        }
    }
    else {
        # Retirer si déjà présent
        if ($user.MemberOf -contains $optionalGroupDN) {
            Remove-UserFromGroup -UserDN $user.DistinguishedName -GroupDN $optionalGroupDN
        }
        Write-Host "Aucun groupe optionnel ajouté (ou retiré s'il était présent)." -ForegroundColor Yellow
    }

    Write-Host "`nProcessus terminé." -ForegroundColor Cyan
}

# ---------------------
# 5. Show-ModificationTracker (simple interface WinForms pour logs)
# ---------------------

function Show-ModificationTracker {
    # Importer les assemblages nécessaires pour l'interface graphique
    Add-Type -AssemblyName PresentationFramework
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    # Définir le chemin du fichier CSV
    $csvFile = "modifications_log.csv"
    # Vérifier si le fichier CSV existe, sinon le créer avec les en-têtes
    if (-not (Test-Path $csvFile)) {
        "Date,Utilisateur,Commentaire" | Out-File -FilePath $csvFile -Encoding UTF8
    }

    # Fonctions internes pour créer les contrôles
    function Create-Label ($text, $font, $location) {
        $label = New-Object System.Windows.Forms.Label
        $label.Text = $text
        $label.Font = $font
        $label.AutoSize = $true
        $label.Location = $location
        return $label
    }

    function Create-Button ($text, $location, $clickAction) {
        $button = New-Object System.Windows.Forms.Button
        $button.Text = $text
        $button.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
        $button.Size = New-Object System.Drawing.Size(150, 40)
        $button.Location = $location
        $button.BackColor = [System.Drawing.Color]::LightGray
        $button.Add_Click($clickAction)
        return $button
    }

    # Charger et afficher les modifications depuis le CSV
    function Load-Modifications {
        $txtModifications.Clear()
        try {
            Import-Csv -Path $csvFile | ForEach-Object {
                # Date en gras, bleu foncé
                $txtModifications.SelectionFont = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
                $txtModifications.SelectionColor = [System.Drawing.Color]::DarkBlue
                $txtModifications.AppendText("Date : " + $_.Date + "`n")

                # Utilisateur & Commentaire en police régulière
                $txtModifications.SelectionFont = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Regular)
                $txtModifications.SelectionColor = [System.Drawing.Color]::Black
                $txtModifications.AppendText("Utilisateur : " + $_.Utilisateur + "`n")
                $txtModifications.AppendText("Commentaire : " + $_.Commentaire + "`n`n")
            } | Out-Null
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Erreur lors du chargement des modifications.", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }

    # Création du formulaire principal
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Suivi des modifications"
    $form.Size = New-Object System.Drawing.Size(800, 700)
    $form.StartPosition = "CenterScreen"

    # Label "Modifications récentes"
    $lblModifications = Create-Label "Modifications récentes :" (New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)) (New-Object System.Drawing.Point(10, 10))
    $form.Controls.Add($lblModifications)

    # RichTextBox pour afficher les modifications
    $txtModifications = New-Object System.Windows.Forms.RichTextBox
    $txtModifications.ReadOnly = $true
    $txtModifications.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $txtModifications.Size = New-Object System.Drawing.Size(760, 300)
    $txtModifications.Location = New-Object System.Drawing.Point(10, 40)
    $txtModifications.BackColor = [System.Drawing.Color]::White
    $form.Controls.Add($txtModifications)

    # Label pour "Ajouter un commentaire"
    $lblCommentaires = Create-Label "Ajouter un commentaire :" (New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)) (New-Object System.Drawing.Point(10, 360))
    $form.Controls.Add($lblCommentaires)

    # TextBox pour le commentaire
    $txtCommentaires = New-Object System.Windows.Forms.TextBox
    $txtCommentaires.Multiline = $true
    $txtCommentaires.ScrollBars = "Vertical"
    $txtCommentaires.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $txtCommentaires.Size = New-Object System.Drawing.Size(760, 150)
    $txtCommentaires.Location = New-Object System.Drawing.Point(10, 390)
    $txtCommentaires.BackColor = [System.Drawing.Color]::LightCyan
    $form.Controls.Add($txtCommentaires)

    # Bouton "Enregistrer"
    $btnSave = Create-Button "Enregistrer" (New-Object System.Drawing.Point(10, 560)) {
        if ($txtCommentaires.Text.Trim() -ne "") {
            try {
                $userName = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
                $newEntry = [PSCustomObject]@{
                    Date        = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    Utilisateur = $userName
                    Commentaire = $txtCommentaires.Text.Trim()
                }
                $newEntry | Export-Csv -Path $csvFile -Append -NoTypeInformation -Encoding UTF8

                Load-Modifications
                $txtCommentaires.Clear()
                [System.Windows.Forms.MessageBox]::Show("Commentaire enregistré avec succès.", "Information",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Information)
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Erreur lors de l'enregistrement.", "Erreur",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
        else {
            [System.Windows.Forms.MessageBox]::Show("Veuillez entrer un commentaire avant de sauvegarder.", "Avertissement",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
    }
    $form.Controls.Add($btnSave)

    # Bouton "Fermer"
    $btnClose = Create-Button "Fermer" (New-Object System.Drawing.Point(170, 560)) {
        $form.Close()
    }
    $form.Controls.Add($btnClose)

    # Charger les modifications existantes
    Load-Modifications

    # Afficher le formulaire
    [void]$form.ShowDialog()
}

# ---------------------
# 6. Menu principal
# ---------------------

function Show-Menu {
    Clear-Host
    Write-Host "===============================" -ForegroundColor Cyan
    Write-Host "       Menu Principal"
    Write-Host "===============================" -ForegroundColor Cyan
    Write-Host "1. Créer un utilisateur" -ForegroundColor Yellow
    Write-Host "2. Modifier un utilisateur (Modify-User)" -ForegroundColor Yellow
    Write-Host "3. Gérer les alias (Ajout/Suppression)" -ForegroundColor Yellow
    Write-Host "4. Gérer les groupes Office 365 (Manage-UserGroups)" -ForegroundColor Yellow
    Write-Host "5. Editer le script (Show-CodeMap)" -ForegroundColor Yellow
    Write-Host "6. Suivi des modifications (Show-ModificationTracker)" -ForegroundColor Yellow
    Write-Host "7. Quitter" -ForegroundColor Yellow
    Write-Host "===============================" -ForegroundColor Cyan
}

# --- ILLUSTRATION : placeholders, vous devrez avoir vos vraies fonctions Create-User, Create-UserFromCSV, Modify-User, etc. déjà définies ---
function Create-User {
    Write-Host "Fonction Create-User à implémenter ou déjà existante..." -ForegroundColor Green
}
function Create-UserFromCSV {
    param (
        [Parameter(Mandatory = $true)]
        [string]$csvFilePath
    )
    Write-Host "Fonction Create-UserFromCSV à implémenter ou déjà existante..." -ForegroundColor Green
}
function Modify-User {
    Write-Host "Fonction Modify-User à implémenter ou déjà existante..." -ForegroundColor Green
}

# ---------------------
# 7. Boucle du menu principal
# ---------------------

do {
    Show-Menu
    $choice = Read-Host -Prompt "Sélectionnez une option"

    switch ($choice) {
        '1' {
            # Créer un utilisateur
            Create-User
            Read-Host -Prompt "Appuyez sur Entrée pour revenir au menu"
        }

        '2' {
            # Modifier un utilisateur
            Modify-User
            Read-Host -Prompt "Appuyez sur Entrée pour revenir au menu"
        }

        '3' {
            # Gérer les alias (ajout, suppression)
            $user = Show-UserList
            if ($user) {
                $userUPN = $user.UserPrincipalName
                Write-Host "Utilisateur sélectionné : $($user.Name) - UPN : $userUPN"

                $aliasAction = Read-Host -Prompt "Sélectionnez une action : 1 - Ajouter un alias, 2 - Supprimer un alias"
                switch ($aliasAction) {
                    '1' {
                        Add-AliasToUser -UserUPN $userUPN
                    }
                    '2' {
                        Remove-AliasFromUser -UserUPN $userUPN
                    }
                    default {
                        Write-Host "Action invalide." -ForegroundColor Red
                    }
                }
            }
            else {
                Write-Host "Aucun utilisateur sélectionné." -ForegroundColor Yellow
            }
            Read-Host -Prompt "Appuyez sur Entrée pour revenir au menu"
        }

        '4' {
            # Gérer les groupes Office 365
            Manage-UserGroups
            Read-Host -Prompt "Appuyez sur Entrée pour revenir au menu"
        }

        '5' {
            # Éditer le script (via Show-CodeMap)
            $scriptPath = "C:\Scripts\Script Thomas GRZESINSKI\Script officiel à tester lundi\SCRIPT_OFFICIEL_AD_USER V3.ps1"
            if (Test-Path $scriptPath) {
                Show-CodeMap -ScriptPath $scriptPath
            }
            else {
                Write-Host "Le script $scriptPath est introuvable. Adaptez le chemin dans le code." -ForegroundColor Red
            }
            Read-Host -Prompt "Appuyez sur Entrée pour revenir au menu"
        }

        '6' {
            # Suivi des modifications
            Show-ModificationTracker
            Read-Host -Prompt "Appuyez sur Entrée pour revenir au menu"
        }

        '7' {
            Write-Host "Au revoir!" -ForegroundColor Green
            break
        }

        default {
            Write-Host "Option invalide, veuillez essayer à nouveau." -ForegroundColor Red
        }
    }

} while ($choice -ne '7')

Write-Host "Fin du script du menu." -ForegroundColor Cyan
