Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName Microsoft.VisualBasic

# Saisie du nom du PC
$hostname = [Microsoft.VisualBasic.Interaction]::InputBox("Nom du PC", "Qui est connecté ?")

if (-not [string]::IsNullOrWhiteSpace($hostname)) {

    # Identifiants d'un compte admin domaine
    $username = "lnh\adm.apps"
    $password = "xxxxxxxxxxxxx"
    $securePassword = ConvertTo-SecureString $password -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential ($username, $securePassword)

    try {
        # Récupère le nom d'utilisateur connecté
        $userSession = Get-WmiObject -Class Win32_ComputerSystem -ComputerName $hostname -Credential $credential -ErrorAction Stop | Select-Object -ExpandProperty UserName

        if ($userSession) {
            $login = $userSession.Split('\')[-1]

            # Requête LDAP pour retrouver l'utilisateur
            $domain = "lnh.local"  # ← à adapter selon ton environnement AD
            $ldapPath = "LDAP://$domain"

            $directoryEntry = New-Object System.DirectoryServices.DirectoryEntry($ldapPath, $username, $password)
            $searcher = New-Object System.DirectoryServices.DirectorySearcher($directoryEntry)
            $searcher.Filter = "(&(objectClass=user)(sAMAccountName=$login))"
            $searcher.PropertiesToLoad.Add("displayName") | Out-Null
            $searcher.PropertiesToLoad.Add("mail") | Out-Null
            $searcher.PropertiesToLoad.Add("title") | Out-Null
            $searcher.PropertiesToLoad.Add("department") | Out-Null

            $result = $searcher.FindOne()

            if ($result -ne $null) {
                $displayName = $result.Properties["displayname"] | Select-Object -First 1
                $email = $result.Properties["mail"] | Select-Object -First 1
                $poste = $result.Properties["title"] | Select-Object -First 1
                $departement = $result.Properties["department"] | Select-Object -First 1

                $info = "Utilisateur connecté sur $hostname :`n`n"
                $info += "Nom de session : $userSession`n"
                $info += "Nom complet    : $displayName`n"
                $info += "Adresse mail   : $email`n"
                $info += "Poste          : $poste`n"
                $info += "Département    : $departement"

                [void][System.Windows.MessageBox]::Show($info, "Informations utilisateur", "OK", "Information")
            } else {
                [void][System.Windows.MessageBox]::Show("Utilisateur $login introuvable dans l'annuaire LDAP.", "Info AD", "OK", "Warning")
            }
        }
        else {
            [void][System.Windows.MessageBox]::Show("Aucun utilisateur connecté sur $hostname ou PC inaccessible.", "Résultat", "OK", "Warning")
        }

    } catch {
        [void][System.Windows.MessageBox]::Show("Erreur de connexion à $hostname :`n$_", "Erreur", "OK", "Error")
    }

} else {
    [void][System.Windows.MessageBox]::Show("Vous devez entrer un nom de PC valide.", "Erreur", "OK", "Error")
}
