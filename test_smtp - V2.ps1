Set-ExecutionPolicy Bypass -Scope Process -Force

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- Creation de la Fenetre ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "Testeur de Serveur SMTP autonome"
$form.Size = New-Object System.Drawing.Size(420, 580)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

# Fonction pour ajouter des labels facilement
function Add-Label ($text, $x, $y) {
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $text
    $label.Location = New-Object System.Drawing.Point($x, $y)
    $label.Size = New-Object System.Drawing.Size(120, 20)
    $form.Controls.Add($label)
}

# --- Champs du Formulaire ---
Add-Label "Serveur SMTP :" 20 20
$txtServeur = New-Object System.Windows.Forms.TextBox
$txtServeur.Location = New-Object System.Drawing.Point(150, 18)
$txtServeur.Size = New-Object System.Drawing.Size(230, 20)
$txtServeur.Text = "smtp.gmail.com"
$form.Controls.Add($txtServeur)

Add-Label "Port :" 20 60
$txtPort = New-Object System.Windows.Forms.TextBox
$txtPort.Location = New-Object System.Drawing.Point(150, 58)
$txtPort.Size = New-Object System.Drawing.Size(60, 20)
$txtPort.Text = "587"
$form.Controls.Add($txtPort)

Add-Label "Mail Expediteur :" 20 100
$txtExpediteur = New-Object System.Windows.Forms.TextBox
$txtExpediteur.Location = New-Object System.Drawing.Point(150, 98)
$txtExpediteur.Size = New-Object System.Drawing.Size(230, 20)
$form.Controls.Add($txtExpediteur)

Add-Label "Mail Destinataire :" 20 140
$txtDestinataire = New-Object System.Windows.Forms.TextBox
$txtDestinataire.Location = New-Object System.Drawing.Point(150, 138)
$txtDestinataire.Size = New-Object System.Drawing.Size(230, 20)
$form.Controls.Add($txtDestinataire)

Add-Label "Nom d'utilisateur :" 20 180
$txtUtilisateur = New-Object System.Windows.Forms.TextBox
$txtUtilisateur.Location = New-Object System.Drawing.Point(150, 178)
$txtUtilisateur.Size = New-Object System.Drawing.Size(230, 20)
$form.Controls.Add($txtUtilisateur)

Add-Label "Mot de passe :" 20 220
$txtPassword = New-Object System.Windows.Forms.TextBox
$txtPassword.Location = New-Object System.Drawing.Point(150, 218)
$txtPassword.Size = New-Object System.Drawing.Size(230, 20)
$txtPassword.PasswordChar = '*'
$form.Controls.Add($txtPassword)

# --- Case Anonyme ---
$chkAnonyme = New-Object System.Windows.Forms.CheckBox
$chkAnonyme.Text = "Autoriser l'envoi anonyme (sans authentification)"
$chkAnonyme.Location = New-Object System.Drawing.Point(20, 260)
$chkAnonyme.Size = New-Object System.Drawing.Size(360, 25)
$form.Controls.Add($chkAnonyme)

# Activation / Desactivation des champs login
$chkAnonyme.Add_CheckedChanged({
    if ($chkAnonyme.Checked) {
        $txtUtilisateur.Enabled = $false
        $txtPassword.Enabled = $false
        $txtUtilisateur.Text = ""
        $txtPassword.Text = ""
    }
    else {
        $txtUtilisateur.Enabled = $true
        $txtPassword.Enabled = $true
    }
})

# --- Zone de Statut ---
$lblStatut = New-Object System.Windows.Forms.Label
$lblStatut.Location = New-Object System.Drawing.Point(20, 380)
$lblStatut.Size = New-Object System.Drawing.Size(360, 130)
$lblStatut.ForeColor = [System.Drawing.Color]::Blue
$lblStatut.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
$form.Controls.Add($lblStatut)

# --- Action du Bouton Envoyer ---
$btnEnvoyer = New-Object System.Windows.Forms.Button
$btnEnvoyer.Text = "Envoyer le mail de test"
$btnEnvoyer.Location = New-Object System.Drawing.Point(20, 310)
$btnEnvoyer.Size = New-Object System.Drawing.Size(360, 45)
$btnEnvoyer.BackColor = [System.Drawing.Color]::LightBlue

$btnEnvoyer.Add_Click({

    # Validation rapide
    if (
        [string]::IsNullOrWhiteSpace($txtServeur.Text) -or
        [string]::IsNullOrWhiteSpace($txtPort.Text) -or
        [string]::IsNullOrWhiteSpace($txtExpediteur.Text) -or
        [string]::IsNullOrWhiteSpace($txtDestinataire.Text)
    ) {
        $lblStatut.ForeColor = [System.Drawing.Color]::Red
        $lblStatut.Text = "Veuillez remplir au minimum le serveur, le port, l'expediteur et le destinataire."
        return
    }

    $lblStatut.ForeColor = [System.Drawing.Color]::DarkOrange
    $lblStatut.Text = "Tentative d'envoi en cours..."
    $form.Refresh()

    # Date / Heure
    $dateJour = Get-Date -Format "dd/MM/yyyy"
    $heureJour = Get-Date -Format "HH:mm"

    $sujet = "Test smtp fait le $dateJour le $heureJour"

    # Mot de passe masque
    $mdpMasque = New-Object String ('*', $txtPassword.Text.Length)
    if ([string]::IsNullOrEmpty($mdpMasque)) {
        $mdpMasque = "[Aucun]"
    }

    # Mode auth
    if ($chkAnonyme.Checked) {
        $modeAuth = "Anonyme"
    }
    else {
        $modeAuth = "Authentifie"
    }

    # Corps du mail
    $corps = @"
Ceci est un mail de test de configuration SMTP autonome.

Voici les parametres renseignes :

- Serveur SMTP : $($txtServeur.Text)
- Port : $($txtPort.Text)
- Mode : $modeAuth
- User : $($txtUtilisateur.Text)
- Mdp : $mdpMasque
- Mail expediteur : $($txtExpediteur.Text)

"@

    try {

        # Force TLS 1.2 / 1.3
        [Net.ServicePointManager]::SecurityProtocol = `
            [Net.SecurityProtocolType]::Tls12 -bor `
            [Net.SecurityProtocolType]::Tls13

        # Parametres SMTP
        $splatting = @{
            SmtpServer = $txtServeur.Text
            Port       = [int]$txtPort.Text
            From       = $txtExpediteur.Text
            To         = $txtDestinataire.Text
            Subject    = $sujet
            Body       = $corps
        }

        # SSL automatique
        if ($txtPort.Text -eq "587" -or $txtPort.Text -eq "465") {
            $splatting.UseSsl = $true
        }

        # Ajout des credentials UNIQUEMENT si mode anonyme desactive
        if (-not $chkAnonyme.Checked) {

            if (
                [string]::IsNullOrWhiteSpace($txtUtilisateur.Text) -or
                [string]::IsNullOrWhiteSpace($txtPassword.Text)
            ) {
                throw "Veuillez renseigner un utilisateur et un mot de passe ou cocher le mode anonyme."
            }

            $securePassword = ConvertTo-SecureString $txtPassword.Text -AsPlainText -Force

            $credentials = New-Object System.Management.Automation.PSCredential(
                $txtUtilisateur.Text,
                $securePassword
            )

            $splatting.Credential = $credentials
        }

        # Envoi
        Send-MailMessage @splatting

        $lblStatut.ForeColor = [System.Drawing.Color]::Green
        $lblStatut.Text = @"
Succes !

Le mail a ete envoye avec succes.

Sujet :
$sujet

Mode utilise :
$modeAuth
"@

    }
    catch {
        $lblStatut.ForeColor = [System.Drawing.Color]::Red
        $lblStatut.Text = "Erreur d'envoi :`n`n$($_.Exception.Message)"
    }

})

$form.Controls.Add($btnEnvoyer)

# Affichage
$form.ShowDialog() | Out-Null