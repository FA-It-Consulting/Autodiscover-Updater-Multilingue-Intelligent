<#
.SYNOPSIS
    Script PowerShell avec interface graphique multilingue (Fr/En) pour mettre à jour les URLs Autodiscover.

.DESCRIPTION
    - Interface graphique Windows Forms.
    - Choix dynamique de la langue (Français ou Anglais).
    - Saisie et application des nouvelles URLs internes et externes.
    - Sauvegarde CSV des configurations avant toute modification.
    - Vérification intelligente : ne modifie que si les URLs sont différentes.

.AUTEUR
    FA-IT Consulting (Farid Amghar)
    Version : 2025-03-27
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName PresentationFramework

# Choix de langue
$choice = [System.Windows.Forms.MessageBox]::Show("Do you want to use English? / Voulez-vous utiliser le français ?", "Language / Langue", "YesNo", "Question")
if ($choice -eq "Yes") {
    $lang = @{
        title = "Autodiscover URL Configuration"
        label_intro = "Enter new Autodiscover URLs to apply to all Exchange servers:"
        internal_label = "Internal URL:"
        external_label = "External URL:"
        button_apply = "Save & Apply"
        msg_empty = "Please fill in both URLs."
        msg_saved = "Current configuration saved at:"
        log_updating = "Updating server:"
        log_skipped = "Already up to date. Skipped."
        log_success = "Updated successfully."
        log_error = "Error on server:"
        log_done = "Update completed."
        msg_error_title = "Error"
    }
} else {
    $lang = @{
        title = "Configuration des URLs Autodiscover"
        label_intro = "Entrez les nouvelles URLs Autodiscover à appliquer sur tous les serveurs Exchange :"
        internal_label = "URL Interne :"
        external_label = "URL Externe :"
        button_apply = "Sauvegarder et Appliquer"
        msg_empty = "Veuillez remplir les deux URLs."
        msg_saved = "Configuration actuelle sauvegardée dans :"
        log_updating = "Mise à jour du serveur :"
        log_skipped = "Déjà à jour. Sauté."
        log_success = "Mis à jour avec succès."
        log_error = "Erreur sur le serveur :"
        log_done = "Mise à jour terminée."
        msg_error_title = "Erreur"
    }
}

# Créer le formulaire
$form = New-Object System.Windows.Forms.Form
$form.Text = $lang.title
$form.Size = New-Object System.Drawing.Size(700, 540)
$form.StartPosition = "CenterScreen"

# Label d'instruction
$label = New-Object System.Windows.Forms.Label
$label.Text = $lang.label_intro
$label.AutoSize = $true
$label.Location = New-Object System.Drawing.Point(10, 10)
$form.Controls.Add($label)

# Champ InternalUrl
$internalLabel = New-Object System.Windows.Forms.Label
$internalLabel.Text = $lang.internal_label
$internalLabel.Location = New-Object System.Drawing.Point(10, 40)
$internalLabel.AutoSize = $true
$form.Controls.Add($internalLabel)

$internalTextBox = New-Object System.Windows.Forms.TextBox
$internalTextBox.Size = New-Object System.Drawing.Size(650, 20)
$internalTextBox.Location = New-Object System.Drawing.Point(10, 60)
$form.Controls.Add($internalTextBox)

# Champ ExternalUrl
$externalLabel = New-Object System.Windows.Forms.Label
$externalLabel.Text = $lang.external_label
$externalLabel.Location = New-Object System.Drawing.Point(10, 90)
$externalLabel.AutoSize = $true
$form.Controls.Add($externalLabel)

$externalTextBox = New-Object System.Windows.Forms.TextBox
$externalTextBox.Size = New-Object System.Drawing.Size(650, 20)
$externalTextBox.Location = New-Object System.Drawing.Point(10, 110)
$form.Controls.Add($externalTextBox)

# Bouton
$updateButton = New-Object System.Windows.Forms.Button
$updateButton.Text = $lang.button_apply
$updateButton.Size = New-Object System.Drawing.Size(200, 30)
$updateButton.Location = New-Object System.Drawing.Point(10, 150)
$form.Controls.Add($updateButton)

# Zone de log
$logBox = New-Object System.Windows.Forms.TextBox
$logBox.Multiline = $true
$logBox.ScrollBars = "Vertical"
$logBox.ReadOnly = $true
$logBox.Size = New-Object System.Drawing.Size(660, 270)
$logBox.Location = New-Object System.Drawing.Point(10, 200)
$form.Controls.Add($logBox)

# Fonction du bouton
$updateButton.Add_Click({
    $internalUrl = $internalTextBox.Text
    $externalUrl = $externalTextBox.Text

    if ([string]::IsNullOrWhiteSpace($internalUrl) -or [string]::IsNullOrWhiteSpace($externalUrl)) {
        [System.Windows.Forms.MessageBox]::Show($lang.msg_empty, $lang.msg_error_title, "OK", "Error")
        return
    }

    $dirs = Get-AutodiscoverVirtualDirectory
    $backupPath = "$env:USERPROFILE\Desktop\Autodiscover_Backup_{0:yyyyMMdd_HHmmss}.csv" -f (Get-Date)
    $dirs | Select-Object Name, Server, InternalUrl, ExternalUrl, Identity | Export-Csv -Path $backupPath -NoTypeInformation -Encoding UTF8

    $logBox.AppendText("💾 $($lang.msg_saved) $backupPath" + [Environment]::NewLine)

    foreach ($dir in $dirs) {
        $logBox.AppendText("➡ $($lang.log_updating) $($dir.Server)..." + [Environment]::NewLine)

        if ($dir.InternalUrl -eq $internalUrl -and $dir.ExternalUrl -eq $externalUrl) {
            $logBox.AppendText("🔄 $($lang.log_skipped)" + [Environment]::NewLine)
            continue
        }

        try {
            Set-AutodiscoverVirtualDirectory -Identity $dir.Identity `
                -InternalUrl $internalUrl `
                -ExternalUrl $externalUrl

            $logBox.AppendText("✅ $($lang.log_success)" + [Environment]::NewLine)
        } catch {
            $logBox.AppendText("❌ $($lang.log_error) $($dir.Server) : $_" + [Environment]::NewLine)
        }
    }

    $logBox.AppendText("✔️ $($lang.log_done)" + [Environment]::NewLine)
})

[void]$form.ShowDialog()
