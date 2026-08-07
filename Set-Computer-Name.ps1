# Load GUI assembly
Add-Type -AssemblyName Microsoft.VisualBasic

# Prompt for the Computer Name
$NewName = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the new Computer Name", "Post-Install Setup", "")

if (-not [string]::IsNullOrWhiteSpace($NewName)) {
    try {
        # Set the name
        Rename-Computer -NewName $NewName -Force
        
        # Display a brief message before restarting
        [System.Windows.Forms.MessageBox]::Show("Computer name set to $NewName. System will restart in 5 seconds.", "Success")
        Start-Sleep -Seconds 5
        Restart-Computer -Force
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error setting name: $($_.Exception.Message)", "Error")
    }
} else {
    [System.Windows.Forms.MessageBox]::Show("No name entered. Skipping rename.", "Warning")
}
