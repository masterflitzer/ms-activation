# Activate Microsoft Windows / Office

-   **Microsoft-Activation-Scripts** by _massgravel_: https://github.com/massgravel/Microsoft-Activation-Scripts.git (Windows with HWID and Office with KMS)
-   **KMS_VL_ALL** by _kkkgo_: https://github.com/kkkgo/KMS_VL_ALL.git (Windows and Office with KMS)

# Install Windows and Office

## Windows

1. Download the **Media Creation Tool** or the image (ISO) here: https://microsoft.com/software-download
1. Create your installation media

## Office

-   Download the XML config file from this repository (for Volume Licensing or 365) or generate online with the **Office Customization Tool**: https://config.office.com/deploymentsettings and save it as `odt.xml` (e.g. `~/Downloads/odt.xml`)

If you want to deploy the beta, change the channel in the 2nd line to: `Channel="BetaChannel"` (probably only works with 365 version)

### Script

Download and run the script `odt.ps1` (right-click and **Run with PowerShell** or run it manually in PowerShell with an optional parameter)

-   `& "~/Downloads/odt.ps1"`
-   `& "~/Downloads/odt.ps1" -config "~/Downloads/odt.xml"`
-   `& "~/Downloads/odt.ps1" -c "~/Downloads/odt.xml"`
-   `& "~/Downloads/odt.ps1" "~/Downloads/odt.xml"`

If your ExecutionPolicy does not allow the script, run these commands in **PowerShell as Administrator** and try again:

-   `Set-ExecutionPolicy RemoteSigned`
-   `Unblock-File "~/Downloads/odt.ps1"`

Alternatively run the cmd wrapper `odt.cmd`

### Manual

-   Download the **Office Deployment Tool** and save it in the Downloads directory: https://microsoft.com/download/confirmation.aspx?id=49117
-   Run it and extract in a directory of your choice (e.g. `~/Downloads/office-deployment-tool`)
-   Open PowerShell and run these commands successively:
    -   `& "~/Downloads/office-deployment-tool/setup.exe" /download "$home/Downloads/odt.xml"`
    -   `& "~/Downloads/office-deployment-tool/setup.exe" /configure "$home/Downloads/odt.xml"`

# Notes:

-   If you want to activate your installation with a Microsoft account (paid 365 account), choose the 365 config
-   If you want to activate your installation with KMS (using one of the activation methods found above), choose the volume licensing (vl) config

Here you can read more about the [Office Customization Tool](https://docs.microsoft.com/deployoffice/overview-of-the-office-customization-tool-for-click-to-run) or the [Office Deployment Tool](https://docs.microsoft.com/deployoffice/overview-office-deployment-tool).
