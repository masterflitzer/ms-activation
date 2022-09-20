# Install & Activate Microsoft Windows / Office

## Activation

-   [**Microsoft-Activation-Scripts** by _massgravel_](https://github.com/massgravel/Microsoft-Activation-Scripts.git): Windows with HWID and Office with KMS
-   [**KMS_VL_ALL** by _kkkgo_](https://github.com/kkkgo/KMS_VL_ALL.git): Windows and Office with KMS

## Install

### Windows

-   Download the **Media Creation Tool** or the image (ISO) here: <https://microsoft.com/software-download>
-   Create your installation media

### Office

-   Download the XML config file from this repository (for Volume Licensing or 365) or generate online with the **Office Customization Tool**: <https://config.office.com/deploymentsettings> and save it (e.g. as `$HOME/Downloads/odt.xml`)

If you want to deploy the beta, change the channel in the 2nd line to `Channel="BetaChannel"` (probably only works with the 365 version)

#### Script

Download and run the script `odt.ps1`

##### CLI (Oneliner)

Download and run directly

```powershell
Invoke-RestMethod -UseBasicParsing -Uri "https://raw.githubusercontent.com/masterflitzer/ms-activation/main/odt.ps1" | Invoke-Expression
```

##### CLI

Download and save the script

```powershell
Invoke-RestMethod -UseBasicParsing -Uri "https://raw.githubusercontent.com/masterflitzer/ms-activation/main/odt.ps1" -OutFile "~/Downloads/odt.ps1"
```

Run the script (you can provide the path to the config file)

-   `& "~/Downloads/odt.ps1"`
-   `& "~/Downloads/odt.ps1" -Config "~/Downloads/odt.xml"`

If your ExecutionPolicy does not allow the execution of the script, run these commands and try running it again:

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process
Unblock-File "~/Downloads/odt.ps1"
```

Alternatively you can just do this:

-   `powershell -ExecutionPolicy Bypass -NoLogo -NoProfile -File "$HOME/Downloads/odt.ps1"`
-   `powershell -ExecutionPolicy Bypass -NoLogo -NoProfile -File "$HOME/Downloads/odt.ps1" -Config "~/Downloads/odt.xml"`

#### Manual

-   Download the [**Office Deployment Tool**](https://microsoft.com/download/confirmation.aspx?id=49117) and save it in the Downloads directory
-   Run it and extract in a directory of your choice (e.g. `~/Downloads/office-deployment-tool`)
-   Open PowerShell and run these commands successively:
    -   `& "~/Downloads/office-deployment-tool/setup.exe" /download "$home/Downloads/odt.xml"`
    -   `& "~/Downloads/office-deployment-tool/setup.exe" /configure "$home/Downloads/odt.xml"`

## Notes

-   If you want to activate your installation with a Microsoft account (paid 365 account), choose the 365 config
-   If you want to activate your installation with KMS (using one of the activation methods found above), choose the volume licensing (vl) config

Here you can read more about the [Office Customization Tool](https://docs.microsoft.com/deployoffice/overview-of-the-office-customization-tool-for-click-to-run) or the [Office Deployment Tool](https://docs.microsoft.com/deployoffice/overview-office-deployment-tool).
