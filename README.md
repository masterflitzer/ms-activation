# Activate Windows and Office

**Microsoft-Activation-Scripts** by *massgravel*: https://github.com/massgravel/Microsoft-Activation-Scripts.git

**KMS_VL_ALL** by *kkkgo*: https://github.com/kkkgo/KMS_VL_ALL.git

# Install Windows and Office
## Windows

1. Download the **Media Creation Tool** (Windows) or the ISO (other OS) here: https://microsoft.com/software-download
1. Run the tool to create an installation media or burn the ISO.

## Office
### Script

- Download XML config file from this repository (for Volume Licensing or 365) or generate online with the **Office Customization Tool**: https://config.office.com/deploymentsettings and save it as *config-office.xml* in the Downloads directory
- Run *office-install.ps1* script as administrator

If your executionpolicy doesn't allow the script, run this command in PowerShell as Administrator: 
```
powershell -executionpolicy bypass -file "$home/downloads/office-install.ps1"
```

### Manually

- Download XML config file from this repository (for Volume Licensing or 365) or generate online with the Office **Customization Tool**: https://config.office.com/deploymentsettings and save it as *config-office.xml* in the Downloads directory
- Download the **Office Deployment Tool** and save it in the Downloads directory: https://microsoft.com/download/confirmation.aspx?id=49117
- Run it and choose a directory to extract.
- Open PowerShell as Administrator and run these commands successively: 
    - change to the previously chosen directory, e.g. `cd "$home/downloads/ms-activation/"`
    - `"./setup.exe" /download "../config-office.xml"`
    - `"./setup.exe" /configure "../config-office.xml"`

Note: You will need the 365 config, if you want to login with a school/work account.

You can read more about the [Office Customization Tool](https://docs.microsoft.com/deployoffice/overview-of-the-office-customization-tool-for-click-to-run) or the [Office Deployment Tool](https://docs.microsoft.com/deployoffice/overview-office-deployment-tool).
