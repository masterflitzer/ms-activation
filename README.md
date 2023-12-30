# Install & Activate Microsoft Windows / Office

## Activation

- [**Microsoft-Activation-Scripts** by _massgravel_](https://github.com/massgravel/Microsoft-Activation-Scripts.git)
- [**KMS_VL_ALL** by _kkkgo_](https://github.com/kkkgo/KMS_VL_ALL.git)

## Install

### Windows

- Download the [Media Creation Tool](https://microsoft.com/software-download)
- Create your installation media

### Office

- Generate your XML config file with the [Office Customization Tool](https://config.office.com/deploymentsettings)
- Save it (e.g. as `$HOME/Downloads/office.xml`)

Alternatively download one of the preconfigured config files from this repository

- Depending on your preferred activation method, you'll need to choose a corresponding configuration
  - Activate with Microsoft 365 account: `office-microsoft-365.xml`
  - Activate with Volume License (VL): `office-volume-license.xml`

```pwsh
# Microsoft 365
irm -OutFile "$HOME/Downloads/office.xml" https://github.com/masterflitzer/ms-activation/raw/main/office-microsoft-365.xml

# Volume License (VL)
irm -OutFile "$HOME/Downloads/office.xml" https://github.com/masterflitzer/ms-activation/raw/main/office-volume-license.xml
```

#### Automatic

- Open PowerShell, execute this one liner and follow the instructions

```pwsh
irm https://github.com/masterflitzer/ms-activation/raw/main/office.ps1 | iex
```

#### Manual

- Download the [**Office Deployment Tool**](https://microsoft.com/download/confirmation.aspx?id=49117)
- Save it (e.g. as `$HOME/Downloads/odt.exe`)
- Run it and extract in a directory of your choice (e.g. `$HOME/Downloads/office-deployment-tool`)
- Open PowerShell and run these commands successively:

```pwsh
cd "$HOME/Downloads/office-deployment-tool"
./setup.exe /download "$HOME/Downloads/office.xml"
./setup.exe /configure "$HOME/Downloads/office.xml"
```

## Notes

Here you can read more about the [Office Customization Tool](https://docs.microsoft.com/deployoffice/overview-of-the-office-customization-tool-for-click-to-run) or the [Office Deployment Tool](https://docs.microsoft.com/deployoffice/overview-office-deployment-tool)
