# Install & Activate Microsoft Windows / Office

## Activation

- [**Microsoft-Activation-Scripts** by _massgravel_](https://github.com/massgravel/Microsoft-Activation-Scripts.git): Windows with HWID and Office with KMS
- [**KMS_VL_ALL** by _kkkgo_](https://github.com/kkkgo/KMS_VL_ALL.git): Windows and Office with KMS

## Install

### Windows

- Download the [Media Creation Tool](https://microsoft.com/software-download)
- Create your installation media

### Office

- Generate your XML config file with the [Office Customization Tool](https://config.office.com/deploymentsettings)
- Save it (e.g. as `$HOME/Downloads/office.xml`)

Alternatively download one of the preconfigured config files from this repository

- Depending on your preferred activation method, you'll need to choose a corresponding configuration
  - Activate with KMS (Volume Licensing): `config-vl.xml`
  - Activate with Microsoft 365 account: `config-365.xml`

```pwsh
# Volume Licensing (VL)
irm -o "$HOME/Downloads/office.xml" https://github.com/masterflitzer/ms-activation/raw/main/config-vl.xml

# Microsoft 365
irm -o "$HOME/Downloads/office.xml" https://github.com/masterflitzer/ms-activation/raw/main/config-365.xml
```

If you want to deploy the beta, change the channel in the 2nd line to `Channel="BetaChannel"` (will probably only work with the **365** version)

#### Script

- Open PowerShell, execute this oneliner and follow the instructions

```pwsh
irm https://github.com/masterflitzer/ms-activation/raw/main/office.ps1 | iex
```

Alternatively:

- Download the script

```pwsh
irm -o $HOME/Downloads/office.ps1 https://github.com/masterflitzer/ms-activation/raw/main/office.ps1
```

- Run the script

You can optionally specify the path to the config file or the office deployment tool, otherwise the script will ask you interactively

- `& $HOME/Downloads/office.ps1`
- `& $HOME/Downloads/office.ps1 -config $HOME/Downloads/office.xml -tool $HOME/Downloads/odt.exe`

If your ExecutionPolicy does not allow the execution of the script, run these commands and try again:

```pwsh
Set-ExecutionPolicy RemoteSigned -Scope Process
Unblock-File $HOME/Downloads/office.ps1
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

Here you can read more about the [Office Customization Tool](https://docs.microsoft.com/deployoffice/overview-of-the-office-customization-tool-for-click-to-run) or the [Office Deployment Tool](https://docs.microsoft.com/deployoffice/overview-office-deployment-tool).
