# microsoft windows & office activation

## activation

- [**microsoft-activation-scripts** by _massgravel_](https://github.com/massgravel/microsoft-activation-scripts.git)

## install

### windows

- [media creation tool](https://microsoft.com/software-download)

### office

#### automated

execute this one liner and follow the instructions to deploy office with the preconfigured office config file from this repository:

```pwsh
irm "https://github.com/masterflitzer/ms-activation/raw/main/office.ps1" | iex
```

#### manually

get office deployment tool:

- download the [office deployment tool](https://microsoft.com/download/details.aspx?id=49117) (e.g. as `$HOME/Downloads/odt.exe`)
- execute it and choose a directory of your choice for extraction (e.g. `$HOME/Downloads/office-deployment-tool`)

generate a custom office config file:

- go to the [office customization tool](https://config.office.com)
- create a new config / import existing config
- customize deployment settings
- export config (e.g. as `$HOME/Downloads/office-deployment-tool/office.xml`)

deploy office:

- open powershell (as admin) and execute these commands (adapt paths if needed):

```pwsh
cd "$HOME/Downloads/office-deployment-tool"
./setup.exe /download office.xml
./setup.exe /configure office.xml
```

## notes

read more about:

- [office customization tool](https://learn.microsoft.com/microsoft-365-apps/admin-center/overview-office-customization-tool)
- [office deployment tool](https://learn.microsoft.com/microsoft-365-apps/deploy/overview-office-deployment-tool)
