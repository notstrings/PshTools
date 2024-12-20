# 一般用
properties:
  configurationVersion: 0.2.0
  resources:
    - resource: Microsoft.WinGet.DSC/WinGetPackage
      id: installIPMsg
      directives:
        description: Install IP Messenger
      settings:
        id: FastCopy.IPMsg
        source: winget

    - resource: Microsoft.WinGet.DSC/WinGetPackage
      id: installPowerToys
      directives:
        description: Install PowerToys
        allowPrerelease: true
      settings:
        id: Microsoft.PowerToys
        source: winget

    - resource: Microsoft.PowerToys.Configure/PowerToysConfigure
      dependsOn:
        - installPowerToys
      directives:
        description: Configure PowerToys
      settings:
        Awake:
          Enabled: true
          KeepDisplayOn: true
          Mode: INDEFINITE


