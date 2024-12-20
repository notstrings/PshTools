properties:
  resources:
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
  configurationVersion: 0.2.0
