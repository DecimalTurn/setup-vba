[![GitHub release](https://img.shields.io/github/v/release/DecimalTurn/setup-vba?style=flat-square)](https://github.com/DecimalTurn/setup-vba/releases)
[![GitHub license](https://img.shields.io/github/license/DecimalTurn/setup-vba?style=flat-square)](LICENSE)
[![GitHub branch status](https://img.shields.io/github/check-runs/DecimalTurn/setup-vba/main)](https://github.com/DecimalTurn/setup-vba/actions/workflows/tests.yml)

# setup-vba

GitHub Action to prepare a Windows runner for VBA automation.

Originally part of [VBA-Build](https://github.com/DecimalTurn/VBA-Build), this action was extracted as a reusable setup step for Office and VBA automation so `VBA-Build` can stay focused on build and test workflows.

It performs three setup tasks:
- Install Microsoft Office (optional, enabled by default)
- Configure VBA security (enable VBOM access and allow macro execution)
- Prepare selected Office apps (open/initialize and close them cleanly)

## Inputs

- `office-apps` (default: `Excel,Word,PowerPoint,Access`)
  - Comma-separated list of Office apps to prepare. **Only these apps are installed** — all others are excluded from the Office installation for shorter installation times.
  - Supported values: `Access`, `Excel`, `Outlook`, `PowerPoint`, `Publisher`, `Word`
- `install-office` (default: `true`)
  - Set to `false` if Office is already installed on the runner.
- `office-package` (default: `office365proplus`)
  - Chocolatey package used when `install-office=true`.
- `office-language` (default: `""`)
  - Language/locale code passed to the Chocolatey package when `install-office=true`. For example `fr-fr` for French, `de-de` for German, `es-es` for Spanish. When unset, Office defaults to matching the OS language (which for the GitHub runners is en-us).

## Outputs

- `office-apps`
  - Normalized comma-separated app list that was prepared.

## Example usage


```yml
name: Build VBA

on:
  workflow_dispatch:
  push:
    branches:
      - main
    paths-ignore:
      - 'README.md'

  permissions:
    contents: read
    id-token: write
    attestations: write

  jobs:
    build:
      runs-on: windows-latest
      steps:
        - name: "Checkout"
          uses: actions/checkout@v5
        - name: "Setup VBA runtime"
          uses: DecimalTurn/setup-vba@1a34c76cc5d0e2ae3f1da2b008a4aa097da59f23 #v0.3.0
          with:
            office-apps: Excel,Word,PowerPoint,Access
        - name: "Build VBA-Enabled Documents"
          id: build_vba
          uses: DecimalTurn/VBA-Build@73d88e15e27d1e80aa402c60981175a491281d25 #v2.0.0
          with:
            source-dir: "./src"
          timeout-minutes: 20
        - name: "Upload Build Artifact"
          uses: actions/upload-artifact@v4
          id: "upload"
          with:
            name: "VBA-Enabled-Documents"
            path: "./src/out/*"
            if-no-files-found: warn
        - name: "Attestation"
          if: ${{ !github.event.repository.private }}
          uses: actions/attest-build-provenance@v2
          with:
            subject-name: "VBA-Enabled-Documents"
            subject-digest: sha256:${{ steps.upload.outputs.artifact-digest }}
```

---

> **💡 Offline / self-hosted runner?** If Office is already installed, set `install-office: false` to skip the Chocolatey install step.

> **🇫🇷 French locale?** Use `office-language: fr-fr` (or `fr-ca` for Canadian French) to install Office in French:
> ```yaml
> - name: "Setup VBA runtime"
>   uses: DecimalTurn/setup-vba@v0
>   with:
>     office-language: fr-fr
> ```


