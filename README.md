# setup-vba

GitHub Action to prepare a Windows runner for VBA automation.

It performs three setup tasks:
- Install Microsoft Office (optional, enabled by default)
- Configure VBA security (enable VBOM access and allow macro execution)
- Prepare selected Office apps (open/initialize and close them cleanly)

## Inputs

- `office-apps` (default: `Excel,Word,PowerPoint,Access`)
	- Comma-separated list of Office apps to prepare.
	- Supported values: `Excel`, `Word`, `PowerPoint`, `Access`.
- `install-office` (default: `true`)
	- Set to `false` if Office is already installed on the runner.
- `office-package` (default: `office365proplus`)
	- Chocolatey package used when `install-office=true`.

## Outputs

- `office-apps`
	- Normalized comma-separated app list that was prepared.

## Example usage

on:
```yaml
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
				uses: DecimalTurn/setup-vba@main
				with:
					office-apps: Excel
			- name: "Build VBA-Enabled Documents"
				id: build_vba
				uses: DecimalTurn/VBA-Build@b9754c8a6c4e7aac10dd68bcc990ef7e2c8237fc #v1.4.0
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