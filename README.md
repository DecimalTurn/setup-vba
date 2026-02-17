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

```yaml
name: Build VBA

on:
	workflow_dispatch:

jobs:
	build:
		runs-on: windows-latest
		steps:
			- uses: actions/checkout@v6

			- name: Setup VBA runtime
				uses: DecimalTurn/setup-vba@main
				with:
					office-apps: Excel,Word

			- name: Build VBA artifacts
				uses: DecimalTurn/VBA-Build@main
				with:
					source-dir: src
```

## Migration for VBA-Build

Run `setup-vba` before `VBA-Build` in workflows, then keep `VBA-Build` focused on build/test logic.