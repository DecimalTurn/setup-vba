# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/), and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added

- **Selective Office install**: The `office-apps` input now controls which apps are actually installed. Only the listed apps are installed; all others are excluded via the [ODT `<ExcludeApp>`](https://learn.microsoft.com/en-us/microsoft-365-apps/deploy/office-deployment-tool-configuration-options#excludeapp-element) element to reduce CI installation time.

## [0.2.0] - 2026-07-09

### Added

- `office-language` input to install Office in a specific locale (e.g. `fr-fr`, `de-de`).

## [0.1.2] - 2026-04-09

### Added

- VBAPM Excel test build workflow ([#2](https://github.com/DecimalTurn/setup-vba/pull/2)).

### Changed

- Truncate Chocolatey logs to last 1000 lines on install failure.

### Dependencies

- Update dependency node to v24 ([#4](https://github.com/DecimalTurn/setup-vba/pull/4)).

## [0.1.1] - 2026-03-13

### Changed

- Remove vba-build subfolder (Marketplace release).

## [0.1.0] - 2026-03-13

### Added

- Initial public release.
