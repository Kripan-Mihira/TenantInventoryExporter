# Changelog

All notable changes to this project will be documented in this file.

The format is based on Keep a Changelog and this project follows Semantic Versioning where practical.

## [Unreleased]

### Added
- 

### Changed
- 

### Fixed
- 

### Security
- 

## [1.0.0] - YYYY-MM-DD

### Added
- Initial public release of the Tenant Inventory Reporter.
- Modular Microsoft 365 tenant-reporting PowerShell script.
- CSV report output, structured logging, run manifest, optional transcript, optional photo export, and optional Excel workbook.
- Pester tests for deterministic helper functions and local CSV export behavior.

### Changed
- Replaced AzureAD module usage with Microsoft Graph PowerShell.
- Replaced fixed local output paths with parameters.
- Replaced raw array concatenation and formatting-in-pipeline patterns with streamed CSV export.

### Security
- Removed hard-coded tenant-specific output paths.
- Removed broad unused Graph write scopes.
- Added guidance to keep exports, identities, and secrets out of source control.
