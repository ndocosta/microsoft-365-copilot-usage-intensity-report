# Contributing

Contributions that improve reliability, documentation, accessibility, or report
usefulness are welcome.

## Workflow

1. Fork the repository and create a focused branch.
2. Keep tenant names, IDs, user data, certificates, logs, and generated CSVs out
   of commits.
3. Follow the existing PowerShell style and include comment-based help for new
   public scripts.
4. Test the smallest relevant path, including PowerShell parsing and
   PSScriptAnalyzer.
5. Open a pull request describing the behavior change and validation performed.

For report changes, use PBIP source files and avoid unrelated formatting churn.
Do not regenerate or overwrite manually maintained report definitions.

By contributing, you agree that your contribution is licensed under the Apache
License 2.0.
