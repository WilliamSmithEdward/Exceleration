# Security policy

## Reporting a vulnerability

Report a vulnerability privately, not in a public issue or pull request:
[open a private report](https://github.com/WilliamSmithEdward/Exceleration/security/advisories/new).
Only the maintainer sees it. Include the Exceleration version, the .NET
version and operating system, the method involved, and the smallest
workbook or code that shows it, with credentials and private data removed.

A confirmed vulnerability is fixed in a release on nuget.org, and the
advisory is published with it, crediting you unless you ask otherwise.

## Supported versions

Only the latest release on nuget.org receives security fixes. Older
releases are not maintained separately; update when a fix ships.

## Scope

Exceleration is a library. It opens no port, reaches no network and starts
no process. What it reads and writes:

- **Files read.** `new Workbook(filePath)` opens the file at the path the
  caller gives, reads it whole through ExcelDataReader (.xlsx, .xlsm, .xlsb
  and .xls) and closes it. It reads no other file.
- **Files written.** Only `new Workbook(filePath, true)` writes: it copies
  the workbook into the folder `Exceleration.dll` was loaded from, under the
  workbook's file name, replacing a file of that name there, and leaves the
  copy behind. Nothing else is written, and nothing is ever written back to
  a workbook.
- **Process state.** Each constructor registers
  `CodePagesEncodingProvider.Instance` with `Encoding.RegisterProvider`,
  which adds the legacy code pages to the whole process.

The workbook is untrusted input. Parsing is ExcelDataReader's; the XML in an
.xlsx is read by .NET's `XmlReader`, which refuses a DTD, so external
entities are not resolved. A workbook that makes the library read or write
any file other than the workbook and the copy described above, reach the
network, or run code counts as a vulnerability. A workbook that throws an
exception is not one on its own; report it as an issue. Memory use is
covered below.

### Reading workbooks you did not create

The whole workbook is held in memory, and a sheet takes room for every cell
from A1 to its last used cell, empty or not, so a file of a few kilobytes
can need gigabytes. Exceleration sets no limit. Check the size and source
of a workbook before reading it, and read untrusted workbooks in a process
whose memory you can cap.

### The copy option

Pass `true` for the copy only for workbooks whose file name you control:
the copy replaces any file of that name in the application's folder.

## How the code is checked

Three workflows check every pull request and every push to `main`, and
their gates decide whether a change can merge: **CI passed**,
**Security passed** and **Malware scan passed**. A gate passes only when
every job before it did, and any unexpected finding fails it, whatever its
severity. Security also runs weekly, so new queries, rules and advisories
reach code that has not changed, and Malware scan runs daily, so new
signatures and rules reach files that have not changed.

- **Code:** CodeQL with GitHub's security-extended queries, for C#
  (extracted from a Release build of net8.0, net9.0 and net10.0) and GitHub
  Actions, and Semgrep with the default, C#, security-audit, secrets and
  GitHub Actions rule sets. Semgrep scans the library, the workflows that
  build and publish it, and the scripts in `scripts/security` that judge the
  scans. A `nosemgrep` comment cannot hide a finding. Results go to the
  repository's code scanning.
- **Workflows:** zizmor audits the GitHub Actions workflows; a finding fails
  Security.
- **Dependencies:** `dotnet list package --vulnerable --include-transitive`
  checks the packages `packages.lock.json` resolves, ExcelDataReader and
  ExcelDataReader.DataSet, after the same locked restore CI builds with, and
  any known vulnerability fails Security.
- **Malware:** ClamAV, with signatures freshclam fetches and verifies on
  every run, and YARA-X, with the YARA Forge rules pinned to a release and
  its SHA-256, scan every file the commit holds and the .nupkg built from it
  with the locked restore, both as the archive and unpacked. On a release
  they scan the very .nupkg that is published. YARA-X runs YARA Forge's full
  rule set. A scan error fails the report as a match does.
- **OpenSSF Scorecard** rates the repository's security practices on every
  change to `main` and weekly, and the README badge shows the result.
  Its Code-Review and Contributors checks assume more than one
  maintainer, such as a second person approving every change, so a
  single-maintainer project cannot score full marks on them. Its Fuzzing
  check finds no C# fuzzer short of OSS-Fuzz or ClusterFuzzLite, so this
  repository has no fuzz workflow; the workbook parsing is ExcelDataReader's.

## Accepted findings

A finding is fixed, or accepted with a written reason in
[.github/security/accepted.toml](https://github.com/WilliamSmithEdward/Exceleration/blob/main/.github/security/accepted.toml)
for CodeQL and Semgrep, or
[.github/security/malware-accepted.toml](https://github.com/WilliamSmithEdward/Exceleration/blob/main/.github/security/malware-accepted.toml)
for ClamAV and YARA-X. A finding entry matches on the tool, the rule and
the file, and in accepted.toml also the text of the flagged line, so an
edited line needs another review; a notice entry, for a warning a tool
raises about its own scan, matches on the tool, the warning and text the
message contains. An entry that no longer matches fails the report. zizmor
keeps its exceptions in `.github/zizmor.yml` or inline beside the line they
excuse, each with its reason.

The current entries: there are none in accepted.toml or
malware-accepted.toml. zizmor's `self-repository` and `superfluous-actions`
rules are turned off in `.github/zizmor.yml`, each with its reason and when
it comes back.

## Pinning and updates

Everything the workflows run is pinned: actions to full commit SHAs,
runners to named OS releases, scanner images to digests, Python tools to
hash-locked lock files, the library's NuGet packages to
`packages.lock.json`, restored in locked mode, the YARA Forge rules to a
release and its SHA-256, and the YARA-X engine to a release and its
SHA-256. `global.json` sets the .NET SDK's floor at 10.0.400 and lets it
roll forward to a newer feature band, so CI builds with the newest .NET 10
SDK. ClamAV's signatures change too often to pin, so freshclam fetches and
verifies them on every run.

Dependabot proposes updates to the GitHub Actions, the Semgrep and ClamAV
images, the hash-locked files in `.github/requirements`, the NuGet packages
and the .NET SDK in `global.json` once a version is a week old, and at once
for a security advisory. The Update YARA rules workflow proposes new YARA
pins in `.github/security/yara.json` each week. A minor or patch update,
and the YARA pull request, merges itself once CI, Security and Malware scan
pass; a third-party major version waits for review.

## Releases

A pushed `vX.Y.Z` tag builds the .nupkg with the locked restore, checks
that the tag matches the package version (`PackageVersion` in the csproj),
and runs Security and Malware scan on the tagged commit and that package.
Nothing is published unless all of them pass. The package goes to
nuget.org through trusted publishing, so no long-lived API key exists to
leak. The GitHub release, titled with the tag, carries the .nupkg,
`Exceleration-<version>-security-report.md` and
`Exceleration-<version>-malware-report.md` beside the scan results they
were made from, and the provenance bundle, with the version's section of
`CHANGELOG.md` as its notes. Started by hand, the Publish workflow is always
a dry run and publishes nothing.

### Verifying a download

Releases after 1.1.1.4 carry a GitHub build provenance attestation for the
.nupkg CI built. Check the copy attached to the GitHub release:

```
gh attestation verify Exceleration.<version>.nupkg --owner WilliamSmithEdward
```

The output names the commit and workflow run that built the file. The
signed bundle is also attached to the release as
`Exceleration-<version>.sigstore.json`, so the check works without asking
GitHub for it: add `--bundle Exceleration-<version>.sigstore.json`.

nuget.org adds its own repository signature to every package it serves,
which changes the file, so the copy from nuget.org does not match the
attestation. Check that copy's signature with `dotnet nuget verify`.

## Repository settings

<!-- repo-standards:begin security-settings. Copied from WilliamSmithEdward/repo-standards, templates/security/settings-block.md. Change it there; the weekly rescan fails a copy that differs. -->
- `main` accepts changes only through a pull request that passes
  **CI passed**, **Security passed** and **Malware scan passed**. The
  ruleset has no bypass, for the owner either, and refuses force-pushes and
  deleting the branch.
- A `v*` release tag cannot be moved or deleted once pushed, except by a
  repository admin.
- A workflow that uses an action not pinned to a full commit SHA fails to
  run. Workflow tokens are read-only unless a job is granted more for
  itself.
- Secret scanning with push protection, Dependabot alerts and security
  updates, and private vulnerability reporting are on.
<!-- repo-standards:end -->
