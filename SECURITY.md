# Security Policy

## Reporting a vulnerability

Please report security vulnerabilities in xl-marinade **privately** — not via public
issues or pull requests.

- **Preferred:** use GitHub's private vulnerability reporting. Open the
  [Security tab](https://github.com/gaspatchio/xl-marinade/security) of this repository
  and click **"Report a vulnerability"**. This opens an advisory visible only to the
  maintainers.
- **Alternative:** email **security@opioinc.com**.

Please include the affected version, a description of the issue and its impact, and,
where possible, a minimal reproduction.

## Reporting bugs privately (not security-related)

Many xl-marinade users work at insurers and cannot post code publicly. You are welcome
to report **any** bug — not just vulnerabilities — by email to **security@opioinc.com**
instead of opening a public issue. When we confirm one:

- we open a public tracking issue with a **rewritten, neutral reproduction** — never
  your workbook, formulas, cell values, or expected outputs;
- your name, employer, and any workbook details stay out of the issue, the commit
  history, and the changelog (attribution is "reported via a private field report");
- we send you the issue link so you can follow the fix, and credit you publicly
  only if you ask us to.

## Our commitments

- We acknowledge new reports within **2 business days**.
- We aim to give an initial assessment (confirm or dismiss) within **7 business days**.
- We keep you informed as we work on a fix and credit you in the published advisory
  unless you ask us not to.

## Supported versions

xl-marinade is pre-1.0; security fixes are released against the **latest** version
published on PyPI. We recommend always tracking the latest release.

## Software Bill of Materials (SBOM)

Every release ships a CycloneDX SBOM (`sbom.cdx.json`, covering the Python dependency
graph) as a GitHub Release asset. Fetch it with:

```bash
gh release download <tag> --repo gaspatchio/xl-marinade --pattern 'sbom.cdx.json'
```

Wheels published to PyPI carry [PEP 740](https://peps.python.org/pep-0740/) build-provenance
attestations, generated automatically via PyPI Trusted Publishing.

## Disclosure

Confirmed vulnerabilities are published as GitHub Security Advisories on this
repository once a fix is available.
