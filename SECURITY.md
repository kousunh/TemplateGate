# Security Policy

## Supported versions

TemplateGate is pre-1.0. Fixes land on the latest released version only.

## Reporting a vulnerability

Please report privately through GitHub Security Advisories:
<https://github.com/kousunh/TemplateGate/security/advisories/new>

Do not open a public issue for a vulnerability. Expect an initial response
within about a week.

Useful to include: the version, the platform, a minimal synthetic document and
policy that reproduce the problem, and what you expected to happen. Please do
not attach real business documents.

## What counts as a vulnerability

TemplateGate is an acceptance gate, so the most serious class of bug is one
that lets a disallowed change pass. Reports worth sending privately include:

- A document change that a policy should reject but the checker reports as PASS.
- A crafted document or policy file that causes code execution, a path
  traversal, or a write outside the intended output path while being read.
- A way for the checked document to influence the policy that governs it.

Note that `semantic.command` runs a command the policy author configured, by
design; that is the operator's own trusted input rather than a vulnerability.
Semantic mode `off`, the default, makes no network calls.
