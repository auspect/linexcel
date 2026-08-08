# Security Policy

## Reporting a vulnerability

Please do not open a public issue for a suspected vulnerability. Report it
privately to `phcalvet@gmail.com` with a description, affected version, and
steps to reproduce. Do not attach confidential workbooks, credentials, or API
keys; use a minimized synthetic reproduction instead.

You can expect an acknowledgement within seven days. Confirmed issues will be
triaged, fixed, and disclosed with a release note when appropriate.

## Supported versions

Security fixes are applied to the latest released version of `linexcel`.

## Scope

`linexcel` analyses workbooks supplied by its caller. The optional AI
documentation feature transmits the documented dossier to the provider you
configure, and no provider is configured for you: nothing is sent until you name
an endpoint or supply a callable. See
[Data handling](https://auspect.github.io/linexcel/guide/data-handling/) for
what each call sends and where it goes before enabling it.