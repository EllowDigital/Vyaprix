# Security Policy — Vyaprix

## Reporting a security issue

Please do not publish exploitable vulnerabilities, production credentials, or private data in a public issue.

Report security concerns privately to the project maintainer at **ellowdigitalindia@gmail.com**. Include the affected component, impact, reproduction steps, and suggested mitigation. Remove credentials, personal information, patient/customer/attendee data, and unredacted logs from the report.

## Important precautions

- Never commit passwords, access tokens, private API keys, database connection strings, production exports, or private logs.
- Treat client-side bundles and `VITE_*` / `EXPO_PUBLIC_*` variables as public.
- Put privileged database, email, payment, deletion, and administrative operations behind a trusted server-side service.
- Use separate development credentials and synthetic test data.
- Revoke or rotate exposed credentials immediately. Deleting a value from the latest commit does not remove it from Git history.
- Review the applicable project agreement and third-party asset licenses before redistributing code or artifacts.

## Scope

This policy covers security concerns in this repository and its documented configuration. It does not authorize testing systems or services without permission.
