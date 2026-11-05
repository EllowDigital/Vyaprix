# Security Policy — Vyaprix

**Vyaprix by EllowLabs** takes the security and privacy of business records, application data, and user information seriously.

We encourage responsible security research and appreciate contributions that help identify and address potential vulnerabilities.

## 🔐 Reporting a Security Vulnerability

If you discover a security vulnerability in Vyaprix, please report it privately rather than creating a public GitHub issue.

**Security Contact:** <ellowdigitalindia@gmail.com>

When submitting a report, please include:

- **Affected Component:** Identify the application module, feature, or dependency involved.
- **Vulnerability Description:** Clearly explain the security issue.
- **Potential Impact:** Describe the possible consequences and affected data or functionality.
- **Reproduction Steps:** Provide clear steps to reproduce the issue.
- **Evidence:** Include relevant screenshots, sanitized logs, or proof-of-concept examples.
- **Suggested Mitigation:** If possible, recommend a way to resolve the vulnerability.

Please allow the maintainer reasonable time to investigate and address reported issues before publicly disclosing technical details.

**Important:** Never include passwords, API keys, access tokens, private customer information, production database exports, or other sensitive information in vulnerability reports.

## 🛡️ Security Best Practices

Anyone developing, deploying, or maintaining Vyaprix should follow these precautions:

### 1. Credential & Secret Protection

- Never commit passwords, API keys, access tokens, private certificates, or database credentials.
- Store sensitive configuration values outside the source code.
- Use environment variables or an appropriate secrets-management solution.
- Never share production credentials through public repositories, screenshots, or logs.
- Immediately revoke and rotate credentials if they become exposed.

### 2. Database & Business Data Security

- Protect SQLite databases, database backups, and exported business records.
- Restrict access to application data directories.
- Avoid exposing customer information, invoices, payment records, or business details.
- Validate imported databases and external data before processing.
- Maintain regular backups and verify recovery procedures.

### 3. Application & API Security

- Validate and sanitize user inputs.
- Apply appropriate permissions to sensitive operations.
- Keep privileged database, email, payment, deletion, and administrative operations behind trusted server-side services.
- Avoid exposing internal implementation details or sensitive error messages.
- Keep application dependencies updated and review security advisories.

### 4. Development & Deployment Security

- Use separate development and production environments.
- Use synthetic data for testing and demonstrations.
- Review dependencies before adding them to the project.
- Do not commit private logs, production exports, or confidential configuration files.
- Treat client-side environment variables, including `VITE_*` and `EXPO_PUBLIC_*`, as publicly accessible.

### 5. Git & Repository Security

- Review staged files before every commit.
- Use `.gitignore` to exclude databases, credentials, environment files, temporary files, and private exports.
- Remember that deleting a secret from the latest commit does not remove it from Git history.
- Rotate exposed credentials and clean affected repository history when necessary.

## 🎯 Security Scope

This policy applies to:

- The Vyaprix application source code.
- Application configuration and documented setup procedures.
- Database handling and local data-management workflows.
- Dependencies and integrations maintained as part of the project.
- Other security-related components within this repository.

This policy does not grant permission to test, scan, access, or exploit systems, servers, accounts, or third-party services without explicit authorization.

## ⚖️ Responsible Disclosure

We support responsible security research.

Security researchers are expected to:

- Act in good faith.
- Avoid accessing, modifying, or deleting data belonging to others.
- Avoid service disruption or denial-of-service testing.
- Respect user privacy and confidentiality.
- Report vulnerabilities privately and provide sufficient information for investigation.

## 📋 Security Updates

Security-related fixes and important updates may be communicated through repository releases, commit history, or project announcements.

Users are encouraged to keep their application and dependencies updated and maintain secure backups of important business records.

---

**Project:** Vyaprix

**Developed & Maintained by:** EllowLabs

**Security Contact:** <ellowdigitalindia@gmail.com>

*Building reliable software with security, privacy, and responsibility in mind.*
