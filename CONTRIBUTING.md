# Contributing to Vyaprix

Thank you for taking the time to improve this project.

## Before proposing a change

- Open an issue or contact the maintainer to confirm the proposed work is appropriate.
- Check the project README for the stack, local setup, and available validation commands.
- Keep pull requests small and focused; explain the goal and the behavior changed.
- Include screenshots only when useful, and make sure they contain no personal information.

## Security and ownership

- Never commit passwords, API tokens, database URLs, private customer/attendee/patient data, internal documents, or production exports.
- Treat browser-exposed environment variables as public. Do not add server secrets to frontend code.
- Use synthetic data for testing and redact logs before sharing.
- Do not replace client branding, copy, verified metadata, or approved business information without authorization.
- See [LICENSE](LICENSE). A public repository does not automatically grant permission to reuse its code or assets.

## Validation

Run the commands available in `package.json` before submitting a change, normally the relevant build and lint checks. In the pull request, list the checks run and any that could not be completed.

Only maintainers or collaborators authorized by the applicable project owner should merge changes to deployment and production configuration.
