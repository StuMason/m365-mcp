# Contributing to M365 MCP

Thanks for your interest in contributing! This document covers how the project maintains itself and how you can help.

## Project Maintenance

This project is designed to be low-maintenance while staying secure and up-to-date.

### Automated Security & Dependencies

**Dependabot** runs daily to keep dependencies secure:

- **Patch/Minor updates** → Auto-merged after CI passes
- **Major updates** → PR created with review checklist, requires manual approval
- **GitHub Actions** → Weekly updates on Mondays

Configuration: [`.github/dependabot.yml`](.github/dependabot.yml)

### Branch Protection

The `main` branch is protected:

- All CI checks must pass (Node 20.x, 22.x, 24.x)
- Admin bypass enabled for maintainers
- No force pushes (except admins)

### CI Pipeline

Every PR runs:

1. **Security audit** - `npm audit`
2. **Format check** - Prettier
3. **Lint** - ESLint
4. **Build** - TypeScript compilation
5. **Test** - Jest with coverage

## How to Contribute

### Reporting Issues

- **Bugs**: Open an issue with reproduction steps
- **Feature requests**: Open an issue describing the use case

### Making Changes

1. Fork the repository
2. Create a feature branch: `git checkout -b feature/your-feature`
3. Make your changes
4. Run tests: `npm test`
5. Run lint: `npm run lint`
6. Commit with conventional commits: `feat:`, `fix:`, `chore:`, etc.
7. Open a PR against `main`

### Adding New Tools

When adding Microsoft Graph API capabilities:

1. Check the [Microsoft Graph API docs](https://learn.microsoft.com/en-us/graph/api/overview)
2. Add types in `src/types/`
3. Add the client method in `src/lib/`
4. Add the MCP tool definition
5. Add tests in `src/__tests__/`
6. Update tool count in README.md and CLAUDE.md, and document the tool in `docs/tools.md`
7. Add changelog entry

### Code Style

- TypeScript strict mode
- Prettier for formatting
- ESLint for linting
- Conventional commits

## Architecture Overview

```text
src/
├── index.ts              # Entry point
├── lib/                  # Core implementation
├── types/                # TypeScript types
└── __tests__/            # Jest tests
```

## Release Process

1. Run `npm version <patch|minor|major> --no-git-tag-version`. It bumps `package.json`,
   and `server.json` and `manifest.json` with it
2. Add changelog entry
3. Merge to main
4. GitHub Actions publishes to npm, attaches `m365-mcp.mcpb` to a GitHub release and
   publishes to the MCP Registry

## Questions?

- Open a [GitHub Issue](https://github.com/StuMason/m365-mcp/issues)
