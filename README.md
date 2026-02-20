# M365 MCP

[![npm version](https://img.shields.io/npm/v/@masonator/m365-mcp.svg)](https://www.npmjs.com/package/@masonator/m365-mcp)
[![MIT License](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
[![Node.js](https://img.shields.io/badge/node-%3E%3D18-brightgreen.svg)](https://nodejs.org)
[![TypeScript](https://img.shields.io/badge/TypeScript-5.8-blue.svg)](https://www.typescriptlang.org/)
[![CI](https://github.com/StuMason/m365-mcp/actions/workflows/ci.yml/badge.svg)](https://github.com/StuMason/m365-mcp/actions/workflows/ci.yml)

MCP server for Microsoft 365 via the Microsoft Graph API.

## Installation

### Claude Code

```bash
claude mcp add m365-mcp -- npx -y @masonator/m365-mcp
```

### Claude Desktop

Add to your Claude Desktop config (`claude_desktop_config.json`):

```json
{
  "mcpServers": {
    "m365-mcp": {
      "command": "npx",
      "args": ["-y", "@masonator/m365-mcp"]
    }
  }
}
```

## Environment Variables

| Variable | Required | Description                         |
| -------- | -------- | ----------------------------------- |
| _TBD_    |          | _Configuration details to be added_ |

## Tools

_Tool listing to be added once implementation is complete._

## Development

```bash
git clone https://github.com/StuMason/m365-mcp.git
cd m365-mcp
npm install
npm run build
npm test
```

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for details.

## License

[MIT](LICENSE) - Stu Mason
