// Pin the timezone for the whole suite.
//
// Tool output now embeds a zone label (`2026-09-11 14:00 BST`), so assertions
// would otherwise depend on the machine's clock settings — passing on a laptop in
// London and failing in CI, which runs UTC. Tests that care about zone handling
// override this explicitly.
process.env['MS365_MCP_TIMEZONE'] = 'Europe/London';
