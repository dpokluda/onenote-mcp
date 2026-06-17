using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using ModelContextProtocol.Server;
using OneNoteMcp.ComClient;
using OneNoteMcp.Configuration;
using OneNoteMcp.McpServer;

namespace OneNoteMcp;

/// <summary>
/// Process entry point. Wires up configuration, logging, the OneNote COM client, and the
/// MCP server (exposed over stdio), then runs until the host shuts down.
/// </summary>
public static class Program
{
    public static async Task<int> Main(string[] args)
    {
        var builder = Host.CreateApplicationBuilder(args);

        // Bind appsettings.json (copied next to the DLL) plus environment-variable overrides.
        builder.Configuration
            .SetBasePath(AppContext.BaseDirectory)
            .AddJsonFile("appsettings.json", optional: true, reloadOnChange: false)
            .AddEnvironmentVariables();

        // MCP uses stdout for the JSON-RPC transport, so every log line must go to stderr.
        builder.Logging.ClearProviders();
        builder.Logging.AddConsole(options => options.LogToStandardErrorThreshold = LogLevel.Trace);

        // Bind the "OneNote" config section and register the COM client as a singleton.
        builder.Services
            .Configure<OneNoteOptions>(builder.Configuration.GetSection(OneNoteOptions.SectionName))
            .AddSingleton<OneNoteComClient>();

        // Register the MCP server over stdio and surface the OneNote tools.
        builder.Services
            .AddMcpServer()
            .WithStdioServerTransport()
            .WithTools<OneNoteMcpServer>();

        await builder.Build().RunAsync().ConfigureAwait(false);
        return 0;
    }
}
