var builder = WebApplication.CreateBuilder(args);

// Add configuration from appsettings.json
builder.Services.AddEndpointsApiExplorer();

var app = builder.Build();

// Dictionary to store token-to-connection-string mappings
// In production, this would come from a secure configuration store
var connectionStrings = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
{
    { "token123", "Provider=MSOLEDBSQL19;Server=localhost;Database=master;Trusted_Connection=yes;" },
    { "devtoken", "Provider=MSOLEDBSQL19;Server=devserver;Database=devdb;Trusted_Connection=yes;" },
    { "prodtoken", "Provider=MSOLEDBSQL19;Server=prodserver;Database=proddb;Trusted_Connection=yes;" }
};

app.MapGet("/api/connectionstring", (string token) =>
{
    if (string.IsNullOrWhiteSpace(token))
    {
        return Results.BadRequest(new { error = "Token is required" });
    }

    if (connectionStrings.TryGetValue(token, out var connectionString))
    {
        return Results.Ok(new { connectionString });
    }

    return Results.NotFound(new { error = "Invalid token" });
});

app.Run();
