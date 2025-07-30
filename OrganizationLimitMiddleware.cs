using CvParser.API.Services.Interfaces;

namespace CvParser.API.Middleware;

public class OrganizationLimitMiddleware
{
    private readonly RequestDelegate _next;

    public OrganizationLimitMiddleware(RequestDelegate next)
    {
        _next = next;
    }

    public async Task InvokeAsync(HttpContext context, IOrganizationLimitService limitService, IUserIdentityService userService)
    {
        var organizationId = userService.GetOrganizationId();
        if (!organizationId.HasValue)
        {
            await _next(context);
            return;
        }

        var endpoint = BuildEndpointKey(context);
        
        if (!await limitService.TryConsumeAsync(organizationId.Value, endpoint))
        {
            context.Response.StatusCode = 429;
            await context.Response.WriteAsync("Daily limit exceeded");
            return;
        }

        await _next(context);
    }

    /// <summary>
    /// Создает ключ эндпоинта для ограничения запросов на основе HTTP метода и пути.
    /// </summary>
    private static string BuildEndpointKey(HttpContext context)
    {
        var pathSegments = context.Request.Path.Value?.Split('/', StringSplitOptions.RemoveEmptyEntries) 
            ?? Array.Empty<string>();

        if (pathSegments.Length < 3)
            return $"{context.Request.Method}_{context.Request.Path}".Replace("/", "_");
            
        return $"{context.Request.Method}_{pathSegments[0]}_{pathSegments[1]}_{pathSegments[2]}";
    }
}