using Microsoft.AspNetCore.Mvc;
using TeamsNotifier.Bot;
using TeamsNotifier.Storage;

namespace TeamsNotifier.Controllers;

[Route("api/notify")]
[ApiController]
public class NotifyController : ControllerBase
{
    private readonly INotifierService _notifier;
    private readonly IConversationReferenceStore _store;
    private readonly IConfiguration _configuration;

    public NotifyController(
        INotifierService notifier,
        IConversationReferenceStore store,
        IConfiguration configuration)
    {
        _notifier = notifier;
        _store = store;
        _configuration = configuration;
    }

    private bool ValidateApiKey()
    {
        var apiKey = _configuration["NotifierApiKey"];
        if (string.IsNullOrEmpty(apiKey))
        {
            return true; // No key configured = allow (dev only)
        }

        if (!Request.Headers.TryGetValue("X-Api-Key", out var providedKey))
        {
            return false;
        }

        return providedKey == apiKey;
    }

    [HttpPost]
    public async Task<ActionResult<NotifyResult>> Send([FromBody] NotifyRequest request)
    {
        if (!ValidateApiKey())
        {
            return Unauthorized(new { error = "Invalid or missing X-Api-Key header" });
        }

        if (string.IsNullOrWhiteSpace(request.EmailOrUpn))
        {
            return BadRequest(new { error = "EmailOrUpn is required" });
        }

        if (string.IsNullOrWhiteSpace(request.Message))
        {
            return BadRequest(new { error = "Message is required" });
        }

        var result = await _notifier.SendToAssigneeAsync(request);
        return result.Sent ? Ok(result) : StatusCode(500, result);
    }

    [HttpPost("card")]
    public async Task<ActionResult<NotifyResult>> SendCard([FromBody] NotifyRequest request)
    {
        if (!ValidateApiKey())
        {
            return Unauthorized(new { error = "Invalid or missing X-Api-Key header" });
        }

        if (string.IsNullOrWhiteSpace(request.EmailOrUpn))
        {
            return BadRequest(new { error = "EmailOrUpn is required" });
        }

        if (string.IsNullOrWhiteSpace(request.Message))
        {
            return BadRequest(new { error = "Message is required" });
        }

        var result = await _notifier.SendAdaptiveCardAsync(request);
        return result.Sent ? Ok(result) : StatusCode(500, result);
    }

    [HttpGet("status/{emailOrUpn}")]
    public ActionResult<object> GetStatus(string emailOrUpn)
    {
        if (!ValidateApiKey())
        {
            return Unauthorized(new { error = "Invalid or missing X-Api-Key header" });
        }

        var isRegistered = _store.TryGet(emailOrUpn) != null;

        return Ok(new
        {
            emailOrUpn,
            registered = isRegistered,
            message = isRegistered ? "User is registered and can receive notifications" : "User not registered - they must install the bot and send a message first"
        });
    }
}