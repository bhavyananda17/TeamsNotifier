using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Integration.AspNet.Core;
using Microsoft.Bot.Schema;
using Polly;
using Polly.Retry;
using TeamsNotifier.Services;
using TeamsNotifier.Storage;

namespace TeamsNotifier.Bot;

public interface INotifierService
{
    Task<NotifyResult> SendToAssigneeAsync(NotifyRequest request, CancellationToken cancellationToken = default);
    Task<NotifyResult> SendAdaptiveCardAsync(NotifyRequest request, CancellationToken cancellationToken = default);
}

public class NotifierService : INotifierService
{
    private readonly CloudAdapter _adapter;
    private readonly IConversationReferenceStore _store;
    private readonly ILogger<NotifierService> _logger;
    private readonly string _appId;
    private readonly AsyncRetryPolicy _retryPolicy;

    public NotifierService(
        CloudAdapter adapter,
        IConversationReferenceStore store,
        IConfiguration configuration,
        ILogger<NotifierService> logger)
    {
        _adapter = adapter;
        _store = store;
        _logger = logger;
        _appId = configuration["MicrosoftAppId"] ?? throw new InvalidOperationException("MicrosoftAppId not configured");

        _retryPolicy = Policy
            .Handle<Exception>()
            .WaitAndRetryAsync(
                3,
                attempt => TimeSpan.FromSeconds(Math.Pow(2, attempt)),
                (exception, timeSpan, attempt, context) =>
                {
                    _logger.LogWarning(exception, "Retry {Attempt} after {Delay}ms sending to Teams", attempt, timeSpan.TotalMilliseconds);
                });
    }

    public async Task<NotifyResult> SendToAssigneeAsync(NotifyRequest request, CancellationToken cancellationToken = default)
    {
        var reference = _store.TryGet(request.EmailOrUpn);

        if (reference == null)
        {
            _logger.LogWarning("No conversation reference found for user: {EmailOrUpn}", request.EmailOrUpn);
            return new NotifyResult(
                Sent: false,
                AssigneeUpn: request.EmailOrUpn,
                Error: "user_not_registered",
                Timestamp: DateTime.UtcNow);
        }

        try
        {
            await _retryPolicy.ExecuteAsync(async ct =>
            {
                await _adapter.ContinueConversationAsync(
                    _appId,
                    reference,
                    async (turnContext, ct) =>
                    {
                        var message = string.IsNullOrEmpty(request.Title)
                            ? request.Message
                            : $"**{request.Title}**\n\n{request.Message}";
                        await turnContext.SendActivityAsync(MessageFactory.Text(message), ct);
                    },
                    ct);
            }, cancellationToken);

            _logger.LogInformation("Sent proactive message to {EmailOrUpn}", request.EmailOrUpn);

            return new NotifyResult(
                Sent: true,
                AssigneeUpn: request.EmailOrUpn,
                Error: null,
                Timestamp: DateTime.UtcNow);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to send message to {EmailOrUpn}", request.EmailOrUpn);
            return new NotifyResult(
                Sent: false,
                AssigneeUpn: request.EmailOrUpn,
                Error: "teams_unavailable",
                Timestamp: DateTime.UtcNow);
        }
    }

    public async Task<NotifyResult> SendAdaptiveCardAsync(NotifyRequest request, CancellationToken cancellationToken = default)
    {
        var reference = _store.TryGet(request.EmailOrUpn);

        if (reference == null)
        {
            _logger.LogWarning("No conversation reference found for user: {EmailOrUpn}", request.EmailOrUpn);
            return new NotifyResult(
                Sent: false,
                AssigneeUpn: request.EmailOrUpn,
                Error: "user_not_registered",
                Timestamp: DateTime.UtcNow);
        }

        var card = BuildAdaptiveCard(request);

        try
        {
            await _retryPolicy.ExecuteAsync(async ct =>
            {
                await _adapter.ContinueConversationAsync(
                    _appId,
                    reference,
                    async (turnContext, ct) =>
                    {
                        var attachment = new Attachment
                        {
                            ContentType = "application/vnd.microsoft.card.adaptive",
                            Content = card
                        };
                        await turnContext.SendActivityAsync(MessageFactory.Attachment(attachment), ct);
                    },
                    ct);
            }, cancellationToken);

            _logger.LogInformation("Sent Adaptive Card to {EmailOrUpn}", request.EmailOrUpn);

            return new NotifyResult(
                Sent: true,
                AssigneeUpn: request.EmailOrUpn,
                Error: null,
                Timestamp: DateTime.UtcNow);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to send Adaptive Card to {EmailOrUpn}", request.EmailOrUpn);
            return new NotifyResult(
                Sent: false,
                AssigneeUpn: request.EmailOrUpn,
                Error: "teams_unavailable",
                Timestamp: DateTime.UtcNow);
        }
    }

    private static object BuildAdaptiveCard(NotifyRequest request)
    {
        var card = new
        {
            type = "AdaptiveCard",
            version = "1.4",
            body = new object[]
            {
                new
                {
                    type = "TextBlock",
                    text = request.Title ?? "Workflow Notification",
                    weight = "Bolder",
                    size = "Medium",
                    wrap = true
                },
                new
                {
                    type = "TextBlock",
                    text = request.Message,
                    wrap = true,
                    spacing = "Medium"
                }
            },
            actions = new object[]
            {
                new
                {
                    type = "Action.OpenUrl",
                    title = request.ActionTitle ?? "Open Workflow",
                    url = request.ActionUrl ?? "https://teams.microsoft.com"
                }
            }
        };

        return card;
    }
}