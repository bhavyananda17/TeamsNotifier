using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Teams;
using Microsoft.Bot.Schema;
using Newtonsoft.Json.Linq;
using TeamsNotifier.Storage;

namespace TeamsNotifier.Bot;

public class TeamsNotifierBot : TeamsActivityHandler
{
    private readonly IConversationReferenceStore _store;
    private readonly ILogger<TeamsNotifierBot> _logger;

    public TeamsNotifierBot(IConversationReferenceStore store, ILogger<TeamsNotifierBot> logger)
    {
        _store = store;
        _logger = logger;
    }

    protected override async Task OnMembersAddedAsync(IList<ChannelAccount> membersAdded, ITurnContext<IConversationUpdateActivity> turnContext, CancellationToken cancellationToken)
    {
        foreach (var member in membersAdded)
        {
            if (member.Id != turnContext.Activity.Recipient.Id)
            {
                await turnContext.SendActivityAsync("Hi! I'm the Teams Notifier bot. You're now registered for 1:1 notifications.", cancellationToken: cancellationToken);
            }
        }
    }

    protected override async Task OnMessageActivityAsync(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken)
    {
        var activity = turnContext.Activity;
        var userEmail = GetUserEmail(activity);

        if (string.IsNullOrEmpty(userEmail))
        {
            await turnContext.SendActivityAsync("Could not determine your email address. Please make sure you're in a 1:1 chat with me.", cancellationToken: cancellationToken);
            return;
        }

        var reference = activity.GetConversationReference();
        _store.AddOrUpdate(userEmail, reference);

        _logger.LogInformation("Registered conversation reference for user: {UserEmail}", userEmail);

        await turnContext.SendActivityAsync($"Registered! I'll send notifications to this 1:1 chat. Your registered email: {userEmail}", cancellationToken: cancellationToken);
    }

    protected override async Task OnConversationUpdateActivityAsync(ITurnContext<IConversationUpdateActivity> turnContext, CancellationToken cancellationToken)
    {
        if (turnContext.Activity.MembersAdded?.Any(m => m.Id == turnContext.Activity.Recipient.Id) == true)
        {
            var userEmail = GetUserEmail(turnContext.Activity);
            if (!string.IsNullOrEmpty(userEmail))
            {
                var reference = turnContext.Activity.GetConversationReference();
                _store.AddOrUpdate(userEmail, reference);
                _logger.LogInformation("Auto-registered conversation reference for user: {UserEmail}", userEmail);
            }
        }

        await base.OnConversationUpdateActivityAsync(turnContext, cancellationToken);
    }

    protected override async Task OnInstallationUpdateAddAsync(ITurnContext<IInstallationUpdateActivity> turnContext, CancellationToken cancellationToken)
    {
        // This fires when the bot is installed for a user (personal scope) or team
        // In personal scope, the activity.ChannelData contains the user's info
        var userEmail = GetUserEmail(turnContext.Activity);
        
        if (!string.IsNullOrEmpty(userEmail))
        {
            var reference = turnContext.Activity.GetConversationReference();
            _store.AddOrUpdate(userEmail, reference);
            _logger.LogInformation("Registered conversation reference via installation for user: {UserEmail}", userEmail);
        }

        // Also send welcome message
        await turnContext.SendActivityAsync("Hi! I'm the Teams Notifier bot. You're now registered for 1:1 notifications.", cancellationToken: cancellationToken);

        await base.OnInstallationUpdateAddAsync(turnContext, cancellationToken);
    }

    protected override async Task OnInstallationUpdateRemoveAsync(ITurnContext<IInstallationUpdateActivity> turnContext, CancellationToken cancellationToken)
    {
        var userEmail = GetUserEmail(turnContext.Activity);
        
        if (!string.IsNullOrEmpty(userEmail))
        {
            _store.Remove(userEmail);
            _logger.LogInformation("Removed conversation reference for uninstalled user: {UserEmail}", userEmail);
        }

        await base.OnInstallationUpdateRemoveAsync(turnContext, cancellationToken);
    }

    private static string? GetUserEmail(IActivity activity)
    {
        if (activity is IMessageActivity messageActivity && messageActivity.From != null)
        {
            if (!string.IsNullOrEmpty(messageActivity.From.AadObjectId))
            {
                return messageActivity.From.AadObjectId;
            }
            
            if (activity.ChannelId == "msteams" && activity.From.Properties != null)
            {
                if (activity.From.Properties.TryGetValue("aadObjectId", out var aadObjectId))
                {
                    var aadId = aadObjectId?.ToString();
                    if (!string.IsNullOrEmpty(aadId))
                        return aadId;
                }
            }
            
            if (!string.IsNullOrEmpty(activity.From.Name) && activity.From.Name.Contains("@"))
            {
                return activity.From.Name;
            }
        }

        if (activity.ChannelId == "msteams" && activity.From.Properties != null)
        {
            if (activity.From.Properties.TryGetValue("email", out var email))
            {
                var emailStr = email?.ToString();
                if (!string.IsNullOrEmpty(emailStr))
                    return emailStr;
            }
            if (activity.From.Properties.TryGetValue("userPrincipalName", out var upn))
            {
                var upnStr = upn?.ToString();
                if (!string.IsNullOrEmpty(upnStr))
                    return upnStr;
            }
        }

        return activity.From?.Name;
    }
}