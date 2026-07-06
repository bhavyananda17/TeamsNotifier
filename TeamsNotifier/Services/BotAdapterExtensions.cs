using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Integration.AspNet.Core;
using Microsoft.Bot.Connector.Authentication;
using TeamsNotifier.Bot;
using TeamsNotifier.Storage;

namespace TeamsNotifier.Services;

public static class BotAdapterExtensions
{
    public static IServiceCollection AddTeamsNotifierBot(this IServiceCollection services, IConfiguration configuration)
    {
        services.AddSingleton<CloudAdapter>(sp =>
        {
            var logger = sp.GetRequiredService<ILogger<CloudAdapter>>();
            var auth = new ConfigurationBotFrameworkAuthentication(configuration);
            var adapter = new CloudAdapter(auth, logger);
            
            adapter.OnTurnError = async (turnContext, exception) =>
            {
                logger.LogError(exception, "Unhandled exception in bot");
                await turnContext.SendActivityAsync("Sorry, something went wrong!");
            };

            return adapter;
        });

        services.AddSingleton<INotifierService, NotifierService>();
        services.AddSingleton<IConversationReferenceStore, ConversationReferenceStore>();
        services.AddTransient<IBot, TeamsNotifierBot>();

        return services;
    }
}