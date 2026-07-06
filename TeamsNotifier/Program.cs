using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Integration.AspNet.Core;
using Microsoft.Bot.Connector.Authentication;
using TeamsNotifier.Bot;
using TeamsNotifier.Controllers;
using TeamsNotifier.Services;
using TeamsNotifier.Storage;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddControllers();

builder.Services.AddSingleton<CloudAdapter>(sp =>
{
    var logger = sp.GetRequiredService<ILogger<CloudAdapter>>();
    var auth = new ConfigurationBotFrameworkAuthentication(builder.Configuration);
    var adapter = new CloudAdapter(auth, logger);
    
    adapter.OnTurnError = async (turnContext, exception) =>
    {
        logger.LogError(exception, "Unhandled exception in bot");
        await turnContext.SendActivityAsync("Sorry, something went wrong!");
    };

    return adapter;
});

builder.Services.AddSingleton<INotifierService, NotifierService>();
builder.Services.AddSingleton<IConversationReferenceStore, ConversationReferenceStore>();
builder.Services.AddTransient<IBot, TeamsNotifierBot>();

var app = builder.Build();

if (app.Environment.IsDevelopment())
{
    app.UseDeveloperExceptionPage();
}

app.UseRouting();
app.MapControllers();

app.Run();