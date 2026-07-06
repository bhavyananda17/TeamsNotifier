# TeamsNotifier - Standalone C# Bot for Proactive 1:1 Teams Notifications

---

## Problem Solved

Microsoft Graph API cannot send proactive 1:1 chat messages with app-only (daemon) authentication — only delegated (interactive user) auth works. This bot uses **Bot Framework + Proactive Messaging** to send notifications from any background service/daemon to a specific user's personal Teams chat without requiring a user to be signed in.

---

## Architecture

```
Background Workflow (daemon, scheduled job, etc.)
        │
        ▼
┌─────────────────────────────────────────┐
│         NotifierService                 │  ← Public API: SendToAssigneeAsync(NotifyRequest)
│   (injected via DI, singleton)          │     Returns NotifyResult with delivery status
└─────────────────────────────────────────┘
        │
        ▼
┌─────────────────────────────────────────┐
│   ConversationReferenceStore            │  ← In-memory ConcurrentDictionary<string, ConversationReference>
│   (keyed by email/UPN)                  │     Swap for Azure Table/Redis in production
└─────────────────────────────────────────┘
        │
        ▼
┌─────────────────────────────────────────┐
│   CloudAdapter.ContinueConversationAsync│  ← Bot Framework proactive messaging
│   (uses MicrosoftAppId + App Password)  │     With Polly retry (exponential backoff)
└─────────────────────────────────────────┘
        │
        ▼
   Microsoft Teams (1:1 chat with bot)
```

---

## Project Structure

| File | Purpose |
|------|---------|
| `Program.cs` | DI setup, CloudAdapter with `ConfigurationBotFrameworkAuthentication`, endpoint routing |
| `appsettings.json` | `MicrosoftAppId`, `MicrosoftAppPassword`, `MicrosoftAppType: MultiTenant`, `NotifierApiKey` |
| `Bot/TeamsNotifierBot.cs` | `TeamsActivityHandler` — captures `ConversationReference` on install, message, conversation update; extracts user email/UPN from Teams channel data |
| `Bot/NotifierService.cs` | `INotifierService` — `SendToAssigneeAsync(NotifyRequest)` and `SendAdaptiveCardAsync(NotifyRequest)` with retry + `NotifyResult` |
| `Bot/NotifyModels.cs` | `NotifyRequest`, `NotifyResult` records |
| `Storage/IConversationReferenceStore.cs` | Interface for pluggable storage |
| `Storage/ConversationReferenceStore.cs` | In-memory `ConcurrentDictionary` implementation (MVP) |
| `Services/BotAdapterExtensions.cs` | `AddTeamsNotifierBot()` — registers all services in one call |
| `Controllers/MessagesController.cs` | `POST /api/messages` — Bot Framework webhook endpoint |
| `Controllers/NotifyController.cs` | `POST /api/notify`, `POST /api/notify/card`, `GET /api/notify/status/{email}` — secured with API key |

---

## Key Features Implemented

### 1. **Org-Wide Install Support** (`OnInstallationUpdateAddAsync`)
- Captures `ConversationReference` the moment an admin pushes the bot to users
- No user action needed — user is registered automatically on install
- `OnInstallationUpdateRemoveAsync` cleans up on uninstall

### 2. **API Key Authentication** (`NotifyController`)
- All `/api/notify/*` endpoints require `X-Api-Key` header
- Configure via `appsettings.json` → `NotifierApiKey`
- Dev mode: if not configured, allows all requests

### 3. **Retry with Polly** (Exponential Backoff)
- 3 retries with 2^attempt second delays (2s, 4s, 8s)
- Handles transient Teams failures gracefully
- Logs each retry attempt

### 4. **Adaptive Cards** (`SendAdaptiveCardAsync`)
- Professional card format with title, message, and action button
- JSON-based card built inline (no external template files needed)

### 5. **Delivery Status Tracking** (`NotifyResult`)
```csharp
public record NotifyResult(
    bool Sent,
    string AssigneeUpn,
    string? Error,         // "user_not_registered" | "teams_unavailable" | "exception"
    DateTime Timestamp
);
```

### 6. **Registration Status Endpoint**
- `GET /api/notify/status/{emailOrUpn}` — checks if user is registered before sending
- Returns `{ registered: true/false, message: "..." }`

---

## Flow

1. **Admin installs bot org-wide** (or user installs in personal scope)
2. **`OnInstallationUpdateAddAsync` fires** → captures `ConversationReference` automatically
3. **User can also send "hi"** → `OnMessageActivityAsync` fires as fallback
4. **Background service calls** `await _notifier.SendToAssigneeAsync(new NotifyRequest { EmailOrUpn = "user@domain.com", Message = "Your deployment finished", Title = "Build Complete", ActionUrl = "https://..." })`
5. `NotifierService` looks up reference → retries up to 3x → sends message/card to that 1:1 chat
6. Returns `NotifyResult` with `Sent=true/false` and error code if failed

---

## Key NuGet Packages

| Package | Version | Purpose |
|---------|---------|---------|
| `Microsoft.Bot.Builder.Integration.AspNet.Core` | 4.22.4 | ASP.NET Core integration, CloudAdapter |
| `Microsoft.Bot.Builder` | 4.22.4 | Core bot SDK, TeamsActivityHandler |
| `Microsoft.Bot.Builder.Teams` | 4.3.0-beta1 | Teams-specific schema |
| `Microsoft.Bot.Connector.Authentication` | 4.22.4 | `ConfigurationBotFrameworkAuthentication` |
| `Polly` | 8.7.0 | Retry policies with exponential backoff |

---

## Configuration (`appsettings.json`)

```json
{
  "MicrosoftAppId": "<AZURE_AD_APP_CLIENT_ID>",
  "MicrosoftAppPassword": "<AZURE_AD_APP_CLIENT_SECRET>",
  "MicrosoftAppType": "MultiTenant",
  "NotifierApiKey": "<OPTIONAL_API_KEY_FOR_NOTIFY_ENDPOINTS>"
}
```

> **Note:** Use **MultiTenant** so any Azure AD user can install the bot.
> **Note:** `NotifierApiKey` is optional — if omitted, endpoints allow all requests (dev only).

---

## Local Testing

```bash
# 1. Start ngrok
ngrok http 5000

# 2. Update Bot Framework registration → Messaging endpoint: https://<ngrok-id>.ngrok-free.app/api/messages

# 3. Run app
dotnet run --project TeamsNotifier/TeamsNotifier.csproj

# 4. In Teams: search for bot by name → install in personal scope
#    Bot auto-registers on install (no "hi" needed), or send "hi" as fallback

# 5. Test proactive plain text:
curl -X POST http://localhost:5000/api/notify \
  -H "Content-Type: application/json" \
  -H "X-Api-Key: your-api-key" \
  -d '{"emailOrUpn":"user@domain.com","message":"Test from background job","title":"Build Complete"}'

# 6. Test Adaptive Card:
curl -X POST http://localhost:5000/api/notify/card \
  -H "Content-Type: application/json" \
  -H "X-Api-Key: your-api-key" \
  -d '{"emailOrUpn":"user@domain.com","message":"Contract review needed","title":"Workflow Assigned","actionUrl":"https://app.example.com/workflow/123","actionTitle":"Open Workflow"}'

# 7. Check registration status:
curl -X GET http://localhost:5000/api/notify/status/user@domain.com \
  -H "X-Api-Key: your-api-key"
```

---

## Adaptive Card Example (rendered in Teams)

```
┌─────────────────────────────────┐
│ 📋 Workflow Assigned to You     │
│─────────────────────────────────│
│ Task:    Contract Review        │
│ By:      bob@company.com        │
│ Due:     2025-02-28             │
│ Priority: High                  │
│─────────────────────────────────│
│ [Open Workflow]                 │
└─────────────────────────────────┘
```

---

## Production Deployment Checklist

| Step | Action |
|------|--------|
| **Persistent Storage** | Replace `ConversationReferenceStore` with Azure Table Storage / Redis / SQL implementation of `IConversationReferenceStore` |
| **Secrets** | Store `MicrosoftAppPassword` and `NotifierApiKey` in Azure Key Vault / App Service Settings — never in `appsettings.json` |
| **Scaling** | In-memory store works for single instance; multi-instance requires distributed cache |
| **Monitoring** | Add Application Insights; log `OnTurnError` and proactive send failures |
| **API Key Rotation** | Implement key rotation strategy for `NotifierApiKey` |

---

## Usage in Background Workflow

```csharp
// Any .NET service (Worker Service, Azure Function, Console App, etc.)
public class MyBackgroundJob
{
    private readonly INotifierService _notifier;

    public MyBackgroundJob(INotifierService notifier) => _notifier = notifier;

    public async Task RunAsync()
    {
        // Plain text
        var result = await _notifier.SendToAssigneeAsync(new NotifyRequest
        {
            EmailOrUpn = "jane.doe@company.com",
            Message = "✅ Build #1234 deployed to staging",
            Title = "Deployment Complete",
            ActionUrl = "https://github.com/org/repo/actions/runs/1234",
            ActionTitle = "View Build"
        });

        if (!result.Sent)
        {
            _logger.LogWarning("Failed to notify {User}: {Error}", result.AssigneeUpn, result.Error);
        }

        // Or Adaptive Card
        await _notifier.SendAdaptiveCardAsync(new NotifyRequest
        {
            EmailOrUpn = "jane.doe@company.com",
            Title = "Workflow Assigned",
            Message = "Contract review needed — due Friday",
            ActionUrl = "https://app.example.com/workflow/456",
            ActionTitle = "Review Now"
        });
    }
}
```

No user context, no interactive login — works as a true daemon.

---

## Error Codes Reference

| Error | Meaning | Action |
|-------|---------|--------|
| `user_not_registered` | No conversation reference found | User must install bot or admin must push org-wide |
| `teams_unavailable` | Teams API failed after retries | Check connectivity, retry later |
| `exception` | Unexpected error | Check logs for details |