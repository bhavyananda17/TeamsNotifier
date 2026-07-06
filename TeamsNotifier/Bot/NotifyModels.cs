namespace TeamsNotifier.Bot;

public record NotifyResult(
    bool Sent,
    string AssigneeUpn,
    string? Error,         // "user_not_registered" | "teams_unavailable" | "exception"
    DateTime Timestamp
);

public record NotifyRequest(
    string EmailOrUpn,
    string Message,
    string? Title = null,
    string? ActionUrl = null,
    string? ActionTitle = null
);