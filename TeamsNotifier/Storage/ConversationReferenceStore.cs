using Microsoft.Bot.Schema;
using System.Collections.Concurrent;

namespace TeamsNotifier.Storage;

public class ConversationReferenceStore : IConversationReferenceStore
{
    private readonly ConcurrentDictionary<string, ConversationReference> _references = new();

    public void AddOrUpdate(string key, ConversationReference reference)
    {
        _references.AddOrUpdate(key.ToLowerInvariant(), reference, (_, _) => reference);
    }

    public ConversationReference? TryGet(string key)
    {
        _references.TryGetValue(key.ToLowerInvariant(), out var reference);
        return reference;
    }

    public bool Remove(string key)
    {
        return _references.TryRemove(key.ToLowerInvariant(), out _);
    }

    public IEnumerable<KeyValuePair<string, ConversationReference>> GetAll()
    {
        return _references;
    }

    public int Count => _references.Count;
}