using Microsoft.Bot.Schema;

namespace TeamsNotifier.Storage;

public interface IConversationReferenceStore
{
    ConversationReference? TryGet(string key);
    void AddOrUpdate(string key, ConversationReference reference);
    bool Remove(string key);
    IEnumerable<KeyValuePair<string, ConversationReference>> GetAll();
}