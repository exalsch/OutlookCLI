using System.Runtime.InteropServices;
using OutlookCLI.Models;

namespace OutlookCLI.Services;

/// <summary>
/// Outlook service using late-binding COM interop (no PIAs required).
/// All COM collection iteration uses index-based access (1-based) to avoid
/// enumerator leaks that cause "too many items open" errors.
/// </summary>
public class OutlookService : IOutlookService
{
    private dynamic? _app;
    private dynamic? _namespace;
    private readonly List<object> _comObjects = new();

    // OlDefaultFolders enum values
    private const int olFolderInbox = 6;
    private const int olFolderCalendar = 9;
    private const int olFolderSentMail = 5;
    private const int olFolderDrafts = 16;
    private const int olFolderDeletedItems = 3;
    private const int olFolderOutbox = 4;
    private const int olFolderJunk = 23;
    private const int olFolderContacts = 10;

    // OlItemType enum values
    private const int olMailItem = 0;
    private const int olAppointmentItem = 1;

    // OlMailRecipientType enum values
    private const int olTo = 1;
    private const int olCC = 2;

    /// <summary>
    /// Maps localized folder names to OlDefaultFolders constants.
    /// Keys are lowercase. Covers: EN, DE, SV, CS, DA, NO, FI, NL, FR, ES, IT,
    /// PT (BR/PT), PL, HU, RO, TR, JA, ZH, KO, RU, AR, HE.
    /// </summary>
    private static readonly Dictionary<string, int> FolderAliases = new(StringComparer.OrdinalIgnoreCase)
    {
        // --- Inbox ---
        ["inbox"] = olFolderInbox,
        // DE
        ["posteingang"] = olFolderInbox,
        // SV
        ["inkorg"] = olFolderInbox,
        // CS
        ["doručená pošta"] = olFolderInbox,
        // DA
        ["indbakke"] = olFolderInbox,
        // NO
        ["innboks"] = olFolderInbox,
        // FI
        ["saapuneet"] = olFolderInbox,
        // NL
        ["postvak in"] = olFolderInbox,
        // FR
        ["boîte de réception"] = olFolderInbox,
        ["boite de reception"] = olFolderInbox,
        // ES
        ["bandeja de entrada"] = olFolderInbox,
        // IT
        ["posta in arrivo"] = olFolderInbox,
        // PT
        ["caixa de entrada"] = olFolderInbox,
        // PL
        ["skrzynka odbiorcza"] = olFolderInbox,
        // HU
        ["beérkezett üzenetek"] = olFolderInbox,
        ["beerkezett uzenetek"] = olFolderInbox,
        // RO
        ["mesaje primite"] = olFolderInbox,
        // TR
        ["gelen kutusu"] = olFolderInbox,
        // JA
        ["受信トレイ"] = olFolderInbox,
        // ZH
        ["收件箱"] = olFolderInbox,
        // KO
        ["받은 편지함"] = olFolderInbox,
        // RU
        ["входящие"] = olFolderInbox,
        // AR
        ["علبة الوارد"] = olFolderInbox,
        // HE
        ["דואר נכנס"] = olFolderInbox,

        // --- Sent Items ---
        ["sent"] = olFolderSentMail,
        ["sent mail"] = olFolderSentMail,
        ["sentmail"] = olFolderSentMail,
        ["sent items"] = olFolderSentMail,
        // DE
        ["gesendete elemente"] = olFolderSentMail,
        // SV
        ["skickat"] = olFolderSentMail,
        // CS
        ["odeslaná pošta"] = olFolderSentMail,
        // DA
        ["sendt post"] = olFolderSentMail,
        // NO
        ["sendte elementer"] = olFolderSentMail,
        // FI
        ["lähetetyt"] = olFolderSentMail,
        // NL
        ["verzonden items"] = olFolderSentMail,
        // FR
        ["éléments envoyés"] = olFolderSentMail,
        ["elements envoyes"] = olFolderSentMail,
        // ES
        ["elementos enviados"] = olFolderSentMail,
        // IT
        ["posta inviata"] = olFolderSentMail,
        // PT
        ["itens enviados"] = olFolderSentMail,
        // PL
        ["elementy wysłane"] = olFolderSentMail,
        ["elementy wyslane"] = olFolderSentMail,
        // HU
        ["elküldött elemek"] = olFolderSentMail,
        ["elkuldott elemek"] = olFolderSentMail,
        // RO
        ["elemente trimise"] = olFolderSentMail,
        // TR
        ["gönderilmiş öğeler"] = olFolderSentMail,
        ["gonderilmis ogeler"] = olFolderSentMail,
        // JA
        ["送信済みアイテム"] = olFolderSentMail,
        // ZH
        ["已发送邮件"] = olFolderSentMail,
        // KO
        ["보낸 편지함"] = olFolderSentMail,
        // RU
        ["отправленные"] = olFolderSentMail,
        // AR
        ["العناصر المرسلة"] = olFolderSentMail,
        // HE
        ["פריטים שנשלחו"] = olFolderSentMail,

        // --- Drafts ---
        ["drafts"] = olFolderDrafts,
        // DE
        ["entwürfe"] = olFolderDrafts,
        // SV
        ["utkast"] = olFolderDrafts,
        // CS
        ["koncepty"] = olFolderDrafts,
        // DA
        ["kladder"] = olFolderDrafts,
        // NO
        ["kladd"] = olFolderDrafts,
        // FI
        ["luonnokset"] = olFolderDrafts,
        // NL
        ["concepten"] = olFolderDrafts,
        // FR
        ["brouillons"] = olFolderDrafts,
        // ES
        ["borradores"] = olFolderDrafts,
        // IT
        ["bozze"] = olFolderDrafts,
        // PT
        ["rascunhos"] = olFolderDrafts,
        // PL
        ["wersje robocze"] = olFolderDrafts,
        // HU
        ["piszkozatok"] = olFolderDrafts,
        // RO
        ["ciorne"] = olFolderDrafts,
        // TR
        ["taslaklar"] = olFolderDrafts,
        // JA
        ["下書き"] = olFolderDrafts,
        // ZH
        ["草稿"] = olFolderDrafts,
        // KO
        ["임시 보관함"] = olFolderDrafts,
        // RU
        ["черновики"] = olFolderDrafts,
        // AR
        ["المسودات"] = olFolderDrafts,
        // HE
        ["טיוטות"] = olFolderDrafts,

        // --- Deleted Items ---
        ["deleted"] = olFolderDeletedItems,
        ["deleted items"] = olFolderDeletedItems,
        ["deleteditems"] = olFolderDeletedItems,
        ["trash"] = olFolderDeletedItems,
        // DE
        ["gelöschte elemente"] = olFolderDeletedItems,
        // SV
        ["borttagna objekt"] = olFolderDeletedItems,
        // CS
        ["odstraněná pošta"] = olFolderDeletedItems,
        // DA
        ["slettet post"] = olFolderDeletedItems,
        // NO
        ["slettede elementer"] = olFolderDeletedItems,
        // FI
        ["poistetut"] = olFolderDeletedItems,
        // NL
        ["verwijderde items"] = olFolderDeletedItems,
        // FR
        ["éléments supprimés"] = olFolderDeletedItems,
        ["elements supprimes"] = olFolderDeletedItems,
        // ES
        ["elementos eliminados"] = olFolderDeletedItems,
        // IT
        ["posta eliminata"] = olFolderDeletedItems,
        // PT-BR
        ["itens excluídos"] = olFolderDeletedItems,
        ["itens excluidos"] = olFolderDeletedItems,
        // PT-PT
        ["itens eliminados"] = olFolderDeletedItems,
        // PL
        ["elementy usunięte"] = olFolderDeletedItems,
        ["elementy usuniete"] = olFolderDeletedItems,
        // HU
        ["törölt elemek"] = olFolderDeletedItems,
        ["torolt elemek"] = olFolderDeletedItems,
        // RO
        ["elemente șterse"] = olFolderDeletedItems,
        ["elemente sterse"] = olFolderDeletedItems,
        // TR
        ["silinmiş öğeler"] = olFolderDeletedItems,
        ["silinmis ogeler"] = olFolderDeletedItems,
        // JA
        ["削除済みアイテム"] = olFolderDeletedItems,
        // ZH
        ["已删除邮件"] = olFolderDeletedItems,
        // KO
        ["지운 편지함"] = olFolderDeletedItems,
        // RU
        ["удалённые"] = olFolderDeletedItems,
        ["удаленные"] = olFolderDeletedItems,
        // AR
        ["العناصر المحذوفة"] = olFolderDeletedItems,
        // HE
        ["פריטים שנמחקו"] = olFolderDeletedItems,

        // --- Outbox ---
        ["outbox"] = olFolderOutbox,
        // DE
        ["postausgang"] = olFolderOutbox,
        // SV
        ["utkorg"] = olFolderOutbox,
        // CS
        ["pošta k odeslání"] = olFolderOutbox,
        // DA
        ["udbakke"] = olFolderOutbox,
        // NO
        ["utboks"] = olFolderOutbox,
        // FI
        ["lähtevät"] = olFolderOutbox,
        // NL
        ["postvak uit"] = olFolderOutbox,
        // FR
        ["boîte d'envoi"] = olFolderOutbox,
        ["boite d'envoi"] = olFolderOutbox,
        // ES
        ["bandeja de salida"] = olFolderOutbox,
        // IT
        ["posta in uscita"] = olFolderOutbox,
        // PT
        ["caixa de saída"] = olFolderOutbox,
        ["caixa de saida"] = olFolderOutbox,
        // PL
        ["skrzynka nadawcza"] = olFolderOutbox,
        // HU
        ["postázandó üzenetek"] = olFolderOutbox,
        ["postazando uzenetek"] = olFolderOutbox,
        // RO
        ["mesaje de ieșire"] = olFolderOutbox,
        ["mesaje de iesire"] = olFolderOutbox,
        // TR
        ["giden kutusu"] = olFolderOutbox,
        // JA
        ["送信トレイ"] = olFolderOutbox,
        // ZH
        ["发件箱"] = olFolderOutbox,
        // KO
        ["보낼 편지함"] = olFolderOutbox,
        // RU
        ["исходящие"] = olFolderOutbox,
        // AR
        ["علبة الصادر"] = olFolderOutbox,
        // HE
        ["דואר יוצא"] = olFolderOutbox,

        // --- Junk Email ---
        ["junk"] = olFolderJunk,
        ["junk mail"] = olFolderJunk,
        ["junkmail"] = olFolderJunk,
        ["junk email"] = olFolderJunk,
        ["spam"] = olFolderJunk,
        // DE
        ["junk-e-mail"] = olFolderJunk,
        // SV
        ["skräppost"] = olFolderJunk,
        // CS
        ["nevyžádaná pošta"] = olFolderJunk,
        // DA
        ["uønsket mail"] = olFolderJunk,
        // NO
        ["søppelpost"] = olFolderJunk,
        // FI
        ["roskaposti"] = olFolderJunk,
        // NL
        ["ongewenste e-mail"] = olFolderJunk,
        // FR
        ["courrier indésirable"] = olFolderJunk,
        ["courrier indesirable"] = olFolderJunk,
        // ES
        ["correo no deseado"] = olFolderJunk,
        // IT
        ["posta indesiderata"] = olFolderJunk,
        // PT
        ["lixo eletrônico"] = olFolderJunk,
        ["lixo eletronico"] = olFolderJunk,
        // PL
        ["wiadomości-śmieci"] = olFolderJunk,
        ["wiadomosci-smieci"] = olFolderJunk,
        // HU
        ["levélszemét"] = olFolderJunk,
        ["levelszemet"] = olFolderJunk,
        // RO
        ["e-mail nedorit"] = olFolderJunk,
        // TR
        ["gereksiz e-posta"] = olFolderJunk,
        // JA
        ["迷惑メール"] = olFolderJunk,
        // ZH
        ["垃圾邮件"] = olFolderJunk,
        // KO
        ["정크 메일"] = olFolderJunk,
        // RU
        ["нежелательная почта"] = olFolderJunk,
        // AR
        ["البريد الإلكتروني غير الهام"] = olFolderJunk,
        // HE
        ["דואר זבל"] = olFolderJunk,

        // --- Calendar ---
        ["calendar"] = olFolderCalendar,
        // DE
        ["kalender"] = olFolderCalendar,
        // SV (same as DE)
        // CS
        ["kalendář"] = olFolderCalendar,
        // DA (same as DE)
        // NO (same as DE)
        // FI
        ["kalenteri"] = olFolderCalendar,
        // NL
        ["agenda"] = olFolderCalendar,
        // FR
        ["calendrier"] = olFolderCalendar,
        // ES
        ["calendario"] = olFolderCalendar,
        // IT (same as ES)
        // PT (same as ES)
        // PL
        ["kalendarz"] = olFolderCalendar,
        // HU
        ["naptár"] = olFolderCalendar,
        ["naptar"] = olFolderCalendar,
        // RO (same as EN)
        // TR
        ["takvim"] = olFolderCalendar,
        // JA
        ["予定表"] = olFolderCalendar,
        // ZH
        ["日历"] = olFolderCalendar,
        // KO
        ["일정"] = olFolderCalendar,
        // RU
        ["календарь"] = olFolderCalendar,
        // AR
        ["التقويم"] = olFolderCalendar,
        // HE
        ["לוח שנה"] = olFolderCalendar,

        // --- Contacts ---
        ["contacts"] = olFolderContacts,
        // DE
        ["kontakte"] = olFolderContacts,
        // SV
        ["kontakter"] = olFolderContacts,
        // CS
        ["kontakty"] = olFolderContacts,
        // DA
        ["kontaktpersoner"] = olFolderContacts,
        // NO (same as SV)
        // FI
        ["yhteystiedot"] = olFolderContacts,
        // NL
        ["contactpersonen"] = olFolderContacts,
        // FR (same as EN)
        // ES
        ["contactos"] = olFolderContacts,
        // IT
        ["contatti"] = olFolderContacts,
        // PT-BR
        ["contatos"] = olFolderContacts,
        // PL (same as CS)
        // HU
        ["névjegyek"] = olFolderContacts,
        ["nevjegyek"] = olFolderContacts,
        // RO
        ["persoane de contact"] = olFolderContacts,
        // TR
        ["kişiler"] = olFolderContacts,
        ["kisiler"] = olFolderContacts,
        // JA
        ["連絡先"] = olFolderContacts,
        // ZH
        ["联系人"] = olFolderContacts,
        // KO
        ["연락처"] = olFolderContacts,
        // RU
        ["контакты"] = olFolderContacts,
        // AR
        ["جهات الاتصال"] = olFolderContacts,
        // HE
        ["אנשי קשר"] = olFolderContacts,
    };

    public void Initialize()
    {
        var outlookType = Type.GetTypeFromProgID("Outlook.Application");
        if (outlookType == null)
        {
            throw new InvalidOperationException("Outlook is not installed on this machine.");
        }

        _app = Activator.CreateInstance(outlookType);
        if (_app == null)
        {
            throw new InvalidOperationException("Failed to create Outlook application instance.");
        }

        Track(_app);
        _namespace = _app.GetNamespace("MAPI");
        Track(_namespace);
    }

    private void Track(object comObject) => _comObjects.Add(comObject);

    private static void Release(object? comObject)
    {
        if (comObject != null)
        {
            try { Marshal.ReleaseComObject(comObject); } catch { }
        }
    }

    public void Dispose()
    {
        foreach (var obj in _comObjects.AsEnumerable().Reverse())
        {
            Release(obj);
        }
        _comObjects.Clear();
        GC.Collect();
        GC.WaitForPendingFinalizers();
    }

    private dynamic GetFolder(string? folderName)
    {
        if (_namespace == null) throw new InvalidOperationException("Service not initialized");

        if (string.IsNullOrEmpty(folderName))
        {
            var folder = _namespace.GetDefaultFolder(olFolderInbox);
            Track(folder);
            return folder;
        }

        if (FolderAliases.TryGetValue(folderName, out int folderId))
        {
            var folder = _namespace.GetDefaultFolder(folderId);
            Track(folder);
            return folder;
        }

        var foundFolder = FindFolderByName(folderName);
        if (foundFolder == null)
        {
            throw new DirectoryNotFoundException($"Folder not found: {folderName}");
        }
        return foundFolder;
    }

    private dynamic? FindFolderByName(string folderName)
    {
        if (_namespace == null) return null;

        // Search within the default mailbox only
        var inbox = _namespace.GetDefaultFolder(olFolderInbox);
        var mailboxRoot = inbox.Parent;
        Release(inbox);

        var found = SearchFolderRecursive(mailboxRoot, folderName);
        if (found != null)
        {
            if (!ReferenceEquals(found, mailboxRoot)) Release(mailboxRoot);
            return found;
        }
        Release(mailboxRoot);

        return null;
    }

    private dynamic? SearchFolderRecursive(dynamic parent, string folderName)
    {
        string parentName = parent.Name;
        if (parentName.Equals(folderName, StringComparison.OrdinalIgnoreCase))
            return parent;

        var subfolders = parent.Folders;
        int subCount = subfolders.Count;
        for (int i = 1; i <= subCount; i++)
        {
            var subfolder = subfolders[i];
            var found = SearchFolderRecursive(subfolder, folderName);
            if (found != null)
            {
                if (!ReferenceEquals(found, subfolder)) Release(subfolder);
                Release(subfolders);
                return found;
            }
            Release(subfolder);
        }
        Release(subfolders);

        return null;
    }

    private dynamic GetDeletedItemsFolder()
    {
        if (_namespace == null) throw new InvalidOperationException("Service not initialized");
        var folder = _namespace.GetDefaultFolder(olFolderDeletedItems);
        Track(folder);
        return folder;
    }

    public IEnumerable<FolderInfo> GetMailFolders()
    {
        if (_namespace == null) throw new InvalidOperationException("Service not initialized");

        var result = new List<FolderInfo>();

        // Start from the default Inbox's parent (the user's mailbox root)
        // instead of _namespace.Folders which includes all stores/shared mailboxes
        var inbox = _namespace.GetDefaultFolder(olFolderInbox);
        var mailboxRoot = inbox.Parent;
        Release(inbox);

        try
        {
            CollectFolders(mailboxRoot, "", result);
        }
        finally
        {
            Release(mailboxRoot);
        }

        return result;
    }

    private void CollectFolders(dynamic folder, string parentPath, List<FolderInfo> result)
    {
        string folderName = folder.Name;
        var fullPath = string.IsNullOrEmpty(parentPath) ? folderName : $"{parentPath}/{folderName}";

        // Only include mail folders
        int defaultItemType = folder.DefaultItemType;
        if (defaultItemType == olMailItem)
        {
            var items = folder.Items;
            int itemCount = items.Count;
            Release(items);

            result.Add(new FolderInfo(
                folderName,
                fullPath,
                itemCount,
                folder.UnReadItemCount
            ));
        }

        var subfolders = folder.Folders;
        int subCount = subfolders.Count;
        for (int i = 1; i <= subCount; i++)
        {
            var subfolder = subfolders[i];
            try
            {
                CollectFolders(subfolder, fullPath, result);
            }
            finally
            {
                Release(subfolder);
            }
        }
        Release(subfolders);
    }

    public IEnumerable<MailMessageSummary> GetMailList(string? folderName, bool unreadOnly, int limit)
    {
        var folder = GetFolder(folderName);
        var items = folder.Items;
        Track(items);

        items.Sort("[ReceivedTime]", true);

        if (unreadOnly)
        {
            items = items.Restrict("[UnRead] = True");
            Track(items);
        }

        var result = new List<MailMessageSummary>();
        int count = 0;
        string folderNameStr = folder.Name;
        int totalItems = items.Count;

        for (int i = 1; i <= totalItems && count < limit; i++)
        {
            dynamic? item = null;
            try
            {
                item = items[i];
                int itemClass = item.Class;
                if (itemClass == 43) // olMail
                {
                    result.Add(MapToSummary(item, folderNameStr));
                    count++;
                }
            }
            catch
            {
                // Skip items that can't be read
            }
            finally
            {
                Release(item);
            }
        }

        return result;
    }

    public IEnumerable<MailMessage> GetMailListFull(string? folderName, bool unreadOnly, int limit)
    {
        var folder = GetFolder(folderName);
        var items = folder.Items;
        Track(items);

        items.Sort("[ReceivedTime]", true);

        if (unreadOnly)
        {
            items = items.Restrict("[UnRead] = True");
            Track(items);
        }

        var result = new List<MailMessage>();
        int count = 0;
        string folderNameStr = folder.Name;
        int totalItems = items.Count;

        for (int i = 1; i <= totalItems && count < limit; i++)
        {
            dynamic? item = null;
            try
            {
                item = items[i];
                int itemClass = item.Class;
                if (itemClass == 43) // olMail
                {
                    result.Add(MapToFull(item, folderNameStr));
                    count++;
                }
            }
            catch
            {
                // Skip items that can't be read
            }
            finally
            {
                Release(item);
            }
        }

        return result;
    }

    public MailMessage? GetMail(string entryId)
    {
        if (_namespace == null) throw new InvalidOperationException("Service not initialized");

        try
        {
            var item = _namespace.GetItemFromID(entryId);
            Track(item);

            int itemClass = item.Class;
            if (itemClass == 43) // olMail
            {
                var folder = item.Parent;
                Track(folder);
                string folderName = folder.Name;
                return MapToFull(item, folderName);
            }
        }
        catch
        {
            return null;
        }

        return null;
    }

    public void OpenMail(string entryId)
    {
        if (_namespace == null) throw new InvalidOperationException("Service not initialized");

        var item = _namespace.GetItemFromID(entryId);
        int itemClass = item.Class;
        if (itemClass != 43) // olMail
        {
            Marshal.ReleaseComObject(item);
            throw new InvalidOperationException("Item is not a mail message");
        }

        item.Display();
        // Do NOT release item - Outlook manages the displayed window
    }

    public IEnumerable<MailMessageSummary> SearchMail(string? query, string? from, DateTime? after, DateTime? before, string? folderName)
    {
        var folder = GetFolder(folderName);
        var items = folder.Items;
        Track(items);

        var conditions = new List<string>();

        if (!string.IsNullOrEmpty(query))
        {
            conditions.Add($"@SQL=(\"urn:schemas:httpmail:subject\" LIKE '%{EscapeSearchString(query)}%' OR \"urn:schemas:httpmail:textdescription\" LIKE '%{EscapeSearchString(query)}%')");
        }

        if (!string.IsNullOrEmpty(from))
        {
            conditions.Add($"@SQL=\"urn:schemas:httpmail:fromemail\" LIKE '%{EscapeSearchString(from)}%'");
        }

        if (after.HasValue)
        {
            conditions.Add($"[ReceivedTime] >= '{after.Value:g}'");
        }

        if (before.HasValue)
        {
            conditions.Add($"[ReceivedTime] <= '{before.Value:g}'");
        }

        if (conditions.Count > 0)
        {
            var filter = string.Join(" AND ", conditions);
            items = items.Restrict(filter);
            Track(items);
        }

        items.Sort("[ReceivedTime]", true);

        var result = new List<MailMessageSummary>();
        string folderNameStr = folder.Name;
        int totalItems = items.Count;

        for (int i = 1; i <= totalItems; i++)
        {
            dynamic? item = null;
            try
            {
                item = items[i];
                int itemClass = item.Class;
                if (itemClass == 43)
                {
                    result.Add(MapToSummary(item, folderNameStr));
                }
            }
            catch
            {
                // Skip items that can't be read
            }
            finally
            {
                Release(item);
            }
        }

        return result;
    }

    public IEnumerable<MailMessageSummary> GetConversation(string entryId, int limit)
    {
        var (items, folderNames) = GetConversationItems(entryId);
        var result = new List<MailMessageSummary>();
        foreach (var (item, folder) in items.Zip(folderNames))
        {
            try
            {
                result.Add(MapToSummary(item, folder));
            }
            finally
            {
                Release(item);
            }
            if (result.Count >= limit) break;
        }
        return result;
    }

    public IEnumerable<MailMessage> GetConversationFull(string entryId, int limit)
    {
        var (items, folderNames) = GetConversationItems(entryId);
        var result = new List<MailMessage>();
        foreach (var (item, folder) in items.Zip(folderNames))
        {
            try
            {
                result.Add(MapToFull(item, folder));
            }
            finally
            {
                Release(item);
            }
            if (result.Count >= limit) break;
        }
        return result;
    }

    private (List<dynamic> items, List<string> folderNames) GetConversationItems(string entryId)
    {
        if (_namespace == null) throw new InvalidOperationException("Service not initialized");

        var sourceItem = _namespace.GetItemFromID(entryId);
        Track(sourceItem);

        string conversationTopic = sourceItem.ConversationTopic;
        Release(sourceItem);

        if (string.IsNullOrEmpty(conversationTopic))
            return (new List<dynamic>(), new List<string>());

        var escapedTopic = conversationTopic.Replace("'", "''");
        var filter = $"[ConversationTopic] = '{escapedTopic}'";

        var seen = new HashSet<string>();
        var allItems = new List<(dynamic item, string folder, DateTime received)>();

        int[] folderIds = { olFolderInbox, olFolderSentMail, olFolderDrafts };
        foreach (var folderId in folderIds)
        {
            dynamic? folder = null;
            try
            {
                folder = _namespace.GetDefaultFolder(folderId);
                var items = folder.Items;
                var filtered = items.Restrict(filter);
                string folderName = folder.Name;
                int count = filtered.Count;

                for (int i = 1; i <= count; i++)
                {
                    dynamic? item = null;
                    try
                    {
                        item = filtered[i];
                        int itemClass = item.Class;
                        if (itemClass == 43)
                        {
                            string eid = item.EntryID;
                            if (seen.Add(eid))
                            {
                                DateTime received = item.ReceivedTime;
                                allItems.Add((item, folderName, received));
                                item = null; // prevent Release below
                            }
                        }
                    }
                    finally
                    {
                        Release(item);
                    }
                }
                Release(filtered);
                Release(items);
            }
            catch { }
            finally
            {
                Release(folder);
            }
        }

        // Sort chronologically (oldest first)
        allItems.Sort((a, b) => a.received.CompareTo(b.received));

        return (
            allItems.Select(x => x.item).ToList(),
            allItems.Select(x => x.folder).ToList()
        );
    }

    private static string EscapeSearchString(string input)
    {
        return input.Replace("'", "''");
    }

    public void SendMail(string[] to, string[]? cc, string subject, string body, bool isHtml = false, string[]? attachments = null)
    {
        if (_app == null) throw new InvalidOperationException("Service not initialized");

        var mail = _app.CreateItem(olMailItem);
        Track(mail);

        mail.To = string.Join(";", to);
        if (cc != null && cc.Length > 0)
            mail.CC = string.Join(";", cc);
        mail.Subject = subject;

        if (isHtml)
        {
            mail.HTMLBody = body;
        }
        else
        {
            mail.Body = body;
        }

        // Add attachments
        if (attachments != null)
        {
            foreach (var attachmentPath in attachments)
            {
                mail.Attachments.Add(attachmentPath);
            }
        }

        mail.Send();
    }

    public string CreateDraft(string[] to, string[]? cc, string subject, string body, bool isHtml = false, string[]? attachments = null)
    {
        if (_app == null) throw new InvalidOperationException("Service not initialized");

        var mail = _app.CreateItem(olMailItem);
        Track(mail);

        mail.To = string.Join(";", to);
        if (cc != null && cc.Length > 0)
            mail.CC = string.Join(";", cc);
        mail.Subject = subject;

        if (isHtml)
        {
            mail.HTMLBody = body;
        }
        else
        {
            mail.Body = body;
        }

        // Add attachments
        if (attachments != null)
        {
            foreach (var attachmentPath in attachments)
            {
                mail.Attachments.Add(attachmentPath);
            }
        }

        mail.Save(); // Save as draft, don't send
        return mail.EntryID;
    }

    public bool MarkAsRead(string entryId, bool read)
    {
        if (_namespace == null) throw new InvalidOperationException("Service not initialized");

        try
        {
            var item = _namespace.GetItemFromID(entryId);
            Track(item);

            int itemClass = item.Class;
            if (itemClass == 43) // olMail
            {
                item.UnRead = !read;
                item.Save();
                return true;
            }
        }
        catch
        {
            return false;
        }

        return false;
    }

    public List<string> GetCategories(string entryId)
    {
        if (_namespace == null) throw new InvalidOperationException("Service not initialized");

        var item = _namespace.GetItemFromID(entryId);
        Track(item);

        int itemClass = item.Class;
        if (itemClass != 43)
            throw new InvalidOperationException("Item is not a mail message");

        string? cats = null;
        try { cats = item.Categories; } catch { }
        return ParseCategories(cats);
    }

    public bool SetCategories(string entryId, string categories)
    {
        if (_namespace == null) throw new InvalidOperationException("Service not initialized");

        try
        {
            var item = _namespace.GetItemFromID(entryId);
            Track(item);

            int itemClass = item.Class;
            if (itemClass == 43)
            {
                item.Categories = categories;
                item.Save();
                return true;
            }
        }
        catch
        {
            return false;
        }

        return false;
    }

    public List<string>? SaveAttachments(string entryId, string outputDirectory)
    {
        if (_namespace == null) throw new InvalidOperationException("Service not initialized");

        try
        {
            var item = _namespace.GetItemFromID(entryId);
            Track(item);

            int itemClass = item.Class;
            if (itemClass != 43) return null; // Not a mail item

            var savedFiles = new List<string>();
            var atts = item.Attachments;
            int attCount = atts.Count;
            for (int i = 1; i <= attCount; i++)
            {
                var attachment = atts[i];
                try
                {
                    string fileName = attachment.FileName;
                    string filePath = Path.Combine(outputDirectory, fileName);

                    // Handle duplicate file names
                    int counter = 1;
                    while (File.Exists(filePath))
                    {
                        var nameWithoutExt = Path.GetFileNameWithoutExtension(fileName);
                        var ext = Path.GetExtension(fileName);
                        filePath = Path.Combine(outputDirectory, $"{nameWithoutExt}_{counter}{ext}");
                        counter++;
                    }

                    attachment.SaveAsFile(filePath);
                    savedFiles.Add(filePath);
                }
                finally
                {
                    Release(attachment);
                }
            }
            Release(atts);

            return savedFiles;
        }
        catch
        {
            return null;
        }
    }

    public string? ExtractSignatureWithImages(string entryId)
    {
        if (_namespace == null) throw new InvalidOperationException("Service not initialized");

        try
        {
            var item = _namespace.GetItemFromID(entryId);
            Track(item);

            int itemClass = item.Class;
            if (itemClass != 43) return null; // Not a mail item

            string htmlBody = item.HTMLBody ?? "";
            if (string.IsNullOrEmpty(htmlBody)) return null;

            // Build a map of Content-ID to base64 data URI
            var cidToDataUri = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            var attachments = item.Attachments;
            int attCount = attachments.Count;
            for (int i = 1; i <= attCount; i++)
            {
                var attachment = attachments[i];
                try
                {
                    // Get the Content-ID (PropertyAccessor)
                    string? contentId = null;
                    try
                    {
                        // PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
                        contentId = attachment.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F");
                    }
                    catch
                    {
                        // Try alternate method - sometimes CID is in the filename pattern
                    }

                    if (!string.IsNullOrEmpty(contentId))
                    {
                        // Save attachment to temp file to get bytes
                        string tempPath = Path.Combine(Path.GetTempPath(), $"outlook_att_{Guid.NewGuid()}{Path.GetExtension(attachment.FileName)}");
                        try
                        {
                            attachment.SaveAsFile(tempPath);

                            // Wait for file to be fully written and handle
                            byte[]? bytes = null;
                            for (int retry = 0; retry < 20; retry++)
                            {
                                try
                                {
                                    // Use FileShare.Read to allow reading even if Outlook still has a handle
                                    using var fs = new FileStream(tempPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                                    bytes = new byte[fs.Length];
                                    int totalRead = 0;
                                    while (totalRead < bytes.Length)
                                    {
                                        int read = fs.Read(bytes, totalRead, bytes.Length - totalRead);
                                        if (read == 0) break;
                                        totalRead += read;
                                    }
                                    if (totalRead == bytes.Length) break;
                                }
                                catch (IOException)
                                {
                                    Thread.Sleep(100);
                                }
                            }

                            if (bytes != null && bytes.Length > 0)
                            {
                                // Determine MIME type
                                string mimeType = GetMimeType(attachment.FileName);
                                string base64 = Convert.ToBase64String(bytes);
                                string dataUri = $"data:{mimeType};base64,{base64}";

                                cidToDataUri[contentId] = dataUri;
                            }
                        }
                        finally
                        {
                            try { File.Delete(tempPath); } catch { }
                        }
                    }
                }
                catch
                {
                    // Skip attachments we can't process
                }
                finally
                {
                    Release(attachment);
                }
            }
            Release(attachments);

            // Extract signature block
            string? signature = ExtractSignatureBlock(htmlBody);
            if (string.IsNullOrEmpty(signature)) return null;

            // Replace cid: references with data URIs
            foreach (var kvp in cidToDataUri)
            {
                signature = signature.Replace($"cid:{kvp.Key}", kvp.Value, StringComparison.OrdinalIgnoreCase);
            }

            return signature;
        }
        catch
        {
            return null;
        }
    }

    private static string? ExtractSignatureBlock(string htmlBody)
    {
        // Try to find signature by id="Signature"
        var signatureMatch = System.Text.RegularExpressions.Regex.Match(htmlBody,
            @"<div[^>]*id=""Signature""[^>]*>(.*?)</div>\s*</div>\s*</body>",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Singleline);

        if (signatureMatch.Success)
        {
            return $"<div id=\"Signature\">{signatureMatch.Groups[1].Value}</div>";
        }

        // Try broader match for Signature div
        signatureMatch = System.Text.RegularExpressions.Regex.Match(htmlBody,
            @"(<div[^>]*id=""Signature""[^>]*>.*?</div>)\s*</body>",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Singleline);

        if (signatureMatch.Success)
        {
            return signatureMatch.Groups[1].Value;
        }

        // Try to find signature by common greeting patterns
        var greetingPatterns = new[]
        {
            @"(<p[^>]*>.*?Viele Grüße.*?)</body>",
            @"(<p[^>]*>.*?Kind regards.*?)</body>",
            @"(<p[^>]*>.*?Best regards.*?)</body>",
            @"(<p[^>]*>.*?Mit freundlichen Grüßen.*?)</body>",
        };

        foreach (var pattern in greetingPatterns)
        {
            var match = System.Text.RegularExpressions.Regex.Match(htmlBody, pattern,
                System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Singleline);
            if (match.Success)
            {
                return match.Groups[1].Value;
            }
        }

        return null;
    }

    private static string GetMimeType(string fileName)
    {
        var ext = Path.GetExtension(fileName).ToLowerInvariant();
        return ext switch
        {
            ".png" => "image/png",
            ".jpg" or ".jpeg" => "image/jpeg",
            ".gif" => "image/gif",
            ".bmp" => "image/bmp",
            ".webp" => "image/webp",
            ".svg" => "image/svg+xml",
            ".ico" => "image/x-icon",
            _ => "application/octet-stream"
        };
    }

    public string ReplyToMail(string entryId, string body, bool replyAll, bool saveAsDraft = false, bool isHtml = true)
    {
        if (_namespace == null) throw new InvalidOperationException("Service not initialized");

        var item = _namespace.GetItemFromID(entryId);
        Track(item);

        int itemClass = item.Class;
        if (itemClass == 43) // olMail
        {
            var reply = replyAll ? item.ReplyAll() : item.Reply();
            Track(reply);
            if (isHtml)
            {
                string existingHtml = reply.HTMLBody ?? "";
                // Insert the new body after <body...> tag
                var bodyTagMatch = System.Text.RegularExpressions.Regex.Match(existingHtml, @"<body[^>]*>", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                if (bodyTagMatch.Success)
                {
                    int insertPos = bodyTagMatch.Index + bodyTagMatch.Length;
                    reply.HTMLBody = existingHtml.Insert(insertPos, $"<div>{body}</div><br>");
                }
                else
                {
                    reply.HTMLBody = $"<div>{body}</div><br>" + existingHtml;
                }
            }
            else
            {
                reply.Body = body + "\n\n" + reply.Body;
            }
            if (saveAsDraft)
            {
                reply.Save();
                return reply.EntryID;
            }
            else
            {
                reply.Send();
                return string.Empty;
            }
        }
        else
        {
            throw new InvalidOperationException("Item is not a mail message");
        }
    }

    public string ForwardMail(string entryId, string[] to, string? body, bool saveAsDraft = false, bool isHtml = true)
    {
        if (_namespace == null) throw new InvalidOperationException("Service not initialized");

        var item = _namespace.GetItemFromID(entryId);
        Track(item);

        int itemClass = item.Class;
        if (itemClass == 43) // olMail
        {
            var forward = item.Forward();
            Track(forward);
            forward.To = string.Join(";", to);
            if (!string.IsNullOrEmpty(body))
            {
                if (isHtml)
                {
                    string existingHtml = forward.HTMLBody ?? "";
                    var bodyTagMatch = System.Text.RegularExpressions.Regex.Match(existingHtml, @"<body[^>]*>", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                    if (bodyTagMatch.Success)
                    {
                        int insertPos = bodyTagMatch.Index + bodyTagMatch.Length;
                        forward.HTMLBody = existingHtml.Insert(insertPos, $"<div>{body}</div><br>");
                    }
                    else
                    {
                        forward.HTMLBody = $"<div>{body}</div><br>" + existingHtml;
                    }
                }
                else
                {
                    forward.Body = body + "\n\n" + forward.Body;
                }
            }
            if (saveAsDraft)
            {
                forward.Save();
                return forward.EntryID;
            }
            else
            {
                forward.Send();
                return string.Empty;
            }
        }
        else
        {
            throw new InvalidOperationException("Item is not a mail message");
        }
    }

    public bool IsInDeletedItems(string entryId)
    {
        if (_namespace == null) throw new InvalidOperationException("Service not initialized");

        try
        {
            var item = _namespace.GetItemFromID(entryId);
            Track(item);

            int itemClass = item.Class;
            if (itemClass == 43)
            {
                var folder = item.Parent;
                Track(folder);
                var deletedItems = GetDeletedItemsFolder();
                string folderEntryId = folder.EntryID;
                string deletedEntryId = deletedItems.EntryID;
                return folderEntryId == deletedEntryId;
            }
        }
        catch
        {
            return false;
        }

        return false;
    }

    public bool DeleteMail(string entryId)
    {
        if (_namespace == null) throw new InvalidOperationException("Service not initialized");

        try
        {
            var item = _namespace.GetItemFromID(entryId);
            Track(item);

            int itemClass = item.Class;
            if (itemClass == 43)
            {
                item.Delete();
                return true;
            }
        }
        catch
        {
            return false;
        }

        return false;
    }

    public void MoveMail(string entryId, string targetFolderName)
    {
        if (_namespace == null) throw new InvalidOperationException("Service not initialized");

        var item = _namespace.GetItemFromID(entryId);
        Track(item);

        int itemClass = item.Class;
        if (itemClass == 43)
        {
            var targetFolder = GetFolder(targetFolderName);
            item.Move(targetFolder);
        }
        else
        {
            throw new InvalidOperationException("Item is not a mail message");
        }
    }

    public IEnumerable<CalendarEventSummary> GetEventList(DateTime? start, DateTime? end, int limit)
    {
        if (_namespace == null) throw new InvalidOperationException("Service not initialized");

        var folder = _namespace.GetDefaultFolder(olFolderCalendar);
        Track(folder);

        var items = folder.Items;
        Track(items);

        items.IncludeRecurrences = true;
        items.Sort("[Start]");

        var effectiveStart = start ?? DateTime.Today;
        var effectiveEnd = end ?? DateTime.Today.AddMonths(1);

        var filter = $"[Start] >= '{effectiveStart:g}' AND [Start] <= '{effectiveEnd:g}'";
        items = items.Restrict(filter);
        Track(items);

        var result = new List<CalendarEventSummary>();
        int count = 0;
        int totalItems = items.Count;

        for (int i = 1; i <= totalItems && count < limit; i++)
        {
            dynamic? item = null;
            try
            {
                item = items[i];
                int itemClass = item.Class;
                if (itemClass == 26) // olAppointment
                {
                    result.Add(MapEventToSummary(item));
                    count++;
                }
            }
            catch
            {
                // Skip items that can't be read
            }
            finally
            {
                Release(item);
            }
        }

        return result;
    }

    public IEnumerable<CalendarEvent> GetEventListFull(DateTime? start, DateTime? end, int limit)
    {
        if (_namespace == null) throw new InvalidOperationException("Service not initialized");

        var folder = _namespace.GetDefaultFolder(olFolderCalendar);
        Track(folder);

        var items = folder.Items;
        Track(items);

        items.IncludeRecurrences = true;
        items.Sort("[Start]");

        var effectiveStart = start ?? DateTime.Today;
        var effectiveEnd = end ?? DateTime.Today.AddMonths(1);

        var filter = $"[Start] >= '{effectiveStart:g}' AND [Start] <= '{effectiveEnd:g}'";
        items = items.Restrict(filter);
        Track(items);

        var result = new List<CalendarEvent>();
        int count = 0;
        int totalItems = items.Count;

        for (int i = 1; i <= totalItems && count < limit; i++)
        {
            dynamic? item = null;
            try
            {
                item = items[i];
                int itemClass = item.Class;
                if (itemClass == 26) // olAppointment
                {
                    result.Add(MapEventToFull(item));
                    count++;
                }
            }
            catch
            {
                // Skip items that can't be read
            }
            finally
            {
                Release(item);
            }
        }

        return result;
    }

    public CalendarEvent? GetEvent(string entryId)
    {
        if (_namespace == null) throw new InvalidOperationException("Service not initialized");

        try
        {
            var item = _namespace.GetItemFromID(entryId);
            Track(item);

            int itemClass = item.Class;
            if (itemClass == 26) // olAppointment
            {
                return MapEventToFull(item);
            }
        }
        catch
        {
            return null;
        }

        return null;
    }

    public void OpenEvent(string entryId)
    {
        if (_namespace == null) throw new InvalidOperationException("Service not initialized");

        var item = _namespace.GetItemFromID(entryId);
        int itemClass = item.Class;
        if (itemClass != 26) // olAppointment
        {
            Marshal.ReleaseComObject(item);
            throw new InvalidOperationException("Item is not a calendar event");
        }

        item.Display();
        // Do NOT release item - Outlook manages the displayed window
    }

    public string CreateEvent(string subject, DateTime start, DateTime end, string? location, string? body, bool isAllDay)
    {
        if (_app == null) throw new InvalidOperationException("Service not initialized");

        var apt = _app.CreateItem(olAppointmentItem);
        Track(apt);

        apt.Subject = subject;
        apt.Start = start;
        apt.End = end;
        apt.AllDayEvent = isAllDay;

        if (!string.IsNullOrEmpty(location))
            apt.Location = location;
        if (!string.IsNullOrEmpty(body))
            apt.Body = body;

        apt.Save();
        return apt.EntryID;
    }

    public void UpdateEvent(string entryId, string? subject, DateTime? start, DateTime? end, string? location, string? body)
    {
        if (_namespace == null) throw new InvalidOperationException("Service not initialized");

        var item = _namespace.GetItemFromID(entryId);
        Track(item);

        int itemClass = item.Class;
        if (itemClass == 26) // olAppointment
        {
            if (!string.IsNullOrEmpty(subject))
                item.Subject = subject;
            if (start.HasValue)
                item.Start = start.Value;
            if (end.HasValue)
                item.End = end.Value;
            if (location != null)
                item.Location = location;
            if (body != null)
                item.Body = body;

            item.Save();
        }
        else
        {
            throw new InvalidOperationException("Item is not a calendar event");
        }
    }

    public bool DeleteEvent(string entryId)
    {
        if (_namespace == null) throw new InvalidOperationException("Service not initialized");

        try
        {
            var item = _namespace.GetItemFromID(entryId);
            Track(item);

            int itemClass = item.Class;
            if (itemClass == 26) // olAppointment
            {
                item.Delete();
                return true;
            }
        }
        catch
        {
            return false;
        }

        return false;
    }

    public bool RespondToMeeting(string entryId, string responseType, string? message)
    {
        if (_namespace == null) throw new InvalidOperationException("Service not initialized");

        try
        {
            var item = _namespace.GetItemFromID(entryId);
            Track(item);

            int itemClass = item.Class;
            if (itemClass == 26) // olAppointment
            {
                dynamic response;
                switch (responseType.ToLowerInvariant())
                {
                    case "accept":
                        response = item.Respond(3); // olMeetingAccepted = 3
                        break;
                    case "decline":
                        response = item.Respond(4); // olMeetingDeclined = 4
                        break;
                    case "tentative":
                        response = item.Respond(2); // olMeetingTentative = 2
                        break;
                    default:
                        return false;
                }

                if (response != null)
                {
                    Track(response);
                    if (!string.IsNullOrEmpty(message))
                    {
                        response.Body = message + "\n\n" + response.Body;
                    }
                    response.Send();
                }

                return true;
            }
        }
        catch
        {
            return false;
        }

        return false;
    }

    public FreeBusyResult GetFreeBusy(string email, DateTime start, DateTime end)
    {
        if (_namespace == null) throw new InvalidOperationException("Service not initialized");

        var recipient = _namespace.CreateRecipient(email);
        Track(recipient);
        recipient.Resolve();

        if (!(bool)recipient.Resolved)
            throw new InvalidOperationException($"Could not resolve recipient: {email}");

        // FreeBusy(Start, MinPerChar, CompleteFormat)
        // Returns string: '0'=free, '1'=tentative, '2'=busy, '3'=OOF
        const int intervalMinutes = 30;
        string freeBusyString = recipient.FreeBusy(start, intervalMinutes, true);

        // Calculate how many characters we need for our range
        int totalMinutes = (int)(end - start).TotalMinutes;
        int charsNeeded = totalMinutes / intervalMinutes;
        if (charsNeeded > freeBusyString.Length)
            charsNeeded = freeBusyString.Length;

        var slots = new List<FreeBusySlot>();
        string? currentStatus = null;
        DateTime? slotStart = null;

        for (int i = 0; i < charsNeeded; i++)
        {
            string status = freeBusyString[i] switch
            {
                '1' => "Tentative",
                '2' => "Busy",
                '3' => "OOF",
                _ => "Free"
            };

            if (status != currentStatus)
            {
                if (currentStatus != null && currentStatus != "Free" && slotStart.HasValue)
                {
                    slots.Add(new FreeBusySlot(slotStart.Value, start.AddMinutes(i * intervalMinutes), currentStatus));
                }
                currentStatus = status;
                slotStart = start.AddMinutes(i * intervalMinutes);
            }
        }

        // Close last slot
        if (currentStatus != null && currentStatus != "Free" && slotStart.HasValue)
        {
            slots.Add(new FreeBusySlot(slotStart.Value, start.AddMinutes(charsNeeded * intervalMinutes), currentStatus));
        }

        return new FreeBusyResult(email, start, end, slots);
    }

    public List<AvailableSlot> FindAvailableSlots(List<string> emails, DateTime start, DateTime end, int durationMinutes, bool includeSelf = true)
    {
        if (_namespace == null) throw new InvalidOperationException("Service not initialized");

        var allEmails = new List<string>(emails);

        // Add self if requested
        if (includeSelf)
        {
            string? selfEmail = null;
            try
            {
                var currentUser = _namespace.CurrentUser;
                Track(currentUser);
                selfEmail = currentUser.Address;
            }
            catch { }

            if (!string.IsNullOrEmpty(selfEmail) && !allEmails.Contains(selfEmail, StringComparer.OrdinalIgnoreCase))
            {
                allEmails.Add(selfEmail);
            }
        }

        // Get free/busy for all attendees
        var freeBusyResults = new List<FreeBusyResult>();
        foreach (var email in allEmails)
        {
            freeBusyResults.Add(GetFreeBusy(email, start, end));
        }

        // Find slots where everyone is free, filtered to business hours (09:00-17:00) on weekdays
        var availableSlots = new List<AvailableSlot>();
        const int intervalMinutes = 30;
        var current = start;

        while (current.AddMinutes(durationMinutes) <= end)
        {
            // Skip weekends
            if (current.DayOfWeek == DayOfWeek.Saturday || current.DayOfWeek == DayOfWeek.Sunday)
            {
                current = current.Date.AddDays(1).AddHours(9);
                continue;
            }

            // Skip outside business hours
            if (current.Hour < 9)
            {
                current = current.Date.AddHours(9);
                continue;
            }
            if (current.Hour >= 17 || current.AddMinutes(durationMinutes).Hour > 17 ||
                (current.AddMinutes(durationMinutes).Hour == 17 && current.AddMinutes(durationMinutes).Minute > 0 && current.AddMinutes(durationMinutes).Date == current.Date))
            {
                current = current.Date.AddDays(1).AddHours(9);
                continue;
            }

            // Check if slot end exceeds 17:00
            var slotEnd = current.AddMinutes(durationMinutes);
            if (slotEnd.TimeOfDay > new TimeSpan(17, 0, 0) && slotEnd.Date == current.Date)
            {
                current = current.Date.AddDays(1).AddHours(9);
                continue;
            }

            // Check all attendees are free for the entire duration
            bool allFree = true;
            foreach (var fb in freeBusyResults)
            {
                foreach (var busySlot in fb.BusySlots)
                {
                    // Check overlap: busy slot overlaps with [current, slotEnd)
                    if (busySlot.Start < slotEnd && busySlot.End > current)
                    {
                        allFree = false;
                        break;
                    }
                }
                if (!allFree) break;
            }

            if (allFree)
            {
                availableSlots.Add(new AvailableSlot(current, slotEnd, durationMinutes, allEmails));
                current = slotEnd; // Skip past this slot to avoid overlapping results
            }
            else
            {
                current = current.AddMinutes(intervalMinutes);
            }
        }

        return availableSlots;
    }

    private static List<string> ParseCategories(string? categories)
    {
        if (string.IsNullOrWhiteSpace(categories))
            return new List<string>();
        return categories.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).ToList();
    }

    private MailMessageSummary MapToSummary(dynamic mail, string folderName)
    {
        string senderEmail;
        try
        {
            senderEmail = mail.SenderEmailAddress ?? "";
        }
        catch
        {
            senderEmail = "";
        }

        var atts = mail.Attachments;
        int attCount = atts.Count;
        Release(atts);

        string? cats = null;
        try { cats = mail.Categories; } catch { }

        string? conversationTopic = null;
        try { conversationTopic = mail.ConversationTopic; } catch { }

        return new MailMessageSummary(
            mail.EntryID,
            mail.Subject ?? "",
            mail.SenderName ?? "",
            senderEmail,
            mail.ReceivedTime,
            mail.UnRead,
            folderName,
            attCount > 0,
            attCount,
            ParseCategories(cats),
            conversationTopic
        );
    }

    private MailMessage MapToFull(dynamic mail, string folderName)
    {
        string senderEmail;
        try
        {
            senderEmail = mail.SenderEmailAddress ?? "";
        }
        catch
        {
            senderEmail = "";
        }

        var toList = new List<string>();
        var ccList = new List<string>();

        var recipients = mail.Recipients;
        int recipientCount = recipients.Count;
        for (int i = 1; i <= recipientCount; i++)
        {
            var recipient = recipients[i];
            try
            {
                int recipientType = recipient.Type;
                string address = recipient.Address ?? recipient.Name;
                if (recipientType == olTo)
                    toList.Add(address);
                else if (recipientType == olCC)
                    ccList.Add(address);
            }
            finally
            {
                Release(recipient);
            }
        }
        Release(recipients);

        var attachmentList = new List<Attachment>();
        var atts = mail.Attachments;
        int attCount = atts.Count;
        for (int i = 1; i <= attCount; i++)
        {
            var att = atts[i];
            try
            {
                attachmentList.Add(new Attachment(
                    att.FileName,
                    att.Size,
                    att.Type.ToString()
                ));
            }
            finally
            {
                Release(att);
            }
        }
        Release(atts);

        string? cats = null;
        try { cats = mail.Categories; } catch { }

        string? conversationTopic = null;
        try { conversationTopic = mail.ConversationTopic; } catch { }

        return new MailMessage(
            mail.EntryID,
            mail.Subject ?? "",
            mail.SenderName ?? "",
            senderEmail,
            mail.ReceivedTime,
            mail.UnRead,
            folderName,
            attCount > 0,
            attCount,
            mail.Body ?? "",
            mail.HTMLBody ?? "",
            toList,
            ccList,
            attachmentList,
            ParseCategories(cats),
            conversationTopic
        );
    }

    private CalendarEventSummary MapEventToSummary(dynamic apt)
    {
        return new CalendarEventSummary(
            apt.EntryID,
            apt.Subject ?? "",
            apt.Start,
            apt.End,
            apt.Location ?? "",
            apt.AllDayEvent,
            apt.IsRecurring
        );
    }

    private CalendarEvent MapEventToFull(dynamic apt)
    {
        var attendees = new List<string>();
        var recipients = apt.Recipients;
        int recipientCount = recipients.Count;
        for (int i = 1; i <= recipientCount; i++)
        {
            var recipient = recipients[i];
            try
            {
                string address = recipient.Address ?? recipient.Name;
                attendees.Add(address);
            }
            finally
            {
                Release(recipient);
            }
        }
        Release(recipients);

        string? recurrencePattern = null;
        bool isRecurring = apt.IsRecurring;
        if (isRecurring)
        {
            try
            {
                var pattern = apt.GetRecurrencePattern();
                Track(pattern);
                recurrencePattern = pattern.RecurrenceType.ToString();
            }
            catch
            {
                recurrencePattern = "Unknown";
            }
        }

        return new CalendarEvent(
            apt.EntryID,
            apt.Subject ?? "",
            apt.Start,
            apt.End,
            apt.Location ?? "",
            apt.AllDayEvent,
            isRecurring,
            apt.Body ?? "",
            attendees,
            apt.Organizer ?? "",
            recurrencePattern
        );
    }
}
