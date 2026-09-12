using System.Drawing.Imaging;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Auth.OAuth2.Responses;
using Google.Apis.PeopleService.v1;
using Google.Apis.PeopleService.v1.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;

namespace Adressen.cls;

internal record LoadContactsResult(List<Contact> Contacts, Dictionary<string, string> GroupMap);

internal class GooglePeopleManager(string secretPath, string tokenDir)
{
    private static volatile PeopleServiceService? _cachedService;  // ohne volatile auf Multi-Core-Systemen nicht korrekt

    internal const string TokenFileName = "Google.Apis.Auth.OAuth2.Responses.TokenResponse-user";

    // true, wenn bei der Autorisierung eine Browser-Anmeldung stattgefunden hat (statt stillem Token-Laden).
    internal bool LastAuthorizationWasInteractive { get; private set; }
    private static readonly SemaphoreSlim _serviceLock = new(1, 1);
    // Fax-Typen werden bei der Verteilung übrig gebliebener Nummern auf die Telefonfelder ausgeschlossen (siehe MapPhoneSlots)
    private static readonly HashSet<string> _faxTypes = new(StringComparer.OrdinalIgnoreCase) { TYPE_FAX, TYPE_HOME_FAX, TYPE_WORK_FAX, TYPE_OTHER_FAX };

    // Zuordnung Programmfeld → konkreter Google-Eintrag. Wird beim Lesen (MapPersonToContact) und beim Schreiben (UpdateContactAsync)
    // mit derselben Regel berechnet, damit garantiert der angezeigte Eintrag geändert bzw. entfernt wird und kein anderer.
    private sealed record PhoneSlots(PhoneNumber? Telefon1, PhoneNumber? Telefon2, PhoneNumber? Mobil, PhoneNumber? Fax);
    private sealed record EmailSlots(EmailAddress? Mail1, EmailAddress? Mail2);
    internal sealed record ContactFieldLabels(string? Telefon1, string? Telefon2, string? Mobil, string? Fax, string? Mail1, string? Mail2);

    // --- GOOGLE API KONSTANTEN ---
    private const string CONTACT_PERSON_FIELDS = "names,memberships,nicknames,addresses,phoneNumbers,emailAddresses,biographies,birthdays,urls,organizations,photos,userDefined,metadata";
    private const string GROUP_FIELDS = "name,clientData,groupType";

    // System-Gruppen
    private const string GROUP_STARRED_RESOURCE = "contactGroups/starred";
    private const string GROUP_STARRED_LABEL = "starred";
    private const string GROUP_MYCONTACTS_RESOURCE = "contactGroups/myContacts";
    private const string GROUP_MYCONTACTS_LABEL = "myContacts";
    private const string STAR_SYMBOL = "★";

    // Standard-Typen für Kontakte
    private const string TYPE_HOME = "home";
    private const string TYPE_WORK = "work";
    private const string TYPE_MOBILE = "mobile";
    private const string TYPE_FAX = "fax";
    private const string TYPE_HOMEPAGE = "homePage";
    private const string TYPE_OTHER = "other";
    private const string TYPE_HOME_FAX = "homeFax";
    private const string TYPE_WORK_FAX = "workFax";
    private const string TYPE_OTHER_FAX = "otherFax";
    private const string TYPE_WORK_MOBILE = "workMobile";
    private const string TYPE_MAIN = "main";  // Hauptnummer, wird von Google als Fallback für fehlende Typen verwendet
    private const string TYPE_PAGER = "pager";  // seltener Typ, den wir als letzten Fallback für Mobilnummern verwenden, da Pager in der Regel keine Handynummern sind

    // UserDefined Keys (benutzerdefinierte Felder)
    private const string KEY_ANREDE = "Anrede";
    private const string KEY_BETREFF = "Betreff";
    private const string KEY_GRUSS = "Grussformel";
    private const string KEY_SCHLUSS = "Schlussformel";

    // ========================================================================
    // 1. PUBLIC API: LOAD, CREATE, UPDATE, DELETE
    // ========================================================================

    public async Task<LoadContactsResult> LoadContactsAsync(CancellationToken token = default)
    {
        try
        {
            var service = await GetServiceAsync(token);
            var groupMap = await GetContactGroupsMapAsync(service, token);
            var peopleRequest = service.People.Connections.List("people/me");
            peopleRequest.PersonFields = CONTACT_PERSON_FIELDS;
            peopleRequest.SortOrder = (PeopleResource.ConnectionsResource.ListRequest.SortOrderEnum)3; // LAST_NAME_ASCENDING
            peopleRequest.PageSize = 2000;
            var response = await peopleRequest.ExecuteAsync(token);
            var contactList = new List<Contact>();
            if (response?.Connections != null)
            {
                foreach (var person in response.Connections) { contactList.Add(MapPersonToContact(person, groupMap)); }
            }
            return new LoadContactsResult(contactList, groupMap);
        }
        catch (TokenResponseException ex) { throw new UnauthorizedAccessException("Google Token abgelaufen", ex); }
    }

    public async Task<Contact> CreateContactAsync(Contact contact, Image? profileImage, ImageFormat? photoFormat, CancellationToken token = default)
    {
        contact.TrimStrings();
        var service = await GetServiceAsync(token);
        var personToCreate = new Person
        {
            Names = [new() {
                HonorificPrefix = contact.Praefix ?? "",
                GivenName = contact.Vorname ?? "",
                MiddleName = contact.Zwischenname ?? "",
                FamilyName = contact.Nachname ?? "",
                HonorificSuffix = contact.Suffix ?? ""
            }],
            Nicknames = !string.IsNullOrWhiteSpace(contact.Nickname)
                ? [new() { Value = contact.Nickname }] : null,
            Organizations = [new() {
                Name = contact.Unternehmen ?? "",
                Title = contact.Position ?? "",
                Type = TYPE_WORK
            }],
            Addresses = [new() {
                StreetAddress = contact.Strasse ?? "",
                PostalCode = contact.PLZ ?? "",
                City = contact.Ort ?? "",
                PoBox = contact.Postfach ?? "",
                Country = contact.Land ?? ""
            }],
            Birthdays = contact.Geburtstag.HasValue ? [new() {
                Date = new Date {
                    Day = contact.Geburtstag.Value.Day,
                    Month = contact.Geburtstag.Value.Month,
                    Year = contact.Geburtstag.Value.Year
                }
            }] : null,
            Urls = !string.IsNullOrWhiteSpace(contact.Internet) ? [new() { Value = contact.Internet }] : null,
            Biographies = !string.IsNullOrWhiteSpace(contact.Notizen) ? [new() { Value = contact.Notizen }] : null
        };

        var emails = new List<EmailAddress>();
        if (!string.IsNullOrWhiteSpace(contact.Mail1)) { emails.Add(new EmailAddress { Value = contact.Mail1, Type = TYPE_HOME }); }
        if (!string.IsNullOrWhiteSpace(contact.Mail2)) { emails.Add(new EmailAddress { Value = contact.Mail2, Type = TYPE_WORK }); }
        if (emails.Count > 0) { personToCreate.EmailAddresses = emails; }

        var phones = new List<PhoneNumber>();
        if (!string.IsNullOrWhiteSpace(contact.Telefon1)) { phones.Add(new PhoneNumber { Value = contact.Telefon1, Type = TYPE_HOME }); }
        if (!string.IsNullOrWhiteSpace(contact.Telefon2)) { phones.Add(new PhoneNumber { Value = contact.Telefon2, Type = TYPE_WORK }); }
        if (!string.IsNullOrWhiteSpace(contact.Mobil)) { phones.Add(new PhoneNumber { Value = contact.Mobil, Type = TYPE_MOBILE }); }
        if (!string.IsNullOrWhiteSpace(contact.Fax)) { phones.Add(new PhoneNumber { Value = contact.Fax, Type = TYPE_FAX }); }
        if (phones.Count > 0) { personToCreate.PhoneNumbers = phones; }

        var userDefined = new List<UserDefined>();
        if (!string.IsNullOrWhiteSpace(contact.Anrede)) { userDefined.Add(new UserDefined { Key = KEY_ANREDE, Value = contact.Anrede }); }
        if (!string.IsNullOrWhiteSpace(contact.Betreff)) { userDefined.Add(new UserDefined { Key = KEY_BETREFF, Value = contact.Betreff }); }
        if (!string.IsNullOrWhiteSpace(contact.Grussformel)) { userDefined.Add(new UserDefined { Key = KEY_GRUSS, Value = contact.Grussformel }); }
        if (!string.IsNullOrWhiteSpace(contact.Schlussformel)) { userDefined.Add(new UserDefined { Key = KEY_SCHLUSS, Value = contact.Schlussformel }); }
        if (userDefined.Count > 0) { personToCreate.UserDefined = userDefined; }

        var createdPerson = await service.People.CreateContact(personToCreate).ExecuteAsync(token);

        contact.RawGooglePerson = createdPerson;
        contact.ResourceName = createdPerson.ResourceName;
        contact.ETag = createdPerson.ETag;
        contact.LastModified = createdPerson.Metadata?.Sources?.FirstOrDefault(static s => s.Type == "CONTACT")?.UpdateTimeDateTimeOffset?.UtcDateTime;

        if (profileImage != null && !string.IsNullOrEmpty(contact.ResourceName))
        {
            var format = photoFormat ?? profileImage.RawFormat;  // Fallback falls doch mal kein Format mitgegeben wird
            var (photoUrl, newETag) = await UploadPhotoInternalAsync(service, contact.ResourceName, profileImage, format, token);
            if (!string.IsNullOrEmpty(photoUrl)) { contact.PhotoUrl = photoUrl; }
            if (!string.IsNullOrEmpty(newETag)) { contact.ETag = newETag; }
        }
        return contact;
    }

    public async Task<Contact> UpdateContactAsync(Contact contact, List<string> changedFields, Dictionary<string, string> groupMap, Contact? originalContactSnapshot, bool checkEmptyGroups = false, CancellationToken token = default)
    {
        contact.TrimStrings();
        var service = await GetServiceAsync(token);

        var personToUpdate = contact.RawGooglePerson != null
            ? Newtonsoft.Json.JsonConvert.DeserializeObject<Person>(Newtonsoft.Json.JsonConvert.SerializeObject(contact.RawGooglePerson)) ?? new Person()
            : new Person();
        personToUpdate.Metadata = null;  // beim Schreiben das Metadata-Objekt aus dem Payload entfernen, sonst PreconditionFailed-Konflikt
        personToUpdate.ResourceName = contact.ResourceName;
        personToUpdate.ETag = contact.ETag;

        if (changedFields.Contains("names"))
        {
            personToUpdate.Names ??= [];
            var primaryName = personToUpdate.Names.FirstOrDefault();  // derselbe Eintrag wie beim Lesen in MapPersonToContact (erster Eintrag) – sonst wird ein anderer Eintrag überschrieben als der angezeigte
            if (primaryName == null)
            {
                primaryName = new Name();
                personToUpdate.Names.Add(primaryName);
            }
            primaryName.HonorificPrefix = contact.Praefix;
            primaryName.FamilyName = contact.Nachname;
            primaryName.GivenName = contact.Vorname;
            primaryName.MiddleName = contact.Zwischenname;
            primaryName.HonorificSuffix = contact.Suffix;
        }

        if (changedFields.Contains("nicknames"))
        {
            personToUpdate.Nicknames ??= [];
            var primaryNick = personToUpdate.Nicknames.FirstOrDefault();  // wie beim Lesen: erster Eintrag
            if (primaryNick == null)
            {
                primaryNick = new Nickname();
                personToUpdate.Nicknames.Add(primaryNick);
            }
            primaryNick.Value = contact.Nickname;
        }

        if (changedFields.Contains("addresses"))
        {
            personToUpdate.Addresses ??= [];
            // Zwingend dieselbe Auswahlregel wie beim Lesen (home → work → other → erste): Sonst wird bei einem Kontakt mit
            // z. B. [work (primär), home] die angezeigte Privatadresse in die Geschäftsadresse geschrieben und diese geht verloren.
            var primaryAddr = GetGoogleAddressByType(personToUpdate, TYPE_HOME, TYPE_WORK, TYPE_OTHER);
            if (primaryAddr == null)
            {
                primaryAddr = new Address { Type = TYPE_HOME };  // mit Typ anlegen, damit der nächste Lesevorgang denselben Eintrag findet
                personToUpdate.Addresses.Add(primaryAddr);
            }
            primaryAddr.StreetAddress = contact.Strasse;
            primaryAddr.PostalCode = contact.PLZ;
            primaryAddr.City = contact.Ort;
            primaryAddr.PoBox = contact.Postfach;
            primaryAddr.Country = contact.Land;
        }

        if (changedFields.Contains("organizations"))
        {
            personToUpdate.Organizations ??= [];
            var primaryOrg = personToUpdate.Organizations.FirstOrDefault();  // wie beim Lesen: erster Eintrag
            if (primaryOrg == null)
            {
                primaryOrg = new Organization();
                personToUpdate.Organizations.Add(primaryOrg);
            }
            primaryOrg.Name = contact.Unternehmen;
            primaryOrg.Title = contact.Position;
        }

        if (changedFields.Contains("birthdays"))
        {
            personToUpdate.Birthdays ??= [];
            var primaryBday = personToUpdate.Birthdays.FirstOrDefault();  // wie beim Lesen: erster Eintrag (Birthdays[0])
            if (contact.Geburtstag.HasValue)
            {
                if (primaryBday == null)
                {
                    primaryBday = new Birthday();
                    personToUpdate.Birthdays.Add(primaryBday);
                }
                primaryBday.Date = new Date { Day = contact.Geburtstag.Value.Day, Month = contact.Geburtstag.Value.Month, Year = contact.Geburtstag.Value.Year };
            }
            else
            {
                if (primaryBday != null) { personToUpdate.Birthdays.Remove(primaryBday); }
            }
        }

        if (changedFields.Contains("phoneNumbers")) { ApplyPhoneChanges(personToUpdate, contact); }

        if (changedFields.Contains("emailAddresses")) { ApplyEmailChanges(personToUpdate, contact); }

        if (changedFields.Contains("urls"))
        {
            personToUpdate.Urls ??= [];
            var primaryUrl = GetGoogleUrlEntry(personToUpdate, TYPE_HOMEPAGE, TYPE_HOME, TYPE_WORK, TYPE_OTHER);  // dieselbe Auswahlregel wie beim Lesen
            if (string.IsNullOrWhiteSpace(contact.Internet))
            {
                if (primaryUrl != null) { personToUpdate.Urls.Remove(primaryUrl); }
            }
            else
            {
                if (primaryUrl == null)
                {
                    primaryUrl = new Url { Type = TYPE_HOMEPAGE };
                    personToUpdate.Urls.Add(primaryUrl);
                }
                primaryUrl.Value = contact.Internet;
            }
        }

        if (changedFields.Contains("biographies"))
        {
            personToUpdate.Biographies ??= [];
            var primaryBio = personToUpdate.Biographies.FirstOrDefault();  // wie beim Lesen: erster Eintrag
            if (string.IsNullOrWhiteSpace(contact.Notizen))
            {
                if (primaryBio != null)
                {
                    personToUpdate.Biographies.Remove(primaryBio);
                }
            }
            else
            {
                if (primaryBio == null)
                {
                    primaryBio = new Biography();
                    personToUpdate.Biographies.Add(primaryBio);
                }
                primaryBio.Value = contact.Notizen;
            }
        }

        if (changedFields.Contains("userDefined"))
        {
            UpdateGoogleUserDef(personToUpdate, KEY_ANREDE, contact.Anrede);
            UpdateGoogleUserDef(personToUpdate, KEY_BETREFF, contact.Betreff);
            UpdateGoogleUserDef(personToUpdate, KEY_GRUSS, contact.Grussformel);
            UpdateGoogleUserDef(personToUpdate, KEY_SCHLUSS, contact.Schlussformel);
        }

        var groupsToRemoveToCheck = new HashSet<string>();
        if (changedFields.Contains("memberships"))
        {
            personToUpdate.Memberships ??= [];

            var knownGroupNames = new HashSet<string>(groupMap.Keys);
            knownGroupNames.Add(GROUP_MYCONTACTS_RESOURCE);
            knownGroupNames.Add(GROUP_STARRED_RESOURCE);

            personToUpdate.Memberships = [.. personToUpdate.Memberships
                .Where(m => m.ContactGroupMembership?.ContactGroupResourceName != null
                         && !knownGroupNames.Contains(m.ContactGroupMembership.ContactGroupResourceName))];

            var desiredGroupNames = new HashSet<string>(contact.GroupNames, StringComparer.OrdinalIgnoreCase);
            if (desiredGroupNames.Remove(STAR_SYMBOL)) { desiredGroupNames.Add(GROUP_STARRED_LABEL); }
            desiredGroupNames.Add(GROUP_MYCONTACTS_LABEL);

            foreach (var groupName in desiredGroupNames)
            {
                var resourceName = string.Empty;

                var existingEntry = groupMap.FirstOrDefault(x => x.Value.Equals(groupName, StringComparison.OrdinalIgnoreCase));
                if (!string.IsNullOrEmpty(existingEntry.Key)) { resourceName = existingEntry.Key; }
                else if (groupName.Equals(GROUP_MYCONTACTS_LABEL, StringComparison.OrdinalIgnoreCase)) { resourceName = GROUP_MYCONTACTS_RESOURCE; }
                else if (groupName.Equals(GROUP_STARRED_LABEL, StringComparison.OrdinalIgnoreCase) || groupName == STAR_SYMBOL) { resourceName = GROUP_STARRED_RESOURCE; }
                else
                {
                    resourceName = await CreateContactGroupInternalAsync(service, groupName, token);
                    if (!string.IsNullOrEmpty(resourceName)) { groupMap[resourceName] = groupName; }
                }

                if (!string.IsNullOrEmpty(resourceName))
                {
                    personToUpdate.Memberships.Add(new Membership { ContactGroupMembership = new ContactGroupMembership { ContactGroupResourceName = resourceName } });
                }
            }
            if (checkEmptyGroups && originalContactSnapshot != null)
            {
                var originalGroups = originalContactSnapshot.GroupNames.Select(g => g == STAR_SYMBOL ? GROUP_STARRED_LABEL : g).ToHashSet(StringComparer.OrdinalIgnoreCase);
                foreach (var rem in originalGroups.Except(desiredGroupNames))
                {
                    var resKey = groupMap.FirstOrDefault(x => x.Value.Equals(rem, StringComparison.OrdinalIgnoreCase)).Key;
                    if (!string.IsNullOrEmpty(resKey)) { groupsToRemoveToCheck.Add(resKey); }
                }
            }
        }
        if (changedFields.Count > 0)
        {
            var updateRequest = service.People.UpdateContact(personToUpdate, contact.ResourceName);
            updateRequest.UpdatePersonFields = string.Join(",", changedFields);
            var updatedPerson = await updateRequest.ExecuteAsync(token);
            contact.RawGooglePerson = updatedPerson;
            contact.ETag = updatedPerson.ETag;
            contact.ResourceName = updatedPerson.ResourceName;
            contact.LastModified = updatedPerson.Metadata?.Sources?.FirstOrDefault(static s => s.Type == "CONTACT")?.UpdateTimeDateTimeOffset?.UtcDateTime;
            if (checkEmptyGroups && groupsToRemoveToCheck.Count > 0) { await CheckAndDeleteEmptyGroupsInternalAsync(service, groupsToRemoveToCheck, token); }
        }
        return contact;
    }

    public async Task<Person> GetRawPersonAsync(string resourceName, CancellationToken token = default)
    {
        var service = await GetServiceAsync(token);
        var req = service.People.Get(resourceName);
        req.PersonFields = CONTACT_PERSON_FIELDS;
        return await req.ExecuteAsync(token);
    }

    public async Task DeleteContactAsync(string resourceName, CancellationToken token = default)
    {
        if (string.IsNullOrWhiteSpace(resourceName)) { return; }
        var service = await GetServiceAsync(token);
        await service.People.DeleteContact(resourceName).ExecuteAsync(token);
    }

    public async Task DeleteContactGroupAsync(string resourceName, CancellationToken token = default)
    {
        if (string.IsNullOrWhiteSpace(resourceName)) { return; }
        var service = await GetServiceAsync(token);
        await service.ContactGroups.Delete(resourceName).ExecuteAsync(token);
    }

    public async Task UpdateContactGroupNameAsync(string resourceName, string newName, CancellationToken token = default)
    {
        if (string.IsNullOrWhiteSpace(resourceName) || string.IsNullOrWhiteSpace(newName)) { return; }
        var service = await GetServiceAsync(token);
        var requestBody = new UpdateContactGroupRequest
        {
            ContactGroup = new ContactGroup { Name = newName },
            UpdateGroupFields = "name"
        };
        var request = service.ContactGroups.Update(requestBody, resourceName);
        await request.ExecuteAsync(token);
    }

    // ========================================================================
    // 2. FOTO API
    // ========================================================================

    public async Task<(string? PhotoUrl, string? ETag)> UpdateContactPhotoAsync(
        string resourceName, Image image, ImageFormat format, CancellationToken token = default)
    {
        var service = await GetServiceAsync(token);
        return await UploadPhotoInternalAsync(service, resourceName, image, format, token);
    }

    public async Task<(string? PhotoUrl, string? ETag)> DeleteContactPhotoAsync(string resourceName, CancellationToken token = default)
    {
        try
        {
            var service = await GetServiceAsync(token);
            var request = service.People.DeleteContactPhoto(resourceName);
            request.PersonFields = "photos,metadata";

            var response = await request.ExecuteAsync(token);
            var person = response?.Person;

            return (person?.Photos?.FirstOrDefault()?.Url, person?.ETag);
        }
        catch (Google.GoogleApiException ex) when (ex.HttpStatusCode == System.Net.HttpStatusCode.NotFound)
        {
            // Foto ist serverseitig bereits gelöscht. Wir fordern die Personendaten manuell neu an, um den korrekten ETag zu erhalten.
            var service = await GetServiceAsync(token);
            var personRequest = service.People.Get(resourceName);
            personRequest.PersonFields = "metadata";
            var currentPerson = await personRequest.ExecuteAsync(token);

            return (null, currentPerson?.ETag);
        }
    }

    // ========================================================================
    // 3. INTERNE HILFSMETHODEN (PRIVATE)
    // ========================================================================

    private async Task<PeopleServiceService> GetServiceAsync(CancellationToken token)
    {
        if (_cachedService != null) { return _cachedService; }

        await _serviceLock.WaitAsync(token);
        try
        {
            if (_cachedService != null) { return _cachedService; }

            var scopes = new[] { PeopleServiceService.Scope.Contacts };
            using var stream = new FileStream(secretPath, FileMode.Open, FileAccess.Read);

            // Browser-Anmeldung erkennen: geschieht immer wenn kein gespeichertes Token vorhanden ist; andernfalls verrät die Dauer der Interaktion, ob ein Browser-Login nötig war.
            var tokenExisted = File.Exists(Path.Combine(tokenDir, TokenFileName));
            var authWatch = System.Diagnostics.Stopwatch.StartNew();
            var credential = await GoogleWebAuthorizationBroker.AuthorizeAsync(GoogleClientSecrets.FromStream(stream).Secrets, scopes, "user", token, new FileDataStore(tokenDir, true));
            authWatch.Stop();
            LastAuthorizationWasInteractive = !tokenExisted || authWatch.ElapsedMilliseconds > 3000;

            _cachedService = new PeopleServiceService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = Application.ProductName,
            });
            return _cachedService;
        }
        finally { _serviceLock.Release(); }
    }

    public static void ClearServiceCache() => _cachedService = null;

    private static async Task<Dictionary<string, string>> GetContactGroupsMapAsync(PeopleServiceService service, CancellationToken token = default)
    {
        var map = new Dictionary<string, string>();
        try
        {
            var req = service.ContactGroups.List();
            req.GroupFields = GROUP_FIELDS;
            var res = await req.ExecuteAsync(token);

            if (res.ContactGroups != null)
            {
                foreach (var g in res.ContactGroups)
                {
                    if (string.IsNullOrEmpty(g.ResourceName)) { continue; }
                    var isUserGroup = g.GroupType == "USER_CONTACT_GROUP";
                    var isStarred = g.ResourceName == GROUP_STARRED_RESOURCE;

                    if (isUserGroup || isStarred)
                    {
                        var name = g.FormattedName ?? g.Name;
                        map[g.ResourceName] = name;
                    }
                }
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Warnung: Kontaktgruppen konnten nicht geladen werden: {ex.Message}");
            // map bleibt leer — Aufrufer muss entscheiden, ob das akzeptabel ist; Alle Gruppen aller Kontakte gehen beim nächsten Sync verloren
        }
        return map;
    }

    private static async Task<string> CreateContactGroupInternalAsync(PeopleServiceService service, string groupName, CancellationToken token = default)
    {
        try
        {
            var group = new ContactGroup { Name = groupName };
            var requestBody = new CreateContactGroupRequest { ContactGroup = group };
            var createdGroup = await service.ContactGroups.Create(requestBody).ExecuteAsync(token);
            return createdGroup.ResourceName;
        }
        catch { return string.Empty; }
    }

    private static async Task CheckAndDeleteEmptyGroupsInternalAsync(PeopleServiceService service, HashSet<string> groupResourceNames, CancellationToken token)
    {
        foreach (var resourceName in groupResourceNames)
        {
            if (resourceName.Contains("system") || resourceName.Contains(GROUP_STARRED_LABEL) || resourceName.Contains(GROUP_MYCONTACTS_LABEL)) { continue; }
            try
            {
                var groupReq = service.ContactGroups.Get(resourceName);
                groupReq.GroupFields = "memberCount";
                var group = await groupReq.ExecuteAsync(token);
                if (group.MemberCount == 0)
                {
                    await service.ContactGroups.Delete(resourceName).ExecuteAsync(token);
                }
            }
            catch { }
        }
    }

    private static async Task<(string? PhotoUrl, string? ETag)> UploadPhotoInternalAsync(
        PeopleServiceService service, string resourceName, Image image, ImageFormat format, CancellationToken token)
    {
        using var clonedImage = new Bitmap(image);
        using var ms = new MemoryStream();
        clonedImage.Save(ms, format);
        var base64Photo = Convert.ToBase64String(ms.ToArray());

        var updatePhotoRequest = new UpdateContactPhotoRequest
        {
            PhotoBytes = base64Photo,
            PersonFields = "photos"
        };
        var response = await service.People.UpdateContactPhoto(updatePhotoRequest, resourceName).ExecuteAsync(token);
        var photoUrl = response?.Person?.Photos?.FirstOrDefault()?.Url;
        var etag = response?.Person?.ETag;
        return (photoUrl, etag);
    }

    private static Contact MapPersonToContact(Person person, Dictionary<string, string> groupMap)
    {
        var addr = GetGoogleAddressByType(person, TYPE_HOME, TYPE_WORK, TYPE_OTHER);
        var phones = MapPhoneSlots(person);
        var emails = MapEmailSlots(person);
        var newContact = new Contact
        {
            RawGooglePerson = person,
            ResourceName = person.ResourceName,
            ETag = person.ETag,
            Praefix = person.Names?.FirstOrDefault()?.HonorificPrefix ?? "",
            Nachname = person.Names?.FirstOrDefault()?.FamilyName ?? "",
            Vorname = person.Names?.FirstOrDefault()?.GivenName ?? "",
            Zwischenname = person.Names?.FirstOrDefault()?.MiddleName ?? "",
            Nickname = person.Nicknames?.FirstOrDefault()?.Value ?? "",
            Suffix = person.Names?.FirstOrDefault()?.HonorificSuffix ?? "",
            Unternehmen = person.Organizations?.FirstOrDefault()?.Name ?? "",
            Position = person.Organizations?.FirstOrDefault()?.Title ?? "",
            Strasse = addr?.StreetAddress ?? "",
            PLZ = addr?.PostalCode ?? "",
            Ort = addr?.City ?? "",
            Postfach = addr?.PoBox ?? "",
            Land = addr?.Country ?? "",
            Notizen = person.Biographies?.FirstOrDefault()?.Value.ReplaceLineEndings() ?? "",
            Internet = GetGoogleUrlByType(person, TYPE_HOMEPAGE, TYPE_HOME, TYPE_WORK, TYPE_OTHER),
            Mail1 = emails.Mail1?.Value ?? "",
            Mail2 = emails.Mail2?.Value ?? "",
            Telefon1 = phones.Telefon1?.Value ?? "",
            Telefon2 = phones.Telefon2?.Value ?? "",
            Mobil = phones.Mobil?.Value ?? "",
            Fax = phones.Fax?.Value ?? "",
        };

        if (person.UserDefined != null)
        {
            foreach (var f in person.UserDefined)
            {
                if (f.Key == KEY_ANREDE) { newContact.Anrede = f.Value; }
                else if (f.Key == KEY_BETREFF) { newContact.Betreff = f.Value; }
                else if (f.Key == KEY_GRUSS) { newContact.Grussformel = f.Value; }
                else if (f.Key == KEY_SCHLUSS) { newContact.Schlussformel = f.Value; }
            }
        }

        if (person.Birthdays != null && person.Birthdays.Count > 0 && person.Birthdays[0].Date != null)
        {
            var bday = person.Birthdays[0].Date;
            try { newContact.Geburtstag = new DateOnly(bday.Year ?? 1900, bday.Month ?? 1, bday.Day ?? 1); }
            catch { }
        }

        if (person.Photos != null)
        {
            var photo = person.Photos.FirstOrDefault(p => !string.IsNullOrEmpty(p.Url));
            if (photo != null && photo.Default__ != true) { newContact.PhotoUrl = photo.Url; }
            //if (photo != null && (!photo.Default__ ?? true)) { newContact.PhotoUrl = photo.Url; }  // !photo.Default__ führt zu einem Kompilierfehler, da ! nur für bool, nicht für bool? definiert ist
        }

        var groupNames = new HashSet<string>();
        if (person.Memberships != null)
        {
            foreach (var m in person.Memberships)
            {
                if (m.ContactGroupMembership?.ContactGroupResourceName != null && groupMap.TryGetValue(m.ContactGroupMembership.ContactGroupResourceName, out var gName))
                {
                    groupNames.Add(gName.Equals(GROUP_STARRED_LABEL, StringComparison.OrdinalIgnoreCase) ? STAR_SYMBOL : gName);
                }
            }
        }
        newContact.GroupNames = [.. groupNames];

        newContact.LastModified = person.Metadata?.Sources?.FirstOrDefault(static s => s.Type == "CONTACT")?.UpdateTimeDateTimeOffset?.UtcDateTime;
        return newContact;
    }

    /// <summary>
    /// Verteilt die Telefonnummern eines Kontakts auf die vier Programmfelder. Jeder Google-Eintrag wird höchstens EINEM Feld zugeordnet
    /// (keine Doppelanzeige, keine Duplikate beim Schreiben). Reihenfolge: erst die bevorzugten Typen je Feld, danach werden übrig gebliebene
    /// Nummern (eigene Labels, zweite Nummer eines Typs, Nummern ohne Typ) auf noch freie Telefonfelder verteilt – nicht auf Fax.
    /// </summary>
    private static PhoneSlots MapPhoneSlots(Person person)
    {
        var numbers = (person.PhoneNumbers ?? []).Where(static p => !string.IsNullOrEmpty(p.Value)).ToList();
        var used = new HashSet<PhoneNumber>(ReferenceEqualityComparer.Instance);
        PhoneNumber? Pick(params string[] types)
        {
            foreach (var type in types)
            {
                var match = numbers.FirstOrDefault(p => !used.Contains(p) && p.Type?.Equals(type, StringComparison.OrdinalIgnoreCase) == true);
                if (match != null) { used.Add(match); return match; }
            }
            return null;
        }
        var telefon1 = Pick(TYPE_HOME, TYPE_OTHER, TYPE_MAIN);
        var telefon2 = Pick(TYPE_WORK, TYPE_MAIN, TYPE_WORK_MOBILE);
        var mobil = Pick(TYPE_MOBILE, TYPE_WORK_MOBILE, TYPE_PAGER);
        var fax = Pick(TYPE_FAX, TYPE_HOME_FAX, TYPE_WORK_FAX, TYPE_OTHER_FAX);
        foreach (var rest in numbers.Where(p => !used.Contains(p) && !_faxTypes.Contains(p.Type ?? "")))
        {
            if (telefon1 == null) { telefon1 = rest; }
            else if (telefon2 == null) { telefon2 = rest; }
            else if (mobil == null) { mobil = rest; }
            else { break; }  // alle Felder belegt – weitere Nummern bleiben bei Google unangetastet
            used.Add(rest);
        }
        return new PhoneSlots(telefon1, telefon2, mobil, fax);
    }

    /// <summary>Verteilt die E-Mail-Adressen analog zu <see cref="MapPhoneSlots"/>; Adressen ohne Typ (bei Google häufig) landen auf dem ersten freien Feld.</summary>
    private static EmailSlots MapEmailSlots(Person person)
    {
        var emails = (person.EmailAddresses ?? []).Where(static e => !string.IsNullOrEmpty(e.Value)).ToList();
        var used = new HashSet<EmailAddress>(ReferenceEqualityComparer.Instance);
        EmailAddress? Pick(params string[] types)
        {
            foreach (var type in types)
            {
                var match = emails.FirstOrDefault(e => !used.Contains(e) && e.Type?.Equals(type, StringComparison.OrdinalIgnoreCase) == true);
                if (match != null) { used.Add(match); return match; }
            }
            return null;
        }
        var mail1 = Pick(TYPE_HOME, TYPE_OTHER);
        var mail2 = Pick(TYPE_WORK);
        foreach (var rest in emails.Where(e => !used.Contains(e)))
        {
            if (mail1 == null) { mail1 = rest; }
            else if (mail2 == null) { mail2 = rest; }
            else { break; }
            used.Add(rest);
        }
        return new EmailSlots(mail1, mail2);
    }

    /// <summary>
    /// Liefert für die Anzeige (Tooltip) das tatsächliche Google-Label des Eintrags, der aktuell hinter jedem Telefon-/E-Mail-Feld steht –
    /// berechnet aus <see cref="Contact.RawGooglePerson"/> mit derselben Zuordnung wie Lesen und Schreiben. null = kein Eintrag (Feld ist neu oder leer).
    /// </summary>
    internal static ContactFieldLabels GetFieldLabels(Contact contact)
    {
        if (contact.RawGooglePerson == null) { return new ContactFieldLabels(null, null, null, null, null, null); }
        var phones = MapPhoneSlots(contact.RawGooglePerson);
        var emails = MapEmailSlots(contact.RawGooglePerson);
        return new ContactFieldLabels(
            PhoneLabel(phones.Telefon1), PhoneLabel(phones.Telefon2), PhoneLabel(phones.Mobil), PhoneLabel(phones.Fax),
            EmailLabel(emails.Mail1), EmailLabel(emails.Mail2));
    }

    private static string? PhoneLabel(PhoneNumber? p) => p == null ? null : TypeLabel(p.Type, p.FormattedType, static t => t switch
    {
        TYPE_HOME => "Privat",
        TYPE_WORK => "Geschäftlich",
        TYPE_MOBILE => "Mobil",
        TYPE_FAX => "Fax",
        TYPE_HOME_FAX => "Fax privat",
        TYPE_WORK_FAX => "Fax geschäftlich",
        TYPE_OTHER_FAX => "Fax sonstige",
        TYPE_WORK_MOBILE => "Mobil geschäftlich",
        TYPE_MAIN => "Hauptnummer",
        TYPE_PAGER => "Pager",
        TYPE_OTHER => "Sonstige",
        "googleVoice" => "Google Voice",
        "workPager" => "Pager geschäftlich",
        _ => null
    });

    private static string? EmailLabel(EmailAddress? e) => e == null ? null : TypeLabel(e.Type, e.FormattedType, static t => t switch
    {
        TYPE_HOME => "Privat",
        TYPE_WORK => "Geschäftlich",
        TYPE_OTHER => "Sonstige",
        _ => null
    });

    // Bekannte Google-Typen auf Deutsch; eigene Labels stehen bei Google direkt im Type-Feld und werden unverändert angezeigt.
    private static string TypeLabel(string? type, string? formattedType, Func<string, string?> knownTypes)
    {
        if (string.IsNullOrWhiteSpace(type)) { return "ohne Label"; }
        return knownTypes(type) ?? (string.IsNullOrWhiteSpace(formattedType) ? type : formattedType);
    }

    private static Address? GetGoogleAddressByType(Person person, string primaryType, params string[] fallbackTypes)
    {
        var addresses = person.Addresses ?? [];
        foreach (var type in fallbackTypes.Prepend(primaryType))
        {
            var match = addresses.FirstOrDefault(a => a.Type?.Equals(type, StringComparison.OrdinalIgnoreCase) == true);
            if (match != null) { return match; }
        }
        return addresses.FirstOrDefault(); // absoluter Fallback: irgendeineAdresse
    }

    private static string GetGoogleUrlByType(Person person, string primaryType, params string[] fallbackTypes)
        => GetGoogleUrlEntry(person, primaryType, fallbackTypes)?.Value ?? string.Empty;

    private static Url? GetGoogleUrlEntry(Person person, string primaryType, params string[] fallbackTypes)  // gemeinsame Auswahlregel für Lesen (MapPersonToContact) und Schreiben (UpdateContactAsync)
    {
        var urls = person.Urls ?? [];
        foreach (var type in fallbackTypes.Prepend(primaryType))
        {
            var match = urls.FirstOrDefault(u => u.Type?.Equals(type, StringComparison.OrdinalIgnoreCase) == true && !string.IsNullOrEmpty(u.Value));
            if (match != null) { return match; }
        }
        return urls.FirstOrDefault(u => !string.IsNullOrEmpty(u.Value));  // absoluter Fallback: irgendeine URL mit Wert
    }

    /// <summary>
    /// Schreibt die vier Telefonfelder zurück. Die Zuordnung Feld → Eintrag wird mit derselben Regel wie beim Lesen berechnet (MapPhoneSlots),
    /// BEVOR etwas verändert wird. Nur tatsächlich geänderte Felder werden angefasst; Typ/Label eines vorhandenen Eintrags bleibt erhalten.
    /// Leeres Feld = Eintrag entfernen; neues Feld ohne Eintrag = neuen Eintrag mit Standardtyp anlegen.
    /// </summary>
    private static void ApplyPhoneChanges(Person person, Contact contact)
    {
        person.PhoneNumbers ??= [];
        var slots = MapPhoneSlots(person);
        ApplyPhoneSlot(person.PhoneNumbers, slots.Telefon1, contact.Telefon1, TYPE_HOME);
        ApplyPhoneSlot(person.PhoneNumbers, slots.Telefon2, contact.Telefon2, TYPE_WORK);
        ApplyPhoneSlot(person.PhoneNumbers, slots.Mobil, contact.Mobil, TYPE_MOBILE);
        ApplyPhoneSlot(person.PhoneNumbers, slots.Fax, contact.Fax, TYPE_FAX);
    }

    private static void ApplyPhoneSlot(IList<PhoneNumber> numbers, PhoneNumber? existing, string? newValue, string defaultType)
    {
        if (string.IsNullOrWhiteSpace(newValue))
        {
            if (existing != null) { numbers.Remove(existing); }
        }
        else if (existing == null) { numbers.Add(new PhoneNumber { Type = defaultType, Value = newValue }); }
        else if (!string.Equals(existing.Value, newValue, StringComparison.Ordinal)) { existing.Value = newValue; }  // Typ bleibt unverändert
    }

    /// <summary>E-Mail-Pendant zu <see cref="ApplyPhoneChanges"/>.</summary>
    private static void ApplyEmailChanges(Person person, Contact contact)
    {
        person.EmailAddresses ??= [];
        var slots = MapEmailSlots(person);
        ApplyEmailSlot(person.EmailAddresses, slots.Mail1, contact.Mail1, TYPE_HOME);
        ApplyEmailSlot(person.EmailAddresses, slots.Mail2, contact.Mail2, TYPE_WORK);
    }

    private static void ApplyEmailSlot(IList<EmailAddress> emails, EmailAddress? existing, string? newValue, string defaultType)
    {
        if (string.IsNullOrWhiteSpace(newValue))
        {
            if (existing != null) { emails.Remove(existing); }
        }
        else if (existing == null) { emails.Add(new EmailAddress { Type = defaultType, Value = newValue }); }
        else if (!string.Equals(existing.Value, newValue, StringComparison.Ordinal)) { existing.Value = newValue; }  // Typ bleibt unverändert
    }

    private static void UpdateGoogleUserDef(Person person, string targetKey, string? newValue)
    {
        person.UserDefined ??= [];
        var existing = person.UserDefined.FirstOrDefault(e => e.Key?.Equals(targetKey, StringComparison.OrdinalIgnoreCase) == true);
        if (string.IsNullOrWhiteSpace(newValue))
        {
            if (existing != null) { person.UserDefined.Remove(existing); }
        }
        else
        {
            if (existing == null)
            {
                existing = new UserDefined { Key = targetKey };
                person.UserDefined.Add(existing);
            }
            existing.Value = newValue;
        }
    }
}