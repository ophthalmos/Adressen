using System.Drawing.Imaging;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Auth.OAuth2.Responses;
using Google.Apis.Oauth2.v2;
using Google.Apis.PeopleService.v1;
using Google.Apis.PeopleService.v1.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store; // Für FileDataStore

namespace Adressen.cls;

// Definition des Rückgabewerts für LoadContacts
internal record LoadContactsResult(List<Contact> Contacts, string UserEmail, Dictionary<string, string> GroupMap);

internal class GooglePeopleManager(string secretPath, string tokenDir)
{
    // ========================================================================
    // 1. PUBLIC API: LOAD, CREATE, UPDATE, DELETE
    // ========================================================================

    public async Task<LoadContactsResult> LoadContactsAsync(HashSet<string> excludedGroups)
    {
        try
        {
            var service = await GetServiceAsync();
            // A. Email abrufen (Direkt über People Service, statt Oauth2Service)
            var userEmail = string.Empty;
            try
            {
                var meReq = service.People.Get("people/me");
                meReq.PersonFields = "emailAddresses";
                var me = await meReq.ExecuteAsync();
                if (me.EmailAddresses != null && me.EmailAddresses.Count > 0)
                {
                    userEmail = me.EmailAddresses.FirstOrDefault()?.Value ?? string.Empty;
                }
            }
            catch //(Exception ex)
            {
                //Utils.ErrTaskDlg(nint.Zero, ex); // sollte nicht den gesamten Import stoppen
            }
            finally
            {
                if (string.IsNullOrEmpty(userEmail)) { userEmail = "Meine Kontakte"; }
            }
            //var userEmail = string.Empty;
            //var emailReq = await oauthService.Userinfo.Get().ExecuteAsync();
            //if (emailReq.VerifiedEmail == true) { userEmail = emailReq.Email; }

            // B. Gruppen laden (Mapping ID -> Name)
            var groupMap = await GetContactGroupsMapAsync(service);

            // C. Kontakte laden
            var peopleRequest = service.People.Connections.List("people/me");
            peopleRequest.PersonFields = "names,memberships,nicknames,addresses,phoneNumbers,emailAddresses,biographies,birthdays,urls,organizations,photos,userDefined";
            peopleRequest.SortOrder = (PeopleResource.ConnectionsResource.ListRequest.SortOrderEnum)3; // LAST_NAME_ASCENDING
            peopleRequest.PageSize = 2000;

            var response = await peopleRequest.ExecuteAsync();
            List<Contact> contactList = [];

            if (response?.Connections != null)
            {
                foreach (var person in response.Connections)
                {
                    contactList.Add(MapPersonToContact(person, groupMap, excludedGroups));
                }
            }
            return new LoadContactsResult(contactList, userEmail, groupMap);
        }
        catch (TokenResponseException ex)
        {
            throw new UnauthorizedAccessException("Google Token abgelaufen", ex);
        }
    }

    public async Task<Contact> CreateContactAsync(Contact contact, Image? profileImage)
    {
        var service = await GetServiceAsync();

        // 1. Person-Objekt bauen (Mapping Local -> Google)
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
                Type = "work"
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
            Urls = !string.IsNullOrWhiteSpace(contact.Internet)
                ? [new() { Value = contact.Internet }] : null,
            Biographies = !string.IsNullOrWhiteSpace(contact.Notizen)
                ? [new() { Value = contact.Notizen }] : null
        };

        // Listen für Mehrfachfelder
        List<EmailAddress> emails = [];
        if (!string.IsNullOrWhiteSpace(contact.Mail1))
        {
            emails.Add(new EmailAddress { Value = contact.Mail1, Type = "home" });
        }

        if (!string.IsNullOrWhiteSpace(contact.Mail2))
        {
            emails.Add(new EmailAddress { Value = contact.Mail2, Type = "work" });
        }

        if (emails.Count > 0)
        {
            personToCreate.EmailAddresses = emails;
        }

        List<PhoneNumber> phones = [];
        if (!string.IsNullOrWhiteSpace(contact.Telefon1))
        {
            phones.Add(new PhoneNumber { Value = contact.Telefon1, Type = "home" });
        }

        if (!string.IsNullOrWhiteSpace(contact.Telefon2))
        {
            phones.Add(new PhoneNumber { Value = contact.Telefon2, Type = "work" });
        }

        if (!string.IsNullOrWhiteSpace(contact.Mobil))
        {
            phones.Add(new PhoneNumber { Value = contact.Mobil, Type = "mobile" });
        }

        if (!string.IsNullOrWhiteSpace(contact.Fax))
        {
            phones.Add(new PhoneNumber { Value = contact.Fax, Type = "fax" });
        }

        if (phones.Count > 0)
        {
            personToCreate.PhoneNumbers = phones;
        }

        List<UserDefined> userDefined = [];
        if (!string.IsNullOrWhiteSpace(contact.Anrede))
        {
            userDefined.Add(new UserDefined { Key = "Anrede", Value = contact.Anrede });
        }

        if (!string.IsNullOrWhiteSpace(contact.Betreff))
        {
            userDefined.Add(new UserDefined { Key = "Betreff", Value = contact.Betreff });
        }

        if (!string.IsNullOrWhiteSpace(contact.Grussformel))
        {
            userDefined.Add(new UserDefined { Key = "Grussformel", Value = contact.Grussformel });
        }

        if (!string.IsNullOrWhiteSpace(contact.Schlussformel))
        {
            userDefined.Add(new UserDefined { Key = "Schlussformel", Value = contact.Schlussformel });
        }

        if (userDefined.Count > 0)
        {
            personToCreate.UserDefined = userDefined;
        }

        // 2. Erstellen
        var createdPerson = await service.People.CreateContact(personToCreate).ExecuteAsync();

        // 3. ID zurückschreiben
        contact.ResourceName = createdPerson.ResourceName;
        contact.ETag = createdPerson.ETag;

        // 4. Foto Upload (falls vorhanden)
        if (profileImage != null && !string.IsNullOrEmpty(contact.ResourceName))
        {
            var photoUrl = await UploadPhotoInternalAsync(service, contact.ResourceName, profileImage, profileImage.RawFormat);
            if (photoUrl != null)
            {
                contact.PhotoUrl = photoUrl;
            }
        }

        return contact;
    }

    public async Task<Contact> UpdateContactAsync(
        Contact contact,
        List<string> changedFields,
        Dictionary<string, string> groupMap,
        Contact? originalContactSnapshot,
        bool checkEmptyGroups)
    {
        var service = await GetServiceAsync();

        // 1. Person-Objekt vorbereiten
        var personToUpdate = new Person
        {
            ResourceName = contact.ResourceName,
            ETag = contact.ETag
        };

        // 2. Mapping der geänderten Felder
        if (changedFields.Contains("names"))
        {
            personToUpdate.Names = [new() {
                HonorificPrefix = contact.Praefix,
                FamilyName = contact.Nachname,
                GivenName = contact.Vorname,
                MiddleName = contact.Zwischenname,
                HonorificSuffix = contact.Suffix
            }];
        }

        if (changedFields.Contains("nicknames"))
        {
            personToUpdate.Nicknames = [new Nickname { Value = contact.Nickname }];
        }

        if (changedFields.Contains("addresses"))
        {
            personToUpdate.Addresses = [new() {
                StreetAddress = contact.Strasse,
                PostalCode = contact.PLZ,
                City = contact.Ort,
                PoBox = contact.Postfach,
                Country = contact.Land
            }];
        }

        if (changedFields.Contains("organizations"))
        {
            personToUpdate.Organizations = [new Organization { Name = contact.Unternehmen, Title = contact.Position }];
        }

        if (changedFields.Contains("birthdays") && contact.Geburtstag.HasValue)
        {
            personToUpdate.Birthdays = [new() {
                Date = new Date {
                    Day = contact.Geburtstag.Value.Day,
                    Month = contact.Geburtstag.Value.Month,
                    Year = contact.Geburtstag.Value.Year
                }
            }];
        }

        if (changedFields.Contains("emailAddresses"))
        {
            personToUpdate.EmailAddresses = [];
            if (!string.IsNullOrWhiteSpace(contact.Mail1))
            {
                personToUpdate.EmailAddresses.Add(new EmailAddress { Value = contact.Mail1, Type = "home" });
            }

            if (!string.IsNullOrWhiteSpace(contact.Mail2))
            {
                personToUpdate.EmailAddresses.Add(new EmailAddress { Value = contact.Mail2, Type = "work" });
            }
        }

        if (changedFields.Contains("phoneNumbers"))
        {
            personToUpdate.PhoneNumbers = [];
            if (!string.IsNullOrWhiteSpace(contact.Telefon1))
            {
                personToUpdate.PhoneNumbers.Add(new PhoneNumber { Value = contact.Telefon1, Type = "home" });
            }

            if (!string.IsNullOrWhiteSpace(contact.Telefon2))
            {
                personToUpdate.PhoneNumbers.Add(new PhoneNumber { Value = contact.Telefon2, Type = "work" });
            }

            if (!string.IsNullOrWhiteSpace(contact.Mobil))
            {
                personToUpdate.PhoneNumbers.Add(new PhoneNumber { Value = contact.Mobil, Type = "mobile" });
            }

            if (!string.IsNullOrWhiteSpace(contact.Fax))
            {
                personToUpdate.PhoneNumbers.Add(new PhoneNumber { Value = contact.Fax, Type = "fax" });
            }
        }

        if (changedFields.Contains("urls"))
        {
            personToUpdate.Urls = [new Url { Value = contact.Internet, Type = "homePage" }];
        }

        if (changedFields.Contains("biographies"))
        {
            personToUpdate.Biographies = [new Biography { Value = contact.Notizen }];
        }

        if (changedFields.Contains("userDefined"))
        {
            personToUpdate.UserDefined = [];
            if (!string.IsNullOrWhiteSpace(contact.Anrede))
            {
                personToUpdate.UserDefined.Add(new UserDefined { Key = "Anrede", Value = contact.Anrede });
            }

            if (!string.IsNullOrWhiteSpace(contact.Betreff))
            {
                personToUpdate.UserDefined.Add(new UserDefined { Key = "Betreff", Value = contact.Betreff });
            }

            if (!string.IsNullOrWhiteSpace(contact.Grussformel))
            {
                personToUpdate.UserDefined.Add(new UserDefined { Key = "Grussformel", Value = contact.Grussformel });
            }

            if (!string.IsNullOrWhiteSpace(contact.Schlussformel))
            {
                personToUpdate.UserDefined.Add(new UserDefined { Key = "Schlussformel", Value = contact.Schlussformel });
            }
        }

        // Gruppen Logik
        HashSet<string> groupsToRemoveToCheck = [];
        if (changedFields.Contains("memberships"))
        {
            personToUpdate.Memberships = [];
            var desiredGroupNames = new HashSet<string>(contact.GroupNames, StringComparer.OrdinalIgnoreCase);

            if (desiredGroupNames.Remove("★"))
            {
                desiredGroupNames.Add("starred");
            }

            desiredGroupNames.Add("myContacts");

            foreach (var groupName in desiredGroupNames)
            {
                string resourceName;
                var existingEntry = groupMap.FirstOrDefault(x => x.Value.Equals(groupName, StringComparison.OrdinalIgnoreCase));

                if (!string.IsNullOrEmpty(existingEntry.Key))
                {
                    resourceName = existingEntry.Key;
                }
                else if (groupName.Equals("myContacts", StringComparison.OrdinalIgnoreCase) || groupName.Equals("starred", StringComparison.OrdinalIgnoreCase))
                {
                    resourceName = "contactGroups/" + groupName.ToLowerInvariant(); // Google mag Kleinschreibung bei Systemgruppen
                }
                else
                {
                    resourceName = await CreateContactGroupInternalAsync(service, groupName);
                    if (string.IsNullOrEmpty(resourceName))
                    {
                        continue;
                    }

                    groupMap[resourceName] = groupName;
                }

                personToUpdate.Memberships.Add(new Membership
                {
                    ContactGroupMembership = new ContactGroupMembership { ContactGroupResourceName = resourceName }
                });
            }

            // Cleanup Check vorbereiten
            if (checkEmptyGroups && originalContactSnapshot != null)
            {
                var originalGroups = originalContactSnapshot.GroupNames
                    .Select(g => g == "★" ? "starred" : g)
                    .ToHashSet(StringComparer.OrdinalIgnoreCase);

                foreach (var rem in originalGroups.Except(desiredGroupNames))
                {
                    var resKey = groupMap.FirstOrDefault(x => x.Value.Equals(rem, StringComparison.OrdinalIgnoreCase)).Key;
                    if (!string.IsNullOrEmpty(resKey))
                    {
                        groupsToRemoveToCheck.Add(resKey);
                    }
                }
            }
        }

        // 3. API Request senden
        if (changedFields.Count > 0)
        {
            var updateRequest = service.People.UpdateContact(personToUpdate, contact.ResourceName);
            updateRequest.UpdatePersonFields = string.Join(",", changedFields);

            var updatedPerson = await updateRequest.ExecuteAsync();

            contact.ETag = updatedPerson.ETag;
            contact.ResourceName = updatedPerson.ResourceName;

            if (checkEmptyGroups && groupsToRemoveToCheck.Count > 0)
            {
                await CheckAndDeleteEmptyGroupsInternalAsync(service, groupsToRemoveToCheck);
            }
        }

        return contact;
    }

    public async Task DeleteContactAsync(string resourceName)
    {
        if (string.IsNullOrWhiteSpace(resourceName))
        {
            return;
        }

        var service = await GetServiceAsync();
        await service.People.DeleteContact(resourceName).ExecuteAsync();
    }

    // ========================================================================
    // 2. FOTO API
    // ========================================================================

    public async Task<string?> UpdateContactPhotoAsync(string resourceName, Image image, ImageFormat format)
    {
        var service = await GetServiceAsync();
        return await UploadPhotoInternalAsync(service, resourceName, image, format);
    }

    public async Task<string?> DeleteContactPhotoAsync(string resourceName)
    {
        var service = await GetServiceAsync();
        var request = service.People.DeleteContactPhoto(resourceName);
        request.PersonFields = "photos";
        var response = await request.ExecuteAsync();
        return response?.Person?.Photos?.FirstOrDefault()?.Url;
    }

    // ========================================================================
    // 3. INTERNE HILFSMETHODEN (PRIVATE)
    // ========================================================================

    private async Task<PeopleServiceService> GetServiceAsync()
    {
        string[] scopes = [PeopleServiceService.Scope.Contacts, PeopleServiceService.Scope.UserinfoEmail, PeopleServiceService.Scope.UserinfoProfile];
        UserCredential credential;
        using (FileStream stream = new(secretPath, FileMode.Open, FileAccess.Read))
        {
            credential = await GoogleWebAuthorizationBroker.AuthorizeAsync(
                GoogleClientSecrets.FromStream(stream).Secrets,
                scopes,
                "user",
                CancellationToken.None,
                new FileDataStore(tokenDir, true));
        }
        return new PeopleServiceService(new BaseClientService.Initializer()
        {
            HttpClientInitializer = credential,
            ApplicationName = Application.ProductName,
        });
    }

    private static async Task<Dictionary<string, string>> GetContactGroupsMapAsync(PeopleServiceService service)
    {
        var map = new Dictionary<string, string>();
        try
        {
            var req = service.ContactGroups.List();
            req.GroupFields = "name,clientData";
            var res = await req.ExecuteAsync();
            if (res.ContactGroups != null)
            {
                foreach (var g in res.ContactGroups)
                {
                    var name = g.FormattedName ?? g.Name;
                    if (!string.IsNullOrEmpty(g.ResourceName))
                    {
                        map[g.ResourceName] = name;
                    }
                }
            }
        }
        catch { }
        return map;
    }

    private static async Task<string> CreateContactGroupInternalAsync(PeopleServiceService service, string groupName)
    {
        try
        {
            var group = new ContactGroup { Name = groupName };
            var requestBody = new CreateContactGroupRequest { ContactGroup = group };
            var createdGroup = await service.ContactGroups.Create(requestBody).ExecuteAsync();
            return createdGroup.ResourceName;
        }
        catch { return string.Empty; }
    }

    private static async Task CheckAndDeleteEmptyGroupsInternalAsync(PeopleServiceService service, HashSet<string> groupResourceNames)
    {
        foreach (var resourceName in groupResourceNames)
        {
            if (resourceName.Contains("system") || resourceName.Contains("starred") || resourceName.Contains("myContacts"))
            {
                continue;
            }

            try
            {
                var groupReq = service.ContactGroups.Get(resourceName);
                groupReq.GroupFields = "memberCount";
                var group = await groupReq.ExecuteAsync();
                if (group.MemberCount == 0)
                {
                    await service.ContactGroups.Delete(resourceName).ExecuteAsync();
                }
            }
            catch { }
        }
    }

    private static async Task<string?> UploadPhotoInternalAsync(PeopleServiceService service, string resourceName, Image image, ImageFormat format)
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

        var response = await service.People.UpdateContactPhoto(updatePhotoRequest, resourceName).ExecuteAsync();
        return response?.Person?.Photos?.FirstOrDefault()?.Url;
    }

    private static Contact MapPersonToContact(Person person, Dictionary<string, string> groupMap, HashSet<string> excludedGroups)
    {
        var newContact = new Contact
        {
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
            Strasse = person.Addresses?.FirstOrDefault()?.StreetAddress ?? "",
            PLZ = person.Addresses?.FirstOrDefault()?.PostalCode ?? "",
            Ort = person.Addresses?.FirstOrDefault()?.City ?? "",
            Postfach = person.Addresses?.FirstOrDefault()?.PoBox ?? "",
            Land = person.Addresses?.FirstOrDefault()?.Country ?? "",
            Notizen = person.Biographies?.FirstOrDefault()?.Value.ReplaceLineEndings() ?? "",
            Internet = person.Urls?.FirstOrDefault()?.Value ?? "",
            Mail1 = person.EmailAddresses?.FirstOrDefault()?.Value ?? "",
            Mail2 = (person.EmailAddresses?.Count > 1) ? person.EmailAddresses[1].Value : "",
            Telefon1 = GetGooglePhoneByType(person, "home") ?? "",
            Telefon2 = GetGooglePhoneByType(person, "work") ?? "",
            Mobil = GetGooglePhoneByType(person, "mobile") ?? "",
            Fax = GetGooglePhoneByType(person, "fax") ?? ""
        };

        if (person.UserDefined != null)
        {
            foreach (var f in person.UserDefined)
            {
                if (f.Key == "Anrede")
                {
                    newContact.Anrede = f.Value;
                }
                else if (f.Key == "Betreff")
                {
                    newContact.Betreff = f.Value;
                }
                else if (f.Key == "Grussformel")
                {
                    newContact.Grussformel = f.Value;
                }
                else if (f.Key == "Schlussformel")
                {
                    newContact.Schlussformel = f.Value;
                }
            }
        }

        if (person.Birthdays != null && person.Birthdays.Count > 0 && person.Birthdays[0].Date != null)
        {
            var bday = person.Birthdays[0].Date;
            try { newContact.Geburtstag = new DateOnly(bday.Year ?? 1900, bday.Month ?? 1, bday.Day ?? 1); } catch { }
        }

        if (person.Photos != null)
        {
            var photo = person.Photos.FirstOrDefault(p => !string.IsNullOrEmpty(p.Url));
            if (photo != null && (!photo.Default__ ?? true))
            {
                newContact.PhotoUrl = photo.Url;
            }
        }

        var groupNames = new HashSet<string>();
        if (person.Memberships != null)
        {
            //MessageBox.Show($"excludedGroups: {string.Join(", ", excludedGroups)}");
            foreach (var m in person.Memberships)
            {
                if (m.ContactGroupMembership?.ContactGroupResourceName != null &&
                    groupMap.TryGetValue(m.ContactGroupMembership.ContactGroupResourceName, out var gName))
                {
                    if (!excludedGroups.Contains(gName))
                    {
                        groupNames.Add(gName.Equals("starred", StringComparison.OrdinalIgnoreCase) ? "★" : gName);
                    }
                }
            }
        }
        newContact.GroupNames = [.. groupNames];
        //MessageBox.Show($"Loaded contact: {newContact.DisplayName} with groups: {string.Join(", ", newContact.GroupNames)}");
        return newContact;
    }

    internal static string GetGooglePhoneByType(Person person, string type)
    {
        foreach (var phone in person.PhoneNumbers ?? [])
        {
            if (phone.Type?.Contains(type, StringComparison.OrdinalIgnoreCase) == true)
            {
                return phone.Value ?? string.Empty;
            }
        }
        return string.Empty;
    }
}
