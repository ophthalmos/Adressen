namespace Adressen.cls;

public interface IContactEntity
{
    string UniqueId
    {
        get;
    }
    string DisplayName
    {
        get;
    }
    string SearchText
    {
        get;
    }
    DateOnly? BirthdayDate
    {
        get;
    }

    bool Reminder
    {
        get => true;
        set
        {
        }
    }

    IList<string> GroupList
    {
        get;
    }
    string? Vorname
    {
        get; set;
    }
    string? Nachname
    {
        get; set;
    }
    string? Mail1
    {
        get; set;
    }
    Task<Image?> GetPhotoAsync(CancellationToken token = default);
    void ResetSearchCache();
}
