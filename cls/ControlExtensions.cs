using System.ComponentModel.DataAnnotations;
using System.Reflection;

namespace Adressen.cls;

public static class ControlExtensions
{
    /// <summary>
    /// Durchsucht rekursiv alle Controls und setzt die MaxLength-Eigenschaft basierend auf den EF Core Datenannotationen.
    /// </summary>
    public static void ApplyMaxLengthFromEntity<T>(this Control container)
    {
        foreach (Control control in container.Controls)
        {
            // Rekursion für verschachtelte Panels, SplitContainer, etc.
            if (control.HasChildren)
            {
                control.ApplyMaxLengthFromEntity<T>();
            }

            // Prüfen, ob es eine TextBox, PaddedTextBox oder MaskedTextBox ist
            if (control is TextBoxBase textBox)
            {
                // Wir holen uns das DataBinding für die "Text"-Eigenschaft
                var textBinding = textBox.DataBindings["Text"];

                if (textBinding != null)
                {
                    // Name der Eigenschaft in der Entity (z.B. "FirstName")
                    var propertyName = textBinding.BindingMemberInfo.BindingField;

                    if (!string.IsNullOrEmpty(propertyName))
                    {
                        var propertyInfo = typeof(T).GetProperty(propertyName);

                        if (propertyInfo != null)
                        {
                            // Suche nach [MaxLength(X)] oder [StringLength(X)]
                            var maxLengthAttr = propertyInfo.GetCustomAttribute<MaxLengthAttribute>();
                            var stringLengthAttr = propertyInfo.GetCustomAttribute<StringLengthAttribute>();

                            // Nimmt den Wert, der gefunden wurde
                            var maxLength = maxLengthAttr?.Length ?? stringLengthAttr?.MaximumLength;

                            if (maxLength.HasValue && maxLength.Value > 0)
                            {
                                // Zuweisung an das UI-Control
                                textBox.MaxLength = maxLength.Value;
                            }
                        }
                    }
                }
            }
        }
    }
}