using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Adressen.cls;

[Table("Fotos")]
public class Foto
{
    [Key]
    [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
    public int Id { get; set; }
    public byte[]? Fotodaten { get; set; }
    public int AdressId { get; set; }

    [ForeignKey("AdressId")]
    public virtual Adresse Adresse { get; set; } = null!;
}
