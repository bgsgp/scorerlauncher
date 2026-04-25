using System.ComponentModel.DataAnnotations;

namespace scorerlauncher;

public class OperatorAccount
{
    [Key]
    public int Id { get; set; }

    [Required]
    [MaxLength(50)]
    public string Username { get; set; } = string.Empty;

    [Required]
    public string PasswordHash { get; set; } = string.Empty;

    public bool IsMaster { get; set; } = false;
    public bool IsAdmin { get; set; } = false;

    public DayOfWeek? AssignedDay { get; set; }
}