using Microsoft.EntityFrameworkCore;

namespace scorerlauncher;

public class ScoreContext : DbContext
{
    public DbSet<OperatorAccount> Operators => Set<OperatorAccount>();

    public string DbPath { get; }

    public ScoreContext()
    {
        DbPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "operators.db");
    }

    protected override void OnConfiguring(DbContextOptionsBuilder options)
        => options.UseSqlite($"Data Source={DbPath}");
}