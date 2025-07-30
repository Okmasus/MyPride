using CvParser.DataAccess.Extensions;
using CvParser.DataAccess.Models;
using CvParser.DataAccess.Models.EducationsModels;
using CvParser.DataAccess.Models.Google;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Metadata.Builders;
using Microsoft.Extensions.Logging;

namespace CvParser.DataAccess;

/// <summary>
/// Контекст данных для взаимодействия с базой данных.
/// </summary>
public sealed class DataContext : DbContext
{
    private readonly ILogger<DataContext> _logger;

    public DataContext(DbContextOptions<DataContext> context, ILogger<DataContext> logger) : base(context)
    {
        _logger = logger;
    }

    /// <summary>
    /// Переопределён для возможности установки значений по-умолчанию и журналирования
    /// </summary>
    public override Task<int> SaveChangesAsync(CancellationToken cancellationToken = default)
    {
        foreach (var entry in ChangeTracker
            .Entries<ResumeModel>()
            .Where(entry => entry.State == EntityState.Added || entry.State == EntityState.Modified))
        {
            entry.Entity.SetDefaultOrganizationInfoIfEmpty();

            if (string.IsNullOrWhiteSpace(entry.Entity.Name))
            {
                _logger.LogWarning(
                    "Resume was saved with empty fullname.\n" +
                    "Resume ID:\t{id}\n" +
                    "Time:\t\t{time}\n" +
                    "Entity state:\t{state}\n",
                    entry.Entity.ID, DateTime.Now, entry.State);
            }
        }

        return base.SaveChangesAsync(cancellationToken);
    }

    private static void ConfigureCodes<TCode>(EntityTypeBuilder<TCode> buildAction)
        where TCode : ConfirmationToken
    {
        buildAction.HasKey(code => code.Code);
        buildAction.HasIndex(code => code.RelatedUserId).IsUnique();
    }

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        modelBuilder.ApplyConfigurationsFromAssembly(typeof(DataContext).Assembly);
    }

    protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
    {
        optionsBuilder.EnableDetailedErrors();
        optionsBuilder.EnableSensitiveDataLogging(true);
    }

    public DbSet<ResumeModel> Resumes { get; set; } = null!;
    public DbSet<LimitBalance> LimitBalances { get; set; } = null!;
    public DbSet<OpenContact> OpenContacts { get; set; } = null!;
    public DbSet<MessageStatistics> MessageStatistics { get; set; } = null!;
    public DbSet<CallsToCandidatesStatistics> CallsToCandidatesStatistics { get; set; } = null!;
    public DbSet<CvStatistics> CvStatistics { get; set; } = null!;
    public DbSet<AddingCandidatesStatistics> AddingCandidatesStatistics { get; set; } = null!;
    public DbSet<DealStatusStatistics> DealStatusStatistics { get; set; } = null!;
    public DbSet<UserDealStatusHistory> UserDealStatusHistories { get; set; } = null!;
    public DbSet<GettingUserDate> GettingUserDatas { get; set; } = null!;
    public DbSet<HHQuery> HHQueries { get; set; } = null!;
    public DbSet<KeySkill> KeySkill { get; set; }
    public DbSet<WorkExperienceModel> WorkExperiences { get; set; } = null!;
    public DbSet<HHQueryResume> HHQueryResumes { get; set; } = null!;
    public DbSet<EntityModel> Entities { get; set; } = null!;
    public DbSet<StatModel> Stats { get; set; } = null!;
    public DbSet<HeadHunterClientErrors> HeadHunterClientErrors { get; set; } = null!;
    public DbSet<RelevantWorkExperienceKeyword> RelevantWorkExperienceKeywords { get; set; } = null!;
    public DbSet<GoogleBenchSourceModel> GoogleBenchSources { get; set; } = null!;
    public DbSet<GoogleSheetBenchSourceRangeModel> GoogleSheetsBenchSourcesRanges { get; set; } = null!;
    public DbSet<GoogleBenchSourceOrganizationModel> GoogleBenchSourceOrganizations { get; set; } = null!;
    public DbSet<OrganizationWatermark> OrganizationWatermarks { get; set; } = null!;
    public DbSet<ResumeFromBlockParsing> ResumesFromBlockParsing { get; set; } = null!;
    public DbSet<CityModel> Cities { get; set; } = null!;
    public DbSet<CountryModel> Countries { get; set; } = null!;
    public DbSet<OrganizationDailyLimit> OrganizationDailyLimits { get; set; } = null!;

}