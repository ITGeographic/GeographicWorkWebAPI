using System;
using System.Collections.Generic;
using Microsoft.EntityFrameworkCore;

namespace GeographicDynamic_DAL.Models;

public partial class GeographicDynamicDbContext : DbContext
{
    public GeographicDynamicDbContext()
    {
    }

    public GeographicDynamicDbContext(DbContextOptions<GeographicDynamicDbContext> options)
        : base(options)
    {
    }

    public virtual DbSet<ColumnName> ColumnNames { get; set; }

    public virtual DbSet<DictionariesCodeDefinition> DictionariesCodeDefinitions { get; set; }

    public virtual DbSet<Dictionary> Dictionaries { get; set; }

    public virtual DbSet<Foto> Fotos { get; set; }

    public virtual DbSet<GadanomriliFotoebi> GadanomriliFotoebis { get; set; }

    public virtual DbSet<Qarsafari> Qarsafaris { get; set; }

    public virtual DbSet<QarsafariArqivi> QarsafariArqivis { get; set; }

    public virtual DbSet<QarsafariGrouped> QarsafariGroupeds { get; set; }

    public virtual DbSet<QarsafariGroupedArqivi> QarsafariGroupedArqivis { get; set; }

    public virtual DbSet<VarjisFarti> VarjisFartis { get; set; }

    public virtual DbSet<WindbreakMdb> WindbreakMdbs { get; set; }

    protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
#warning To protect potentially sensitive information in your connection string, you should move it out of source code. You can avoid scaffolding the connection string by using the Name= syntax to read it from configuration - see https://go.microsoft.com/fwlink/?linkid=2131148. For more guidance on storing connection strings, see https://go.microsoft.com/fwlink/?LinkId=723263.
        => optionsBuilder.UseSqlServer("Server=ALEX\\SQLEXPRESS;Database=Geographic_Dynamic_DB;Trusted_Connection=False;User ID=sa;Password=123;TrustServerCertificate=yes");

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        modelBuilder.UseCollation("SQL_Latin1_General_CP1_CI_AS");

        modelBuilder.Entity<ColumnName>(entity =>
        {
            entity.ToTable("ColumnName");

            entity.Property(e => e.Id).HasColumnName("ID");
            entity.Property(e => e.AccessName).HasMaxLength(255);
            entity.Property(e => e.DataType).HasMaxLength(255);
            entity.Property(e => e.ExcelName).HasMaxLength(255);
            entity.Property(e => e.GroupMethod).HasMaxLength(250);
            entity.Property(e => e.Sqlname)
                .HasMaxLength(255)
                .HasColumnName("SQLName");
        });

        modelBuilder.Entity<DictionariesCodeDefinition>(entity =>
        {
            entity.ToTable("DictionariesCodeDefinition");

            entity.Property(e => e.Definition).HasMaxLength(150);
        });

        modelBuilder.Entity<Dictionary>(entity =>
        {
            entity.HasKey(e => e.Id).HasName("PK_TableDictionaries");

            entity.Property(e => e.Id).HasColumnName("ID");
            entity.Property(e => e.Name).HasMaxLength(50);
        });

        modelBuilder.Entity<Foto>(entity =>
        {
            entity.HasKey(e => e.Id).HasName("PK__foto__3214EC27D1B50520");

            entity.ToTable("foto");

            entity.Property(e => e.Id).HasColumnName("ID");
            entity.Property(e => e.FotoN).HasColumnName("fotoN");
        });

        modelBuilder.Entity<GadanomriliFotoebi>(entity =>
        {
            entity.HasKey(e => e.Id).HasName("PK__gadanomr__3214EC27C819CE0A");

            entity.ToTable("gadanomriliFotoebi");

            entity.Property(e => e.Id).HasColumnName("ID");
            entity.Property(e => e.LiterId).HasColumnName("Liter_ID");
            entity.Property(e => e.PhotoDate).HasMaxLength(500);
            entity.Property(e => e.PhotoN).HasColumnName("Photo_N");
            entity.Property(e => e.UniqId)
                .HasMaxLength(255)
                .HasColumnName("UNIQ_ID");
        });

        modelBuilder.Entity<Qarsafari>(entity =>
        {
            entity.HasKey(e => e.Id).HasName("PK__qarsafar__3214EC27DF13B916");

            entity.ToTable("qarsafari");

            entity.Property(e => e.AdmMun).HasMaxLength(255);
            entity.Property(e => e.CadCod).HasMaxLength(255);
            entity.Property(e => e.CityTownVillage).HasMaxLength(255);
            entity.Property(e => e.Company).HasMaxLength(255);
            entity.Property(e => e.DaTe1).HasMaxLength(255);
            entity.Property(e => e.Date).HasMaxLength(255);
            entity.Property(e => e.Date2).HasMaxLength(255);
            entity.Property(e => e.Date3).HasMaxLength(255);
            entity.Property(e => e.EtapiId).HasColumnName("EtapiID");
            entity.Property(e => e.FieldOperator).HasMaxLength(255);
            entity.Property(e => e.GisOperator).HasMaxLength(255);
            entity.Property(e => e.IsUniqLiterNull).HasMaxLength(50);
            entity.Property(e => e.LandFieldOperator).HasMaxLength(255);
            entity.Property(e => e.LandGisOperator).HasMaxLength(255);
            entity.Property(e => e.LegalPerson).HasMaxLength(255);
            entity.Property(e => e.Municipality).HasMaxLength(255);
            entity.Property(e => e.Note).HasMaxLength(255);
            entity.Property(e => e.Note1).HasMaxLength(255);
            entity.Property(e => e.Note11).HasMaxLength(255);
            entity.Property(e => e.OverlapCadCode).HasMaxLength(255);
            entity.Property(e => e.Owner).HasMaxLength(255);
            entity.Property(e => e.Owners).HasMaxLength(255);
            entity.Property(e => e.ProjectId).HasColumnName("ProjectID");
            entity.Property(e => e.Region).HasMaxLength(255);
            entity.Property(e => e.Rrr).HasMaxLength(50);
            entity.Property(e => e.Sakutreba).HasMaxLength(50);
            entity.Property(e => e.ShapeArea).HasMaxLength(255);
            entity.Property(e => e.Uid).HasMaxLength(50);
            entity.Property(e => e.UniqIdNew).HasMaxLength(255);
            entity.Property(e => e.WoodyPlantPercent).HasMaxLength(255);
            entity.Property(e => e.WoodyPlantSpecies).HasMaxLength(255);
        });

        modelBuilder.Entity<QarsafariArqivi>(entity =>
        {
            entity
                .HasNoKey()
                .ToTable("qarsafariArqivi");

            entity.Property(e => e.AdmMun).HasMaxLength(255);
            entity.Property(e => e.CadCod).HasMaxLength(255);
            entity.Property(e => e.CityTownVillage).HasMaxLength(255);
            entity.Property(e => e.Company).HasMaxLength(255);
            entity.Property(e => e.DaTe1).HasMaxLength(255);
            entity.Property(e => e.Date).HasMaxLength(255);
            entity.Property(e => e.Date2).HasMaxLength(255);
            entity.Property(e => e.Date3).HasMaxLength(255);
            entity.Property(e => e.EtapiId).HasColumnName("EtapiID");
            entity.Property(e => e.FieldOperator).HasMaxLength(255);
            entity.Property(e => e.GisOperator).HasMaxLength(255);
            entity.Property(e => e.IsUniqLiterNull).HasMaxLength(50);
            entity.Property(e => e.LandFieldOperator).HasMaxLength(255);
            entity.Property(e => e.LandGisOperator).HasMaxLength(255);
            entity.Property(e => e.LegalPerson).HasMaxLength(255);
            entity.Property(e => e.Municipality).HasMaxLength(255);
            entity.Property(e => e.Note).HasMaxLength(255);
            entity.Property(e => e.Note1).HasMaxLength(255);
            entity.Property(e => e.Note11).HasMaxLength(255);
            entity.Property(e => e.OverlapCadCode).HasMaxLength(255);
            entity.Property(e => e.Owner).HasMaxLength(255);
            entity.Property(e => e.Owners).HasMaxLength(255);
            entity.Property(e => e.ProjectId).HasColumnName("ProjectID");
            entity.Property(e => e.Region).HasMaxLength(255);
            entity.Property(e => e.Rrr).HasMaxLength(50);
            entity.Property(e => e.Sakutreba).HasMaxLength(50);
            entity.Property(e => e.ShapeArea).HasMaxLength(255);
            entity.Property(e => e.Uid).HasMaxLength(50);
            entity.Property(e => e.UniqIdNew).HasMaxLength(255);
            entity.Property(e => e.WoodyPlantPercent).HasMaxLength(255);
            entity.Property(e => e.WoodyPlantSpecies).HasMaxLength(255);
        });

        modelBuilder.Entity<QarsafariGrouped>(entity =>
        {
            entity.HasKey(e => e.Id).HasName("PK__qarsafar__E8E7955E4456C1CE");

            entity.ToTable("qarsafariGrouped");

            entity.Property(e => e.Id).HasColumnName("ID");
            entity.Property(e => e.AdmMun).HasMaxLength(255);
            entity.Property(e => e.CadCod).HasMaxLength(255);
            entity.Property(e => e.CityTownVillage).HasMaxLength(255);
            entity.Property(e => e.Company).HasMaxLength(255);
            entity.Property(e => e.DaTe1).HasMaxLength(255);
            entity.Property(e => e.Date).HasMaxLength(255);
            entity.Property(e => e.Date2).HasMaxLength(255);
            entity.Property(e => e.Date3).HasMaxLength(255);
            entity.Property(e => e.EtapiId).HasColumnName("EtapiID");
            entity.Property(e => e.FieldOperator).HasMaxLength(255);
            entity.Property(e => e.GisOperator).HasMaxLength(255);
            entity.Property(e => e.LandFieldOperator).HasMaxLength(255);
            entity.Property(e => e.LandGisOperator).HasMaxLength(255);
            entity.Property(e => e.LegalPerson).HasMaxLength(255);
            entity.Property(e => e.Municipality).HasMaxLength(255);
            entity.Property(e => e.Note).HasMaxLength(255);
            entity.Property(e => e.Note1).HasMaxLength(255);
            entity.Property(e => e.Note11).HasMaxLength(255);
            entity.Property(e => e.OverlapCadCode).HasMaxLength(255);
            entity.Property(e => e.Owner).HasMaxLength(255);
            entity.Property(e => e.Owners).HasMaxLength(255);
            entity.Property(e => e.ProjectId).HasColumnName("ProjectID");
            entity.Property(e => e.Region).HasMaxLength(255);
            entity.Property(e => e.Sakutreba).HasMaxLength(50);
            entity.Property(e => e.Uid).HasMaxLength(50);
        });

        modelBuilder.Entity<QarsafariGroupedArqivi>(entity =>
        {
            entity
                .HasNoKey()
                .ToTable("qarsafariGroupedArqivi");

            entity.Property(e => e.AdmMun).HasMaxLength(255);
            entity.Property(e => e.CadCod).HasMaxLength(255);
            entity.Property(e => e.CityTownVillage).HasMaxLength(255);
            entity.Property(e => e.Company).HasMaxLength(255);
            entity.Property(e => e.DaTe1).HasMaxLength(255);
            entity.Property(e => e.Date).HasMaxLength(255);
            entity.Property(e => e.Date2).HasMaxLength(255);
            entity.Property(e => e.Date3).HasMaxLength(255);
            entity.Property(e => e.EtapiId).HasColumnName("EtapiID");
            entity.Property(e => e.FieldOperator).HasMaxLength(255);
            entity.Property(e => e.GisOperator).HasMaxLength(255);
            entity.Property(e => e.Id).HasColumnName("ID");
            entity.Property(e => e.LandFieldOperator).HasMaxLength(255);
            entity.Property(e => e.LandGisOperator).HasMaxLength(255);
            entity.Property(e => e.LegalPerson).HasMaxLength(255);
            entity.Property(e => e.Municipality).HasMaxLength(255);
            entity.Property(e => e.Note).HasMaxLength(255);
            entity.Property(e => e.Note1).HasMaxLength(255);
            entity.Property(e => e.Note11).HasMaxLength(255);
            entity.Property(e => e.OverlapCadCode).HasMaxLength(255);
            entity.Property(e => e.Owner).HasMaxLength(255);
            entity.Property(e => e.Owners).HasMaxLength(255);
            entity.Property(e => e.ProjectId).HasColumnName("ProjectID");
            entity.Property(e => e.Region).HasMaxLength(255);
            entity.Property(e => e.Sakutreba).HasMaxLength(50);
            entity.Property(e => e.Uid).HasMaxLength(50);
        });

        modelBuilder.Entity<VarjisFarti>(entity =>
        {
            entity.ToTable("VarjisFarti");

            entity.Property(e => e.Id).HasColumnName("ID");
            entity.Property(e => e.AreaNameId).HasColumnName("AreaNameID");
            entity.Property(e => e.SaxeobaId).HasColumnName("SaxeobaID");
            entity.Property(e => e.VarjisFarti1).HasColumnName("varjisFarti");

            entity.HasOne(d => d.Saxeoba).WithMany(p => p.VarjisFartis)
                .HasForeignKey(d => d.SaxeobaId)
                .HasConstraintName("FK_VarjisFarti_Dictionaries1");
        });

        modelBuilder.Entity<WindbreakMdb>(entity =>
        {
            entity.HasKey(e => e.Id).HasName("PK__Windbrea__3214EC076D389C12");

            entity.ToTable("WindbreakMDB");

            entity.Property(e => e.AdmMun).HasMaxLength(255);
            entity.Property(e => e.Aleks).HasColumnName("ALEKS");
            entity.Property(e => e.CadCod).HasMaxLength(255);
            entity.Property(e => e.CityTownVillage).HasMaxLength(255);
            entity.Property(e => e.Company).HasMaxLength(255);
            entity.Property(e => e.DaTe1).HasMaxLength(255);
            entity.Property(e => e.Date).HasMaxLength(255);
            entity.Property(e => e.Date2).HasMaxLength(255);
            entity.Property(e => e.Date3).HasMaxLength(255);
            entity.Property(e => e.FieldOperator).HasMaxLength(255);
            entity.Property(e => e.GisOperator).HasMaxLength(255);
            entity.Property(e => e.IsUniqLiterNull).HasMaxLength(50);
            entity.Property(e => e.LandFieldOperator).HasMaxLength(255);
            entity.Property(e => e.LandGisOperator).HasMaxLength(255);
            entity.Property(e => e.LegalPerson).HasMaxLength(255);
            entity.Property(e => e.Municipality).HasMaxLength(255);
            entity.Property(e => e.Note).HasMaxLength(255);
            entity.Property(e => e.Note1).HasMaxLength(255);
            entity.Property(e => e.Note11).HasMaxLength(255);
            entity.Property(e => e.OverlapCadCode).HasMaxLength(255);
            entity.Property(e => e.Owner).HasMaxLength(255);
            entity.Property(e => e.Owners).HasMaxLength(255);
            entity.Property(e => e.PhotoN).HasMaxLength(255);
            entity.Property(e => e.Region).HasMaxLength(255);
            entity.Property(e => e.Rrr).HasMaxLength(50);
            entity.Property(e => e.Sakutreba).HasMaxLength(50);
            entity.Property(e => e.ShapeArea).HasMaxLength(255);
            entity.Property(e => e.Test)
                .HasMaxLength(255)
                .HasColumnName("TEST");
            entity.Property(e => e.Test1)
                .HasMaxLength(255)
                .HasColumnName("TEST1");
            entity.Property(e => e.Test2)
                .HasMaxLength(255)
                .HasColumnName("TEST2");
            entity.Property(e => e.Test3).HasColumnName("TEST3");
            entity.Property(e => e.Uid).HasMaxLength(50);
            entity.Property(e => e.UniqIdNew).HasMaxLength(255);
            entity.Property(e => e.WoodyPlantPercent).HasMaxLength(255);
            entity.Property(e => e.WoodyPlantSpecies).HasMaxLength(255);
        });

        OnModelCreatingPartial(modelBuilder);
    }

    partial void OnModelCreatingPartial(ModelBuilder modelBuilder);
}
