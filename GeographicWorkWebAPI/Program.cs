using BotReestriClassLibrary.Interface;
using BotReestriClassLibrary.Repository;
using GeographicDynamic_DAL.Configurations;
using GeographicDynamic_DAL.DTOs.Windbreak;
using GeographicDynamic_DAL.Interface;
using GeographicDynamic_DAL.Models;
using GeographicDynamic_DAL.Repository;
using Microsoft.EntityFrameworkCore;
//using Microsoft.Extensions.Configuration;


var builder = WebApplication.CreateBuilder(args);

var MyAllowSpecificOrigins = "_myAllowSpecificOrigins";
// Add services to the container.

builder.Services.AddDbContext<GeographicDynamicDbContext>(options => options.UseSqlServer(builder.Configuration.GetConnectionString("Geographic_Dynamic_Connection")));
//var configuration = new ConfigurationBuilder()
//    .SetBasePath(builder.Environment.ContentRootPath)
//    .AddJsonFile("appsettings.json")
//    .Build(); // Initialize the configuration object



builder.Services.AddControllers();
// Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

builder.Services.AddCors(options =>
{
    options.AddPolicy(name: MyAllowSpecificOrigins,
        policy =>
        {
            policy.AllowAnyOrigin()
                .AllowAnyHeader()
                .AllowAnyMethod();
        });
});
builder.Services.AddTransient<DictionaryDTO>();
builder.Services.AddTransient<IVarjisFarti, VarjisFartiRepository>();
builder.Services.AddTransient<IAero, AeroGadagebaRepository>();
builder.Services.AddTransient<IAeroWithObjectId, AeroGadagebaWithObjectIdRepository>();
builder.Services.AddTransient<IWindbreak, WindbreakRepository>();
builder.Services.AddTransient<IColumnName, ColumnNameRepository>();
builder.Services.AddTransient<IChromeBot, ChromeBotRepository>();
builder.Services.AddAutoMapper(typeof(MapperConfig));
var app = builder.Build();

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();

}
app.UseCors(MyAllowSpecificOrigins);

app.UseHttpsRedirection();

app.UseAuthorization();

app.MapControllers();

app.Run();
