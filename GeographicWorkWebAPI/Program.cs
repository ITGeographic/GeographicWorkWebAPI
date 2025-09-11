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

// Services
builder.Services.AddDbContext<GeographicDynamicDbContext>(options =>
    options.UseSqlServer(builder.Configuration.GetConnectionString("Geographic_Dynamic_Connection")));
builder.Services.AddControllers(options =>
{
    options.Filters.Add(new Microsoft.AspNetCore.Mvc.Authorization.AllowAnonymousFilter());
});
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();
builder.Services.AddCors(options =>
{
    options.AddPolicy("_myAllowSpecificOrigins",
        policy => policy.AllowAnyOrigin().AllowAnyHeader().AllowAnyMethod());
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

// Base path (subfolder)
app.UsePathBase("/GeographicWorkBack");

// Middleware
app.UseCors("_myAllowSpecificOrigins");
//app.UseHttpsRedirection();
app.UseAuthorization();

// Swagger (ყველა environment-ში)
app.UseSwagger();
app.UseSwaggerUI(c =>
{
    c.SwaggerEndpoint("/GeographicWorkBack/swagger/v1/swagger.json", "My API V1");
    c.RoutePrefix = "swagger";
});

app.MapControllers();

app.Run();
