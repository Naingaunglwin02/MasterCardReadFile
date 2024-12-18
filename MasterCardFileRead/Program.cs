

using MasterCardFileRead.Services;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.

builder.Services.AddControllers();

builder.Services.AddTransient<EcommerceTransaction>();
builder.Services.AddTransient<OtherTransaction>();
<<<<<<< HEAD
builder.Services.AddTransient<IssuingTransaction>();
builder.Services.AddTransient<RejectTransaction>();
=======
builder.Services.AddTransient<PosTransaction>();
>>>>>>> features/hhz
// Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

var app = builder.Build();

//Console.WriteLine("hello tesing hhz");

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseHttpsRedirection();

app.UseAuthorization();

app.MapControllers();

app.Run();
