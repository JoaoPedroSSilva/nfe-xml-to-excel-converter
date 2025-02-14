
namespace NFeConverter {
    public class Program {
        public static void Main(string[] args) {
            var builder = WebApplication.CreateBuilder(args);

            // Add services to the container.

            builder.Services.AddCors(options => {
                options.AddPolicy("AllowLocalhost",
                    policy => {
                        policy.WithOrigins("http://127.0.0.1:5500", "https://127.0.0.1:5500")
                        .AllowAnyMethod()
                        .AllowAnyHeader()
                        .AllowCredentials();

                    });
            });

            builder.Services.AddControllers();
            // Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
            builder.Services.AddEndpointsApiExplorer();
            builder.Services.AddSwaggerGen();

            var app = builder.Build();

            app.UseCors("AllowLocalhost");

            // Configure the HTTP request pipeline.
            if (app.Environment.IsDevelopment()) {
                app.UseSwagger();
                app.UseSwaggerUI();
            }

            app.UseHttpsRedirection();

            app.UseAuthorization();


            app.MapControllers();

            app.Run();
        }
    }
}
