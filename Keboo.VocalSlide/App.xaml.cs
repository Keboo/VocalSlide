using Keboo.VocalSlide.Services;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using System.Windows;
using System.Windows.Threading;

namespace Keboo.VocalSlide;

public partial class App : Application
{
    private readonly IHost _host;

    public App()
    {
        _host = CreateHostBuilder([]).Build();
    }

    [STAThread]
    private static void Main(string[] args)
    {
        App app = new();
        app.InitializeComponent();
        app.Run();
    }

    protected override async void OnStartup(StartupEventArgs e)
    {
        await _host.StartAsync().ConfigureAwait(true);

        MainWindow = _host.Services.GetRequiredService<MainWindow>();
        MainWindow.Show();

        base.OnStartup(e);
    }

    protected override async void OnExit(ExitEventArgs e)
    {
        await _host.StopAsync().ConfigureAwait(true);
        _host.Dispose();
        base.OnExit(e);
    }

    public static IHostBuilder CreateHostBuilder(string[] args)
    {
        Dispatcher uiDispatcher = Dispatcher.CurrentDispatcher;

        return Host.CreateDefaultBuilder(args)
            .ConfigureAppConfiguration((_, configurationBuilder)
                => configurationBuilder.AddUserSecrets(typeof(App).Assembly))
            .ConfigureServices((_, services) =>
            {
                services.AddSingleton(uiDispatcher);

                services.AddSingleton<MainWindow>();
                services.AddSingleton<MainWindowViewModel>();

                services.AddSingleton<IPresenterNotesParser, PresenterNotesParser>();
                services.AddSingleton<IPowerPointSessionService, PowerPointSessionService>();
                services.AddSingleton<ILocalTranscriptionService, WhisperTranscriptionService>();
                services.AddSingleton<ILocalSlideEvaluationService, OllamaSlideEvaluationService>();
                services.AddSingleton<IModelDownloadService, ModelDownloadService>();
                services.AddSingleton<AutoAdvancePolicy>();
            });
    }
}
