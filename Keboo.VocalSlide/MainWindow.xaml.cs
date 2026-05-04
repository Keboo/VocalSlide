using Keboo.VocalSlide.Dialogs;
using MaterialDesignThemes.Wpf;
using System.Windows.Input;

namespace Keboo.VocalSlide;

/// <summary>
/// Interaction logic for MainWindow.xaml
/// </summary>
public partial class MainWindow
{
    public MainWindow(MainWindowViewModel viewModel)
    {
        DataContext = viewModel;
        InitializeComponent();

        CommandBindings.Add(new CommandBinding(ApplicationCommands.Close, OnClose));

        viewModel.RequestOpenDialog = async dialogId =>
        {
            try
            {
                object content = dialogId switch
                {
                    "settings" => new SettingsDialogView { DataContext = viewModel },
                    "slides" => new SlidesDialogView { DataContext = viewModel },
                    _ => throw new InvalidOperationException($"Unknown dialog identifier: {dialogId}")
                };

                await DialogHost.Show(content, "RootDialog").ConfigureAwait(true);
            }
            catch (InvalidOperationException)
            {
                // A dialog is already open; ignore the request.
            }
        };
    }

    private void OnClose(object sender, ExecutedRoutedEventArgs e)
    {
        Close();
    }
}
