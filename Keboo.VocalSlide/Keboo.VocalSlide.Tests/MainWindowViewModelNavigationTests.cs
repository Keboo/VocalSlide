using Keboo.VocalSlide.Models;
using Keboo.VocalSlide.Services;
using System.IO;
using System.Windows.Threading;

namespace Keboo.VocalSlide.Tests;

public class MainWindowViewModelNavigationTests
{
    [Test]
    public async Task PreviousSlideCommand_CanExecute_WhenCurrentSlideIsPastFirstSlide()
    {
        MainWindowViewModel viewModel = CreateViewModel();
        viewModel.IsPowerPointConnected = true;
        viewModel.IsSlideShowRunning = true;
        viewModel.CurrentSlideNumber = 2;

        await Assert.That(viewModel.PreviousSlideCommand.CanExecute(null)).IsTrue();
    }

    [Test]
    public async Task PreviousSlideCommand_CannotExecute_OnFirstSlide()
    {
        MainWindowViewModel viewModel = CreateViewModel();
        viewModel.IsPowerPointConnected = true;
        viewModel.IsSlideShowRunning = true;
        viewModel.CurrentSlideNumber = 1;

        await Assert.That(viewModel.PreviousSlideCommand.CanExecute(null)).IsFalse();
    }

    [Test]
    public async Task GoToSelectedSlideCommand_CanExecute_WhenSelectionDiffersFromCurrentSlide()
    {
        MainWindowViewModel viewModel = CreateViewModel();
        viewModel.IsPowerPointConnected = true;
        viewModel.IsSlideShowRunning = true;
        viewModel.CurrentSlideNumber = 1;
        viewModel.SelectedSlide = new PowerPointSlideInfo(4, "Slide 4", "Notes", "Prompt");

        await Assert.That(viewModel.GoToSelectedSlideCommand.CanExecute(null)).IsTrue();
    }

    [Test]
    public async Task GoToSelectedSlideCommand_CannotExecute_WhenSelectionMatchesCurrentSlide()
    {
        MainWindowViewModel viewModel = CreateViewModel();
        viewModel.IsPowerPointConnected = true;
        viewModel.IsSlideShowRunning = true;
        viewModel.CurrentSlideNumber = 4;
        viewModel.SelectedSlide = new PowerPointSlideInfo(4, "Slide 4", "Notes", "Prompt");

        await Assert.That(viewModel.GoToSelectedSlideCommand.CanExecute(null)).IsFalse();
    }

    private static MainWindowViewModel CreateViewModel()
    {
        Mock<IPowerPointSessionService> powerPointSessionService = new();
        Mock<ILocalTranscriptionService> transcriptionService = new();
        Mock<ILocalSlideEvaluationService> slideEvaluationService = new();
        Mock<IModelDownloadService> modelDownloadService = new();

        string modelStorageDirectory = Path.Combine(Path.GetTempPath(), "Keboo.VocalSlide.Tests");

        modelDownloadService.SetupGet(service => service.StorageDirectory).Returns(modelStorageDirectory);
        modelDownloadService
            .Setup(service => service.GetStoredFilePath(It.IsAny<string>()))
            .Returns((string fileName) => Path.Combine(modelStorageDirectory, fileName));
        modelDownloadService
            .Setup(service => service.GetStoredFilePathFromUrl(It.IsAny<string>()))
            .Returns((string _) => Path.Combine(modelStorageDirectory, "model.gguf"));

        return new MainWindowViewModel(
            powerPointSessionService.Object,
            transcriptionService.Object,
            slideEvaluationService.Object,
            modelDownloadService.Object,
            new AutoAdvancePolicy(),
            Dispatcher.CurrentDispatcher);
    }
}
