using Keboo.VocalSlide.Models;
using Keboo.VocalSlide.Services;

namespace Keboo.VocalSlide.Tests;

public class SlideAutomationIndexTests
{
    [Test]
    public async Task Build_WhenSlidesContainPromptsAcrossDeck_IndexesOnlyPromptSlidesInOrder()
    {
        PowerPointSlideInfo[] slides =
        [
            new PowerPointSlideInfo(1, "Intro", "Notes", string.Empty),
            new PowerPointSlideInfo(2, "Problem", "Notes", "Switch when I describe the problem."),
            new PowerPointSlideInfo(3, "Background", "Notes", string.Empty),
            new PowerPointSlideInfo(4, "Architecture", "Notes", "Switch when I explain the architecture."),
            new PowerPointSlideInfo(5, "Wrap Up", "Notes", "Switch when I summarize the plan.")
        ];

        IReadOnlyList<PowerPointSlideInfo> index = SlideAutomationIndex.Build(slides);

        await Assert.That(index.Count).IsEqualTo(3);
        await Assert.That(index[0].SlideNumber).IsEqualTo(2);
        await Assert.That(index[1].SlideNumber).IsEqualTo(4);
        await Assert.That(index[2].SlideNumber).IsEqualTo(5);
    }

    [Test]
    public async Task BuildCandidates_WhenCurrentSlideIsInMiddle_ReturnsPromptSlidesBeforeAndAfterIt()
    {
        IReadOnlyList<PowerPointSlideInfo> index =
        [
            new PowerPointSlideInfo(1, "Intro", "Notes", "Switch when I restate the problem."),
            new PowerPointSlideInfo(3, "Architecture", "Notes", "Switch when I explain the current architecture."),
            new PowerPointSlideInfo(5, "Wrap Up", "Notes", "Switch when I summarize the rollout plan.")
        ];

        IReadOnlyList<PowerPointSlideInfo> candidates = SlideAutomationIndex.BuildCandidates(index, 3);

        await Assert.That(candidates.Count).IsEqualTo(2);
        await Assert.That(candidates[0].SlideNumber).IsEqualTo(1);
        await Assert.That(candidates[1].SlideNumber).IsEqualTo(5);
    }

    [Test]
    public async Task FindNearest_WhenSlidesExistOnBothSides_PicksTheClosestIndexedPrompt()
    {
        IReadOnlyList<PowerPointSlideInfo> index =
        [
            new PowerPointSlideInfo(1, "Intro", "Notes", "Prompt"),
            new PowerPointSlideInfo(4, "Architecture", "Notes", "Prompt"),
            new PowerPointSlideInfo(6, "Wrap Up", "Notes", "Prompt")
        ];

        PowerPointSlideInfo? nearest = SlideAutomationIndex.FindNearest(index, 5);

        await Assert.That(nearest).IsNotNull();
        await Assert.That(nearest!.SlideNumber).IsEqualTo(4);
    }
}
