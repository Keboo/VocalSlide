using Keboo.VocalSlide.Models;

namespace Keboo.VocalSlide.Services;

public interface IPowerPointSessionService
{
    Task<PresentationSessionSnapshot> ConnectAsync(CancellationToken cancellationToken = default);

    Task<SlideShowState> GetSlideShowStateAsync(CancellationToken cancellationToken = default);

    Task AdvanceToNextSlideAsync(CancellationToken cancellationToken = default);

    Task GoToSlideAsync(int slideNumber, CancellationToken cancellationToken = default);
}
