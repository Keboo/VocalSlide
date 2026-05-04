using Keboo.VocalSlide.Models;

namespace Keboo.VocalSlide.Services;

public interface IPresenterNotesParser
{
    PresenterNotesParseResult Parse(string? rawNotes);
}
