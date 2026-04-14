using Keboo.VocalSlide.Infrastructure;
using Keboo.VocalSlide.Models;
using System.Text;
using System.Windows.Threading;

namespace Keboo.VocalSlide.Services;

public sealed class PowerPointSessionService : IPowerPointSessionService
{
    private const int MsoTrue = -1;
    private const int MsoPlaceholderShapeType = 14;
    private const int PlaceholderBody = 2;
    private const int PlaceholderVerticalBody = 6;

    private readonly IPresenterNotesParser _presenterNotesParser;
    private readonly Dispatcher _dispatcher;

    public PowerPointSessionService(IPresenterNotesParser presenterNotesParser, Dispatcher dispatcher)
    {
        _presenterNotesParser = presenterNotesParser;
        _dispatcher = dispatcher;
    }

    public Task<PresentationSessionSnapshot> ConnectAsync(CancellationToken cancellationToken = default) =>
        _dispatcher.InvokeAsync(CreateSnapshotCore, DispatcherPriority.Background, cancellationToken).Task;

    public Task<SlideShowState> GetSlideShowStateAsync(CancellationToken cancellationToken = default) =>
        _dispatcher.InvokeAsync(CreateSlideShowStateCore, DispatcherPriority.Background, cancellationToken).Task;

    public Task AdvanceToNextSlideAsync(CancellationToken cancellationToken = default) =>
        _dispatcher.InvokeAsync(AdvanceToNextSlideCore, DispatcherPriority.Background, cancellationToken).Task;

    public Task GoToSlideAsync(int slideNumber, CancellationToken cancellationToken = default) =>
        _dispatcher.InvokeAsync(() => GoToSlideCore(slideNumber), DispatcherPriority.Background, cancellationToken).Task;

    private PresentationSessionSnapshot CreateSnapshotCore()
    {
        object? application = null;
        object? presentation = null;
        object? slides = null;
        object? slide = null;

        try
        {
            application = GetApplication();
            presentation = GetPresentation(application);
            dynamic presentationDynamic = presentation;
            slides = presentationDynamic.Slides;
            dynamic slidesDynamic = slides;

            int slideCount = Convert.ToInt32(slidesDynamic.Count);
            List<PowerPointSlideInfo> slideInfos = new(slideCount);

            for (int index = 1; index <= slideCount; index++)
            {
                slide = slidesDynamic[index];
                dynamic slideDynamic = slide;

                string notesText = ReadNotesText(slide);
                PresenterNotesParseResult parsedNotes = _presenterNotesParser.Parse(notesText);

                slideInfos.Add(new PowerPointSlideInfo(
                    Convert.ToInt32(slideDynamic.SlideNumber),
                    ReadTitle(slide),
                    parsedNotes.PresenterNotes,
                    parsedNotes.AutomationPrompt));

                ComInteropHelper.ReleaseIfNeeded(slide);
                slide = null;
            }

            return new PresentationSessionSnapshot(
                ((string?)presentationDynamic.Name)?.Trim() ?? "PowerPoint Presentation",
                CreateSlideShowStateCore(application),
                slideInfos);
        }
        finally
        {
            ComInteropHelper.ReleaseIfNeeded(slide);
            ComInteropHelper.ReleaseIfNeeded(slides);
            ComInteropHelper.ReleaseIfNeeded(presentation);
            ComInteropHelper.ReleaseIfNeeded(application);
        }
    }

    private SlideShowState CreateSlideShowStateCore()
    {
        object? application = null;

        try
        {
            application = GetApplication();
            return CreateSlideShowStateCore(application);
        }
        finally
        {
            ComInteropHelper.ReleaseIfNeeded(application);
        }
    }

    private static SlideShowState CreateSlideShowStateCore(object application)
    {
        object? slideShowWindow = null;
        object? slideShowView = null;

        try
        {
            dynamic applicationDynamic = application;
            dynamic slideShowWindows = applicationDynamic.SlideShowWindows;
            if (Convert.ToInt32(slideShowWindows.Count) <= 0)
            {
                return new SlideShowState(false, 0);
            }

            slideShowWindow = slideShowWindows[1];
            dynamic slideShowWindowDynamic = slideShowWindow;
            slideShowView = slideShowWindowDynamic.View;
            dynamic slideShowViewDynamic = slideShowView;

            return new SlideShowState(true, Convert.ToInt32(slideShowViewDynamic.CurrentShowPosition));
        }
        finally
        {
            ComInteropHelper.ReleaseIfNeeded(slideShowView);
            ComInteropHelper.ReleaseIfNeeded(slideShowWindow);
        }
    }

    private void AdvanceToNextSlideCore()
    {
        object? application = null;
        object? slideShowWindow = null;
        object? slideShowView = null;

        try
        {
            application = GetApplication();
            slideShowWindow = GetSlideShowWindow(application);
            dynamic slideShowWindowDynamic = slideShowWindow;
            slideShowView = slideShowWindowDynamic.View;
            dynamic slideShowViewDynamic = slideShowView;
            slideShowViewDynamic.Next();
        }
        finally
        {
            ComInteropHelper.ReleaseIfNeeded(slideShowView);
            ComInteropHelper.ReleaseIfNeeded(slideShowWindow);
            ComInteropHelper.ReleaseIfNeeded(application);
        }
    }

    private void GoToSlideCore(int slideNumber)
    {
        object? application = null;
        object? slideShowWindow = null;
        object? slideShowView = null;

        try
        {
            application = GetApplication();
            slideShowWindow = GetSlideShowWindow(application);
            dynamic slideShowWindowDynamic = slideShowWindow;
            slideShowView = slideShowWindowDynamic.View;
            dynamic slideShowViewDynamic = slideShowView;
            slideShowViewDynamic.GotoSlide(slideNumber);
        }
        finally
        {
            ComInteropHelper.ReleaseIfNeeded(slideShowView);
            ComInteropHelper.ReleaseIfNeeded(slideShowWindow);
            ComInteropHelper.ReleaseIfNeeded(application);
        }
    }

    private static object GetApplication()
    {
        return ComInteropHelper.GetRunningObject("PowerPoint.Application");
    }

    private static object GetPresentation(object application)
    {
        dynamic applicationDynamic = application;
        dynamic presentations = applicationDynamic.Presentations;

        if (Convert.ToInt32(presentations.Count) <= 0)
        {
            throw new InvalidOperationException("PowerPoint is running, but no presentation is open.");
        }

        dynamic slideShowWindows = applicationDynamic.SlideShowWindows;
        if (Convert.ToInt32(slideShowWindows.Count) > 0)
        {
            object? slideShowWindow = null;
            try
            {
                slideShowWindow = slideShowWindows[1];
                dynamic slideShowWindowDynamic = slideShowWindow;
                return slideShowWindowDynamic.Presentation;
            }
            finally
            {
                ComInteropHelper.ReleaseIfNeeded(slideShowWindow);
            }
        }

        object? activePresentation = null;
        try
        {
            activePresentation = applicationDynamic.ActivePresentation;
        }
        catch
        {
            // Fall back to the first open presentation below.
        }

        return activePresentation ?? presentations[1];
    }

    private static object GetSlideShowWindow(object application)
    {
        dynamic applicationDynamic = application;
        dynamic slideShowWindows = applicationDynamic.SlideShowWindows;
        if (Convert.ToInt32(slideShowWindows.Count) <= 0)
        {
            throw new InvalidOperationException("PowerPoint is connected, but the slide show is not running.");
        }

        return slideShowWindows[1];
    }

    private static string ReadTitle(object slide)
    {
        object? titleShape = null;

        try
        {
            dynamic slideDynamic = slide;
            titleShape = slideDynamic.Shapes.Title;
            if (titleShape is not null)
            {
                dynamic titleShapeDynamic = titleShape;
                if (Convert.ToInt32(titleShapeDynamic.HasTextFrame) == MsoTrue &&
                    Convert.ToInt32(titleShapeDynamic.TextFrame.HasText) == MsoTrue)
                {
                    return ((string?)titleShapeDynamic.TextFrame.TextRange.Text)?.Trim() ?? GetSlideName(slide);
                }
            }
        }
        catch
        {
            // Fall back to the PowerPoint slide name when no title placeholder exists.
        }
        finally
        {
            ComInteropHelper.ReleaseIfNeeded(titleShape);
        }

        return GetSlideName(slide);
    }

    private static string ReadNotesText(object slide)
    {
        object? notesPage = null;
        object? shapes = null;
        object? shape = null;

        try
        {
            dynamic slideDynamic = slide;
            notesPage = slideDynamic.NotesPage;
            dynamic notesPageDynamic = notesPage;
            shapes = notesPageDynamic.Shapes;
            dynamic shapesDynamic = shapes;

            StringBuilder builder = new();
            HashSet<string> seenText = new(StringComparer.Ordinal);
            int shapeCount = Convert.ToInt32(shapesDynamic.Count);

            for (int index = 1; index <= shapeCount; index++)
            {
                shape = shapesDynamic[index];
                string text = ReadNotesShapeText(shape);
                if (!string.IsNullOrWhiteSpace(text) && seenText.Add(text))
                {
                    if (builder.Length > 0)
                    {
                        builder.AppendLine();
                    }

                    builder.Append(text);
                }

                ComInteropHelper.ReleaseIfNeeded(shape);
                shape = null;
            }

            return builder.ToString();
        }
        finally
        {
            ComInteropHelper.ReleaseIfNeeded(shape);
            ComInteropHelper.ReleaseIfNeeded(shapes);
            ComInteropHelper.ReleaseIfNeeded(notesPage);
        }
    }

    private static string GetSlideName(object slide)
    {
        dynamic slideDynamic = slide;
        return ((string?)slideDynamic.Name)?.Trim() ?? "Untitled slide";
    }

    private static string ReadNotesShapeText(object shape)
    {
        dynamic shapeDynamic = shape;

        if (!ShapeCanContainText(shapeDynamic))
        {
            return string.Empty;
        }

        int shapeType = TryGetInt(() => shapeDynamic.Type, defaultValue: 0);
        if (shapeType == MsoPlaceholderShapeType)
        {
            int placeholderType = TryGetInt(() => shapeDynamic.PlaceholderFormat.Type, defaultValue: -1);
            if (placeholderType is not PlaceholderBody and not PlaceholderVerticalBody)
            {
                return string.Empty;
            }
        }

        string text = ReadShapeText(shapeDynamic);
        return string.IsNullOrWhiteSpace(text) ? string.Empty : text.Trim();
    }

    private static bool ShapeCanContainText(dynamic shapeDynamic)
    {
        try
        {
            return Convert.ToInt32(shapeDynamic.HasTextFrame) == MsoTrue &&
                   Convert.ToInt32(shapeDynamic.TextFrame.HasText) == MsoTrue;
        }
        catch
        {
            return false;
        }
    }

    private static string ReadShapeText(dynamic shapeDynamic)
    {
        try
        {
            string? text = (string?)shapeDynamic.TextFrame.TextRange.Text;
            if (!string.IsNullOrWhiteSpace(text))
            {
                return text;
            }
        }
        catch
        {
            // Fall back to TextFrame2 below.
        }

        try
        {
            string? text = (string?)shapeDynamic.TextFrame2.TextRange.Text;
            return text ?? string.Empty;
        }
        catch
        {
            return string.Empty;
        }
    }

    private static int TryGetInt(Func<object?> accessor, int defaultValue)
    {
        try
        {
            object? value = accessor();
            return value is null ? defaultValue : Convert.ToInt32(value);
        }
        catch
        {
            return defaultValue;
        }
    }
}
